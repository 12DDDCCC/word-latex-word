#!/usr/bin/env python3
"""TikZ 表格解析模块

将 LaTeX TikZ 表格解析为 JSON 格式，实现 100% 无损转换。

算法: draw-first boundary deduction
1. 从 draw 线段坐标建立精确的行列边界网格
2. 水平线的 x 范围差异推断合并单元格区域
3. 在无垂直线分隔的区域内，用节点 x 坐标细分列
4. 节点映射到单元格，检测 gridSpan/vMerge

辅助函数已移至:
- _tikz_geometry.py: 几何计算辅助函数
- _tikz_grid.py: 网格/线段处理函数
"""
import re
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import match_balanced_braces as _match_balanced_braces, to_subscript as _to_subscript
from shared.unit_convert import pt_to_sz as _pt_to_sz, cm_to_twips as _cm_to_twips

# 从子模块导入并 re-export（保持向后兼容）
from _tikz_geometry import (
    _empty_json,
    _cluster_coords,
    _merge_boundary,
    _find_interval,
    _find_nearest_index,
    _get_outer_border_sz,
    _clean_tikz_text,
)
from _tikz_grid import (
    _parse_segments,
    _classify_segments,
    _build_grid_from_segments,
    _normalize_draw_cmd,
    _extract_draw_coords,
    _parse_node_commands,
    _collect_all_positions,
)


# LaTeX font 命令 → Word size_pt
_FONT_MAP = {
    'tiny': 6, 'scriptsize': 8, 'footnotesize': 9,
    'small': 10, 'normalsize': 12, 'large': 14,
    'Large': 16, 'LARGE': 18, 'huge': 20, 'Huge': 24,
}


def tikz_to_json(tikz_body, table_env='', layout_spec=None):
    """TikZ 表格 → JSON (100% 无损)

    Args:
        tikz_body: \\begin{tikzpicture}...\\end{tikzpicture} 内的内容
        table_env: table 环境内容（含 caption）
        layout_spec: 排版规格

    Returns:
        dict: 符合 gen_table_from_json.py 输入格式的 JSON
    """
    # 0. 检查 meta 注释（caption行信息 + 结构元数据）
    has_caption_row = False
    caption_text = ''
    meta_json = None
    for line in tikz_body.split('\n'):
        m = re.match(r'\s*%\s*meta:has_caption_row=(\d)', line)
        if m and int(m.group(1)) == 1:
            has_caption_row = True
        m2 = re.match(r'\s*%\s*meta:caption_text=(.*)', line)
        if m2:
            caption_text = m2.group(1).strip()
        m3 = re.match(r'\s*%\s*meta:json=(.+)', line)
        if m3:
            try:
                import base64
                meta_json = json.loads(base64.b64decode(m3.group(1).strip()).decode('utf-8'))
            except Exception:
                meta_json = None

    # 1. 提取 \draw 线段
    h_segments = []  # (y, x_start, x_end, width_pt)
    v_segments = []  # (x, y_start, y_end, width_pt)
    draw_pattern = r'\\draw\[([^\]]*)\]\s*\(([^,]+),([^)]+)\)\s*--\s*\(([^,]+),([^)]+)\)\s*;'
    for m in re.finditer(draw_pattern, tikz_body):
        attrs = m.group(1)
        x1, y1 = float(m.group(2)), float(m.group(3))
        x2, y2 = float(m.group(4)), float(m.group(5))

        lw_m = re.search(r'line width=([0-9.]+)pt', attrs)
        width_pt = float(lw_m.group(1)) if lw_m else 0.4

        if abs(y1 - y2) < 0.001:  # 水平线
            xs, xe = min(x1, x2), max(x1, x2)
            h_segments.append((y1, xs, xe, width_pt))
        elif abs(x1 - x2) < 0.001:  # 垂直线
            ys, ye = min(y1, y2), max(y1, y2)
            v_segments.append((x1, ys, ye, width_pt))

    # 2. 提取 \node 文本
    nodes = []  # (x, y, text, is_bold, font_size, text_width)
    node_pattern = r'\\node\[([^\]]*)\]\s*(?:at\s*)?\(([^,]+),([^)]+)\)\s*\{'
    for m in re.finditer(node_pattern, tikz_body):
        attrs = m.group(1)
        x, y = float(m.group(2)), float(m.group(3))
        brace_start = m.end() - 1
        text = _match_balanced_braces(tikz_body, brace_start)

        font_m = re.search(r'font=\\(\w+)', attrs)
        font_cmd = font_m.group(1) if font_m else 'normalsize'
        size_pt = _FONT_MAP.get(font_cmd, 12)

        tw_m = re.search(r'text width=([0-9.]+)cm', attrs)
        text_width = float(tw_m.group(1)) if tw_m else 0

        is_bold = '\\textbf{' in text or 'font=\\bfseries' in attrs or 'font=\\textbf' in attrs

        nodes.append((x, y, text, is_bold, size_pt, text_width))

    if not nodes:
        return _empty_json()

    # 3. 列边界推导 — 频率投票 + 垂直线增强算法
    # 3a. 节点按 y 聚类为逻辑行
    node_y_abs = [(abs(y), x, y, text, is_bold, size_pt, text_width)
                  for x, y, text, is_bold, size_pt, text_width in nodes]
    node_y_abs.sort(key=lambda t: t[0])
    row_groups = []  # [(y_center, [nodes_in_row])]
    for item in node_y_abs:
        y_abs = item[0]
        if row_groups and abs(y_abs - row_groups[-1][0]) < 0.1:
            row_groups[-1][1].append(item)
        else:
            row_groups.append((y_abs, [item]))

    # 3b. 统计每个节点 x 聚类出现的行频率
    all_x_vals = sorted(set(round(n[1], 3) for n in node_y_abs))
    x_clusters = _cluster_coords(all_x_vals, tol=0.1)

    x_cluster_row_count = {}
    for xc in x_clusters:
        count = 0
        for _, row_nodes in row_groups:
            row_x_vals = [round(n[1], 3) for n in row_nodes]
            row_x_clusters = _cluster_coords(row_x_vals, tol=0.1)
            if any(abs(rxc - xc) < 0.1 for rxc in row_x_clusters):
                count += 1
        x_cluster_row_count[xc] = count

    # 3c. 频率阈值
    if not x_cluster_row_count:
        return _empty_json()

    max_freq = max(x_cluster_row_count.values())
    threshold = max(max_freq * 0.4, 2)

    col_centers = sorted(xc for xc, freq in x_cluster_row_count.items()
                         if freq >= threshold)

    if len(col_centers) < 2:
        col_centers = sorted(xc for xc, freq in x_cluster_row_count.items()
                             if freq > 1)
    if len(col_centers) < 2:
        widest_row = max(row_groups, key=lambda g: len(g[1])) if row_groups else ([], [])
        col_centers = sorted(set(round(n[1], 3) for n in widest_row[1]))

    if not col_centers:
        return _empty_json()

    # 3d. 从列中心推导列边界
    x_pos = []
    if len(col_centers) >= 2:
        min_gap = min(col_centers[i+1] - col_centers[i] for i in range(len(col_centers)-1))
        x_pos.append(col_centers[0] - min_gap / 2)
        for i in range(len(col_centers) - 1):
            x_pos.append((col_centers[i] + col_centers[i+1]) / 2)
        x_pos.append(col_centers[-1] + min_gap / 2)
    else:
        x_pos = [0.0, col_centers[0] * 2 if col_centers[0] > 0 else 5.0]

    # 用 draw 线段的精确坐标校准 x_pos
    draw_x_coords = set()
    for _, xs, xe, _ in h_segments:
        draw_x_coords.add(round(xs, 3))
        draw_x_coords.add(round(xe, 3))
    for vx, _, _, _ in v_segments:
        draw_x_coords.add(round(abs(vx), 3))

    for dx in sorted(draw_x_coords):
        nearest_idx = min(range(len(x_pos)), key=lambda i: abs(x_pos[i] - dx))
        dist = abs(x_pos[nearest_idx] - dx)
        if dist < 0.15:
            x_pos[nearest_idx] = dx
        elif dx < x_pos[0] - 0.01:
            x_pos[0] = dx
        elif dx > x_pos[-1] + 0.01:
            x_pos[-1] = dx

    # 3e. 行边界
    y_from_draw = set()
    for y, xs, xe, _ in h_segments:
        y_from_draw.add(round(abs(y), 3))

    y_pos = sorted(y_from_draw)

    # 用节点行中心补充行边界
    row_centers = sorted(set(round(g[0], 3) for g in row_groups))

    if row_centers and y_pos and row_centers[0] < y_pos[0] - 0.05:
        gap = y_pos[0] - row_centers[0]
        title_top = row_centers[0] - min(gap, 0.3)
        title_bottom = (row_centers[0] + y_pos[0]) / 2
        y_pos.extend([title_top, title_bottom])
        y_pos.sort()

    # 补充行边界
    new_y_bounds = []
    for i in range(len(y_pos) - 1):
        top, bottom = y_pos[i], y_pos[i + 1]
        centers_in = [c for c in row_centers if top + 0.05 < c < bottom - 0.05]

        data_centers = []
        for c in centers_in:
            for g_y, g_nodes in row_groups:
                if abs(g_y - c) < 0.05:
                    n_nodes = len(g_nodes)
                    n_cols_in_row = len(set(round(n[1], 3) for n in g_nodes))
                    if n_nodes <= 2 and n_cols_in_row <= 2:
                        continue
                    data_centers.append(c)
                    break

        if len(data_centers) >= 2:
            for j in range(len(data_centers) - 1):
                new_y_bounds.append((data_centers[j] + data_centers[j+1]) / 2)

    for b in new_y_bounds:
        y_pos.append(b)
    y_pos.sort()

    if len(y_pos) < 2:
        return _empty_json()

    num_cols = len(x_pos) - 1
    num_rows = len(y_pos) - 1

    if num_cols <= 0 or num_rows <= 0:
        return _empty_json()

    # 4-5. 生成 grid_cols 和 row_heights
    grid_cols = []
    for i in range(num_cols):
        w_cm = x_pos[i + 1] - x_pos[i]
        grid_cols.append({'width_twips': _cm_to_twips(w_cm)})

    row_heights = []
    for i in range(num_rows):
        h_cm = y_pos[i + 1] - y_pos[i]
        tw = _cm_to_twips(h_cm) if h_cm > 0 else 400
        row_heights.append(tw)

    # 元数据覆盖
    if meta_json and 'cw' in meta_json:
        meta_cols = meta_json['cw']
        if len(meta_cols) == num_cols:
            grid_cols = [{'width_twips': w} for w in meta_cols]
        meta_rh = meta_json.get('rh', [])
        if len(meta_rh) == num_rows:
            row_heights = list(meta_rh)

        if 'xp' in meta_json and 'yp' in meta_json:
            meta_xp = meta_json['xp']
            meta_yp = meta_json['yp']
            if len(meta_xp) == len(x_pos):
                x_pos = list(meta_xp)
            if len(meta_yp) >= 2 and len(meta_yp) != len(y_pos):
                y_pos = list(meta_yp)
                num_rows = len(y_pos) - 1
                grid_cols = [{'width_twips': _cm_to_twips(x_pos[i + 1] - x_pos[i])}
                             for i in range(num_cols)]
                row_heights = [_cm_to_twips(y_pos[i + 1] - y_pos[i]) if y_pos[i + 1] - y_pos[i] > 0 else 400
                               for i in range(num_rows)]
                if len(meta_cols) == num_cols:
                    grid_cols = [{'width_twips': w} for w in meta_cols]
                if len(meta_rh) == num_rows:
                    row_heights = list(meta_rh)
            elif len(meta_yp) == len(y_pos):
                y_pos = list(meta_yp)

    # 6. 将 node 映射到单元格 + 检测 gridSpan/vMerge
    cell_map = {}

    col_center_positions = [(x_pos[i] + x_pos[i + 1]) / 2
                            for i in range(num_cols)]
    col_half_widths = [(x_pos[i + 1] - x_pos[i]) / 2
                       for i in range(num_cols)]

    # Pass 0: 识别 vMerge 居中节点
    vm_center_nodes = {}
    vm_mapped_ys = set()
    for gi, (g_y, g_nodes) in enumerate(row_groups):
        n_nodes = len(g_nodes)
        n_x = len(set(round(n[1], 3) for n in g_nodes))
        if n_nodes <= 2 and n_x <= 2:
            if gi > 0 and len(row_groups[gi - 1][1]) >= 3:
                above_n_nodes = len(row_groups[gi - 1][1])
                if above_n_nodes < n_nodes * 4:
                    above_x_vals = [round(n[1], 3) for n in row_groups[gi - 1][1]]
                    cur_x_vals = [round(n[1], 3) for n in g_nodes]
                    x_overlap = any(
                        any(abs(cx - ax) < 0.3 for ax in above_x_vals)
                        for cx in cur_x_vals
                    )
                    if x_overlap and g_y not in vm_mapped_ys:
                        above_y = row_groups[gi - 1][0]
                        above_ri = _find_interval(above_y, y_pos)
                        if above_ri is not None:
                            vm_center_nodes[gi] = above_ri
                            vm_mapped_ys.add(g_y)

    # Pass 1: 映射节点到行+最近列
    row_node_map = {}
    for x, y, text, is_bold, size_pt, text_width in nodes:
        y_abs = abs(y)
        mapped_ri = None
        for gi, target_ri in vm_center_nodes.items():
            g_y = row_groups[gi][0]
            if abs(y_abs - g_y) < 0.05:
                mapped_ri = target_ri
                break

        if mapped_ri is None:
            ri = _find_interval(y_abs, y_pos)
        else:
            ri = mapped_ri

        if ri is None:
            continue
        if ri not in row_node_map:
            row_node_map[ri] = []
        row_node_map[ri].append((x, y, text, is_bold, size_pt, text_width))

    for ri, row_nodes in row_node_map.items():
        row_nodes.sort(key=lambda n: n[0])

        center_nodes = []
        span_candidates = []
        all_node_classifications = []

        for node in row_nodes:
            x = node[0]
            ci = _find_interval(x, x_pos)
            if ci is None:
                continue

            dist = abs(x - col_center_positions[ci]) if ci < len(col_center_positions) else 999
            half_w = col_half_widths[ci] if ci < len(col_half_widths) else 0.5

            is_center = dist <= half_w * 0.6
            all_node_classifications.append((ci, node, is_center, dist))

        center_cis = sorted(set(ci for ci, _, is_center, _ in all_node_classifications if is_center))
        has_gaps = False
        if len(center_cis) >= 2 and len(center_cis) < num_cols:
            for i in range(len(center_cis) - 1):
                if center_cis[i+1] - center_cis[i] > 2:
                    has_gaps = True
                    break

        if has_gaps:
            gap_cols = set(range(num_cols)) - set(center_cis)
            for ci, node, is_center, dist in all_node_classifications:
                x = node[0]
                if not is_center:
                    span_candidates.append((ci, node, dist))
                    continue
                best_span_ci = ci
                best_span_gs = 1
                best_span_dist = dist
                best_covers_gap = False
                for try_ci in range(max(0, ci - 3), min(num_cols, ci + 4)):
                    for try_gs in range(2, min(num_cols - try_ci + 1, ci + 4)):
                        span_center = sum(col_center_positions[try_ci + s]
                                          for s in range(try_gs)) / try_gs
                        d = abs(x - span_center)
                        span_cols = set(range(try_ci, try_ci + try_gs))
                        covers_gap = bool(span_cols & gap_cols)
                        if covers_gap and d < col_half_widths[ci]:
                            if not best_covers_gap or d < best_span_dist:
                                best_span_dist = d
                                best_span_ci = try_ci
                                best_span_gs = try_gs
                                best_covers_gap = True
                        elif not best_covers_gap and d < best_span_dist:
                            best_span_dist = d
                            best_span_ci = try_ci
                            best_span_gs = try_gs
                if best_span_gs > 1:
                    span_candidates.append((best_span_ci, node, best_span_dist))
                else:
                    center_nodes.append((ci, node))
        else:
            for ci, node, is_center, dist in all_node_classifications:
                if is_center:
                    center_nodes.append((ci, node))
                else:
                    span_candidates.append((ci, node, dist))

        # Pass 2: 先放列中心节点，再处理跨列候选
        for ci, node in center_nodes:
            if (ri, ci) in cell_map:
                continue
            cell_map[(ri, ci)] = {
                'text': node[2], 'is_bold': node[3],
                'size_pt': node[4], 'gridSpan': 1,
            }

        for nearest_ci, node, dist in span_candidates:
            x = node[0]
            occupied = set()
            for (r, c), info in cell_map.items():
                if r == ri:
                    gs = info.get('gridSpan', 1)
                    for s in range(gs):
                        occupied.add(c + s)

            best_ci = nearest_ci
            best_gs = 1
            best_dist = abs(x - col_center_positions[nearest_ci]) if nearest_ci < len(col_center_positions) else 999

            for try_ci in range(max(0, nearest_ci - 2), min(num_cols, nearest_ci + 3)):
                for try_gs in range(2, num_cols - try_ci + 1):
                    try_cols = set(range(try_ci, try_ci + try_gs))
                    if try_cols & occupied:
                        continue
                    span_center = sum(col_center_positions[try_ci + s]
                                      for s in range(try_gs)) / try_gs
                    d = abs(x - span_center)
                    min_half_w = min(col_half_widths[try_ci + s] for s in range(try_gs))
                    if d < best_dist - min_half_w * 0.3:
                        best_dist = d
                        best_ci = try_ci
                        best_gs = try_gs
                    elif d < best_dist + min_half_w * 0.3 and try_gs < best_gs:
                        best_dist = d
                        best_ci = try_ci
                        best_gs = try_gs

            cell_map[(ri, best_ci)] = {
                'text': node[2], 'is_bold': node[3],
                'size_pt': node[4], 'gridSpan': best_gs,
            }
            for s in range(best_gs):
                occupied.add(best_ci + s)

    # vMerge 检测
    vm_center_target_rows = set(vm_center_nodes.values())
    vmerge_map = {}
    row_wide_span_count = {}
    for ri in range(num_rows):
        count = 0
        for ci in range(num_cols):
            if (ri, ci) in cell_map:
                if cell_map[(ri, ci)].get('gridSpan', 1) > 1:
                    count += 1
        row_wide_span_count[ri] = count

    for ci in range(num_cols):
        ri = 0
        while ri < num_rows:
            if row_wide_span_count.get(ri, 0) >= 2:
                ri += 1
                continue
            covered = False
            for prev_ci in range(ci):
                if (ri, prev_ci) in cell_map:
                    gs_c = cell_map[(ri, prev_ci)].get('gridSpan', 1)
                    if prev_ci + gs_c > ci:
                        covered = True
                        break
            if covered:
                ri += 1
                continue
            if (ri, ci) in cell_map:
                info = cell_map[(ri, ci)]
                gs = info.get('gridSpan', 1)
                span = 1
                while ri + span < num_rows:
                    next_ri = ri + span
                    spanned_by_wide_cell = False
                    for prev_ci in range(ci):
                        if (next_ri, prev_ci) in cell_map:
                            pgs = cell_map[(next_ri, prev_ci)].get('gridSpan', 1)
                            if pgs > 1 and prev_ci + pgs > ci:
                                spanned_by_wide_cell = True
                                break
                    if spanned_by_wide_cell:
                        break
                    has_wide_span_in_next = False
                    for next_ci in range(num_cols):
                        if (next_ri, next_ci) in cell_map:
                            if cell_map[(next_ri, next_ci)].get('gridSpan', 1) > 2:
                                has_wide_span_in_next = True
                                break
                    if has_wide_span_in_next:
                        break
                    if next_ri in vm_center_target_rows:
                        break
                    next_row_content_count = 0
                    for next_ci in range(num_cols):
                        if (next_ri, next_ci) in cell_map:
                            next_text = cell_map[(next_ri, next_ci)].get('text', '')
                            if next_text:
                                next_row_content_count += 1
                    if 0 < next_row_content_count <= 2:
                        break
                    next_covered = False
                    for prev_ci in range(ci):
                        if (next_ri, prev_ci) in cell_map:
                            pgs = cell_map[(next_ri, prev_ci)].get('gridSpan', 1)
                            if prev_ci + pgs > ci:
                                next_covered = True
                                break
                    if next_covered:
                        break
                    all_empty = True
                    for s in range(gs):
                        if (next_ri, ci + s) in cell_map:
                            all_empty = False
                            break
                    if all_empty:
                        span += 1
                    else:
                        break
                if span > 1:
                    vmerge_map[(ri, ci)] = span
                    for s in range(1, span):
                        if (ri + s, ci) not in cell_map:
                            cell_map[(ri + s, ci)] = {
                                'text': '', 'is_bold': False,
                                'size_pt': 12, 'gridSpan': gs,
                                '_vmerge_continue': True,
                            }
                        else:
                            cell_map[(ri + s, ci)]['_vmerge_continue'] = True
                            cell_map[(ri + s, ci)]['gridSpan'] = gs
                ri += span
            else:
                ri += 1

    # 元数据 cell 覆盖
    if meta_json and 'cells' in meta_json:
        new_cell_map = {}
        new_vmerge_map = {}
        meta_cells = meta_json['cells']
        for c in meta_cells:
            ri = c['r']
            ci = c['c']
            gs = c.get('gs', 1)
            vm = c.get('vm', None)
            text = c.get('t', '')
            is_bold = c.get('b', False)
            borders = c.get('bd', {})
            for vertical_edge in ('left', 'right'):
                borders[vertical_edge] = {
                    'val': 'nil',
                    'sz': '0',
                    'color': 'auto',
                    'space': '0',
                }
            if ri < 0 or ri >= num_rows or ci < 0 or ci >= num_cols:
                continue
            if gs > num_cols - ci:
                gs = num_cols - ci
            cell_info = {
                'text': text, 'is_bold': is_bold,
                'size_pt': 12, 'gridSpan': gs,
                'borders': borders,
                'vAlign': c.get('va', ''),
            }
            if vm == 'continue':
                cell_info['_vmerge_continue'] = True
            elif vm == 'restart':
                new_vmerge_map[(ri, ci)] = 1
            new_cell_map[(ri, ci)] = cell_info

        for (ri, ci) in new_vmerge_map:
            gs = new_cell_map[(ri, ci)].get('gridSpan', 1)
            span = 1
            for next_ri in range(ri + 1, num_rows):
                if (next_ri, ci) in new_cell_map:
                    next_info = new_cell_map[(next_ri, ci)]
                    if next_info.get('_vmerge_continue') and next_info.get('gridSpan') == gs:
                        span += 1
                    else:
                        break
                else:
                    break
            new_vmerge_map[(ri, ci)] = span
        cell_map = new_cell_map
        vmerge_map = new_vmerge_map

    # 7. 边框线段 → cell borders
    cell_borders = {}
    for ri in range(num_rows):
        for ci in range(num_cols):
            cell_borders[(ri, ci)] = {}

    for y_val, xs, xe, width_pt in h_segments:
        sz = _pt_to_sz(width_pt)
        border = {'val': 'single', 'sz': str(sz), 'color': '000000', 'space': '0'}
        row_edge = _find_nearest_index(abs(y_val), y_pos)
        if row_edge is None:
            if abs(y_val) <= y_pos[0]:
                row_edge = 0
            elif abs(y_val) >= y_pos[-1]:
                row_edge = len(y_pos) - 1
            else:
                continue
        col_start = _find_nearest_index(xs, x_pos)
        if col_start is None:
            if xs <= x_pos[0]:
                col_start = 0
            else:
                continue
        col_end = _find_nearest_index(xe, x_pos)
        if col_end is None:
            if xe >= x_pos[-1]:
                col_end = num_cols
            else:
                continue
        if row_edge > 0:
            for ci in range(col_start, col_end):
                cell_borders[(row_edge - 1, ci)]['bottom'] = dict(border)
        if row_edge < num_rows:
            for ci in range(col_start, col_end):
                cell_borders[(row_edge, ci)]['top'] = dict(border)

    for key in cell_borders:
        for edge in ('top', 'bottom', 'left', 'right'):
            if edge not in cell_borders[key]:
                cell_borders[key][edge] = {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'}

    # 8. 提取 caption
    caption = ''
    cap_m = re.search(r'\\caption\{', table_env)
    if cap_m:
        caption = _match_balanced_braces(table_env, cap_m.end() - 1)

    # 9. 组装 JSON
    rows_data = []
    for ri in range(num_rows):
        cells_data = []
        col_pos = 0
        while col_pos < num_cols:
            cell_info = cell_map.get((ri, col_pos), {})
            gs = cell_info.get('gridSpan', 1)
            if gs > 1 and col_pos + gs > num_cols:
                gs = num_cols - col_pos
            text = _clean_tikz_text(cell_info.get('text', ''))
            is_bold = cell_info.get('is_bold', False)
            size_pt = cell_info.get('size_pt', 12)
            borders = cell_info.get('borders') or cell_borders.get((ri, col_pos), {})
            for vertical_edge in ('left', 'right'):
                borders[vertical_edge] = {
                    'val': 'nil',
                    'sz': '0',
                    'color': 'auto',
                    'space': '0',
                }
            is_vm_continue = cell_info.get('_vmerge_continue', False)
            if is_vm_continue:
                vm = 'continue'
            elif (ri, col_pos) in vmerge_map:
                vm = 'restart'
            else:
                vm = ''
            paragraphs = []
            if text:
                run = {'text': text}
                fmt = {}
                if is_bold:
                    fmt['bold'] = True
                if size_pt != 12:
                    fmt['size_pt'] = size_pt
                fmt['font_ascii'] = 'Times New Roman'
                if fmt:
                    run['format'] = fmt
                paragraphs.append({
                    'align': 'center',
                    'runs': [run],
                })
            cell_data = {
                'col_start': col_pos,
                'gridSpan': gs,
                'vMerge': vm,
                'borders': borders,
            }
            if is_vm_continue:
                cell_data['vAlign'] = 'center'
            elif cell_info.get('vAlign'):
                cell_data['vAlign'] = cell_info.get('vAlign')
            if paragraphs:
                cell_data['paragraphs'] = paragraphs
            cells_data.append(cell_data)
            col_pos += gs
        row_data = {
            'row_height': str(row_heights[ri]),
            'row_height_rule': 'atLeast',
            'cells': cells_data,
        }
        if ri == 0 and any(c.get('paragraphs', []) and
                           any(r.get('format', {}).get('bold') for r in c.get('paragraphs', [{}]))
                           for c in cells_data):
            row_data['is_header'] = True
        rows_data.append(row_data)

    # vMerge 边框合并
    for ri, row in enumerate(rows_data):
        for ci, cell in enumerate(row['cells']):
            if cell.get('vMerge') == 'restart':
                span_ri = ri + 1
                while span_ri < len(rows_data):
                    span_cell = None
                    for sc in rows_data[span_ri]['cells']:
                        if sc.get('col_start') == cell.get('col_start'):
                            span_cell = sc
                            break
                    if span_cell and span_cell.get('vMerge') == 'continue':
                        span_bottom = span_cell.get('borders', {}).get('bottom', {})
                        if span_bottom.get('val') == 'single':
                            cell['borders']['bottom'] = dict(span_bottom)
                        span_ri += 1
                    else:
                        break

    # caption 行
    if has_caption_row and caption_text and num_cols > 0:
        caption_cell = {
            'col_start': 0,
            'gridSpan': num_cols,
            'vMerge': 'restart',
            'borders': {
                'top': {'val': 'single', 'sz': '8', 'color': '000000', 'space': '0'},
                'bottom': {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'},
                'left': {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'},
                'right': {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'},
            },
            'paragraphs': [{
                'align': 'center',
                'runs': [{'text': _clean_tikz_text(caption_text),
                          'format': {'bold': True, 'font_ascii': 'Times New Roman'}}],
            }],
        }
        caption_row = {
            'row_height': '400',
            'row_height_rule': 'atLeast',
            'cells': [caption_cell],
            'is_header': True,
        }
        rows_data.insert(0, caption_row)

    # 表格级边框
    tbl_borders = {}
    top_sz = _get_outer_border_sz(h_segments, y_pos, 0, 'top')
    if top_sz:
        tbl_borders['top'] = {'val': 'single', 'sz': str(top_sz), 'color': '000000', 'space': '0'}
    bottom_sz = _get_outer_border_sz(h_segments, y_pos, -1, 'bottom')
    if bottom_sz:
        tbl_borders['bottom'] = {'val': 'single', 'sz': str(bottom_sz), 'color': '000000', 'space': '0'}

    result = {
        'grid_cols': grid_cols,
        'table_properties': {
            'borders': tbl_borders,
            'width': str(sum(gc['width_twips'] for gc in grid_cols)),
        },
        'rows': rows_data,
    }

    if caption:
        result['position'] = {'table_caption': caption}

    return result
