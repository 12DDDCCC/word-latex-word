#!/usr/bin/env python3
"""TikZ 网格/线段处理函数

包含:
- _parse_segments: 解析 TikZ draw 命令中的线段
- _classify_segments: 将线段分类为水平/垂直/斜线
- _build_grid_from_segments: 从线段构建行列位置
- _normalize_draw_cmd: 规范化 draw 命令格式
- _extract_draw_coords: 从 draw 命令中提取坐标
- _parse_node_commands: 解析 TikZ node 命令
- _collect_all_positions: 收集所有位置信息
"""

import re

from shared.unit_convert import pt_to_sz as _pt_to_sz
from _tikz_geometry import _cluster_coords, _merge_boundary, _find_interval, _find_nearest_index


def _parse_segments(draw_cmds):
    """解析 TikZ draw 命令中的线段

    Args:
        draw_cmds: draw 命令字符串列表

    Returns:
        tuple: (h_segments, v_segments) 水平和垂直线段列表
        每条线段: (pos, start, end, width_pt)
    """
    h_segments = []  # (y, x_start, x_end, width_pt)
    v_segments = []  # (x, y_start, y_end, width_pt)

    for cmd in draw_cmds:
        # 提取坐标对
        coords = re.findall(
            r'\(\s*([\d.\-]+)\s*,\s*([\d.\-]+)\s*\)', cmd)
        if len(coords) < 2:
            continue

        # 提取线宽
        w_m = re.search(r'line\s+width\s*=\s*([\d.]+)\s*(?:pt|mm|cm)', cmd)
        width_pt = float(w_m.group(1)) if w_m else 0.4
        if 'mm' in (w_m.group(0) if w_m else ''):
            width_pt *= 2.835
        elif 'cm' in (w_m.group(0) if w_m else ''):
            width_pt *= 28.35

        # 分析线段
        for i in range(len(coords) - 1):
            x1, y1 = float(coords[i][0]), float(coords[i][1])
            x2, y2 = float(coords[i + 1][0]), float(coords[i + 1][1])

            if abs(y1 - y2) < 0.01:  # 水平线
                x_start, x_end = min(x1, x2), max(x1, x2)
                h_segments.append((y1, x_start, x_end, width_pt))
            elif abs(x1 - x2) < 0.01:  # 垂直线
                y_start, y_end = min(y1, y2), max(y1, y2)
                v_segments.append((x1, y_start, y_end, width_pt))

    return h_segments, v_segments


def _classify_segments(h_segments, v_segments):
    """将线段分类为边框线、内部线和虚线

    Args:
        h_segments: 水平线段列表
        v_segments: 垂直线段列表

    Returns:
        dict: 分类结果
    """
    result = {
        'outer_h': [],   # 外框水平线
        'outer_v': [],   # 外框垂直线
        'inner_h': [],   # 内部水平线
        'inner_v': [],   # 内部垂直线
        'dashed': [],    # 虚线
    }

    if not h_segments and not v_segments:
        return result

    # 找到最外围边界
    if h_segments:
        y_min = min(abs(y) for y, _, _, _ in h_segments)
        y_max = max(abs(y) for y, _, _, _ in h_segments)
    else:
        y_min = y_max = 0

    if v_segments:
        x_min = min(abs(x) for x, _, _, _ in v_segments)
        x_max = max(abs(x) for x, _, _, _ in v_segments)
    else:
        x_min = x_max = 0

    tol = 0.01
    for y, xs, xe, w in h_segments:
        is_outer = abs(abs(y) - y_min) < tol or abs(abs(y) - y_max) < tol
        if is_outer:
            result['outer_h'].append((y, xs, xe, w))
        else:
            result['inner_h'].append((y, xs, xe, w))

    for x, ys, ye, w in v_segments:
        is_outer = abs(abs(x) - x_min) < tol or abs(abs(x) - x_max) < tol
        if is_outer:
            result['outer_v'].append((x, ys, ye, w))
        else:
            result['inner_v'].append((x, ys, ye, w))

    return result


def _build_grid_from_segments(h_segments, v_segments):
    """从线段构建行列位置

    Args:
        h_segments: 水平线段列表
        v_segments: 垂直线段列表

    Returns:
        tuple: (y_positions, x_positions, h_by_row, v_by_col)
        - y_positions: 排序后的 y 坐标列表
        - x_positions: 排序后的 x 坐标列表
        - h_by_row: 每行的水平线段信息
        - v_by_col: 每列的垂直线段信息
    """
    # 收集所有 y 和 x 坐标
    y_values = [abs(y) for y, _, _, _ in h_segments]
    x_values = [abs(x) for x, _, _, _ in v_segments]

    # 聚类
    y_positions = _cluster_coords(y_values, tol=0.05)
    x_positions = _cluster_coords(x_values, tol=0.05)

    # 按行列组织线段
    h_by_row = {}
    for y, xs, xe, w in h_segments:
        row = _find_nearest_index(abs(y), y_positions)
        if row is not None:
            h_by_row.setdefault(row, []).append((xs, xe, w))

    v_by_col = {}
    for x, ys, ye, w in v_segments:
        col = _find_nearest_index(abs(x), x_positions)
        if col is not None:
            v_by_col.setdefault(col, []).append((ys, ye, w))

    return y_positions, x_positions, h_by_row, v_by_col


def _normalize_draw_cmd(cmd):
    """规范化 draw 命令格式

    统一空格、坐标格式等
    """
    # 移除多余空格
    cmd = re.sub(r'\s+', ' ', cmd.strip())
    # 统一坐标格式
    cmd = re.sub(r'\(\s*', '(', cmd)
    cmd = re.sub(r'\s*\)', ')', cmd)
    cmd = re.sub(r'\s*,\s*', ',', cmd)
    return cmd


def _extract_draw_coords(cmd):
    """从 draw 命令中提取坐标

    Args:
        cmd: draw 命令字符串

    Returns:
        list: [(x, y), ...] 坐标列表
    """
    coords = []
    for m in re.finditer(r'\(\s*([\d.\-]+)\s*,\s*([\d.\-]+)\s*\)', cmd):
        coords.append((float(m.group(1)), float(m.group(2))))
    return coords


def _parse_node_commands(tex_code):
    """解析 TikZ node 命令

    Args:
        tex_code: TikZ 代码字符串

    Returns:
        list: [{'x': float, 'y': float, 'text': str, 'anchor': str}, ...]
    """
    nodes = []
    # 匹配 \node[...] at (x,y) {text}; 或 \node (name) at (x,y) {text};
    for m in re.finditer(
        r'\\node\s*(?:\((\w+)\))?\s*(?:\[([^\]]*)\])?\s*'
        r'at\s*\(\s*([\d.\-]+)\s*,\s*([\d.\-]+)\s*\)\s*'
        r'(?:\{([^}]*)\})?\s*;',
        tex_code
    ):
        name = m.group(1) or ''
        opts = m.group(2) or ''
        x = float(m.group(3))
        y = float(m.group(4))
        text = m.group(5) or ''

        # 提取 anchor
        anchor = ''
        anchor_m = re.search(r'anchor\s*=\s*(\w+)', opts)
        if anchor_m:
            anchor = anchor_m.group(1)

        nodes.append({
            'name': name,
            'x': x,
            'y': y,
            'text': text,
            'anchor': anchor,
            'options': opts,
        })

    # 也匹配在 draw 命令中内嵌的 node
    for m in re.finditer(
        r'node\s*(?:\[([^\]]*)\])?\s*(?:\{([^}]*)\})',
        tex_code
    ):
        opts = m.group(1) or ''
        text = m.group(2) or ''

        # 从 draw 命令上下文中提取坐标比较复杂，这里简化处理
        # 只记录文本
        if text.strip():
            nodes.append({
                'name': '',
                'x': 0,
                'y': 0,
                'text': text,
                'anchor': '',
                'options': opts,
                'embedded': True,
            })

    return nodes


def _collect_all_positions(nodes, h_segments, v_segments):
    """收集所有位置信息（用于确定行列边界）

    Args:
        nodes: 解析后的 node 列表
        h_segments: 水平线段列表
        v_segments: 垂直线段列表

    Returns:
        tuple: (y_positions, x_positions)
    """
    # 从线段获取坐标
    y_values = [abs(y) for y, _, _, _ in h_segments]
    x_values = [abs(x) for x, _, _, _ in v_segments]

    # 从 node 获取坐标
    for node in nodes:
        if not node.get('embedded'):
            y_values.append(abs(node['y']))
            x_values.append(abs(node['x']))

    # 聚类合并
    y_positions = _cluster_coords(y_values, tol=0.05)
    x_positions = _cluster_coords(x_values, tol=0.05)

    return y_positions, x_positions