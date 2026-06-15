#!/usr/bin/env python3
"""TikZ 几何计算辅助函数

包含:
- _cluster_coords: 聚类合并坐标值
- _merge_boundary: 将精确边界值合并到位置列表
- _find_interval: 找到 val 所在的区间索引
- _find_nearest_index: 找到最接近 val 的位置索引
- _get_outer_border_sz: 从最外围水平线段获取 sz 值
- _clean_tikz_text: 清理 TikZ 节点文本中的 LaTeX 命令
- _empty_json: 返回空表格 JSON
"""

import re

from shared.unit_convert import pt_to_sz as _pt_to_sz


def _empty_json():
    """返回空表格 JSON"""
    return {
        'grid_cols': [],
        'table_properties': {'borders': {}, 'width': '0'},
        'rows': [],
    }


def _cluster_coords(values, tol=0.05):
    """聚类合并坐标值：相近值（差值 < tol）归为同一组，取均值"""
    if not values:
        return []
    values = sorted(values)
    groups = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] < tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _merge_boundary(pos_list, value, tol=0.05):
    """将精确边界值合并到位置列表中"""
    for i, p in enumerate(pos_list):
        if abs(p - value) < tol:
            pos_list[i] = (p + value) / 2
            return
    pos_list.append(value)
    pos_list.sort()


def _find_interval(val, positions):
    """找到 val 所在的区间索引 (positions[i] <= val < positions[i+1])
    超出范围时映射到最接近的边界区间"""
    if not positions or len(positions) < 2:
        return 0
    for i in range(len(positions) - 1):
        if positions[i] - 0.01 <= val <= positions[i + 1] + 0.01:
            return i
    # val 在最上面的边界之上 → 映射到第一行
    if val < positions[0]:
        return 0
    # val 在最下面的边界之下 → 映射到最后一行
    return len(positions) - 2


def _find_nearest_index(val, positions):
    """找到最接近 val 的位置索引，tolerance=0.05cm"""
    best_idx = None
    best_dist = float('inf')
    for i, p in enumerate(positions):
        d = abs(val - p)
        if d < best_dist:
            best_dist = d
            best_idx = i
    if best_dist > 0.05:  # 收紧: 0.1 → 0.05
        return None
    return best_idx


def _get_outer_border_sz(h_segments, y_pos, idx, edge):
    """从最外围水平线段获取 sz 值"""
    if not h_segments:
        return 0
    target_y = y_pos[idx] if idx >= 0 else y_pos[-1]
    for y, xs, xe, w_pt in h_segments:
        if abs(abs(y) - target_y) < 0.01:
            return _pt_to_sz(w_pt)
    return 0


def _clean_tikz_text(text):
    """清理 TikZ 节点文本中的 LaTeX 命令"""
    t = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
    t = re.sub(r'\\textit\{([^}]*)\}', r'\1', t)
    t = re.sub(r'\\emph\{([^}]*)\}', r'\1', t)
    t = re.sub(r'\\underline\{([^}]*)\}', r'\1', t)
    t = re.sub(r'\\\w+\{([^}]*)\}', r'\1', t)
    t = t.replace('\\%', '%').replace('\\_', '_').replace('\\&', '&')
    t = t.replace('\\#', '#').replace('\\$', '$')
    t = t.replace('\\textasciitilde{}', '~').replace('\\textasciicircum{}', '^')
    t = re.sub(r'[${}]', '', t).strip()
    return t