"""单位转换工具 — pt/sz/twips/cm之间的转换

2个skill中重复的单位转换逻辑合并到此模块：
- latex_table_parser: _pt_to_sz(), _cm_to_twips()
- tikz_table_gen: width_to_pt()
"""


def pt_to_sz(pt):
    """TikZ 线宽 pt → Word sz (1/8pt 单位)

    分段映射，与 Word booktabs 标准对齐：
    - ≥1.0pt (粗线/toprule/bottomrule) → sz=8 (1.0pt)
    - ≥0.7pt (中粗/midrule) → sz=6 (0.75pt)
    - <0.7pt (细线) → sz=4 (0.5pt)
    """
    if pt >= 1.0:
        return 8
    elif pt >= 0.7:
        return 6
    else:
        return 4


def cm_to_twips(cm):
    """cm → twips (1cm ≈ 567 twips)"""
    return round(cm * 567)


def width_to_pt(sz):
    """Word sz(1/8pt单位) → LaTeX线宽pt"""
    if sz >= 8:
        return 1.2
    elif sz >= 4:
        return 0.4
    return 0.0