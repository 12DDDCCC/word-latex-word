"""共享工具函数 — 跨skill使用的通用功能

模块结构：
- caption_utils: Caption文本格式化（normalize/clean/strip）
- caption_detect: Caption检测（通过字体样式/文本模式判断是否为caption）
- latex_text_utils: LaTeX文本处理（转义/括号匹配/上下标）
- latex_parse_utils: LaTeX解析常量和辅助函数
- unit_convert: 单位转换（pt/sz/twips/cm）
- word_xml_utils: Word XML命名空间操作
"""

# Caption格式化
from .caption_utils import (
    normalize_caption, clean_caption, strip_caption_prefix,
    clean_caption_prefix_in_tex
)

# Caption检测
from .caption_detect import (
    get_para_font_size_xml, is_caption_by_font_style, is_caption_paragraph,
    is_table_paragraph, get_body_font_size, find_caption_and_context,
    find_legend_paragraphs, find_context_body_text
)

# LaTeX文本处理
from .latex_text_utils import (
    match_balanced_braces, escape_latex, to_subscript, to_superscript,
    clean_latex_text
)

# LaTeX解析常量和辅助函数
from .latex_parse_utils import (
    LATEX_SIZE_PT, LATEX_SIZE_PT_11, LATEX_SIZE_PT_12,
    WEIGHT_MAP, SHAPE_MAP, FONT_CODE_TO_NAME,
    BS, cmd, la_size_to_pt, size_style, len_to_mm, find_balanced_braces
)

# 兼容别名（子模块使用 _cmd/_la_size_to_pt/_size_style/_len_to_mm/_read 等旧名）
_cmd = cmd
_la_size_to_pt = la_size_to_pt
_size_style = size_style
_len_to_mm = len_to_mm
def _read(path):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()

# 单位转换
from .unit_convert import pt_to_sz, cm_to_twips, width_to_pt

# Word XML工具
from .word_xml_utils import (
    W_NS, tag_local, wattr, get_all_attrs,
    iter_runs_recursive, get_run_color, get_run_text, get_run_bold
)
