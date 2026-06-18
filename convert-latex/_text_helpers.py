#!/usr/bin/env python3
"""文本辅助函数：标题编号剥离、参考文献检测、摘要关键词提取"""

import re
import sys
from pathlib import Path

# shared模块导入
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

# 包内绝对导入（非Python包，不能使用相对导入）
from template_spec_extract import APPENDIX_KEYWORDS
from numbering_system import REF_KEYWORDS


def _strip_heading_number(latex, text):
    """从标题 latex 中剥离自动编号前缀和多余格式

    LaTeX 的 \\section{} 会自动编号和加粗，需要去掉：
    1. Word 标题中的编号（避免与 LaTeX 自动编号重复）
    2. \\textbf{} 包裹（section 命令自带加粗）
    """
    num_match = re.match(r'^(\d+(?:\.\d+)*)[\s、．.]*', text)
    if not num_match:
        # 无编号，只需去掉 \textbf{} 包裹
        clean = re.sub(r'\\textbf\{([^}]*)\}', r'\1', latex)
        return clean.strip()

    # 首先去掉所有 \textbf{} 包裹，提取内部文本
    result = re.sub(r'\\textbf\{([^}]*)\}', r'\1', latex)

    # 然后去掉开头的数字编号（如 "2.1 ", "1 ", "3" 等）
    # 数字后面可以跟空白/标点，也可以直接跟汉字（如 "2研究方法"）
    result = re.sub(r'^\d+(?:\.\d+)*[\s、．.]*', '', result)

    # 清理前后空白
    result = result.strip()

    return result


def _find_reference_start(paragraphs):
    """检测参考文献段落的起始位置

    优先检测 heading 级别的参考文献标题，再检测纯文本段落。

    Args:
        paragraphs: text_result['paragraphs']

    Returns:
        int or None: 参考文献起始段落的 para_index
    """
    ref_start_para = None
    for p in paragraphs:
        text_lower = p['text'].lower().strip()
        if p['heading_level'] and any(kw in text_lower for kw in REF_KEYWORDS):
            ref_start_para = p['para_index']
            break
    if ref_start_para is None:
        for p in paragraphs:
            text_stripped = p['text'].strip()
            text_lower = text_stripped.lower()
            if text_lower in ('references', 'reference', '参考文献', 'bibliography', '参考文献列表'):
                ref_start_para = p['para_index']
                break
    return ref_start_para


def _extract_abstract_keywords(paragraphs, ref_start_para, skeleton_info, layout_spec):
    """提取摘要和关键词的LaTeX行

    Args:
        paragraphs: text_result['paragraphs']
        ref_start_para: 参考文献起始位置（None则不截断）
        skeleton_info: 模板骨架信息
        layout_spec: 排版规格

    Returns:
        tuple: (abstract_lines, keywords_lines)
    """
    abstract_env = skeleton_info.get('abstract_env', 'abstract')
    abstract_cmd = skeleton_info.get('abstract_cmd')
    keywords_cmd = skeleton_info.get('keywords_cmd')
    keywords_env = skeleton_info.get('keywords_env', 'keywords')
    abs_spec = layout_spec.get('abstract', {}) if layout_spec else {}
    abs_font_size = abs_spec.get('font_size', '')
    abstract_lines = []
    keywords_lines = []

    def _strip_abstract_label(text):
        text = re.sub(
            r'^\s*\\textbf\{\s*Abstract\s*[:：.]?\s*\}\s*',
            '',
            text,
            flags=re.IGNORECASE,
        )
        return re.sub(r'^\s*Abstract\s*[:：.]?\s*', '', text, flags=re.IGNORECASE)

    for p in paragraphs:
        pi = p['para_index']
        semantic = p.get('semantic_type', 'unknown')
        if ref_start_para and pi >= ref_start_para:
            continue
        if semantic == 'abstract':
            if not abstract_lines:
                if not abstract_cmd:
                    abstract_lines.append(f'\\begin{{{abstract_env}}}')
                    if abs_font_size:
                        abstract_lines.append(f'\\{abs_font_size}')
            if p['latex'] and p['latex'].strip():
                abs_text = _strip_abstract_label(p['latex'])
                abstract_lines.append(abs_text)
        elif semantic == 'keywords':
            kw_text = re.sub(r'^\s*Key\s*words?\s*[:：]\s*|^\s*关键词\s*[:：]\s*', '', p['latex'], flags=re.IGNORECASE)
            if keywords_cmd:
                keywords_lines.append(f'{keywords_cmd}{{{kw_text}}}')
            else:
                keywords_lines.append(f'\\begin{{{keywords_env}}}')
                keywords_lines.append(kw_text)
                keywords_lines.append(f'\\end{{{keywords_env}}}')
    if abstract_lines:
        if abstract_cmd:
            optional = (skeleton_info.get('abstract_cmd_optional') or '').strip()
            opt_arg = f'[{optional}]' if optional else ''
            abstract_lines = [f'{abstract_cmd}{opt_arg}{{{" ".join(abstract_lines)}}}']
        else:
            abstract_lines.append(f'\\end{{{abstract_env}}}')
    return abstract_lines, keywords_lines
