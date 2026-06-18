#!/usr/bin/env python3
r"""BBL解析模块 — 解析.bbl文件提取参考文献映射

从 natbib 的 \bibitem 格式中提取 key → author_year 映射。
"""
import re
import os
import sys
from pathlib import Path

# 导入共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import match_balanced_braces as _match_balanced_braces

_ACCENT_MAP = {
    r"\'\i": "í", r"\'{\i}": "í",
    r"\'a": "á", r"\'e": "é", r"\'i": "í",
    r"\'o": "ó", r"\'u": "ú", r"\'{a}": "á",
    r"\'{e}": "é", r"\'{i}": "í", r"\'{o}": "ó",
    r"\'{u}": "ú",
}


def _clean_latex_name(text):
    if not text:
        return text
    for pattern in sorted(_ACCENT_MAP, key=len, reverse=True):
        text = text.replace(pattern, _ACCENT_MAP[pattern])
    text = re.sub(r'\\[a-zA-Z]+\{([^}]*)\}', r'\1', text)
    return text


def _parse_author_list(author_str):
    """解析bbl中的完整作者列表为作者名列表

    输入: "Wang, Ding, and Ma" 或 "Baker, Basu, ..., and van der Woude"
    输出: ["Wang", "Ding", "Ma"] 或 ["Baker", "Basu", ..., "van der Woude"]
    """
    # 去掉末尾的 "and "
    author_str = re.sub(r',?\s+and\s+', ', ', author_str)
    # 按逗号分割
    parts = [p.strip() for p in author_str.split(',') if p.strip()]
    return parts


def _format_author_list(authors, max_authors=None):
    """将作者列表格式化为引用显示格式

    1位: "Author"
    2位: "Author1 and Author2"
    3位: "Author1, Author2, and Author3"
    3+位(max_authors截断): "Author1, Author2, Author3, ..."
    """
    if max_authors and len(authors) > max_authors:
        shown = authors[:max_authors]
        return ', '.join(shown) + ', et al.'
    if len(authors) == 1:
        return authors[0]
    if len(authors) == 2:
        return authors[0] + ' and ' + authors[1]
    # 3+位: "A, B, and C"
    return ', '.join(authors[:-1]) + ', and ' + authors[-1]


def _parse_bbl(bbl_path):
    r"""解析.bbl文件，提取 \bibitem 的 key→author_year 映射和完整内容

    natbib的\bibitem格式: \bibitem[{ShortAuthor(Year)FullAuthors}]{key}
    例如: \bibitem[{Wang et~al.(2022a)Wang, Ding, and Ma}]{44}
    提取为: '44' → 'Wang, Ding, and Ma (2022a)'

    对于同年多篇(a/b标记)，使用完整作者列表区分不同论文。

    Returns:
        dict: {
            'cite_map': {key: author_year_str},
            'bbl_content': str,
        }
    """
    with open(bbl_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    cite_map = {}
    # 匹配 \bibitem[{label}]{key} — label中含ShortAuthor(Year)FullAuthor
    # 方括号内可能包含嵌套花括号，需要平衡匹配
    i = 0
    numeric_index = 0
    while i < len(content):
        m = re.match(r'\\bibitem', content[i:])
        if not m:
            i += 1
            continue

        pos = i + len('\\bibitem')
        label = None

        # 检查可选参数 [...]
        if pos < len(content) and content[pos] == '[':
            # 平衡匹配方括号
            depth = 1
            j = pos + 1
            while j < len(content) and depth > 0:
                if content[j] == '[':
                    depth += 1
                elif content[j] == ']':
                    depth -= 1
                j += 1
            label = content[pos+1:j-1]  # 去掉外层方括号
            pos = j

        # 匹配 {key}
        key_match = re.match(r'\{([^}]+)\}', content[pos:])
        if key_match:
            key = key_match.group(1)
            numeric_index += 1

            if label:
                # label格式: {ShortAuthor(Year)FullAuthorList} 或 Author(Year)
                # 去掉可能的外层花括号
                label_inner = label
                if label_inner.startswith('{') and label_inner.endswith('}'):
                    label_inner = label_inner[1:-1]

                # 预处理: \natexlab{a} → a (同年多篇标记)
                label_inner = re.sub(r'\\natexlab\{([^}]*)\}', r'\1', label_inner)
                label_inner = _clean_latex_name(label_inner)
                # 清理花括号（年份中的{a}等）
                label_inner = label_inner.replace('{', '').replace('}', '')

                # 提取 Author(Year)FullAuthor 格式
                # 例如: Wang et~al.(2022a)Wang, Ding, and Ma
                # 年份可能是 2025 或 2022a (natexlab已展开)
                full_match = re.match(r'(.+?)\((\d{4}[a-z]?)\)(.+)?', label_inner)
                if full_match:
                    short_author = full_match.group(1).strip()
                    year = full_match.group(2)
                    full_author = full_match.group(3)  # 可能为None

                    # 去掉年份中的a/b后缀，同姓同年不同论文无需区分
                    year = re.sub(r'[a-z]$', '', year)

                    if full_author:
                        # 有完整作者列表，统一使用缩写格式
                        # a/b标记本身已区分同姓同年不同论文，无需额外显示多位作者
                        full_author = re.sub(r'\s+', ' ', full_author.strip())
                        full_author = full_author.replace('~', ' ')
                        full_author = _clean_latex_name(full_author)
                        authors = _parse_author_list(full_author)
                        if len(authors) <= 2:
                            display = _format_author_list(authors)
                        else:
                            display = authors[0] + ' et al.'
                        cite_map[key] = f'{display} ({year})'
                    else:
                        cite_map[key] = f'{short_author} ({year})'
                else:
                    author_clean = _clean_latex_name(label_inner.replace('~', ' '))
                    cite_map[key] = author_clean
            else:
                cite_map[key] = str(numeric_index)

        # 跳到下一个可能的bibitem
        i = pos + (len(key_match.group(0)) if key_match else 1)

    return {'cite_map': cite_map, 'bbl_content': content}


def _find_bbl_file(tex_dir):
    """查找tex文件所在目录的.bbl文件"""
    for f in Path(tex_dir).iterdir():
        if f.suffix.lower() == '.bbl' and f.is_file():
            return str(f)
    return None
