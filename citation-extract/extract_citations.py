#!/usr/bin/env python3
r"""
从Word文档中无损提取引用标记
识别策略: 以括号()为边界定位引用标记 + 纯红色数字
颜色信息作为附加属性保存
"""
import zipfile, json, re, os, sys
from xml.etree import ElementTree as ET
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 从shared模块导入XML工具（消除重复代码）
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))
from shared.word_xml_utils import W_NS, tag_local, wattr, iter_runs_recursive, get_run_color, get_run_text, get_run_bold
from word_link_citations import extract_word_link_citations


def _parse_numbers(text):
    """从引用文本中提取编号列表
    支持: '1,2,3', '5-7', '22–23–24'(链式破折号)
    """
    nums = []
    text = text.strip()
    parts = [p.strip() for p in text.split(',') if p.strip()]
    for part in parts:
        dash_parts = re.split(r'[-–—]', part)
        dash_parts = [p.strip() for p in dash_parts if p.strip()]
        if len(dash_parts) > 2:
            for dp in dash_parts:
                if dp.isdigit():
                    nums.append(int(dp))
        elif len(dash_parts) == 2:
            s, e = dash_parts
            if s.isdigit() and e.isdigit():
                si, ei = int(s), int(e)
                if ei - si <= 50:
                    nums.extend(range(si, ei + 1))
                else:
                    nums.extend([si, ei])
        elif len(dash_parts) == 1 and dash_parts[0].isdigit():
            nums.append(int(dash_parts[0]))
    return sorted(set(nums))


def extract_citations(docx_path, bib_path=None):
    with zipfile.ZipFile(docx_path, 'r') as z:
        content = z.read('word/document.xml').decode('utf-8')

    root = ET.fromstring(content)
    body = root.find(f'{{{W_NS}}}body')

    # 扫描章节标题
    section_map = {}
    current_section = ''
    para_count = 0
    for child in body:
        if tag_local(child) != 'p':
            continue
        para_count += 1
        pPr = next((c for c in child if tag_local(c) == 'pPr'), None)
        if pPr is not None:
            pStyle = next((c for c in pPr if tag_local(c) == 'pStyle'), None)
            if pStyle is not None:
                sv = wattr(pStyle, 'val') or ''
                if 'Heading' in sv or 'heading' in sv or 'TOC' in sv:
                    ht = ''.join(t.text for r in iter_runs_recursive(child)
                                 for t in r.iter(f'{{{W_NS}}}t') if t.text)
                    current_section = ht.strip()
        section_map[para_count] = current_section

    # 遍历段落提取引用标记
    citations = []
    para_idx = 0
    current_heading = ''

    for child in body:
        if tag_local(child) != 'p':
            continue
        para_idx += 1
        current_heading = section_map.get(para_idx, current_heading)

        runs = list(iter_runs_recursive(child))

        # 构建字符级属性
        full_text = ''
        char_colors = []
        char_bold = []

        for r in runs:
            color = get_run_color(r)
            text = get_run_text(r)
            bold = get_run_bold(r)
            full_text += text
            char_colors.extend([color] * len(text))
            char_bold.extend([bold] * len(text))

        if not full_text:
            continue

        # === 策略1: 括号引用 ===
        i = 0
        while i < len(full_text):
            if full_text[i] != '(':
                i += 1
                continue

            open_pos = i
            close_pos = -1
            depth = 1
            j = i + 1
            while j < len(full_text) and depth > 0:
                if full_text[j] == '(':
                    depth += 1
                elif full_text[j] == ')':
                    depth -= 1
                    if depth == 0:
                        close_pos = j
                j += 1

            if close_pos < 0:
                i += 1
                continue

            inner = full_text[open_pos+1:close_pos]
            is_cite = bool(re.match(r'^[\d,\-\s;]+$', inner))
            has_red = 'EE0000' in char_colors[open_pos:close_pos+1]

            if has_red and is_cite:
                nums = _parse_numbers(inner)
                if nums and all(n <= 100 for n in nums):
                    inner_colors = char_colors[open_pos+1:close_pos]
                    red_inner = ''.join(ch for ci, ch in enumerate(inner)
                                        if ci < len(inner_colors) and inner_colors[ci] == 'EE0000')
                    black_inner = ''.join(ch for ci, ch in enumerate(inner)
                                          if ci < len(inner_colors) and inner_colors[ci] != 'EE0000')
                    full_mark = full_text[open_pos:close_pos+1]
                    ctx_before = full_text[max(0, open_pos-40):open_pos]
                    ctx_after = full_text[close_pos+1:min(len(full_text), close_pos+41)]
                    has_bold = any(char_bold[j] for j in range(open_pos, close_pos+1))

                    citations.append({
                        'type': 'bracket',
                        'text': full_mark,
                        'inner': inner,
                        'numbers': nums,
                        'has_red': has_red,
                        'red_part': red_inner.strip(),
                        'black_part': black_inner.strip(),
                        'bold': has_bold,
                        'para_index': para_idx,
                        'char_offset': open_pos,
                        'section': current_heading,
                        'before': ctx_before,
                        'after': ctx_after,
                    })

            i = close_pos + 1

        # === 策略2: 纯红色数字片段(无括号) ===
        ri = 0
        while ri < len(full_text):
            if char_colors[ri] != 'EE0000':
                ri += 1
                continue
            seg_start = ri
            while ri < len(full_text) and char_colors[ri] == 'EE0000':
                ri += 1
            seg_text = full_text[seg_start:ri]

            # 纯数字或逗号分隔数字，且不在括号中
            if re.match(r'^[\d,\-\s]+$', seg_text.strip()):
                # 检查是否已被括号策略捕获
                already = any(c['para_index'] == para_idx
                              and c['char_offset'] <= seg_start < c['char_offset'] + len(c['text'])
                              for c in citations)
                if not already:
                    nums = _parse_numbers(seg_text.strip())
                    if nums:
                        before = full_text[max(0, seg_start-40):seg_start]
                        after = full_text[ri:min(len(full_text), ri+40)]
                        citations.append({
                            'type': 'red_number',
                            'text': seg_text.strip(),
                            'inner': seg_text.strip(),
                            'numbers': nums,
                            'red_part': seg_text.strip(),
                            'black_part': '',
                            'bold': any(char_bold[j] for j in range(seg_start, ri)),
                            'para_index': para_idx,
                            'char_offset': seg_start,
                            'section': current_heading,
                            'before': before,
                            'after': after,
                        })

    red_citations = list(citations)
    link_result = extract_word_link_citations(docx_path, bib_path) if bib_path else {
        'total_link_citations': 0,
        'citations': [],
    }
    citations.extend(link_result.get('citations', []))

    return {
        'source_file': os.path.basename(docx_path),
        'total_citations': len(citations),
        'total_red_citations': len(red_citations),
        'total_link_citations': link_result.get('total_link_citations', 0),
        'citations': citations,
        'red_citations': red_citations,
        'link_citations': link_result.get('citations', []),
    }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_citations.py <docx_file> [output.json] [references.bib]")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'citations.json'

    bib_path = sys.argv[3] if len(sys.argv) > 3 else None
    result = extract_citations(docx_path, bib_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Extracted {result['total_citations']} citations from {result['source_file']}")
    for c in result['citations']:
        nums_str = ','.join(str(n) for n in c['numbers'])
        parts = []
        if c.get('black_part'): parts.append(f"black:{c['black_part']}")
        if c.get('red_part'): parts.append(f"red:{c['red_part']}")
        print(f"  [{c['para_index']}] {c['text']} -> [{nums_str}]  ({' '.join(parts)})  sec: {c['section']}")
