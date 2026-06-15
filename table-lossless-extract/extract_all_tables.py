"""
从docx内部XML零损失提取所有表格的完整信息
包括：合并单元格、边框、底纹、对齐、字体、行高、段落格式
以及表格在文档中的位置（前导标题、段落、序号）
"""
import zipfile
import xml.etree.ElementTree as ET
import json
import sys
import os
import re
from pathlib import Path

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.word_xml_utils import W_NS, tag_local, wattr, get_all_attrs
from shared.caption_utils import clean_caption as _clean_caption

ET.register_namespace('w', W_NS)

# caption 正则：匹配"表1"、"表 1"、"表1.1"、"Table 1"、"Table1."、"Tab. 1"、"Talbe 1"(拼写错误)等
_RE_TABLE_CAP = re.compile(
    r'^\s*表\s*[\d\.]+|^Table\s*[\d\.]+|^Tab\.?\s*[\d\.]+|^Talbe\s*[\d\.]+',
    re.IGNORECASE
)

_RE_TABLE_CAP_STRICT = re.compile(
    r'^\s*(?:表|Table|Tab\.?|Talbe)\s*\d+(?:\.\d+)*(?=\s|[.．,，、:：。])',
    re.IGNORECASE
)


def _starts_table_caption_text(text):
    """Return True when text explicitly starts with a table caption label."""
    return bool(_RE_TABLE_CAP_STRICT.match((text or '').strip()))


def _is_short_body_continuation(text):
    """A short paragraph may be a split tail before a following table caption."""
    stripped = (text or '').strip()
    return bool(stripped) and len(stripped) <= 80 and not _starts_table_caption_text(stripped)


def extract_tables(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as doc:
            content = doc.read().decode('utf-8')

    root = ET.fromstring(content)

    def get_all_attrs(elem):
        result = {}
        for key, val in elem.attrib.items():
            if '}' in key:
                result[key.split('}')[1]] = val
            else:
                result[key] = val
        return result

    def parse_borders(borders_elem):
        result = {}
        if borders_elem is None:
            return result
        for b in borders_elem:
            result[tag_local(b)] = get_all_attrs(b)
        return result

    def parse_cell_props(tcPr):
        props = {
            'gridSpan': 1, 'vMerge': None, 'width': None,
            'shading': None, 'vAlign': None, 'borders': {},
            'tcMar': {},
        }
        if tcPr is None:
            return props
        for prop in tcPr:
            local = tag_local(prop)
            if local == 'gridSpan':
                v = wattr(prop, 'val')
                if v: props['gridSpan'] = int(v)
            elif local == 'vMerge':
                v = wattr(prop, 'val')
                props['vMerge'] = 'restart' if v == 'restart' else 'continue'
            elif local == 'tcW':
                props['width'] = wattr(prop, 'w')
            elif local == 'shd':
                props['shading'] = get_all_attrs(prop)
            elif local == 'vAlign':
                props['vAlign'] = wattr(prop, 'val')
            elif local == 'tcBorders':
                props['borders'] = parse_borders(prop)
            elif local == 'tcMar':
                for m in prop:
                    props['tcMar'][tag_local(m)] = wattr(m, 'w')
        return props

    def parse_paragraph(p_elem):
        para = {
            'text': '', 'align': None,
            'space_before': None, 'space_after': None,
            'line_spacing': None, 'runs': [],
        }
        pPr = next((c for c in p_elem if tag_local(c) == 'pPr'), None)
        if pPr:
            for prop in pPr:
                local = tag_local(prop)
                if local == 'jc': para['align'] = wattr(prop, 'val')
                elif local == 'spacing':
                    para['space_before'] = wattr(prop, 'before')
                    para['space_after'] = wattr(prop, 'after')
                    para['line_spacing'] = wattr(prop, 'line')
        for r in p_elem.iter(f'{{{W_NS}}}r'):
            run_info = {'text': '', 'format': {}}
            t_elem = next((c for c in r if tag_local(c) == 't'), None)
            text = t_elem.text if t_elem is not None and t_elem.text else ''
            run_info['text'] = text
            rPr = next((c for c in r if tag_local(c) == 'rPr'), None)
            if rPr:
                for prop in rPr:
                    local = tag_local(prop)
                    if local == 'rFonts':
                        run_info['format']['font_ascii'] = wattr(prop, 'ascii')
                        run_info['format']['font_eastAsia'] = wattr(prop, 'eastAsia')
                    elif local == 'sz':
                        val = wattr(prop, 'val')
                        if val:
                            run_info['format']['size_half_pt'] = int(val)
                            run_info['format']['size_pt'] = int(val) / 2
                    elif local == 'b':
                        run_info['format']['bold'] = True
                        bval = wattr(prop, 'val')
                        if bval == '0': run_info['format']['bold'] = False
                    elif local == 'i':
                        run_info['format']['italic'] = True
                    elif local == 'vertAlign':
                        run_info['format']['vertAlign'] = wattr(prop, 'val')
                    elif local == 'u':
                        run_info['format']['underline'] = wattr(prop, 'val')
                    elif local == 'color':
                        run_info['format']['color'] = wattr(prop, 'val')
            if run_info['format'] or text:
                para['runs'].append(run_info)
        para['text'] = ''.join(r['text'] for r in para['runs']) if para['runs'] else ''
        return para

    def get_para_text(p_elem):
        """获取段落纯文本"""
        text = ''
        for r in p_elem.iter(f'{{{W_NS}}}r'):
            for t in r.iter(f'{{{W_NS}}}t'):
                if t.text: text += t.text
        return text.strip()

    def get_para_align(p_elem):
        """获取段落对齐方式"""
        pPr = next((c for c in p_elem if tag_local(c) == 'pPr'), None)
        if pPr:
            jc = next((c for c in pPr if tag_local(c) == 'jc'), None)
            if jc:
                return wattr(jc, 'val')
        return None

    def get_para_indent(p_elem):
        """获取段落首行缩进（twips）"""
        pPr = next((c for c in p_elem if tag_local(c) == 'pPr'), None)
        if pPr:
            ind = next((c for c in pPr if tag_local(c) == 'ind'), None)
            if ind:
                first = wattr(ind, 'firstLine')
                if first:
                    return int(first)
        return 0

    def get_para_max_font_size(p_elem):
        """获取段落中最大字号（half-pt），用于区分表例/图例"""
        max_size = 0
        for r in p_elem.iter(f'{{{W_NS}}}r'):
            rPr = next((c for c in r if tag_local(c) == 'rPr'), None)
            if rPr:
                sz = next((c for c in rPr if tag_local(c) == 'sz'), None)
                if sz:
                    val = wattr(sz, 'val')
                    if val:
                        max_size = max(max_size, int(val))
        return max_size

    def _get_body_font_size(body_elem):
        """统计正文字号（出现次数最多的half-pt值）"""
        from collections import Counter
        size_counter = Counter()
        for child in body_elem:
            if tag_local(child) == 'p':
                ptext = get_para_text(child).strip()
                if not ptext:
                    continue
                sz = get_para_max_font_size(child)
                if sz > 0:
                    size_counter[sz] += 1
        if size_counter:
            return size_counter.most_common(1)[0][0]
        return None

    def _is_table_caption_para(p_elem, body_font_halfpt=None):
        """判断段落是否为表例(caption)

        表例与正文的格式区别：
        - 字号小于正文（正文12pt=24half-pt，表例通常9pt=18half-pt）
        - 居中对齐或左对齐（正文通常是两端对齐+首行缩进）
        - 无首行缩进（正文通常有首行缩进）
        - 文本以"表X"或"Table X"开头

        识别优先级：字号差异(最可靠) > 居中对齐+文本模式 > 文本模式+无缩进
        关键：字号差异是独立判断条件，不依赖文本模式匹配
        """
        ptext = get_para_text(p_elem).strip()
        if not ptext:
            return False

        # 格式特征1：字号明显小于正文 → 表例（最可靠，独立判断）
        para_size = get_para_max_font_size(p_elem)
        if body_font_halfpt and para_size > 0:
            if para_size <= body_font_halfpt - 4:  # 比正文小2pt以上(4 half-pt)
                return True

        # 以下规则需要文本模式匹配
        if not _RE_TABLE_CAP.match(ptext):
            return False

        # 格式特征2：字号略小（比正文小1-2pt）
        if body_font_halfpt and para_size > 0:
            if para_size < body_font_halfpt - 2:  # 比正文小1pt以上
                return True

        # 格式特征3：居中对齐
        align = get_para_align(p_elem)
        if align == 'center':
            return True

        # 格式特征4：无首行缩进 + 字号可获取且与正文相同或更小
        indent = get_para_indent(p_elem)
        if indent == 0 and para_size > 0 and body_font_halfpt:
            if para_size <= body_font_halfpt:
                return True

        return False

    def _find_table_legend_paragraphs(body_elem, tbl_idx, body_children, body_font_halfpt):
        """从表格后开始扫描，收集表例段落

        核心判定：用字体样式差异（字号 < 正文2pt以上）区分表例与正文。
        表例特征：sz ≤ body_font_halfpt - 4（如正文24→表例≤20，即≤10pt）
        停止条件：遇到正文段落（字号≥body_font）或 heading 或下一个表格
        """
        legends = []
        for fi in range(tbl_idx + 1, min(tbl_idx + 10, len(body_children))):
            fchild = body_children[fi]
            flocal = tag_local(fchild)
            if flocal == 'p':
                ftext = get_para_text(fchild).strip()
                if not ftext:
                    continue
                # 停止：标题
                pPr = next((c for c in fchild if tag_local(c) == 'pPr'), None)
                if pPr:
                    pStyle = next((c for c in pPr if tag_local(c) == 'pStyle'), None)
                    if pStyle:
                        style_val = wattr(pStyle, 'val') or ''
                        if 'Heading' in style_val or 'heading' in style_val:
                            break
                # 核心判定：用字体样式判断是表例还是正文
                if _is_table_caption_para(fchild, body_font_halfpt):
                    legends.append(ftext)
                else:
                    # 非表例 → 正文，停止收集
                    break
            elif flocal == 'tbl':
                break
        return legends

    def _find_table_context_body(body_elem, tbl_idx, body_children, direction, body_font_halfpt):
        """查找表格附近的正文段落文本

        direction='above': 向前查找最近的正文段落
        direction='below': 向后查找最近的正文段落（跳过表例段落）
        """
        if direction == 'above':
            for k in range(tbl_idx - 1, max(tbl_idx - 15, -1), -1):
                prev_child = body_children[k]
                prev_local = tag_local(prev_child)
                if prev_local == 'p':
                    prev_text = get_para_text(prev_child).strip()
                    if not prev_text:
                        continue
                    # 跳过表例段落
                    if _RE_TABLE_CAP.match(prev_text):
                        continue
                    return prev_text
            return ""
        else:  # below
            for j in range(tbl_idx + 1, min(tbl_idx + 15, len(body_children))):
                next_child = body_children[j]
                next_local = tag_local(next_child)
                if next_local == 'p':
                    next_text = get_para_text(next_child).strip()
                    if not next_text:
                        continue
                    # 跳过表例段落
                    if _RE_TABLE_CAP.match(next_text):
                        continue
                    return next_text
                elif next_local == 'tbl':
                    break
            return ""

    # ===== 提取位置信息 =====
    body = root.find(f'{{{W_NS}}}body')
    if body is None:
        body = root

    # 统计正文字号，用于区分表例/图例
    body_font_halfpt = _get_body_font_size(body)

    # 遍历body的直接子元素，记录表格位置
    # para_count 只统计 p 元素，不统计 tbl 元素，与 text_extract.py 的段落计数对齐
    table_positions = []
    para_count = 0
    current_heading = ''
    preceding_paras = []  # 最近5个段落文本，用于caption查找和上下文

    # 先收集所有body子元素信息，用于后续获取following_paragraph
    body_children = list(body)

    for idx, child in enumerate(body_children):
        local = tag_local(child)
        if local == 'p':
            ptext = get_para_text(child)
            # 检测标题（Heading样式）
            pPr = next((c for c in child if tag_local(c) == 'pPr'), None)
            if pPr:
                pStyle = next((c for c in pPr if tag_local(c) == 'pStyle'), None)
                if pStyle:
                    style_val = wattr(pStyle, 'val') or ''
                    if 'Heading' in style_val or 'heading' in style_val:
                        current_heading = ptext
            preceding_paras.append(ptext)
            if len(preceding_paras) > 5:
                preceding_paras.pop(0)
            para_count += 1
        elif local == 'tbl':
            # 向前查找caption：从preceding_paras中用格式特征+文本模式匹配
            table_caption = ''
            # 先用格式特征判断（优先）
            for ptext in reversed(preceding_paras):
                if _RE_TABLE_CAP.match(ptext.strip()):
                    table_caption = ptext.strip()
                    break

            # 获取紧邻前一段落文本（最后50字用于匹配定位）
            # 跳过表例段落，找最近的正文段落
            context_above = ''
            for ptext in reversed(preceding_paras):
                ptext_stripped = ptext.strip()
                if not ptext_stripped:
                    continue
                # 跳过表例段落
                if _RE_TABLE_CAP.match(ptext_stripped):
                    continue
                context_above = ptext_stripped[-50:] if ptext_stripped else ''
                break

            # 向后查找following段落：遍历后续body子元素，收集紧接的p元素文本
            following_paras = []
            for fi in range(idx + 1, len(body_children)):
                fchild = body_children[fi]
                flocal = tag_local(fchild)
                if flocal == 'p':
                    ftext = get_para_text(fchild)
                    following_paras.append(ftext)
                    if len(following_paras) >= 2:
                        break
                elif flocal == 'tbl':
                    break

            # 向后查找caption：如果前面没找到，用格式特征检查表格后面的段落。
            # Word中偶尔会出现“表格 → 正文短残段 → TableN表例”的顺序，
            # 因此允许跳过最多一个很短的正文续句继续查找显式表例。
            if not table_caption:
                skipped_short_body = False
                for fchild_idx in range(idx + 1, len(body_children)):
                    fchild = body_children[fchild_idx]
                    flocal = tag_local(fchild)
                    if flocal == 'p':
                        ftext = get_para_text(fchild).strip()
                        if not ftext:
                            continue
                        if _is_table_caption_para(fchild, body_font_halfpt) or _starts_table_caption_text(ftext):
                            table_caption = ftext
                            break
                        if (
                            not skipped_short_body
                            and _is_short_body_continuation(ftext)
                            and fchild_idx + 1 < len(body_children)
                            and tag_local(body_children[fchild_idx + 1]) == 'p'
                        ):
                            next_text = get_para_text(body_children[fchild_idx + 1]).strip()
                            if _starts_table_caption_text(next_text):
                                skipped_short_body = True
                                continue
                        break  # 非空非表例段落，停止
                    elif flocal == 'tbl':
                        break

            # 获取紧邻后一段落文本（前50字用于匹配定位）
            # 跳过表例段落，找最近的正文段落
            context_below = ''
            for fchild_idx in range(idx + 1, min(idx + 10, len(body_children))):
                fchild = body_children[fchild_idx]
                flocal = tag_local(fchild)
                if flocal == 'p':
                    ftext = get_para_text(fchild).strip()
                    if not ftext:
                        continue
                    # 跳过表例段落
                    if _is_table_caption_para(fchild, body_font_halfpt):
                        continue
                    context_below = ftext[:50] if ftext else ''
                    break
                elif flocal == 'tbl':
                    break

            # 提取表例段落。若已经通过隔一段正文找到了表例，则不再用紧邻扫描覆盖。
            legend_paragraphs = []
            if table_caption:
                legend_paragraphs = [table_caption]
            else:
                legend_paragraphs = _find_table_legend_paragraphs(body, idx, body_children, body_font_halfpt)

            # 组合完整 caption
            if table_caption and legend_paragraphs:
                caption_full = ' '.join(legend_paragraphs)
            elif table_caption:
                caption_full = table_caption
            elif legend_paragraphs:
                caption_full = ' '.join(legend_paragraphs)
            else:
                caption_full = ""

            # 增强上下文：完整正文段落文本
            context_above_text = _find_table_context_body(body, idx, body_children, 'above', body_font_halfpt)
            context_below_text = _find_table_context_body(body, idx, body_children, 'below', body_font_halfpt)

            table_positions.append({
                'paragraph_index': para_count,
                'current_heading': current_heading,
                'table_caption': table_caption,
                'legend_paragraphs': legend_paragraphs,
                'caption_full': caption_full,
                'context_above': context_above,
                'context_below': context_below,
                'context_above_text': context_above_text,
                'context_below_text': context_below_text,
            })
            preceding_paras = []

    # ===== 提取表格 =====
    tables = list(root.iter(f'{{{W_NS}}}tbl'))
    all_tables = []

    for ti, tbl in enumerate(tables):
        tblPr = next((c for c in tbl if tag_local(c) == 'tblPr'), None)
        table_props = {'borders': {}, 'width': None, 'cellMargins': {}, 'style': None}
        if tblPr:
            for prop in tblPr:
                local = tag_local(prop)
                if local == 'tblBorders':
                    table_props['borders'] = parse_borders(prop)
                elif local == 'tblW':
                    table_props['width'] = wattr(prop, 'w')
                elif local == 'tblCellMar':
                    for m in prop:
                        table_props['cellMargins'][tag_local(m)] = wattr(m, 'w')
                elif local == 'tblStyle':
                    table_props['style'] = wattr(prop, 'val')

        grid_cols = []
        tblGrid = next((c for c in tbl if tag_local(c) == 'tblGrid'), None)
        if tblGrid:
            for gc in tblGrid:
                if tag_local(gc) == 'gridCol':
                    grid_cols.append({'width_twips': int(wattr(gc, 'w') or '0')})

        rows_elems = list(tbl.iter(f'{{{W_NS}}}tr'))
        # 只取直接子tr（避免嵌套表格的tr）
        direct_trs = [c for c in tbl if tag_local(c) == 'tr']
        # 如果直接子tr不够，说明可能有嵌套结构，用iter
        trs_to_use = direct_trs if direct_trs else rows_elems

        rows_data = []
        for ri, tr in enumerate(trs_to_use):
            trPr = next((c for c in tr if tag_local(c) == 'trPr'), None)
            row_height = row_height_rule = None
            is_header = False
            if trPr:
                for prop in trPr:
                    local = tag_local(prop)
                    if local == 'trHeight':
                        row_height = wattr(prop, 'val')
                        row_height_rule = wattr(prop, 'hRule')
                    elif local == 'tblHeader':
                        is_header = True

            cells_elems = [c for c in tr if tag_local(c) == 'tc']
            cells_data = []
            col_pos = 0

            for ci, tc in enumerate(cells_elems):
                tcPr = next((c for c in tc if tag_local(c) == 'tcPr'), None)
                props = parse_cell_props(tcPr)

                paras_elems = [c for c in tc if tag_local(c) == 'p']
                paragraphs = [parse_paragraph(p) for p in paras_elems]
                full_text = ''.join(p['text'] for p in paragraphs)

                cells_data.append({
                    'text': full_text,
                    'gridSpan': props['gridSpan'],
                    'vMerge': props['vMerge'],
                    'col_start': col_pos,
                    'width': props['width'],
                    'shading': props['shading'],
                    'vAlign': props['vAlign'],
                    'borders': props['borders'],
                    'tcMar': props['tcMar'],
                    'paragraphs': paragraphs,
                })
                col_pos += props['gridSpan']

            rows_data.append({
                'row_height': row_height,
                'row_height_rule': row_height_rule,
                'is_header': is_header,
                'cells': cells_data
            })

        position = table_positions[ti] if ti < len(table_positions) else {}

        # 如果外部没找到caption，检查表格第一行是否为表例
        if not position.get('table_caption') and rows_data:
            first_row = rows_data[0]
            first_cell_text = first_row['cells'][0]['text'].strip() if first_row['cells'] else ''
            # 匹配"表X"/"Table X"/"Talbe X"(拼写错误)开头的文本
            if _RE_TABLE_CAP.match(first_cell_text) or re.match(r'^Talbe\s*\d', first_cell_text, re.IGNORECASE):
                position['table_caption'] = first_cell_text
                first_row['is_caption_row'] = True

        all_tables.append({
            'table_index': ti + 1,
            'position': position,
            'table_properties': table_props,
            'grid_cols': grid_cols,
            'rows': rows_data
        })

    output = {
        'source_file': os.path.basename(docx_path),
        'total_tables': len(all_tables),
        'tables': all_tables
    }
    return output


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_all_tables.py <docx_file> [output.json]")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'all_tables_complete.json'

    result = extract_tables(docx_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Extracted {result['total_tables']} tables from {result['source_file']}")
    for t in result['tables']:
        pos = t.get('position', {})
        print(f"  Table {t['table_index']}: {len(t['rows'])} rows, {len(t['grid_cols'])} cols")
        print(f"    Position: heading='{pos.get('current_heading','')}' caption='{pos.get('table_caption','')}'")
        print(f"    Context: prev='{pos.get('prev_text','')[:40]}' next='{pos.get('next_text','')[:40]}'")
