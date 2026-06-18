"""DOCX 结构审计：生成后检查浮动容器 / 关系悬空 / XML 合法性。

用于固化 P0 修复——主链路 ``_full_width_block_elements`` 已用连续分节符表示
跨栏浮动体，不再产出 ``w:tblpPr`` 浮动表格容器。本模块提供一个纯函数审计器，
供测试做回归护栏，也可在生成 DOCX 后做"打开前结构审计"。

设计原则：只读，无副作用，无外部依赖（仅 zipfile + lxml）。
"""
from __future__ import annotations

import zipfile
from pathlib import Path
from lxml import etree

# OOXML 命名空间
_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
# document.xml 中 r:embed / r:id 等属性使用 officeDocument relationships 命名空间
_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
# .rels 关系文件本身使用 package relationships 命名空间（注意与 _R 不同）
_PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships'
_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

# 主链路跨栏浮动体不再用 tblpPr；若审计发现 >0 即为回归信号。
_EXPECTED_TBLPPR_COUNT = 0


def audit_docx_structure(docx_path):
    """审计 docx 结构健康度，返回结构化问题报告。

    Args:
        docx_path: docx 文件路径（str 或 Path）。

    Returns:
        dict，字段含义：
        - ``xml_valid``: word/document.xml 能否被 lxml 解析
        - ``tblpPr_count``: 浮动表格定位容器数量（主链路应为 0）
        - ``docpr_ids``: 所有 wp:docPr 的 id
        - ``docpr_duplicate_ids``: 重复的 docPr id（应为空）
        - ``image_rel_count``: rels 中的图片关系数
        - ``dangling_rids_in_doc``: 文档引用但 rels 未注册的 rId（悬空引用）
        - ``dangling_image_rels``: rels 注册了图片但文档未引用（孤儿关系）
    """
    result = {
        'xml_valid': False,
        'tblpPr_count': 0,
        'skill_floating_tblpPr_count': 0,
        'non_skill_tblpPr_count': 0,
        'docpr_ids': [],
        'docpr_duplicate_ids': [],
        'image_rel_count': 0,
        'dangling_rids_in_doc': [],
        'dangling_image_rels': [],
    }

    docx_path = Path(docx_path)
    try:
        with zipfile.ZipFile(docx_path) as zf:
            doc_xml = zf.read('word/document.xml')
            try:
                rels_xml = zf.read('word/_rels/document.xml.rels')
            except KeyError:
                rels_xml = None
    except (FileNotFoundError, zipfile.BadZipFile, KeyError):
        return result

    try:
        root = etree.fromstring(doc_xml)
    except etree.XMLSyntaxError:
        return result
    result['xml_valid'] = True

    floating_tbls = root.findall(f'.//{{{_W}}}tbl')
    for tbl in floating_tbls:
        if tbl.find(f'{{{_W}}}tblPr/{{{_W}}}tblpPr') is None:
            continue
        result['tblpPr_count'] += 1
        desc = tbl.find(f'{{{_W}}}tblPr/{{{_W}}}tblDescription')
        desc_val = desc.get(f'{{{_W}}}val') if desc is not None else ''
        if str(desc_val or '').startswith('skill-pdf-float-wrap'):
            result['skill_floating_tblpPr_count'] += 1
        else:
            result['non_skill_tblpPr_count'] += 1

    docpr_ids = [
        el.get('id') for el in root.findall(f'.//{{{_WP}}}docPr') if el.get('id')
    ]
    result['docpr_ids'] = docpr_ids
    seen = set()
    for rid in docpr_ids:
        if rid in seen and rid not in result['docpr_duplicate_ids']:
            result['docpr_duplicate_ids'].append(rid)
        seen.add(rid)

    rel_ids, image_rels = _collect_relationships(rels_xml)
    result['image_rel_count'] = len(image_rels)

    referenced = _collect_referenced_rids(root)
    result['dangling_rids_in_doc'] = sorted(referenced - rel_ids)
    result['dangling_image_rels'] = sorted(image_rels - referenced)
    return result


def _collect_relationships(rels_xml):
    """从 document.xml.rels 解析所有关系 id 与图片关系 id。"""
    rel_ids = set()
    image_rels = set()
    if rels_xml is None:
        return rel_ids, image_rels
    try:
        rels_root = etree.fromstring(rels_xml)
    except etree.XMLSyntaxError:
        return rel_ids, image_rels
    for rel in rels_root.findall(f'{{{_PKG_R}}}Relationship'):
        rid = rel.get('Id')
        rtype = (rel.get('Type') or '').lower()
        if rid:
            rel_ids.add(rid)
        if 'image' in rtype:
            image_rels.add(rid)
    return rel_ids, image_rels


def _collect_referenced_rids(root):
    """收集 document.xml 中所有被引用的关系 id（r:embed / r:id / r:link）。"""
    referenced = set()
    for el in root.iter():
        for attr in ('embed', 'id', 'link'):
            val = el.get(f'{{{_R}}}{attr}')
            if val:
                referenced.add(val)
    return referenced


def is_clean(docx_path, allow_skill_floating=False):
    """便捷断言：docx 是否通过结构审计（无 tblpPr / 无重复 docPr / 无悬空 rId）。"""
    report = audit_docx_structure(docx_path)
    if allow_skill_floating:
        tblppr_ok = report['non_skill_tblpPr_count'] == 0
    else:
        tblppr_ok = report['tblpPr_count'] == _EXPECTED_TBLPPR_COUNT
    return (
        report['xml_valid']
        and tblppr_ok
        and not report['docpr_duplicate_ids']
        and not report['dangling_rids_in_doc']
    )
