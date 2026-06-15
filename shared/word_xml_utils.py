"""Word XML工具 — 跨skill使用的Word XML命名空间操作

3个skill中重复的XML工具函数合并到此模块：
- citation-extract: tag_local, wattr, _iter_runs_recursive, _get_run_color/text/bold
- table-lossless-extract: tag_local, wattr, get_all_attrs
- citation-extract/cross_ref_builder: W_NS 常量
"""

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def tag_local(e):
    """提取XML元素的本地标签名（去除命名空间前缀）"""
    return e.tag.split('}')[1] if '}' in e.tag else e.tag


def wattr(e, name):
    """获取Word命名空间属性值"""
    return e.get(f'{{{W_NS}}}{name}')


def get_all_attrs(elem):
    """获取元素所有属性，去除命名空间前缀"""
    result = {}
    for key, val in elem.attrib.items():
        if '}' in key:
            result[key.split('}')[1]] = val
        else:
            result[key] = val
    return result


def iter_runs_recursive(elem):
    """递归遍历所有run，包括任何嵌套容器(hyperlink/sdt/等)"""
    for child in elem:
        ln = tag_local(child)
        if ln == 'r':
            yield child
        else:
            yield from iter_runs_recursive(child)


def get_run_color(r):
    """获取run的颜色属性"""
    for c in r:
        if tag_local(c) != 'rPr':
            continue
        for cc in c:
            if tag_local(cc) == 'color':
                return wattr(cc, 'val') or 'auto'
    return 'auto'


def get_run_text(r):
    """获取run中的文本内容"""
    return ''.join(t.text for t in r.iter(f'{{{W_NS}}}t') if t.text)


def get_run_bold(r):
    """获取run的粗体属性"""
    for c in r:
        if tag_local(c) != 'rPr':
            continue
        for cc in c:
            if tag_local(cc) == 'b':
                v = wattr(cc, 'val')
                return v != '0' if v is not None else True
    return False
