#!/usr/bin/env python3
r"""
排版规格报告生成器
生成排版规格的JSON/Markdown可读报告
"""
from collections import OrderedDict

# 导入排版规格提取模块
import os, sys
_this_dir = os.path.dirname(os.path.abspath(__file__))
if _this_dir not in sys.path:
    sys.path.insert(0, _this_dir)
try:
    from layout_spec_extract import LayoutSpecExtractor
    _HAS_LAYOUT_EXTRACT = True
except ImportError:
    _HAS_LAYOUT_EXTRACT = False


def extract_layout_spec(cls_path, sty_paths=None):
    """从.cls/.sty文件提取完整排版规格"""
    try:
        ext = LayoutSpecExtractor(cls_path, sty_paths or [])
        spec = ext.extract_all()
        return _ordered_to_dict(spec)
    except Exception as e:
        print(f'[排版规格] 提取失败: {e}')
        return None


def _ordered_to_dict(obj):
    """递归转换OrderedDict为普通dict"""
    if isinstance(obj, OrderedDict):
        return {k: _ordered_to_dict(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_ordered_to_dict(i) for i in obj]
    elif isinstance(obj, dict):
        return {k: _ordered_to_dict(v) for k, v in obj.items()}
    return obj


def _write_layout_report(spec, class_name, report_path):
    """生成排版规格的Markdown可读报告"""
    lines = [f'# 排版规格报告: {class_name}.cls', '']
    base_size = spec.get('base_size', 10)
    lines.append(f'> 基准字号: {base_size}pt')
    lines.append('')

    # 页面布局
    layout = spec.get('page_layout', {})
    if layout:
        lines.append('## 1. 页面布局')
        pw = layout.get('paperwidth', '')
        ph = layout.get('paperheight', '')
        if pw and ph:
            lines.append(f'- 纸张: {pw} × {ph}')
        for k in ('textheight', 'textwidth', 'oddsidemargin', 'evensidemargin',
                   'topmargin', 'headheight', 'headsep', 'footskip', 'columnsep'):
            if k in layout:
                lines.append(f'- {k}: {layout[k]}')
        lines.append('')

    # 栏数
    cols = spec.get('columns', '')
    if cols:
        lines.append('## 2. 栏数')
        lines.append(f'- {cols}')
        lines.append('')

    # 字体族
    fonts = spec.get('fonts', {})
    if fonts:
        lines.append('## 3. 字体族')
        for k, v in fonts.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 论文标题
    title = spec.get('title', {})
    if title:
        lines.append('## 4. 论文标题')
        _emit_size_block_md(lines, title)
        lines.append('')

    # 作者
    author = spec.get('author', {})
    if author:
        lines.append('## 5. 作者')
        _emit_size_block_md(lines, author)
        if 'font_declaration' in author:
            lines.append(f'- 字体声明: `{author["font_declaration"]}`')
        lines.append('')

    # 摘要
    abstract = spec.get('abstract', {})
    if abstract:
        lines.append('## 6. 摘要')
        for k, v in abstract.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 章节标题
    headings = spec.get('headings', {})
    if headings:
        lines.append('## 7. 章节标题')
        for level, h in headings.items():
            parts = [f'字号: {h.get("size_name","?")}({h.get("size_pt","?")}pt)']
            parts.append(f'字重: {h.get("weight","normal")}')
            parts.append(f'字形: {h.get("shape","normal")}')
            if 'font_family' in h:
                parts.append(f'字体族: {h["font_family"]}')
            if 'alignment' in h:
                parts.append(f'对齐: {h["alignment"]}')
            if 'before_skip' in h:
                parts.append(f'前间距: {h["before_skip"]}')
            if 'after_skip' in h:
                parts.append(f'后间距: {h["after_skip"]}')
            lines.append(f'- **{level}**: {", ".join(parts)}')
        lines.append('')

    # 正文
    body = spec.get('body_text', {})
    if body:
        lines.append('## 8. 正文')
        for k, v in body.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # Caption/Table/Footnote/References
    for section_key, section_title in [
        ('caption', '9. Caption (图/表标题)'),
        ('table', '10. 表格'),
        ('footnote', '11. 脚注'),
        ('references', '12. 参考文献'),
    ]:
        sec = spec.get(section_key, {})
        if sec:
            lines.append(f'## {section_title}')
            for k, v in sec.items():
                lines.append(f'- {k}: {v}')
            lines.append('')

    # 编号格式
    numbering = spec.get('numbering', {})
    if numbering:
        lines.append('## 13. 编号格式')
        for k, v in numbering.items():
            lines.append(f'- {k}: `{v}`')
        lines.append('')

    # 间距
    spacing = spec.get('spacing', {})
    if spacing:
        lines.append('## 14. 间距')
        for k, v in spacing.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 页眉页脚
    hdr = spec.get('header_footer', {})
    if hdr:
        lines.append('## 15. 页眉页脚')
        for k, v in hdr.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 颜色
    colors = spec.get('colors', {})
    if colors:
        lines.append('## 16. 颜色')
        for k, v in colors.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 列表
    lists = spec.get('lists', {})
    if lists:
        lines.append('## 17. 列表')
        for k, v in lists.items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 特殊环境
    envs = spec.get('special_environments', {})
    if envs:
        lines.append('## 18. 特殊环境')
        for name, info in envs.items():
            entry = f'- **{name}**: '
            if isinstance(info, dict):
                parts = [f'{k}={v}' for k, v in info.items()]
                entry += ', '.join(parts)
            else:
                entry += str(info)
            lines.append(entry)
        lines.append('')

    # 自定义命令(前20个)
    cmds = spec.get('custom_commands', {})
    if cmds:
        lines.append('## 19. 自定义命令 (前20个)')
        count = 0
        for name, info in cmds.items():
            if count >= 20:
                break
            nargs = info.get('nargs', 0) if isinstance(info, dict) else 0
            body = info.get('body', '') if isinstance(info, dict) else str(info)
            lines.append(f'- \\{name} (参数: {nargs}): `{body[:80]}`')
            count += 1
        lines.append('')

    # 依赖宏包
    pkgs = spec.get('packages', [])
    if pkgs:
        lines.append('## 20. 依赖宏包')
        lines.append(', '.join(pkgs[:50]))
        lines.append('')

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


def _emit_size_block_md(lines, spec_dict):
    """输出字号/字重/字形/对齐信息(Markdown报告用)"""
    if spec_dict.get('size_name') or spec_dict.get('size_pt'):
        sn = spec_dict.get('size_name', '?')
        sp = spec_dict.get('size_pt', '?')
        lines.append(f'- 字号: {sn} ({sp}pt)')
    if 'weight' in spec_dict:
        lines.append(f'- 字重: {spec_dict["weight"]}')
    if 'shape' in spec_dict:
        lines.append(f'- 字形: {spec_dict["shape"]}')
    if 'alignment' in spec_dict:
        lines.append(f'- 对齐: {spec_dict["alignment"]}')
