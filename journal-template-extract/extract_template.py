#!/usr/bin/env python3
r"""
从LaTeX期刊模板(.cls/.sty/.tex/.bst)中提取写作结构
生成目标期刊的空白LaTeX写作文件

用法: python extract_template.py <模板目录> [期刊名称] [输出目录]
  模板目录: 包含.cls/.sty/.tex/.bst等文件的目录
  期刊名称: 目标期刊缩写(如acp, amt, bg等), 默认从.cfg推断
  输出目录: 生成文件的目录, 默认为模板目录的父目录
"""
import os, sys, json, shutil, re

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 确保当前目录在搜索路径中
_this_dir = os.path.dirname(os.path.abspath(__file__))
if _this_dir not in sys.path:
    sys.path.insert(0, _this_dir)

# 从子模块导入
from _template_scanner import find_template_files
from _cls_cfg_parser import parse_cls_file, parse_cfg_file, parse_template_tex
from _template_tex_builder import build_template_tex as _build_complete_tex
from _tex_transform import _write_full_spec
from _layout_report import extract_layout_spec, _write_layout_report


def _template_builder_spec(layout_spec, cls_info, target_journal, journal_options=None, tex_info=None):
    """Adapt legacy layout-spec output to build_template_tex()'s current schema."""
    spec = dict(layout_spec or {})

    document = dict(spec.get('document') or {})
    document.setdefault('documentclass', cls_info.get('class_name') or 'article')

    if not document.get('options'):
        options = []
        known_journal_options = set(journal_options or [])
        if target_journal and target_journal in known_journal_options:
            options.append(target_journal)
        document['options'] = options
    spec['document'] = document

    required_packages = dict(spec.get('required_packages') or {})
    for pkg_entry in cls_info.get('required_packages', []):
        for pkg_name in str(pkg_entry).split(','):
            pkg_name = pkg_name.strip()
            if pkg_name:
                required_packages.setdefault(pkg_name, None)
    for usepackage in (tex_info or {}).get('preamble_extra', []):
        m = re.search(r'\\usepackage\s*(?:\[[^\]]*\])?\{([^}]+)\}', usepackage)
        if not m:
            continue
        for pkg_name in m.group(1).split(','):
            pkg_name = pkg_name.strip()
            if pkg_name:
                required_packages.setdefault(pkg_name, None)
    spec['required_packages'] = required_packages

    if 'special_envs' not in spec:
        spec['special_envs'] = spec.get('special_environments') or {}

    return spec


def generate_latex_file(template_dir, journal_name=None, output_dir=None):
    """从模板提取信息并生成目标期刊的空白LaTeX文件"""
    files = find_template_files(template_dir)

    if not files['cls']:
        print(f'Error: No .cls file found in {template_dir}')
        return None

    cls_path = files['cls'][0]
    cls_info = parse_cls_file(cls_path)

    # 解析cfg获取期刊列表
    cfg_result = {'journals': {}, 'journal_options': []}
    if files['cfg']:
        for cfg in files['cfg']:
            r = parse_cfg_file(cfg)
            cfg_result['journals'].update(r['journals'])
            cfg_result['journal_options'].extend(r['journal_options'])

    # 合并cls和cfg的期刊选项
    all_journal_options = list(dict.fromkeys(
        cls_info['journals'] + cfg_result['journal_options']
    ))

    # 确定期刊名称
    if journal_name:
        target_journal = journal_name
    elif cfg_result['journals']:
        target_journal = list(cfg_result['journals'].keys())[0]
    elif all_journal_options:
        target_journal = all_journal_options[0]
    else:
        target_journal = cls_info['class_name']

    # 解析template.tex
    tex_info = None
    if files['tex']:
        tex_info = parse_template_tex(files['tex'][0])

    # 需要复制的支撑文件
    support_files = files['cls'] + files['sty'] + files['bst'] + files['cfg']

    # 生成输出目录
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(template_dir))
    output_dir = os.path.join(output_dir, f'{target_journal}_paper')
    os.makedirs(output_dir, exist_ok=True)

    # 复制支撑文件
    for src in support_files:
        fname = os.path.basename(src)
        dst = os.path.join(output_dir, fname)
        if not os.path.exists(dst):
            shutil.copy2(src, dst)

    # 排版规格提取 (先生成，供 .tex 文件使用)
    print('[排版规格] 提取中...')
    layout_spec = extract_layout_spec(cls_path, files.get('sty', []))

    # 解析 template.tex 结构 (如果有)
    # 生成完整 .tex 文件 - 强制包含所有模板要求的段落和排版规格注释
    print('[1/5] 生成完整 .tex 文件...')
    tex_path = os.path.join(output_dir, f'{target_journal}_paper.tex')
    builder_spec = _template_builder_spec(
        layout_spec, cls_info, target_journal, all_journal_options, tex_info
    )
    built_tex_path = _build_complete_tex(
        builder_spec,
        template_tex_path=files['tex'][0] if files['tex'] else None,
        output_dir=output_dir,
        journal_name=target_journal,
    )
    if os.path.abspath(built_tex_path) != os.path.abspath(tex_path):
        shutil.copy2(built_tex_path, tex_path)
    print(f'[1/5] → {tex_path}')

    # 生成结构信息JSON
    info = {
        'class_name': cls_info['class_name'],
        'version': cls_info['version'],
        'date': cls_info['date'],
        'target_journal': target_journal,
        'journals_from_cls': cls_info['journals'],
        'journals_from_cfg': cfg_result['journal_options'],
        'all_journal_options': all_journal_options,
        'custom_commands': cls_info['custom_commands'],
        'custom_environments': cls_info['custom_environments'],
        'support_files': [os.path.basename(f) for f in support_files],
    }

    if layout_spec:
        info['layout_spec'] = layout_spec
        # 生成独立排版规格文件
        print('[2/5] 生成排版规格 JSON...')
        spec_path = os.path.join(output_dir, f'{target_journal}_layout_spec.json')
        with open(spec_path, 'w', encoding='utf-8') as f:
            json.dump(layout_spec, f, ensure_ascii=False, indent=2)
        print(f'[2/5] → {spec_path}')

        # 生成可读报告
        print('[3/5] 生成排版规格报告...')
        report_path = os.path.join(output_dir, f'{target_journal}_layout_report.md')
        _write_layout_report(layout_spec, cls_info['class_name'], report_path)
        print(f'[3/5] → {report_path}')

        # 生成 Word 样式映射
        print('[4/5] 生成 Word 样式映射...')
        from layout_spec_extract import spec_to_word_styles
        word_styles = spec_to_word_styles(layout_spec)
        ws_path = os.path.join(output_dir, f'{target_journal}_word_styles.json')
        with open(ws_path, 'w', encoding='utf-8') as f:
            json.dump({'layout_spec': layout_spec, 'word_styles': word_styles},
                      f, ensure_ascii=False, indent=2)
        print(f'[4/5] → {ws_path}')

        # 生成完整规格文档 (v3.1新增)
        print('[5/5] 生成完整规格文档...')
        full_spec_path = os.path.join(os.path.dirname(output_dir), f'{target_journal}_full_spec.md')
        _write_full_spec(layout_spec, full_spec_path)
        print(f'[5/5] → {full_spec_path}')
    else:
        print('[2-5/5] 排版规格提取失败，跳过')

    info_path = os.path.join(output_dir, 'template_info.json')
    with open(info_path, 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False, indent=2)

    return {
        'tex_path': tex_path,
        'info_path': info_path,
        'output_dir': output_dir,
        'journal': target_journal,
        'support_files': [os.path.basename(f) for f in support_files],
        'layout_spec': layout_spec,
    }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_template.py <模板目录> [期刊名称] [输出目录]")
        print("  模板目录: 包含.cls/.sty/.tex/.bst等文件的目录")
        print("  期刊名称: 目标期刊缩写(如acp)")
        print("  输出目录: 生成文件的目录")
        sys.exit(1)

    template_dir = sys.argv[1]
    journal_name = sys.argv[2] if len(sys.argv) > 2 else None
    output_dir = sys.argv[3] if len(sys.argv) > 3 else None

    result = generate_latex_file(template_dir, journal_name, output_dir)
    if result:
        print(f'Generated journal template:')
        print(f'  Journal: {result["journal"]}')
        print(f'  TeX: {result["tex_path"]}')
        print(f'  Info: {result["info_path"]}')
        print(f'  Support files: {result["support_files"]}')
