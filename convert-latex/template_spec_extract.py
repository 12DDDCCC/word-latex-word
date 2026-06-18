#!/usr/bin/env python3
"""模板规格提取

使用template-extract-lite提取模板规格，替代旧版generate_latex_file。
"""

import json
import shutil
import subprocess
import sys
from pathlib import Path

# 包内绝对导入（非Python包，不能使用相对导入）
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from spec_adapter import SpecAdapter
from skeleton_builder import _build_skeleton_from_spec, class_options_from_spec
from template_example_inference import infer_example_settings

# 语义关键词（基础集合，可从spec.section_commands动态扩展）
INTRO_KEYWORDS = {'introduction', '引言', '绪论', '1 introduction'}
CONCLUSION_KEYWORDS = {'conclusions', 'conclusion', '结论', '总结', 'concluding remarks'}
APPENDIX_KEYWORDS = {'appendix', 'appendices', '附录', 'supplementary'}
DECL_KEYWORD_MAP = {
    'codeavailability':     ['code availability', '代码可用性', 'code and data availability'],
    'dataavailability':     ['data availability', '数据可用性'],
    'sampleavailability':   ['sample availability', '样品可用性'],
    'competinginterests':   ['competing interests', '利益冲突', 'conflict of interest', 'conflicts of interest'],
    'authorcontribution':   ['author contribution', '作者贡献', 'author contributions'],
    'acknowledgements':     ['acknowledgements', 'acknowledgment', '致谢', 'acknowledgements'],
}


def _extend_keywords_from_spec(spec):
    """从spec.section_commands和spec.special_envs动态扩展关键词集合

    不同模板可能有不同的章节别名命令（如\orisection），其section_title
    提供了该命令对应的默认标题文本，可作为关键词补充。
    """
    if not spec:
        return INTRO_KEYWORDS, CONCLUSION_KEYWORDS, APPENDIX_KEYWORDS, DECL_KEYWORD_MAP

    intro_kw = set(INTRO_KEYWORDS)
    concl_kw = set(CONCLUSION_KEYWORDS)
    app_kw = set(APPENDIX_KEYWORDS)
    decl_kw = dict(DECL_KEYWORD_MAP)

    # 从section_commands扩展
    sec_cmds = spec.get('section_commands', {})
    for cmd_name, cmd_info in sec_cmds.items():
        if isinstance(cmd_info, dict):
            alias_of = cmd_info.get('alias_of', '')
            title = cmd_info.get('section_title', '').lower()
            if alias_of == 'introduction' and title:
                intro_kw.add(title)
            elif alias_of == 'conclusions' and title:
                concl_kw.add(title)

    # 从special_envs扩展
    special_envs = spec.get('special_envs', {})
    for env_name, env_info in special_envs.items():
        if isinstance(env_info, dict):
            title = env_info.get('section_title', '').lower()
            if title and env_name in decl_kw:
                decl_kw[env_name].append(title)
            elif title and env_name == 'appendix':
                app_kw.add(title)

    return intro_kw, concl_kw, app_kw, decl_kw


def _build_cls_from_ins(template_dir, output_dir):
    """Build missing .cls files from template .ins/.dtx sources into output_dir."""
    ins_files = sorted(template_dir.glob('*.ins'))
    if not ins_files:
        return []

    output_dir.mkdir(parents=True, exist_ok=True)
    build_dir = output_dir / '_template_class_build'
    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir.mkdir(parents=True, exist_ok=True)

    support_suffixes = {'.ins', '.dtx', '.drv', '.cfg', '.sty', '.bst'}
    for src in template_dir.iterdir():
        if src.is_file() and src.suffix.lower() in support_suffixes:
            shutil.copy2(str(src), str(build_dir / src.name))

    for ins in sorted(build_dir.glob('*.ins')):
        subprocess.run(
            ['latex', '-interaction=nonstopmode', ins.name],
            cwd=str(build_dir),
            capture_output=True,
            timeout=180,
            check=False,
        )

    generated = []
    for src in sorted(build_dir.glob('*.cls')):
        dst = output_dir / src.name
        shutil.copy2(str(src), str(dst))
        generated.append(dst)
    for src in sorted(build_dir.glob('*.bst')):
        dst = output_dir / src.name
        if not dst.exists():
            shutil.copy2(str(src), str(dst))
    shutil.rmtree(build_dir, ignore_errors=True)
    return generated


def _ensure_cls_files(template_dir, output_dir):
    cls_files = sorted(template_dir.glob('*.cls'))
    if cls_files:
        return cls_files
    cls_files = sorted(output_dir.glob('*.cls'))
    if cls_files:
        return cls_files
    return _build_cls_from_ins(template_dir, output_dir)


def _copy_support_files(template_dir, cls_path, out_subdir):
    support_files = []
    seen = set()
    for base in (template_dir, Path(cls_path).parent):
        if not base.exists():
            continue
        for fp in sorted(base.iterdir()):
            if not fp.is_file() or fp.suffix.lower() not in ('.cls', '.bst', '.cfg', '.sty'):
                continue
            if fp.name in seen:
                continue
            dst = out_subdir / fp.name
            if not dst.exists():
                shutil.copy2(str(fp), str(dst))
            seen.add(fp.name)
            support_files.append(fp.name)
    return support_files


def extract_template_spec(template_dir, journal, output_dir, config_mode=None):
    """使用template-extract-lite提取模板规格（替代旧版generate_latex_file）

    config_mode: 模板配置模式名 ('manuscript', 'final', 'discussions')
                 None 时默认 'manuscript'

    Returns:
        dict: 与旧版generate_latex_file()返回值同构，新增skeleton_info和spec字段
    """
    # 延迟导入：其他skill模块
    import sys
    SKILL_DIR = Path(__file__).resolve().parent.parent
    if str(SKILL_DIR / 'journal-template-extract') not in sys.path:
        sys.path.insert(0, str(SKILL_DIR / 'journal-template-extract'))
    if str(SKILL_DIR / 'template-extract-lite') not in sys.path:
        sys.path.insert(0, str(SKILL_DIR / 'template-extract-lite'))

    from template_extract_lite import extract_spec, derive_skeleton_info, derive_metadata_block
    from extract_template import generate_latex_file  # 旧版回退

    template_dir = Path(template_dir)
    output_dir = Path(output_dir)

    # 查找.cls文件
    cls_files = _ensure_cls_files(template_dir, output_dir)
    if not cls_files:
        print(f'WARNING: No .cls file found in {template_dir}, falling back to old extractor')
        return generate_latex_file(str(template_dir), journal, str(output_dir))

    cls_path = cls_files[0]
    print(f'  使用 template-extract-lite 提取: {cls_path.name}')

    # 1. 提取spec
    try:
        spec = extract_spec(str(cls_path), journal=journal)
    except Exception as e:
        print(f'WARNING: template-extract-lite failed: {e}, falling back to old extractor')
        return generate_latex_file(str(template_dir), journal, str(output_dir))

    # 2. 推导skeleton_info
    spec.setdefault('document_class', {}).setdefault('class_name', cls_path.stem)
    skeleton_info = derive_skeleton_info(spec, cls_path=str(cls_path))
    example_settings = infer_example_settings(template_dir, cls_path.stem, spec)
    if example_settings:
        skeleton_info.update({
            key: value for key, value in example_settings.items()
            if key in {
                'abstract_cmd', 'abstract_env', 'abstract_after_maketitle',
                'abstract_cmd_optional', 'keywords_cmd', 'keywords_env', 'bib_style',
            } and value not in ('', None)
        })
        citation_command = example_settings.get('citation_command')
        if citation_command:
            spec.setdefault('bibliography_format', {})['citation_command'] = citation_command
        if example_settings.get('bib_style'):
            spec.setdefault('bibliography_format', {})['bst_file'] = example_settings['bib_style']

    # 3. 推导layout_spec
    adapter = SpecAdapter(spec, skeleton_info, cls_path=str(cls_path),
                          template_dir=str(template_dir), config_mode=config_mode)
    layout_spec = adapter.to_layout_spec()

    # 4. 生成元数据区和骨架tex
    metadata_block = derive_metadata_block(spec, skeleton_info)
    dc = spec.get('document_class', {})
    cls_name = dc['class_name']
    # 只在cls声明了manuscript选项时才添加
    doc_options = class_options_from_spec(spec, journal, config_mode)
    opts = ','.join(doc_options)
    skeleton_tex = _build_skeleton_from_spec(spec, skeleton_info, cls_name, opts, metadata_block)

    # 5. 写入文件
    out_subdir = output_dir / f'{journal}_paper'
    out_subdir.mkdir(parents=True, exist_ok=True)

    tex_path = out_subdir / f'{journal}_paper.tex'
    tex_path.write_text(skeleton_tex, encoding='utf-8')

    spec_path = out_subdir / f'{journal}_template_spec.json'
    with open(spec_path, 'w', encoding='utf-8') as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    layout_spec_path = out_subdir / f'{journal}_layout_spec.json'
    with open(layout_spec_path, 'w', encoding='utf-8') as f:
        json.dump(layout_spec, f, ensure_ascii=False, indent=2)

    # 6. 复制支撑文件
    support_files = _copy_support_files(template_dir, cls_path, out_subdir)

    return {
        'tex_path': str(tex_path),
        'info_path': str(spec_path),
        'layout_spec_path': str(layout_spec_path),
        'output_dir': str(out_subdir),
        'journal': journal,
        'support_files': support_files,
        'layout_spec': layout_spec,
        'skeleton_info': skeleton_info,
        'spec': spec,
        'metadata_block': metadata_block,
        'doc_options': doc_options,
    }
