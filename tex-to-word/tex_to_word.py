#!/usr/bin/env python3
r"""
LaTeX → Word 无损转换管道
核心思路:
  1. Pandoc基础转换(文字+章节+引用)
  2. 公式: 占位符 + latex_to_omml skill 直接插入 OMML
  3. 表格回填(TikZ → python-docx重建)

用法: python tex_to_word.py <tex文件> [-o 输出.docx] [--ref-doc 模板.docx]
"""
import os, sys, re, json, subprocess, tempfile, shutil
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 导入交叉引用构建器
_CROSSREF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'citation-extract')
if os.path.isdir(_CROSSREF_DIR):
    sys.path.insert(0, _CROSSREF_DIR)
try:
    from cross_ref_builder import detect_cite_style_from_tex
    _HAS_CROSSREF = True
except ImportError:
    _HAS_CROSSREF = False

# 从子模块导入所有公共函数（保持向后兼容）
from _bbl_parser import _parse_bbl, _parse_author_list, _format_author_list, _find_bbl_file
from _pandoc_prep import prepare_tex_for_pandoc
from _docx_insert import (
    postprocess_docx, embed_images_in_docx, add_tikz_table_to_docx,
    latex_to_omml, _restore_equation_numbers,
    _add_table_borders, _set_cell_border,
    _set_font, _apply_fonts_to_doc, apply_template_word_layout,
    apply_template_word_styles,
    restore_front_matter_from_tex,
)
from _tex_extraction import (
    extract_images_from_tex, extract_tikz_tables_from_tex, parse_tikz_table,
    extract_formulas_from_tex, extract_tables_from_tex,
    _parse_tabular, _split_cells, _collect_table_captions,
)
from _pdf_float_guidance import apply_pdf_float_guidance


def find_support_files(tex_dir):
    """查找tex文件所在目录的支撑文件"""
    files = {'cls': [], 'sty': [], 'bst': [], 'bib': [], 'bbl': [], 'cfg': []}
    for f in Path(tex_dir).iterdir():
        if not f.is_file():
            continue
        ext = f.suffix.lstrip('.').lower()
        if ext in files:
            files[ext].append(str(f))
    return files


def find_pandoc():
    """Find Pandoc from PATH or Windows application discovery."""
    def _env_path(name, *parts):
        value = os.environ.get(name)
        return (Path(value).joinpath(*parts) if value else None)

    def _usable(exe):
        try:
            result = subprocess.run(
                [str(exe), '--version'],
                capture_output=True, text=True, encoding='utf-8',
                errors='replace', timeout=10, check=False)
            return result.returncode == 0 and (result.stdout or '').lower().startswith('pandoc ')
        except PermissionError:
            if os.name == 'nt':
                return _powershell_can_run(exe)
            return False
        except Exception:
            return False

    def _powershell_candidates():
        candidates = []
        for name in ('pwsh', 'powershell'):
            found = shutil.which(name)
            if found:
                candidates.append(Path(found))
        system_powershell = _env_path(
            'SystemRoot', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
        if system_powershell:
            candidates.append(system_powershell)
        candidates.append(Path('powershell.exe'))
        return candidates

    def _ps_quote(value):
        return "'" + str(value).replace("'", "''") + "'"

    def _powershell_can_run(exe):
        script = '& ' + _ps_quote(exe) + ' --version'
        for powershell in _powershell_candidates():
            try:
                result = subprocess.run(
                    [str(powershell), '-NoProfile', '-Command', script],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='replace', timeout=30, check=False)
                if result.returncode == 0 and (result.stdout or '').lower().startswith('pandoc '):
                    return True
            except Exception:
                pass
        return False

    def _powershell_find_pandoc(root):
        if root is None:
            return None
        root_text = str(root)
        root_literal = root_text.replace("'", "''")
        if 'winget' in root_text.lower() and 'packages' in root_text.lower():
            script = (
                f"$root = '{root_literal}'; "
                "$packages = Get-ChildItem -LiteralPath $root -Directory -Filter 'JohnMacFarlane.Pandoc*' "
                "-ErrorAction SilentlyContinue; "
                "foreach ($package in $packages) { "
                "  $versions = Get-ChildItem -LiteralPath $package.FullName -Directory -Filter 'pandoc-*' "
                "  -ErrorAction SilentlyContinue | Sort-Object Name -Descending; "
                "  foreach ($version in $versions) { "
                "    $exe = Join-Path $version.FullName 'pandoc.exe'; "
                "    $exe; return "
                "  } "
                "} "
            )
        else:
            script = (
                f"$root = '{root_literal}'; "
                "Get-ChildItem -Path $root -Recurse -Filter pandoc.exe "
                "-ErrorAction SilentlyContinue | "
                "Select-Object -First 1 -ExpandProperty FullName"
            )
        for powershell in _powershell_candidates():
            try:
                result = subprocess.run(
                    [
                        str(powershell), '-NoProfile', '-Command',
                        script,
                    ],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='replace', timeout=30, check=False)
                source = result.stdout.strip().splitlines()
                if source:
                    candidate = source[0].strip()
                    if candidate:
                        return candidate
            except Exception:
                pass
        return None

    def _powershell_find_known_pandoc():
        script = (
            "$roots = @("
            "(Join-Path $env:ProgramFiles 'WinGet\\Packages'), "
            "(Join-Path $env:ProgramFiles 'Pandoc'), "
            "(Join-Path $env:LOCALAPPDATA 'Microsoft\\WinGet\\Packages')"
            "); "
            "foreach ($root in $roots) { "
            "  if (-not (Test-Path -LiteralPath $root)) { continue } "
            "  $packages = Get-ChildItem -LiteralPath $root -Directory -Filter 'JohnMacFarlane.Pandoc*' "
            "  -ErrorAction SilentlyContinue; "
            "  foreach ($package in $packages) { "
            "    $versions = Get-ChildItem -LiteralPath $package.FullName -Directory -Filter 'pandoc-*' "
            "    -ErrorAction SilentlyContinue | Sort-Object Name -Descending; "
            "    foreach ($version in $versions) { "
            "      $exe = Join-Path $version.FullName 'pandoc.exe'; "
            "      $exe; return "
            "    } "
            "  } "
            "  $direct = Join-Path $root 'pandoc.exe'; "
            "  if (Test-Path -LiteralPath $direct) { $direct; return } "
            "} "
        )
        for powershell in _powershell_candidates():
            try:
                result = subprocess.run(
                    [str(powershell), '-NoProfile', '-Command', script],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='replace', timeout=30, check=False)
                source = result.stdout.strip().splitlines()
                if source:
                    candidate = source[0].strip()
                    if candidate:
                        return candidate
            except Exception:
                pass
        return None

    def _find_under(root):
        try:
            root = Path(root)
        except TypeError:
            return None
        if not root.is_dir():
            return None
        root_text = str(root).lower()
        is_winget_packages = 'winget' in root_text and 'packages' in root_text
        names = {'pandoc.exe', 'pandoc.EXE', 'Pandoc.exe'}
        for dirpath, dirnames, filenames in os.walk(root, onerror=lambda exc: None):
            for filename in filenames:
                if filename in names or filename.lower() == 'pandoc.exe':
                    candidate = Path(dirpath) / filename
                    if _usable(candidate):
                        return str(candidate)
                    if is_winget_packages:
                        return str(candidate)
                    try:
                        exists = candidate.exists()
                    except OSError:
                        exists = False
                    if exists:
                        return str(candidate)
        return None

    env_pandoc = os.environ.get('PANDOC_EXE') or os.environ.get('PANDOC')
    if env_pandoc:
        return env_pandoc

    pandoc = shutil.which('pandoc')
    if pandoc:
        return pandoc
    if os.name == 'nt':
        try:
            result = subprocess.run(
                ['where.exe', 'pandoc'],
                capture_output=True, text=True, encoding='utf-8',
                errors='replace', timeout=10, check=False)
            candidates = [line.strip() for line in result.stdout.splitlines() if line.strip()]
            candidates.sort(key=lambda item: (
                'appdata\\local' in item.lower(),
                'program files' not in item.lower(),
            ))
            if candidates:
                return candidates[0]
        except Exception:
            pass
    path_entries = [
        Path(entry) for entry in os.environ.get('PATH', '').split(os.pathsep)
        if 'pandoc' in entry.lower()
    ]
    path_entries.sort(key=lambda item: (
        'appdata\\local' not in str(item).lower(),
        'program files' in str(item).lower(),
    ))
    for entry in path_entries:
        for exe_name in ('pandoc.exe', 'pandoc.EXE'):
            candidate = entry / exe_name
            try:
                exists = candidate.exists()
            except OSError:
                exists = False
            if _usable(candidate) or exists:
                return str(candidate)
    if os.name == 'nt':
        for powershell in _powershell_candidates():
            try:
                result = subprocess.run(
                    [
                        str(powershell), '-NoProfile', '-Command',
                        '(Get-Command pandoc -ErrorAction Stop).Source'
                    ],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='replace', timeout=10, check=False)
                source = result.stdout.strip()
                if source:
                    return source
            except Exception:
                pass
        known_pandoc = _powershell_find_known_pandoc()
        if known_pandoc:
            return known_pandoc
        try:
            home = Path.home()
        except RuntimeError:
            home = None
        roots = [
            _env_path('ProgramFiles', 'Pandoc'),
            _env_path('ProgramFiles', 'WinGet', 'Packages'),
            _env_path('LOCALAPPDATA', 'Microsoft', 'WinGet', 'Packages'),
            Path(__file__).resolve().parents[2] / 'test' / '_tmp' / 'tools' / 'pandoc',
        ]
        if home:
            roots[3:3] = [
                home / 'AppData' / 'Local' / 'Microsoft' / 'WinGet' / 'Packages',
            ]
        for root in (root for root in roots if root is not None):
            found = _powershell_find_pandoc(root)
            if found:
                return found
            found = _find_under(root)
            if found:
                return found
        userprofile = os.environ.get('USERPROFILE')
        fallback_roots = []
        program_files = os.environ.get('ProgramFiles') or r'C:\Program Files'
        fallback_roots.append(Path(program_files) / 'WinGet' / 'Packages')
        if userprofile:
            fallback_roots.append(
                Path(userprofile) / 'AppData' / 'Local' / 'Microsoft' / 'WinGet' / 'Packages'
            )
        package = 'JohnMacFarlane.Pandoc_Microsoft.Winget.Source_8wekyb3d8bbwe'
        for root in fallback_roots:
            for version in ('pandoc-3.10', 'pandoc-3.9.0.2', 'pandoc-3.9', 'pandoc-3.8'):
                candidate = root / package / version / 'pandoc.exe'
                if _usable(candidate):
                    return str(candidate)
        for root in fallback_roots:
            try:
                root = Path(root)
                if not root.is_dir():
                    continue
                for dirpath, dirnames, filenames in os.walk(root, onerror=lambda exc: None):
                    for filename in filenames:
                        if filename.lower() == 'pandoc.exe':
                            return str(Path(dirpath) / filename)
            except Exception:
                pass
    raise RuntimeError('Pandoc executable not found')


def _run_pandoc_subprocess(cmd):
    """Run Pandoc directly, falling back to PowerShell for WinGet-protected paths."""
    try:
        return subprocess.run(
            cmd, capture_output=True, text=True, encoding='utf-8', timeout=600)
    except PermissionError:
        if os.name != 'nt':
            raise
        cmdline = subprocess.list2cmdline([str(part) for part in cmd])
        cmd_result = subprocess.run(
            ['cmd', '/c', cmdline],
            capture_output=True, text=True, encoding='utf-8',
            errors='replace', timeout=600)
        if cmd_result.returncode == 0:
            return cmd_result
        powershell = (
            shutil.which('powershell')
            or str(Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32' / 'WindowsPowerShell' / 'v1.0' / 'powershell.exe')
        )
        def _ps_quote(value):
            return "'" + str(value).replace("'", "''") + "'"

        ps_cmd = '& ' + _ps_quote(cmd[0])
        if len(cmd) > 1:
            ps_cmd += ' ' + ' '.join(_ps_quote(part) for part in cmd[1:])
        return subprocess.run(
            [powershell, '-NoProfile', '-Command', ps_cmd],
            capture_output=True, text=True, encoding='utf-8',
            errors='replace', timeout=600)


def run_pandoc(tex_path, output_docx, ref_doc=None, bib_file=None):
    """执行Pandoc转换"""
    cmd = [
        find_pandoc(), tex_path,
        '-o', output_docx,
        '--from', 'latex',
        '--to', 'docx',
        '--standalone',
        '--mathml',
    ]

    if ref_doc and os.path.exists(ref_doc):
        cmd.extend(['--reference-doc', ref_doc])

    if bib_file and os.path.exists(bib_file):
        cmd.extend(['--bibliography', bib_file, '--citeproc'])

    print(f'[1/4] Running Pandoc: {" ".join(cmd)}')
    try:
        result = _run_pandoc_subprocess(cmd)
    except subprocess.TimeoutExpired:
        raise RuntimeError('Pandoc 超时（>600s）')
    if result.returncode != 0:
        print(f'  Pandoc stderr: {result.stderr}')
        if not os.path.exists(output_docx):
            raise RuntimeError(f'Pandoc failed: {result.stderr}')
    print(f'  Output: {output_docx}')
    return output_docx


def tex_to_word(
    tex_path,
    output_path=None,
    ref_doc=None,
    bib_file=None,
    config_mode=None,
    use_reference_doc_styles=None,
    use_pdf_float_guidance=True,
    use_pdf_float_wrap=None,
    use_pdf_float_reflow=False,
):
    """主管道: tex → word

    config_mode: 模板配置模式名 ('manuscript', 'final', 'discussions')
                 None 时默认 'manuscript'（经典版）
    """
    tex_path = os.path.abspath(tex_path)
    tex_dir = os.path.dirname(tex_path)
    if use_pdf_float_wrap is None:
        # 修改：所有模式都启用四周环绕，避免图片独占一页导致空白
        use_pdf_float_wrap = True

    if output_path is None:
        base = os.path.splitext(os.path.basename(tex_path))[0]
        output_path = os.path.join(os.getcwd(), f'{base}.docx')
    output_path = os.path.abspath(output_path)

    # 查找支撑文件
    support = find_support_files(tex_dir)
    if not bib_file:
        bib_file = support['bib'][0] if support['bib'] else None
    bbl_file = support['bbl'][0] if support['bbl'] else None

    print(f'Input: {tex_path}')
    print(f'Output: {output_path}')
    print(f'Bib: {bib_file}')
    print(f'BBL: {bbl_file}')
    if ref_doc:
        print(f'Reference DOCX: {ref_doc}')
    if use_reference_doc_styles is None:
        use_reference_doc_styles = bool(ref_doc and os.path.exists(ref_doc))
    print(f'Reference-doc styles primary: {use_reference_doc_styles}')

    # 自动检测引用样式
    cite_style = 'apa'
    if _HAS_CROSSREF:
        with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
            tex_content = f.read()
        cite_style = detect_cite_style_from_tex(tex_content)
        if bbl_file:
            parsed_cites = _parse_bbl(bbl_file).get('cite_map', {})
            if parsed_cites and all(re.fullmatch(r'\d+', str(value)) for value in parsed_cites.values()):
                cite_style = 'numbered'
        print(f'Cite style: {cite_style}')

    # 工作目录
    work_dir = tempfile.mkdtemp(prefix='tex2word_')
    print(f'Work dir: {work_dir}')

    try:
        # Step 1: 预处理tex
        print('\n[Step 1] Preprocessing LaTeX for Pandoc...')
        prepared_tex, _, tikz_tables, display_formula_data, cite_map = prepare_tex_for_pandoc(tex_path, work_dir, bbl_path=bbl_file)

        # 复制支撑文件到工作目录
        for f in support['cls'] + support['sty'] + support['bst'] + support['cfg']:
            shutil.copy2(f, work_dir)
        if bib_file:
            shutil.copy2(bib_file, work_dir)

        # Step 2: 提取公式、表格、图片和TikZ表格
        print('\n[Step 2] Extracting formulas, tables and images...')
        formulas = extract_formulas_from_tex(tex_path)
        tables = extract_tables_from_tex(tex_path)
        images = extract_images_from_tex(tex_path)
        print(f'  Formulas: {len(formulas)} (display: {sum(1 for f in formulas if f["type"]=="display")}, inline: {sum(1 for f in formulas if f["type"]=="inline")})')
        print(f'  Tables: {len(tables)}')
        print(f'  Images: {len(images)}')
        print(f'  TikZ Tables: {len(tikz_tables)}')
        if use_pdf_float_guidance:
            guidance_stats = apply_pdf_float_guidance(tex_path, images, tikz_tables)
            if guidance_stats.get("enabled"):
                print(
                    "  [pdf guidance] matched "
                    f"{guidance_stats.get('matched', 0)}/{guidance_stats.get('total', 0)} "
                    f"cross-column floats from {Path(guidance_stats.get('pdf', '')).name} "
                    f"(top={guidance_stats.get('top', 0)}, "
                    f"inline={guidance_stats.get('inline', 0)}, "
                    f"delay-after={guidance_stats.get('delayed', 0)})"
                )
            else:
                print(f"  [pdf guidance] skipped: {guidance_stats.get('reason')}")

        # 保存提取结果
        extract_json = os.path.join(work_dir, 'extract_info.json')
        with open(extract_json, 'w', encoding='utf-8') as f:
            json.dump({'formulas': formulas, 'tables': tables, 'images': images, 'tikz_tables': tikz_tables}, f, ensure_ascii=False, indent=2)

        # Step 3: Pandoc转换
        pandoc_bib = bib_file if not bbl_file else None
        print('\n[Step 3] Running Pandoc conversion...')
        pandoc_output = os.path.join(work_dir, 'pandoc_output.docx')
        run_pandoc(prepared_tex, pandoc_output, ref_doc=ref_doc, bib_file=pandoc_bib)

        # Step 4: 后处理
        print('\n[Step 4] Post-processing Word document...')
        shutil.copy2(pandoc_output, output_path)

        # 加载 layout_spec (从tex目录的template_spec.json)
        layout_spec = None
        spec_paths = (
            list(Path(tex_dir).glob("*_paper/*_layout_spec.json")) +
            list(Path(tex_dir).glob("*layout_spec.json")) +
            list(Path(tex_dir).glob("*_paper/*_template_spec.json")) +
            list(Path(tex_dir).glob("*template_spec.json"))
        )
        if spec_paths:
            try:
                layout_spec = json.loads(Path(spec_paths[0]).read_text(encoding='utf-8', errors='ignore'))
                print(f'  Loaded layout_spec from {spec_paths[0].name}')
            except Exception:
                layout_spec = None

        # 加载 style_spec (从CLS提取的页面几何信息)
        from _docx_insert import _load_template_word_style
        style_spec = _load_template_word_style(tex_path, config_mode=config_mode)

        postprocess_docx(
            output_path,
            display_formula_data,
            tables,
            cite_map=cite_map,
            cite_style=cite_style,
            bbl_path=bbl_file,
            layout_spec=layout_spec,
            style_spec=style_spec,
            tex_path=tex_path,
        )

        # Step 5: 嵌入图片
        print('\n[Step 4.5] Applying template page setup before asset insertion...')
        if use_reference_doc_styles:
            apply_template_word_layout(
                output_path, tex_path, config_mode=config_mode,
                layout_spec=layout_spec)
        else:
            apply_template_word_styles(
                output_path, tex_path, config_mode=config_mode, layout_spec=layout_spec)

        if images:
            print('\n[Step 5] Embedding images...')
            embed_images_in_docx(
                output_path, images, tex_dir,
                layout_spec=layout_spec, style_spec=style_spec)

        # Step 6: 添加TikZ表格
        if tikz_tables:
            print('\n[Step 6] Adding TikZ tables...')
            add_tikz_table_to_docx(
                output_path, tikz_tables, tex_dir,
                layout_spec=layout_spec, style_spec=style_spec)

        if use_reference_doc_styles:
            print('\n[Step 7] Preserving reference-doc Word styles...')
            print('  Skipped template-derived style backfill; reference DOCX remains primary.')
        else:
            print('\n[Step 7] Applying template-derived Word styles...')
            apply_template_word_styles(output_path, tex_path, config_mode=config_mode, layout_spec=layout_spec)

        # Step 8: Restore front matter (Correspondence etc.)
        print('\n[Step 8] Restoring front matter...')
        restore_front_matter_from_tex(output_path, tex_path)
        if not use_reference_doc_styles:
            print('  [front matter] reapplying template-derived styles...')
            apply_template_word_styles(
                output_path, tex_path, config_mode=config_mode,
                layout_spec=layout_spec)

        if use_pdf_float_reflow:
            print('\n[Step 9] PDF-guided post-render float reflow...')
            try:
                from _pdf_float_reflow import reflow_cross_column_floats
                reflow_stats = reflow_cross_column_floats(output_path)
                if reflow_stats.get("enabled"):
                    accepted = reflow_stats.get("accepted", [])
                    print(f'  [float reflow] accepted {len(accepted)} move(s)')
                    for item in accepted:
                        print(
                            '    '
                            f'page {item.get("blank_page")} blank={item.get("blank_pt")}pt, '
                            f'min-page={item.get("min_page")}, '
                            f'floats={item.get("float_count")}, '
                            f'align={item.get("alignment_improved")}, '
                            f'score {item.get("score_before")} -> {item.get("score_after")}'
                        )
                        if item.get("section"):
                            print(f'      section: {item.get("section")}')
                    if reflow_stats.get("debug_dir"):
                        print(f'  [float reflow] debug: {reflow_stats.get("debug_dir")}')
                else:
                    print(f'  [float reflow] skipped: {reflow_stats.get("reason")}')
            except Exception as exc:
                print(f'  [float reflow] failed: {exc}')

        if use_pdf_float_wrap:
            print('\n[Step 9.5] PDF-guided floating wrap for cross-column floats...')
            try:
                from _pdf_float_wrap import wrap_cross_column_floats
                source_pdf_path = Path(tex_path).with_suffix('.pdf')
                wrap_stats = wrap_cross_column_floats(
                    output_path,
                    guidance_items=[*images, *tikz_tables],
                    source_pdf_path=source_pdf_path if source_pdf_path.exists() else None,
                    force_all=True,
                    verify_render=False,
                    max_iterations=1,
                )
                if wrap_stats.get("enabled"):
                    accepted = wrap_stats.get("accepted", [])
                    print(f'  [float wrap] accepted {len(accepted)} container(s)')
                    for item in accepted:
                        captions = item.get("captions") or []
                        first = captions[0] if captions else ""
                        print(
                            '    '
                            f'pdf-page={item.get("source_page")}, '
                            f'floats={item.get("float_count")}, '
                            f'drawings={item.get("converted_drawings")}, '
                            f'score {item.get("score_before")} -> {item.get("score_after")}'
                        )
                        if first:
                            print(f'      first: {first[:90]}')
                    if wrap_stats.get("debug_dir"):
                        print(f'  [float wrap] debug: {wrap_stats.get("debug_dir")}')
                else:
                    print(f'  [float wrap] skipped: {wrap_stats.get("reason")}')
            except Exception as exc:
                print(f'  [float wrap] failed: {exc}')

        print(f'\nDone! Output: {output_path}')
        return output_path

    finally:
        debug_dir = os.path.join(os.path.dirname(output_path), '_tex2word_work')
        if os.path.exists(debug_dir):
            try:
                shutil.rmtree(debug_dir)
            except PermissionError:
                pass
        try:
            shutil.move(work_dir, debug_dir)
        except Exception:
            pass
        if os.path.exists(debug_dir):
            print(f'Debug files saved: {debug_dir}')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='LaTeX → Word 无损转换')
    parser.add_argument('tex_file', help='LaTeX源文件路径')
    parser.add_argument('-o', '--output', help='输出Word文件路径')
    parser.add_argument('--ref-doc', help='Word样式模板(.docx)')
    parser.add_argument('--bib', help='BibTeX参考文献文件(.bib)')
    parser.add_argument('--config-mode', default=None,
                        help='模板配置模式 (manuscript/final/discussions)')
    parser.add_argument('--template-style-backfill', action='store_true',
                        help='即使使用 --ref-doc，也继续用LaTeX模板反推样式覆盖Word样式')
    parser.add_argument('--no-pdf-float-guidance', action='store_true',
                        help='不使用已编译PDF中的图表位置来指导Word跨栏图表分页')
    parser.add_argument('--pdf-float-wrap', dest='pdf_float_wrap', action='store_true',
                        default=None,
                        help='Convert cross-column floats to editable square-wrap containers (default for final mode)')
    parser.add_argument('--no-pdf-float-wrap', dest='pdf_float_wrap', action='store_false',
                        help='Disable final-mode default PDF-guided square-wrap float handling')
    parser.add_argument('--pdf-float-reflow', action='store_true',
                        help='Render generated Word to PDF and try moving next-page cross-column floats upward to reduce large blank areas')
    args = parser.parse_args()

    reference_styles = None
    if args.ref_doc:
        reference_styles = not args.template_style_backfill
    tex_to_word(args.tex_file, args.output, args.ref_doc,
                bib_file=args.bib, config_mode=args.config_mode,
                use_reference_doc_styles=reference_styles,
                use_pdf_float_guidance=not args.no_pdf_float_guidance,
                use_pdf_float_wrap=args.pdf_float_wrap,
                use_pdf_float_reflow=args.pdf_float_reflow)
