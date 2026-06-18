"""Infer structural commands from the example TeX shipped with a template."""

import re
from pathlib import Path


def _without_comments(text):
    lines = []
    for line in text.splitlines():
        line = re.sub(r'(?<!\\)%.*$', '', line)
        if line.strip():
            lines.append(line)
    return '\n'.join(lines)


def _select_example(template_dir, class_name, spec):
    candidates = []
    class_pattern = re.compile(
        rf'\\documentclass(?:\[[^\]]*\])?\{{{re.escape(class_name)}\}}')
    for path in Path(template_dir).glob('*.tex'):
        text = path.read_text(encoding='utf-8', errors='ignore')
        if class_pattern.search(_without_comments(text)):
            candidates.append((path, text))
    if not candidates:
        return None, ''
    if spec.get('bibliography_format', {}).get('style') == 'natbib':
        harv = next((item for item in candidates if 'harv' in item[0].stem.lower()), None)
        if harv:
            return harv
    return candidates[0]


def _existing_bst_style(template_dir, example_text, spec):
    bst_files = list(Path(template_dir).glob('*.bst'))
    available = {path.stem: path for path in bst_files}
    match = re.search(r'\\bibliographystyle\{([^}]+)\}', example_text)
    if match and match.group(1) in available:
        return match.group(1)
    if spec.get('bibliography_format', {}).get('style') == 'natbib':
        harv = next((name for name in available if 'harv' in name.lower()), None)
        if harv:
            return harv
    return next(iter(available), '')


def infer_example_settings(template_dir, class_name, spec):
    """Return only signals demonstrated by a top-level template example."""
    path, raw_text = _select_example(template_dir, class_name, spec)
    if not path:
        return {}
    text = _without_comments(raw_text)
    settings = {'example_path': str(path)}

    abstract_cmd = re.search(r'\\abstract(?:\[([^\]]*)\])?\{', text)
    abstract_env = re.search(r'\\begin\{abstract\}', text)
    if abstract_cmd and (not abstract_env or abstract_cmd.start() < abstract_env.start()):
        settings['abstract_cmd'] = r'\abstract'
        if abstract_cmd.group(1):
            settings['abstract_cmd_optional'] = abstract_cmd.group(1).strip()
        abstract_pos = abstract_cmd.start()
    elif abstract_env:
        settings['abstract_env'] = 'abstract'
        abstract_pos = abstract_env.start()
    else:
        abstract_pos = -1

    trigger_positions = [
        match.start() for pattern in (r'\\maketitle\b', r'\\end\{frontmatter\}')
        for match in re.finditer(pattern, text)
    ]
    if abstract_pos >= 0 and trigger_positions:
        settings['abstract_after_maketitle'] = abstract_pos > min(trigger_positions)

    if re.search(r'\\keywords\s*\{', text):
        settings['keywords_cmd'] = r'\keywords'
    else:
        keyword_env = re.search(r'\\begin\{(keywords?)\}', text)
        if keyword_env:
            settings['keywords_env'] = keyword_env.group(1)

    if re.search(r'\\citep\s*\{', text):
        settings['citation_command'] = 'citep'
    elif re.search(r'\\cite\s*\{', text):
        settings['citation_command'] = 'cite'

    settings['bib_style'] = _existing_bst_style(template_dir, text, spec)
    return settings
