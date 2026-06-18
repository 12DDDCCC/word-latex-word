"""Caption utilities shared by extraction and LaTeX generation."""

import re


def normalize_caption(text):
    """Legacy normalizer kept for callers that explicitly request normalization."""
    if not text:
        return ''
    t = re.sub(r'Talbe', 'Table', str(text), flags=re.IGNORECASE)
    t = re.sub(r'^表\s*(\d+)', r'Table \1', t)
    t = re.sub(r'^图\s*(\d+)', r'Figure \1', t)
    t = re.sub(
        r'^(Table|Figure|Fig\.)\s*(\d+(?:\.\d)?)[\s:：,，.、]*',
        r'\1 \2: ',
        t,
        flags=re.IGNORECASE,
    )
    return t


def clean_caption(text):
    """Escape LaTeX special characters without rewriting caption wording."""
    if not text:
        return ''
    escape_map = {
        '\\': r'\textbackslash{}',
        '%': r'\%',
        '_': r'\_',
        '&': r'\&',
        '#': r'\#',
        '{': r'\{',
        '}': r'\}',
        '$': r'\$',
        '~': r'\textasciitilde{}',
        '^': r'\textasciicircum{}',
    }
    return ''.join(escape_map.get(ch, ch) for ch in str(text)).strip()


def strip_caption_prefix(caption):
    """Remove only a leading visible source caption number for \\caption{}."""
    if not caption:
        return caption
    patterns = [
        r'^\s*图\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
        r'^\s*表\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
        r'^\s*Figure\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
        r'^\s*Fig\.?\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
        r'^\s*Table\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
        r'^\s*Talbe\s*\d+(?:\.\d+)*\s*[\.:：,，、]*\s*',
    ]
    text = str(caption)
    for pat in patterns:
        text = re.sub(pat, '', text, flags=re.IGNORECASE)
    return text.strip()


def clean_caption_prefix_in_tex(tex_content):
    """Remove duplicated source numbers inside already generated \\caption{} text."""
    patterns = [
        r'(\\caption\{)\s*Table\s+\d+(?:\.\d+)?[\.:：,，、]\s*',
        r'(\\caption\{)\s*Figure\s+\d+(?:\.\d+)?[\.:：,，、]\s*',
        r'(\\caption\{)\s*表\s*\d+(?:\.\d+)?[\.:：,，、]\s*',
        r'(\\caption\{)\s*图\s*\d+(?:\.\d+)?[\.:：,，、]\s*',
    ]
    for pat in patterns:
        tex_content = re.sub(pat, r'\1', tex_content, flags=re.IGNORECASE)
    return tex_content
