#!/usr/bin/env python3
r"""
模板文件扫描器
扫描目录中的LaTeX模板文件(.cls/.sty/.tex/.bst/.cfg)
"""
from pathlib import Path

# 不应作为顶级section的环境(表格/公式内部组件)
_SKIP_ENVS = {
    'tabular', 'tabularx', 'longtable', 'array', 'cases',
    'matrix', 'pmatrix', 'bmatrix', 'vmatrix', 'aligned',
    'gathered', 'split', 'multline', 'equation', 'align',
    'alignat', 'flalign', 'gather', 'subequations',
    'algorithmic', 'algorithmicx', 'lstlisting',
    'minipage', 'column', 'cell', 'row',
}

# 标准LaTeX选项(非期刊选项)
_STD_OPTS = {
    'a4paper', 'a5paper', 'b5paper', 'letterpaper', 'legalpaper',
    'executivepaper', '10pt', '11pt', '12pt', 'oneside', 'twoside',
    'draft', 'final', 'fleqn', 'leqno', 'titlepage', 'notitlepage',
    'openright', 'openany', 'twocolumn', 'onecolumn', 'landscape',
    'portrait', 'leqno', 'fleqn', 'openbib',
}

# 排版/格式选项(非期刊选项)
_FORMAT_OPTS = {
    'manuscript', 'proof', 'online', 'noline', 'nohyperref',
    'debug', 'noref', 'noauthor', 'nolastpage', 'forHTML',
    'hvmath', 'corrigendum', 'editorialnote', 'editorialnotediscussion',
    'preface', 'screen', 'print', 'submit', 'preprint',
    'review', 'doubleblind', 'singleblind', 'anonymous',
    'lineno', 'nolineno', 'number', 'nonumber',
    'endnote', 'hyperref', 'nohyperref',
    'proofreadingchanges', 'copyediting', 'smsps',
}


def find_template_files(directory):
    """扫描目录中的LaTeX模板文件"""
    files = {
        'cls': [], 'sty': [], 'bst': [], 'tex': [],
        'cfg': [], 'other': []
    }
    for f in Path(directory).rglob('*'):
        if f.is_file():
            ext = f.suffix.lower()
            name = f.name.lower()
            if ext == '.cls':
                files['cls'].append(str(f))
            elif ext == '.sty':
                files['sty'].append(str(f))
            elif ext == '.bst':
                files['bst'].append(str(f))
            elif ext == '.tex' and ('template' in name or 'sample' in name or 'example' in name):
                files['tex'].append(str(f))
            elif ext == '.cfg':
                files['cfg'].append(str(f))
            else:
                if ext not in ('.zip', '.docx', '.pdf', '.log', '.aux',
                               '.bbl', '.blg', '.synctex.gz', '.fdb_latexmk',
                               '.fls', '.4ct', '.4tc', '.out', '.tmp', '.xref',
                               '.gz', '.tar', '.dvi', '.ps', '.xdv'):
                    files['other'].append(str(f))
    return files
