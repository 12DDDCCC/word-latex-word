#!/usr/bin/env python3
"""LLM 确认模块 — 对各提取阶段的结果进行 LLM 审核

支持的确认阶段：
  - text: 文本提取（语义分类、标题/摘要/关键词/正文）
  - formula: 公式提取（OMML→LaTeX 转换正确性）
  - image: 图片提取（图例与图片对应）
  - table: 表格提取（表例与表格对应）
  - chem: 化学式下标（是否该加下标）

用法:
  from llm_verify import verify_stage
  result = verify_stage('text', extracted_data)
"""

import json, os, sys, re
from pathlib import Path

# 从 verify_extract 导入共享函数
from verify_extract import collect_chem_items

sys.stdout.reconfigure(encoding='utf-8')

# ── LLM 后端选择 ──────────────────────────────────────

def _call_llm(prompt, system_msg="You are a helpful academic paper assistant."):
    """调用 LLM，优先使用 Anthropic API，其次使用 Ollama"""
    # 方案1: Anthropic Claude API
    try:
        import anthropic
        client = anthropic.Anthropic()
        resp = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=system_msg,
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.content[0].text
    except Exception:
        pass

    # 方案2: Ollama 本地
    try:
        import urllib.request
        data = json.dumps({
            "model": "qwen2.5:7b",
            "prompt": prompt,
            "system": system_msg,
            "stream": False,
        }).encode('utf-8')
        req = urllib.request.Request(
            "http://localhost:11434/api/generate",
            data=data,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            return result.get('response', '')
    except Exception:
        pass

    # 方案3: 无 LLM 可用，返回跳过
    return None


# ── 各阶段确认 Prompt ──────────────────────────────────

def _build_text_verify_prompt(paragraphs):
    """构建文本提取确认 prompt"""
    lines = []
    lines.append("请审核以下从Word论文中提取的段落语义分类是否正确。")
    lines.append("对每个段落，检查 semantic_type 是否准确，如有错误请指出。")
    lines.append("")
    lines.append("语义类型定义：")
    lines.append("- title: 文章标题（首段、居中、加粗、长文本）")
    lines.append("- author: 作者姓名（标题后、居中、短文本）")
    lines.append("- affiliation: 作者机构（居中、含大学/研究所等关键词）")
    lines.append("- abstract: 摘要正文")
    lines.append("- abstract_label: 'Abstract' 标签（单独一行）")
    lines.append("- keywords: 关键词")
    lines.append("- heading: 章节标题")
    lines.append("- body: 正文段落")
    lines.append("- figure_caption: 图说明（以'图/Figure/Fig.'开头）")
    lines.append("- table_caption: 表说明（以'表/Table'开头）")
    lines.append("- reference: 参考文献标题")
    lines.append("- acknowledgement: 致谢")
    lines.append("- declaration: 声明（Code/Data availability等）")
    lines.append("- empty: 空段落")
    lines.append("- unknown: 未识别")
    lines.append("")
    lines.append("段落列表（仅列出非body/empty段落）：")

    for i, p in enumerate(paragraphs):
        st = p.get('semantic_type', '?')
        if st in ('body', 'empty'):
            continue
        txt = p['text'][:120].replace('\n', ' ')
        lines.append(f"  [{i}] semantic_type={st} | {txt}")

    lines.append("")
    lines.append("请以 JSON 格式输出审核结果：")
    lines.append('{"verified": true/false, "issues": [{"index": N, "current": "xxx", "correct": "yyy", "reason": "..."}]}')
    return '\n'.join(lines)


def _build_formula_verify_prompt(paragraphs_with_formula):
    """构建公式提取确认 prompt"""
    lines = []
    lines.append("请审核以下从Word论文OMML公式转换的LaTeX是否正确。")
    lines.append("检查：1) 数学符号是否正确 2) 上下标是否正确 3) 分数/积分/求和等结构是否正确")
    lines.append("")

    for i, p in enumerate(paragraphs_with_formula):
        txt = p['text'][:80]
        latex = p.get('latex', '')[:200]
        lines.append(f"[{i}] 原文: {txt}")
        lines.append(f"    LaTeX: {latex}")
        lines.append("")

    lines.append('请以JSON格式输出: {"verified": true/false, "issues": [{"index": N, "error": "...", "fix": "..."}]}')
    return '\n'.join(lines)


def _build_image_verify_prompt(images, captions):
    """构建图片提取确认 prompt"""
    lines = []
    lines.append("请审核以下图片与图说明的对应关系是否正确。")
    lines.append("")

    for i, (img, cap) in enumerate(zip(images, captions)):
        img_name = img.get('image_file', '?') if isinstance(img, dict) else str(img)
        cap_text = cap.get('text', '')[:100] if isinstance(cap, dict) else str(cap)[:100]
        lines.append(f"[{i}] 图片: {img_name}")
        lines.append(f"    图说明: {cap_text}")
        lines.append("")

    lines.append('请以JSON格式输出: {"verified": true/false, "issues": [{"index": N, "error": "..."}]}')
    return '\n'.join(lines)


def _build_table_verify_prompt(tables, captions):
    """构建表格提取确认 prompt"""
    lines = []
    lines.append("请审核以下表格与表说明的对应关系是否正确。")
    lines.append("")

    for i, (tbl, cap) in enumerate(zip(tables, captions)):
        tbl_rows = len(tbl.get('rows', [])) if isinstance(tbl, dict) else '?'
        cap_text = cap.get('text', '')[:100] if isinstance(cap, dict) else str(cap)[:100]
        lines.append(f"[{i}] 表格: {tbl_rows}行")
        lines.append(f"    表说明: {cap_text}")
        lines.append("")

    lines.append('请以JSON格式输出: {"verified": true/false, "issues": [{"index": N, "error": "..."}]}')
    return '\n'.join(lines)


def _build_chem_verify_prompt(chem_items):
    """构建化学式下标确认 prompt"""
    lines = []
    lines.append("请审核以下化学式/缩写的下标处理是否正确。")
    lines.append("规则：只有真正的化学式（如CO2→CO₂）才加下标；模型名/卫星名/缩写不加下标。")
    lines.append("")

    for i, item in enumerate(chem_items):
        original = item.get('original', '?')
        converted = item.get('converted', '?')
        action = item.get('action', '?')
        lines.append(f"[{i}] 原文: {original} → 转换: {converted} (动作: {action})")

    lines.append("")
    lines.append('请以JSON格式输出: {"verified": true/false, "issues": [{"index": N, "original": "xxx", "correct_action": "keep/subscript", "reason": "..."}]}')
    return '\n'.join(lines)


# ── 公共接口 ──────────────────────────────────────────

def verify_stage(stage, data, context=None):
    """对指定阶段的提取结果进行 LLM 确认

    Args:
        stage: 'text' | 'formula' | 'image' | 'table' | 'chem'
        data: 提取结果数据
        context: 可选的原始文档上下文

    Returns:
        dict: {
            'verified': bool,        # 是否全部通过
            'llm_available': bool,   # LLM 是否可用
            'issues': list,          # 发现的问题
            'raw_response': str,     # LLM 原始回复
        }
    """
    builders = {
        'text': _build_text_verify_prompt,
        'formula': _build_formula_verify_prompt,
        'image': _build_image_verify_prompt,
        'table': _build_table_verify_prompt,
        'chem': _build_chem_verify_prompt,
    }

    if stage not in builders:
        return {'verified': False, 'llm_available': False, 'issues': [f'未知阶段: {stage}'], 'raw_response': ''}

    prompt = builders[stage](data)
    system_msg = "You are an academic paper formatting expert. Reply in the requested JSON format only."

    response = _call_llm(prompt, system_msg)

    if response is None:
        # LLM 不可用，使用规则引擎进行基本验证
        return _rule_based_verify(stage, data)

    # 解析 LLM 回复
    try:
        # 提取 JSON
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            result = json.loads(json_match.group())
            return {
                'verified': result.get('verified', False),
                'llm_available': True,
                'issues': result.get('issues', []),
                'raw_response': response,
            }
    except json.JSONDecodeError:
        pass

    return {
        'verified': False,
        'llm_available': True,
        'issues': ['LLM 回复无法解析为 JSON'],
        'raw_response': response,
    }


def _rule_based_verify(stage, data):
    """当 LLM 不可用时，使用规则引擎进行基本验证"""

    issues = []

    if stage == 'text':
        # 检查：是否有 title、abstract
        types = set()
        for p in data:
            types.add(p.get('semantic_type', 'unknown'))

        if 'title' not in types:
            issues.append({'error': '未找到文章标题 (semantic_type=title)'})
        if 'abstract' not in types:
            issues.append({'error': '未找到摘要 (semantic_type=abstract)'})

        # 检查 figure_caption 是否以图/Figure 开头
        for i, p in enumerate(data):
            st = p.get('semantic_type')
            txt = p.get('text', '')
            if st == 'figure_caption' and not re.match(r'^\s*(图|Figure|Fig\.?)', txt, re.IGNORECASE):
                issues.append({'index': i, 'error': f'figure_caption 但文本不以图/Figure开头: {txt[:50]}'})
            if st == 'table_caption' and not re.match(r'^\s*(表|Table)', txt, re.IGNORECASE):
                issues.append({'index': i, 'error': f'table_caption 但文本不以表/Table开头: {txt[:50]}'})

    elif stage == 'chem':
        # 检查已知排除项
        _KNOWN_NON_CHEM = ['GCASv2', 'GOSAT2', 'OCO2', 'OCO3', 'MODIS', 'TCCON', 'MOPITT']
        for i, item in enumerate(data):
            original = item.get('original', '')
            converted = item.get('converted', '')
            for exc in _KNOWN_NON_CHEM:
                if exc in original and '$_{' in converted:
                    issues.append({
                        'index': i,
                        'original': original,
                        'correct_action': 'keep',
                        'reason': f'{exc} 是模型/卫星名，不应加下标',
                    })

    return {
        'verified': len(issues) == 0,
        'llm_available': False,
        'issues': issues,
        'raw_response': '(rule-based verification, LLM unavailable)',
    }


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='LLM 确认模块')
    parser.add_argument('stage', choices=['text', 'formula', 'image', 'table', 'chem'])
    parser.add_argument('json_file', help='提取结果 JSON 文件')
    args = parser.parse_args()

    with open(args.json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if args.stage == 'text':
        paragraphs = data.get('paragraphs', data) if isinstance(data, dict) else data
        result = verify_stage('text', paragraphs)
    elif args.stage == 'chem':
        tex_path = Path(args.json_file)
        if tex_path.suffix == '.tex':
            items = collect_chem_items(tex_path.read_text(encoding='utf-8'))
        else:
            items = data if isinstance(data, list) else data.get('chem_items', [])
        result = verify_stage('chem', items)
    else:
        result = verify_stage(args.stage, data)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result['verified']:
        print(f"\n发现 {len(result['issues'])} 个问题需要修复", file=sys.stderr)
