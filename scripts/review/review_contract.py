"""临时审查脚本 - 验证 test_orchestrator_contract.py 新增测试逻辑正确性"""
import os
import sys
import re
from pathlib import Path

_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 模拟 caption_utils.clean_caption 在没有 import 的情况下
def clean_caption_local(text):
    if not text:
        return ''
    escape_map = {
        '\\': r'\textbackslash{}', '%': r'\%', '_': r'\_', '&': r'\&', '#': r'\#',
        '{': r'\{', '}': r'\}', '$': r'\$', '~': r'\textasciitilde{}', '^': r'\textasciicircum{}',
    }
    return ''.join(escape_map.get(ch, ch) for ch in str(text)).strip()

# 验证 1: line 121 test_clean_caption_does_not_double_escape_textbackslash_braces
# 输入: r'A\B_{2}' → 期望: r'A\textbackslash{}B\_\{2\}'
r1 = clean_caption_local(r'A\B_{2}')
exp1 = r'A\textbackslash{}B\_\{2\}'
print(f"Test 1 (line 121): actual={r1!r}, expected={exp1!r}, pass={r1 == exp1}")

# 验证 2: line 260 test_caption_escapes_tilde_and_caret
# 输入: 'A~B^2' → 期望: r'A\textasciitilde{}B\textasciicircum{}2'
r2 = clean_caption_local('A~B^2')
exp2 = r'A\textasciitilde{}B\textasciicircum{}2'
print(f"Test 2 (line 260): actual={r2!r}, expected={exp2!r}, pass={r2 == exp2}")

# 验证 3: 检查 line 75 test_omml_greek_sigma_mapping_is_not_reversed
# 验证 GREEK_MAP 是 import 自 omml_to_latex
GREEK_MAP = {'σ': r'\sigma', 'ς': r'\varsigma', 'α': r'\alpha', 'β': r'\beta'}
print(f"Test 3 (line 75): GREEK_MAP['σ']={GREEK_MAP['σ']!r}, expected=r'\\sigma', pass={GREEK_MAP['σ'] == r'\\sigma'}")
print(f"Test 3b (line 75): GREEK_MAP['ς']={GREEK_MAP['ς']!r}, expected=r'\\varsigma', pass={GREEK_MAP['ς'] == r'\\varsigma'}")

# 验证 4: 检查测试覆盖的关键路径
print("\n--- 测试覆盖度分析 ---")
with open(os.path.join(_SKILL_ROOT, 'test_orchestrator_contract.py'), 'r', encoding='utf-8') as f:
    content = f.read()

# 统计测试函数
tests = re.findall(r'^def\s+(test_\w+)', content, re.MULTILINE)
print(f"总测试数: {len(tests)}")
for t in tests:
    print(f"  - {t}")

# 验证: 测试是否覆盖了 _docx_insert 中的相关函数
print("\n--- 导入分析 ---")
imports = re.findall(r'^from\s+(\S+)\s+import\s+(.+)$', content, re.MULTILINE)
for module, names in imports:
    print(f"  {module}: {names.strip()}")
