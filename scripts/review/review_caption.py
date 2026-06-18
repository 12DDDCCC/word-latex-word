"""临时审查脚本 - 验证 caption_utils.py 的转义逻辑"""
import os
import sys
_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILL_ROOT, 'shared'))

from caption_utils import clean_caption

results = []

# 验证 1: ~ 转义
r1 = clean_caption('A~B^2')
exp1 = r'A\textasciitilde{}B\textasciicircum{}2'
results.append(('Test1_A~B^2', r1 == exp1, repr(r1), repr(exp1)))

# 验证 2: 已有 test_orchestrator_contract.py:120 的测试用例
r2 = clean_caption(r'A\B_{2}')
exp2 = r'A\textbackslash{}B\_\{2\}'
results.append(('Test2_existing', r2 == exp2, repr(r2), repr(exp2)))

# 验证 3: 多个特殊字符
r3 = clean_caption(r'A\B_{2}&C#D%')
exp3 = r'A\textbackslash{}B\_\{2\}\&C\#D\%'
results.append(('Test3_multi', r3 == exp3, repr(r3), repr(exp3)))

# 验证 4: 单独 ~
r4 = clean_caption('A~B')
exp4 = r'A\textasciitilde{}B'
results.append(('Test4_tilde', r4 == exp4, repr(r4), repr(exp4)))

# 验证 5: 单独 ^
r5 = clean_caption('A^B')
exp5 = r'A\textasciicircum{}B'
results.append(('Test5_caret', r5 == exp5, repr(r5), repr(exp5)))

# 验证 6: 多重
r6 = clean_caption('A~~B^^C')
exp6 = r'A\textasciitilde{}\textasciitilde{}B\textasciicircum{}\textasciicircum{}C'
results.append(('Test6_multi_tilde', r6 == exp6, repr(r6), repr(exp6)))

# 验证 7: 空字符串
r7 = clean_caption('')
results.append(('Test7_empty', r7 == '', repr(r7), "''"))

# 验证 8: None
r8 = clean_caption(None)
results.append(('Test8_none', r8 == '', repr(r8), "''"))

# 验证 9: 普通文本不变
r9 = clean_caption('Hello World')
exp9 = 'Hello World'
results.append(('Test9_normal', r9 == exp9, repr(r9), repr(exp9)))

# 验证 10: $ 转义（要确认是否在 caption 中)
r10 = clean_caption('Price: $100')
exp10 = r'Price: \$100'
results.append(('Test10_dollar', r10 == exp10, repr(r10), repr(exp10)))

# 输出
print("=" * 80)
print("Caption 转义审查结果")
print("=" * 80)
all_pass = True
for name, passed, actual, expected in results:
    status = "PASS" if passed else "FAIL"
    if not passed:
        all_pass = False
    print(f"[{status}] {name}")
    if not passed:
        print(f"  Actual:   {actual}")
        print(f"  Expected: {expected}")
print("=" * 80)
print(f"总览: {sum(1 for _,p,_,_ in results if p)}/{len(results)} 通过")
print("=" * 80)

sys.exit(0 if all_pass else 1)
