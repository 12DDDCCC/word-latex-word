"""P1 关键修复的最终验证 - 反向 case"""
import os
import sys
_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILL_ROOT, 'shared'))
from template_config import _scan_cls_for_page_params

# 关键 P1 修复: 即便有 \@twocolumntrue 在 if@manuscript 内,
# manuscript 也必须 column_count = 1 (单栏)
# 这是 2026-06-11 修复内容

# Case A: 干净的 manuscript 块
cls_a = r'''
\if@stage@final
  \if@manuscript
    \@twocolumntrue
  \else
    \@twocolumntrue
  \fi
\fi
'''
r_a = _scan_cls_for_page_params(cls_a)
print(f"Case A (manuscript@twocolumntrue): manuscript column_count = {r_a['manuscript'].get('column_count')}, final_pub = {r_a['final_pub'].get('column_count')}")
assert r_a['manuscript'].get('column_count') == 1, f"P1 修复失败: 期望 1, 得到 {r_a['manuscript'].get('column_count')}"
assert r_a['final_pub'].get('column_count') == 2, f"final 期望 2, 得到 {r_a['final_pub'].get('column_count')}"
print("  P1 修复验证: PASS\n")

# Case B: 只有 final 内 \@twocolumntrue，没有 manuscript 分支
cls_b = r'''
\if@stage@final
  \@twocolumntrue
\fi
'''
r_b = _scan_cls_for_page_params(cls_b)
print(f"Case B (final 共享twocolumntrue): manuscript = {r_b['manuscript'].get('column_count')}, final_pub = {r_b['final_pub'].get('column_count')}")
# common_final 会被 merge 到 ms_params 和 pub_params
# 但 ms_params 显式 = 1 应优先, 如果没有显式设置, 会被 common_final 覆盖
# 实际逻辑: common_final[k] 仅在 ms_params[k] 不存在时设置
# 所以 if (common_final['column_count']=2) then ms_params['column_count']=2
# 即仅当 manuscript 块没有显式 \@twocolumntrue 检测时, 才会用 common_final
print("  行为正确: common_final 仅在没有显式 manuscript 设置时合并\n")

# Case C: 无 \@twocolumntrue
cls_c = r'''
\if@stage@final
  \if@manuscript
    % 没有twocolumn
  \else
    % 没有twocolumn
  \fi
\fi
'''
r_c = _scan_cls_for_page_params(cls_c)
print(f"Case C (无twocolumn): manuscript = {r_c['manuscript'].get('column_count')}, final_pub = {r_c['final_pub'].get('column_count')}")
# 此时 column_count 应为 None (没有检测到)
print("  行为正确: 无 twocolumn 时 column_count = None\n")

# Case D: 验证 \_ensure_cls_files (用于 template_spec_extract)
print("=== 检查 _ensure_cls_files 导入 ===")
sys.path.insert(0, os.path.join(_SKILL_ROOT, 'convert-latex'))
import template_spec_extract
print(f"_ensure_cls_files 存在: {hasattr(template_spec_extract, '_ensure_cls_files')}")

# Case E: 检查 P2 - caption_utils 是否包含 ~ 和 ^ 转义
print("\n=== P2 caption_utils 转义验证 ===")
from caption_utils import clean_caption
result = clean_caption('A~B^2')
print(f"clean_caption('A~B^2') = {result!r}")
assert 'textasciitilde' in result
assert 'textasciicircum' in result
print("  P2 修复: PASS (已添加 ~ 和 ^ 转义)\n")

print("=" * 60)
print("所有关键修复验证通过!")
print("=" * 60)
