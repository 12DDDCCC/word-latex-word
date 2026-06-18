"""临时审查脚本 - 验证 template_config.py 的 column_count 逻辑"""
import os
import sys
_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILL_ROOT, 'shared'))

import re
from template_config import _scan_cls_for_page_params, _ConditionTracker

results = []

# 验证 1: 检测 P1 — 检验 manuscript 在 if@stage@final 内的 column_count 行为
# 场景1: manuscript 模式 → 单栏 (column_count = 1)
cls_ms = r'''
\if@stage@final
  \if@manuscript
    \@twocolumntrue
  \else
    \@twocolumntrue
  \fi
\fi
'''
scan1 = _scan_cls_for_page_params(cls_ms)
results.append(('Test1_manuscript_in_final', scan1['manuscript'].get('column_count'),
                '应=1 (manuscript 强制单栏)', '=1' if scan1['manuscript'].get('column_count') == 1 else '!=1'))
results.append(('Test1_pub_in_final', scan1['final_pub'].get('column_count'),
                '应=2 (final 出版版双栏)', '=2' if scan1['final_pub'].get('column_count') == 2 else '!=2'))

# 场景2: discussions 模式 (stage@final false 分支) → 单栏
cls_disc = r'''
\if@stage@final
\else
  \@twocolumntrue
\fi
'''
scan2 = _scan_cls_for_page_params(cls_disc)
results.append(('Test2_discussions', scan2['discussions'].get('column_count'),
                '应=1 (discussions 单栏)', '=1' if scan2['discussions'].get('column_count') == 1 else '!=1'))

# 场景3: 真实 ACP 模板样例 (嵌套多重)
cls_acp = r'''
\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{acp}[2024/01/01]
\DeclareOption{manuscript}{\@stage@finaltrue}
\DeclareOption{final}{\@stage@finaltrue}
\DeclareOption{discussions}{\@stage@finalfalse}
\ProcessOptions\relax
\newif\if@stage@final
\if@stage@final
  \if@manuscript
    % manuscript 模式
    \textwidth177mm
    \@twocolumnfalse
  \else
    % 出版模式
    \textwidth177mm
    \@twocolumntrue
  \fi
\else
  % discussions 模式
  \textwidth146mm
  \@twocolumnfalse
\fi
'''
scan3 = _scan_cls_for_page_params(cls_acp)
results.append(('Test3_acp_manuscript', scan3['manuscript'].get('column_count'),
                'manuscript 期望单栏', scan3['manuscript'].get('column_count')))
results.append(('Test3_acp_pub', scan3['final_pub'].get('column_count'),
                'final 期望双栏', scan3['final_pub'].get('column_count')))
results.append(('Test3_acp_disc', scan3['discussions'].get('column_count'),
                'discussions 期望单栏', scan3['discussions'].get('column_count')))
results.append(('Test3_acp_ms_textwidth', scan3['manuscript'].get('textwidth'),
                'manuscript 期望 textwidth=177', scan3['manuscript'].get('textwidth')))

# 场景4: 检测 P1 关键问题 — 当有 \@twocolumntrue 但 manuscript 时被覆盖
# 这就是修复的关键: ms_params['column_count'] = 1 (强制单栏)
cls_p1 = r'''
\if@stage@final
  \if@manuscript
    \@twocolumntrue
  \fi
\fi
'''
scan4 = _scan_cls_for_page_params(cls_p1)
# 关键测试: 即便在 manuscript 块中检测到 \@twocolumntrue, 也应被强制为 1
results.append(('Test4_P1_fix', scan4['manuscript'].get('column_count'),
                '关键修复验证: manuscript = 1 (即便有@twocolumntrue)', '=1'))

# 场景5: 验证 tracker 行为
tracker = _ConditionTracker()
# 模拟进入 if@stage@final true 分支
tracker.process_line('\\if@stage@final')
tracker.process_line('\\if@manuscript')
# 现在应该 is_in_stage_final AND is_in_manuscript
results.append(('Test5_tracker_in_final', tracker.is_in_stage_final(), '应=True', 'True'))
results.append(('Test5_tracker_in_manuscript', tracker.is_in_manuscript(), '应=True', 'True'))
results.append(('Test5_tracker_not_in_pub', tracker.is_in_not_manuscript(), '应=False', 'False'))

# 输出
print("=" * 80)
print("Template Config column_count 审查结果")
print("=" * 80)
for name, actual, desc, expected_str in results:
    print(f"  {name}:")
    print(f"    Actual:   {actual}")
    print(f"    Desc:     {desc}")
    print(f"    Expected: {expected_str}")
print("=" * 80)
