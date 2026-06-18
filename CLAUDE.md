**不要使用死编码，在进行word格式转换成latex格式以及latex格式转换成word格式时，需要从latex模板中进行提取格式要求才行**

## 多模板编译约束（避免不同执行环境产物不一致）

以下约束源于实测发现：同一 skill 在不同运行环境下（不同 Python 路径、不同 LLM 执行者）可能产出 documentclass/排版不一致的结果。

1. **config_mode 必须显式传递**：调用 `extract_template_spec` / `convert_direct` / orchestrator CLI 时，必须显式传 `--config-mode`（`manuscript`/`final`/`discussions`）。不要依赖 `None` 默认值——当 `config_mode` 为空时多处会静默回退到 `manuscript`，导致 elsarticle 等模板落回单栏 preprint，与双栏 final 产物不一致。

2. **elsarticle 双栏归一化不依赖 spec 完整性**：`_normalize_class_options_for_spec` 的 spec 为空回退分支也必须做 `twocolumn+manuscript → 3p` 归一化（已由 `_normalize_elsarticle_3p` 保障）。修改该函数时，确保两个分支（spec 齐全 / spec 为空）都经过 elsarticle 归一化，否则模板提取失败时产物会退回单栏。

3. **编译后检查 bbl URL 预警**：`_run_latex_steps` 在 bibtex 后会调用 `_warn_bbl_long_urls`。若输出 `[bib-url-warn]`，说明 references.bib 里存在超长（>70 字符）或重复等价 URL，会导致单栏超宽。根因在 bib 源数据，需清洗 references.bib 去重/缩短 URL，而非改 bst。

4. **引用风格由模板 bibstyle 决定，不需人为统一**：不同模板的引用形式（numeric vs author-year）由 `\bibliographystyle{}` 决定。例如 acp/elsarticle-harv 用 author-year，nsr 用 numeric。LaTeX 端和 Word 端各自遵循模板 bibstyle 即为正确行为，不要试图强制两端一致。

