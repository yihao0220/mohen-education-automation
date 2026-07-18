# 回归样本说明

这里不直接存放大体积题目/答案原文件，真实样本文档仍保留在本地业务目录：

- `格式处理/待清洗文件/`
- `格式处理/原格式/`
- `墨痕快刀/待录入文档/`

仓库只跟踪本说明文件。以下资产可以在本机复现，但含真实文档派生事实，均被 `.gitignore` 排除，不会随 Git 克隆或拉取：

1. `样本清单.json`
   - 记录本地样本的文件名、类别、学科猜测、风险标签
   - 作为本机“当前样本全集”的轻量索引
2. `预检基线/`
   - 在本机保存三类真实样本的 `DocumentProfile.json`、`ActionPlan.json` 及对应 Markdown
   - 当前 `DocumentProfile` schema 为 1.1，包含段落角色、置信度与命中证据
   - JSON 是机器权威结果，Markdown 是由 JSON 生成的人工审核视图
   - 当前角色层固定 `automatic_exclusion_enabled=false`，动作计划固定 `execution_enabled=false`；只验证结构、题块和原始坐标，不执行 WPS 按键
3. `文档族分析/`
   - 在本机保存由三份现有 Profile 生成的 `DocumentFamilyReport.json/.md`
   - 当前三份样本跨项目、跨学科，保守结果为 3 个单例异常候选，只用于验证兼容性和安全边界
   - 报告固定 `classification_mode=advisory_only`、`automatic_rule_binding_enabled=false`、`production_execution_enabled=false`
4. P1b 页面视觉产物
   - PDF、页面 PNG、页面清单、视觉审核、校准报告和只读批次清单必须写在原题目录与 Git 仓库之外
   - Mac 连续预览固定 `page_truth_authority=false`；Windows WPS COM 或显式 `--attest-wps-export` 的 WPS PDF 才可提供生产页面真值
5. 其他预期结果文件
   - 比如某份卷子应该切出多少题块
   - 某份答案是否应拆小题
   - 某份样本有哪些已知风险标签

刷新样本清单命令：

```powershell
python tools/build_regression_manifest.py
```

刷新单份只读预检基线：

```powershell
.\.venv\Scripts\python.exe .\tools\build_document_preflight.py "D:\...\原题.docx" --output-dir ".\回归样本\预检基线"
```

预检实现入口为 `shared_core/document_preflight.py`，角色证据入口为 `shared_core/document_roles.py`，验证入口为 `test_document_preflight.py`。所有输出必须位于原题之外，并在生成前后核对源文件 SHA256。

刷新 P1a 文档族报告：

```powershell
.\.venv\Scripts\python.exe .\tools\analyze_document_families.py .\回归样本\预检基线 --output-dir .\回归样本\文档族分析
```

分族实现入口为 `shared_core/document_families.py`，验证入口为 `test_document_families.py`。现有三份跨项目基线不能用于校准生产阈值；下一步必须使用同一真实项目批次复核。

新机器克隆仓库后不会自动获得上述真实基线。需要从本机业务目录重新生成，并先确认输出目录位于原题和仓库之外；不得把真实 DOCX、PDF、PNG 或派生 JSON 提交到 GitHub。
