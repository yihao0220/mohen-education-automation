# 删除废弃 DeepSeek/AI 答案清洗旁路

## 现象

项目仍保留两个已经不属于生产工作流的 AI 答案清洗入口：`格式处理/format_answers_ai.py` 和 `格式处理/format_answers_deepseek.py`。后者仅返回停用提示，前者是独立脚本，均未被当前 `格式处理/main.py`、审核门禁或答案录入主链引用。

## 根因

早期曾尝试使用 DeepSeek 直接参与答案清洗；后来生产链转为确定性模板、共享 `AnswerUnit`、题答映射与审核门禁，但为兼容历史调用继续保留了旧脚本和一条兼容测试，项目技术栈说明也没有同步清理。

## 改动点

- 删除 `格式处理/format_answers_ai.py`；
- 删除 `格式处理/format_answers_deepseek.py`；
- 删除 `test_low_risk_cleanup.py` 中只验证停用兼容入口的测试；
- 更新 `AGENTS.md`，明确当前生产链不依赖外部 AI 服务；
- 删除 `格式处理/common.py` 中未使用的 `llm` 日志级别；
- 在历史更新日志顶部补充当前状态，不改写既有历史记录。

## 测试样例

- 运行低风险清理定向测试，确认其余共享模板和字体行为不变；
- 运行格式处理与共享内核相关离线测试；
- 全仓检索生产代码，确认不存在旧模块导入、API 地址、环境变量或调用入口；
- 检查 Git diff，确认没有改动题目解析、答案清洗模板、审核门禁或 WPS 执行逻辑。

## 测试命令

```bash
python -m unittest -q test_low_risk_cleanup.py
python -m pytest -q -p no:cacheprovider test_shared_core_flow.py test_low_risk_cleanup.py
rg -n -i "format_answers_ai|format_answers_deepseek|deepseek|siliconflow|SILICONFLOW_API_KEY|DEEPSEEK_API_KEY" .
git diff --check
```

## 结论

本次只删除废弃旁路，不改变当前三段生产流程：题目录入、答案清洗、答案录入。历史归档和旧版本更新日志保留，仅用于说明过去发生过什么，不参与运行。

## 下次排查入口

若未来重新引入 AI，只能从新的结构化建议契约和影子验证开始，不得恢复或复制本次删除的旧脚本。先检查主规格 `docs/superpowers/specs/2026-08-20-mohen-intelligence-release-execution-spec.md` 中的智能参与边界。
