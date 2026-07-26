---
name: mohen-archive-check
description: 检查墨痕教育开发线程中的变更是否需要归档，并将结果写入 /Users/xiaosheng/工作/全局工作区/40_事业与收入/墨痕教育-记忆库/。Use when the user says “检查归档”, asks whether a recent 墨痕教育 change should be archived, or after bug fixes, regex/template/parser/recognition changes, input-output behavior changes, archive-rule updates, testing-rule updates, or stable preference updates across 题目录入、格式转换、答案录入.
---

# Mohen Archive Check

## Overview

Use this skill only for 墨痕教育's memory and archive governance.

Keep all long-term records under the Storage Map paths below. Do not invent new storage roots.

The default command phrase is `检查归档`.

## Storage Map

- Long-term memory root: `/Users/xiaosheng/工作/全局工作区/40_事业与收入/墨痕教育-记忆库/`
- Skill path: `/Users/xiaosheng/工作/全局工作区/40_事业与收入/墨痕教育/.claude/skills/mohen-archive-check/`
- Project evidence archive: `/Users/xiaosheng/工作/全局工作区/40_事业与收入/墨痕教育/问题归档/`

Read [references/archive-templates.md](references/archive-templates.md) for exact file templates and naming rules.

## Decision Flow

### 1. Confirm whether this turn should be archived

Archive when any of the following is true:

- A bug was fixed.
- A regex, template, parser, recognition rule, or split/merge rule changed.
- The input/output behavior of `题目录入`, `格式转换`, or `答案录入` changed.
- The user updated archive rules, testing rules, workflow rules, or stable preferences.

Do not archive when:

- The turn is only brainstorming or casual discussion.
- No behavior, rule, or reusable preference changed.
- The user is asking a question with no resulting decision.

If this turn does not qualify, respond clearly with `本次可不归档` and give one short reason. Do not create empty records.

### 2. Classify the record

Choose exactly one primary class:

- `模块问题`: a change tied to `题目录入`, `格式转换`, or `答案录入`
- `全局规则`: a change to archive/test/process rules
- `工作偏好`: a stable preference, path rule, language preference, or command phrase

If a change touches more than one class, archive the primary class first and cross-reference the rest in `关联记录`.

### 3. Write immediately

When the current thread already contains enough facts, write the archive directly instead of stopping at advice.

Required write behavior:

1. Update the relevant `INDEX.md`.
2. Create the record file with the required naming rule.
3. If the class is `模块问题`, also check whether the project archive under `问题归档/` needs a matching evidence record or index update.

### 4. Output after archiving

Return a short checklist with:

- Archive result: `需要归档` or `本次可不归档`
- Class and destination path
- Files updated in the long-term memory library
- Whether the project evidence archive also needs updating

## Writing Rules

- Keep the long-term memory library focused on reusable summaries, rules, and preferences.
- Keep the project `问题归档/` focused on evidence: symptom, root cause, changed files, test samples, test commands, result.
- Use progressive disclosure everywhere:
  - L1: total index
  - L2: category or module index
  - L3: single record
  - L4: samples, commands, risks, linked entries
- Prefer concise Chinese by default unless the user asks otherwise.
- Never store secrets, tokens, or temporary chat noise.

## Example Triggers

- `检查归档`
- `这次答案录入的 bug 修好了，检查归档`
- `我刚改了格式转换的正则，看看要不要归档`
- `以后都默认中文说明，检查归档`
- `把不要放到 C 盘这条记住，检查归档`
