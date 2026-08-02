# Windows 代码包 ZIP 中文文件名乱码

## L1 摘要

- 模块：题目录入 / Windows 跨机交付
- 日期：2026-08-02
- 状态：已修复并验证
- 结论：macOS `/usr/bin/zip` 生成的包未给中文条目写入 UTF-8 标记，文件名在压缩包内已经乱码；改用 Python `zipfile` 并校验 UTF-8 标记后恢复正常。

## L2 现象与根因

- 用户在 Windows 解压 `墨痕教育-系统代码优化-目录学科路由-20260802-Windows.zip` 后看到中文文件名乱码。
- 检查压缩包发现中文条目名称已经变成乱码，且 `flag_bits=0x0`，没有 ZIP UTF-8 标记。
- 这是打包阶段的文件名编码问题，不是 Python 源码内容或 Windows WPS 文档乱码。

## L3 改动与验证

- 新包使用纯英文包名和顶层目录名，由 Python `zipfile` 写入。
- 最终包共 268 个条目，其中 157 个中文条目均带 UTF-8 标记 `0x800`；CRC 检查无坏文件，并确认 `墨痕快刀/wps_helper.py` 以原名存在。
- 解压后核心定向测试 33 项通过；核心 Python 文件编译通过；包内不含 `.git`、`.venv`、`.env`、缓存、认证文件或 sqlite 状态库。
- 最终包：`mohen-windows-wps-bound-plan-20260802-utf8.zip`
- SHA256：`8f95de8e2bc5f25f1ce05f90631464fc6f0e12cae3c2bc43c1258606fbc342a5`

## L4 边界与复用

- 以后从 macOS 交付包含中文路径的 Windows ZIP，必须在交付前检查中文条目的 UTF-8 标记，并执行一次实际解压验收。
- 旧乱码包应视为废弃，不作为测试输入。
- 该验证证明压缩包结构、文件名和离线测试正常，不代替 Windows/WPS 实机选区验收。
