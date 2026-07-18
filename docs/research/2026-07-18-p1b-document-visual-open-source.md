# P1b 文档视觉链路 GitHub 开源项目调研

> 日期：2026-07-18  
> 目的：为“WPS 导出 PDF → 逐页 PNG → 页面区域 JSON → 视觉角色审核”选择最小依赖。  
> 结论：P1b 直接复用 `pdfplumber + pypdfium2 + Pillow`，保留 `pdftoppm` 兼容路径；WPS COM 只借鉴安全调用方式；暂不引入大型 OCR/布局模型。

## 1. 能够直接套用的

### [pdfplumber](https://github.com/jsvine/pdfplumber)

- 许可证：MIT。
- 能力：提取机器生成 PDF 的文字、图片、线条和边界框，也提供 JSON 对象输出和视觉调试能力。
- 本地状态：Codex bundled runtime 已安装 `pdfplumber 0.11.9`。
- P1b 用法：读取 WPS 导出的 PDF，生成页面尺寸、文字行、图片对象和 PDF 坐标。
- 边界：它适合机器生成 PDF；对图片内部文字不做 OCR，图片角色仍需视觉审核或可选 OCR 增强。

### [pypdfium2](https://github.com/pypdfium2-team/pypdfium2)

- 许可证：Apache-2.0 OR BSD-3-Clause；底层 PDFium 使用 BSD 风格许可证。
- 能力：提供 Windows、macOS 和 Linux 的预编译包，可直接把 PDF 页面渲染为 PIL 图片。
- P1b 用法：作为默认逐页 PNG 渲染器，避免 Windows 新机额外配置 Poppler 可执行文件。

### [pdf2image](https://github.com/Belval/pdf2image) / Poppler

- `pdf2image` 许可证：MIT；它封装 `pdftoppm` 和 `pdftocairo`。
- 能力：稳定地把 PDF 页面转成 PIL 图片，支持 DPI、页码、输出目录、超时和大文件路径模式。
- 本地状态：Mac bundled runtime 已提供 `pdftoppm` 和 `pdfinfo`。
- P1b 用法：作为显式指定或缺少 `pypdfium2` 时的兼容路径。

直接采用上述组合的原因是依赖已存在、接口足够小，而且页面渲染和版面事实可以分别验证。

## 2. 方法思路修改一下能用的

### [wps-cli](https://github.com/jjchen17/wps-cli)

- 许可证：MIT。
- 可借鉴点：`Documents.Open` 使用 `ConfirmConversions=False`、`ReadOnly=True`、`AddToRecentFiles=False`；关闭文档时使用“不保存更改”。
- 需要修改：项目只需要一个只读 DOCX→PDF 适配器，不引入它的完整会话、编辑和 CLI 框架。

### [harness-anything 的 WPS 后端](https://github.com/yb2460/harness-anything/blob/master/cli_anything/wps/utils/wps_backend.py)

- 许可证：MIT。
- 可借鉴点：WPS Writer 的 COM ProgID、`wdFormatPDF=17`、应用版本读取和 Windows/pywin32 延迟依赖。
- 需要修改：其通用 `open_document()` 默认没有只读参数；P1b 必须强制只读打开、禁止写回、导出失败清理半成品，并在关闭后再次核对源 SHA256。

### [self-service-printer 的 WPS 转 PDF 实现](https://github.com/233kun/self-service-printer/blob/master/convert/convert_wps.py)

- 可借鉴点：WPS Writer 可用 `SaveAs2(..., 17)` 生成 PDF，并可用 `pypdf`核对页数。
- 需要修改：P1b 优先使用 `ExportAsFixedFormat`，仅在接口不可用时回退到 `SaveAs2`；源文档必须只读打开，不能复用其业务状态管理。

以上项目只提供接口兼容线索。实现使用本项目自己的最小适配器，不复制外部项目的大段代码。

## 3. 学习它的核心方法思路

### [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)

- 许可证：Apache-2.0。
- 核心思路：先做页面布局检测，再对文字、表格、公式、图片区域分别识别，输出带坐标的 JSON/Markdown。
- P1b 借鉴：视觉区域必须带边界框、类型、置信度和证据；低置信度结果进入人工复核。
- 当前不直接引入：依赖和模型体积远超本阶段需要，且 WPS 生成 PDF 通常已有文字层。

### [LayoutParser](https://github.com/Layout-Parser/layout-parser)

- 许可证：Apache-2.0。
- 核心思路：用统一的 `Layout`/区域对象表示边界框、类型和文本，并支持裁剪、筛选、OCR 和可视化。
- P1b 借鉴：`region_id + bbox + role + decision + question binding` 的数据模型。

### [Docling](https://github.com/docling-project/docling)

- 许可证：MIT。
- 核心思路：将多种文档解析成统一结构模型，并保存页面、元素、表格和图片关系。
- P1b 借鉴：视觉解析只能作为旁路增强，必须与现有 WPS/OOXML 坐标真值对照，不能替代原始 Range。

### [UniLM / LayoutLM](https://github.com/microsoft/unilm)

- 许可证：MIT。
- 核心思路：联合建模文字、二维布局和页面图像，适合判断视觉丰富文档中的语义角色。
- P1b 借鉴：题目图片归属不能只看图片大小或文本关键词，应综合文字、位置、图像和上下文。
- 当前不直接引入：需要训练或微调数据，超出 P1b 的确定性审核范围。

## 4. 暂不直接使用的项目

### [PyMuPDF](https://github.com/pymupdf/PyMuPDF)

- 能力很适合：PDF 渲染、文字/图片坐标、JSON 和任意 DPI 页面图像。
- 许可证：AGPL-3.0，另有商业许可。
- 决策：本项目当前不新增 PyMuPDF 依赖，避免在许可边界未单独确认时把它接入生产工具链。现有 `pdfplumber + pdftoppm` 已能满足 P1b。

## 5. 最终选型

```text
WPS 只读导出：本项目最小 pywin32 适配器
PDF 版面事实：pdfplumber
逐页 PNG：pypdfium2（默认）/ pdftoppm（兼容）
图片裁剪与哈希：Pillow
OCR/深度布局模型：P1b 不设为必需依赖
```

该组合优先保证可验证、可回退和许可证清晰，不追求一次性自动理解所有页面。无法确认的页面区域固定标为 `unknown`，由人工门禁处理。
