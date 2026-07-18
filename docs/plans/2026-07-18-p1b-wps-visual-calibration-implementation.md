# P1b WPS Visual Calibration Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a read-only P1b sidecar that turns WPS-rendered PDFs into page PNG/layout JSON, records confirmed visual roles, calibrates P1a against execution-rule ground truth, and produces a blocking batch-gate report without enabling production F1.

**Architecture:** Keep rendering facts, review decisions, and calibration conclusions in separate JSON contracts. WPS exports the source document to PDF in read-only mode; Poppler renders pages; `pdfplumber` extracts page-coordinate text and image facts; a review layer binds regions to roles and question IDs; a calibration layer compares P1a clusters with human rule-family labels. Existing `DocumentProfile.json`, `ActionPlan.json`, and production parsers remain unchanged.

**Tech Stack:** Python 3.12, pywin32/WPS COM on Windows, macOS Quick Look + WebKit for development preview, `pypdfium2` with `pdftoppm` fallback, `pdfplumber`, Pillow, pytest, python-docx/OOXML.

---

## Constraints

- The user has separately authorized local Git initialization and a private GitHub handoff; source documents and derived review artifacts must remain ignored.
- Do not modify `wps_helper.py`, `墨痕快刀/core_parser.py`, or source DOCX files.
- Reject every output path that can overwrite or pollute an input directory.
- Hash each source before and after WPS export.
- Keep `automatic_rule_binding_enabled=false` and `production_execution_enabled=false`.

### Task 1: Record the open-source dependency decision

**Files:**
- Create: `docs/research/2026-07-18-p1b-document-visual-open-source.md`

**Steps:**

1. Record `pdfplumber` and Poppler/`pdf2image` under “directly reusable”.
2. Record `jjchen17/wps-cli` and `yb2460/harness-anything` under “adapt before use”.
3. Record PaddleOCR, LayoutParser, Docling, and LayoutLM under “learn the core method”.
4. Record PyMuPDF's useful API and AGPL constraint; do not add it as a dependency.
5. Verify all three reuse categories, license evidence, and local dependency choices are present.

Run:

```bash
rg -n "直接套用|修改后使用|学习核心思路|许可证|PyMuPDF" docs/research/2026-07-18-p1b-document-visual-open-source.md
```

### Task 2: Implement deterministic PDF page manifests

**Files:**
- Create: `shared_core/document_render.py`
- Create: `test_document_render.py`

**Steps:**

1. Write failing tests using a two-page synthetic PDF with text and one image.
2. Assert two PNGs, page sizes, relative paths, SHA256, text/image regions, source hash stability, and disabled production execution.
3. Add safety tests for source/output overlap, missing PDF/Poppler, invalid PNG, and invalid bounding boxes.
4. Run `python -m pytest -q -p no:cacheprovider test_document_render.py` and confirm failure.
5. Implement `sha256_file()`, `discover_pdftoppm()`, `render_pdf_pages()`, `extract_pdf_layout()`, `build_page_render_manifest()`, atomic JSON writes, and JSON-derived Markdown.
6. Use deterministic region IDs and record rendered crop SHA256 for visually comparable regions.
7. Run the focused tests and confirm they pass.

### Task 3: Implement a read-only WPS PDF exporter adapter

**Files:**
- Modify: `shared_core/document_render.py`
- Modify: `test_document_render.py`

**Steps:**

1. Write fake-COM tests before implementation.
2. Require lazy pywin32 import and ProgID fallback beginning with `KWPS.Application`.
3. Require `ConfirmConversions=False`, `ReadOnly=True`, and `AddToRecentFiles=False`.
4. Require PDF export through `ExportAsFixedFormat` or `SaveAs2(..., 17)` fallback.
5. Require `Close(SaveChanges=0)`, source hash verification, and failed partial-PDF cleanup.
6. Implement `export_docx_to_pdf_with_wps(..., dispatch_factory=None)` without importing pywin32 on macOS tests.
7. Run the focused fake-COM lifecycle tests.

### Task 4: Implement visual review contracts

**Files:**
- Create: `shared_core/document_visual_review.py`
- Create: `test_document_visual_review.py`

**Steps:**

1. Write failing tests for a template with one entry per page region.
2. Default each entry to `role=unknown`, `decision=review`, `review_status=draft`.
3. Write rejection tests for source/manifest mismatch, missing or duplicate region IDs, invalid roles, confirmed unknowns, excluded regions bound to questions, and included question content without one question ID.
4. Implement fixed roles: document banner/title, score instruction, exercise label, group heading, question text/media, and unknown.
5. Implement validation, review summaries, blocking reasons, atomic JSON writes, and JSON-derived Markdown.
6. Run `test_document_visual_review.py`.

### Task 5: Implement P1b threshold calibration and batch gates

**Files:**
- Create: `shared_core/document_family_calibration.py`
- Create: `test_document_family_calibration.py`

**Steps:**

1. Write synthetic threshold tests against human rule-family labels.
2. Report false splits and false merges for every deterministic candidate threshold.
3. Recommend the highest zero-error threshold when available.
4. Mark one ground-truth family as `batch_scoped` with a missing-negative-control warning.
5. Write gate failures for missing documents, duplicate SHA values, unknown/draft regions, decoration in F1, question media without binding, incomplete hash checks, and any enabled production switch.
6. Reuse `cluster_document_ids()` without changing P1a defaults.
7. Implement `build_p1b_calibration_report()` plus atomic JSON/Markdown writes.
8. Run `test_document_family_calibration.py`.

### Task 6: Add thin command-line tools

**Files:**
- Create: `tools/build_document_render.py`
- Create: `tools/build_visual_review.py`
- Create: `tools/calibrate_document_families.py`
- Create: `test_p1b_cli.py`

**Steps:**

1. Write failing CLI tests for offline PDF input, default WPS mode, overlap rejection, review-template generation, and calibration.
2. Implement these commands:

```powershell
python tools\build_document_render.py SOURCE.docx --output-dir OUTPUT
python tools\build_document_render.py SOURCE.docx --pdf-input WPS.pdf --attest-wps-export --output-dir OUTPUT
python tools\build_visual_review.py PageRenderManifest.json --output-dir OUTPUT
python tools\calibrate_document_families.py DocumentFamilyReport.json GroundTruth.json --reviews REVIEW_DIR --output-dir OUTPUT
```

3. Refuse WPS temporary files beginning with `~$` or `.~`.
4. Use Chinese error messages and non-zero exits for blockers.
5. Run `test_p1b_cli.py`.

### Task 7: Build the future-biology P1b trial batch

**Files:**
- Create outside the source directory: caller-selected PDF/PNG/JSON artifact directory.
- Create: `回归样本/P1b未来高二生物/README.md` with reproducible commands and compact report references, not all page PNGs.

**Steps:**

1. Process only formal DOCX files under `选必一活页` and `选必二活页`; ignore lock files and answer directories; expect 52 documents.
2. Export with Windows WPS and build page manifests. If WPS COM is unavailable, record the exact blocker and do not substitute another renderer as WPS truth.
3. Confirm the user-provided truth for `第1章　作业1　细胞生活的环境.docx`.
4. Exclude chapter banner, title, score instructions, `对点训练`, `题组一`, and `综合强化`.
5. Retain question images for 3, 4, 7, 8, 9, 10, 12, 14, and 15; question 8 must lose only `综合强化`.
6. Propagate confirmed repeated-media decisions only when role and context remain consistent; otherwise leave `unknown`.
7. Build rule-equivalence ground truth and calibrate P1a. If all 52 documents are one family, keep the threshold `batch_scoped`.
8. Verify 52/52 source hashes unchanged and that the source directories contain no new artifacts.

### Task 8: Regress, archive, and update the architecture ledger

**Files:**
- Modify: `问题归档/INDEX.md`
- Modify: `问题归档/题目录入/INDEX.md`
- Create: `问题归档/题目录入/2026-07-18-P1b-WPS页面视觉校准与批次门禁.md`
- Modify: `docs/墨痕教育架构问题工程思维分析拆解.md`
- Modify only if required: `README.md`, `AGENTS.md`

**Steps:**

1. Run:

```bash
python -m pytest -q -p no:cacheprovider \
  test_document_render.py \
  test_document_visual_review.py \
  test_document_family_calibration.py \
  test_p1b_cli.py \
  test_document_preflight.py \
  test_document_families.py
```

2. Clean test PDFs, PNGs, temporary JSON/Markdown, and harness-created lock files.
3. Update archive indexes before writing the individual issue record.
4. Record phenomenon, root cause, changes, test cases, commands, conclusion, and next diagnostic entrypoint.
5. Mark P1b complete only if the 52-document WPS-rendered gate passes; otherwise record the remaining Windows WPS validation gap.
6. Read back every final report, inspect representative page PNGs, check encoding and links, and report unverified risks.
