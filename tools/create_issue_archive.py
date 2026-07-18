from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE_ROOT = ROOT / "问题归档"
TEMPLATE_PATH = ARCHIVE_ROOT / "_templates" / "问题记录模板.md"
MODULES = {
    "question": "题目录入",
    "format": "格式转换",
    "answer": "答案录入",
}


def slugify(title: str) -> str:
    return title.strip().replace(" ", "-")


def main() -> None:
    parser = argparse.ArgumentParser(description="创建问题归档记录")
    parser.add_argument("module", choices=MODULES.keys(), help="归档模块")
    parser.add_argument("title", help="问题标题")
    args = parser.parse_args()

    module_name = MODULES[args.module]
    today = date.today().isoformat()
    filename = f"{today}-{slugify(args.title)}.md"
    target_dir = ARCHIVE_ROOT / module_name
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / filename

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    target_path.write_text(template.replace("{标题}", args.title), encoding="utf-8")

    index_path = target_dir / "INDEX.md"
    if index_path.exists():
        content = index_path.read_text(encoding="utf-8").rstrip()
    else:
        content = f"# {module_name}问题索引\n\n| 日期 | 标题 | 状态 | 文件 |\n|------|------|------|------|"
    row = f"\n| {today} | {args.title} | 待处理 | [{filename}](./{filename}) |"
    if row not in content:
        index_path.write_text(content + row + "\n", encoding="utf-8")

    print(target_path)


if __name__ == "__main__":
    main()
