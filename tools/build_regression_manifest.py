import json
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
QUESTION_DIR = PROJECT_ROOT / "墨痕快刀"
FORMAT_DIR = PROJECT_ROOT / "格式处理"
REGRESSION_DIR = PROJECT_ROOT / "回归样本"
REGRESSION_MANIFEST_PATH = REGRESSION_DIR / "样本清单.json"


def _portable_file_name(file_path: Path) -> str:
    return str(file_path).replace("\\", "/").rsplit("/", 1)[-1]


def build_manifest_entry(
    *,
    file_path: Path,
    sample_kind: str,
    subject: str | None = None,
    question_count: int | None = None,
    tags: list[str] | None = None,
    requires_subquestion_split: bool | None = None,
) -> dict:
    return {
        "file_name": _portable_file_name(file_path),
        "relative_path": str(file_path),
        "sample_kind": sample_kind,
        "subject": subject or "未知",
        "question_count": question_count,
        "tags": tags or [],
        "requires_subquestion_split": requires_subquestion_split,
    }


def _guess_subject_from_name(name: str) -> str:
    if any(key in name for key in ["英语", "English", "外语"]):
        return "英语"
    if any(key in name for key in ["物理", "化学", "生物", "科学", "理综"]):
        return "理科"
    if any(key in name for key in ["地理", "历史", "政治", "语文", "文综", "道法"]):
        return "文科"
    return "未知"


def _guess_tags_from_name(name: str) -> list[str]:
    tags = []
    if any(key in name for key in ["图", "示意", "装置"]):
        tags.append("image")
    if any(key in name for key in ["材料", "阅读", "据此", "下题"]):
        tags.append("material")
    if any(key in name for key in ["答案", "解析"]):
        tags.append("answer")
    return tags


def sync_regression_manifest() -> str:
    REGRESSION_DIR.mkdir(parents=True, exist_ok=True)
    sample_specs = [
        (FORMAT_DIR / "待清洗文件", "raw_answer_doc"),
        (FORMAT_DIR / "原格式", "formatted_answer_reference"),
        (QUESTION_DIR / "待录入文档", "question_doc"),
    ]

    entries = []
    for folder, sample_kind in sample_specs:
        if not folder.exists():
            continue
        for path in sorted(folder.iterdir()):
            if path.suffix.lower() not in {".doc", ".docx"} or path.name.startswith("~$"):
                continue
            entries.append(
                build_manifest_entry(
                    file_path=path.relative_to(PROJECT_ROOT),
                    sample_kind=sample_kind,
                    subject=_guess_subject_from_name(path.name),
                    tags=_guess_tags_from_name(path.name),
                    requires_subquestion_split=None,
                )
            )

    REGRESSION_MANIFEST_PATH.write_text(
        json.dumps({"entries": entries}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(REGRESSION_MANIFEST_PATH)


def main():
    path = sync_regression_manifest()
    print(path)


if __name__ == "__main__":
    main()
