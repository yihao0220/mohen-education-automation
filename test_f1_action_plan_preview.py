import pytest

from tools.execute_f1_action_plan import (
    APPROVED_PILOT_SHA256,
    execute_actions,
    validate_execution_profile,
)
from tools.preview_f1_action_plan import preview_actions, validate_preview_plan


class FakeRange:
    def __init__(self, text: str, start: int, end: int, table_count: int = 0):
        self.Text = text
        self.Start = start
        self.End = end
        self.Tables = type("FakeTables", (), {"Count": table_count})()
        self.selected = False

    def Select(self):
        self.selected = True


class FakeParagraph:
    def __init__(self, text: str, index: int, table_id: int | None = None):
        self.Range = FakeRange(text, index * 100, index * 100 + len(text))
        self.table_id = table_id


class FakeParagraphs:
    def __init__(self, texts: list[str], table_ids: dict[int, int] | None = None):
        table_ids = table_ids or {}
        self._items = [
            FakeParagraph(text, index, table_ids.get(index))
            for index, text in enumerate(texts, 1)
        ]
        self.Count = len(self._items)

    def __call__(self, index: int):
        return self._items[index - 1]


class FakeDocument:
    def __init__(self, texts: list[str], table_ids: dict[int, int] | None = None):
        self.Paragraphs = FakeParagraphs(texts, table_ids=table_ids)
        self.selected_ranges: list[FakeRange] = []

    def Range(self, start: int, end: int):
        selected = [
            paragraph.Range.Text
            for paragraph in self.Paragraphs._items
            if paragraph.Range.Start >= start and paragraph.Range.End <= end
        ]
        selected_table_ids = {
            paragraph.table_id
            for paragraph in self.Paragraphs._items
            if paragraph.Range.Start >= start
            and paragraph.Range.End <= end
            and paragraph.table_id is not None
        }
        result = FakeRange("\n".join(selected), start, end, len(selected_table_ids))
        self.selected_ranges.append(result)
        return result


def _plan():
    return {
        "mode": "preview_only",
        "execution_enabled": False,
        "actions": [
            {
                "sequence": 1,
                "key": "F1",
                "question_ids": ["1", "2"],
                "preview": "材料题。1．第一题",
                "source_ref": {"paragraph_start": 2, "paragraph_end": 4},
            },
            {
                "sequence": 2,
                "key": "F1",
                "question_ids": ["3"],
                "preview": "3．第三题",
                "source_ref": {"paragraph_start": 5, "paragraph_end": 5},
            },
        ],
    }


def test_preview_actions_selects_ranges_without_pressing_f1():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    prompts: list[str] = []

    receipt = preview_actions(
        doc,
        _plan(),
        confirm=lambda prompt: prompts.append(prompt) or "",
        output=lambda _message: None,
    )

    assert receipt == {"status": "completed", "selected_actions": 2, "keypress_count": 0}
    assert [(item.Start, item.End) for item in doc.selected_ranges] == [(200, 405), (500, 505)]
    assert all(item.selected for item in doc.selected_ranges)
    assert len(prompts) == 2


def test_preview_actions_finds_nearby_range_when_wps_has_extra_paragraph():
    doc = FakeDocument(
        ["标题", "材料题。", "1．第一题", "2．第二题", "WPS 额外段落", "3．第三题"]
    )
    messages: list[str] = []

    receipt = preview_actions(
        doc,
        _plan(),
        confirm=lambda _prompt: "",
        output=messages.append,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt == {"status": "completed", "selected_actions": 2, "keypress_count": 0}
    assert [(item.Start, item.End) for item in selected] == [(200, 405), (600, 605)]
    assert any("偏移 +1 段" in message for message in messages)


def test_preview_actions_does_not_choose_overlapping_range_before_anchor():
    plan = _plan()
    plan["actions"][1]["source_ref"] = {"paragraph_start": 5, "paragraph_end": 7}
    doc = FakeDocument(
        [
            "标题",
            "材料题。",
            "1．第一题",
            "2．第二题",
            "WPS 额外段落",
            "3．第三题",
            "第三题选项",
            "第三题结尾",
        ]
    )

    receipt = preview_actions(
        doc,
        plan,
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt["keypress_count"] == 0
    assert (selected[1].Start, selected[1].End) == (600, 805)


def test_preview_actions_accepts_same_material_when_wps_formats_question_marker_differently():
    plan = _plan()
    plan["actions"][1]["preview"] = (
        "下图为地球自转示意图。据此完成3～4题。 3．图中能正确表示地球自转方向"
    )
    plan["actions"][1]["source_ref"] = {"paragraph_start": 5, "paragraph_end": 6}
    doc = FakeDocument(
        [
            "标题",
            "材料题。",
            "1．第一题",
            "2．第二题",
            "下图为地球自转示意图。据此完成3～4题。",
            "/3.图中能正确表示地球自转方向",
        ]
    )

    receipt = preview_actions(
        doc,
        plan,
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt["keypress_count"] == 0
    assert (selected[1].Start, selected[1].End) == (
        doc.Paragraphs(5).Range.Start,
        doc.Paragraphs(6).Range.End,
    )


def test_preview_actions_uses_start_preview_when_short_material_is_followed_by_wps_slash_marker():
    plan = _plan()
    plan["actions"] = [plan["actions"][0]]
    plan["actions"][0]["preview"] = "读图，完成6～7题。 6．图中 A 点的昼长为"
    plan["actions"][0]["source_ref"] = {
        "paragraph_start": 3,
        "paragraph_end": 8,
        "table_indexes": [],
        "start_preview": "读图，完成6～7题。",
        "end_preview": "A．太原 B．长春 C．南昌 D．昆明",
    }
    doc = FakeDocument(
        [
            "标题",
            "WPS 表格单元格占位",
            "读图，完成6～7题。",
            "",
            "/6.图中 A 点的昼长为",
            "A．24小时 B．12小时",
            "7．下列各地白昼最长的是",
            "A．太原 B．长春 C．南昌 D．昆明",
        ]
    )

    receipt = preview_actions(
        doc,
        plan,
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt == {"status": "completed", "selected_actions": 1, "keypress_count": 0}
    assert (selected[-1].Start, selected[-1].End) == (
        doc.Paragraphs(3).Range.Start,
        doc.Paragraphs(8).Range.End,
    )


def test_preview_actions_legacy_plan_uses_material_sentence_before_question_marker():
    plan = _plan()
    plan["actions"] = [plan["actions"][0]]
    plan["actions"][0]["preview"] = "读图，完成6～7题。 6．图中 A 点的昼长为"
    plan["actions"][0]["source_ref"] = {
        "paragraph_start": 20,
        "paragraph_end": 25,
        "table_indexes": [],
    }
    doc = FakeDocument(
        [
            *[f"前置段落 {index}" for index in range(1, 20)],
            "读图，完成6～7题。",
            "",
            "/6.图中 A 点的昼长为",
            "A．24小时 B．12小时",
            "7．下列各地白昼最长的是",
            "A．太原 B．长春 C．南昌 D．昆明",
        ]
    )

    receipt = preview_actions(
        doc,
        plan,
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt == {"status": "completed", "selected_actions": 1, "keypress_count": 0}
    assert (selected[-1].Start, selected[-1].End) == (
        doc.Paragraphs(20).Range.Start,
        doc.Paragraphs(25).Range.End,
    )


def test_preview_actions_can_find_unique_range_more_than_five_paragraphs_away():
    doc = FakeDocument(
        ["标题", "材料题。", "1．第一题", "2．第二题"]
        + [f"WPS 额外段落 {index}" for index in range(1, 8)]
        + ["3．第三题"]
    )

    receipt = preview_actions(
        doc,
        _plan(),
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt["keypress_count"] == 0
    assert (selected[1].Start, selected[1].End) == (1200, 1205)


def test_preview_plan_rejects_any_executable_plan():
    plan = _plan()
    plan["execution_enabled"] = True

    try:
        validate_preview_plan(plan)
    except ValueError as exc:
        assert "execution_enabled" in str(exc)
    else:
        raise AssertionError("可执行计划不应进入无按键预览")


def test_preview_actions_uses_start_and_end_anchors_to_include_native_table():
    plan = _plan()
    plan["actions"] = [plan["actions"][0]]
    plan["actions"][0]["source_ref"] = {
        "paragraph_start": 2,
        "paragraph_end": 4,
        "table_indexes": [1],
        "start_preview": "材料题。",
        "end_preview": "2．第二题",
    }
    doc = FakeDocument(
        ["标题", "材料题。", "表头", "单元格甲", "单元格乙", "2．第二题"],
        table_ids={3: 1, 4: 1, 5: 1},
    )

    receipt = preview_actions(
        doc,
        plan,
        confirm=lambda _prompt: "",
        output=lambda _message: None,
    )

    selected = [item for item in doc.selected_ranges if item.selected]
    assert receipt == {"status": "completed", "selected_actions": 1, "keypress_count": 0}
    assert (selected[-1].Start, selected[-1].End) == (200, 605)
    assert selected[-1].Tables.Count == 1


def test_preview_actions_blocks_table_action_when_selected_range_has_no_table():
    plan = _plan()
    plan["actions"] = [plan["actions"][0]]
    plan["actions"][0]["source_ref"] = {
        "paragraph_start": 2,
        "paragraph_end": 4,
        "table_indexes": [1],
        "start_preview": "材料题。",
        "end_preview": "2．第二题",
    }
    doc = FakeDocument(["标题", "材料题。", "普通段落", "2．第二题"])

    with pytest.raises(ValueError, match="表格数量"):
        preview_actions(
            doc,
            plan,
            confirm=lambda _prompt: "",
            output=lambda _message: None,
        )


def test_execute_actions_stops_before_f1_without_exact_confirmation():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    presses: list[str] = []

    receipt = execute_actions(
        doc,
        _plan(),
        confirm=lambda _prompt: "s",
        press_f1=lambda: presses.append("f1"),
        sleep=lambda _seconds: None,
        output=lambda _message: None,
    )

    assert receipt == {
        "status": "stopped",
        "selected_actions": 1,
        "executed_actions": 0,
        "keypress_count": 0,
    }
    assert presses == []


def test_execute_actions_defaults_to_one_action_canary():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    answers = iter(["f1", ""])
    presses: list[str] = []

    receipt = execute_actions(
        doc,
        _plan(),
        confirm=lambda _prompt: next(answers),
        press_f1=lambda: presses.append("f1"),
        sleep=lambda _seconds: None,
        output=lambda _message: None,
    )

    assert receipt == {
        "status": "canary_completed",
        "selected_actions": 1,
        "executed_actions": 1,
        "keypress_count": 1,
    }
    assert presses == ["f1"]


def test_execute_actions_can_run_all_actions_with_confirmation_per_action():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    answers = iter(["f1", "", "f1", ""])
    presses: list[str] = []

    receipt = execute_actions(
        doc,
        _plan(),
        max_actions=None,
        confirm=lambda _prompt: next(answers),
        press_f1=lambda: presses.append("f1"),
        sleep=lambda _seconds: None,
        output=lambda _message: None,
    )

    assert receipt == {
        "status": "completed",
        "selected_actions": 2,
        "executed_actions": 2,
        "keypress_count": 2,
    }
    assert presses == ["f1", "f1"]


def test_execute_actions_can_stop_after_checking_temporary_plugin_result():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    answers = iter(["f1", "s"])
    presses: list[str] = []

    receipt = execute_actions(
        doc,
        _plan(),
        max_actions=None,
        confirm=lambda _prompt: next(answers),
        press_f1=lambda: presses.append("f1"),
        sleep=lambda _seconds: None,
        output=lambda _message: None,
    )

    assert receipt == {
        "status": "stopped",
        "selected_actions": 1,
        "executed_actions": 1,
        "keypress_count": 1,
    }


def test_execute_actions_reports_completed_when_stop_is_entered_after_last_action():
    doc = FakeDocument(["标题", "材料题。", "1．第一题", "2．第二题", "3．第三题"])
    answers = iter(["f1", "", "f1", "s"])

    receipt = execute_actions(
        doc,
        _plan(),
        max_actions=None,
        confirm=lambda _prompt: next(answers),
        press_f1=lambda: None,
        sleep=lambda _seconds: None,
        output=lambda _message: None,
    )

    assert receipt == {
        "status": "completed",
        "selected_actions": 2,
        "executed_actions": 2,
        "keypress_count": 2,
    }


def test_execution_profile_still_blocks_native_tables():
    with pytest.raises(RuntimeError, match="原生表格"):
        validate_execution_profile(
            {
                "fingerprint": {"table_count": 1},
                "source": {"sha256": APPROVED_PILOT_SHA256},
            }
        )


def test_execution_profile_rejects_unapproved_document_hash():
    with pytest.raises(RuntimeError, match="8/8 WPS 选区验收"):
        validate_execution_profile(
            {
                "fingerprint": {"table_count": 0},
                "source": {"sha256": "not-approved"},
            }
        )
