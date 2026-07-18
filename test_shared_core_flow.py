from shared_core.answer_core import build_answer_units_from_paragraph_texts
from shared_core.models import DocNode
from shared_core.question_core import build_question_units_from_nodes, build_question_units_from_wps_spans
from shared_core.review import build_review_report, format_question_warning_details


def test_question_units():
    nodes = [
        DocNode(index=1, text="某房地产开发商计划在某城市进行商业性建筑开发，下图示意该城市部分区域和地租信息图层。据此完成下列问题。"),
        DocNode(index=2, text="", has_inline_media=True),
        DocNode(index=3, text="1．甲区域地租水平明显高于周边区域的主要原因是（    ）"),
        DocNode(index=4, text="A．距市中心较近    B．交通便捷    C．人口密度大    D．环境优美"),
        DocNode(index=5, text="2．若该房地产开发商拟建一个豪华酒店，最适宜建在（    ）"),
        DocNode(index=6, text="A．甲区域          B．乙区域      C．丙区域        D．丁区域"),
    ]
    q_nodes = [{"idx": 1, "type": "READING", "text": nodes[0].text}]
    units = build_question_units_from_wps_spans(
        doc_name="高一地理限训4.docx",
        subject_name="文科",
        nodes=nodes,
        q_nodes=q_nodes,
        end_p=6,
    )
    assert len(units) == 1
    unit = units[0]
    assert unit.question_type == "material_choice"
    assert unit.media_blocks == [2]
    assert "image_related_question" in unit.warnings


def test_answer_units():
    paragraphs = [
        "1．B",
        "解析：这是第1题解析。",
        "2．",
        "答案：（1）甲（2）乙",
        "解析：这是第2题解析。",
    ]
    units = build_answer_units_from_paragraph_texts(paragraphs)
    assert len(units) == 2
    assert units[0].answer_mode == "whole"
    assert units[1].answer_mode == "subquestion"
    assert len(units[1].answer_items) == 2


def test_review_report():
    nodes = [
        DocNode(index=1, text="9.可利用太阳光在新型复合催化剂表面实现高效分解水制氢,主要过程如图所示:"),
        DocNode(index=2, text="", has_inline_media=True),
        DocNode(index=3, text="下列说法错误的是 (      )"),
        DocNode(index=4, text="A.过程中实现了光能转化为化学能"),
        DocNode(index=5, text="B.反应中存在极性键的断裂和形成"),
        DocNode(index=6, text="C.过程Ⅰ吸收能量"),
        DocNode(index=7, text="D.过程Ⅲ发生的反应为H2O2H2↑+O2↑"),
    ]
    q_units = build_question_units_from_wps_spans(
        doc_name="限训4-化学.docx",
        subject_name="理科",
        nodes=nodes,
        q_nodes=[{"idx": 1, "type": "STD", "text": nodes[0].text}],
        end_p=7,
    )
    a_units = build_answer_units_from_paragraph_texts(["9．D", "解析：图示过程Ⅲ放出氧气。"])
    report = build_review_report("限训4-化学-已清洗.docx", q_units, a_units)
    assert report.summary["question_count"] == 1
    assert report.summary["answer_count"] == 1


def test_auto_detect_material_units_from_nodes():
    nodes = [
        DocNode(index=1, text="中心城市对周边地区的经济辐射能力受多种要素的影响。下图为“2017年我国两个中心城市经济辐射能力示意图”，图中数值越大，表明城市在该要素上的辐射力越强。完成1～2题。"),
        DocNode(index=2, text="", has_inline_media=True),
        DocNode(index=3, text="1．总体来看，成都的经济辐射力较西安强，其主要原因不包括 (　　)"),
        DocNode(index=4, text="A．成都经济发展水平更高         B．成都科技水平更高"),
        DocNode(index=5, text="2．下列有关增强西安经济辐射力的措施，叙述不合理的是 (　　)"),
        DocNode(index=6, text="A．提升西安城市行政等级，增强服务范围"),
    ]
    units = build_question_units_from_nodes("高二地理.docx", "文科", nodes)
    assert len(units) == 1
    assert units[0].node_type == "READING"
    assert units[0].source_span == (1, 6)


def test_auto_detect_science_figure_question():
    nodes = [
        DocNode(index=1, text="9.可利用太阳光在新型复合催化剂表面实现高效分解水制氢,主要过程如图所示:"),
        DocNode(index=2, text="", has_inline_media=True, page_break_before=True),
        DocNode(index=3, text="下列说法错误的是 (      )"),
        DocNode(index=4, text="A.过程中实现了光能转化为化学能"),
        DocNode(index=5, text="B.反应中存在极性键的断裂和形成"),
        DocNode(index=6, text="10.某实验继续如下。"),
    ]
    units = build_question_units_from_nodes("限训4-化学.docx", "理科", nodes)
    assert len(units) == 2
    assert units[0].source_span == (1, 5)
    assert "cross_page_media" in units[0].warnings


def test_question_warning_details_use_chinese_descriptions():
    details = format_question_warning_details(
        ["image_related_question", "sparse_options", "image_between_stem_and_options"]
    )
    assert details == [
        "这是图片/图表相关题，题干依赖图片或图表内容",
        "选项行偏少，可能存在漏抓或选项未识别完整",
        "题干和选项之间夹有图片或表格，边界可能受影响",
    ]


if __name__ == "__main__":
    test_question_units()
    test_answer_units()
    test_review_report()
    test_auto_detect_material_units_from_nodes()
    test_auto_detect_science_figure_question()
    test_question_warning_details_use_chinese_descriptions()
    print("ALL PASSED")
