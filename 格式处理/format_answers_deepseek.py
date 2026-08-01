# coding: utf-8
"""已停用的 DeepSeek 答案清洗入口；保留函数名兼容旧调用。"""


def process_document_llm(doc=None):
    del doc
    print("   ⚠️ format_answers_deepseek.py 已停用，避免继续走失效旧链路。")
    print("   👉 请改用: python 格式处理/main.py")
    return False


def run_llm_engine(doc):
    return process_document_llm(doc)


if __name__ == "__main__":
    process_document_llm()
