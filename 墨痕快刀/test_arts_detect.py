import sys
sys.path.append('E:/PYTHON/practice/墨痕教育/core')
import config
from core_parser import detect_subject

class MockRange:
    def __init__(self, text):
        self.Text = text
    def __call__(self, start, end):
        return self
    @property
    def End(self):
        return len(self.Text)

class MockDoc:
    def __init__(self, name, text):
        self.Name = name
        self.Range = MockRange(text)

text = """2.6《哈姆莱特》同步习题
一、基础夯实
1.下列词语中加点字的注音,全都正确的一项是(　　)
...
二、延伸阅读
阅读下面的文字,完成第6~9题。
第五幕
第二场　城堡中的厅堂(节选)
...
"""
doc = MockDoc("document1.docx", text)
subj = detect_subject(doc)
print(subj['name'])
