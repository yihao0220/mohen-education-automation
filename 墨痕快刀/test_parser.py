import re
text = """1．A
解析：这是解析
2．B
解析：这也是解析
"""
print(re.match(r'^\s*\d+[．.]', "1．A"))
print(re.match(r'^\s*\d+[．.]', "解析：这是解析"))
