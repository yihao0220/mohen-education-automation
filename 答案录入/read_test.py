import os
import glob

files = glob.glob(r"E:\PYTHON\practice\*\.sisyphus\test_answers.docx")
if files:
    file_path = files[0]
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    with open('test_answers.txt', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Saved to test_answers.txt")
