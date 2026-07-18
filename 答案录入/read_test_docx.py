from docx import Document

def read_docx(file_path):
    doc = Document(file_path)
    for para in doc.paragraphs:
        if para.text.strip():
            print(para.text)

if __name__ == "__main__":
    read_docx(".sisyphus/test_answers.docx")
