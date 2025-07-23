from docx import Document

# В реальной реализации здесь будет сбор данных и генерация docx

def generate_docx_list(event_id):
    doc = Document()
    doc.add_heading('Список докладов', 0)
    f = BytesIO()
    doc.save(f)
    return f.getvalue()

def generate_docx_report(event_id):
    doc = Document()
    doc.add_heading('Отчет о проведении конференции', 0)
    f = BytesIO()
    doc.save(f)
    return f.getvalue()

def generate_docx_papers(event_id):
    doc = Document()
    doc.add_heading('Список предоставленных к публикации статей', 0)
    f = BytesIO()
    doc.save(f)
    return f.getvalue()

from io import BytesIO
