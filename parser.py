from docx import Document
from pdfminer.high_level import extract_text
import os

def parse_docx(file_path):
    doc = Document(file_path)
    text = "\n".join([p.text for p in doc.paragraphs])
    return text

def parse_pdf(file_path):
    return extract_text(file_path)

def parse_resume(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        text = parse_docx(file_path)
    elif ext == ".pdf":
        text = parse_pdf(file_path)
    else:
        raise ValueError("Unsupported file type")

    lines = text.splitlines()
    name = lines[0] if lines else "Candidate"
    return {
        "name": name,
        "summary": "Summary not parsed",
        "experience": "Experience not parsed",
        "skills": "Skills not parsed"
    }
