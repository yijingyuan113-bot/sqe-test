import os
import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def read_docx(path):
    doc = Document(path)
    texts = []
    for para in doc.paragraphs:
        if para.text.strip():
            texts.append(para.text)
    return texts

def extract_questions_from_texts(texts):
    """Extract questions from texts - this is a simplified parser"""
    questions = []
    current_q = None
    for line in texts:
        line = line.strip()
        if not line:
            continue
        # Detect question patterns
        if line.startswith('Q') or line.startswith('Question') or line.startswith('题目'):
            if current_q:
                questions.append(current_q)
            current_q = {'raw': line, 'text': line}
        elif current_q:
            current_q['raw'] += '\n' + line
            current_q['text'] += '\n' + line
    if current_q:
        questions.append(current_q)
    return questions

# List docx files
sqe_dir = r'C:\Users\Administrator\Desktop\Desktop\SQE'
docx_files = [f for f in os.listdir(sqe_dir) if f.endswith('.docx')]
print(f"Found docx files: {docx_files}")

# Read first docx to understand structure
for f in docx_files[:3]:
    path = os.path.join(sqe_dir, f)
    print(f"\n=== {f} ===")
    texts = read_docx(path)
    print(f"Total paragraphs: {len(texts)}")
    if texts:
        print("First 10 lines:")
        for t in texts[:10]:
            print(f"  {t[:100]}")