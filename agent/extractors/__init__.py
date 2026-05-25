import os

from .pdf import extract_pdf
from .docx import extract_docx
from .xlsx import extract_xlsx
from .pptx import extract_pptx

OFFICE_EXTS = {".pdf", ".docx", ".xlsx", ".pptx"}
PLAINTEXT_EXTS = {".txt", ".md", ".csv", ".tsv", ".log", ".json", ".xml", ".yaml", ".yml"}
SUPPORTED = OFFICE_EXTS | PLAINTEXT_EXTS


def extract_text(filename: str, content: bytes) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return extract_pdf(content)
    if ext == ".docx":
        return extract_docx(content)
    if ext == ".xlsx":
        return extract_xlsx(content)
    if ext == ".pptx":
        return extract_pptx(content)
    if ext in PLAINTEXT_EXTS:
        try:
            return content.decode("utf-8")
        except UnicodeDecodeError:
            return content.decode("latin-1", errors="replace")
    raise ValueError(f"Unsupported file type: {ext}")
