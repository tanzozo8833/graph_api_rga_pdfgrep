import io

import pdfplumber


def extract_pdf(content: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            parts.append(f"[PAGE {i}]")
            text = page.extract_text() or ""
            parts.append(text)
    return "\n".join(parts)
