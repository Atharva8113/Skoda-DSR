import fitz
from pathlib import Path

files = [
    r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\new bl and shipper\1\90533303.pdf',
    r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\new bl and shipper\1\DB3GL0M91.PDF'
]

keywords = ["PREMIUM", "EVERGREEN", "AUDI", "VOLKSWAGEN", "SKODA", "CELKOV"]

for f in files:
    print(f"--- {Path(f).name} ---")
    try:
        doc = fitz.open(f)
        text = "".join(page.get_text().upper() for page in doc)
        for kw in keywords:
            if kw in text:
                print(f"Found keyword: {kw}")
        # print first 1000 chars anyway
        print(text[:1000])
        doc.close()
    except Exception as e:
        print(f"Error: {e}")
