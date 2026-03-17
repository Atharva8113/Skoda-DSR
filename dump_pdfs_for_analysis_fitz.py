import fitz
import os

files_to_dump = [
    (r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\new bl and shipper\1\90533303.pdf', 'eg_bl.txt'),
    (r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\new bl and shipper\1\DB3GL0M91.PDF', 'eg_inv.txt'),
    (r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\HBL  DESTRA0000007883  MBL  HLCUHAM2602AQGK2\1394415430-INVOICE.pdf', 'hbl_mbl_inv.txt'),
    (r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\HBL  DESTRA0000007883  MBL  HLCUHAM2602AQGK2\MCOP0101_333191936.pdf', 'hbl_mbl_mcop.txt'),
    (r'c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda DSR\16-03-26\HBL  DESTRA0000007883  MBL  HLCUHAM2602AQGK2\WWTAN_88A2341BB1964215970B9B251B22ADDC.pdf', 'hbl_mbl_wwtan.txt')
]

for pdf_path, out_path in files_to_dump:
    if os.path.exists(pdf_path):
        print(f"Dumping {pdf_path} to {out_path}")
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(text)
            doc.close()
        except Exception as e:
            print(f"Error dumping {pdf_path}: {e}")
    else:
        print(f"Path does not exist: {pdf_path}")
