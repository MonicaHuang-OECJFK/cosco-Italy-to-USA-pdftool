import sys
import os
from openpyxl import Workbook
from cosco_extract import extract
from to_excel import write_rate_sheet


def main():
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "cosco_pdf.pdf"

    direct, outports = extract(pdf_path)

    wb = Workbook()
    write_rate_sheet(wb.active,         "Direct Ports", direct)
    write_rate_sheet(wb.create_sheet(), "Outports",     outports)

    out_path = os.path.join(os.getcwd(), "pdf_extract.xlsx")
    wb.save(out_path)
    print(f"Saved: {out_path}") 


if __name__ == "__main__":
    main()
