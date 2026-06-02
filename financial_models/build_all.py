"""Build the three purpose-built Azalea Hospice workbooks from the shared engine.

    python financial_models/build_all.py

Outputs to financial_models/output/:
    Azalea_SBA_Loan_Package.xlsx
    Azalea_Investor_Model.xlsx
    Azalea_Operations_Dashboard.xlsx
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from financial_models.engine.build import Book
from financial_models.engine.workbooks import assemble_sba, assemble_investor, assemble_ops

OUT = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT, exist_ok=True)

TARGETS = [
    ("Azalea — SBA Loan Package", assemble_sba, "Azalea_SBA_Loan_Package.xlsx"),
    ("Azalea — Investor Model", assemble_investor, "Azalea_Investor_Model.xlsx"),
    ("Azalea — Operations Dashboard", assemble_ops, "Azalea_Operations_Dashboard.xlsx"),
]


ENGINE_TABS = {"Revenue Model", "Staffing & Payroll", "Operating Budget", "Detailed P&L",
               "Debt Schedule", "Cash Flow & BS", "Actuals vs Model"}


def polish(bk):
    """Professional finishing touches: tab colors + opening view."""
    for ws in bk.wb.worksheets:
        if ws.title == "Inputs":
            ws.sheet_properties.tabColor = "1F3864"      # navy
        elif ws.title in ENGINE_TABS:
            ws.sheet_properties.tabColor = "808080"      # grey (engine)
        else:
            ws.sheet_properties.tabColor = "2E5496"      # blue (audience)
        ws.sheet_view.zoomScale = 100
        ws.sheet_view.showGridLines = False
    # open on the first audience tab (sheet index 1) with A1 selected
    bk.wb.active = 1
    for ws in bk.wb.worksheets:
        ws.sheet_view.tabSelected = (ws is bk.wb.worksheets[1])


def main():
    for title, assemble, fname in TARGETS:
        bk = Book(title)
        assemble(bk)
        polish(bk)
        bk.wb.properties.title = title
        bk.wb.properties.creator = "Azalea Hospice Financial Model"
        bk.wb.calculation.fullCalcOnLoad = True
        path = os.path.join(OUT, fname)
        bk.wb.save(path)
        print(f"  built {fname:42} ({len(bk.wb.sheetnames)} tabs)")
    print("Done.")


if __name__ == "__main__":
    main()
