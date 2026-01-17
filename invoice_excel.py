import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import re

# Invoice maker function.
def func_invoice(name, invoice_no, current_date, utr, nin, week_ending, company_name, company_road, company_city, 
                company_postcode, site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed, site_thurs, 
                pay_thurs, site_fri, pay_fri, site_sat, pay_sat, site_sun, pay_sun, bank_name, sort_code, account_no, expenses):
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # --- Calculations ---
    # Convert inputs to float for math
    gross_val = sum([float(pay_mon), float(pay_tues), float(pay_wed), float(pay_thurs), float(pay_fri), float(pay_sat), float(pay_sun)])
    tax_val = gross_val * 0.20
    total_val = (gross_val - tax_val) + float(expenses)

    # --- Styles ---
    bold_font = Font(bold=True, name="Arial")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    wrap_alignment = Alignment(wrap_text=True, vertical="top")
    right_alignment = Alignment(horizontal="right")

    # --- TOP SECTION ---
    ws["A1"] = "Name"; ws["A1"].font = bold_font
    ws["B1"] = name

    ws["D1"] = "Invoice No"; ws["D1"].font = bold_font
    ws["E1"] = invoice_no; ws["E1"].alignment = right_alignment

    ws["D3"] = "Date"; ws["D3"].font = bold_font
    ws["E3"] = current_date; ws["E3"].alignment = right_alignment

    # --- TO SECTION (With Text Wrapping) ---
    ws["A4"] = "TO"; ws["A4"].font = bold_font
    
    # We put the company name in A5 and enable wrap_text so "Solutions" drops down if the cell is narrow
    ws["A5"] = company_name
    ws["A5"].alignment = wrap_alignment
    
    ws["A9"] = company_road
    ws["A10"] = company_city
    ws["A11"] = company_postcode

    # --- ID SECTION ---
    ws["D9"] = "UTR"; ws["D9"].font = bold_font
    ws["E9"] = utr; ws["E9"].alignment = right_alignment

    ws["D11"] = "NIN"; ws["D11"].font = bold_font
    ws["E11"] = nin; ws["E11"].alignment = right_alignment

    ws["A16"] = "Week Ending"; ws["A16"].font = bold_font
    ws["B16"] = week_ending

    # --- DAILY ENTRIES (Matching your Grid) ---
    days = [
        ("Monday", site_mon, pay_mon, 19),
        ("Tuesday", site_tues, pay_tues, 22),
        ("Wednesday", site_wed, pay_wed, 25),
        ("Thursday", site_thurs, pay_thurs, 28),
        ("Friday", site_fri, pay_fri, 31),
        ("Saturday", site_sat, pay_sat, 34),
        ("Sunday", site_sun, pay_sun, 37)
    ]

    for day, site, pay, row in days:
        ws[f"A{row}"] = day
        ws[f"B{row}"] = site
        ws[f"A{row+1}"] = "Job Name"
        ws[f"B{row+1}"] = f"£{float(pay):.2f}"

    # --- FOOTER & TOTALS (With Yellow Highlights) ---
    ws["A40"] = "Gross"; ws["A40"].font = bold_font
    ws["B40"] = f"£{gross_val:.2f}"
    
    ws["A41"] = "        -20%"
    ws["B41"] = f"£{tax_val:.2f}"
    
    ws["A42"] = "Expenses"; ws["A42"].font = bold_font
    ws["B42"] = f"£{float(expenses):.2f}"
    
    # Yellow Highlight for the Total
    ws["A44"] = "Total"; ws["A44"].font = bold_font
    ws["B44"] = f"£{total_val:.2f}"
    ws["B44"].fill = yellow_fill
    ws["B44"].font = bold_font

    # Bank Details
    ws["D40"] = "Bank"; ws["D40"].font = bold_font
    ws["E40"] = bank_name; ws["E40"].alignment = right_alignment

    ws["D41"] = "Sort Code"; ws["D41"].font = bold_font
    ws["E41"] = sort_code; ws["E41"].alignment = right_alignment

    ws["D42"] = "Account No."; ws["D42"].font = bold_font
    ws["E42"] = account_no; ws["E42"].alignment = right_alignment

    # Setting column width to force the wrap on company name
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 20

    file_name = re.sub(r"\s+", "_", name).lower() + "_invoice.xlsx"
    wb.save(file_name)
    return file_name
