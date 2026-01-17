import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import re

def func_invoice(name, invoice_no, current_date, utr, nin, week_ending, company_name, company_road, company_city, 
                company_postcode, site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed, site_thurs, 
                pay_thurs, site_fri, pay_fri, site_sat, pay_sat, site_sun, pay_sun, bank_name, sort_code, account_no, expenses):
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # --- Calculation Logic ---
    payment_info = func_pay(float(pay_mon), float(pay_tues), float(pay_wed), float(pay_thurs), float(pay_fri), float(pay_sat), float(pay_sun))
    
    # --- Styling Presets ---
    bold_f = Font(bold=True, name="Arial", size=11)
    reg_f = Font(name="Arial", size=11)
    align_left = Alignment(horizontal="left")
    align_right = Alignment(horizontal="right")

    # --- TOP SECTION (Name & Invoice Details) ---
    ws["A1"] = "Name"; ws["A1"].font = bold_f
    ws["B1"] = name; ws["B1"].font = reg_f

    ws["D1"] = "Invoice No"; ws["D1"].font = bold_f
    ws["E1"] = invoice_no; ws["E1"].font = reg_f
    ws["E1"].alignment = align_right

    ws["D3"] = "Date"; ws["D3"].font = bold_f
    ws["E3"] = current_date; ws["E3"].font = reg_f
    ws["E3"].alignment = align_right

    # --- BILL TO SECTION (Left Side) ---
    ws["A4"] = "TO"; ws["A4"].font = bold_f
    ws["A5"] = company_name; ws["A5"].font = reg_f
    ws["A9"] = company_road; ws["A9"].font = reg_f
    ws["A10"] = company_city; ws["A10"].font = reg_f
    ws["A11"] = company_postcode; ws["A11"].font = reg_f

    # --- COMPLIANCE SECTION (Right Side) ---
    ws["D9"] = "UTR"; ws["D9"].font = bold_f
    ws["E9"] = utr; ws["E9"].font = reg_f
    ws["E9"].alignment = align_right

    ws["D11"] = "NIN"; ws["D11"].font = bold_f
    ws["E11"] = nin; ws["E11"].font = reg_f
    ws["E11"].alignment = align_right

    # --- WEEK ENDING ---
    ws["A16"] = "Week Ending"; ws["A16"].font = bold_f
    ws["B16"] = week_ending; ws["B16"].font = reg_f

    # --- JOB ENTRIES (Rows 19 to 38) ---
    # We use the exact day/site/pay structure from your image
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
        ws[f"A{row}"] = day; ws[f"A{row}"].font = bold_f
        ws[f"B{row}"] = site; ws[f"B{row}"].font = reg_f
        
        ws[f"A{row+1}"] = "Job Name"; ws[f"A{row+1}"].font = reg_f
        ws[f"B{row+1}"] = f"£{pay}"; ws[f"B{row+1}"].font = reg_f

    # --- TOTALS & BANKING (Row 40 onwards) ---
    # Left: Pay Breakdown
    ws["A40"] = "Gross"; ws["A40"].font = bold_f
    ws["B40"] = f"£{payment_info['gross']:.2f}"
    
    ws["A41"] = "        -20%"; ws["A41"].font = reg_f
    ws["B41"] = f"£{payment_info['tax']:.2f}"
    
    ws["A42"] = "Expenses"; ws["A42"].font = bold_f
    ws["B42"] = f"£{float(expenses):.2f}"
    
    ws["A44"] = "Total"; ws["A44"].font = bold_f
    ws["B44"] = f"£{payment_info['take_home'] + float(expenses):.2f}"
    ws["B44"].font = bold_f

    # Right: Bank Details
    ws["D40"] = "Bank"; ws["D40"].font = bold_f
    ws["E40"] = bank_name; ws["E40"].font = reg_f
    ws["E40"].alignment = align_right

    ws["D41"] = "Sort Code"; ws["D41"].font = bold_f
    ws["E41"] = sort_code; ws["E41"].font = reg_f
    ws["E41"].alignment = align_right

    ws["D42"] = "Account No."; ws["D42"].font = bold_f
    ws["E42"] = account_no; ws["E42"].font = reg_f
    ws["E42"].alignment = align_right

    # --- Column Width Adjustment ---
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20

    file_name = re.sub(r"\s+", "_", name).lower() + "_invoice.xlsx"
    wb.save(file_name)
    return file_name

def func_pay(mon, tues, wed, thurs, fri, sat, sun):
    gross = sum([mon, tues, wed, thurs, fri, sat, sun])
    tax = gross * 0.20
    take_home = gross - tax
    return {"gross": gross, "tax": tax, "take_home": take_home}
