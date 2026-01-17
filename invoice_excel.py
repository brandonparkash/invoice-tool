import openpyxl, re
from openpyxl.styles import Font, Alignment, PatternFill

def func_invoice(name, invoice_no, current_date, utr, nin, week_ending, 
                site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed, 
                site_thurs, pay_thurs, site_fri, pay_fri, site_sat, pay_sat, 
                site_sun, pay_sun, bank_name, sort_code, account_no, expenses):
    
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Organize data for the loop
    sites = [site_mon, site_tues, site_wed, site_thurs, site_fri, site_sat, site_sun]
    pays = [pay_mon, pay_tues, pay_wed, pay_thurs, pay_fri, pay_sat, pay_sun]
    
    # Financial Logic
    # Formula: $$Total = (Gross - (Gross \times 0.20)) + Expenses$$
    gross = sum(float(x) if x else 0 for x in pays)
    tax = gross * 0.20
    total_due = (gross - tax) + float(expenses if expenses else 0)

    # Styles
    bold_font = Font(bold=True, name="Arial")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    wrap_text = Alignment(wrap_text=True, vertical="top", horizontal="left")
    align_right = Alignment(horizontal="right")

    # Header Section
    ws["A1"], ws["B1"] = "Name", name
    ws["D1"], ws["E1"] = "Invoice No", invoice_no
    ws["D3"], ws["E3"] = "Date", current_date
    ws["A1"].font = ws["D1"].font = ws["D3"].font = bold_font
    ws["E1"].alignment = ws["E3"].alignment = align_right

    # Recipient Section (Avant Compliance Solutions)
    ws["A4"] = "TO"
    ws["A4"].font = bold_font
    ws["A5"] = "Avant Compliance\nSolutions" 
    ws["A5"].alignment = wrap_text
    ws["A9"], ws["A10"], ws["A11"] = "34 Grove Park", "Rainham", "RM13 7DA"

    # Worker Compliance IDs
    ws["D9"], ws["E9"] = "UTR", utr
    ws["D11"], ws["E11"] = "NIN", nin
    ws["D9"].font = ws["D11"].font = bold_font
    ws["E9"].alignment = ws["E11"].alignment = align_right

    ws["A16"], ws["B16"] = "Week Ending", week_ending
    ws["A16"].font = bold_font

    # The 7-Day Work Grid
    day_labels = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    current_row = 19
    for i in range(7):
        ws[f"A{current_row}"] = day_labels[i]
        ws[f"B{current_row}"] = sites[i]
        ws[f"A{current_row+1}"] = "Job Name"
        ws[f"B{current_row+1}"] = f"£{float(pays[i] if pays[i] else 0):.2f}"
        ws[f"A{current_row}"].font = bold_font
        current_row += 3

    # Totals Section
    ws["A40"], ws["B40"] = "Gross", f"£{gross:.2f}"
    ws["A41"], ws["B41"] = "        -20%", f"£{tax:.2f}"
    ws["A42"], ws["B42"] = "Expenses", f"£{float(expenses if expenses else 0):.2f}"
    ws["A44"], ws["B44"] = "Total", f"£{total_due:.2f}"
    
    ws["A40"].font = ws["A42"].font = ws["A44"].font = ws["B44"].font = bold_font
    ws["B44"].fill = yellow_fill # Yellow Highlight

    # Bank Details
    ws["D40"], ws["E40"] = "Bank", bank_name
    ws["D41"], ws["E41"] = "Sort Code", sort_code
    ws["D42"], ws["E42"] = "Account No.", account_no
    ws["D40"].font = ws["D41"].font = ws["D42"].font = bold_font
    ws["E40"].alignment = ws["E41"].alignment = ws["E42"].alignment = align_right

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 25

    filename = f"Invoice_{invoice_no}_{name.replace(' ', '_')}.xlsx"
    wb.save(filename)
    return filename
