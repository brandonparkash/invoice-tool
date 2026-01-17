import openpyxl
import re


# Invoice maker function.
def func_invoice(name, invoice_no, current_date, utr, nin, week_ending, site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed,
            site_thurs, pay_thurs, site_fri, pay_fri, site_sat, pay_sat, site_sun, pay_sun, bank_name, sort_code, account_no, expenses):
    
    wb = openpyxl.Workbook()
    ws = wb.active
    payment_info = func_pay(int(pay_mon), int(pay_tues), int(pay_wed), int(pay_thurs), int(pay_fri), int(pay_sat), int(pay_sun))

    ws["A1"] = "Name"
    ws["D1"] = "Invoice No"
    ws["D3"] = "Date"
    ws["D9"] = "UTR"
    ws["D11"] = "NIN"
    ws["A16"] = "Week Ending"
    
    ws["B1"] = name
    ws["E1"] = invoice_no
    ws["E3"] = current_date
    ws["E9"] = utr
    ws["E11"] = nin
    ws["B16"] = week_ending

    # Form headers
    ws["A20"] = "Job Name"
    ws["A23"] = "Job Name"
    ws["A26"] = "Job Name"
    ws["A29"] = "Job Name"
    ws["A32"] = "Job Name"
    ws["A35"] = "Job Name"
    ws["A38"] = "Job Name"
    
    # Custom information
    ws["A4"] = "TO"
    ws["A5"] = "*****"
    ws["A6"] = "**********"
    ws["A7"] = "*********"

    ws["A9"] = "** ***** ****"
    ws["A10"] = "****"
    ws["A11"] = "*******"
    ws["A12"] = "**** ***"

    # Job information for each day
    ws["A19"] = "Monday"
    ws["B19"] = site_mon
    ws["B20"] = (f"£{pay_mon}")

    ws["A22"] = "Tuesday"
    ws["B22"] = site_tues
    ws["B23"] = (f"£{pay_tues}")

    ws["A25"] = "Wednesday"
    ws["B25"] = site_wed
    ws["B26"] = (f"£{pay_wed}")

    ws["A28"] = "Thursday"
    ws["B28"] = site_thurs
    ws["B29"] = (f"£{pay_thurs}")

    ws["A31"] = "Friday"
    ws["B31"] = site_fri
    ws["B32"] = (f"£{pay_fri}")

    ws["A34"] = "Saturday"
    ws["B34"] = site_sat
    ws["B35"] = (f"£{pay_sat}")

    ws["A37"] = "Sunday"
    ws["B37"] = site_sun
    ws["B38"] = (f"£{pay_sun}")

    # Calculating pay
    ws["A40"] = "Gross"
    ws["A41"] = "Tax (20%)"
    ws["A42"] = "Expenses"
    ws["A44"] = "Total"

    ws["B40"] = (f"£{payment_info["gross"]}")
    ws["B41"] = (f"£{payment_info["tax"]}")
    ws["B42"] = (f"£{expenses}")
    ws["B44"] = (f"£{payment_info["take_home"]}")

    # Bank Details
    ws["D40"] = "Name"
    ws["D41"] = "Sort Code"
    ws["D42"] = "Account No."

    ws["E40"] = bank_name
    ws["E41"] = sort_code
    ws["E42"] = account_no

    file_name = re.sub(r"\s+", "_", name).lower()
    wb.save(f"{file_name}invoice.xlsx")
    return f"{file_name}invoice.xlsx"


# Calculator to calculate take home pay and tax.
def func_pay(mon, tues, wed, thurs, fri, sat, sun):
    gross = sum([mon, tues, wed, thurs, fri, sat, sun])
    tax = (gross / 10) * 2
    take_home = gross - tax

    return {"gross": gross, "tax": tax, "take_home": take_home}

