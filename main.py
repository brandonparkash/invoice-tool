import os
from fastapi import FastAPI, Request, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from invoice_excel import func_invoice

# Initialising the server
app = FastAPI()

# Mount templates and static files
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        "index.html", {
            "request": request,
            "title": "Subcontractor Invoice Tool"
        }
    )

@app.post("/", response_class=HTMLResponse)
async def handle_invoice(
    request: Request,
    background_tasks: BackgroundTasks,
    name: str = Form(...),
    invoice_no: str = Form(...),
    current_date: str = Form(...),
    utr: str = Form(...),
    nin: str = Form(...),
    company_name: str = Form(...),
    company_road: str = Form(...),
    company_city: str = Form(...),
    company_postcode: str = Form(...),
    week_ending: str = Form(...),
    # Optional fields default to empty strings or 0
    site_mon: str = Form(""), pay_mon: str = Form("0"),
    site_tues: str = Form(""), pay_tues: str = Form("0"),
    site_wed: str = Form(""), pay_wed: str = Form("0"),
    site_thurs: str = Form(""), pay_thurs: str = Form("0"),
    site_fri: str = Form(""), pay_fri: str = Form("0"),
    site_sat: str = Form(""), pay_sat: str = Form("0"),
    site_sun: str = Form(""), pay_sun: str = Form("0"),
    bank_name: str = Form(...),
    sort_code: str = Form(...),
    account_no: str = Form(...),
    expenses: str = Form("0")
):
    # Generate the file
    fname = func_invoice(name, invoice_no, current_date, utr, nin, week_ending, company_name, company_road,
                        company_city, company_postcode, site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed,
                        site_thurs, pay_thurs, site_fri, pay_fri, site_sat, pay_sat, site_sun, pay_sun, bank_name,
                        sort_code, account_no, expenses)

    file_path = Path.cwd() / fname
    
    # Schedule file deletion after sending
    background_tasks.add_task(os.remove, file_path)

    return FileResponse(
        path=file_path,
        filename=fname,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
