from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from pathlib import Path

from invoice_excel import func_invoice


# Initialising the server
app = FastAPI()

# Showing the server the route to the html/css paths
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


# GET route to return html template
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        "index.html", {
            "request": request,
            "title": "Subcontractor Invoice Automation"
        }
    )

# POST route to capture form values
@app.post("/", response_class=HTMLResponse)
async def handle_invoice(
    request: Request,
    name: str = Form(...),
    invoice_no: str = Form(...),
    current_date: str = Form(...),
    utr: str = Form(...),
    nin: str = Form(...),
    week_ending: str = Form(...),
    site_mon: str = Form(""),
    pay_mon: str = Form("0"),
    site_tues: str = Form(""),
    pay_tues: str = Form("0"),
    site_wed: str = Form(""),
    pay_wed: str = Form("0"),
    site_thurs: str = Form(""),
    pay_thurs: str = Form("0"),
    site_fri: str = Form(""),
    pay_fri: str = Form("0"),
    site_sat: str = Form(""),
    pay_sat: str = Form("0"),
    site_sun: str = Form(""),
    pay_sun: str = Form("0"),
    bank_name: str = Form(...),
    sort_code: str = Form(...),
    account_no: str = Form(...),
    expenses: str = Form("0")
):
    fname = func_invoice(name, invoice_no, current_date, utr, nin, week_ending, site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed,
        site_thurs, pay_thurs, site_fri, pay_fri, site_sat, pay_sat, site_sun, pay_sun, bank_name, sort_code, account_no, expenses)

    return FileResponse(
        path=Path.cwd() / fname,
        filename=fname,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )