import os, smtplib, httpx
from fastapi import FastAPI, Request, Form, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from email.message import EmailMessage
from invoice_excel import func_invoice
from pydantic import BaseModel
import json

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# Mount static files - THIS WAS MISSING!
app.mount("/static", StaticFiles(directory="static"), name="static")

# SMTP Configuration using environment variables
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
EMAIL_ADDR = os.getenv("EMAIL_ADDR", "")
EMAIL_PASS = os.getenv("EMAIL_PASS", "")

class PromptRequest(BaseModel):
    prompt: str

def send_email(target, file_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = "Subcontractor Invoice Attached"
        msg['From'] = EMAIL_ADDR
        msg['To'] = target
        msg.set_content("Please find the generated invoice attached.")
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', 
                               subtype='octet-stream', filename=file_path.name)
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDR, EMAIL_PASS)
            server.send_message(msg)
    except Exception as e:
        print(f"Email sending failed: {e}")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/slm-process")
async def process_slm(data: PromptRequest):
    try:
        # Sends prompt to local Phi3 via Ollama
        async with httpx.AsyncClient() as client:
            response = await client.post("http://localhost:11434/api/generate", json={
                "model": "invoice-phi3",
                "prompt": data.prompt,
                "stream": False,
                "format": "json"
            }, timeout=45.0)
            
            if response.status_code != 200:
                raise HTTPException(status_code=500, detail="AI processing failed")
                
        return JSONResponse(content=response.json())
    except httpx.TimeoutException:
        raise HTTPException(status_code=504, detail="AI processing timeout")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")

@app.post("/generate")
async def handle_invoice(
    background_tasks: BackgroundTasks,
    name: str = Form(...), invoice_no: str = Form(...), current_date: str = Form(...),
    utr: str = Form(...), nin: str = Form(...), week_ending: str = Form(...),
    site_mon: str = Form(""), pay_mon: str = Form("0"),
    site_tues: str = Form(""), pay_tues: str = Form("0"),
    site_wed: str = Form(""), pay_wed: str = Form("0"),
    site_thurs: str = Form(""), pay_thurs: str = Form("0"),
    site_fri: str = Form(""), pay_fri: str = Form("0"),
    site_sat: str = Form(""), pay_sat: str = Form("0"),
    site_sun: str = Form(""), pay_sun: str = Form("0"),
    bank_name: str = Form(...), sort_code: str = Form(...), account_no: str = Form(...),
    expenses: str = Form("0"), delivery: str = Form("download"), email_to: str = Form("")
):
    try:
        fname = func_invoice(name, invoice_no, current_date, utr, nin, week_ending, 
                            site_mon, pay_mon, site_tues, pay_tues, site_wed, pay_wed, 
                            site_thurs, pay_thurs, site_fri, pay_fri, site_sat, pay_sat, 
                            site_sun, pay_sun, bank_name, sort_code, account_no, expenses)
        
        file_path = Path.cwd() / fname
        
        if delivery == "email" and email_to:
            background_tasks.add_task(send_email, email_to, file_path)
            background_tasks.add_task(os.remove, file_path)
            return JSONResponse(content={"status": "Success", "message": f"Invoice sent to {email_to}"})
        
        return FileResponse(path=file_path, filename=fname, 
                          background=background_tasks.add_task(os.remove, file_path))
                          
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Invoice generation failed: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
