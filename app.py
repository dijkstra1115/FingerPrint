from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
from modules.report_generator import generate_report

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 允许所有来源，您可以根据需要限制
    allow_credentials=False,
    allow_methods=["*"],  # 允许所有方法
    allow_headers=["*"],  # 允许所有头部
)

# 靜態下載路徑 or 動態下載路徑
# app.mount("/download", StaticFiles(directory="output"), name="download_file")
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    print("Request URL:", request.url)
    print("Available routes:", app.routes)
    print("Static URL:", request.url_for('static', path="css/styles.css"))  # 添加调试信息
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/submit-personality-test", response_class=HTMLResponse)
async def submit_personality_test(
    request: Request,
    user_name: str = Form(...),
    pricing_plan: str = Form(...),
    L1_code: str = Form(...), L1_left_value: int = Form(...), L1_right_value: int = Form(...),
    L2_code: str = Form(...), L2_left_value: int = Form(...), L2_right_value: int = Form(...),
    L3_code: str = Form(...), L3_left_value: int = Form(...), L3_right_value: int = Form(...),
    L4_code: str = Form(...), L4_left_value: int = Form(...), L4_right_value: int = Form(...),
    L5_code: str = Form(...), L5_left_value: int = Form(...), L5_right_value: int = Form(...),
    R1_code: str = Form(...), R1_left_value: int = Form(...), R1_right_value: int = Form(...),
    R2_code: str = Form(...), R2_left_value: int = Form(...), R2_right_value: int = Form(...),
    R3_code: str = Form(...), R3_left_value: int = Form(...), R3_right_value: int = Form(...),
    R4_code: str = Form(...), R4_left_value: int = Form(...), R4_right_value: int = Form(...),
    R5_code: str = Form(...), R5_left_value: int = Form(...), R5_right_value: int = Form(...)
):
    data = {
        "L1_code": L1_code, "L1_left_value": L1_left_value, "L1_right_value": L1_right_value,
        "L2_code": L2_code, "L2_left_value": L2_left_value, "L2_right_value": L2_right_value,
        "L3_code": L3_code, "L3_left_value": L3_left_value, "L3_right_value": L3_right_value,
        "L4_code": L4_code, "L4_left_value": L4_left_value, "L4_right_value": L4_right_value,
        "L5_code": L5_code, "L5_left_value": L5_left_value, "L5_right_value": L5_right_value,
        "R1_code": R1_code, "R1_left_value": R1_left_value, "R1_right_value": R1_right_value,
        "R2_code": R2_code, "R2_left_value": R2_left_value, "R2_right_value": R2_right_value,
        "R3_code": R3_code, "R3_left_value": R3_left_value, "R3_right_value": R3_right_value,
        "R4_code": R4_code, "R4_left_value": R4_left_value, "R4_right_value": R4_right_value,
        "R5_code": R5_code, "R5_left_value": R5_left_value, "R5_right_value": R5_right_value,
    }

    generate_report(user_name, pricing_plan, data)
    file_path = request.url_for('download_file', filename=f"{user_name}_{pricing_plan}.docx")
    return templates.TemplateResponse("download.html", {"request": request, "file_path": file_path})

@app.post("/generate_report")
async def generate_report_api(request: Request):
    body = await request.json()
    user_name = body.get('user_name')
    pricing_plan = body.get('pricing_plan')
    data = body.get('data')

    generate_report(user_name, pricing_plan, data)
    report_url = str(request.url_for('download_file', filename=f"{user_name}_{pricing_plan}.docx"))
    return JSONResponse(content={"report_url": report_url})

@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join("output", filename)
    return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=filename)

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000, log_level="error")