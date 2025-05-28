from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
import shutil
import os
import uuid

from tools.balance_proyectado import process_file as process_balance
from tools.facturacion import process_file as process_facturacion
from tools.inventario import process_file as process_inventario

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key="clave_super_secreta")
app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/img", StaticFiles(directory="static/img"), name="img")

templates = Jinja2Templates(directory="templates")

UPLOAD_FOLDER = "uploads"
DOWNLOAD_FOLDER = "downloads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

PROCESSORS = {
    "balance-proyectado": process_balance,
    "facturacion": process_facturacion,
    "inventario": process_inventario
}

@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse("dashboard.html", {"request": request})

@app.post("/procesar")
async def procesar_archivo(
    request: Request,
    file: UploadFile = File(...),
    tool: str = Form(...),
    fileType: str = Form(...)
):
    try:
        extension = os.path.splitext(file.filename)[1]
        temp_filename = f"{uuid.uuid4().hex}{extension}"
        temp_path = os.path.join(UPLOAD_FOLDER, temp_filename)

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        procesador = PROCESSORS.get(tool)
        if not procesador:
            return {"error": "Herramienta no encontrada"}

        output_path = procesador(temp_path)
        output_name = os.path.basename(output_path)
        final_path = os.path.join(DOWNLOAD_FOLDER, output_name)

        shutil.move(output_path, final_path)

        return FileResponse(final_path, filename=output_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return {"error": f"Error al procesar archivo: {str(e)}"}
