from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import os
from pydantic import BaseModel
from fastapi.responses import FileResponse
from views import generar_graficos_y_pdf
app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
UPLOAD_DIR = "uploads"
REPORTS_DIR = "reports"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)
class InformeRequest(BaseModel):
    filename: str
@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    # Validar extensión del archivo
    if not file.filename.endswith((".xlsx", ".csv")):
        raise HTTPException(status_code=400, detail="Formato de archivo no válido. Solo se permiten .xlsx y .csv")

    # Guardar archivo temporalmente
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as f:
        f.write(await file.read())
    return {"message": f"Archivo {file.filename} cargado exitosamente"}

@app.post("/generar_informe/")
def generar_informe(data: InformeRequest):
    filename = data.filename
    file_path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    pdf_name = generar_graficos_y_pdf(file_path, filename, REPORTS_DIR)
    return {"pdf": pdf_name}

@app.get("/descargar_informe/{pdf_name}")
def descargar_informe(pdf_name: str):
    pdf_path = os.path.join(REPORTS_DIR, pdf_name)
    if not os.path.exists(pdf_path):
        raise HTTPException(status_code=404, detail="Informe no encontrado")
    return FileResponse(
        pdf_path,
        media_type='application/pdf',
        filename=pdf_name,
        headers={"Content-Disposition": f'inline; filename="{pdf_name}"'})
      