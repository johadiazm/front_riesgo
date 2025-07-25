from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import os
from views import generar_informe_en_word, generar_graficos_y_pdf
from pydantic import BaseModel
from fastapi.responses import FileResponse
app = FastAPI()
from fastapi import Query

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://front-riesgo-bek.vercel.app"],
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
    formato: str = "pdf" 
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
    formato = data.formato
    file_path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    if formato == "pdf":
        informe_name = generar_graficos_y_pdf(file_path, filename, REPORTS_DIR)
    elif formato == "word":
        informe_name = generar_informe_en_word(file_path, None, filename, REPORTS_DIR)
    else:
        raise HTTPException(status_code=400, detail="Formato no soportado")
    return {"informe": informe_name}


@app.get("/descargar_informe/{nombre}")
def descargar_informe(nombre: str, formato: str = Query("pdf")):
    base, _ = os.path.splitext(nombre)
    if formato == "pdf":
        ext = ".pdf"
        media_type = "application/pdf"
    elif formato == "word":
        ext = ".docx"
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        raise HTTPException(status_code=400, detail="Formato no soportado")

    informe_name = base + ext
    informe_path = os.path.join(REPORTS_DIR, informe_name)
    if not os.path.exists(informe_path):
        raise HTTPException(status_code=404, detail="Informe no encontrado")

    return FileResponse(
        informe_path,
        media_type=media_type,
        filename=informe_name,
        headers={"Content-Disposition": f'inline; filename="{informe_name}"'}
    )