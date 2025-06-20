from fastapi import FastAPI, File, UploadFile, HTTPException
import os

app = FastAPI()

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

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

