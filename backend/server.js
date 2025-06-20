const express = require("express");
const multer = require("multer");
const path = require("path");
const cors = require("cors");

const app = express();
const UPLOAD_DIR = path.join(__dirname, "uploads");
 app.use(cors());
// Configurar multer para manejar la carga de archivos
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});


const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    if (ext !== ".xlsx" && ext !== ".csv") {
      return cb(new Error("Formato de archivo no válido. Solo se permiten .xlsx y .csv"));
    }
    cb(null, true);
  },
});

// Crear la carpeta de uploads si no existe
const fs = require("fs");
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR);
}

// Endpoint para subir archivos
app.post("/upload", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ detail: "No se ha cargado ningún archivo." });
  }
  res.json({ message: `Archivo ${req.file.originalname} cargado exitosamente` });
});

// Manejo de errores
app.use((err, req, res, next) => {
  if (err) {
    return res.status(400).json({ detail: err.message });
  }
  next();
});

// Iniciar el servidor
const PORT = 8000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});