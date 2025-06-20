import React, { useState } from "react";
import '../styles/FileUpload.css';

const FileUpload = () => {
  const [fileName, setFileName] = useState("");
  const [pdfName, setPdfName] = useState("");

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      
    } else {
      setFileName("");
    }
  };

  const handleFileUpload = async (event) => {
    event.preventDefault();
    const fileInput = document.getElementById("file-upload");
    const file = fileInput.files[0];

    if (!file) {
      alert("Por favor, selecciona un archivo.");
      return;
    }

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("http://localhost:8000/upload/", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        alert(`Error: ${errorData.detail}`);
      } else {
        const data = await response.json();
        alert(data.message);
      }
    } catch (error) {
      console.error("Error al cargar el archivo:", error);
      alert("Ocurrió un error al cargar el archivo.");
    }
  };
  const [loading, setLoading] = useState(false);
  const handleGenerarInforme = async () => {
    setLoading(true);
    if (!fileName) {
      alert("Primero debes subir un archivo.");
      return;
    }
    try {
      const response = await fetch("http://localhost:8000/generar_informe/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: fileName }),
      });
      if (!response.ok) {
        const errorData = await response.json();
        alert(`Error: ${errorData.detail}`);
      } else {
        const data = await response.json();
        setPdfName(data.pdf);
        alert("Informe generado. Ahora puedes descargarlo.");
      }
    } catch (error) {
      alert("Ocurrió un error al generar el informe.");
    }
    setLoading(false);
  };
const handleDescargarInforme = () => {
    if (pdfName) {
      window.open(`http://localhost:8000/descargar_informe/${pdfName}`, "_blank");
    }
  };

  return (
    <div>
      <form onSubmit={handleFileUpload}>
        <label htmlFor="file-upload">Cargar archivo (.xlsx, .csv): </label>
        <input
          type="file"
          id="file-upload"
          accept=".xlsx, .csv"
          onChange={handleFileChange}
        />
        <button type="submit">Subir archivo</button>
      </form>
      {fileName && <p>Archivo seleccionado: {fileName}</p>}
      <button onClick={handleGenerarInforme}>Generar Informe</button>
      
       {loading && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", marginTop: 30 }}>
          <img src="/ventana.gif" alt="Cargando..." style={{ width: 80, height: 80 }} />
          <p>Generando informe, por favor espere...</p>
        </div>
      )}
      {pdfName && (
  <>
    <button onClick={handleDescargarInforme}>
      Descargar Informe PDF
    </button>
    <iframe
      src={`http://localhost:8000/descargar_informe/${pdfName}`}
      width="100%"
      height="600px"
      title="Vista previa del informe"
      style={{ border: "1px solid #ccc", marginTop: "20px" }}
    />
    
  </>
)}
</div>
  );
};

export default FileUpload;