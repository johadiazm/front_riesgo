import React, { useState } from "react";
import '../styles/FileUpload.css';

const FileUpload = () => {
  const [fileName, setFileName] = useState("");
  const [pdfName, setPdfName] = useState("");
  const [wordName, setWordName] = useState("");
  const [loading, setLoading] = useState(false);
  const baseURL = "https://proyecto-riesgo-psicosocial.onrender.com";

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    setFileName(file ? file.name : "");
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
      const response = await fetch(`${baseURL}/upload/`, {
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
      alert("Ocurrió un error al cargar el archivo.");
    }
  };

  // Genera ambos informes (PDF y Word) al mismo tiempo
  const handleGenerarAmbosInformes = async () => {
    setLoading(true);
    if (!fileName) {
      alert("Primero debes subir un archivo.");
      setLoading(false);
      return;
    }
    try {
      // Generar PDF
      const responsePDF = await fetch(`${baseURL}/generar_informe/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: fileName, formato: "pdf" }),
      });
      // Generar Word
      const responseWord = await fetch(`${baseURL}/generar_informe/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: fileName, formato: "word" }),
      });

      if (!responsePDF.ok || !responseWord.ok) {
        alert("Error al generar uno de los informes.");
      } else {
        const dataPDF = await responsePDF.json();
        const dataWord = await responseWord.json();
        setPdfName(dataPDF.informe);
        setWordName(dataWord.informe);
        alert("Informes generados. Ahora puedes descargarlos.");
      }
    } catch (error) {
      alert("Ocurrió un error al generar los informes.");
    }
    setLoading(false);
  };

  // Descarga el informe seleccionado
  const handleDescargarInforme = (formato) => {
    let informe = formato === "pdf" ? pdfName : wordName;
    if (informe) {
      window.open(`${baseURL}/descargar_informe/${informe}?formato=${formato}`, "_blank");
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

      {/* Un solo botón para generar ambos informes */}
      <button onClick={handleGenerarAmbosInformes}>Generar Informe</button>

      {loading && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", marginTop: 30 }}>
          <img src="/ventana.gif" alt="Cargando..." style={{ width: 80, height: 80 }} />
          <p>Generando informe, por favor espere...</p>
        </div>
      )}

      {/* Botones para descargar PDF o Word solo si ambos existen */}
      {pdfName && wordName && (
        <>
          <button onClick={() => handleDescargarInforme('pdf')}>
            Descargar Informe PDF
          </button>
          <button onClick={() => handleDescargarInforme('word')}>
            Descargar Informe Word
          </button>
          {/* Vista previa solo del PDF */}
          <iframe
            src={`${baseURL}/descargar_informe/${pdfName}`}
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