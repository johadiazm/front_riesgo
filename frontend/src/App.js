import React from "react";
import FileUpload from "./components/FileUpload";
import Header from "./components/Header";


function App() {
  return (
    <div className="App">
      <Header />
      <div className="upload-container">
        <h3>Subir Archivos</h3>
      <FileUpload />
      </div>
    </div>
  );
}

export default App;
