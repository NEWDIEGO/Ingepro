import React, { useState } from 'react';
import TransmittalForm from './TransmittalForm';
import './TransmittalComponent.css'; // Importar los estilos

const TransmittalComponent = () => {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [isFormOpen, setIsFormOpen] = useState(false);

  // Simula la selección de archivos en la biblioteca de SharePoint
  const handleFileSelection = (files) => {
    setSelectedFiles(files);
  };

  // Habilitar o deshabilitar el botón
  const isButtonDisabled = selectedFiles.length === 0;

  // Abrir el formulario de Transmittal
  const openForm = () => {
    setIsFormOpen(true);
  };

  // Cerrar el formulario de Transmittal
  const closeForm = () => {
    setIsFormOpen(false);
  };

  return (
    <div className="transmittal-component">
      <h1>Asignar a Transmittal</h1>

      {/* Ejemplo de lista de archivos (simulación de selección) */}
      <div className="file-list">
        <p>Selecciona los archivos:</p>
        <ul>
          <li><input type="checkbox" onChange={() => handleFileSelection(['file1'])}/> Archivo 1</li>
          <li><input type="checkbox" onChange={() => handleFileSelection(['file2'])}/> Archivo 2</li>
          <li><input type="checkbox" onChange={() => handleFileSelection(['file3'])}/> Archivo 3</li>
        </ul>
      </div>

      <button
        onClick={openForm}
        disabled={isButtonDisabled}
        className="assign-button"
      >
        Asignar a Transmittal
      </button>

      {isFormOpen && (
        <div className="transmittal-form-wrapper">
          <TransmittalForm
            selectedFiles={selectedFiles}
            onClose={closeForm}
          />
        </div>
      )}
    </div>
  );
};

export default TransmittalComponent;
