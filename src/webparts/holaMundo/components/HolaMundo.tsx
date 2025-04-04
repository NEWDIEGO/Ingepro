import * as React from 'react';
import { useState } from 'react';
import TransmittalForm from './TransmittalForm';

interface IFile {
  name: string;
  url: string;
}

const AssignToTransmittalComponent: React.FC = () => {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [isButtonDisabled, setIsButtonDisabled] = useState(true);

  const handleFileSelection = (files: IFile[]): void => { // Tipo de retorno `void` agregado
    const fileObjects = files.map(file => new File([], file.name));
    setSelectedFiles(fileObjects);
    setIsButtonDisabled(fileObjects.length === 0);
  };

  return (
    <div>
      <button
        onClick={() => handleFileSelection([{ name: 'Archivo1.pdf', url: '/documentos/Archivo1.pdf' }])}
        disabled={isButtonDisabled}
      >
        Asignar a Transmittal
      </button>

      {selectedFiles.length > 0 && (
        <TransmittalForm selectedFiles={selectedFiles} />
      )}
    </div>
  );
};

export default AssignToTransmittalComponent;
