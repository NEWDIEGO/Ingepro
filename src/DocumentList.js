import React, { useEffect, useState } from 'react';
import { sp } from "@pnp/sp/presets/all";

const DocumentList = () => {
    const [documents, setDocuments] = useState([]);
    const [selectedFiles, setSelectedFiles] = useState([]); // Estado para manejar archivos seleccionados
    const [actionCount, setActionCount] = useState({});

    useEffect(() => {
        sp.setup({
            sp: {
                baseUrl: "https://ntechilespa.sharepoint.com/sites/ingp", // Cambia esta URL a la de tu sitio de SharePoint
            },
        });

        async function fetchData() {
            try {
                // Obtener archivos desde SharePoint
                const documentos = await sp.web.lists.getByTitle("Distribucion").items();
                setDocuments(documentos);
            } catch (error) {
                console.error("❌ Error al obtener datos de SharePoint:", error);
            }
        }

        fetchData();
    }, []);

    const handleFileSelection = (doc) => {
        // Maneja la selección de archivos, agregándolos o removiéndolos de la lista seleccionada
        if (selectedFiles.includes(doc)) {
            setSelectedFiles(selectedFiles.filter(file => file !== doc));
        } else {
            setSelectedFiles([...selectedFiles, doc]);
        }
    };

    const handleAssignTransmittal = (archivo, accion, destinatario) => {
        const newData = { archivo, accion, destinatario };
    
        // Obtener datos previos de localStorage
        const storedData = JSON.parse(localStorage.getItem("transmittalData")) || [];
    
        // Agregar nuevo dato
        storedData.push(newData);
    
        // Guardar en localStorage
        localStorage.setItem("transmittalData", JSON.stringify(storedData));
    
        console.log("Datos guardados en localStorage:", storedData);
    };
    

    return (
        <div>
            <h3>Lista de Documentos</h3>
            <ul>
                {documents.map((doc) => (
                    <li key={doc.UniqueId}>
                        <input 
                            type="checkbox" 
                            onChange={() => handleFileSelection(doc)} 
                            checked={selectedFiles.includes(doc)}
                        />
                        {doc.Name}
                    </li>
                ))}
            </ul>
            <button 
                onClick={handleAssignTransmittal} 
                disabled={selectedFiles.length === 0} // Habilitar solo si hay archivos seleccionados
            >
                Asignar a Transmittal
            </button>
        </div>
    );
};

export default DocumentList;
