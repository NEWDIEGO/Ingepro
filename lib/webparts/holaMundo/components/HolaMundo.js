import * as React from 'react';
import { useState } from 'react';
import TransmittalForm from './TransmittalForm';
var AssignToTransmittalComponent = function () {
    var _a = useState([]), selectedFiles = _a[0], setSelectedFiles = _a[1];
    var _b = useState(true), isButtonDisabled = _b[0], setIsButtonDisabled = _b[1];
    var handleFileSelection = function (files) {
        var fileObjects = files.map(function (file) { return new File([], file.name); });
        setSelectedFiles(fileObjects);
        setIsButtonDisabled(fileObjects.length === 0);
    };
    return (React.createElement("div", null,
        React.createElement("button", { onClick: function () { return handleFileSelection([{ name: 'Archivo1.pdf', url: '/documentos/Archivo1.pdf' }]); }, disabled: isButtonDisabled }, "Asignar a Transmittal"),
        selectedFiles.length > 0 && (React.createElement(TransmittalForm, { selectedFiles: selectedFiles }))));
};
export default AssignToTransmittalComponent;
//# sourceMappingURL=HolaMundo.js.map