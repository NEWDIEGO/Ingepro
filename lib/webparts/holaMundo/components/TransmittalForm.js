import * as React from 'react';
import { useState } from 'react';
import styles from './TransmittalForm.module.scss';
var TransmittalForm = function (_a) {
    var selectedFiles = _a.selectedFiles;
    var _b = useState(''), selectedTransmittal = _b[0], setSelectedTransmittal = _b[1];
    var _c = useState(''), selectedAction = _c[0], setSelectedAction = _c[1];
    var _d = useState(''), selectedResponsable = _d[0], setSelectedResponsable = _d[1];
    var _e = useState(''), selectedNumeroTransmittal = _e[0], setSelectedNumeroTransmittal = _e[1];
    var _f = useState(''), selectedDestinatario = _f[0], setSelectedDestinatario = _f[1];
    var _g = useState(true), isSubmitDisabled = _g[0], setIsSubmitDisabled = _g[1];
    var handleExit = function () {
        setSelectedTransmittal('');
        setSelectedAction('');
        setSelectedResponsable('');
        setSelectedNumeroTransmittal('');
        setSelectedDestinatario('');
        setIsSubmitDisabled(true);
    };
    var handleSubmit = function () {
        if (!isSubmitDisabled) {
            console.log('Archivos seleccionados:', selectedFiles);
            console.log('Transmittal enviado con éxito con los siguientes detalles:', {
                selectedTransmittal: selectedTransmittal,
                selectedAction: selectedAction,
                selectedResponsable: selectedResponsable,
                selectedNumeroTransmittal: selectedNumeroTransmittal,
                selectedDestinatario: selectedDestinatario,
            });
            handleExit();
        }
    };
    var checkSubmitEnabled = function () {
        console.log("Checking if submit should be enabled:", {
            selectedTransmittal: selectedTransmittal,
            selectedAction: selectedAction,
            selectedResponsable: selectedResponsable,
            selectedDestinatario: selectedDestinatario,
        });
        if (selectedTransmittal === 'Construcción') {
            setIsSubmitDisabled(!(selectedAction && selectedResponsable));
        }
        else if (selectedTransmittal === 'Distribución') {
            setIsSubmitDisabled(!(selectedAction && selectedDestinatario));
        }
        else if (selectedTransmittal === 'Devolución') {
            setIsSubmitDisabled(!selectedAction);
        }
    };
    var handleTransmittalChange = function (event) {
        setSelectedTransmittal(event.target.value);
        setSelectedAction('');
        setSelectedResponsable('');
        setSelectedNumeroTransmittal('');
        setSelectedDestinatario('');
        setIsSubmitDisabled(true);
    };
    var handleActionChange = function (event) {
        setSelectedAction(event.target.value);
        checkSubmitEnabled();
    };
    var handleResponsableChange = function (event) {
        setSelectedResponsable(event.target.value);
        checkSubmitEnabled();
    };
    var handleDestinatarioChange = function (event) {
        setSelectedDestinatario(event.target.value);
        checkSubmitEnabled();
    };
    var handleNumeroTransmittalChange = function (event) {
        setSelectedNumeroTransmittal(event.target.value);
        checkSubmitEnabled();
    };
    return (React.createElement("div", { className: styles.transmittalForm },
        React.createElement("h2", null, "Enviar a Transmittal"),
        React.createElement("div", { className: styles.formGroup },
            React.createElement("label", { htmlFor: "transmittal" }, "Tipo de Transmittal"),
            React.createElement("select", { id: "transmittal", value: selectedTransmittal, onChange: handleTransmittalChange, className: styles.dropdown },
                React.createElement("option", { value: "", disabled: true }, "Seleccione una opci\u00F3n"),
                React.createElement("option", { value: "Devoluci\u00F3n" }, "Devoluci\u00F3n"),
                React.createElement("option", { value: "Distribuci\u00F3n" }, "Distribuci\u00F3n"),
                React.createElement("option", { value: "Construcci\u00F3n" }, "Construcci\u00F3n"))),
        selectedTransmittal === 'Devolución' && (React.createElement("div", { className: styles.formGroup },
            React.createElement("label", { htmlFor: "action" }, "Acci\u00F3n"),
            React.createElement("select", { id: "action", value: selectedAction, onChange: handleActionChange, className: styles.dropdown },
                React.createElement("option", { value: "", disabled: true }, "Seleccione una opci\u00F3n"),
                React.createElement("option", { value: "Recibido Aprobado" }, "Recibido Aprobado"),
                React.createElement("option", { value: "Recibido Aprobado con comentarios" }, "Recibido Aprobado con comentarios"),
                React.createElement("option", { value: "Recibido con comentarios" }, "Recibido con comentarios")))),
        selectedTransmittal === 'Distribución' && (React.createElement(React.Fragment, null,
            React.createElement("div", { className: styles.formGroup },
                React.createElement("label", { htmlFor: "action" }, "Acci\u00F3n"),
                React.createElement("select", { id: "action", value: selectedAction, onChange: handleActionChange, className: styles.dropdown },
                    React.createElement("option", { value: "", disabled: true }, "Seleccione una opci\u00F3n"),
                    React.createElement("option", { value: "Certificado para Construcci\u00F3n" }, "Certificado para Construcci\u00F3n"),
                    React.createElement("option", { value: "Emitido para Revisi\u00F3n y Comentarios" }, "Emitido para Revisi\u00F3n y Comentarios"),
                    React.createElement("option", { value: "Informaci\u00F3n" }, "Informaci\u00F3n"),
                    React.createElement("option", { value: "Licitaci\u00F3n" }, "Licitaci\u00F3n"))),
            React.createElement("div", { className: styles.formGroup },
                React.createElement("label", { htmlFor: "destinatario" }, "Destinatario"),
                React.createElement("select", { id: "destinatario", value: selectedDestinatario, onChange: handleDestinatarioChange, className: styles.dropdown },
                    React.createElement("option", { value: "", disabled: true }, "Seleccione un destinatario"),
                    React.createElement("option", { value: "ARAUCO" }, "ARAUCO"))))),
        selectedTransmittal === 'Construcción' && (React.createElement(React.Fragment, null,
            React.createElement("div", { className: styles.formGroup },
                React.createElement("label", { htmlFor: "action" }, "Acci\u00F3n"),
                React.createElement("select", { id: "action", value: selectedAction, onChange: handleActionChange, className: styles.dropdown },
                    React.createElement("option", { value: "", disabled: true }, "Seleccione una opci\u00F3n"),
                    React.createElement("option", { value: "Emitido para Construcci\u00F3n" }, "Emitido para Construcci\u00F3n"))),
            React.createElement("div", { className: styles.formGroup },
                React.createElement("label", { htmlFor: "responsable" }, "Responsable"),
                React.createElement("select", { id: "responsable", value: selectedResponsable, onChange: handleResponsableChange, className: styles.dropdown },
                    React.createElement("option", { value: "", disabled: true }, "Seleccione un responsable"),
                    React.createElement("option", { value: "Cesar Jimenez" }, "Cesar Jimenez"),
                    React.createElement("option", { value: "Claudia SANTANA" }, "Claudia SANTANA"),
                    React.createElement("option", { value: "Dennis Rivas" }, "Dennis Rivas"))),
            React.createElement("div", { className: styles.formGroup },
                React.createElement("label", { htmlFor: "numeroTransmittal" }, "N\u00FAmero Transmittal"),
                React.createElement("select", { id: "numeroTransmittal", value: selectedNumeroTransmittal, onChange: handleNumeroTransmittalChange, className: styles.dropdown },
                    React.createElement("option", { value: "", disabled: true }, "Seleccione un n\u00FAmero de transmittal (opcional)"))))),
        React.createElement("div", { className: styles.buttonGroup },
            React.createElement("button", { onClick: handleSubmit, className: styles.acceptButton, disabled: isSubmitDisabled }, "Aceptar"),
            React.createElement("button", { onClick: handleExit, className: styles.cancelButton }, "Salir"))));
};
export default TransmittalForm;
//# sourceMappingURL=TransmittalForm.js.map