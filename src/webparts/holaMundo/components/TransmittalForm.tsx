import * as React from 'react';
import { useState } from 'react';
import styles from './TransmittalForm.module.scss';

interface TransmittalFormProps {
  selectedFiles: File[]; // Prop para recibir los archivos seleccionados
}

const TransmittalForm: React.FC<TransmittalFormProps> = ({ selectedFiles }) => {
  const [selectedTransmittal, setSelectedTransmittal] = useState<string>('');
  const [selectedAction, setSelectedAction] = useState<string>('');
  const [selectedResponsable, setSelectedResponsable] = useState<string>('');
  const [selectedNumeroTransmittal, setSelectedNumeroTransmittal] = useState<string>('');
  const [selectedDestinatario, setSelectedDestinatario] = useState<string>('');
  const [isSubmitDisabled, setIsSubmitDisabled] = useState<boolean>(true);

  const handleExit = (): void => {
    setSelectedTransmittal('');
    setSelectedAction('');
    setSelectedResponsable('');
    setSelectedNumeroTransmittal('');
    setSelectedDestinatario('');
    setIsSubmitDisabled(true);
  };

  const handleSubmit = (): void => {
    if (!isSubmitDisabled) {
      console.log('Archivos seleccionados:', selectedFiles);
      console.log('Transmittal enviado con éxito con los siguientes detalles:', {
        selectedTransmittal,
        selectedAction,
        selectedResponsable,
        selectedNumeroTransmittal,
        selectedDestinatario,
      });
      handleExit();
    }
  };
  
  const checkSubmitEnabled = (): void => {
    console.log("Checking if submit should be enabled:", {
      selectedTransmittal,
      selectedAction,
      selectedResponsable,
      selectedDestinatario,
    });
  
    if (selectedTransmittal === 'Construcción') {
      setIsSubmitDisabled(!(selectedAction && selectedResponsable));
    } else if (selectedTransmittal === 'Distribución') {
      setIsSubmitDisabled(!(selectedAction && selectedDestinatario));
    } else if (selectedTransmittal === 'Devolución') {
      setIsSubmitDisabled(!selectedAction);
    }
  };

  const handleTransmittalChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    setSelectedTransmittal(event.target.value);
    setSelectedAction('');
    setSelectedResponsable('');
    setSelectedNumeroTransmittal('');
    setSelectedDestinatario('');
    setIsSubmitDisabled(true);
  };

  const handleActionChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    setSelectedAction(event.target.value);
    checkSubmitEnabled();
  };

  const handleResponsableChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    setSelectedResponsable(event.target.value);
    checkSubmitEnabled();
  };

  const handleDestinatarioChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    setSelectedDestinatario(event.target.value);
    checkSubmitEnabled();
  };

  const handleNumeroTransmittalChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    setSelectedNumeroTransmittal(event.target.value);
    checkSubmitEnabled();
  };

  return (
    <div className={styles.transmittalForm}>
      <h2>Enviar a Transmittal</h2>
      <div className={styles.formGroup}>
        <label htmlFor="transmittal">Tipo de Transmittal</label>
        <select
          id="transmittal"
          value={selectedTransmittal}
          onChange={handleTransmittalChange}
          className={styles.dropdown}
        >
          <option value="" disabled>Seleccione una opción</option>
          <option value="Devolución">Devolución</option>
          <option value="Distribución">Distribución</option>
          <option value="Construcción">Construcción</option>
        </select>
      </div>

      {selectedTransmittal === 'Devolución' && (
        <div className={styles.formGroup}>
          <label htmlFor="action">Acción</label>
          <select
            id="action"
            value={selectedAction}
            onChange={handleActionChange}
            className={styles.dropdown}
          >
            <option value="" disabled>Seleccione una opción</option>
            <option value="Recibido Aprobado">Recibido Aprobado</option>
            <option value="Recibido Aprobado con comentarios">Recibido Aprobado con comentarios</option>
            <option value="Recibido con comentarios">Recibido con comentarios</option>
          </select>
        </div>
      )}

      {selectedTransmittal === 'Distribución' && (
        <>
          <div className={styles.formGroup}>
            <label htmlFor="action">Acción</label>
            <select
              id="action"
              value={selectedAction}
              onChange={handleActionChange}
              className={styles.dropdown}
            >
              <option value="" disabled>Seleccione una opción</option>
              <option value="Certificado para Construcción">Certificado para Construcción</option>
              <option value="Emitido para Revisión y Comentarios">Emitido para Revisión y Comentarios</option>
              <option value="Información">Información</option>
              <option value="Licitación">Licitación</option>
            </select> 
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="destinatario">Destinatario</label>
            <select
              id="destinatario"
              value={selectedDestinatario}
              onChange={handleDestinatarioChange}
              className={styles.dropdown}
            >
              <option value="" disabled>Seleccione un destinatario</option>
              <option value="ARAUCO">ARAUCO</option>
            </select>
          </div>
        </>
      )}

      {selectedTransmittal === 'Construcción' && (
        <>
          <div className={styles.formGroup}>
            <label htmlFor="action">Acción</label>
            <select
              id="action"
              value={selectedAction}
              onChange={handleActionChange}
              className={styles.dropdown}
            >
              <option value="" disabled>Seleccione una opción</option>
              <option value="Emitido para Construcción">Emitido para Construcción</option>
            </select>
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="responsable">Responsable</label>
            <select
              id="responsable"
              value={selectedResponsable}
              onChange={handleResponsableChange}
              className={styles.dropdown}
            >
              <option value="" disabled>Seleccione un responsable</option>
              <option value="Cesar Jimenez">Cesar Jimenez</option>
              <option value="Claudia SANTANA">Claudia SANTANA</option>
              <option value="Dennis Rivas">Dennis Rivas</option>
            </select>
          </div>

          <div className={styles.formGroup}>
            <label htmlFor="numeroTransmittal">Número Transmittal</label>
            <select
              id="numeroTransmittal"
              value={selectedNumeroTransmittal}
              onChange={handleNumeroTransmittalChange}
              className={styles.dropdown}
            >
              <option value="" disabled>Seleccione un número de transmittal (opcional)</option>
            </select>
          </div>
        </>
      )}

      <div className={styles.buttonGroup}>
        <button
          onClick={handleSubmit}
          className={styles.acceptButton}
          disabled={isSubmitDisabled}
        >
          Aceptar
        </button>
        <button onClick={handleExit} className={styles.cancelButton}>
          Salir
        </button>
      </div>
    </div>
  );
};

export default TransmittalForm;
