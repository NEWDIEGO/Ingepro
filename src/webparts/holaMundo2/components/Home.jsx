async function asignarADistribucion() {
  console.log("Asignando archivos a distribución...");

  // Obtener valores seleccionados
  const actionValue = document.getElementById("action").value;
  const recipientValue = document.getElementById("recipient").value;
  const checkboxes = document.querySelectorAll('.checkbox:checked');

  // Verificar que hay archivos seleccionados
  if (actionValue && recipientValue && checkboxes.length > 0) {
      let listaDeArchivos = [];

      checkboxes.forEach((checkbox) => {
          const row = checkbox.closest("tr");
          if (row) {
              const nombreArchivo = row.cells[1].innerText;
              listaDeArchivos.push(nombreArchivo);
          }
      });

      console.log("Archivos seleccionados:", listaDeArchivos);

      // URL de la API REST de SharePoint para insertar en la lista "Distribución"
      let url = "https://ntechilespa.sharepoint.com/sites/ingp/_api/web/lists/getbytitle('Distribucion')/items";

      // Definir encabezados para la autenticación y formato de datos
      let headers = {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value
      };

      // Insertar cada archivo en SharePoint
      for (let i = 0; i < listaDeArchivos.length; i++) {
          let data = {
              __metadata: { type: "SP.Data.DistribucionListItem" }, // Nombre de la lista en SharePoint
              Nombre: listaDeArchivos[i], // Nombre del archivo
              Accion: actionValue, // Acción asignada
              Destinatario: recipientValue // Destinatario asignado
          };

          try {
              let response = await fetch(url, {
                  method: "POST",
                  headers: headers,
                  body: JSON.stringify(data),
                  credentials: "include"
              });

              if (!response.ok) {
                  throw new Error("Error al almacenar datos en SharePoint: " + response.statusText);
              }

              let result = await response.json();
              console.log("Datos almacenados correctamente en SharePoint:", result);

          } catch (error) {
              console.error("Error al guardar datos en SharePoint:", error);
              alert("Hubo un problema al guardar los datos.");
              return;
          }
      }

      alert("Archivos asignados correctamente en SharePoint");
      window.location.reload(); // Refrescar la página para ver los cambios

  } else {
      alert("Debe seleccionar una acción, un destinatario y al menos un archivo.");
  }
}

// Agregar el evento al botón "Asignar a Distribución"
document.getElementById('assignDistributionButton').addEventListener("click", async () => {
    console.log("Botón 'Asignar a Distribución' presionado.");

    const actionValue = document.getElementById("action").value;
    const recipientValue = document.getElementById("recipient").value;
    const checkboxes = document.querySelectorAll('.checkbox:checked');

    // Definir numeroTransmittal correctamente antes de su uso
    let numeroTransmittal = document.getElementById("numeroTransmittal")?.value || "No asignado";

    if (!numeroTransmittal) {
        console.error("Error: numeroTransmittal tiene un valor vacío.");
        return;
    }

    if (actionValue && recipientValue && checkboxes.length > 0) {
        checkboxes.forEach((checkbox) => {
            const row = checkbox.closest("tr");
            if (row) {
                const nombreArchivo = row.cells[1].innerText;
                row.cells[2].innerText = "Sí";
                row.cells[3].innerText = actionValue;
                row.cells[4].innerText = recipientValue;

                console.log("Llamando a guardarEnSharePoint:", nombreArchivo, actionValue, recipientValue, numeroTransmittal);
                window.opener.guardarEnSharePoint(nombreArchivo, actionValue, recipientValue, numeroTransmittal);
            }
        });

        console.log("Llamando a actualizarContadores:", actionValue, recipientValue);
        window.opener.actualizarContadores(actionValue, recipientValue);
        alert("Archivos asignados correctamente");
    } else {
        alert("Debe seleccionar una acción, un destinatario y al menos un archivo");
    }
});
