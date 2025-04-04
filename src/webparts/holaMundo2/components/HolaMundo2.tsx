import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IHolaMundo2Props } from "./IHolaMundo2Props";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { agregarADistribucion } from "../services/spService";
import "@pnp/sp/items/list";
import "@pnp/sp/items/get-all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { Logger, LogLevel } from "@pnp/logging";

Logger.subscribe(console);
Logger.activeLogLevel = LogLevel.Info; // o .Verbose si necesitas más detalle

let _sp: any;

fetch('https://example.sharepoint.com/api/data')
    .then(response => response.json())
    .then(data => console.log(data))
    .catch(error => console.error('Error atrapado:', error));

export const getSP = (context?: WebPartContext) => {
  if (!_sp && context) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

const sp: SPFI = spfi("https://ntechilespa.sharepoint.com/sites/ingp");

export const guardarEnSharePoint = async (
  nombreArchivo: string,
  accion: string,
  destinatario: string,
  numeroTransmittal: string,
  fecha: string
): Promise<void> => {
  console.log("✅ Iniciando función guardarEnSharePoint");
  try {
    const payload = {
      Title: nombreArchivo,
      Accion: accion,
      Destinatario: destinatario,
      NumeroTransmittal: numeroTransmittal,
      fecha: fecha
    };
    console.log("📤 Datos a enviar:", payload);

    const res = await sp.web.lists.getByTitle("Distribucion").items.add(payload);

    console.log("✅ Elemento guardado en SharePoint:", res);
  } catch (error) {
    console.error("❌ Error al guardar en SharePoint:", error.message, error);
  }
};

interface IDocumentoTecnico {
  id: number;
  name: string;
}
async function verificarConexion(): Promise<void> {
  try {
    const listas = await sp.web.lists.select("Title", "Id").expand("Fields")();
    console.log("Listas en SharePoint:", listas);
  } catch (error) {
    console.error("Error al conectar con SharePoint:", error);
  }
}

verificarConexion()
  .catch(error => console.error("Error en la conexión:", error))
  declare global {
    interface Window {
      // Función para guardar un archivo en SharePoint, incluyendo número de transmittal y fecha.
      guardarEnSharePoint: (
        nombreArchivo: string,
        accion: string,
        destinatario: string,
        numeroTransmittal: number, // Nuevo parámetro agregado
        fecha: string // Nuevo parámetro agregado
      ) => Promise<void>;
  
      // Función para actualizar los contadores en la interfaz.
      actualizarContadores: (actionValue: string, recipientValue: string) => Promise<void>;
    }
  }

interface IHolaMundo2State { 
  lista: IDocumentoTecnico[]; // Lista de documentos técnicos disponibles en la aplicación.
  actions: string[]; // Lista de acciones posibles que pueden asignarse a los documentos.
  recipients: string[]; // Lista de destinatarios a los que se pueden asignar los documentos.
  isAssignDisabled: boolean; // Botón "Asignar a Distribución" habilitado/inhabilitado
  showTable: boolean;  // Nuevo estado para alternar la vista
  actionCount: { [key: string]: number }; //accion contador
}
export default class HolaMundo2 extends React.Component<IHolaMundo2Props, IHolaMundo2State> { 

  // Constructor del componente, recibe `props` como argumento.
  private : SPFI; // Declaramos `sp` en la clase
  
  constructor(props: IHolaMundo2Props) {
    super(props);
    this.state = {
      lista: [], // Lista de documentos técnicos.
      actions: [], //accion
      recipients: [], //destinatario
      isAssignDisabled: true, // Comienza deshabilitado
      showTable: false, // Indica si se debe mostrar la tabla.
      actionCount: {} //accion contador
      //accionSeleccionada: null, // Agregar aquí la acción seleccionada
      //contadorAccion: 0 // Agregar aquí el contador de la acción seleccionada
    }
    // Se enlazan los métodos a la instancia de la clase para garantizar el acceso al contexto `this`.

    // Esto es necesario en clases de React que extienden `Component`, ya que los métodos 
    // no están automáticamente ligados al contexto de la instancia de la clase.

    this.mostrarVentana = this.mostrarVentana.bind(this); // Enlaza el método `mostrarVentana` al contexto de la clase.
    this.handleButtonClick = this.handleButtonClick.bind(this); // Enlaza el método `handleButtonClick` al contexto de la clase.
    this.guardarEnSharePoint = this.guardarEnSharePoint.bind(this); // Enlaza el método `guardarEnSharePoint` al contexto de la clase.
    this.actualizarContadores = this.actualizarContadores.bind(this); // Enlaza el método `actualizarContadores` al contexto de la clase.

    // Se agregan las funciones `guardarEnSharePoint` y `actualizarContadores` al objeto global `window`,
    // permitiendo que sean accesibles desde cualquier parte de la aplicación.

    window.guardarEnSharePoint = this.guardarEnSharePoint.bind(this); // Se asigna la función `guardarEnSharePoint` al objeto global `window`.
    
    window.actualizarContadores = this.actualizarContadores.bind(this); // Se asigna la función `actualizarContadores` al objeto global `window`.
  }

  //Ahora podemos usar el contexto de SharePoint de forma segura

  async verificarConexion(): Promise<void> {
    try {
      const listas = await sp.web.lists.select("Title", "Id").expand("Fields")();
      console.log("Listas en SharePoint:", listas);
    } catch (error) {
      console.error("Error al conectar con SharePoint:", error);
    }
  }

public render(): JSX.Element {
  return (
    <div>
      <h2>Hola Mundo 2</h2>
      <button onClick={this.handleButtonClick}>Asignar a Transmittal</button>
    </div>
  );
}

// Función `mostrarVentana`: Abre una nueva ventana emergente y muestra el estado actual del componente en formato JSON.
mostrarVentana = (): void => {

  const popupWindow = window.open("about:blank", "_blank", "width=800,height=600");

  if (!popupWindow) {
      alert("Por favor, permite las ventanas emergentes en tu navegador.");
      return; // Si no se pudo abrir, se muestra una alerta y se detiene la ejecución.
  }

  popupWindow.document.close();
}

private async actualizarContadorAccion(actionTitle: string): Promise<void> {
  try {
    // Paso 1: Verificar si la acción ya existe en la lista
    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Accion')/items?$filter=Title eq '${encodeURIComponent(actionTitle)}'&$select=Id,ContadorAccion`;

    const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value.length > 0) {
      // La acción ya existe, actualizar el contador
      const accionId = data.value[0].Id;
      const nuevoContador = (data.value[0].ContadorAccion || 0) + 1;

      const updateUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Accion')/items(${accionId})`;

      const updateBody = JSON.stringify({
        ContadorAccion: nuevoContador
      });

      await this.props.context.spHttpClient.get(updateUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          //'X-RequestDigest': (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: updateBody
      });

      console.log(`✅ Acción "${actionTitle}" actualizada con contador: ${nuevoContador}`);
    } else {
      // La acción no existe, crear un nuevo registro
      const createUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Accion')/items`;

      const createBody = JSON.stringify({
        Title: actionTitle,
        ContadorAccion: 1
      });

      await this.props.context.spHttpClient.get(createUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          //'X-RequestDigest': (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value
        },
        body: createBody
      });

      console.log(`✅ Nueva acción "${actionTitle}" creada con contador: 1`);
    }
  } catch (error) {
    console.error("❌ Error al actualizar contador de acciones:", error);
  }
}

private async guardarEnSharePoint(
  nombreArchivo: string,
  accion: string,
  destinatario: string,
  numeroTransmittal: number,
  fecha: string
): Promise<void> {
  try {
    console.log(`📌 Guardando en SharePoint: ${nombreArchivo} / ${accion} / ${destinatario} / ${numeroTransmittal} / ${fecha}`);

    // Verifica que los valores no estén vacíos
    if (!nombreArchivo || !accion || !destinatario || !numeroTransmittal || !fecha) {
      throw new Error("❌ Faltan datos obligatorios para guardar en SharePoint.");
    }

    // Enviar los datos a la lista "Distribucion"
    const response = await sp.web.lists.getByTitle("Distribucion").items.add({
      __metadata: { type: "SP.Data.DistribucionListItem" },
      Title: nombreArchivo, // Nombre del archivo
      accion: accion, // Acción a realizar (verifica el nombre en SharePoint)
      destinatario: destinatario, // Destinatario del archivo
      NumeroTransmittal: numeroTransmittal, // Número único de transmittal
      fecha: fecha // Fecha de distribución (verifica si el campo existe)
    });

    if (response?.data) {
      console.log(`✅ Archivo guardado en SharePoint correctamente:`, response.data);
      alert(`Archivo guardado correctamente:\nArchivo: ${nombreArchivo}\nAcción: ${accion}\nDestinatario: ${destinatario}\nNúmero Transmittal: ${numeroTransmittal}\nFecha: ${fecha}`);
    } else {
      throw new Error("No se recibió una respuesta válida de SharePoint.");
    }
  } catch (error) {
    console.error("Error al guardar en SharePoint:", error);
    alert("Error al guardar en SharePoint. Revisa la consola para más detalles.");
  }
}

handleAsignarDistribucion = async (): Promise<void> => {

  // Obtiene el valor seleccionado en el campo "action" (acción a realizar).
  const actionValue = (document.getElementById("action") as HTMLSelectElement)?.value;

  // Obtiene el valor seleccionado en el campo "recipient" (destinatario del archivo).
  const recipientValue = (document.getElementById("recipient") as HTMLSelectElement)?.value;

  // Selecciona todos los checkboxes marcados con la clase "checkbox" (archivos seleccionados).
  const selectedFiles = Array.from(document.querySelectorAll('.checkbox:checked'))
      .map(cb => cb.closest("tr")?.cells[1].innerText); // Obtiene el nombre del archivo de la segunda celda (índice 1) de la fila.

  // Verifica si no se han seleccionado archivos, una acción o un destinatario, y muestra una alerta.
  if (selectedFiles.length === 0 || !actionValue || !recipientValue) {
      alert("Debe seleccionar al menos un archivo, una acción y un destinatario.");
      return;
  }

  // Se actualiza el estado para contar cuántas veces se ha asignado cada acción.
  this.setState((prevState) => ({
    actionCount: {
      ...prevState.actionCount,
      [actionValue]: (prevState.actionCount[actionValue] || 0) + 1 // Incrementa el contador de la acción seleccionada.
    }
  }), 
  () => {
    // Callback que se ejecuta después de que `setState` haya actualizado el estado.
    console.log(`Acción '${actionValue}' asignada ${this.state.actionCount[actionValue]} veces.`);
    console.log("Nuevo estado de actionCount:", this.state.actionCount);
  });
  try {
    for (const file of selectedFiles) {
      const resultado = await agregarADistribucion("prueba.docx", "Acción", "Destinatario", "123456");
      console.log(`✅ Archivo '${file}' asignado correctamente:`, resultado);
    }
  // Guardar en el estado
  
  // Función incorrectamente anidada dentro de `handleAsignarDistribucion`, lo que puede generar errores.
  // Probablemente debería definirse fuera de esta función.
  this.handleAsignarDistribucion = async () => {
    const resultado = await agregarADistribucion("prueba.docx", "Acción", "Destinatario", "123456");
    console.log("Resultado de la asignación:", resultado);
    }
  } catch (error) {
    console.error("Error al asignar la distribución:", error);
  }

  // Muestra en la consola el estado actualizado después de la asignación.
  console.log("Datos guardados en el estado:", this.state);
};

public async componentDidMount(): Promise<void> {

  await this.getData();

  await this.getActions();

  await this.getRecipients();
}

private getData = async (): Promise<void> => {
  try {

    const response = await this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('DocumentosTecnicos')/items?$select=Id,File/Name&$expand=File`,
      SPHttpClient.configurations.v1
    );

    const data: { value: { Id: number; File: { Name: string } }[] } = await response.json();

    const lista = data.value.map((item) => ({
      id: item.Id, // Identificador del documento.
      name: item.File?.Name || "Sin nombre", // Nombre del archivo o "Sin nombre" si no está definido.
    }));
    
    this.setState({ lista });
  } catch (error) {

    console.error("Error fetching data:", error);
  }
}

private getActions = async (): Promise<void> => {
  try {

    const response = await this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Accion')/items?$select=Title`,
      SPHttpClient.configurations.v1
    );

    const data: { value: { Title: string }[] } = await response.json();

    const actions = data.value.map((item) => item.Title);

    this.setState({ actions });

      actions.forEach(async (action) => {
        await this.actualizarContadorAccion(action);
      });
  } catch (error) {

    console.error("Error fetching actions:", error);
  }
}

private getRecipients = async (): Promise<void> => {
  try {

    const response = await this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Destinatarios')/items?$select=Title`,
      SPHttpClient.configurations.v1
    );

    const data: { value: { Title: string }[] } = await response.json();

    const recipients = data.value.map((item) => item.Title);

    this.setState({ recipients });
  } catch (error) {

    console.error("Error fetching recipients:", error);
  }
}

private handleButtonClick = async (): Promise<void> => {

  const popupWindow = window.open("", "_blank");

  console.log("Botón 'Asignar a Transmittal' fue clickeado");

  const contadorAcciones: Record<string, number> = {};
  const checkboxes = document.querySelectorAll('.checkbox:checked');
  checkboxes.forEach((checkbox) => {
    const row = checkbox.closest("tr");
    const accion = row?.cells[3].innerText;
    if (accion) {
      contadorAcciones[accion] = (contadorAcciones[accion] || 0) + 1;
    }
  });

  for (const accion in contadorAcciones) {
    if (!Object.prototype.hasOwnProperty.call(contadorAcciones, accion)) {
      continue;
    }
  
    const cantidad = contadorAcciones[accion];
    const items = await sp.web.lists
      .getByTitle("Accion")
      .items.select("Id", "Title", "Acciones asignadas")
      .filter(`Title eq '${accion}'`)();
  
    if (items.length > 0) {
      const item = items[0];
      const id = item.Id;
      const actual = item["Acciones asignadas"] || 0;
  
      await sp.web.lists
        .getByTitle("Accion")
        .items.getById(id)
        .update({
          "Acciones asignadas": actual + cantidad,
        });
    }
  }

  if (popupWindow) {

      popupWindow.document.write(`
      <html>
        <head>
          <title>Documentos Técnicos</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 10px; text-align: center; }
            button:disabled { background: #ccc; cursor: not-allowed; }
          </style>
        </head>
        <body>
          <h2 style="text-align: center;">Documentos Técnicos</h2>
          <div style="text-align: center;">
            <button id="clearButton">Limpiar</button>
            <select id="action">
              <option value="">Seleccione una acción</option>
              ${(this.state.actions || []).map(action => `<option value="${action}">${action}</option>`).join("")}
            </select>

            <select id="recipient">
              <option value="">Seleccione un destinatario</option>
              ${(this.state.recipients || []).map(recipient => `<option value="${recipient}">${recipient}</option>`).join("")}
            </select>
            <button id="assignDistributionButton">
              Asignar a Distribución
            </button>
            <input type="text" id="searchBar" placeholder="Buscar archivo..." />
            <button id="searchButton">Buscar</button>
            <button id="clearSearch">Limpiar búsqueda</button>
            <button id="saveButton">Guardar</button>
            <button onClick={actualizarContadorAccion} disabled={!accionSeleccionada}>
              Guardar Acción
            </button>
            <button>Asignar a Transmittal</button>
          </div>
          <table>
            <thead>
              <tr>
                <th><input type="checkbox" id="masterCheckbox"/></th>
                <th>Nombre del Archivo</th>
                <th>¿Distribuido?</th>
                <th>Acción</th>
                <th>Destinatario</th>
              </tr>
            </thead>

            <tbody id="tableBody">
              ${(this.state.lista || []).map(item => `
                <tr>
                  <td><input type="checkbox" class="checkbox" data-id="${item.id}" /></td>
                  <td>${item.name}</td>
                  <td>No</td>
                  <td>Sin acción</td>
                  <td>Sin destinatario</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
          <script>
            
            // Función para validar el formulario y habilitar/deshabilitar el botón de asignación

            function validateForm() {
            let checkboxesChecked = document.querySelectorAll('.checkbox:checked').length > 0;
            let actionSelected = document.getElementById('action').value !== "";
            let recipientSelected = document.getElementById('recipient').value !== "";

            console.log("Checkboxes seleccionados:", checkboxesChecked);
            console.log("Acción seleccionada:", actionSelected);
            console.log("Destinatario seleccionado:", recipientSelected);

            let numeroTransmittal = '12345';
            if (!numeroTransmittal) {
                console.error("Error: numeroTransmittal tiene un valor vacío.");
            }

            let assignButton = document.getElementById('assignDistributionButton');
            assignButton.disabled = !(checkboxesChecked && actionSelected && recipientSelected);
            }

            // Asignar eventos de cambio a selects y checkboxes

            document.getElementById('action').addEventListener('change', validateForm);
            document.getElementById('recipient').addEventListener('change', validateForm);
            document.querySelectorAll('.checkbox').forEach(checkbox => {
              checkbox.addEventListener('change', validateForm);
            });

            // boton para asignar distribución

            document.getElementById('assignDistributionButton').addEventListener("click", async () => {
              console.log("Botón 'Asignar a Distribución' presionado.");
              
              const actionValue = document.getElementById("action").value;
              const recipientValue = document.getElementById("recipient").value;
            
              const checkboxes = document.querySelectorAll('.checkbox:checked');

              if (actionValue && recipientValue && checkboxes.length > 0) {
                  checkboxes.forEach((checkbox) => {
                      const row = checkbox.closest("tr");
                      if (row) {
                          const nombreArchivo = row.cells[1].innerText;
                          const numeroTransmittal = Math.floor(100000 + Math.random() * 900000);
                          const fecha = new Date().toISOString().slice(0, 19).replace("T", " ");

                          row.cells[2].innerText = "Sí";
                          row.cells[3].innerText = actionValue;
                          row.cells[4].innerText = recipientValue;

                          console.log("Llamando a guardarEnSharePoint:", nombreArchivo, actionValue, recipientValue, numeroTransmittal, fecha);

                      }
                  });
                  console.log("Llamando a actualizarContadores:", actionValue, recipientValue);
                  window.opener.actualizarContadores(actionValue, recipientValue);
                  alert("Archivos asignados correctamente");
              } else {
                alert("Debe seleccionar una acción, un destinatario y al menos un archivo");
              }
            });

            // Evento para seleccionar o deseleccionar todos los checkboxes

            document.getElementById("masterCheckbox").addEventListener("change", function() {
                let checkboxes = document.querySelectorAll(".checkbox");
                checkboxes.forEach(checkbox => {
                    checkbox.checked = this.checked;
                });
                validateForm();
            });

            // Botón limpiar
            
            document.getElementById('clearButton').addEventListener('click', function() {
                // Obtiene todos los checkboxes con la clase 'checkbox'
                let checkboxes = document.querySelectorAll('.checkbox');

                // Recorre cada checkbox y lo desmarca
                checkboxes.forEach(checkbox => {
                    checkbox.checked = false;
                });

                // Restablece los selects a su valor predeterminado
                const actionSelect = document.getElementById("action");
                const recipientSelect = document.getElementById("recipient");

                if (actionSelect) actionSelect.selectedIndex = 0;
                if (recipientSelect) recipientSelect.selectedIndex = 0;

                // Ejecuta la validación nuevamente (si la función está definida)
                if (typeof validateForm === "function") {
                    validateForm();
                } else {
                    console.warn("La función validateForm no está definida.");
                }
            });

            // BUSCAR ELEMENTO/S

            document.getElementById("searchButton").addEventListener("click", function() {
                let searchValue = document.getElementById("searchBar").value.toLowerCase(); // Obtener texto de búsqueda
                let rows = document.querySelectorAll("tbody tr"); // Obtener todas las filas de la tabla
                let foundCount = 0; // Contador de resultados encontrados
                
                rows.forEach(row => {
                    let fileNameCell = row.cells[1]; // Segunda columna (nombre del archivo)
                    let checkbox = row.querySelector(".checkbox"); // Buscar checkbox en la fila

                    if (fileNameCell.innerText.toLowerCase().includes(searchValue)) {
                        row.style.backgroundColor = "#ffff99"; // Resaltar la fila en amarillo
                        if (checkbox) checkbox.checked = true; // Marcar checkbox
                        foundCount++; // Incrementar el contador de resultados encontrados
                    } else {
                        row.style.backgroundColor = ""; // Restaurar color original
                        if (checkbox) checkbox.checked = false; // Desmarcar checkbox
                    }
                });

                alert("Resultados: " + foundCount);

                validateForm(); // Actualizar validación del formulario
            });

            // Botón para limpiar la búsqueda

            document.getElementById("clearSearch").addEventListener("click", function() {
                document.getElementById("searchBar").value = ""; // Vaciar el campo de búsqueda
                let rows = document.querySelectorAll("tbody tr");

                rows.forEach(row => {
                    row.style.backgroundColor = ""; // Restaurar colores
                    let checkbox = row.querySelector(".checkbox");
                    if (checkbox) checkbox.checked = false; // Desmarcar todos los checkboxes
                });

                validateForm(); // Actualizar validación del formulario
            });          

            // boton "guardar"

            document.getElementById("saveButton").addEventListener("click", async () => {
                console.log("Botón 'Guardar' presionado.");

                // Obtiene todos los checkboxes seleccionados
                const checkboxes = document.querySelectorAll(".checkbox:checked");

                if (checkboxes.length === 0) {
                    alert("Debe seleccionar al menos un archivo para guardar.");
                    return;
                }

                for (let checkbox of checkboxes) {
                    const row = checkbox.closest("tr");
                    if (row) {
                        const nombreArchivo = row.cells[1].innerText;
                        const accion = row.cells[3].innerText;
                        const destinatario = row.cells[4].innerText;

                        // Genera un número transmittal aleatorio
                        const numeroTransmittal = Math.floor(100000 + Math.random() * 900000);

                        // Obtiene la fecha actual en formato YYYY-MM-DD HH:MM:SS
                        const fecha = new Date().toISOString().slice(0, 19).replace("T", " ");

                        // Depuración: Verificar valores antes de enviar
                        console.log("Datos antes de guardar en SharePoint:", {
                            nombreArchivo, accion, destinatario, numeroTransmittal, fecha
                        });

                        // Llamamos a la API REST de SharePoint para guardar los datos
                        console.log("Llamando a guardarEnSharePoint:", nombreArchivo, accion, destinatario, numeroTransmittal, fecha);

                    }
                }

                alert("Datos guardados correctamente en SharePoint.");
            });


            // Ejecutar la validación al abrir la ventana

            validateForm(); 
          </script>
        </body>
      </html>
    `);            

  popupWindow.document.close();

    setTimeout(() => {
      const assignButton = popupWindow.document.getElementById("assignDistributionButton") as HTMLButtonElement;
      
      if (assignButton) {
        assignButton.disabled = false; // Habilita el botón antes de agregar el evento
      
        assignButton.addEventListener("click", () => { 

          alert("Distribuido"); // Alerta cuando se presiona el botón
        });
      
      } else {
        console.error("Error: No se encontró el botón 'Asignar a Distribución'");
      }
    }, 500);

popupWindow.document.getElementById('assignDistributionButton')?.addEventListener('click', async () => {
  try {
      console.log("Botón 'Asignar a Distribución' presionado."); // Mensaje en la consola para confirmar el evento.

      const actionValue = (popupWindow.document.getElementById("action") as HTMLSelectElement)?.value;

      const recipientValue = (popupWindow.document.getElementById("recipient") as HTMLSelectElement)?.value;

      console.log(`Acción seleccionada: ${actionValue}, Destinatario seleccionado: ${recipientValue}`);

      if (actionValue && recipientValue) {
          await this.actualizarContadores(actionValue, recipientValue);
      } else {
          alert("Debe seleccionar una acción y un destinatario.");
      }
  } catch (error) {
      console.error("Error en actualizarContadores:", error);
  }
  });
  }
}

private async actualizarContadores(actionValue: string, recipientValue: string): Promise<void> {
  if (!actionValue || !recipientValue) {
      alert("Debe seleccionar una acción y un destinatario.");
      return;
  }

  try {
      console.log(`📌 Buscando acción en SharePoint con filtro: Title eq '${actionValue}'`);
      const accionItems = await sp.web.lists.getByTitle("Accion").items
          .filter(`Title eq '${actionValue}'`).top(1)();

      let accionContador = 0;
      let accionItemId = null;

      if (accionItems.length > 0) {
          accionItemId = accionItems[0].Id;
          accionContador = (accionItems[0].Acciones_x0020_asignadas || 0) + 1;

          await sp.web.lists.getByTitle("Accion").items.getById(accionItemId).update({
            __metadata: { type: "SP.Data.AccionListItem" },
            Acciones_x0020_asignadas: accionContador
          });

          console.log(`Acción actualizada en SharePoint: ${actionValue} N°${accionContador}`);
      } else {

          const newItem = await sp.web.lists.getByTitle("Accion").items.add({
            __metadata: { type: "SP.Data.AccionListItem" },
            Title: actionValue,
            Acciones_x0020_asignadas: 1
          });

          accionItemId = newItem.data.Id;
          accionContador = 1;

          console.log(`Nueva acción creada en SharePoint: ${actionValue} N°${accionContador}`);
      }
      console.log(`📌 Buscando destinatario "${recipientValue}" en SharePoint...`);

      const destinatarioItems = await sp.web.lists.getByTitle("Destinatarios").items
          .filter(`Title eq '${recipientValue}'`).top(1)();

      let destinatarioContador = 0;
      let destinatarioItemId = null;

      if (destinatarioItems.length > 0) {
          destinatarioItemId = destinatarioItems[0].Id;
          destinatarioContador = (destinatarioItems[0].Destinatarios_x0020_asignados || 0) + 1;

          await sp.web.lists.getByTitle("Destinatarios").items.getById(destinatarioItemId).update({
            __metadata: { type: "SP.Data.AccionListItem" },
            Destinatarios_x0020_asignados: destinatarioContador
          });

          console.log(`✅ Destinatario actualizado en SharePoint: ${recipientValue} N°${destinatarioContador}`);
      } else {

          const newItem = await sp.web.lists.getByTitle("Destinatarios").items.add({
            __metadata: { type: "SP.Data.AccionListItem" },
            Title: recipientValue,
            Destinatarios_x0020_asignados: 1
          });

          destinatarioItemId = newItem.data.Id;
          destinatarioContador = 1;

          console.log(`✅ Nuevo destinatario creado en SharePoint: ${recipientValue} N°${destinatarioContador}`);
      }

      alert("Datos almacenados correctamente en SharePoint.");
  } catch (error) {
      console.error("❌ Error al actualizar datos en SharePoint:", error);
  }
}

}