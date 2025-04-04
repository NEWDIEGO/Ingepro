import * as React from "react";
import { useState, useEffect } from "react";
import { /*ISPHttpClientOptions*/ SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IHolaMundo2Props } from "./IHolaMundo2Props";
import { SPFI, spfi } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IHolaMundo2Props } from "./IHolaMundo2Props";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { agregarADistribucion } from "../services/spService";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
//import { SPFx } from "@pnp/sp/behaviors/spfx";
//import { sp } from "../services/spConfig"; // Asegúrate de la ruta correcta
// Se crea una instancia de SPFI (SharePoint Fluent Interface) utilizando `spfi`, 
// que permite interactuar con la API de SharePoint Online. 
// Se le pasa la URL base del sitio de SharePoint donde se realizarán las operaciones.

const sp: SPFI = spfi("https://ntechilespa.sharepoint.com/sites/ingp");

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

//url para localhost
// Se extiende la interfaz global de `window` para agregar dos funciones personalizadas 
// que estarán disponibles en el objeto global `window`.
// Estas funciones permitirán interactuar con SharePoint desde cualquier parte de la aplicación.



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

// Se define la interfaz `IHolaMundo2State`, que representa el estado del componente `HolaMundo2`.
// Esta interfaz se utiliza para tipar el estado en un componente de React.

interface IHolaMundo2State { 
  lista: IDocumentoTecnico[]; // Lista de documentos técnicos disponibles en la aplicación.
  actions: string[]; // Lista de acciones posibles que pueden asignarse a los documentos.
  recipients: string[]; // Lista de destinatarios a los que se pueden asignar los documentos.
  isAssignDisabled: boolean; // Botón "Asignar a Distribución" habilitado/inhabilitado
  showTable: boolean;  // Nuevo estado para alternar la vista
  actionCount: { [key: string]: number }; //accion contador
}

// Se define la clase `HolaMundo2`, que extiende de `React.Component`.
// Esta clase representa un componente de React con estado, que gestiona la asignación de documentos en SharePoint.


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
      actionCount: {}, //accion contador
      accionSeleccionada: null, // Agregar aquí la acción seleccionada
      contadorAccion: 0 // Agregar aquí el contador de la acción seleccionada
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
    // 📌 ✅ FUNCIONES PARA ACTUALIZAR EL ESTADO
    setAccionSeleccionada = (accion: string) => {
      this.setState({ accionSeleccionada: accion });
  };

  setContadorAccion = (contador: number) => {
      this.setState({ contadorAccion: contador });
  };
  // ✅ Ahora podemos usar el contexto de SharePoint de forma segura

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
  
  // Convierte el estado del componente a una cadena JSON con formato legible.
  // `null, 2` significa que no se aplicará transformación en los datos (`null`) 
  // y que se usará una indentación de 2 espacios para mejorar la legibilidad.
  
  // Abre una nueva ventana emergente con un tamaño de 600x400 píxeles.
  // `window.open` crea una nueva ventana o pestaña con los parámetros dados.

  const popupWindow = window.open("about:blank", "_blank", "width=800,height=600");

  // Verifica si la ventana emergente se abrió correctamente.


  const popupWindow = window.open("about:blank", "_blank", "width=800,height=600");

  if (!popupWindow) {
      alert("Por favor, permite las ventanas emergentes en tu navegador.");
      return; // Si no se pudo abrir, se muestra una alerta y se detiene la ejecución.
  }


  // Escribe el contenido HTML dentro de la ventana emergente.
  // Se utiliza `document.write()` para insertar un documento HTML completo en la nueva ventana.

  // Cierra el documento de la ventana emergente para finalizar la escritura del contenido.
  

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

      await this.props.context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
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

      await this.props.context.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
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

// Método privado `guardarEnSharePoint`: Guarda un nuevo registro en la lista "Distribucion" de SharePoint.
// Recibe como parámetros el nombre del archivo, la acción y el destinatario.
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
      Title: nombreArchivo, // Nombre del archivo
      Accion: accion, // Acción a realizar (verifica el nombre en SharePoint)
      Destinatario: destinatario, // Destinatario del archivo
      NumeroTransmittal: numeroTransmittal, // Número único de transmittal
      FechaDistribucion: fecha // Fecha de distribución (verifica si el campo existe)
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

// Método `handleAsignarDistribucion`: Maneja la asignación de archivos a una distribución en SharePoint.
// Verifica si se han seleccionado archivos, una acción y un destinatario antes de proceder.
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

// Método del ciclo de vida de React `componentDidMount`.
// Se ejecuta automáticamente después de que el componente se ha montado en el DOM.
// En este caso, se utiliza para obtener datos iniciales necesarios para el funcionamiento del componente.
public async componentDidMount(): Promise<void> {

  // Llama al método `getData` para obtener la lista de documentos técnicos desde SharePoint.

  await this.getData();

  // Llama al método `getActions` para obtener la lista de acciones disponibles.

  await this.getActions();

  // Llama al método `getRecipients` para obtener la lista de destinatarios.

  await this.getRecipients();
}

// Método privado `getData`: Obtiene la lista de documentos técnicos desde SharePoint
// y actualiza el estado del componente con la información obtenida.
private getData = async (): Promise<void> => {
  try {

    // Realiza una solicitud GET a la API de SharePoint para obtener los elementos de la lista "DocumentosTecnicos".
    // Se seleccionan los campos `Id` y `File/Name` y se expande la propiedad `File` para acceder al nombre del archivo.

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

    // Convierte la respuesta en un objeto JSON y define la estructura esperada.

    const data: { value: { Id: number; File: { Name: string } }[] } = await response.json();

    // Mapea los datos obtenidos y los transforma en un formato compatible con la interfaz `IDocumentoTecnico`.

    const data: { value: { Id: number; File: { Name: string } }[] } = await response.json();

    const lista = data.value.map((item) => ({
      id: item.Id, // Identificador del documento.
      name: item.File?.Name || "Sin nombre", // Nombre del archivo o "Sin nombre" si no está definido.
    }));

    // Actualiza el estado del componente con la lista de documentos obtenida.
    
    this.setState({ lista });
  } catch (error) {

    // Captura y muestra un mensaje de error en caso de que la solicitud falle.

    console.error("Error fetching data:", error);
  }
}

// Método privado `getActions`: Obtiene la lista de acciones disponibles desde la lista "Accion" en SharePoint
// y actualiza el estado del componente con las acciones obtenidas.
private getActions = async (): Promise<void> => {
  try {

    // Realiza una solicitud GET a la API de SharePoint para obtener los elementos de la lista "Accion".
    // Se selecciona solo el campo `Title`, que representa el nombre de cada acción.

private getActions = async (): Promise<void> => {
  try {

    const response = await this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Accion')/items?$select=Title`,
      SPHttpClient.configurations.v1
    );

    // Convierte la respuesta en un objeto JSON y define la estructura esperada.

    const data: { value: { Title: string }[] } = await response.json();

    // Mapea los datos obtenidos y extrae solo los títulos de las acciones.

    const actions = data.value.map((item) => item.Title);

    // Actualiza el estado del componente con la lista de acciones obtenida.

    this.setState({ actions });
        // Llamar a actualizarContadorAccion para cada acción
    const data: { value: { Title: string }[] } = await response.json();

    const actions = data.value.map((item) => item.Title);

    this.setState({ actions });

      actions.forEach(async (action) => {
        await this.actualizarContadorAccion(action);
      });
  } catch (error) {

    // Captura y muestra un mensaje de error en caso de que la solicitud falle.

    console.error("Error fetching actions:", error);
  }
}

// Método privado `getRecipients`: Obtiene la lista de destinatarios desde la lista "Destinatarios" en SharePoint
// y actualiza el estado del componente con los destinatarios obtenidos.
private getRecipients = async (): Promise<void> => {
  try {

    // Realiza una solicitud GET a la API de SharePoint para obtener los elementos de la lista "Destinatarios".
    // Se selecciona solo el campo `Title`, que representa el nombre de cada destinatario.

private getRecipients = async (): Promise<void> => {
  try {

    const response = await this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Destinatarios')/items?$select=Title`,
      SPHttpClient.configurations.v1
    );

    // Convierte la respuesta en un objeto JSON y define la estructura esperada.

    const data: { value: { Title: string }[] } = await response.json();

    // Mapea los datos obtenidos y extrae solo los títulos de los destinatarios.

    const recipients = data.value.map((item) => item.Title);

    // Actualiza el estado del componente con la lista de destinatarios obtenida.

    this.setState({ recipients });
  } catch (error) {

    // Captura y muestra un mensaje de error en caso de que la solicitud falle.

    const data: { value: { Title: string }[] } = await response.json();

    const recipients = data.value.map((item) => item.Title);

    this.setState({ recipients });
  } catch (error) {

    console.error("Error fetching recipients:", error);
  }
}

// Método `handleButtonClick`: Abre una ventana emergente con una tabla de documentos técnicos y opciones para asignación.
// Esta ventana permite al usuario seleccionar documentos, asignarles una acción y un destinatario, y luego guardar la información en SharePoint.
private handleButtonClick = async (): Promise<void> => {

  // Abre una nueva ventana emergente sin URL inicial.

  const popupWindow = window.open("", "_blank");

  // Mensaje en consola para registrar que el botón fue presionado.

  console.log("Botón 'Asignar a Transmittal' fue clickeado");

  // Verifica si la ventana emergente se abrió correctamente.

  if (popupWindow) {

      // Escribe el contenido HTML de la ventana emergente.

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

            let numeroTransmittal = document.querySelector("#idDelInput")?.value || "";
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
                          window.opener.guardarEnSharePoint(nombreArchivo, accion, destinatario, numeroTransmittal, fecha);

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
                        window.opener.guardarEnSharePoint(nombreArchivo, accion, destinatario, numeroTransmittal, fecha);
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

    // Se ejecuta después de 500ms para asegurarse de que los elementos estén cargados en la ventana emergente.

    setTimeout(() => {
      const assignButton = popupWindow.document.getElementById("assignDistributionButton") as HTMLButtonElement;
      
      if (assignButton) {
        assignButton.disabled = false; // Habilita el botón antes de agregar el evento
      
        assignButton.addEventListener("click", () => { 

          // Aquí se debe agregar la lógica para manejar los datos del transmittal.

          alert("Distribuido"); // Alerta cuando se presiona el botón
        });
      
      } else {
        console.error("Error: No se encontró el botón 'Asignar a Distribución'");
      }
    }, 500);

    // Evento adicional para manejar la asignación de distribución en la ventana emergente
    // Se agrega un evento `click` al botón con id `assignDistributionButton` en la ventana emergente.

// Se obtiene el botón con ID 'assignDistributionButton' dentro de la ventana emergente.
// `?.` asegura que si el elemento no existe, no se lance un error.
popupWindow.document.getElementById('assignDistributionButton')?.addEventListener('click', async () => {
  try {
      console.log("Botón 'Asignar a Distribución' presionado."); // Mensaje en la consola para confirmar el evento.

      // Se obtiene el valor seleccionado en el campo de acción (<select id="action">).
      const actionValue = (popupWindow.document.getElementById("action") as HTMLSelectElement)?.value;

      // Se obtiene el valor seleccionado en el campo de destinatario (<select id="recipient">).
      const recipientValue = (popupWindow.document.getElementById("recipient") as HTMLSelectElement)?.value;

      // Se muestra en consola qué acción y destinatario fueron seleccionados.
      console.log(`Acción seleccionada: ${actionValue}, Destinatario seleccionado: ${recipientValue}`);

      // Se valida que ambos valores (acción y destinatario) estén definidos antes de continuar.
      if (actionValue && recipientValue) {
          // Se llama a la función `actualizarContadores`, que posiblemente actualiza estadísticas
          // o algún estado de la aplicación sobre las asignaciones realizadas.
          await this.actualizarContadores(actionValue, recipientValue);
      } else {
          // Si no se seleccionaron valores, se muestra una alerta indicando que son obligatorios.
          alert("Debe seleccionar una acción y un destinatario.");
      }
  } catch (error) {
      // En caso de que ocurra un error en la ejecución de `actualizarContadores`,
      // se captura y se muestra en la consola.
      console.error("Error en actualizarContadores:", error);
  }
});
// Método `actualizarContadores`: Incrementa los contadores de acciones y destinatarios en SharePoint.
// Recibe los valores de acción (`actionValue`) y destinatario (`recipientValue`) como parámetros.
  }
}



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

      // 🔹 Buscar la acción en la lista "Accion"

      const accionItems = await sp.web.lists.getByTitle("Accion").items
          .filter(`Title eq '${actionValue}'`).top(1)();

      let accionContador = 0;
      let accionItemId = null;

      if (accionItems.length > 0) {
          accionItemId = accionItems[0].Id;
          accionContador = (accionItems[0].Acciones_x0020_asignadas || 0) + 1;

          // 🔹 Actualizar en SharePoint usando REST API (PATCH)
          await sp.web.lists.getByTitle("Accion").items.getById(accionItemId).update({
              Acciones_x0020_asignadas: accionContador
          await sp.web.lists.getByTitle("Accion").items.getById(accionItemId).update({
            __metadata: { type: "SP.Data.AccionListItem" },
            Acciones_x0020_asignadas: accionContador
          });

          console.log(`Acción actualizada en SharePoint: ${actionValue} N°${accionContador}`);
      } else {
          // Si la acción no existe, crearla (POST)
          const newItem = await sp.web.lists.getByTitle("Accion").items.add({
              Title: actionValue,
              Acciones_x0020_asignadas: 1


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

      // 🔹 Buscar destinatario en la lista "Destinatarios"
      console.log(`📌 Buscando destinatario "${recipientValue}" en SharePoint...`);

      const destinatarioItems = await sp.web.lists.getByTitle("Destinatarios").items
          .filter(`Title eq '${recipientValue}'`).top(1)();

      let destinatarioContador = 0;
      let destinatarioItemId = null;

      if (destinatarioItems.length > 0) {
          destinatarioItemId = destinatarioItems[0].Id;
          destinatarioContador = (destinatarioItems[0].Destinatarios_x0020_asignados || 0) + 1;

          // 🔹 Actualizar en SharePoint usando REST API (PATCH)
          await sp.web.lists.getByTitle("Destinatarios").items.getById(destinatarioItemId).update({
              Destinatarios_x0020_asignados: destinatarioContador
          await sp.web.lists.getByTitle("Destinatarios").items.getById(destinatarioItemId).update({
            __metadata: { type: "SP.Data.AccionListItem" },
            Destinatarios_x0020_asignados: destinatarioContador
          });

          console.log(`✅ Destinatario actualizado en SharePoint: ${recipientValue} N°${destinatarioContador}`);
      } else {
          // Si el destinatario no existe, crearlo (POST)
          const newItem = await sp.web.lists.getByTitle("Destinatarios").items.add({
              Title: recipientValue,
              Destinatarios_x0020_asignados: 1


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