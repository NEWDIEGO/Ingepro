var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from "react";
import { /*ISPHttpClientOptions*/ SPHttpClient } from "@microsoft/sp-http";
import { spfi } from "@pnp/sp";
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
var sp = spfi("https://ntechilespa.sharepoint.com/sites/ingp");
//url para localhost
// Se extiende la interfaz global de `window` para agregar dos funciones personalizadas 
// que estarán disponibles en el objeto global `window`.
// Estas funciones permitirán interactuar con SharePoint desde cualquier parte de la aplicación.
function verificarConexion() {
    return __awaiter(this, void 0, void 0, function () {
        var listas, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, sp.web.lists.select("Title", "Id").expand("Fields")()];
                case 1:
                    listas = _a.sent();
                    console.log("Listas en SharePoint:", listas);
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    console.error("Error al conectar con SharePoint:", error_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
verificarConexion()
    .catch(function (error) { return console.error("Error en la conexión:", error); });
// Se define la clase `HolaMundo2`, que extiende de `React.Component`.
// Esta clase representa un componente de React con estado, que gestiona la asignación de documentos en SharePoint.
var HolaMundo2 = /** @class */ (function (_super) {
    __extends(HolaMundo2, _super);
    function HolaMundo2(props) {
        var _this = _super.call(this, props) || this;
        // 📌 ✅ FUNCIONES PARA ACTUALIZAR EL ESTADO
        _this.setAccionSeleccionada = function (accion) {
            _this.setState({ accionSeleccionada: accion });
        };
        _this.setContadorAccion = function (contador) {
            _this.setState({ contadorAccion: contador });
        };
        // Función `mostrarVentana`: Abre una nueva ventana emergente y muestra el estado actual del componente en formato JSON.
        _this.mostrarVentana = function () {
            // Convierte el estado del componente a una cadena JSON con formato legible.
            // `null, 2` significa que no se aplicará transformación en los datos (`null`) 
            // y que se usará una indentación de 2 espacios para mejorar la legibilidad.
            // Abre una nueva ventana emergente con un tamaño de 600x400 píxeles.
            // `window.open` crea una nueva ventana o pestaña con los parámetros dados.
            var popupWindow = window.open("about:blank", "_blank", "width=800,height=600");
            // Verifica si la ventana emergente se abrió correctamente.
            if (!popupWindow) {
                alert("Por favor, permite las ventanas emergentes en tu navegador.");
                return; // Si no se pudo abrir, se muestra una alerta y se detiene la ejecución.
            }
            // Escribe el contenido HTML dentro de la ventana emergente.
            // Se utiliza `document.write()` para insertar un documento HTML completo en la nueva ventana.
            // Cierra el documento de la ventana emergente para finalizar la escritura del contenido.
            popupWindow.document.close();
        };
        // Método `handleAsignarDistribucion`: Maneja la asignación de archivos a una distribución en SharePoint.
        // Verifica si se han seleccionado archivos, una acción y un destinatario antes de proceder.
        _this.handleAsignarDistribucion = function () { return __awaiter(_this, void 0, void 0, function () {
            var actionValue, recipientValue, selectedFiles, _i, selectedFiles_1, file, resultado, error_2;
            var _this = this;
            var _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        actionValue = (_a = document.getElementById("action")) === null || _a === void 0 ? void 0 : _a.value;
                        recipientValue = (_b = document.getElementById("recipient")) === null || _b === void 0 ? void 0 : _b.value;
                        selectedFiles = Array.from(document.querySelectorAll('.checkbox:checked'))
                            .map(function (cb) { var _a; return (_a = cb.closest("tr")) === null || _a === void 0 ? void 0 : _a.cells[1].innerText; });
                        // Verifica si no se han seleccionado archivos, una acción o un destinatario, y muestra una alerta.
                        if (selectedFiles.length === 0 || !actionValue || !recipientValue) {
                            alert("Debe seleccionar al menos un archivo, una acción y un destinatario.");
                            return [2 /*return*/];
                        }
                        // Se actualiza el estado para contar cuántas veces se ha asignado cada acción.
                        this.setState(function (prevState) {
                            var _a;
                            return ({
                                actionCount: __assign(__assign({}, prevState.actionCount), (_a = {}, _a[actionValue] = (prevState.actionCount[actionValue] || 0) + 1 // Incrementa el contador de la acción seleccionada.
                                , _a))
                            });
                        }, function () {
                            // Callback que se ejecuta después de que `setState` haya actualizado el estado.
                            console.log("Acci\u00F3n '".concat(actionValue, "' asignada ").concat(_this.state.actionCount[actionValue], " veces."));
                            console.log("Nuevo estado de actionCount:", _this.state.actionCount);
                        });
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 6, , 7]);
                        _i = 0, selectedFiles_1 = selectedFiles;
                        _c.label = 2;
                    case 2:
                        if (!(_i < selectedFiles_1.length)) return [3 /*break*/, 5];
                        file = selectedFiles_1[_i];
                        return [4 /*yield*/, agregarADistribucion("prueba.docx", "Acción", "Destinatario", "123456")];
                    case 3:
                        resultado = _c.sent();
                        console.log("\u2705 Archivo '".concat(file, "' asignado correctamente:"), resultado);
                        _c.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 2];
                    case 5:
                        // Guardar en el estado
                        // Función incorrectamente anidada dentro de `handleAsignarDistribucion`, lo que puede generar errores.
                        // Probablemente debería definirse fuera de esta función.
                        this.handleAsignarDistribucion = function () { return __awaiter(_this, void 0, void 0, function () {
                            var resultado;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, agregarADistribucion("prueba.docx", "Acción", "Destinatario", "123456")];
                                    case 1:
                                        resultado = _a.sent();
                                        console.log("Resultado de la asignación:", resultado);
                                        return [2 /*return*/];
                                }
                            });
                        }); };
                        return [3 /*break*/, 7];
                    case 6:
                        error_2 = _c.sent();
                        console.error("Error al asignar la distribución:", error_2);
                        return [3 /*break*/, 7];
                    case 7:
                        // Muestra en la consola el estado actualizado después de la asignación.
                        console.log("Datos guardados en el estado:", this.state);
                        return [2 /*return*/];
                }
            });
        }); };
        // Método privado `getData`: Obtiene la lista de documentos técnicos desde SharePoint
        // y actualiza el estado del componente con la información obtenida.
        _this.getData = function () { return __awaiter(_this, void 0, void 0, function () {
            var response, data, lista, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.props.context.spHttpClient.get("".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('DocumentosTecnicos')/items?$select=Id,File/Name&$expand=File"), SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        lista = data.value.map(function (item) {
                            var _a;
                            return ({
                                id: item.Id,
                                name: ((_a = item.File) === null || _a === void 0 ? void 0 : _a.Name) || "Sin nombre", // Nombre del archivo o "Sin nombre" si no está definido.
                            });
                        });
                        // Actualiza el estado del componente con la lista de documentos obtenida.
                        this.setState({ lista: lista });
                        return [3 /*break*/, 4];
                    case 3:
                        error_3 = _a.sent();
                        // Captura y muestra un mensaje de error en caso de que la solicitud falle.
                        console.error("Error fetching data:", error_3);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Método privado `getActions`: Obtiene la lista de acciones disponibles desde la lista "Accion" en SharePoint
        // y actualiza el estado del componente con las acciones obtenidas.
        _this.getActions = function () { return __awaiter(_this, void 0, void 0, function () {
            var response, data, actions, error_4;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.props.context.spHttpClient.get("".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('Accion')/items?$select=Title"), SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        actions = data.value.map(function (item) { return item.Title; });
                        // Actualiza el estado del componente con la lista de acciones obtenida.
                        this.setState({ actions: actions });
                        // Llamar a actualizarContadorAccion para cada acción
                        actions.forEach(function (action) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, this.actualizarContadorAccion(action)];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); });
                        return [3 /*break*/, 4];
                    case 3:
                        error_4 = _a.sent();
                        // Captura y muestra un mensaje de error en caso de que la solicitud falle.
                        console.error("Error fetching actions:", error_4);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Método privado `getRecipients`: Obtiene la lista de destinatarios desde la lista "Destinatarios" en SharePoint
        // y actualiza el estado del componente con los destinatarios obtenidos.
        _this.getRecipients = function () { return __awaiter(_this, void 0, void 0, function () {
            var response, data, recipients, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.props.context.spHttpClient.get("".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('Destinatarios')/items?$select=Title"), SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        recipients = data.value.map(function (item) { return item.Title; });
                        // Actualiza el estado del componente con la lista de destinatarios obtenida.
                        this.setState({ recipients: recipients });
                        return [3 /*break*/, 4];
                    case 3:
                        error_5 = _a.sent();
                        // Captura y muestra un mensaje de error en caso de que la solicitud falle.
                        console.error("Error fetching recipients:", error_5);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        // Método `handleButtonClick`: Abre una ventana emergente con una tabla de documentos técnicos y opciones para asignación.
        // Esta ventana permite al usuario seleccionar documentos, asignarles una acción y un destinatario, y luego guardar la información en SharePoint.
        _this.handleButtonClick = function () { return __awaiter(_this, void 0, void 0, function () {
            var popupWindow;
            var _this = this;
            var _a;
            return __generator(this, function (_b) {
                popupWindow = window.open("", "_blank");
                // Mensaje en consola para registrar que el botón fue presionado.
                console.log("Botón 'Asignar a Transmittal' fue clickeado");
                // Verifica si la ventana emergente se abrió correctamente.
                if (popupWindow) {
                    // Escribe el contenido HTML de la ventana emergente.
                    popupWindow.document.write("\n      <html>\n        <head>\n          <title>Documentos T\u00E9cnicos</title>\n          <style>\n            body { font-family: Arial, sans-serif; margin: 20px; }\n            table { width: 100%; border-collapse: collapse; margin-top: 20px; }\n            th, td { border: 1px solid black; padding: 10px; text-align: center; }\n            button:disabled { background: #ccc; cursor: not-allowed; }\n          </style>\n        </head>\n        <body>\n          <h2 style=\"text-align: center;\">Documentos T\u00E9cnicos</h2>\n          <div style=\"text-align: center;\">\n            <button id=\"clearButton\">Limpiar</button>\n            <select id=\"action\">\n              <option value=\"\">Seleccione una acci\u00F3n</option>\n              ".concat((this.state.actions || []).map(function (action) { return "<option value=\"".concat(action, "\">").concat(action, "</option>"); }).join(""), "\n            </select>\n\n            <select id=\"recipient\">\n              <option value=\"\">Seleccione un destinatario</option>\n              ").concat((this.state.recipients || []).map(function (recipient) { return "<option value=\"".concat(recipient, "\">").concat(recipient, "</option>"); }).join(""), "\n            </select>\n            <button id=\"assignDistributionButton\">\n              Asignar a Distribuci\u00F3n\n            </button>\n            <input type=\"text\" id=\"searchBar\" placeholder=\"Buscar archivo...\" />\n            <button id=\"searchButton\">Buscar</button>\n            <button id=\"clearSearch\">Limpiar b\u00FAsqueda</button>\n            <button id=\"saveButton\">Guardar</button>\n            <button onClick={actualizarContadorAccion} disabled={!accionSeleccionada}>\n              Guardar Acci\u00F3n\n            </button>\n            <button>Asignar a Transmittal</button>\n          </div>\n          <table>\n            <thead>\n              <tr>\n                <th><input type=\"checkbox\" id=\"masterCheckbox\"/></th>\n                <th>Nombre del Archivo</th>\n                <th>\u00BFDistribuido?</th>\n                <th>Acci\u00F3n</th>\n                <th>Destinatario</th>\n              </tr>\n            </thead>\n\n            <tbody id=\"tableBody\">\n              ").concat((this.state.lista || []).map(function (item) { return "\n                <tr>\n                  <td><input type=\"checkbox\" class=\"checkbox\" data-id=\"".concat(item.id, "\" /></td>\n                  <td>").concat(item.name, "</td>\n                  <td>No</td>\n                  <td>Sin acci\u00F3n</td>\n                  <td>Sin destinatario</td>\n                </tr>\n              "); }).join(""), "\n            </tbody>\n          </table>\n          <script>\n            \n            // Funci\u00F3n para validar el formulario y habilitar/deshabilitar el bot\u00F3n de asignaci\u00F3n\n\n            function validateForm() {\n            let checkboxesChecked = document.querySelectorAll('.checkbox:checked').length > 0;\n            let actionSelected = document.getElementById('action').value !== \"\";\n            let recipientSelected = document.getElementById('recipient').value !== \"\";\n\n            console.log(\"Checkboxes seleccionados:\", checkboxesChecked);\n            console.log(\"Acci\u00F3n seleccionada:\", actionSelected);\n            console.log(\"Destinatario seleccionado:\", recipientSelected);\n\n            let numeroTransmittal = document.querySelector(\"#idDelInput\")?.value || \"\";\n            if (!numeroTransmittal) {\n                console.error(\"Error: numeroTransmittal tiene un valor vac\u00EDo.\");\n            }\n\n            let assignButton = document.getElementById('assignDistributionButton');\n            assignButton.disabled = !(checkboxesChecked && actionSelected && recipientSelected);\n            }\n\n            // Asignar eventos de cambio a selects y checkboxes\n\n            document.getElementById('action').addEventListener('change', validateForm);\n            document.getElementById('recipient').addEventListener('change', validateForm);\n            document.querySelectorAll('.checkbox').forEach(checkbox => {\n              checkbox.addEventListener('change', validateForm);\n            });\n\n            // boton para asignar distribuci\u00F3n\n\n            document.getElementById('assignDistributionButton').addEventListener(\"click\", async () => {\n              console.log(\"Bot\u00F3n 'Asignar a Distribuci\u00F3n' presionado.\");\n              \n              const actionValue = document.getElementById(\"action\").value;\n              const recipientValue = document.getElementById(\"recipient\").value;\n              \n              const checkboxes = document.querySelectorAll('.checkbox:checked');\n\n              if (actionValue && recipientValue && checkboxes.length > 0) {\n                  checkboxes.forEach((checkbox) => {\n                      const row = checkbox.closest(\"tr\");\n                      if (row) {\n                          const nombreArchivo = row.cells[1].innerText;\n                          row.cells[2].innerText = \"S\u00ED\";\n                          row.cells[3].innerText = actionValue;\n                          row.cells[4].innerText = recipientValue;\n\n                          console.log(\"Llamando a guardarEnSharePoint:\", nombreArchivo, actionValue, recipientValue, numeroTransmittal, fecha);\n                          window.opener.guardarEnSharePoint(nombreArchivo, accion, destinatario, numeroTransmittal, fecha);\n                      }\n                  });\n                  console.log(\"Llamando a actualizarContadores:\", actionValue, recipientValue);\n                  window.opener.actualizarContadores(actionValue, recipientValue);\n                  alert(\"Archivos asignados correctamente\");\n              } else {\n                alert(\"Debe seleccionar una acci\u00F3n, un destinatario y al menos un archivo\");\n              }\n            });\n\n            // Evento para seleccionar o deseleccionar todos los checkboxes\n\n            document.getElementById(\"masterCheckbox\").addEventListener(\"change\", function() {\n                let checkboxes = document.querySelectorAll(\".checkbox\");\n                checkboxes.forEach(checkbox => {\n                    checkbox.checked = this.checked;\n                });\n                validateForm();\n            });\n\n            // Bot\u00F3n limpiar\n            \n            document.getElementById('clearButton').addEventListener('click', function() {\n                // Obtiene todos los checkboxes con la clase 'checkbox'\n                let checkboxes = document.querySelectorAll('.checkbox');\n\n                // Recorre cada checkbox y lo desmarca\n                checkboxes.forEach(checkbox => {\n                    checkbox.checked = false;\n                });\n\n                // Restablece los selects a su valor predeterminado\n                const actionSelect = document.getElementById(\"action\");\n                const recipientSelect = document.getElementById(\"recipient\");\n\n                if (actionSelect) actionSelect.selectedIndex = 0;\n                if (recipientSelect) recipientSelect.selectedIndex = 0;\n\n                // Ejecuta la validaci\u00F3n nuevamente (si la funci\u00F3n est\u00E1 definida)\n                if (typeof validateForm === \"function\") {\n                    validateForm();\n                } else {\n                    console.warn(\"La funci\u00F3n validateForm no est\u00E1 definida.\");\n                }\n            });\n\n            // BUSCAR ELEMENTO/S\n\n            document.getElementById(\"searchButton\").addEventListener(\"click\", function() {\n                let searchValue = document.getElementById(\"searchBar\").value.toLowerCase(); // Obtener texto de b\u00FAsqueda\n                let rows = document.querySelectorAll(\"tbody tr\"); // Obtener todas las filas de la tabla\n                let foundCount = 0; // Contador de resultados encontrados\n                \n                rows.forEach(row => {\n                    let fileNameCell = row.cells[1]; // Segunda columna (nombre del archivo)\n                    let checkbox = row.querySelector(\".checkbox\"); // Buscar checkbox en la fila\n\n                    if (fileNameCell.innerText.toLowerCase().includes(searchValue)) {\n                        row.style.backgroundColor = \"#ffff99\"; // Resaltar la fila en amarillo\n                        if (checkbox) checkbox.checked = true; // Marcar checkbox\n                        foundCount++; // Incrementar el contador de resultados encontrados\n                    } else {\n                        row.style.backgroundColor = \"\"; // Restaurar color original\n                        if (checkbox) checkbox.checked = false; // Desmarcar checkbox\n                    }\n                });\n\n                alert(\"Resultados: \" + foundCount);\n\n                validateForm(); // Actualizar validaci\u00F3n del formulario\n            });\n\n            // Bot\u00F3n para limpiar la b\u00FAsqueda\n\n            document.getElementById(\"clearSearch\").addEventListener(\"click\", function() {\n                document.getElementById(\"searchBar\").value = \"\"; // Vaciar el campo de b\u00FAsqueda\n                let rows = document.querySelectorAll(\"tbody tr\");\n\n                rows.forEach(row => {\n                    row.style.backgroundColor = \"\"; // Restaurar colores\n                    let checkbox = row.querySelector(\".checkbox\");\n                    if (checkbox) checkbox.checked = false; // Desmarcar todos los checkboxes\n                });\n\n                validateForm(); // Actualizar validaci\u00F3n del formulario\n            });          \n\n            // boton \"guardar\"\n\n            document.getElementById(\"saveButton\").addEventListener(\"click\", async () => {\n                console.log(\"Bot\u00F3n 'Guardar' presionado.\");\n\n                // Obtiene todos los checkboxes seleccionados\n                const checkboxes = document.querySelectorAll(\".checkbox:checked\");\n\n                if (checkboxes.length === 0) {\n                    alert(\"Debe seleccionar al menos un archivo para guardar.\");\n                    return;\n                }\n\n                for (let checkbox of checkboxes) {\n                    const row = checkbox.closest(\"tr\");\n                    if (row) {\n                        const nombreArchivo = row.cells[1].innerText;\n                        const accion = row.cells[3].innerText;\n                        const destinatario = row.cells[4].innerText;\n\n                        // Genera un n\u00FAmero transmittal aleatorio\n                        const numeroTransmittal = Math.floor(100000 + Math.random() * 900000);\n\n                        // Obtiene la fecha actual en formato YYYY-MM-DD HH:MM:SS\n                        const fecha = new Date().toISOString().slice(0, 19).replace(\"T\", \" \");\n\n                        // Depuraci\u00F3n: Verificar valores antes de enviar\n                        console.log(\"Datos antes de guardar en SharePoint:\", {\n                            nombreArchivo, accion, destinatario, numeroTransmittal, fecha\n                        });\n\n                        // Llamamos a la API REST de SharePoint para guardar los datos\n                        window.opener.guardarEnSharePoint(nombreArchivo, accion, destinatario, numeroTransmittal, fecha);\n                    }\n                }\n\n                alert(\"Datos guardados correctamente en SharePoint.\");\n            });\n\n\n            // Ejecutar la validaci\u00F3n al abrir la ventana\n\n            validateForm(); \n          </script>\n        </body>\n      </html>\n    "));
                    popupWindow.document.close();
                    // Se ejecuta después de 500ms para asegurarse de que los elementos estén cargados en la ventana emergente.
                    setTimeout(function () {
                        var assignButton = popupWindow.document.getElementById("assignDistributionButton");
                        if (assignButton) {
                            assignButton.disabled = false; // Habilita el botón antes de agregar el evento
                            assignButton.addEventListener("click", function () {
                                // Aquí se debe agregar la lógica para manejar los datos del transmittal.
                                alert("Distribuido"); // Alerta cuando se presiona el botón
                            });
                        }
                        else {
                            console.error("Error: No se encontró el botón 'Asignar a Distribución'");
                        }
                    }, 500);
                    // Evento adicional para manejar la asignación de distribución en la ventana emergente
                    // Se agrega un evento `click` al botón con id `assignDistributionButton` en la ventana emergente.
                    // Se obtiene el botón con ID 'assignDistributionButton' dentro de la ventana emergente.
                    // `?.` asegura que si el elemento no existe, no se lance un error.
                    (_a = popupWindow.document.getElementById('assignDistributionButton')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', function () { return __awaiter(_this, void 0, void 0, function () {
                        var actionValue, recipientValue, error_6;
                        var _a, _b;
                        return __generator(this, function (_c) {
                            switch (_c.label) {
                                case 0:
                                    _c.trys.push([0, 4, , 5]);
                                    console.log("Botón 'Asignar a Distribución' presionado."); // Mensaje en la consola para confirmar el evento.
                                    actionValue = (_a = popupWindow.document.getElementById("action")) === null || _a === void 0 ? void 0 : _a.value;
                                    recipientValue = (_b = popupWindow.document.getElementById("recipient")) === null || _b === void 0 ? void 0 : _b.value;
                                    // Se muestra en consola qué acción y destinatario fueron seleccionados.
                                    console.log("Acci\u00F3n seleccionada: ".concat(actionValue, ", Destinatario seleccionado: ").concat(recipientValue));
                                    if (!(actionValue && recipientValue)) return [3 /*break*/, 2];
                                    // Se llama a la función `actualizarContadores`, que posiblemente actualiza estadísticas
                                    // o algún estado de la aplicación sobre las asignaciones realizadas.
                                    return [4 /*yield*/, this.actualizarContadores(actionValue, recipientValue)];
                                case 1:
                                    // Se llama a la función `actualizarContadores`, que posiblemente actualiza estadísticas
                                    // o algún estado de la aplicación sobre las asignaciones realizadas.
                                    _c.sent();
                                    return [3 /*break*/, 3];
                                case 2:
                                    // Si no se seleccionaron valores, se muestra una alerta indicando que son obligatorios.
                                    alert("Debe seleccionar una acción y un destinatario.");
                                    _c.label = 3;
                                case 3: return [3 /*break*/, 5];
                                case 4:
                                    error_6 = _c.sent();
                                    // En caso de que ocurra un error en la ejecución de `actualizarContadores`,
                                    // se captura y se muestra en la consola.
                                    console.error("Error en actualizarContadores:", error_6);
                                    return [3 /*break*/, 5];
                                case 5: return [2 /*return*/];
                            }
                        });
                    }); });
                    // Método `actualizarContadores`: Incrementa los contadores de acciones y destinatarios en SharePoint.
                    // Recibe los valores de acción (`actionValue`) y destinatario (`recipientValue`) como parámetros.
                }
                return [2 /*return*/];
            });
        }); };
        _this.state = {
            lista: [],
            actions: [],
            recipients: [],
            isAssignDisabled: true,
            showTable: false,
            actionCount: {},
            accionSeleccionada: null,
            contadorAccion: 0 // Agregar aquí el contador de la acción seleccionada
        };
        // Se enlazan los métodos a la instancia de la clase para garantizar el acceso al contexto `this`.
        // Esto es necesario en clases de React que extienden `Component`, ya que los métodos 
        // no están automáticamente ligados al contexto de la instancia de la clase.
        _this.mostrarVentana = _this.mostrarVentana.bind(_this); // Enlaza el método `mostrarVentana` al contexto de la clase.
        _this.handleButtonClick = _this.handleButtonClick.bind(_this); // Enlaza el método `handleButtonClick` al contexto de la clase.
        _this.guardarEnSharePoint = _this.guardarEnSharePoint.bind(_this); // Enlaza el método `guardarEnSharePoint` al contexto de la clase.
        _this.actualizarContadores = _this.actualizarContadores.bind(_this); // Enlaza el método `actualizarContadores` al contexto de la clase.
        // Se agregan las funciones `guardarEnSharePoint` y `actualizarContadores` al objeto global `window`,
        // permitiendo que sean accesibles desde cualquier parte de la aplicación.
        window.guardarEnSharePoint = _this.guardarEnSharePoint.bind(_this); // Se asigna la función `guardarEnSharePoint` al objeto global `window`.
        window.actualizarContadores = _this.actualizarContadores.bind(_this); // Se asigna la función `actualizarContadores` al objeto global `window`.
        return _this;
    }
    // ✅ Ahora podemos usar el contexto de SharePoint de forma segura
    HolaMundo2.prototype.verificarConexion = function () {
        return __awaiter(this, void 0, void 0, function () {
            var listas, error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists.select("Title", "Id").expand("Fields")()];
                    case 1:
                        listas = _a.sent();
                        console.log("Listas en SharePoint:", listas);
                        return [3 /*break*/, 3];
                    case 2:
                        error_7 = _a.sent();
                        console.error("Error al conectar con SharePoint:", error_7);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    HolaMundo2.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement("h2", null, "Hola Mundo 2"),
            React.createElement("button", { onClick: this.handleButtonClick }, "Asignar a Transmittal")));
    };
    HolaMundo2.prototype.actualizarContadorAccion = function (actionTitle) {
        return __awaiter(this, void 0, void 0, function () {
            var url, response, data, accionId, nuevoContador, updateUrl, updateBody, createUrl, createBody, error_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 7, , 8]);
                        url = "".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('Accion')/items?$filter=Title eq '").concat(encodeURIComponent(actionTitle), "'&$select=Id,ContadorAccion");
                        return [4 /*yield*/, this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        if (!(data.value.length > 0)) return [3 /*break*/, 4];
                        accionId = data.value[0].Id;
                        nuevoContador = (data.value[0].ContadorAccion || 0) + 1;
                        updateUrl = "".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('Accion')/items(").concat(accionId, ")");
                        updateBody = JSON.stringify({
                            ContadorAccion: nuevoContador
                        });
                        return [4 /*yield*/, this.props.context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=verbose',
                                    'Content-Type': 'application/json;odata=verbose',
                                    //'X-RequestDigest': (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value,
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: updateBody
                            })];
                    case 3:
                        _a.sent();
                        console.log("\u2705 Acci\u00F3n \"".concat(actionTitle, "\" actualizada con contador: ").concat(nuevoContador));
                        return [3 /*break*/, 6];
                    case 4:
                        createUrl = "".concat(this.props.context.pageContext.web.absoluteUrl, "/_api/web/lists/GetByTitle('Accion')/items");
                        createBody = JSON.stringify({
                            Title: actionTitle,
                            ContadorAccion: 1
                        });
                        return [4 /*yield*/, this.props.context.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=verbose',
                                    'Content-Type': 'application/json;odata=verbose',
                                    //'X-RequestDigest': (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value
                                },
                                body: createBody
                            })];
                    case 5:
                        _a.sent();
                        console.log("\u2705 Nueva acci\u00F3n \"".concat(actionTitle, "\" creada con contador: 1"));
                        _a.label = 6;
                    case 6: return [3 /*break*/, 8];
                    case 7:
                        error_8 = _a.sent();
                        console.error("❌ Error al actualizar contador de acciones:", error_8);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    // Método privado `guardarEnSharePoint`: Guarda un nuevo registro en la lista "Distribucion" de SharePoint.
    // Recibe como parámetros el nombre del archivo, la acción y el destinatario.
    HolaMundo2.prototype.guardarEnSharePoint = function (nombreArchivo, accion, destinatario, numeroTransmittal, fecha) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        console.log("\uD83D\uDCCC Guardando en SharePoint: ".concat(nombreArchivo, " / ").concat(accion, " / ").concat(destinatario, " / ").concat(numeroTransmittal, " / ").concat(fecha));
                        // Verifica que los valores no estén vacíos
                        if (!nombreArchivo || !accion || !destinatario || !numeroTransmittal || !fecha) {
                            throw new Error("❌ Faltan datos obligatorios para guardar en SharePoint.");
                        }
                        return [4 /*yield*/, sp.web.lists.getByTitle("Distribucion").items.add({
                                Title: nombreArchivo,
                                Accion: accion,
                                Destinatario: destinatario,
                                NumeroTransmittal: numeroTransmittal,
                                FechaDistribucion: fecha // Fecha de distribución (verifica si el campo existe)
                            })];
                    case 1:
                        response = _a.sent();
                        if (response === null || response === void 0 ? void 0 : response.data) {
                            console.log("\u2705 Archivo guardado en SharePoint correctamente:", response.data);
                            alert("Archivo guardado correctamente:\nArchivo: ".concat(nombreArchivo, "\nAcci\u00F3n: ").concat(accion, "\nDestinatario: ").concat(destinatario, "\nN\u00FAmero Transmittal: ").concat(numeroTransmittal, "\nFecha: ").concat(fecha));
                        }
                        else {
                            throw new Error("No se recibió una respuesta válida de SharePoint.");
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        error_9 = _a.sent();
                        console.error("Error al guardar en SharePoint:", error_9);
                        alert("Error al guardar en SharePoint. Revisa la consola para más detalles.");
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    // Método del ciclo de vida de React `componentDidMount`.
    // Se ejecuta automáticamente después de que el componente se ha montado en el DOM.
    // En este caso, se utiliza para obtener datos iniciales necesarios para el funcionamiento del componente.
    HolaMundo2.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: 
                    // Llama al método `getData` para obtener la lista de documentos técnicos desde SharePoint.
                    return [4 /*yield*/, this.getData()];
                    case 1:
                        // Llama al método `getData` para obtener la lista de documentos técnicos desde SharePoint.
                        _a.sent();
                        // Llama al método `getActions` para obtener la lista de acciones disponibles.
                        return [4 /*yield*/, this.getActions()];
                    case 2:
                        // Llama al método `getActions` para obtener la lista de acciones disponibles.
                        _a.sent();
                        // Llama al método `getRecipients` para obtener la lista de destinatarios.
                        return [4 /*yield*/, this.getRecipients()];
                    case 3:
                        // Llama al método `getRecipients` para obtener la lista de destinatarios.
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    HolaMundo2.prototype.actualizarContadores = function (actionValue, recipientValue) {
        return __awaiter(this, void 0, void 0, function () {
            var accionItems, accionContador, accionItemId, newItem, destinatarioItems, destinatarioContador, destinatarioItemId, newItem, error_10;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!actionValue || !recipientValue) {
                            alert("Debe seleccionar una acción y un destinatario.");
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 12, , 13]);
                        console.log("\uD83D\uDCCC Buscando acci\u00F3n en SharePoint con filtro: Title eq '".concat(actionValue, "'"));
                        return [4 /*yield*/, sp.web.lists.getByTitle("Accion").items
                                .filter("Title eq '".concat(actionValue, "'")).top(1)()];
                    case 2:
                        accionItems = _a.sent();
                        accionContador = 0;
                        accionItemId = null;
                        if (!(accionItems.length > 0)) return [3 /*break*/, 4];
                        accionItemId = accionItems[0].Id;
                        accionContador = (accionItems[0].Acciones_x0020_asignadas || 0) + 1;
                        // 🔹 Actualizar en SharePoint usando REST API (PATCH)
                        return [4 /*yield*/, sp.web.lists.getByTitle("Accion").items.getById(accionItemId).update({
                                Acciones_x0020_asignadas: accionContador
                            })];
                    case 3:
                        // 🔹 Actualizar en SharePoint usando REST API (PATCH)
                        _a.sent();
                        console.log("Acci\u00F3n actualizada en SharePoint: ".concat(actionValue, " N\u00B0").concat(accionContador));
                        return [3 /*break*/, 6];
                    case 4: return [4 /*yield*/, sp.web.lists.getByTitle("Accion").items.add({
                            Title: actionValue,
                            Acciones_x0020_asignadas: 1
                        })];
                    case 5:
                        newItem = _a.sent();
                        accionItemId = newItem.data.Id;
                        accionContador = 1;
                        console.log("Nueva acci\u00F3n creada en SharePoint: ".concat(actionValue, " N\u00B0").concat(accionContador));
                        _a.label = 6;
                    case 6:
                        console.log("\uD83D\uDCCC Buscando destinatario \"".concat(recipientValue, "\" en SharePoint..."));
                        return [4 /*yield*/, sp.web.lists.getByTitle("Destinatarios").items
                                .filter("Title eq '".concat(recipientValue, "'")).top(1)()];
                    case 7:
                        destinatarioItems = _a.sent();
                        destinatarioContador = 0;
                        destinatarioItemId = null;
                        if (!(destinatarioItems.length > 0)) return [3 /*break*/, 9];
                        destinatarioItemId = destinatarioItems[0].Id;
                        destinatarioContador = (destinatarioItems[0].Destinatarios_x0020_asignados || 0) + 1;
                        // 🔹 Actualizar en SharePoint usando REST API (PATCH)
                        return [4 /*yield*/, sp.web.lists.getByTitle("Destinatarios").items.getById(destinatarioItemId).update({
                                Destinatarios_x0020_asignados: destinatarioContador
                            })];
                    case 8:
                        // 🔹 Actualizar en SharePoint usando REST API (PATCH)
                        _a.sent();
                        console.log("\u2705 Destinatario actualizado en SharePoint: ".concat(recipientValue, " N\u00B0").concat(destinatarioContador));
                        return [3 /*break*/, 11];
                    case 9: return [4 /*yield*/, sp.web.lists.getByTitle("Destinatarios").items.add({
                            Title: recipientValue,
                            Destinatarios_x0020_asignados: 1
                        })];
                    case 10:
                        newItem = _a.sent();
                        destinatarioItemId = newItem.data.Id;
                        destinatarioContador = 1;
                        console.log("\u2705 Nuevo destinatario creado en SharePoint: ".concat(recipientValue, " N\u00B0").concat(destinatarioContador));
                        _a.label = 11;
                    case 11:
                        alert("Datos almacenados correctamente en SharePoint.");
                        return [3 /*break*/, 13];
                    case 12:
                        error_10 = _a.sent();
                        console.error("❌ Error al actualizar datos en SharePoint:", error_10);
                        return [3 /*break*/, 13];
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    return HolaMundo2;
}(React.Component));
export default HolaMundo2;
//# sourceMappingURL=HolaMundo2.js.map