import { spfi } from "@pnp/sp";
import "@pnp/sp/webs"; // Importar extensiones necesarias
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { SPFx } from "@pnp/sp"; // Importa SPFx para el contexto de SPFx
// Función para obtener la instancia de SP configurada
export var getSP = function (context) {
    // Crea una instancia de SPFI usando el contexto de SPFx
    return spfi().using(SPFx(context));
};
//# sourceMappingURL=pnpjsConfig.js.map