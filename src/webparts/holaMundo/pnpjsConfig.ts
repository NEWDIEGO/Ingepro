import { SPFI, spfi } from "@pnp/sp";
import "@pnp/sp/webs";  // Importar extensiones necesarias
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { SPFx } from "@pnp/sp"; // Importa SPFx para el contexto de SPFx
import { WebPartContext } from "@microsoft/sp-webpart-base"; // Importa el tipo correcto para context

// Función para obtener la instancia de SP configurada
export const getSP = (context: WebPartContext): SPFI => { // Usa WebPartContext en lugar de any
  // Crea una instancia de SPFI usando el contexto de SPFx
  return spfi().using(SPFx(context));
};
