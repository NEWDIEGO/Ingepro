import { SPFI, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


// Inicializar SharePoint PnP.js con la URL del sitio
const sp: SPFI = spfi("https://ntechilespa.sharepoint.com/sites/ingp");

import "@pnp/sp/items/list";
//import { IList } from "@pnp/sp/lists";
//import { IItem, IItems } from "@pnp/sp/items";


export const getSP = (): SPFI => sp;
export { sp };

// Función para agregar un elemento a la lista Distribución
export const agregarADistribucion = async (
    nombre: string,
    accion: string,
    destinatario: string,
    numeroTransmittal: string
): Promise<boolean> => {
    try {
        await sp.web.lists.getByTitle("Distribución").items.add({
            Title: nombre,
            Accion: accion,
            Destinatario: destinatario,
            NumeroTransmittal: numeroTransmittal
        });

        return true; // ✅ Retorna true si la operación fue exitosa
    } catch (error) {
        console.error("Error al agregar a Distribución:", error);
        return false; // ❌ Retorna false si hubo un error
    }
};