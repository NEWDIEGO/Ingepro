import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";

export default class HolaMundoWebPart extends BaseClientSideWebPart<{}> {
    private sp: SPFI; // Declara el objeto SPFI como propiedad de la clase

    public async render(): Promise<void> {
        // Inicializar SPFI con el contexto de SPFx
        this.sp = spfi().using(SPFx(this.context)); // Usa this.context de BaseClientSideWebPart

        try {
            // Obtener todas las listas del sitio actual
            console.log("Obteniendo listas desde SharePoint...");
            const lists = await this.sp.web.lists();
            console.log("Listas obtenidas correctamente:", lists);

            if (!lists) {
                console.error("No se pudieron obtener las listas.");
                return;
            }

            // Contar todas las listas
            console.log("Número de listas y bibliotecas: ", lists.length);

            // Filtrar solo las bibliotecas de documentos (BaseTemplate 101)
            const documentLibraries = lists.filter(list => list.BaseTemplate === 101);

            console.log("Número de bibliotecas de documentos: ", documentLibraries.length);
        } catch (error) {
            console.error("Error al obtener las listas: ", error);
        }
    }
}
