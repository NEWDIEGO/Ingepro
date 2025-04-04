import * as React from "react";
import { IHolaMundo2Props } from "./IHolaMundo2Props";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
interface IDocumentoTecnico {
    id: number;
    name: string;
}
declare global {
    interface Window {
        guardarEnSharePoint: (nombreArchivo: string, accion: string, destinatario: string, numeroTransmittal: number, // Nuevo parámetro agregado
        fecha: string) => Promise<void>;
        actualizarContadores: (actionValue: string, recipientValue: string) => Promise<void>;
    }
}
interface IHolaMundo2State {
    lista: IDocumentoTecnico[];
    actions: string[];
    recipients: string[];
    isAssignDisabled: boolean;
    showTable: boolean;
    actionCount: {
        [key: string]: number;
    };
}
export default class HolaMundo2 extends React.Component<IHolaMundo2Props, IHolaMundo2State> {
    private: SPFI;
    constructor(props: IHolaMundo2Props);
    setAccionSeleccionada: (accion: string) => void;
    setContadorAccion: (contador: number) => void;
    verificarConexion(): Promise<void>;
    render(): JSX.Element;
    mostrarVentana: () => void;
    private actualizarContadorAccion;
    private guardarEnSharePoint;
    handleAsignarDistribucion: () => Promise<void>;
    componentDidMount(): Promise<void>;
    private getData;
    private getActions;
    private getRecipients;
    private handleButtonClick;
    private actualizarContadores;
}
export {};
//# sourceMappingURL=HolaMundo2.d.ts.map