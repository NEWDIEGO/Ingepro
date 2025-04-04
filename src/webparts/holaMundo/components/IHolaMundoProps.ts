import { SPFI } from "@pnp/sp";  // Importa SPFI si necesitas pasar la instancia de SharePoint

export interface IHolaMundoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
  sp: SPFI;  // Incluye esto si necesitas pasar la instancia de SP
}
