import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHolaMundo2Props {
  context: WebPartContext; // Asegúrate de importar esto si es necesario
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}