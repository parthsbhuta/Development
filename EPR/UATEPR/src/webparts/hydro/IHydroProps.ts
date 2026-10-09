import { WebPartContext } from '@microsoft/sp-webpart-base';
import SharePointService from '../service/SharePointService';

export interface IHydroProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  spService: SharePointService;
}
