import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IMsGraphSpFxProps {
  graphClient: MSGraphClientV3;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}