import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveRequestFormProps {
  context: WebPartContext;
  title: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
