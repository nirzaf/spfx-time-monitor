import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveHistoryProps {
  context: WebPartContext;
  title: string;
  description: string;
  defaultView: string;
  itemsPerPage: number;
  showAnalytics: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
