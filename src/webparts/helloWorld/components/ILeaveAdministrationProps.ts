import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveAdministrationProps {
  context: WebPartContext;
  title: string;
  description: string;
  defaultView: string;
  showPendingOnly: boolean;
  itemsPerPage: number;
  allowBulkActions: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}