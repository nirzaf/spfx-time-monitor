import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITeamCalendarProps {
  context: WebPartContext;
  title: string;
  description: string;
  defaultView: string;
  showWeekends: boolean;
  allowExport: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
