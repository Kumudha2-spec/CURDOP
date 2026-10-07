import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICurdOperationsSpfxProps {
  description: string;
  listName: string;
  defaultRating: number;
  showRating: boolean;
  allowAnonymous: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  context: WebPartContext;
}