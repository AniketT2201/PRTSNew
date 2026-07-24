import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPrtsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  currentSPContext?: any;
  themeVariant?: any;
  userEmail?:any;
  ItemID?:any;
  EmployeeId?:any;
  context?:any;
  Appadmin?:any;
  Maintenance?:any;
  SysAdmin?:any;  
  subcontext?:WebPartContext;
  Reader?:any;
}
