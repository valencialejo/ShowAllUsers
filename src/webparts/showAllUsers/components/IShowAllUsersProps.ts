import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IShowAllUsersProps {
  description: string;
  context: WebPartContext;
  webparttype:string;
  InitDate:Date;
  EndDate:Date;
}
