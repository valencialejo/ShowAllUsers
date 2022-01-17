import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IShowAllUsersProps {
  description: string;
  context: WebPartContext;
  webparttype:string;
  TodayDate:Date;
  WeekDate:Date;
  MonthDate:Date;
}
