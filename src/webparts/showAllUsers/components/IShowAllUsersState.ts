import { IDates } from "./IDates";
import { IUser } from "./IUser"; 

export interface IShowAllUsersState{
    users:Array<IUser>;
    usersView:Array<IUser>;
    dateofSearch:IDates;
    user:IUser;
}