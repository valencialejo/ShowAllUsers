export interface IUser{
    id:string;
    displayName:string;
    givenName:string;
    surname:string;
    mail:string;
    mobilePhone:string;
    userPrincipalName:string;
    birthday?:any;
    birthdayDate?:string;
    birthdayMonth?:string;
    accountEnabled?:Boolean;
    aboutMe?:string;
    jobTitle?:string;
    department?:string;
    profilePhoto?:string;
}