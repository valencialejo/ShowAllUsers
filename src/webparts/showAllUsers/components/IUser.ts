export interface IUser{
    displayName:string;
    givenName:string;
    surname:string;
    mail:string;
    mobilePhone:string;
    userPrincipalName:string;
    birthday?:Date;
    birthdayDate?:string;
    birthdayMonth?:string;
    accountEnabled?:Boolean;
    aboutMe?:string;
}