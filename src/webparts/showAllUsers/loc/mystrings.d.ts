declare interface IShowAllUsersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  //SearchFor: string;
  //SearchForValidationErrorMessage:string;
}

declare module 'ShowAllUsersWebPartStrings' {
  const strings: IShowAllUsersWebPartStrings;
  export = strings;
}
