import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneButtonType,
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ShowAllUsersWebPartStrings';
import ShowAllUsers from './components/ShowAllUsers';
import { IShowAllUsersProps } from './components/IShowAllUsersProps';
import { thProperties } from 'office-ui-fabric-react';

export interface IShowAllUsersWebPartProps {
  description: string;
  webparttype:string;
  TodayDate:Date;
  WeekDate:Date;
  MonthDate:Date;
}

export default class ShowAllUsersWebPart extends BaseClientSideWebPart<IShowAllUsersWebPartProps> {

  protected onInit():Promise<void>{

    return new Promise<void>((resolve,_reject)=>{

      var currentDate:Date=new Date(); var currentDay=currentDate.getUTCDate(); var currentMonth=currentDate.getUTCMonth();
      this.properties.TodayDate=new Date(2000,currentMonth,currentDay);
      this.properties.WeekDate=new Date(2000,currentMonth,currentDay+5);
      this.properties.MonthDate=new Date(2000,currentMonth,currentDay+30);

      console.log(this.properties.TodayDate);
      console.log(this.properties.WeekDate);
      console.log(this.properties.MonthDate);
      resolve(undefined);
    });
  }

  public render(): void {
    const element: React.ReactElement<IShowAllUsersProps> = React.createElement(
      ShowAllUsers,
      {
        description: this.properties.description,
        context:this.context,
        webparttype:this.properties.webparttype,
        TodayDate:this.properties.TodayDate,
        WeekDate:this.properties.WeekDate,
        MonthDate:this.properties.MonthDate
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  } 

  private ButtonClick(oldVal: any): any {
    
    // var currentDate:Date=new Date(); var currentDay=currentDate.getUTCDate(); var currentMonth=currentDate.getUTCMonth();
    
    // this.properties.InitDate=new Date(2000,currentMonth,currentDay);
    // this.properties.EndDate=new Date(2000,currentMonth,currentDay);
    // var type=this.properties.webparttype;
    // var date=this.properties.InitDate;
    // var days: number;

    // if (type=='today'){
    //   days=0; 
    // }else if(type=='week'){
    //   days=5;
    // }else if(type=='month'){
    //   days=30;
    // }

    // this.properties.EndDate.setDate(date.getDate()+days);
    console.log(this.properties.TodayDate); console.log(this.properties.WeekDate); console.log(this.properties.MonthDate);
    console.log('¡Updated succesfully!');
}  

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneChoiceGroup('webparttype',{
                  label:"Webpart type",
                  options:[
                    {
                      checked:true,
                      key:"today",
                      text:"Hoy",
                      iconProps:{
                        officeFabricIconFontName:'GotoToday'
                      }
                    },
                    {
                      key:"week",
                      text:"Semana",
                      iconProps:{
                        officeFabricIconFontName:'CalendarWeek'
                      }
                    },
                    {
                      key:"month",
                      text:"Mes",
                      iconProps:{
                        officeFabricIconFontName:'Calendar'
                      }
                    },
                  ],
                }
                ),
                PropertyPaneButton('updateButton',{
                  text:'Actualizar',
                  onClick:this.ButtonClick.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
