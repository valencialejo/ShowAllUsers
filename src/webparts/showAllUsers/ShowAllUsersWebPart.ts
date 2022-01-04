import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ShowAllUsersWebPartStrings';
import ShowAllUsers from './components/ShowAllUsers';
import { IShowAllUsersProps } from './components/IShowAllUsersProps';
import { thProperties } from 'office-ui-fabric-react';

export interface IShowAllUsersWebPartProps {
  description: string;
  webparttype:string;
}

export default class ShowAllUsersWebPart extends BaseClientSideWebPart<IShowAllUsersWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IShowAllUsersProps > = React.createElement(
      ShowAllUsers,
      {
        description: this.properties.description,
        context:this.context,
        webparttype:this.properties.webparttype
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
                  ]
                }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}
