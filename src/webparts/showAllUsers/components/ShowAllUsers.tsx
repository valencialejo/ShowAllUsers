import * as React from 'react';
import styles from './ShowAllUsers.module.scss';
import { IShowAllUsersProps } from './IShowAllUsersProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IUser } from './IUser';
import { IShowAllUsersState } from './IShowAllUsersState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
} from 'office-ui-fabric-react';

import * as strings from 'ShowAllUsersWebPartStrings';
import { now } from 'lodash';
//import { SearchFor } from 'ShowAllUsersWebPartStrings';

//Configurar las columnas para el componente DetailsList
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  },
  // {
  //   key: 'givenName',
  //   name: 'Given Name',
  //   fieldName: 'givenName',
  //   minWidth: 50,
  //   maxWidth: 100,
  //   isResizable: true
  // },
  // {
  //   key: 'surName',
  //   name: 'SurName',
  //   fieldName: 'surname',
  //   minWidth: 50,
  //   maxWidth: 100,
  //   isResizable: true
  // },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true
  },
  // {
  //   key: 'mobilePhone',
  //   name: 'mobile Phone',
  //   fieldName: 'mobilePhone',
  //   minWidth: 50,
  //   maxWidth: 100,
  //   isResizable: true
  // },
  // {
  //   key: 'userPrincipalName',
  //   name: 'User Principal Name',
  //   fieldName: 'userPrincipalName',
  //   minWidth: 200,
  //   maxWidth: 200,
  //   isResizable: true
  // },
  {
    key: 'birthday',
    name: 'Birthday',
    fieldName: 'birthday',
    minWidth: 200,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'aboutMe',
    name: 'aboutMe',
    fieldName: 'aboutMe',
    minWidth: 200,
    maxWidth: 200,
    isResizable: true
  },
  // {
  //   key: 'birthdayMonth',
  //   name: 'BirthdayMonth',
  //   fieldName: 'birthdayMonth',
  //   minWidth: 200,
  //   maxWidth: 200,
  //   isResizable: true
  // },
];

export default class ShowAllUsers extends React.Component<IShowAllUsersProps, IShowAllUsersState> {

  constructor(props: IShowAllUsersProps, state: IShowAllUsersState) {
    super(props);

    //Inicializar el State
    let date=new Date('2021-06-22');
    
    this.state = {
      user: undefined,
      users: [],
      usersView:[],
      dateofSearch: {
        fullDate:date.toString(),
        day:date.getUTCDate().toString(),
        month:(date.getUTCMonth()+1).toString(),
      },
    };
  }

  public componentDidMount(): void {
    this.fetchUserDetails();
  }
  @autobind
  public _search(): void {
    this.fetchUserDetails();
  }

  // @autobind
  // private _onSearchForChanged(newValue: string): void {
  //   this.setState({
  //     searchFor: newValue,
  //   });
  // }
  // private _getSearchForErrorMessage(value: string): string {
  //   return (value == null || value.length == 0 || value.indexOf(" ") < 0)
  //     ? ''
  //     : `${strings.SearchForValidationErrorMessage}`;
  // }

  public fetchUserDetails(): void {

    var birthdaytype:string;

    // if(this.props.webparttype=='today'){
    //   birthdaytype=
    // }

    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client
        .api('users')
        .version("v1.0")
        .select("*")
        .filter(`accountEnabled eq true`)
        .top(999)
        // and startswith(givenname,'${escape(this.state.searchFor)}')
        .get((error: any, response, rawResponse?: any) => {

          if (error) {
            console.error("Message is : " + error);
            return;
          }
          //console.log(response);

          //Preparar el array de salida
          var allUsers: Array<IUser> = new Array<IUser>();

          //Mapear el la respuesta JSON el array de salida
          response.value.map(async(item: IUser) => {
            if(item.mail!=null){
              let user:IUser={
                displayName: item.displayName,
                givenName: item.displayName,
                surname: item.surname,
                mail: item.mail,
                mobilePhone: item.mobilePhone,
                userPrincipalName: item.userPrincipalName,};
  
              await client
                .api(`users/${item.mail}`)
                .version("v1.0")
                .select("birthday,aboutMe")
  
                .get().then((response) => {
                  let userBirthday=new Date(response.birthday);
                  this.setState({ user:{...user, birthday:response.birthday,birthdayDate:userBirthday.getUTCDate().toString(),birthdayMonth:(userBirthday.getUTCMonth()+1).toString(),aboutMe:response.aboutMe} });
                  console.log(user.displayName+user.mail+response.birthday+(response.birthday>'1997-06-01' && response.birthday<'1997-06-30'));
                });
              
              allUsers.push(this.state.user);
            }
          });
          
          //console.log(Date.UTC(date.getFullYear(),date.getMonth(),date.getDate(),date.getHours(),date.getMinutes(),date.getSeconds()));
          //console.log(this)
          this.setState({ users: allUsers });
          //console.log(this.state.users);
          //console.log(this.props.webparttype);
        });
    });
  }

  public render(): React.ReactElement<IShowAllUsersProps> {
    return (
      <div className={styles.showAllUsers}>
        {/* <TextField
          label={strings.SearchFor}
          required={true}
          value={this.state.searchFor}
          onChanged={this._onSearchForChanged}
          onGetErrorMessage={this._getSearchForErrorMessage}
        /> */}

        {/* <p className={styles.title}>
          <PrimaryButton
            text='Search'
            title='Search'
            onClick={this._search}
          />
      </p> */}
        {
          (this.state.users != null && this.state.users.length > 0) ?
            <p className={styles.row}>
              <p>{this.props.webparttype}</p>
              <DetailsList
                items={this.state.users.filter(user=>user.birthdayDate==this.state.dateofSearch.day && user.birthdayMonth==this.state.dateofSearch.month)}
                columns={_usersListColumns}
                setKey='set'
                checkboxVisibility={CheckboxVisibility.onHover}
                selectionMode={SelectionMode.single}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                compact={true}
              />
            </p>
            : null
        }
      </div>
    );
  }
}
