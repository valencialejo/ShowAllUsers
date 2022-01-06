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
  List,
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
    var InitDate:Date=new Date(2000,0,7);
    var EndDate:Date=new Date(2000,0,12);

    this.state = {
      user: undefined,
      users: [],
      usersView:[],
      dateofSearch: {
        fullInitDate:InitDate.toString(),
        dayInitDate:InitDate.getUTCDate().toString(),
        monthInitDate:(InitDate.getUTCMonth()+1).toString(),
        fullEndDate:EndDate.toString(),
        dayEndDate:EndDate.getUTCDate().toString(),
        monthEndDate:(EndDate.getUTCMonth()+1).toString(),
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

                  let userBirthday1=new Date(response.birthday);
                  let userBirthdayDay=userBirthday1.getUTCDate();
                  let userBirthdayMonth=userBirthday1.getMonth();
                  let userBirthday=new Date(2000,userBirthdayMonth,userBirthdayDay);

                  // var InitDate:Date=new Date(2000,0,7);
                  // var EndDate:Date=new Date(2000,0,12);

                  // let str1=userBirthday.toString();
                  // let str2=(userBirthday>=InitDate && userBirthday<=EndDate).toString();
                  // console.log(str1.concat(str2));
                  
                  this.setState({ user:{...user, birthday:userBirthday,aboutMe:response.aboutMe} });
                  //console.log(user.displayName+user.mail+response.birthday+(response.birthday>'1997-06-01' && response.birthday<'1997-06-30'));
                }); 
              allUsers.push(this.state.user);
            }
          });
          
          this.setState({ users: allUsers });
          //console.log(this.state.users);
          //console.log(this.props.webparttype);
        });
    });
  }

  public render(): React.ReactElement<IShowAllUsersProps> {
    return (
      <div>
        {this.state.users.filter(user=>user.birthday>=this.props.InitDate && user.birthday<=this.props.EndDate ).map(filteredUser=>(
          <>
          <TextField
            label='Nombre'
            value={filteredUser.displayName} />
          <TextField
            label='Cumpleaños'
            value={filteredUser.birthday.toString()}
          />
          </>
        ))}
      </div>
    );
    // return (
    //   <div className={styles.showAllUsers}>
    //     {/* <TextField
    //       label={strings.SearchFor}
    //       required={true}
    //       value={this.state.searchFor}
    //       onChanged={this._onSearchForChanged}
    //       onGetErrorMessage={this._getSearchForErrorMessage}
    //     /> */}

    //     {/* <p className={styles.title}>
    //       <PrimaryButton
    //         text='Search'
    //         title='Search'
    //         onClick={this._search}
    //       />
    //   </p> */}
    //     {
    //       (this.state.users != null && this.state.users.length > 0) ?
    //         <p className={styles.row}>
    //           <p>{this.props.webparttype}</p>
    //           <DetailsList
    //             items={this.state.users.filter(user=>user.birthdayDate==this.state.dateofSearch.dayInitDate && user.birthdayMonth==this.state.dateofSearch.monthInitDate)}
    //             columns={_usersListColumns}
    //             setKey='set'
    //             checkboxVisibility={CheckboxVisibility.onHover}
    //             selectionMode={SelectionMode.single}
    //             layoutMode={DetailsListLayoutMode.fixedColumns}
    //             compact={true}
    //           />
    //         </p>
    //         : null
    //     }
    //   </div>
    // );
  // }
}
}
