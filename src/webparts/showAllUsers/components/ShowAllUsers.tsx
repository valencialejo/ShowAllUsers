import * as React from 'react';
import * as ReactDOM from 'react-dom';

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
    var InitDate: Date = new Date(2000, 0, 7);
    var EndDate: Date = new Date(2000, 0, 12);

    this.state = {
      user: undefined,
      users: [],
      usersView: [],
      dateofSearch: {
        fullInitDate: InitDate.toString(),
        dayInitDate: InitDate.getUTCDate().toString(),
        monthInitDate: (InitDate.getUTCMonth() + 1).toString(),
        fullEndDate: EndDate.toString(),
        dayEndDate: EndDate.getUTCDate().toString(),
        monthEndDate: (EndDate.getUTCMonth() + 1).toString(),
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

  private _formatDate(date: Date): string {
    const day = date.toLocaleString('default', { day: '2-digit' });
    const month = date.toLocaleString('default', { month: 'short' });
    const year = date.toLocaleString('default', { year: 'numeric' });
    return day + ' ' + month[0].toUpperCase() + month.substring(1, month.length);
  }

  public async blobToB64(blob) {
    return new Promise((resolve, reject) => { const reader = new FileReader(); reader.readAsDataURL(blob); reader.onload = () => resolve(reader.result); reader.onerror = error => reject(error); });
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
          response.value.map(async (item: IUser) => {
            if (item.mail != null) {
              let user: IUser = {
                profilePhoto: await this.getProfilePhoto(client, item),
                ...await this.getBirthday(client, item),
                id: item.id,
                displayName: item.displayName,
                givenName: item.givenName,
                surname: item.surname,
                mail: item.mail,
                mobilePhone: item.mobilePhone,
                jobTitle: item.jobTitle,
                userPrincipalName: item.userPrincipalName,
              };

              this.setState({ user: user });
              console.log(this.state.user);
              allUsers.push(this.state.user);
            }
          });
          this.setState({ users: allUsers });
        });
    });
  }

  public async getBirthday(client, item): Promise<object> {

    return await client
      .api(`users/${item.id}`)
      .version("v1.0")
      .select("birthday,aboutMe,department")
      .get().then((response1) => {

        var userBirthday: any;
        let userBirthday1 = new Date(response1.birthday);
        let userBirthdayDay = userBirthday1.getUTCDate();
        let userBirthdayMonth = userBirthday1.getUTCMonth();
        let userBirthdayYear = userBirthday1.getFullYear();



        if (userBirthdayYear == 0) {
          userBirthday = null;
        } else {
          userBirthday = new Date(2000, userBirthdayMonth, userBirthdayDay);
        }
        console.log(item.displayName + ";" + item.mail + ";" + response1.birthday + ";" + userBirthdayDay + ";" + userBirthdayMonth + ";" + userBirthday);
        return { birthday: userBirthday, aboutMe: response1.aboutMe, department: response1.department };
      });
  }

  public async getProfilePhoto(client, item): Promise<string> {
    try {
      return await client
        .api(`users/${item.id}/photo/$value`)
        .responseType("blob")
        .get()
        .then(async (blob: Blob, error, callback) => {
          if (error) {
            return " ";
          }
          return await this.blobToB64(blob);
        });
    } catch (error) {
      return " ";
    }

  }

  public render(): React.ReactElement<IShowAllUsersProps> {

    return (
      <div className={styles.showAllUsers}>
        <div className={styles.todayBirthday}>
          <div className={styles.title}>Cumpleañeros de hoy</div>
          {this.state.users.filter(user => user.birthday >= this.props.TodayDate && user.birthday <= this.props.TodayDate).map(filteredUser => (
            <>
              <div className={styles.birthdayCard}>
                <div className={styles.birthdayBackground}>
                  <div className={styles.background1}>
                    <img src={require('../imgs/today1.png')} alt="Error" />
                  </div>
                  <div className={styles.background2}>
                    <img src={require('../imgs/today2.png')} alt="Error" />
                  </div>
                </div>
                <div className={styles.birthdayCardProfileImg}>
                  <img src={filteredUser.profilePhoto} className={styles.profilePhoto} />
                </div>
                <div className={styles.birthdayContent}>
                  <div className={styles.displayName}>
                    <p className={styles.name}>{filteredUser.givenName}</p>
                    <p className={styles.surname}>{filteredUser.surname}</p>
                  </div>
                  <p className={styles.jobTitle}>{filteredUser.jobTitle}</p>
                  <hr className={styles.line}></hr>
                  <p className={styles.department}>{filteredUser.department}</p>
                  <p className={styles.aboutMeTitle}>Mis gustos:</p>
                  <p className={styles.aboutMe}>{filteredUser.aboutMe}Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Gravida dictum fusce ut placerat orci. Odio eu feugiat pretium nibh ipsum consequat. Nullam ac tortor vitae purus faucibus. Mauris cursus mattis molestie a iaculis at erat pellentesque adipiscing. Quis varius quam quisque id diam vel. Lectus nulla at volutpat diam ut venenatis tellus in. Leo urna molestie at elementum eu facilisis sed odio morbi. Ut tristique et egestas quis. Congue nisi vitae suscipit tellus mauris a diam maecenas sed. Volutpat commodo sed egestas egestas fringilla. Dapibus ultrices in iaculis nunc sed augue lacus. Amet risus nullam eget felis eget. Dignissim sodales ut eu sem. Ut ornare lectus sit amet est placerat in egestas. Tristique magna sit amet purus gravida quis blandit turpis cursus. Orci dapibus ultrices in iaculis nunc sed. Enim ut tellus elementum sagittis vitae et leo duis ut.</p>
                </div>
              </div>
            </>
          ))}
        </div>

        <div className={styles.weekBirthday}>
          <h1 className={styles.title}>Cumpleañeros de la semana</h1>
          {this.state.users.filter(user => user.birthday > this.props.TodayDate && user.birthday <= this.props.WeekDate).sort((a, b) => { return a.birthday > b.birthday ? 1 : a.birthday < b.birthday ? -1 : 0; }).map(filteredUser => (
            <>
              <div className={styles.birthdayCard}>
                <div className={styles.birthdayImg}>
                  <img src={require('../imgs/week.png')} className={styles.backgroundImg} alt="Error" />
                </div>
                <div className={styles.birthdayCardProfileImg}>
                  <img src={filteredUser.profilePhoto} className={styles.profilePhoto} />
                </div>
                <div className={styles.birthdayContent}>
                  <div className={styles.displayName}>
                    <p className={styles.name}>{filteredUser.givenName}</p>
                    <p className={styles.surname}>{filteredUser.surname}</p>
                  </div>
                  <p className={styles.jobTitle}>{filteredUser.jobTitle}</p>
                  <hr className={styles.line}></hr>
                  <p className={styles.department}>{this._formatDate(filteredUser.birthday)}</p>
                  <p className={styles.department}>{filteredUser.department}</p>
                  <p className={styles.aboutMeTitle}>Mis gustos:</p>
                  <p className={styles.aboutMe}>{filteredUser.aboutMe}Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Gravida dictum fusce ut placerat orci. Odio eu feugiat pretium nibh ipsum consequat. Nullam ac tortor vitae purus faucibus. Mauris cursus mattis molestie a iaculis at erat pellentesque adipiscing. Quis varius quam quisque id diam vel. Lectus nulla at volutpat diam ut venenatis tellus in. Leo urna molestie at elementum eu facilisis sed odio morbi. Ut tristique et egestas quis. Congue nisi vitae suscipit tellus mauris a diam maecenas sed. Volutpat commodo sed egestas egestas fringilla. Dapibus ultrices in iaculis nunc sed augue lacus. Amet risus nullam eget felis eget. Dignissim sodales ut eu sem. Ut ornare lectus sit amet est placerat in egestas. Tristique magna sit amet purus gravida quis blandit turpis cursus. Orci dapibus ultrices in iaculis nunc sed. Enim ut tellus elementum sagittis vitae et leo duis ut.</p>
                </div>
              </div>
            </>
          ))}
        </div>

        <div className={styles.monthBirthday}>
          <h1 className={styles.title}>Cumpleañeros del mes</h1>
          {this.state.users.filter(user => user.birthday > this.props.WeekDate && user.birthday <= this.props.MonthDate).sort((a, b) => { return a.birthday > b.birthday ? 1 : a.birthday < b.birthday ? -1 : 0; }).map(filteredUser => (
            <>
              <div className={styles.birthdayCard}>
                <div className={styles.birthdayImg}>
                  <img src={require('../imgs/month.png')} className={styles.backgroundImg} alt="Error" />
                </div>
                <div className={styles.birthdayContent}>
                  <p className={styles.displayName}>{filteredUser.displayName}</p>
                  <p className={styles.jobTitle}>{filteredUser.jobTitle}</p>
                  <hr className={styles.line}></hr>
                  <p className={styles.department}>{this._formatDate(filteredUser.birthday)}</p>
                  <p className={styles.department}>{filteredUser.department}</p>
                  <p className={styles.aboutMeTitle}>Mis gustos:</p>
                  <p className={styles.aboutMe}>{filteredUser.aboutMe}Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Gravida dictum fusce ut placerat orci. Odio eu feugiat pretium nibh ipsum consequat. Nullam ac tortor vitae purus faucibus. Mauris cursus mattis molestie a iaculis at erat pellentesque adipiscing. Quis varius quam quisque id diam vel. Lectus nulla at volutpat diam ut venenatis tellus in. Leo urna molestie at elementum eu facilisis sed odio morbi. Ut tristique et egestas quis. Congue nisi vitae suscipit tellus mauris a diam maecenas sed. Volutpat commodo sed egestas egestas fringilla. Dapibus ultrices in iaculis nunc sed augue lacus. Amet risus nullam eget felis eget. Dignissim sodales ut eu sem. Ut ornare lectus sit amet est placerat in egestas. Tristique magna sit amet purus gravida quis blandit turpis cursus. Orci dapibus ultrices in iaculis nunc sed. Enim ut tellus elementum sagittis vitae et leo duis ut.</p>
                </div>
              </div>
            </>
          ))}
        </div>
      </div>
    );
  }
}
