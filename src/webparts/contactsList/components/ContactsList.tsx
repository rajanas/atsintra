import * as React from 'react';
import styles from './ContactsList.module.scss';
import { IContactsListProps } from './IContactsListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPersonaProps, IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { sp } from '@pnp/sp';
import ListService from '../../../shared/services/ListService';
const examplePersona: IPersonaSharedProps = {
  imageUrl: "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=suresh@aureustechsystems.com&UA=0&size=HR96x96&sc=1551151046966",
  imageInitials: 'AL',
  text: 'Annie Lindqvist',
  secondaryText: 'Software Engineer',
  tertiaryText: 'In a meeting',
  optionalText: 'Available at 4:00pm'
};
export default class ContactsList extends React.Component<IContactsListProps, any> {
  constructor(props) {
    super(props);
    this.state = {
      items: []
    };

  }
  public componentDidMount() {
    let listName = 'Directory';
    let selectedFields: string[] = ['Contact/ID', 'Contact/EMail', 'Contact/Department','Contact/FirstName',
    'Contact/LastName','Contact/WorkPhone'];
    let expandFields = ["Contact"];
    let top = 10;
console.log(sp.site.getWebUrlFromPageUrl(window.location.href));

    let listService = new ListService('/');
    listService.getListItemsByQuery(listName, selectedFields, expandFields, top).
      then(resp => {
        this.setState(
          {
            items: resp
          }
        );
        console.log(resp);
      });

  }
  private _onRenderTertiaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Icon iconName={'Phone'} className={'ms-JobIconExample'} />
        {props.tertiaryText}
      </div>
    );
  }
  private _onRenderSecondaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Icon iconName={'MailSolid'}  className={'ms-JobIconExample'} />
        {props.secondaryText}
      </div>
    );
  }
  public render(): React.ReactElement<IContactsListProps> {
    let items = this.state.items.map(item => {
      let personprops: IPersonaSharedProps = {
        imageUrl: `/_layouts/15/userphoto.aspx?size=L&accountname=${item.Contact.EMail}`,

        text: item.Contact.FirstName + ' ' + item.Contact.LastName,
        secondaryText: item.Contact.EMail,
        tertiaryText: item.Contact.WorkPhone

      };
      return (<Persona
        {...personprops}
        size={PersonaSize.size72}
        onRenderSecondaryText={this._onRenderSecondaryText}
        onRenderTertiaryText={this._onRenderTertiaryText}
      />);
    });
    return (
      <div className={styles.contactsList}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className="ms-PersonaExample">
              {
                items
              }

              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
