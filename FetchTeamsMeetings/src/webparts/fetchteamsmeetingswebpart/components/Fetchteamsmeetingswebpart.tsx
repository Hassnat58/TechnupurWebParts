import * as React from 'react';
import styles from './Fetchteamsmeetingswebpart.module.scss';
import type { IFetchteamsmeetingswebpartProps } from './IFetchteamsmeetingswebpartProps';
import { DetailsList, TextField, DetailsListLayoutMode, IColumn } from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';


//Interface
export interface IUser {
  displayName: string;
  jobTitle: string;
  department: string;
  mobilePhone: string;
  mail: string;
  officeLocation: string;
  extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute : string;
  extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet : string;

}

export interface IUserState {
  userState: IUser[];
  filteredUsers: IUser[];
}

export default class StaffDirectory extends React.Component<IFetchteamsmeetingswebpartProps, IUserState> {
  constructor(props: IFetchteamsmeetingswebpartProps) {
    super(props);
    this.state = { userState: [], filteredUsers: [] };
  }

  public allusers: IUser[] = [];

  
  //GetUsers Function to Fetch Employees
  public GetUsers = (): void => {
    this.props.context.msGraphClientFactory
      .getClient('3')
      .then((msGraphClient: MSGraphClientV3): void => {
        msGraphClient.api("users").version("v1.0").select("id,displayName,mail,extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute,extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet").get((err, res: any) => {
            console.table("Output", res);
          res.value.map((result: any, index: number) => {
            this.allusers.push({
              displayName: result.displayName,
              jobTitle: result.jobTitle, 
              department: result.department, 
              mobilePhone: result.mobilePhone, 
              mail: result.mail,
              officeLocation: result.officeLocation,
              extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute : result.extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute,
              extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet : result.extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet



             
            });
          });
          this.setState({ userState: this.allusers, filteredUsers: this.allusers });
        });
      });
  };

  componentDidMount() {
    this.GetUsers();
  }



  //Search Functionality
  public handleSearch = (event: React.ChangeEvent<HTMLInputElement>): void => {
  
    const value = event.target.value; // Correctly access the value property
    const filteredUsers = this.state.userState.filter((user) => {
      const displayName = user.displayName || '';
      const mail = user.mail || '';
      return displayName.toLowerCase().includes(value.toLowerCase()) || mail.toLowerCase().includes(value.toLowerCase());
    });
    this.setState({ filteredUsers });
  };

  
  //Render Function
  public render(): React.ReactElement<IFetchteamsmeetingswebpartProps> {
    const columns: IColumn[] = [
      { key: 'name', name: 'Name', fieldName: 'displayName', minWidth: 150, isResizable: true },
      { key: 'jobTitle', name: 'Designation', fieldName: 'jobTitle', minWidth: 150, isResizable: true },
      { key: 'department', name: 'Department', fieldName: 'department', minWidth: 150, isResizable: true },
      { key: 'mobilePhone', name: 'Phone No.', fieldName: 'mobilePhone', minWidth: 120, isResizable: true },
      { key: 'email', name: 'Email', fieldName: 'mail', minWidth: 200, isResizable: true },
      { key: 'officeLocation', name: 'Location', fieldName: 'officeLocation', minWidth: 150, isResizable: true },
      { key: 'extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute', name: 'customAttribute', fieldName: 'extension_7e7e6b63c6af495aa165f2ed82585ad9_customAttribute', minWidth: 150, isResizable: true },
      { key: 'extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet', name: 'Intranet', fieldName: 'extension_7e7e6b63c6af495aa165f2ed82585ad9_Intranet', minWidth: 150, isResizable: true },
      
    ];

    return (
      <div className={styles.staffDirectoryContainer}>
        <h2 className={styles.title}>
          <span className={styles.icon}>👥</span> Staff Directory
        </h2>
        <TextField
  placeholder="Search"
  onChange={this.handleSearch}
  className={styles.searchInput}
  iconProps={{ iconName: 'Search' }}
/>
        <DetailsList
          items={this.state.filteredUsers}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionPreservedOnEmptyClick={true}
        />
      </div>
    );
  }
}
