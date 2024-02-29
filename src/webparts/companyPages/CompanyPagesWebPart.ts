import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CompanyPagesWebPart.module.scss';
import * as strings from 'CompanyPagesWebPartStrings';
import iconPageImage from './assets/icon-page.png';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ICompanyPagesWebPartProps {
  description: string;
  selectedList: string; 
  Title: string;
  URL: string;
  Company: string;
}

export default class CompanyPagesWebPart extends BaseClientSideWebPart<ICompanyPagesWebPartProps> {

  private availableLists: IPropertyPaneDropdownOption[] = [];

  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value.toLowerCase();
        console.log('User Email using User Principal Name:', this.userEmail);
        // Now you can use this.userEmail as needed
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  }

  public render(): void {
    // Use async/await to wait for userDetails to complete before rendering
    this.userDetails().then(() => {
      const decodedDescription = decodeURIComponent(this.properties.description);
      console.log(decodedDescription);
      this.domElement.innerHTML = `
        <div class="${styles['group-information']}">
          <div class="${styles['parent-div']}">
            <div>
              <h2>${decodedDescription}</h2>
              <div id="buttonsContainer"></div>
            </div>
          </div>
        </div>`;
  
      this._renderButtons();
    });
  }

  private _renderButtons(): void {
    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#buttonsContainer');
    buttonsContainer?.classList.add(styles.buttonsContainer);
    
    const adminEmailSplit: string[] = this.userEmail.split('.admin@');
    if (this.userEmail.includes(".admin@")){
        console.log("Admin Email after split: ", adminEmailSplit);
    }
    let otherUsersSplit = "";
    if (this.userEmail.includes("_") && this.userEmail.includes("#ext#")){
    const parts = this.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    otherUsersSplit =  secondPart.split('.com')[0];
      console.log("User's company after split: ", otherUsersSplit);
    }

    const siteUrl = this.context.pageContext.web.absoluteUrl;
    console.log("siteUrl: ", siteUrl);

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items?$orderby=Created%20desc&$top=5000`;

    fetch(apiUrl, {
        method: 'GET',
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': ''
        }
    })
    .then(response => response.json())
    .then(data => {
        console.log("Api response: ", data);
        let buttonsCreated = 0; // Variable to keep track of the number of buttons created
        if (data.value && data.value.length > 0) {
            data.value.forEach((item: ICompanyPagesWebPartProps) => {
                if (buttonsCreated >= 6) {
                    console.log("Maximum number of buttons created, loop is exited");
                    return; // Exit the loop if the maximum number of buttons is reached
                }

                if((this.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.userEmail.includes(".admin@") && !otherUsersSplit) || (this.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){                  
                    console.log("Creating button for ", item.Title);
                    const buttonDiv: HTMLDivElement = document.createElement('div');
                    buttonDiv.classList.add(styles['content-box']); // Apply the 'content-box' class styles from YourStyles.module.scss
                    console.log("item.URL.includes(siteUrl)",item.URL.includes(siteUrl));
                    buttonDiv.onclick = (event) => {
                      if(item.URL.includes(siteUrl)){
                        event.preventDefault(); // Prevent the default behavior of the click event
                        window.location.href = item.URL; // Navigate to the 'URL' from the API response in the same tab
                      }else{
                        window.open(item.URL, '_blank'); // Open the 'Url' from the API response in a new tab
                      }
                    };

                    const imgContainer: HTMLDivElement = document.createElement('div');
                    imgContainer.classList.add(styles['content-box-img-container']); // Apply the image container styles
                    const img: HTMLImageElement = document.createElement('img');
                    img.src = iconPageImage; // Use the imported image URL
                    imgContainer.appendChild(img); // Append the image to the container
                    buttonDiv.appendChild(imgContainer); // Append the image container to the button

                    const titleSpan: HTMLDivElement = document.createElement('div');
                    titleSpan.classList.add(styles['content-box-text-container']);
                    titleSpan.textContent = item.Title; // Use the 'Title' from the API response
                    buttonDiv.appendChild(titleSpan); // Append the title to the button

                    buttonsContainer!.appendChild(buttonDiv); // Append the button to the buttons container
                    buttonsCreated++; // Increment the count of buttons created
                } else {
                    console.log("No button creation for: ", item.Title);
                }
            });
        } else {
            const noDataMessage: HTMLDivElement = document.createElement('div');
            noDataMessage.textContent = 'No applications available for the user.';
            buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
        }
    })
    .catch(error => {
        console.error("Error fetching user data: ", error);
    });
}


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList') {
      this.setListTitle(newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _loadLists(): void {
    const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
    //SPHttpClient is a class provided by Microsoft that allows developers to perform HTTP requests to SharePoint REST APIs or other endpoints within SharePoint or the host environment. It is used for making asynchronous network requests to SharePoint or other APIs in SharePoint Framework web parts, extensions, or other components.
    this.context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1)
    //SPHttpClientResponse is the response object returned after making a request using SPHttpClient. It contains information about the response, such as status code, headers, and the response body.
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this.availableLists = data.value.map((list) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists:', error);
      });
  }

  private setListTitle(selectedList: string): void {
    this.properties.selectedList = selectedList;

    this.context.propertyPane.refresh();
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
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title For The Application"
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select A List',
                  options: this.availableLists,
                }),
              ],
            },
          ],
        }
      ]
    };
  }
}
