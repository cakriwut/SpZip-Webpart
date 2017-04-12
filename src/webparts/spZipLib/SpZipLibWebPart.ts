import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpZipLib.module.scss';
import * as strings from 'spZipLibStrings';
import { ISpZipLibWebPartProps } from './ISpZipLibWebPartProps';

//--
import { ISPListItem, ISPListItems } from './SpOperation';
import MockHttpClient from './MockHttpClient';
import {
  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import * as jQuery from 'jquery';
//--

export default class SpZipLibWebPart extends BaseClientSideWebPart<ISpZipLibWebPartProps> {

   private _getMockListData() : Promise<ISPListItems> {    
      return MockHttpClient.get()
        .then((data: ISPListItem[]) => {
          var listItemData : ISPListItems = { value: data};
          return listItemData;
        }) as Promise<ISPListItems>;    
  }

  private _getListData() : Promise<ISPListItems>{
  
    var restUrl : string = this.context.pageContext.web.absoluteUrl;
    restUrl += "/_api/web/lists/GetByTitle('";
    restUrl += this.properties.library;
    restUrl += "')/items?$expand=File&$select=Title,id,File,FileSystemObjectType";


    return this.context.spHttpClient.get(restUrl,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {      
      return response.json().then((responseFormatted: any)=>{
         var formattedResponse: ISPListItems = { value:[]};
         responseFormatted.value.map((object:any, i:number) =>{
            if(object['FileSystemObjectType'] =='0'){
              var spListItem : ISPListItem = {
                'Id': object["ID"],
                 'Title': object["Title"],
                 'Name':  object["File"]["Name"],
                 'ServerRelativeUrl': object["File"]["ServerRelativeUrl"]
              }
              formattedResponse.value.push(spListItem);
            }
         });
         
         return formattedResponse;
      });
    });
  }
  
  private _renderList(items: ISPListItem[]): void{
     let html: string = '';
     items.forEach((item: ISPListItem) => {
        html += `
          <ul>
            <li>
              <a href="${item.ServerRelativeUrl}">${item.Name}</a>   
                <button class="download-Button" data-itemid="${item.Id}">
                  <span class="${styles.label}">Download zip</span>
                </button>
            </li>
          </ul>
        `;
     }); 
     
     const listContainer : Element = this.domElement.querySelector("#listContainer");
     listContainer.innerHTML = html;
     this._setButtonEventHandler();
  }

  private _setButtonEventHandler():void{
    const webPart: SpZipLibWebPart = this;
    jQuery("#listContainer",this.domElement).on("click","button.download-Button",(source) =>  { webPart.downloadZip(source); });
  }

  private downloadZip(source: JQueryEventObject): void{
    var itemID:string = jQuery(source.currentTarget).data("itemid");
    var azureUrl : string = this.properties.azureFunction;
    azureUrl += "&siteUrl=" + this.context.pageContext.web.absoluteUrl;
    azureUrl += "&listTitle=" + this.properties.library;
    azureUrl += "&itemId=" + itemID;

    window.location.href = azureUrl; 
  }

  private _renderListAsync() : void{
    if(Environment.type === EnvironmentType.Local ) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    } else if(
       Environment.type === EnvironmentType.SharePoint ||
       Environment.type === EnvironmentType.ClassicSharePoint
    ){
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
            </div>
          </div>
          <div class="ms-Grid-row>
            <div class="ms-Grid-col ms-u-lg12>
              <div id='listContainer'></div>
            </div>
          </div>
        </div>
      </div>`;

     this._renderListAsync();     
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
                PropertyPaneTextField('library', {
                      label: "Document Library"
                  }),
                PropertyPaneTextField('azureFunction', {
                      label: "Azure Function URL"
                  })
              ]
            }
          ]
        }
      ]
    };
  }
}