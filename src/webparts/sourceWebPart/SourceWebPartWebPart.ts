import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SourceWebPartWebPartStrings';
import SourceWebPart from './components/SourceWebPart';
import { ISourceWebPartProps } from './components/ISourceWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';

import { RxJsEventEmitter } from '../../libraries/rxJsEventEmitter/RxJsEventEmitter';
import { EventData } from '../../libraries/rxJsEventEmitter/EventData';

export interface ISourceWebPartWebPartProps {
  description: string;
  listTitle: string;
  field: string;
  listItems: any;
  fieldOptions: any;
  selectedItem: any;
  htmlCode: string;
}

export default class SourceWebPartWebPart extends BaseClientSideWebPart<ISourceWebPartWebPartProps> {

  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  private _defaultHtml = 
`<style type="text/css">
.itemsStyle{
  list-style: none;
  padding: 8px;
  font-size: large;
  font-weight: 400;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif
}

.listStyle{
  list-style: none;
  padding: 0px;
}

.itemStyleActive{
  list-style: none;
  padding: 8px;
  color: white;
  background: cadetblue;
  font-size: large;
  font-weight: bold;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif
}
</style>

<div>
  <ul class="listStyle">
    {{#each items}}
      <li id="{{../webpartId}}{{this.Id}}" class="itemsStyle">{{#showSelectedField this}}{{/showSelectedField}}</li>
    {{/each}}
  </ul>
</div>`;

  public render(): void {
    const element: React.ReactElement<ISourceWebPartProps > = React.createElement(
      SourceWebPart,
      {
        description: this.properties.description,
        fieldComplete: this.properties.listTitle && this.properties.field ? true: false,
        items: this.properties.listItems,
        selectedField: this.properties.field,
        setSelectedItem: this._setSelectedItem,
        listTitle: this.properties.listTitle,
        selectedItem: this.properties.selectedItem,
        htmlCode: this.properties.htmlCode? this.properties.htmlCode : this._defaultHtml
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getSiteLists = async (): Promise<any> => {
    return await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      '/_api/web/lists?$filter=Hidden eq false and BaseType eq 0', SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      });
  }

  private _mapLists = async () => {
    //Get site lists
    const lists = await this._getSiteLists();
    this._listOptions = lists.value.map(list => {
      return {
        key: list.Title,
        text: list.Title
      };
    });
  }

  private _getListFields = async (listTitle: string): Promise<any> => {
    return await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      `/_api/web/lists/GetByTitle('${listTitle}')/Fields?$filter=Hidden eq false and ReadOnlyField eq false`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      });
  }
  
  private _mapFields = async () => {
    //Get list fields
    const fields = await this._getListFields(this.properties.listTitle);
    this.properties.fieldOptions = fields.value.map(field => {
      return {
        key: field.Title,
        text: field.Title
      };
    });
    this._fieldOptionsDisabled = false;

    //to send to the receiver webpart
    this._sendFields();
  }

  private _getListItems = async (listTitle: string): Promise<any> => {
    return await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('${listTitle}')/Items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      });
  }

  private _mapListItems = async () => {
    const listItems = await this._getListItems(this.properties.listTitle);
    this.properties.listItems = listItems.value.map(item => {
      return item;
    });
  }

  private _sendFields = () => {
    this._eventEmitter.emit("recieveFields", { fields: this.properties.fieldOptions } as EventData);
  }

  private _sendSelectedItem = () => {
    if(this.properties.selectedItem)
      this._eventEmitter.emit("receiveSelectedItem", { selectedItem: this.properties.selectedItem } as EventData);
  }

  private _setSelectedItem = (item: any) => {
    this.properties.selectedItem = item;
  }

  private _listOptions: IPropertyPaneDropdownOption[] = [];
  
  private _fieldOptionsDisabled : boolean = true;

  protected onInit(): Promise<void>{
    return super.onInit().then(_ => {
      this._eventEmitter.on("receiverWebpartStarted", this._sendFields.bind(this));
      this._eventEmitter.on("receiverWebpartMounted", this._sendSelectedItem.bind(this));
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneConfigurationStart(){
    
    await this._mapLists();
    if(this.properties.field){
      await this._mapFields()
      .then(async ()=>{
        await this._mapListItems();
      });
      this._fieldOptionsDisabled = false;
    }
    
    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(){

    await this._mapFields()
    .then(async()=>{
      await this._mapListItems();
    });

    this.context.propertyPane.refresh();
  }

  private _onTemplateFieldChanged(){

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: '',
              groupFields: [
                PropertyPaneDropdown('listTitle',{
                  label: 'Lists',
                  options: this._listOptions,
                }),
                PropertyPaneDropdown('field',{
                  label: 'Fields',
                  options: this.properties.fieldOptions,
                  disabled: this._fieldOptionsDisabled
                }),
                PropertyFieldCodeEditor('htmlCode', {
                  label: 'Edit Template',
                  panelTitle: 'Edit Template',
                  initialValue: this.properties.htmlCode? this.properties.htmlCode: this._defaultHtml,
                  onPropertyChange: this._onTemplateFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.HTML
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
