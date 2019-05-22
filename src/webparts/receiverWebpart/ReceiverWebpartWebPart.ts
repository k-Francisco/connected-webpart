import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';
import {PropertyFieldMultiSelect} from '@pnp/spfx-property-controls';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'ReceiverWebpartWebPartStrings';
import ReceiverWebpart from './components/ReceiverWebpart';
import { IReceiverWebpartProps } from './components/IReceiverWebpartProps';

export interface IReceiverWebpartWebPartProps {
  description: string;
  dropDownSelect: string[];
  listTitle: string;
  lookupColumn: string;
  htmlCode: string;
}

export default class ReceiverWebpartWebPart extends BaseClientSideWebPart<IReceiverWebpartWebPartProps> {

  private _defaultHtml: string = 
`<style type="text/css">
.listStyleReceiver{
  list-style: none;
  padding: 0px;
  flex-grow: 1;
}

.divStyle{
  display: flex;
}

.divStyle .edit{
  display:none;
}

.divStyle:hover .edit{
  display: flex;
}

.edit{
  align-self:center;
  padding:8px;
  cursor: pointer;
}

.edit:hover{
  color: #017E9D;
}

.editForm{
  display:flex;
  padding: 8px;
}

.save{
  padding:8px;
  cursor: pointer;
}

.save:hover{
  color: #017E9D;
}

.visible>div{
  display:block;
}

.visible>ul{
  display:none;
}

.hidden>div{
  display:none;
}

.hidden>ul{
  display:block;
}

</style>

{{#each listItems}}
  <div class="divStyle hidden">
    <ul class="listStyleReceiver">
    {{#each ../fields}}
    <li>{{this}} : {{#showField ../this this }}{{/showField}}</li>
    {{/each}}
    </ul>

    <div id="{{../webpartId}}{{this.Id}}" class="editForm">
      <form>
        {{#each ../fields}}
          {{this}}: <input type="text" name="{{this}}" placeholder="{{#showField ../this this }}{{/showField}}"/>
          <br/>
        {{/each}}
      </form>
      <button id="saveForm">Save</button>
    </div>
    <span class="edit">Edit</span>
  </div>
{{/each}}

`;

  public render(): void {
    const element: React.ReactElement<IReceiverWebpartProps > = React.createElement(
      ReceiverWebpart,
      {
        description: this.properties.description,
        dropDownSelect: this.properties.dropDownSelect,
        context: this.context,
        listTitle: this.properties.listTitle,
        lookupColumn: this.properties.lookupColumn,
        htmlCode: this.properties.htmlCode ? this.properties.htmlCode : this._defaultHtml
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onInit(): Promise<void>{
    return super.onInit().then(_ => {

    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
    this._listFields = fields.value.map(field => {
      return {
        key: field.Title,
        text: field.Title
      };
    });
    this._fieldOptionsDisabled = false;
  }

  private _listFields: IPropertyPaneDropdownOption[] = [];
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldOptionsDisabled : boolean = true;

  protected async onPropertyPaneConfigurationStart(){
    await this._mapLists();
    if(this.properties.dropDownSelect){
      await this._mapFields();
    }

    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(){
    await this._mapFields();
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
                options: this._listOptions
              }),
              PropertyPaneDropdown('lookupColumn',{
                label: 'Lookup Column',
                options: this._listFields,
                disabled: this._fieldOptionsDisabled
              }),
              PropertyFieldMultiSelect('dropDownSelect',{
                label: 'Select fields to display',
                options: this._listFields,
                selectedKeys: this.properties.dropDownSelect,
                key: 'dropDownSelectFieldId'
              }),
              PropertyFieldCodeEditor('htmlCode', {
                label: 'Edit Template',
                panelTitle: 'Edit Template',
                initialValue: this.properties.htmlCode ? this.properties.htmlCode : this._defaultHtml,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
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
