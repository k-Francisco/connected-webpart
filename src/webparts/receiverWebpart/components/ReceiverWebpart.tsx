import * as React from 'react';
import styles from './ReceiverWebpart.module.scss';
import { IReceiverWebpartProps } from './IReceiverWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as Handlebars from 'handlebars';
import * as $ from 'jquery';

import { RxJsEventEmitter } from '../../../libraries/rxJsEventEmitter/RxJsEventEmitter';
import { EventData } from '../../../libraries/rxJsEventEmitter/EventData';

const WEBPART_ID: String = "417117e7-a6e7-4296-a870-7f744eb7fac1";

export default class ReceiverWebpart extends React.Component<IReceiverWebpartProps, {
  selectedItem: any,
  listItems: any,
  stateContext: any}> {

  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  public componentDidMount(){
    this._eventEmitter.emit("receiverWebpartMounted", { });
    
    $(document).on("click", "#saveForm", e=>{
      
      e.currentTarget.closest('div').childNodes.forEach(element => {
          if(element.tagName == "FORM"){
            this._updateItem(e.currentTarget.closest('div').id.replace(WEBPART_ID, ""),
            $(element).serializeArray());
          }
      });

    });

    $(document).on("click", ".edit",e => {

      if(e.currentTarget.innerText.toLowerCase() == "edit")
        e.currentTarget.innerText= "Cancel"
      else
        e.currentTarget.innerText= "Edit"
      
      if(e.currentTarget.closest('div').classList.contains("hidden")){
        e.currentTarget.closest('div').classList.remove("hidden");
        e.currentTarget.closest('div').classList.add("visible");
      }
      else{
        e.currentTarget.closest('div').classList.remove("visible");
        e.currentTarget.closest('div').classList.add("hidden");
      }
    });
  }

  public componentWillReceiveProps(nextProps){
    this.setState({
      stateContext: {
        listItems: this.state.listItems,
        fields: nextProps.dropDownSelect,
        webpartId: WEBPART_ID
      }
    });
  }

  constructor(props){
    super(props);

    this.state = {
      selectedItem: null,
      listItems: null,
      stateContext: null
    };

    Handlebars.registerHelper('showField', (item, field) => {
      return item[field];
    });

    this._eventEmitter.on("receiveSelectedItem", this.receiveSelectedItem.bind(this));
  }

  public render(): React.ReactElement<IReceiverWebpartProps> {
    return (
      <div>
        {
          this.state.selectedItem && this.state.listItems ? 
          <div className={styles.header}>
            <span className={styles.headerTitle}>{this.props.listTitle}</span>
          </div>
          :
          <div></div>
        }

        {
          this.state.selectedItem ? 
            this.state.listItems ? 
              <div dangerouslySetInnerHTML={this._createMarkup(this.props.htmlCode)}></div>
              :
              <div>Please select fields to display</div>
          :
          <div>Please select an Item</div>
        }
      </div>
    );
  }

  protected async receiveSelectedItem(data: EventData) {
    this.setState({
      selectedItem: data.selectedItem
    });

    const listItems = await this._getListItems(this.props.listTitle);
    if(listItems){
      this.setState({
        listItems: listItems.value.filter(item => item[this.props.lookupColumn+'Id'] == this.state.selectedItem.Id)
      }, ()=>{
        this.setState({
          stateContext: {
            listItems: this.state.listItems,
            fields: this.props.dropDownSelect,
            webpartId: WEBPART_ID
          }
        });
      });
    }
  }

  private _createMarkup = (html: string) => {

    let template = Handlebars.compile(html);
    let html2 = template(this.state.stateContext);

    return {__html: html2};
  }

  private _getListItems = async (listTitle: string): Promise<any> => {
    return await this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('${listTitle}')/Items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      });
  }

  private _updateItem = async (itemId: string, fields: any) => {
    
    const listItemEntityTypeName = await this._getListItemEntityTypeName();
    let itemBody: string = `{"__metadata":{"type":"${listItemEntityTypeName.value}"},`;

    fields.forEach((field,index) => {
      itemBody += `"${field.name}":"${field.value}"`
      
      if(index != fields.length-1)
        itemBody+=`,`
      else
        itemBody+=`}`
    });

    console.log(itemBody);

    await this.props.context.spHttpClient.post(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': "*",
          'X-HTTP-Method': 'MERGE'
        },
        body: itemBody
      }
    )
    .then(async(response) =>{
      const listItems = await this._getListItems(this.props.listTitle);
      if(listItems){
          this.setState({
            listItems: listItems.value.filter(item => item[this.props.lookupColumn+'Id'] == this.state.selectedItem.Id)
          }, ()=>{
            this.setState({
              stateContext: {
                listItems: this.state.listItems,
                fields: this.props.dropDownSelect,
                webpartId: WEBPART_ID
              }
            });
          });
        }
    });
  }

  private _getListItemEntityTypeName = async (): Promise<any> => {
    return await this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/ListItemEntityTypeFullName`,
    SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    });
  }

}
