import * as React from 'react';
import styles from './SourceWebPart.module.scss';
import { ISourceWebPartProps } from './ISourceWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { RxJsEventEmitter } from '../../../libraries/rxJsEventEmitter/RxJsEventEmitter';
import { EventData } from '../../../libraries/rxJsEventEmitter/EventData';
import * as Handlebars from 'handlebars';
import * as $ from 'jquery';

export default class SourceWebPart extends React.Component<ISourceWebPartProps, {selectedItem: any}> {

  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  private _selectedId: string = null;


  public componentDidMount(){
    if(this.props.selectedItem){
      this._selectedId = this.props.selectedItem['Id'];
      $('#'+this.props.selectedItem['Id']).addClass('itemStyleActive');
    }

    $(".itemsStyle").on("click", (e)=>{
      if(this._selectedId)
        $('#'+this._selectedId).removeClass('itemStyleActive');

      let item = this.props.items.filter(
        item => item['Id'] == e.currentTarget.id)[0];
      this._itemClicked(item);

      $('#'+item['Id']).addClass('itemStyleActive');
      this._selectedId = item['Id'];
    });
  }

  constructor(props){
    super(props);

    Handlebars.registerHelper('showSelectedField', (item) => {
      return item[this.props.selectedField];
    });
    
  }

  public render(): React.ReactElement<ISourceWebPartProps> {

    return (
        <div>
          {
            this.props.fieldComplete ?
              <div className={styles.header}>
                <span className={styles.headerTitle}>{this.props.listTitle}</span>
              </div>
            :
              <div></div>
          }

          {
            this.props.fieldComplete ?
            <div dangerouslySetInnerHTML={this._createMarkup(this.props.htmlCode)}></div>
            :
            <div>Please configure the webpart properly</div>
          }

      </div>
    );
  }

  private _createMarkup = (html: string) => {

    let context = {
      items: this.props.items,
      selectedItem: this.props.selectedItem
    };

    let template = Handlebars.compile(html);
    let html2 = template(context);

    return {__html: html2};
  }

  private _itemClicked = (item: any) => {
    this.props.setSelectedItem(item);
    this.setState({
      selectedItem: item
    });
    //to send to the receiver webpart
    this._eventEmitter.emit("receiveSelectedItem", { selectedItem: item } as EventData);
  }
}
