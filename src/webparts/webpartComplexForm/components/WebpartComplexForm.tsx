import * as React from 'react';
import styles from './WebpartComplexForm.module.scss';
import { IWebpartComplexFormProps } from './IWebpartComplexFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import Editor from 'react-medium-editor';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { ISPListData } from '@microsoft/sp-page-context/lib/SPList';
import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';

import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import {
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';
import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse, SPHttpClientConfiguration
} from '@microsoft/sp-http';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';


import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { Promise } from 'es6-promise';
import * as lodash from 'lodash';
import * as jquery from 'jquery';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import ReactFileReader from 'react-file-reader';

import { Checkbox, CheckboxGroup } from 'react-checkbox-group';

import { default as pnp, ItemAddResult, Web } from "sp-pnp-js";

require('medium-editor/dist/css/medium-editor.css');
require('medium-editor/dist/css/themes/default.css');


var App = React.createClass({

  getInitialState() {
    return { text: 'Enter Rich Text Description' };
  },

  render() {
    var divStyle = {
      background: "#eee",
      padding: "10px",
      margin: "1px",
      width: "100%",
      height: "140px",
    };

    return (
      <div style={divStyle}>
        <Editor
          text={this.state.text}
          onChange={this.props.handleChange}
        />
      </div>
    );
  },
  handleChange(text, medium) {
    this.setState({ text: text });
  }
});


export default class WebpartComplexForm extends React.Component<IWebpartComplexFormProps, {}> {

  public state: IWebpartComplexFormProps;
  constructor(props, context) {
    super(props);
    this.state = {
      spHttpClient: this.props.spHttpClient,
      description: "",
      ProjectName: "Select Project",
      ProjectsArray: [],
      siteurl: this.props.siteurl,
      Building: "",
      Floor: "",
      GridLine: "",
      Subject: "",
      SubContractor: "",
      CreatedDate: "01-01-1900",
      RequiredDate: "01-01-1900",
      Disciplined: "Civil",
      Description: "",
      Requirement: "",
      Comments: "",
      ItemGuid: this.GenerateGuid().toString(),
      loading: false,
      UploadedFilesArray: [],
      CurrentUser: "",
      UserGroup: "",
      IRFINumber: "",
      IRFISeriesId: "",
      IRFIReference: "",
    };
    this.onChangeDeleteDocument = this.onChangeDeleteDocument.bind(this);
  }

  private CreateNewItem(): void {
  }
  onChangeDeleteDocument(val) {
    var array = this.state.UploadedFilesArray;
    var MainIndex = val.currentTarget.dataset.id.toString(); // Let's say it's Bob.
    var indextoDelete = 0;

    for (var i = 0; i < array.length; i++) {
      if (array[i] != undefined) {
        var temp = array[i].toString().split('|');
        if (temp[1] == MainIndex) {
          indextoDelete = i;

        }
      }
    }
    delete array[indextoDelete];
    this.state.UploadedFilesArray = [];
    this.state.UploadedFilesArray = array;
    this.setState({ UploadedFilesArray: array });
    let getData = [];
    let str = [];
    for (let i = 0; i < this.state.UploadedFilesArray.length; i++) {
      if (this.state.UploadedFilesArray[i] != undefined) {
        var tempx = this.state.UploadedFilesArray[i].toString().split('|');
        str.push(<li key={tempx[0]} onClick={this.onChangeDeleteDocument.bind(this)} data-id={tempx[1]}> Uploaded File : {tempx[0]} - <a className={styles.MyHeadingsAnchor}>Delete </a></li>);
      }
    }
    getData.push(<ul>{str}</ul>);
    this.updateDocumentLibrary(MainIndex);


  }


  private updateDocumentLibrary(MainIndex) {

    var reactHandler = this;
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("SitePages", "");
    console.log(NewSiteUrl);

    const body: string = JSON.stringify({
      '__metadata': {
        'type': `SP.Data.MyDocsItem`
      },
      'Deleted': `Yes`
    });
    return this.props.spHttpClient.post(`${NewSiteUrl}/_api/web/lists/getbytitle('MyDocs')/items(${MainIndex})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': "*",
          'X-HTTP-Method': 'MERGE'
        },
        body: body
      }).then((response) => {
        console.log(response);
      });


  }

  GenerateGuid() {
    var date = new Date();
    var guid = date.valueOf();
    return guid;
  }

  componentDidMount() {
   this._renderListAsync();

  }

  private _renderListAsync(): void {
    var reactHandler = this;
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("SitePages", "");
    console.log(NewSiteUrl);
    jquery.ajax({
      url: `${NewSiteUrl}/_api/web/lists/getbytitle('Projects')/items?&$select=Title,ID`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        var myObject = JSON.stringify(resultData.d.results);
        this.setState({ ProjectsArray: resultData.d.results,
                        Disciplined:"Civil",
          })
      }.bind(this),
      error: function (jqXHR, textStatus, errorThrown) {
      }

    });
  }


  DisciplineChanged = (newFruits) => {
    this.setState({
      Disciplined: newFruits
    });
  }
  public onSelectDate(event: any): void {
    this.setState({ CreatedDate: event._d });
  }
  public onSelectDateRequired(event: any): void {
    this.setState({ RequiredDate: event._d });
  }

  public OnchangeSubContractor(event: any): void {
    this.setState({ SubContractor: event.target.value });
  }

  public OnchangeSubject(event: any): void {
    this.setState({ Subject: event.target.value });
  }

  public onChangeSelect(event: any): void {
    this.setState({
      ProjectName: event.target.value,
    });

  }

  private prototypes(str) {
    var size = 4;
    var s = String(str);
    while (s.length < (size || 2)) { s = "0" + s; }
    return s;
  }



  handleFiles = files => {
    var TemFileGuidName = [];
    var component = this;
    component.setState({ loading: true });
    var FileExtension = this.getFileExtension1(files.fileList[0].name);
    var date = new Date();
    var guid = date.valueOf();
    if (this.state.ItemGuid == "-1") {
      this.setState({ ItemGuid: guid });
    }
    //alert(this.state.ItemGuid);   
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    console.log(NewSiteUrl);
    let webx = new Web(NewSiteUrl);

    var FinalName = guid + FileExtension;


    webx.get().then(r => {
      var myBlob = this._base64ToArrayBuffer(files.base64);
      webx.getFolderByServerRelativeUrl("MyDocs")
        .files.add(FinalName.toString(), myBlob, true)
        .then(function (data) {
          var RelativeUrls = "MyDocs/" + FinalName;//files.fileList[0].name;
          webx.getFolderByServerRelativeUrl(RelativeUrls).getItem().then(item => {
            // updating Start
            TemFileGuidName[0] = files.fileList[0].name + "|" + item["ID"];
            webx.lists.getByTitle("MyDocs").items.getById(item["ID"]).update({
              Guid0: guid.toString(),
              ActualName: files.fileList[0].name
            }).then(r => {
              component.setState({ loading: false });
              component.setState({ UploadedFilesArray: component.state.UploadedFilesArray.concat(TemFileGuidName) });
            });
          }); //Retrive Doc Info End
        });
    });
  }

  private getFileExtension1(filename) {
    return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;
  }


  private _base64ToArrayBuffer(base64) {
    var binary_string = window.atob(base64.split(',')[1]);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
      bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
  }

  public OnchangeFloor(event: any): void {
    this.setState({ Floor: event.target.value });
  }

  public OnchangeGridLine(event: any): void {
    this.setState({ GridLine: event.target.value });
  }
  public OnchangeBuilding(event: any): void {
    this.setState({ Building: event.target.value });
  }

  public render(): React.ReactElement<IWebpartComplexFormProps> {
    let content;
    if (this.state.loading) {
      content = <div><img src="https://*****************/sites/dev/SiteAssets/loadingnew.gif" /></div>;
    } else {
      // content = <div>{this.state.UploadedFilesArray}</div>;
    }


    let getData = [];
    let str = [];
    for (let i = 0; i < this.state.UploadedFilesArray.length; i++) {
      if (this.state.UploadedFilesArray[i] != undefined) {
        var tempx = this.state.UploadedFilesArray[i].toString().split('|');
        str.push(<li key={tempx[0]} onClick={this.onChangeDeleteDocument.bind(this)} data-id={tempx[1]}> Uploaded File : {tempx[0]} - <a className={styles.MyHeadingsAnchor}>Delete </a></li>);
      }
    }
    getData.push(<ul>{str}</ul>);


    var options = this.state.ProjectsArray.map(function (item, i) {
      var Trmp = item["ID"] + ";#" + item["Title"]
      return <option value={Trmp} key={item["ID"]}>{item["Title"]}</option>
    });
    return (
      <div className={styles.addNewRfi} >
        <h1> SPFX - Complex Form </h1>
        <div>
          {content}
          {getData}
        </div>
        <div className={styles.row}>
          <div className={styles.label}>
            Select Dropdown
           </div>
          <div >
            <select value={this.state.ProjectName} className={styles.myinput} onChange={this.onChangeSelect.bind(this)}>{options}
            </select>
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.label}>
            1)Input Value
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.Building} onChange={this.OnchangeBuilding.bind(this)} />
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.label}>
            2)Input Value 2
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.Floor} onChange={this.OnchangeFloor.bind(this)} />
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.label}>
            3)Input Value 3
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.GridLine} onChange={this.OnchangeGridLine.bind(this)} />
          </div>
        </div>


        <div className={styles.row}>
          <div className={styles.label}>
            Input Value 4
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.Subject} onChange={this.OnchangeSubject.bind(this)} />
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.label}>
           Input Value 5
           </div>
          <div >
            <input type="text" className={styles.myinput} value={this.state.SubContractor} onChange={this.OnchangeSubContractor.bind(this)} />
          </div>
        </div>

        <div className={styles.rowDate}>
          <div className={styles.label}>
            Date
           </div>
          <div className={styles.myinput}>
            <Datetime onChange={this.onSelectDate.bind(this)} />
          </div>
        </div>



        <div className={styles.rowDate}>
          <div className={styles.label}>
            Date Column 2
           </div >
          <div className={styles.myinput}>
            <Datetime onChange={this.onSelectDateRequired.bind(this)} />
          </div>
        </div>


        <div className={styles.row}>
          <div className={styles.label}>
            Check box
           </div >
          <div className={styles.myinput}>
           

      <CheckboxGroup name="fruits" value={this.state.Disciplined} onChange={this.DisciplineChanged}>
      <label> kiwi</label><Checkbox value="kiwi"/>
      <label> pineapple</label> <Checkbox value="pineapple"/>
      <label>  watermelon</label><Checkbox value="watermelon"/>
      <label>  Others</label><Checkbox value="Others"/>
      <label>  Apple</label><Checkbox value="Apple"/>
      <label>  Peach</label><Checkbox value="Peach"/>
      </CheckboxGroup>

          </div>
        </div>

        <div className={styles.rowDate}>
          <div className={styles.label}>
            Rich Textbox Value 1 - Editor
           </div >
          <div className={styles.myinput} ref="RefDescription">
            <App />
          </div>
        </div>


        <div className={styles.rowDate}>
          <div className={styles.label}>
            Mini Editor
           </div >
          <div className={styles.myinput} ref="RefRequirement">
            <App />
          </div>
        </div>


        <div className={styles.rowDate}>
          <div className={styles.label}>
            Minit Editor 2
           </div >
          <div className={styles.myinput} ref="RefComments">
            <App />
          </div>
        </div>

        <div className={styles.row}>
          <ReactFileReader fileTypes={[".csv", ".xlsx", ".Docx"]} handleFiles={this.handleFiles.bind(this)} base64={true} >
            <button className='btn'>Upload</button>
          </ReactFileReader>


        </div>



        <div className={styles.row}>
          <div  >
            <button id="btn_add" className={styles.button} onClick={this.CreateNewItem.bind(this)}>Create New ListItem </button>
          </div>
        </div>

      </div >
    );
  }
}
