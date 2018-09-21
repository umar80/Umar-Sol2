/*require('sp-init');
require('microsoft-ajax');
require('script-resx');
require('sp-runtime');
require('sharepoint');
require('sharepoint-init');
require('SP-UI-Dialog');
*/
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {SPComponentLoader} from '@microsoft/sp-loader';

import pnp, { CamlQuery, sp } from "sp-pnp-js";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

import MockHttpClient from './MockHttpClient';
import InjioDialog from './InjioDialog';
import * as JQuery from 'jquery';
import styles from './InjioServiceDirectoryDetailPageWebPart.module.scss';
import * as strings from 'InjioServiceDirectoryDetailPageWebPartStrings';
import { ServiceDirectory, ServiceDirectorys } from './ServiceDirectoryList';
import { ServiceContact, ServiceContacts } from './ServiceContacts';
import { AdditionalAttribute, AdditionalAttributes } from './ServiceDirectoryAdditionalAttributes';
import {ServiceComment, ServiceComments} from './ServiceComments';
import { BaseDialog, IDialogConfiguration, Dialog } from '@microsoft/sp-dialog';


export interface IInjioServiceDirectoryDetailPageWebPartProps {
  description: string;
}

export default class InjioServiceDirectoryDetailPageWebPart extends BaseClientSideWebPart<IInjioServiceDirectoryDetailPageWebPartProps> {


  private myParam:any = "";
  protected onInit():Promise<void>{
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }


  private loadSP() : Promise<any> {
    var globalExportsName = null, p = null;
    var promise = new Promise<any>((resolve, reject) => {
      globalExportsName = '$_global_init'; p = (window[globalExportsName] ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/init.js', { globalExportsName }));
      p.catch((error) => { })
        .then(($_global_init): Promise<any> => {
          globalExportsName = 'Sys'; p = (window[globalExportsName] ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'Sys'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((SP): Promise<any> => {
          globalExportsName = 'SP'; p = ((window[globalExportsName] && window[globalExportsName].ClientContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.js', { globalExportsName }));
          return p;
        })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.Init.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.UI.Dialog.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((SP) => {
          resolve(SP);
        });
    });
    return promise;
  }


  /*private _loadSPJSOMScripts() {
 
    const siteColUrl = this.context.pageContext.web.absoluteUrl;
    
    console.log("Site Collectoin URL"+siteColUrl);
    try {
      SPComponentLoader.loadScript(siteColUrl+'/_layouts/15/sp.init.js', {
        globalExportsName: '$_global_init'
      })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl+'/_layouts/15/sp.ui.dialog.js', {
            globalExportsName: 'SP.UI.Dialog'
          });
        }).catch((error) => {         
          
      });
      
        /*.then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/reputation.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js', {
            globalExportsName: 'jQuery'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/siteassets/jquery.rateit.min.js', {            
            globalExportsName: 'jQuery',            
          });
        })        
        .then((): void => {          
          SPComponentLoader.loadCss(siteColUrl + '/siteassets/rateit.css');
          jQuery('rateit').rateit();
        })
      } catch (error) {

      }
  }*/

//#region GET MOCKUP LIST DATA
  private _getMockListData(): Promise<ServiceDirectorys> {
    return MockHttpClient.getServiceDirectoryItem()
      .then((data: ServiceDirectory[]) => {
        var listData: ServiceDirectorys={ value: data };
        return listData;
      }) as Promise<ServiceDirectorys>;
  }

  private _getMockContactListData(): Promise<ServiceContacts> {
    return MockHttpClient.getServiceContacts()
      .then((data: ServiceContact[]) => {
        var listData: ServiceContacts={ value: data };
        return listData;
      }) as Promise<ServiceContacts>;
  }

  private _getMockAdditionalAttributeListData(): Promise<AdditionalAttributes> {
    return MockHttpClient.getAdditionalAttributes()
      .then((data: AdditionalAttribute[]) => {
        var listData: AdditionalAttributes={ value: data };
        return listData;
      }) as Promise<AdditionalAttributes>;
  }

  private _getMockCommentsData():Promise<ServiceComments>{
    return MockHttpClient.getServiceComments()
    .then((data:ServiceComment[]) => {
      var listData:ServiceComments={value:data};
      return listData;
    }) as Promise<ServiceComments>;


  }
  //#endregion


//#region  RENDERING
  public render(): void {
    //GETTING QUERY STRING
    
    //let myParam = 5;
    var queryParms = new UrlQueryParameterCollection(window.location.href);
      console.log(queryParms);
     this.myParam = queryParms.getValue("SID");
      console.log(this.myParam);

    //LOCAL WORKBENCH ENVIRONMENT
    if (Environment.type === EnvironmentType.Local) {
        this._getMockListData().then((response) => {
          this._renderMain(response.value);
        });

        this._getMockAdditionalAttributeListData().then((response) => {
          this._renderServiceDirectoryAdditionalAttributes(response.value);
        });
       
        this._getMockCommentsData().then((response) => {
          this._renderServiceDirectoryComments(response.value);
        });

        this._getMockContactListData().then((response) => {
          this._renderServiceDirectoryContacts(response.value);
        });
    }
    else if(Environment.type == EnvironmentType.SharePoint 
      || Environment.type == EnvironmentType.ClassicSharePoint )
    {
    //  this._loadSPJSOMScripts();
        //GET DATA FROM SHAREPOINT LISTS.
        if(this.myParam != null || this.myParam != undefined)
        {
          this.loadSP();
          this._getServiceProviderData(this.myParam);
        }
        
       // this._getCommentsData(myParam);
       // this._getServiceDataProviderData(myParam);
       // this._getContactsData(myParam);
    }

    


   /* this.domElement.innerHTML = `
      <div class="${ styles.injioServiceDirectoryDetailPage }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${myParm}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;*/
  }
//#endregion  RENDERING


protected _getServiceProviderData(providerId):void{
  let _query="";
  let whereClause = `<Eq><FieldRef Name='ID'/><Value Type='Text'>${providerId}</Value></Eq>`;
//  whereClause
  _query=`
        <View>
          <Query>
            <Where>${whereClause}</Where>
          </Query>
        </View>`;
        console.log(this);
        this.getServiceDataProviderData(_query)
        .then((response) => {

            this._renderMain(response);
        });
        
}

protected _getCommentsData(providerId):void{
  let _query="";
  let whereClause = `<Eq><FieldRef Name='Service_x003a_ID' Lookup="TRUE"/><Value Type='Lookup'>${providerId}</Value></Eq>`;
//  whereClause
  _query=`
        <View Scope="Recursive">
          <Query>
            <Where>${whereClause}</Where>
          </Query>
        </View>`;

        this.getCommentsData(_query)
        .then((response) => {

            this._renderServiceDirectoryComments(response);
        });
}


protected _getServiceContractsData(providerId):void{
  let _query="";
  let whereClause = `<Eq><FieldRef Name='Service_x003a_ID' Lookup="TRUE"/><Value Type='Lookup'>${providerId}</Value></Eq>`;
//  whereClause
  _query=`
        <View>
          <Query>
            <Where>${whereClause}</Where>
          </Query>
        </View>`;

        this.getServiceContractsData(_query)
        .then((response) => {

            this._renderServiceDirectoryAdditionalAttributes(response);
        });
}


protected _getContactsData(providerId):void{
  let _query="";
  let whereClause = `<Eq><FieldRef Name='Service_x003a_ID' Lookup="TRUE"/><Value Type='Lookup'>${providerId}</Value></Eq>`;
  _query=`
        <View>
          <Query>
            <Where>${whereClause}</Where>
          </Query>
        </View>`;

        this.getContactsData(_query)
        .then((response) => {
            console.log(response);
            this._renderServiceDirectoryContacts(response);
        });
}

protected getServiceDataProviderData(xml:string):Promise<ServiceDirectory[]>{
  const q:CamlQuery={ViewXml:xml};
  return pnp.sp.web.lists.getByTitle("Services Directory").getItemsByCAMLQuery(q).then((response)=>{
    return response;    
  });
}

protected getCommentsData(xml:string):Promise<ServiceComment[]>{
  const q:CamlQuery={ViewXml:xml};
  //let srvComments =  pnp.sp.web.lists.getByTitle("SERVICES DIRECTORY COMMENTS").select("Comments","Created", "Author/Title","Author/Name").expand("Author");
  //console.log(srvComments);
 // return srvComments.getItemsByCAMLQuery(q)
  
  return pnp.sp.web.lists.getByTitle("SERVICES DIRECTORY COMMENTS").items.filter("ServiceId eq '5'").select("ID","Comments","Created", "Author/Title","Author/Name", "Service/ID").expand("Author","Service").orderBy("Created",false).get().then((response)=>{
    console.log("Response 11");
    console.log(response);
    return response;
  });
}



protected getServiceContractsData(xml:string):Promise<AdditionalAttribute[]>{
  const q:CamlQuery={ViewXml:xml};
  return pnp.sp.web.lists.getByTitle("SERVICES DIRECTORY ADDITIONAL ATTRIBUTES").getItemsByCAMLQuery(q).then((response)=>{
    return response;
  });
}

protected getContactsData(xml:string):Promise<ServiceContact[]>{
  const q:CamlQuery={ViewXml:xml};
  return pnp.sp.web.lists.getByTitle("SERVICE CONTACTS").getItemsByCAMLQuery(q).then((response)=>{
    return response;
  });
}

protected _renderMain(response){

    console.log(this);
    if(response != null)
    {
      this.domElement.innerHTML = `    
      <div id="MainContent" class=${styles.injioServiceDirectoryDetailPage}>
          <div id="serviceDirectory" >
              <!-- Services Directory Main Content will go here -->
          </div>	
          <div id="contract" class=${styles.contract}>
           <!--    Contracts Content will go here -->
          </div>
          <div id="comments" class=${styles.comments}>
          <!-- COMMENTS CONTENT WILL GO HERE -->
          </div>
          <div id="contacts" class=${styles.contacts} >
          <!-- Contacts CONTENT WILL GO HERE -->
        </div>
      </div>`;
      this._renderServiceDirectoryList(response);
      //this._renderServiceDirectoryContracts(response);
    }
    this._getServiceContractsData(this.myParam);
}

protected _renderServiceDirectoryList(response){
    let serviceDirectoryHTML="";
    console.log(this);
    if(response != null)
    {
        serviceDirectoryHTML=`<div id="leftContent" class=${styles["left-content"]} >
              <div class=${styles.logo}><img class=${styles.logo} src="https://webvine.sharepoint.com/sites/MIqbalTest/PublishingImages/Lists/Service%20Directory/AllItems/asus.jpg" /></div>
            <div style="position:relative; left:10px;">
              <div class=${styles.providerTitle} >${response[0].Title != null?response[0].Title:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Address:</span> ${response[0].Address != null?response[0].Address:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Region:</span>  ${response[0].Region != null?response[0].Region:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Country:</span> ${response[0].Country != null?response[0].Country:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Website:</span> ${response[0].Website != null?response[0].Website:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Phone:</span> ${response[0].Phone != null?response[0].Phone:"-"}</div>
              <div><span class=${styles.fieldCaptions}>ABN:</span> ${response[0].ABN != null?response[0].ABN:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Rating:</span> ${response[0].AverageRating != null?response[0].AverageRating:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Contact:</span> ${response[0].Contact != null?response[0].Contact:"-"}</div>
              <div><span class=${styles.fieldCaptions}>Service Type:</span> ${response[0].ServiceType != null?response[0].ServiceType:"-"}</div>
            </div>
            <div class=${styles["right-content"]}>
                <iframe src="${response[0].LocationMap != null?response[0].LocationMap:'-'}" width="200" height="200" frameborder="0" style="border:0" allowfullscreen></iframe>
            </div>
              <div class=${styles.providerDescription}>${response[0].Description != null?response[0].Description:"-"}</div>
        </div>
        `;
        this.domElement.querySelector('#serviceDirectory').innerHTML=`${serviceDirectoryHTML}`;
    }
}


protected _renderServiceDirectoryComments(response){
  console.log(response);
    let serviceDirectoryCommentsHTML="";
    let i = 0;
        serviceDirectoryCommentsHTML = `<div>COMMENTS<span><input type="button" class="${styles.button}" value="ADD Commment" id="btnAddComment" /></span></div>
                                        <div id="commentsHeader" class=${styles.commentsHeader}>
                                            <div>Description</div>
                                            <div>Created By</div>
                                            <div>Created</div>
                                      </div>`;
    if(response != null)
    {
      for(i=0;i<response.length;i++)
      {
 
          //new Intl.DateTimeFormat().format(new Date().getDate())
         // console.log(new Intl.DateTimeFormat('en-AU', options).format(response[i].Created));
          console.log(new Date(response[i].Created).toLocaleDateString());
          console.log("Response 2001");
          console.log(response[i]);
          serviceDirectoryCommentsHTML+=`<div id="commentsContent" class=${styles.commentsContent}>
                                            <div id="commentsItem" class=${styles.commentsItem}>
                                                <div>${response[i].Comments}</div>
                                                <div>${response[i].Author.Title}</div>
                                                <div>${new Date(response[i].Created).toLocaleDateString()}</div>
                                            </div>
                                            </div>`;
      }
      
      this.domElement.querySelector('#comments').innerHTML=`${serviceDirectoryCommentsHTML}`;
      this.domElement.querySelector('#btnAddComment').addEventListener('click', () => { this.showDialog("comments");});
      console.log("1004");
    }
    this._getContactsData(this.myParam);
}

  
protected _renderServiceDirectoryContacts(response){
    let serviceDirectoryContactsHTML="";
    let j=0;
        serviceDirectoryContactsHTML=`<div>CONTACTS <span><input type="button" class="${styles.button}" value="ADD Contacts" id="btnAddContact" /></span></div>
                                      <div id="contactsHeader" class=${styles.contactsHeader}>
                                          <div>First Name</div>
                                          <div>Last Name</div>
                                          <div>Job Title</div>
                                          <div>Phone</div>
                                          <div>Email</div>
                                     </div>`;
    if(response != null)
    {

        for(j=0;j<response.length;j++){
            serviceDirectoryContactsHTML+=`<div id="contactsContent" class=${styles.commentsContent}>
                                            <div id="contactItems" class=${styles.contactItems} >
                                                <div>${response[j].FirstName}</div>
                                                <div>${response[j].LastName}</div>
                                                <div>${response[j].JobTitle}</div>
                                                <div>${response[j].BusinessPhone}</div>
                                                <div>${response[j].EMail}</div>
                                            </div>
                                          </div>`;
        }
      this.domElement.querySelector('#contacts').innerHTML=`${serviceDirectoryContactsHTML}`;
      this.domElement.querySelector('#btnAddContact').addEventListener('click', () => { this.showDialog("contacts"); });  
    } 
  
}

 protected _renderServiceDirectoryAdditionalAttributes(response){
    let serviceDirectoryContractsHTML= "";
        serviceDirectoryContractsHTML= `<div>THQ<span><input type="button" class="${styles.button}" value="ADD" id="btnAddTHQ" /></span></div>
                                        <div id="THQHeaders" class=${styles.THQHeader}>
                                          <div>Start Date</div>
                                          <div>End Date</div>
                                          <div>Notice Period Date</div>
                                          <div>Kennards Primary Contact</div>
                                       </div>`;
                                        
    if(response != null)
    {
      //console.log(response);
      serviceDirectoryContractsHTML+=`<div id="contractsContent" class=${styles.commentsContent}>
                                        <div id="THQItem" class=${styles.THQItem}>
                                            <div>${new Date(response[0].StartDate).toLocaleDateString() }</div>
                                            <div>${new Date(response[0].EndDate).toLocaleDateString()}</div>
                                            <div>${new Date(response[0].NoticePeriodDate).toLocaleDateString()}</div>
                                            <div>${response[0].PrimaryContact}</div>
                                        </div>
                                      </div>`;
      this.domElement.querySelector('#contract').innerHTML=`${serviceDirectoryContractsHTML}`;
      this.domElement.querySelector('#btnAddTHQ').addEventListener('click', () => { this.showDialog("contract");});   
    }
    this._getCommentsData(this.myParam);

  }

  protected _openDialog():void{    
    let dialog=new InjioDialog({isBlocking:true});    
    dialog.show();
  }   

  protected varDialogCalledFor = "";
  protected webpartObj = null;
  protected showDialog(calledFor){

    this.webpartObj = this;
    console.log(this.webpartObj);

            console.log('yep');
            let formUrl = "";
            let formTitle = "";
            this.varDialogCalledFor = calledFor;
            console.log(this.varDialogCalledFor);
          switch(calledFor)
          {
            case "comments":
              formUrl   = "/sites/IngiyoLight/ServicesDirectory/Lists/Services%20Directory%20Comments/NewForm.aspx";
              formTitle ="Add Comments";
            break;
            case "contacts":
              formUrl   = "/sites/IngiyoLight/ServicesDirectory/Lists/ServiceComments/NewForm.aspx";
              formTitle ="Add Contacts";
            break;
            case "contract":
              formUrl   = "/sites/IngiyoLight/ServicesDirectory/Lists/Services%20Directory%20Additional%20Attributes/NewForm.aspx";
              formTitle ="Add Contract";
            break;
          }
            let options : SP.UI.IDialogOptions = {
              url: formUrl,
              title: formTitle,
              allowMaximize: false,
              showClose: true,
              width: 800,
              height: 530,
              dialogReturnValueCallback: this.refreshCommonCallback//(calledFor=="comments"?this.refreshCommentsCallback:(calledFor=="contacts"?this.refreshContactsCallback:(calledFor=="contract"?this.refreshContractCallback:this.refreshContactsCallback)))
            };
            SP.UI.ModalDialog.showModalDialog(options);
            console.log("1002");
    }


    protected refreshCommentsCallback2(dialogResult, returnValue )
    {
      
      switch(dialogResult){
        case SP.UI.DialogResult.invalid: 
            break;
        case SP.UI.DialogResult.cancel: 
            console.log("Cancel called");
            break;
        case SP.UI.DialogResult.OK: 
            console.log("Save called");
            //var objRef = new InjioServiceDirectoryDetailPageWebPart();
            //objRef._getCommentsData(5);
          //  InjioServiceDirectoryDetailPageWebPart.call(objRef.refreshData());
            break;
        }
    }
protected refreshCommonCallback(dialogResult, returnValue ):Promise<any>
  {
    
    var promise = new Promise<any>((resolve,reject)=>{
      switch(dialogResult){
        case SP.UI.DialogResult.invalid: 
            break;
        case SP.UI.DialogResult.cancel: 
            console.log("Cancel called");
            break;
        case SP.UI.DialogResult.OK: 
            console.log("Save called");
            var objRef = new InjioServiceDirectoryDetailPageWebPart();
            objRef.refreshData();
            break;
        }
      }).then(():void=>{
        var objRef = new InjioServiceDirectoryDetailPageWebPart();
            console.log("1010");
            objRef.render();
      });
    return promise;
  }

protected refreshData():void
  {
      console.log("1009");
      console.log(this);
      console.log(window.location.href);
      window.location.href = window.location.href;
     // paramref._getServiceProviderData(5);
  }

protected refreshContractCallback(dialogResult,returnValue)
  {
    switch(dialogResult){
      case SP.UI.DialogResult.invalid: 
          break;
      case SP.UI.DialogResult.cancel: 
          console.log("Cancel called");
          break;
      case SP.UI.DialogResult.OK: 
          console.log("Save called");
          this._getServiceContractsData(this.myParam);
          break;
      }

  }

protected refreshContactsCallback(dialogResult,returnValue)
  {
      switch(dialogResult)
      {
         case SP.UI.DialogResult.invalid: 
           break;
         case SP.UI.DialogResult.cancel: 
           console.log("Cancel called");
           break;
         case SP.UI.DialogResult.OK: 
          console.log("Save called");
          this._getContactsData(this.myParam);
          break;
      }    
  }

  // Dialog callback
/*private CloseCallback(result, target) {
  if (result == SP.UI.DialogResult.OK) {
      // Run OK Code
      // reload your page again
      console.log("1");
  }
  if (result == SP.UI.DialogResult.cancel) {
      // Run Cancel Code
      console.log("2");
  }
}*/

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
