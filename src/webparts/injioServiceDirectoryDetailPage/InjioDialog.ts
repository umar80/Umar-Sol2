import {BaseDialog,IDialogShowOptions,IDialogConfiguration} from "@microsoft/sp-dialog";
import styles from './InjioServiceDirectoryDetailPageWebPart.module.scss';

export default class InjioDialog extends BaseDialog {
    public message: string;
    public colorCode: string;
   
    public render(): void {     
      this.domElement.innerHTML=`
        <div>
         <!-- Hello World!!-->
         <div><input type="button" class="${styles.button}" value="CLOSE" id="btnClose" /></div>
         <div><iframe width="500px" height="500px" src="https://webvine.sharepoint.com/sites/IngiyoLight/ServicesDirectory/Lists/ServiceComments/NewForm.aspx?IsDlg=1&https://webvine.sharepoint.com/sites/IngiyoLight/ServicesDirectory/_layouts/15/workbench.aspx?ID=5"/></div>
        </div>`; 
        
        this.domElement.querySelector('#btnClose').addEventListener('click', () => { this.closeDialog(); });   
    }
   
    public getConfig(): IDialogConfiguration {
      return {
        isBlocking: true
        
      };
    }    

    public closeDialog(): Promise<void>
    {
        return super.close();
    }

   }