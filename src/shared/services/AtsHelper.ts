import { sp } from "@pnp/sp";
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneLabel, PropertyPaneLink,
    PropertyPaneSlider, PropertyPaneToggle, PropertyPaneDropdown, IPropertyPaneDropdownOption,WebPartContext
  } from '@microsoft/sp-webpart-base';
export default class AtsHelper {
    public static async  LoadProperties(lstfetched:boolean,selectedList:string,context:WebPartContext,listFieldOptions:Array<IPropertyPaneDropdownOption>): Promise<Array<IPropertyPaneDropdownOption>>{
        if (lstfetched) {
         await  sp.web.lists.getById(selectedList).fields.filter("Group ne '_Hidden' and Hidden eq false").get().then(fields => {
            listFieldOptions = [];
            fields.map(lst => {
              listFieldOptions.push({ key: lst.Id, text: lst.Title });
            });
            lstfetched = false;        
            
          }).catch(error => {
            console.log(error);
            listFieldOptions = [];
           
          });
          return listFieldOptions;
        }
      }
}
