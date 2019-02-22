import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  WebPartContext
} from "@microsoft/sp-webpart-base";
import { JSONParser } from "@pnp/odata";

import * as strings from "HelloWorldWebPartStrings";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";

import {
  PropertyFieldCodeEditor,
  PropertyFieldCodeEditorLanguages
} from "@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import { sp } from "@pnp/sp";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { isEmpty } from "@microsoft/sp-lodash-subset";
export interface IHelloWorldWebPartProps {
  description: string;
  itemTemplate: string;
  selectedList: string;
  selectedFields: number[] | string[];
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<
  IHelloWorldWebPartProps
> {
  private isListFetched: boolean = false;
  private areFieldsLoads: boolean = false;
  private listFieldOptions: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        itemTemplate: this.properties.itemTemplate,
        selectedList: this.properties.selectedList,
        selectedFields: this.properties.selectedFields
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.LoadProperties(
      this.isListFetched,
      this.properties.selectedList,
      this.context,
      this.listFieldOptions
    );

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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker("list", {
                  label: "Select a list",
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: (
                    propertyPath: string,
                    oldValue: any,
                    newValue: any
                  ) => {
                    if (newValue==="") {
                      console.log(newValue);
                      this.properties.selectedFields=null;
                    }else{
                      this.properties.selectedList = newValue;
                      console.log(this.properties.selectedList);
                      this.isListFetched = true;
                      this.context.propertyPane.refresh();
                    }
                     
                    
                  },
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId"
                }),
                PropertyFieldMultiSelect("selectedFields", {
                  key: "selectedFieldsKey",
                  label: "Select Fields",
                  options: this.listFieldOptions,
                  selectedKeys: this.properties.selectedFields
                }),
                PropertyFieldCodeEditor("itemTemplate", {
                  label: "Edit HTML Code",
                  panelTitle: "Edit HTML Code",
                  initialValue: this.properties.itemTemplate,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: "codeEditorFieldId",
                  language: PropertyFieldCodeEditorLanguages.HTML
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private setFieldOptions(selectedList: string) {
    sp.web.lists
      .getById(selectedList)
      .fields.filter("Group ne '_Hidden' and Hidden eq false")
      .get()
      .then(fields => {
        console.log(fields);
        this.listFieldOptions = [];
        fields.map(lst => {
          this.listFieldOptions.push({ key: lst.Id, text: lst.Title });
        });
        this.context.propertyPane.refresh();
        this.isListFetched = false;
      })
      .catch(error => {
        console.log(error);
        this.listFieldOptions = [];
      });
  }

  private LoadProperties(isListFetched: boolean, selectedList: string,context: WebPartContext,listFieldOptions: IDropdownOption[] ) {
   // this.listFieldOptions = [];
   if (isListFetched) 
    {
      this.setFieldOptions(this.properties.selectedList);
    }

    if(this.properties.selectedList){
      this.setFieldOptions(this.properties.selectedList);

    }

  }
}
