import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneLabel, PropertyPaneLink,
  PropertyPaneSlider, PropertyPaneToggle, PropertyPaneDropdown, IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { sp } from "@pnp/sp";
export interface IHelloWorldWebPartProps {
  description: string;
  itemTemplate: string;
  selectedList: string;
  selectedFields: number[] | string[];
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private lstfetched: boolean = false;
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
        itemTemplate: this.properties.itemTemplate
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this.lstfetched) {

      sp.web.lists.getById(this.properties.selectedList).fields.filter("Group ne '_Hidden' and Hidden eq false").get().then(fields => {
        this.listFieldOptions = [];
        fields.map(lst => {
          this.listFieldOptions.push({ key: lst.Id, text: lst.Title });
        });
        console.log(this.listFieldOptions); console.log(fields);
        this.context.propertyPane.refresh();
        this.lstfetched = false;
      }
      ).catch(error => {
        console.log(error);
        this.listFieldOptions = [];

      }
      );
    }

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
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => {
                    this.properties.selectedList = newValue;
                    console.log(this.properties.selectedList);
                    this.lstfetched = true;
                    this.context.propertyPane.refresh();
                  },
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldMultiSelect('listFields', {
                  key: 'multiSelect',
                  label: "Select Fields",
                  options: this.listFieldOptions,
                  selectedKeys: this.properties.selectedFields
                }),
                PropertyFieldCodeEditor('itemTemplate', {
                  label: 'Edit HTML Code',
                  panelTitle: 'Edit HTML Code',
                  initialValue: this.properties.itemTemplate,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
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
