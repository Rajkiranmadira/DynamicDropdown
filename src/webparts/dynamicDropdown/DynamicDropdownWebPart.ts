import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {sp} from '@pnp/sp/presets/all'

import * as strings from 'DynamicDropdownWebPartStrings';
import DynamicDropdown from './components/DynamicDropdown';
import { IDynamicDropdownProps } from './components/IDynamicDropdownProps';

export interface IDynamicDropdownWebPartProps {
  description: string;
}

export default class DynamicDropdownWebPart extends BaseClientSideWebPart<IDynamicDropdownWebPartProps> {
  

  protected onInit(): Promise<void> {
    return super.onInit().then(message => {
      sp.setup({
        spfxContext:this.context as any
      })
    });
  }

  public async render(): Promise<void> {
    const element: React.ReactElement<IDynamicDropdownProps> = React.createElement(
      DynamicDropdown,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        singleValueOptions: await this.getChoiceFields(this.context.pageContext.web.absoluteUrl,'singleValueDropdown'),
        multiValueOptions: await this.getChoiceFields(this.context.pageContext.web.absoluteUrl,'multiValueDropdown')
        
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

  private async getChoiceFields(siteurl:string,field:string):Promise<any>{
    try{
      const reponse=await fetch(`${siteurl}/_api/web/lists/GetByTitle('FluentuiDropdown')/fields?$filter=EntityPropertyName eq '${field}'`,{
        method:'GET',
        headers:{
          'Accept':'application/json;odata=nometadata'
        }
      });
      const data=await reponse.json();
      const choices=data?.value[0]?.Choices||[];
      return choices.map((choice:any)=>({
        key:choice,
        text:choice
      }));
    }
    catch(error){
      console.error('Error while fetching chocie',error);
      throw error;
    }
  }
}



