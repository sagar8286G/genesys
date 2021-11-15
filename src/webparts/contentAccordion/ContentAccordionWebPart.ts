import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ContentAccordionWebPartStrings';
import ContentAccordion from './components/ContentAccordion';
import { IContentAccordionProps } from './components/IContentAccordionProps';

export interface IContentAccordionWebPartProps {
  description: string;
  listName:string;
  accordionTitle: string;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
}

export default class ContentAccordionWebPart extends BaseClientSideWebPart <IContentAccordionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IContentAccordionProps> = React.createElement(
      ContentAccordion,
      {
        description: this.properties.description,
        context:this.context,
        listName:this.properties.listName,
        accordionTitle: this.properties.accordionTitle,
        allowZeroExpanded: this.properties.allowZeroExpanded,
        allowMultipleExpanded: this.properties.allowMultipleExpanded
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
                }),                
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneTextField('accordionTitle', {
                  label: strings.accordionTitle
                }),
                PropertyPaneToggle('allowZeroExpanded', {
                  label: 'Allow zero expanded',
                  checked: this.properties.allowZeroExpanded,
                  key: 'allowZeroExpanded',
                }),
                PropertyPaneToggle('allowMultipleExpanded', {
                  label: 'Allow multiple expand',
                  checked: this.properties.allowMultipleExpanded,
                  key: 'allowMultipleExpanded',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
