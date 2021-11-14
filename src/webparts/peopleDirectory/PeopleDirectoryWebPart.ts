import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { SPComponentLoader } from '@microsoft/sp-loader';
require('../../common/common.css');
import * as strings from 'PeopleDirectoryWebPartStrings';
// import PeopleDirectory from './components/PeopleDirectory';
// import { IPeopleDirectoryProps}  from './components/IPeopleDirectoryProps';
import { PeopleDirectory, IPeopleDirectoryProps } from './components/PeopleDirectory/';

export interface IPeopleDirectoryWebPartProps {
  // description: string;
  title: string;
}

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPeopleDirectoryProps> = React.createElement(
      PeopleDirectory,
      {
        // description: this.properties.description
        webUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        title: this.properties.title,
        displayMode: this.displayMode,
        locale: this.getLocaleId(),
        onTitleUpdate: (newTitle: string) => {
          // after updating the web part title in the component
          // persist it in web part properties
          this.properties.title = newTitle;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
  //   return super.onInit();
  // }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getLocaleId(): string {
    return this.context.pageContext.cultureInfo.currentUICultureName;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        // {
        //   header: {
        //     description: strings.PropertyPaneDescription
        //   },
        //   groups: [
        //     {
        //       groupName: strings.BasicGroupName,
        //       groupFields: [
        //         PropertyPaneTextField('description', {
        //           label: strings.DescriptionFieldLabel
        //         })
        //       ]
        //     }
        //   ]
        // }
      ]
    };
  }
}
