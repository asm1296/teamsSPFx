import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as pnp from 'sp-pnp-js';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/items/list";

import * as strings from 'TeamsReportTabWebPartStrings';
import TeamsReportTab from './components/TeamsReportTab';
import { ITeamsReportTabProps } from './components/ITeamsReportTabProps';

export interface ITeamsReportTabWebPartProps {
  description: string;
  title : string;
  subtitle : string;
}

export default class TeamsReportTabWebPart extends BaseClientSideWebPart <ITeamsReportTabWebPartProps> {

  public render(): void {

    if (this.context.sdks.microsoftTeams){
      this.properties.title = "Welcome to Teams",
      this.properties.subtitle = "Building custom enterprise tabs for your business";
    }
    else {
      this.properties.title = "Welcome to SharePoint",
      this.properties.subtitle = "Customize SharePoint experiences using Web Parts.";
    }
    const element: React.ReactElement<ITeamsReportTabProps> = React.createElement(
      TeamsReportTab,
      {
        description: this.properties.description,
        title : this.properties.title,
        subtitle : this.properties.subtitle
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
}
