import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PoliciesViewerWebPartStrings';
import PoliciesViewer from './components/PoliciesViewer';
import { IPoliciesViewerProps } from './components/IPoliciesViewerProps';
import { setup as pnpSetup } from "@pnp/common";
import { sp } from '@pnp/sp';
import pnp from 'sp-pnp-js';

export interface IPoliciesViewerWebPartProps {
  description: string;
}

export default class PoliciesViewerWebPart extends BaseClientSideWebPart<IPoliciesViewerWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
      });
      sp.setup({
        spfxContext: this.context
      });
      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  
  public render(): void {
    const element: React.ReactElement<IPoliciesViewerProps > = React.createElement(
      PoliciesViewer,
      {
        serviceScope: this.context.serviceScope,
        imageGalleryName: 'Site Pages',
        imagesToDisplay: 20
        // description: this.properties.description
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
