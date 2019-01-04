import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PoliciesViewerWebPart.module.scss';
import * as strings from 'PoliciesViewerWebPartStrings';

import { setup as pnpSetup } from "@pnp/common";


import * as $ from 'jquery';
import { JQuery } from 'jquery';
const Masonry: any = require('masonry-layout');
const jQueryBridget: any = require('jquery-bridget/jquery-bridget');
jQueryBridget('masonry', Masonry, $);

export interface IPoliciesViewerWebPartProps {
  description: string;
}

export default class PoliciesViewerWebPart extends BaseClientSideWebPart<IPoliciesViewerWebPartProps> {

  private $masonry: any = undefined;

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      pnpSetup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    
    this.domElement.innerHTML = `
      <h2>${this.properties.description}</h2>
      <div class="${styles.policiesViewer}"></div>`;

      const $container: JQuery = $(`.${styles.policiesViewer}`, this.domElement);
    for (let i: number = 0; i < 15; i++) {
      const height: number = Math.floor(Math.random() * (200 - 100 + 1)) + 100;
      $container.append(`<img src="http://lorempixel.com/150/${height}/?d=${new Date().getTime().toString()}" width="150" height="${height}" />`);
    }

    if (this.renderedOnce) {
      this.$masonry.masonry('destroy');
    }

    this.$masonry = ($container as any).masonry({
      itemSelector: 'img',
      columnWidth: 150,
      gutter: 10
    });
    /*this.domElement.innerHTML = `
      <div class="${ styles.policiesViewer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      */
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
