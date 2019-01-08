import * as React from 'react';
import styles from './PoliciesViewer.module.scss';
import { IPoliciesViewerProps } from './IPoliciesViewerProps';

import { ServiceScope, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  IDocumentCardPreviewProps
 } from 'office-ui-fabric-react/lib/DocumentCard';
 import { BaseComponent } from 'office-ui-fabric-react/lib/Utilities';
 import { TagPicker, IBasePicker, ITag, TagItem } from 'office-ui-fabric-react/lib/Pickers';
import { ISitePagesService, SitePagesService } from '../services/SitePagesService';
import { ITaxonomyService, TaxonomyService } from '../services/TaxonomyService';
import * as moment from 'moment';
import { ITerms } from '@pnp/sp-taxonomy';
import { RelatedItemManagerImpl } from '@pnp/sp/src/relateditems';

export default class PoliciesViewer extends React.Component<IPoliciesViewerProps, {}> {

  private SitePagesServiceInstance: ISitePagesService;
  private TaxonomyServiceInstance: ITaxonomyService;
  private Tax
  private _noOfItems: number;
  private card: any;
  private renderedCard: any = "";
  private imagesJSON = [];
  private allPages: any = [];
  
  private _filterTags: ITag[] = [].map(item => ({ key: item, name: item }));
  private _selectedTags: ITag[] = [].map(item => ({ key: item, name: item }));

  private _picker = React.createRef<IBasePicker<ITag>>();

  constructor(props: IPoliciesViewerProps) {
    super(props);

    this.state = {
      galleryItems: null,
      isLoading: true,
      showErrorMessage: false
    };

    let serviceScope: ServiceScope;
    serviceScope = this.props.serviceScope;

    this._noOfItems = this.props.imagesToDisplay;

    // Based on the type of environment, return the correct instance of the ImageGalleryServiceInstance interface
    if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
      // Mapping to be used when webpart runs in SharePoint.
      this.SitePagesServiceInstance = serviceScope.consume(SitePagesService.serviceKey);
      this.TaxonomyServiceInstance = serviceScope.consume(TaxonomyService.serviceKey);
    }

    const emptyTags: ITag[] = [];
    this.loadPages(emptyTags);

    this.TaxonomyServiceInstance.getDepartments().then((departments: ITerms) => {
      departments.get().then(terms => {
        terms.forEach(department => {
          this._filterTags.push({ key: 'Dept-' + department.Id, name: department.Name }); //department.Description);
        });
      });
      
      console.log('Termset loaded');
    });

    this.SitePagesServiceInstance.getRoles().then(roles => {
      roles.forEach(role => {
        this._filterTags.push({ key: 'Role-' + role.ID, name: role.Title});
      });
    });

    this.SitePagesServiceInstance.getDepartments().then(roles => {
      roles.forEach(role => {
        this._filterTags.push({ key: 'Dept-' + role.ID, name: role.Title});
      });
    });

  }

  private loadPages(tags: ITag[]) {
    this.SitePagesServiceInstance.getSitePages(this._noOfItems, tags).then((sitePages: any[]) => {
      this.allPages = sitePages;
      this.setState({ isLoading: false });
    });
  }

  private createCards = () => {
    
    
    let cards = [];

    

    {this.allPages.forEach(page => { 
    
      const previewProps: IDocumentCardPreviewProps = {
        previewImages: [
          {
            previewImageSrc: String(require('./document-preview.png')),
            iconSrc: String(require('./icon-ppt.png')),
            // width: 318,
            // height: 196,
            accentColor: '#ce4b1f'
          }
        ],
      };
      
      const webAbsoluteUrl = this.SitePagesServiceInstance._pageContext.web.absoluteUrl;
      const tenancyUrl = this.SitePagesServiceInstance._pageContext.web.absoluteUrl.replace(this.SitePagesServiceInstance._pageContext.web.serverRelativeUrl,'');
      let iconUrl = this.SitePagesServiceInstance._pageContext.web.absoluteUrl + '/_layouts/15/getpreview.ashx?clientType=docLibGrid&guidFile={';
      iconUrl += page.UniqueId;
      iconUrl += '}&guidSite={9928f210-b6c2-46bf-b232-8ac7fae55229}&guidWeb={cb3fd4ca-4597-44d6-889f-7ae9eb51b066}&resolution=Width1024';
      previewProps.previewImages[0].previewImageSrc = iconUrl;

      const card = (<div className="ms-Grid-col ms-u-sm-12 ms-u-md-6 ms-u-lg-4">
      <DocumentCard onClickHref={webAbsoluteUrl + '/SitePages/' + page.FileLeafRef}>
      <DocumentCardPreview { ...previewProps } />
      <DocumentCardTitle title={ page.Title } />
      <DocumentCardActivity
        activity={ moment(page.Modified).fromNow() }
        people={
          [
            { name: 'Last modified:', profileImageSrc: '' }
          ]
        }
      />
      </DocumentCard>
      </div>);
      cards.push(card);
     })
    }
    {
      return cards;
    }

    
  
}

  public render(): JSX.Element {
    const details = (
      <div className={styles.PoliciesViewer}>
      <div className="ms-Grid">
      <div className={styles.row}>
        <div className="ms-font-su ms-Grid-col ms-u-sm-12 ms-u-md-12 ms-u-lg-12">
          Policies
        </div>
      </div>
      <div className={styles.paddedBottomRow}>
        <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
          <span className="ms-font-l">Filter policies: </span>
        </div>
        <div className="ms-Grid-col ms-sm12 ms-md10 ms-lg10">
          <TagPicker
            onResolveSuggestions={this._onFilterChangedNoFilter}
            onItemSelected={this._onItemSelected}
            onChange={this._onPolicyFilterChanged}
            getTextFromItem={this._getTextFromItem}
            pickerSuggestionsProps={{
              suggestionsHeaderText: 'Suggested Tags',
              noResultsFoundText: 'No Color Tags Found'
            }}
            itemLimit={2}
            disabled={ false }
            inputProps={{
              onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
              onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
              'aria-label': 'Tag Picker'
            }}
            /*
            componentRef={this._picker}
            */
          />
        </div>
      </div>
      
      <div className="ms-Grid-row">
        { this.createCards() }
      </div>
      </div>
      </div>
    );
    return details;
    // </div></div>return this.createCards();
  }

  private _getTextFromItem(item: ITag): string {
    return item.name;
  }

  private _onPolicyFilterChanged = (tags: ITag[]): ITag[] => {
    this.loadPages(tags);
    
    return tags;
  };

  private _onFilterChanged = (filterText: string, tagList: ITag[]): ITag[] => {
    return filterText
      ? this._filterTags
          .filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0)
          .filter(tag => !this._listContainsDocument(tag, tagList))
      : [];
  };

  private _onFilterChangedNoFilter = (filterText: string, tagList: ITag[]): ITag[] => {
    return filterText ? this._filterTags.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0) : [];
  };

  private _onItemSelected = (item: ITag): ITag | null => {
    if (this._picker.current && this._listContainsDocument(item, this._picker.current.items)) {
      return null;
    }
    return item;
  };

  private _listContainsDocument(tag: ITag, tagList?: ITag[]) {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }
}
