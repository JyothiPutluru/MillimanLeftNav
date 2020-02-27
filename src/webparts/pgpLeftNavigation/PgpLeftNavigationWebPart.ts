import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  DynamicDataSharedDepth,
  IPropertyPaneConditionalGroup,
  IWebPartPropertiesMetadata,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'PgpLeftNavigationWebPartStrings';
import PgpLeftNavigation from './components/PgpLeftNavigation';
import { IPgpLeftNavigationProps } from './components/IPgpLeftNavigationProps';
import { SearchComponentType } from '../../models/SearchComponentType';
import { IDynamicDataPropertyDefinition, IDynamicDataCallables } from '@microsoft/sp-dynamic-data';
import { IRefinementFilter } from '../../models/ISearchResult';
import ISearchResultSourceData from '../../models/ISearchResultSourceData';
import { DynamicProperty } from '@microsoft/sp-component-base';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { ISearchRecievedProperties } from './components/pgp.models';
import IRefinerSourceData from '../../models/IRefinerSourceData';

export interface IPgpLeftNavigationWebPartProps {
  description: string;
  searchQuery: DynamicProperty<string>;
  searchResultsWebPart: DynamicProperty<ISearchResultSourceData>;
  refinerSourceData: DynamicProperty<IRefinerSourceData>;
  queryKeywords: string;
  defaultQueryKeywords: string;
  taxonomyTermstoreName: string;
  taxonomyTermstoreId: string;
  taxonomyTermGroupId: string;
  taxonomyTermsetId: string;
  taxonomyTermId: string;
  mappedField: string;
  showTitle: boolean;
  title: string;
  titleLinkUrl: string;
  isTreeView: boolean;
}

export default class PgpLeftNavigationWebPart extends BaseClientSideWebPart<IPgpLeftNavigationWebPartProps> {

  private searchRecievedProperties: ISearchRecievedProperties = {
    Refiners: [],
    SearchBoxWebPart: "",
    QueryKeywords: "",
    TermData: null,
    RefinerSourceData: null
  };
  public render(): void {
    const element: React.ReactElement<IPgpLeftNavigationProps> = React.createElement(
      PgpLeftNavigation,
      {
        description: this.properties.description,
        context: this.context,
        searchQuery: this.properties.searchQuery.tryGetValue(),
        refinerSourceData: this.properties.refinerSourceData.tryGetValue(),
        queryKeywords: this.properties.queryKeywords,
        defaultQueryKeywords: this.properties.defaultQueryKeywords,
        transmitSearchProperties: this.transmitSearchProperties,
        taxonomyTermstoreId: this.properties.taxonomyTermstoreId,
        taxonomyTermGroupId: this.properties.taxonomyTermGroupId,
        taxonomyTermsetId: this.properties.taxonomyTermsetId,
        taxonomyTermId: this.properties.taxonomyTermId,
        mappedField: this.properties.mappedField,
        showTitle: this.properties.showTitle,
        title: this.properties.title,
        titleLinkUrl: this.properties.titleLinkUrl,
        isTreeView: this.properties.isTreeView
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private transmitSearchProperties = (changedProperties: ISearchRecievedProperties): void => {
    if (changedProperties) {
      this.searchRecievedProperties.RefinerSourceData = changedProperties.RefinerSourceData;
      this.context.dynamicDataSourceManager.notifyPropertyChanged("refinersWebPart");

      this.searchRecievedProperties.QueryKeywords = changedProperties.QueryKeywords;
      this.context.dynamicDataSourceManager.notifyPropertyChanged("queryKeywords");

      this.searchRecievedProperties.SearchBoxWebPart = changedProperties.SearchBoxWebPart;
      this.context.dynamicDataSourceManager.notifyPropertyChanged("defaultQueryKeywords");

      this.searchRecievedProperties.TermData = changedProperties.TermData;
      this.context.dynamicDataSourceManager.notifyPropertyChanged("termData");
    }
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private validateTermstoreidValue(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Provide a termstore id value';
    }
    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let showTitle = ((this.properties.showTitle != undefined || this.properties.showTitle != null) ? this.properties.showTitle : true);
    let isTreeView = ((this.properties.isTreeView != undefined || this.properties.isTreeView != null) ? this.properties.isTreeView : true);
    let allEventsUrl: any = [];

    if (showTitle) {
      allEventsUrl = [PropertyPaneTextField('title', {
        label: strings.Title
      }), PropertyPaneTextField('titleLinkUrl', {
        label: strings.TitleLinkUrl
      })];
    }
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("taxonomyTermstoreId", {
                  label: strings.TaxonomyTermstoreId,
                  onGetErrorMessage:this.validateTermstoreidValue.bind(this)
                }),
                PropertyPaneTextField("taxonomyTermGroupId", {
                  label: strings.TaxonomyGroupId
                }),
                PropertyPaneTextField("taxonomyTermsetId", {
                  label: strings.TaxonomyTermstoreId
                }),
                PropertyPaneTextField("taxonomyTermId", {
                  label: strings.TaxonomyTermId
                }),
                PropertyPaneTextField("mappedField", {
                  label: strings.MappedField
                }),
                PropertyPaneToggle('isTreeView', {
                  label: strings.IsTreeeView,
                  checked: isTreeView
                }),
                PropertyPaneToggle('showTitle', {
                  label: strings.ShowTitle,
                  checked: showTitle
                })
              ].concat(allEventsUrl)
            }
          ]
        },
        {
          groups: [{
            primaryGroup: {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            secondaryGroup: {
              groupName: "Connect to Source",
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: "searchQuery",
                  fields: [
                    PropertyPaneDynamicField("searchQuery", {
                      label: "Select a query source"
                    })
                  ]
                }),
                PropertyPaneDynamicFieldSet({
                  label: SearchComponentType.RefinersWebPart,
                  fields: [
                    PropertyPaneDynamicField("refinerSourceData", {
                      label: "Select a refiners source"
                    })
                  ]
                })
              ],
              sharedConfiguration: {
                depth: DynamicDataSharedDepth.Property
              }
            }, showSecondaryGroup: !false
          } as IPropertyPaneConditionalGroup]
        }
      ]
    };
  }
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'searchQuery': {
        dynamicPropertyType: 'string'
      },
      'refinerSourceData': {
        dynamicPropertyType: 'string'
      }
    };
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'defaultQueryKeywords',
        title: 'defaultQueryKeywords'
      },
      {
        id: 'refinersWebPart',
        title: 'refinersWebPart'
      },
      {
        id: 'queryKeywords',
        title: 'queryKeywords'
      },
      {
        id: "termData",
        title: "termData"
      }
    ];
  }

  public getPropertyValue(propertyId: string) {
    switch (propertyId) {
      case 'refinersWebPart':
        return this.searchRecievedProperties.RefinerSourceData;
      case 'queryKeywords':
        return this.searchRecievedProperties.QueryKeywords;
      case 'defaultQueryKeywords':
        return this.searchRecievedProperties.SearchBoxWebPart;
      case 'termData':
        return this.searchRecievedProperties.TermData;
    }

    throw new Error('Bad property id');
  }
}
