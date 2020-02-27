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
  IWebPartPropertiesMetadata
} from '@microsoft/sp-webpart-base';

import PgpNavigationDetails from './components/PgpNavigationDetails';
import { IPgpNavigationDetailsProps } from './components/IPgpNavigationDetailsProps';
import { DynamicProperty } from '@microsoft/sp-component-base';
import { ITerm } from '@pnp/sp-taxonomy';
import * as strings from 'PgpNavigationDetailsWebPartStrings';

export interface IPgpNavigationDetailsWebPartProps {
  termData:DynamicProperty<ITerm>;
  description: string;
}

export default class PgpNavigationDetailsWebPart extends BaseClientSideWebPart<IPgpNavigationDetailsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPgpNavigationDetailsProps > = React.createElement(
      PgpNavigationDetails,
      {
        termData:this.properties.termData.tryGetValue(),
        description: this.properties.description
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
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'termData': {
        dynamicPropertyType: 'string'
      }
    };
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
              primaryGroup:{
                groupName:strings.BasicGroupName,
                groupFields:[]
              },
              secondaryGroup:{
              groupName: "Connect to Source",
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'termData',
                  fields: [
                    PropertyPaneDynamicField('termData', {
                      label: "Select a source"
                    })
                  ]
                })
              ],
              sharedConfiguration: {
                depth: DynamicDataSharedDepth.Property
              }
            },showSecondaryGroup: !!this.properties.termData.tryGetSource()
          }as IPropertyPaneConditionalGroup
          ]
        }
      ]
    };
  }
}
