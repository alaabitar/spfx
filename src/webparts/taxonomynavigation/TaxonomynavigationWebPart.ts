import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TaxonomynavigationWebPartStrings';
import Taxonomynavigation from './components/Taxonomynavigation';
import { ITaxonomynavigationProps } from './components/ITaxonomynavigationProps';

export interface ITaxonomynavigationWebPartProps {
  description: string;
}

export default class TaxonomynavigationWebPart extends BaseClientSideWebPart<ITaxonomynavigationWebPartProps> {
  
  public render(): void {
    const element: React.ReactElement<ITaxonomynavigationProps > = React.createElement(
      Taxonomynavigation,
      {
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl
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
