import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import { 
  IPropertyPaneConfiguration, 
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneLabel
} from "@microsoft/sp-property-pane";

import * as strings from 'TableOfContentsWebPartStrings';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps } from './components/ITableOfContentsProps';

export interface ITableOfContentsWebPartProps {
  title: string;
  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;
  showPreviousPageLink: boolean;
  previousPageText: string;
  historyCount: number;
  enableStickyMode: boolean;
  hideInMobileView: boolean;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: this.handleUpdateProperty,

        showHeading2: this.properties.showHeading1,
        showHeading3: this.properties.showHeading2,
        showHeading4: this.properties.showHeading3,

        showPreviousPageLink: this.properties.showPreviousPageLink,
        previousPageText: this.properties.previousPageText,

        enableStickyMode: this.properties.enableStickyMode,
        webpartId: this.context.instanceId,

        hideInMobileView: this.properties.hideInMobileView,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Saves new value for the title property.
   */
  private handleUpdateProperty = (newValue: string) => {
    this.properties.title = newValue;
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
            description: strings.propertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneCheckbox('showHeading1', {
                  text: strings.showHeading1FieldLabel
                }),
                PropertyPaneCheckbox('showHeading2', {
                  text: strings.showHeading2FieldLabel
                }),
                PropertyPaneCheckbox('showHeading3', {
                  text: strings.showHeading3FieldLabel
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneToggle('showPreviousPageLink', {
                  label: strings.showPreviousPageViewLabel
                }),
                PropertyPaneTextField('previousPageText', {
                  //label: strings.PreviousPageFieldLabel,
                  description: strings.previousPageFieldDescription,
                  disabled: !this.properties.showPreviousPageLink,
                  onGetErrorMessage: this.checkToggleField,
                  value: strings.previousPageDefaultValue
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneToggle('enableStickyMode', {
                  label: strings.enableStickyModeLabel
                }),
                PropertyPaneLabel('enabldeStickyModeDescription', {
                  text: strings.enableStickyModeDescription
                }),
                PropertyPaneToggle('hideInMobileView', {
                  label: strings.hideInMobileViewLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private checkToggleField = (value: string): string => {
    if (value === ""){
      return strings.errorToggleFieldEmpty;
    }
    else {
      return "";
    }
  }

}
