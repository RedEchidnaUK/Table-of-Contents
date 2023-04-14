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
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

import * as strings from 'TableOfContentsWebPartStrings';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps } from './components/ITableOfContentsProps';

export interface ITableOfContentsWebPartProps {
  //title: string;
  hideTitle: boolean;
  titleText: string;
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

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    // Consume the ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();
    this.setCSSVariables(this._themeVariant.semanticColors);
    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    return super.onInit();
  }

  private setCSSVariables(theming: any): any {
    if (!theming) { return null; }
    let themingKeys = Object.keys(theming);
    if (themingKeys !== null) {
      themingKeys.forEach(key => {
        this.domElement.style.setProperty(`--${key}`, theming[key]);
      });
    }
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.setCSSVariables(this._themeVariant.semanticColors);
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents,
      {

        themeVariant: this._themeVariant,

        hideTitle: this.properties.hideTitle,
        titleText: this.properties.titleText,

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
  /*private handleUpdateProperty = (newValue: string) => {
    this.properties.title = newValue;
  }*/

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
                PropertyPaneToggle('hideTitle', {
                  label: strings.hideTitleFieldLabel
                }),
                PropertyPaneTextField('titleText', {
                  description: strings.titleFieldDescription,
                  disabled: this.properties.hideTitle,
                  onGetErrorMessage: this.checkToggleField,
                  value: strings.titleDefaultValue
                })
              ]
            },
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
    if (value === "") {
      return strings.errorToggleFieldEmpty;
    }
    else {
      return "";
    }
  }

}
