import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneDropdown
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
  hideTitle: boolean;
  titleText: string;
  searchText: boolean;
  searchMarkdown: boolean;
  searchCollapsible: boolean;
  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;
  showHeading4: boolean;
  showPreviousPageLinkTitle: boolean;
  showPreviousPageLinkAbove: boolean;
  showPreviousPageLinkBelow: boolean;
  previousPageText: string;
  historyCount: number;
  enableStickyMode: boolean;
  hideInMobileView: boolean;
  listStyle: string;

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
    // return super.onInit()
    return super.onInit().then(_ => {
      if (this.properties.searchText === undefined) {
        this.properties.searchText = true;
        this.properties.showHeading4 = true;
      }
    });
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

        searchText: this.properties.searchText,
        searchMarkdown: this.properties.searchMarkdown,
        searchCollapsible: this.properties.searchCollapsible,

        showHeading2: this.properties.showHeading1,
        showHeading3: this.properties.showHeading2,
        showHeading4: this.properties.showHeading3,
        showHeading5: this.properties.showHeading4,

        showPreviousPageLinkTitle: this.properties.showPreviousPageLinkTitle,
        showPreviousPageLinkAbove: this.properties.showPreviousPageLinkAbove,
        showPreviousPageLinkBelow: this.properties.showPreviousPageLinkBelow,
        previousPageText: this.properties.previousPageText,

        enableStickyMode: this.properties.enableStickyMode,
        webpartId: this.context.instanceId,

        hideInMobileView: this.properties.hideInMobileView,

        listStyle: this.properties.listStyle,
        isEditMode: this.displayMode == DisplayMode.Edit,
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
    let showHeading4: any;
    let showPreviousPageLinkTitle: any;

    if (this.properties.searchMarkdown) {
      showHeading4 = PropertyPaneCheckbox('showHeading4', {
        text: strings.showHeading4FieldLabel
      })
    }
    else {
      showHeading4 = PropertyPaneCheckbox('showHeading4', {
        text: strings.showHeading4FieldLabel,
        disabled: true
      });
    }

    if (this.properties.hideTitle) {
      showPreviousPageLinkTitle = PropertyPaneCheckbox('showPreviousPageLinkTitle', {
        text: strings.showPreviousPageTitleLabel,
        disabled: true
      })
    }
    else {
      showPreviousPageLinkTitle = PropertyPaneCheckbox('showPreviousPageLinkTitle', {
        text: strings.showPreviousPageTitleLabel
      });
    }

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
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('searchWebpartsLabel', {
                  text: strings.searchWebpartsLabel
                }),
                PropertyPaneCheckbox('searchText', {
                  text: strings.searchText,
                }),
                PropertyPaneCheckbox('searchMarkdown', {
                  text: strings.searchMarkdown
                }),
                PropertyPaneCheckbox('searchCollapsible', {
                  text: strings.searchCollapsible
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('showHeadingLevelsLabel', {
                  text: strings.showHeadingLevelsLabel
                }),
                PropertyPaneCheckbox('showHeading1', {
                  text: strings.showHeading1FieldLabel
                }),
                PropertyPaneCheckbox('showHeading2', {
                  text: strings.showHeading2FieldLabel
                }),
                PropertyPaneCheckbox('showHeading3', {
                  text: strings.showHeading3FieldLabel
                }),
                showHeading4,
                PropertyPaneDropdown('listStyle', {
                  label: strings.listStyle,
                  options: [
                    { key: 'default', text: 'Default' },
                    { key: 'disc', text: 'Disc' },
                    { key: 'circle', text: 'Circle' },
                    { key: 'square', text: 'Square' },
                    { key: 'none', text: 'None' }
                  ],
                  selectedKey: "default"
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('previousPageLabel', {
                  text: strings.showPreviousPageViewLabel
                }),
                showPreviousPageLinkTitle,
                PropertyPaneCheckbox('showPreviousPageLinkAbove', {
                  text: strings.showPreviousPageAboveLabel
                }),
                PropertyPaneCheckbox('showPreviousPageLinkBelow', {
                  text: strings.showPreviousPageBelowLabel
                }),
                PropertyPaneTextField('previousPageText', {
                  label: strings.previousPageFieldLabel,
                  disabled: (!this.properties.showPreviousPageLinkTitle || this.properties.hideTitle) && !this.properties.showPreviousPageLinkAbove && !this.properties.showPreviousPageLinkBelow,
                  onGetErrorMessage: this.checkToggleField,
                  value: strings.previousPageDefaultValue
                }),
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
