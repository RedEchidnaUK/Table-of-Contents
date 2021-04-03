declare interface ITableOfContentsWebPartStrings {
  propertyPaneDescription: string;
  titleDefaultValue: string;
  titleFieldDescription: string;
  hideTitleFieldLabel: string;
  showHeading1FieldLabel: string;
  showHeading2FieldLabel: string;
  showHeading3FieldLabel: string;
  showPreviousPageViewLabel: string;
  previousPageFieldLabel: string;
  previousPageFieldDescription: string;
  previousPageDefaultValue: string;
  enableStickyModeLabel: string;
  enableStickyModeDescription: string;
  hideInMobileViewLabel: string;
  errorToggleFieldEmpty: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
