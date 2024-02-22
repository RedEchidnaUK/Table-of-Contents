declare interface ITableOfContentsWebPartStrings {
  searchWebpartsLabel: string;
  searchText: string;
  searchMarkdown: string;
  searchCollapsible: string;
  propertyPaneDescription: string;
  titleDefaultValue: string;
  titleFieldDescription: string;
  hideTitleFieldLabel: string;
  showHeadingLevelsLabel: string;
  showHeading1FieldLabel: string;
  showHeading2FieldLabel: string;
  showHeading3FieldLabel: string;
  showHeading4FieldLabel: string;
  listStyle: string;
  showPreviousPageViewLabel: string;
  previousPageFieldLabel: string;
  previousPageDefaultValue: string;
  showPreviousPageTitleLabel: string;
  showPreviousPageAboveLabel: string;
  showPreviousPageBelowLabel: string;
  enableStickyModeLabel: string;
  enableStickyModeDescription: string;
  hideInMobileViewLabel: string;
  errorToggleFieldEmpty: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
