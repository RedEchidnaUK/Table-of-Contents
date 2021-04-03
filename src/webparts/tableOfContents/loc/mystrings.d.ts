declare interface ITableOfContentsWebPartStrings {
  propertyPaneDescription: string;
  showHeading1FieldLabel: string;
  showHeading2FieldLabel: string;
  showHeading3FieldLabel: string;
  showPreviousPageViewLabel: string;
  previousPageFieldLabel: string;
  previousPageFieldDescription: string;
  previousPageDefaultValue: string;
  hideInMobileViewLabel: string;
  errorToggleFieldEmpty: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
