declare interface IRecentDocumentStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'recentDocumentStrings' {
  const strings: IRecentDocumentStrings;
  export = strings;
}
