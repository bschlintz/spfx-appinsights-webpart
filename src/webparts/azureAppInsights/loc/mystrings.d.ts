declare interface IAzureAppInsightsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AdvancedGroupName: string;
  EnabledFieldLabel: string;
  InstrumentationKeyFieldLabel: string;
  TrackUserIdFieldLabel: string;
  TrackUserIdFieldDescription: string;
  TrackExceptionsFieldLabel: string;
  CloudRoleFieldLabel: string;
  CloudRoleInstanceFieldLabel: string;
  ExcludedDependencyTargetsFieldLabel: string;
  ExcludedDependencyTargetsFieldDescription: string;
  OptionalTextPlaceholder: string;
}

declare module 'AzureAppInsightsWebPartStrings' {
  const strings: IAzureAppInsightsWebPartStrings;
  export = strings;
}
