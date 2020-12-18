define([], function() {
  return {
    "PropertyPaneDescription": "Configure the Azure App Insights web part.",
    "BasicGroupName": "Settings",
    "AdvancedGroupName": "Advanced",
    "EnabledFieldLabel": "Tracking Enabled",
    "InstrumentationKeyFieldLabel": "Instrumentation Key",
    "TrackUserIdFieldLabel": "Include User Login in Analytics Data",
    "TrackUserIdFieldDescription": "If enabled, the user's login name will be sent with each telemetry item in the format of user@domain.com.",
    "TrackExceptionsFieldLabel": "Include Errors in Analytics Data",
    "CloudRoleFieldLabel": "Custom Cloud Role Identifier",
    "CloudRoleInstanceFieldLabel": "Custom Cloud Role Instance Identifier",
    "ExcludedDependencyTargetsFieldLabel": "Excluded Dependency Targets List",
    "ExcludedDependencyTargetsFieldDescription": "Add a dependency target on each line to have it excluded from dependency tracking. This is useful to exclude certain SharePoint Online services which may generate a lot of noise. Partial host matching is supported. Each line must be at least 5 characters long (ex: a.com).",
    "OptionalTextPlaceholder": "(Optional) Leave blank if not needed"
  }
});
