import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { ApplicationInsights, DistributedTracingModes } from '@microsoft/applicationinsights-web';

import * as strings from 'AzureAppInsightsWebPartStrings';
import { ITelemetryItem } from '@microsoft/applicationinsights-web';

export interface IAzureAppInsightsWebPartProps {
  enabled: boolean;
  instrumentationKey: string;
  trackUserId: boolean;
  trackExceptions: boolean;
  cloudRole: string;
  cloudRoleInstance: string;
  excludedDependencyTargets: string;
}

const DEFAULT_PROPERTIES: IAzureAppInsightsWebPartProps = {
  enabled: false,
  instrumentationKey: "00000000-0000-0000-0000-000000000000",
  trackUserId: false,
  trackExceptions: true,
  cloudRole: "sharepoint-page",
  cloudRoleInstance: "",
  excludedDependencyTargets: 'browser.pipe.aria.microsoft.com\nbusiness.bing.com\nmeasure.office.com\nofficeapps.live.com\noutlook.office365.com\noutlook.office.com'
};

export default class AzureAppInsightsWebPart extends BaseClientSideWebPart<IAzureAppInsightsWebPartProps> {
  private _appInsights: ApplicationInsights = undefined;
  private _config: IAzureAppInsightsWebPartProps = undefined;
  private _excludedDependencies: string[] = [];

  private _appInsightsInitializer = (telemetryItem: ITelemetryItem): boolean | void => {
    if (telemetryItem) {
      if (!!this._config.cloudRole) {
        telemetryItem.tags['ai.cloud.role'] = this._config.cloudRole;
      }
      if (!!this._config.cloudRoleInstance) {
        telemetryItem.tags['ai.cloud.roleInstance'] = this._config.cloudRoleInstance;
      }
      if (telemetryItem.baseType === 'RemoteDependencyData' && telemetryItem.baseData.target && telemetryItem.baseData.target) {
        const isExcluded = this._excludedDependencies.some(target => telemetryItem.baseData.target.toLowerCase().indexOf(target.toLowerCase()) !== -1);
        if (isExcluded) return false; // don't track
      }
    }
  }

  public async onInit(): Promise<void> {
    this._config = {
      ...DEFAULT_PROPERTIES,
      ...this.properties
    };

    if (!this._config.enabled) {
      console.info(`[AzureAppInsights Web Part] Tracking disabled. No analytics will be logged.`);
      return;
    }

    if (!this._config.instrumentationKey || this._config.instrumentationKey === DEFAULT_PROPERTIES.instrumentationKey) {
      console.warn(`[AzureAppInsights Web Part] Instrumentation key not provided. No analytics will be logged.`);
      return;
    }

    if (!!this._config.excludedDependencyTargets) {
      this._excludedDependencies = this._config.excludedDependencyTargets.split('\n').filter(n => n && !!n.trim() && n.trim().length > 5);
    }

    const userId: string = this._config.trackUserId ? this.context.pageContext.user.loginName.replace(/([\|:;=])/g, '') : undefined;

    // App Insights JS Documentation: https://github.com/microsoft/applicationinsights-js
    this._appInsights = new ApplicationInsights({
      config: {
        // Instrumentation key that you obtained from the Azure Portal.
        instrumentationKey: this._config.instrumentationKey,

        // An optional account id, if your app groups users into accounts. No spaces, commas, semicolons, equals, or vertical bars
        accountId: userId,

        // If true, Fetch requests are not autocollected. Default is true
        disableFetchTracking: false,

        // If true, AJAX & Fetch request headers is tracked, default is false.
        enableRequestHeaderTracking: true,

        // If true, AJAX & Fetch request's response headers is tracked, default is false.
        enableResponseHeaderTracking: true,

        // Default false. If true, include response error data text in dependency event on failed AJAX requests.
        enableAjaxErrorStatusText: true,

        // Default false. Flag to enable looking up and including additional browser window.performance timings in the reported ajax (XHR and fetch) reported metrics.
        enableAjaxPerfTracking: true,

        // If true, unhandled promise rejections will be autocollected and reported as a javascript error. When disableExceptionTracking is true (dont track exceptions) the config value will be ignored and unhandled promise rejections will not be reported.
        enableUnhandledPromiseRejectionTracking: true,

        // If true, the SDK will add two headers ('Request-Id' and 'Request-Context') to all CORS requests tocorrelate outgoing AJAX dependencies with corresponding requests on the server side. Default is false
        enableCorsCorrelation: true,

        // If true, exceptions are not autocollected. Default is false.
        disableExceptionTracking: !this._config.trackExceptions,

        // Sets the distributed tracing mode. If AI_AND_W3C mode or W3C mode is set, W3C trace context headers (traceparent/tracestate) will be generated and included in all outgoing requests. AI_AND_W3C is provided for back-compatibility with any legacy Application Insights instrumented services.
        distributedTracingMode: DistributedTracingModes.AI_AND_W3C
      }
    });

    this._appInsights.loadAppInsights();
    this._appInsights.addTelemetryInitializer(this._appInsightsInitializer);
    this._appInsights.setAuthenticatedUserContext(userId, userId, true);
    this._appInsights.trackPageView();
  }

  public render(): void {
    if (this.displayMode === DisplayMode.Edit) {
      this.domElement.innerHTML = `
        <div>
          <h3>Azure App Insights</h3>
          <div>
            To edit the settings for this web part, click the edit pencil. This message is only visible while the page is in edit mode.
          </div>
        </div>
      `;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('enabled', {
                  label: strings.EnabledFieldLabel
                }),
                PropertyPaneTextField('instrumentationKey', {
                  label: strings.InstrumentationKeyFieldLabel
                }),
                PropertyPaneToggle('trackUserId', {
                  label: strings.TrackUserIdFieldLabel,
                }),
                PropertyPaneLabel('trackUserId', {
                  text: strings.TrackUserIdFieldDescription
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('trackExceptions', {
                  label: strings.TrackExceptionsFieldLabel
                }),
                PropertyPaneTextField('excludedDependencyTargets', {
                  label: strings.ExcludedDependencyTargetsFieldLabel,
                  description: strings.ExcludedDependencyTargetsFieldDescription,
                  multiline: true,
                  resizable: true,
                  rows: 10
                }),
                PropertyPaneTextField('cloudRole', {
                  label: strings.CloudRoleFieldLabel,
                  placeholder: strings.OptionalTextPlaceholder
                }),
                PropertyPaneTextField('cloudRoleInstance', {
                  label: strings.CloudRoleInstanceFieldLabel,
                  placeholder: strings.OptionalTextPlaceholder
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
