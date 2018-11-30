import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AnalyticsSettingsButton, IAnalyticsSettingsButtonProps } from './components/AnalyticsSettingsButton';
import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

const LOG_SOURCE: string = 'AnalyticsApplicationCustomizer';

export interface IAnalyticsApplicationCustomizerProperties {
  googleTrackingId: string;
}

export default class AnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public async onInit(): Promise<void> {
    let googleTrackingId: string = this.properties.googleTrackingId;

    //Wire up settings button if the user is a site owner
    if (this.checkIfCurrentUserIsSiteOwner()) {
      this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceHolders);
    }

    //If we have a google tracking id, setup ga on page, otherwise log warning
    if (googleTrackingId) {
      console.log(LOG_SOURCE, `Google Analytics Tracking ID: ${googleTrackingId}`);
      this.setupGoogleAnalytics(googleTrackingId);
    } else {
      console.warn(LOG_SOURCE, "Google Analytics Tracking ID: Not Specified");
    }
  }

  private renderPlaceHolders(): void {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);

      if (this._bottomPlaceholder.domElement) {
        const analyticsSettingsButton: React.ReactElement<IAnalyticsSettingsButtonProps> = React.createElement(AnalyticsSettingsButton, {
          googleTrackingId: this.properties.googleTrackingId,
          onSaveFunction: this.updateGoogleTrackingId.bind(this)
        } as IAnalyticsSettingsButtonProps);
        ReactDOM.render(analyticsSettingsButton, this._bottomPlaceholder.domElement);
      }
    }
  }

  private checkIfCurrentUserIsSiteOwner(): boolean {
    return this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
  }

  private setupGoogleAnalytics(googleTrackingId: string): void {
    //Add Google Analytics Script Tag to Page
    var gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${googleTrackingId}`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);

    //Invoke Google Analytics Page Tracker
    window["dataLayer"] = window["dataLayer"] || [];
    window["gtag"] = window["gtag"] || function gtag() {
      window["dataLayer"].push(arguments);
    };
    window["gtag"]('js', new Date());
    window["gtag"]('config', googleTrackingId);
  }

  private async updateGoogleTrackingId(googleTrackingId: string): Promise<void> {
    let customAction = await this.getCustomActionByComponentId(this.manifest.id);

    try {
      let postCustomActionUri = `${this.context.pageContext.web.absoluteUrl}/_api/web/usercustomactions('${customAction.Id}')`;
      await this.context.spHttpClient.post(postCustomActionUri, SPHttpClient.configurations.v1, {
        headers: {
          "X-HTTP-Method": "MERGE",
          "content-type": "application/json; odata=nometadata"
        },
        body: JSON.stringify({
          "ClientSideComponentProperties": `{ "googleTrackingId": "${googleTrackingId}" }`
        })
      });
    } catch (error) {
      console.log(`ERROR: Unable to update custom action with id ${customAction.id}`, error);
    }
  }

  private async getCustomActionByComponentId(componentId: string): Promise<any> {
    try {
      let getCustomActionUri = `${this.context.pageContext.web.absoluteUrl}/_api/web/usercustomactions`;
      getCustomActionUri += `?$filter=ClientSideComponentId eq guid'${componentId}'`;
      let getCustomActionResponse: SPHttpClientResponse = await this.context.spHttpClient.get(getCustomActionUri, SPHttpClient.configurations.v1);
      let getCustomActionResult = await getCustomActionResponse.json();
      return getCustomActionResult.value && getCustomActionResult.value.length > 0 ? getCustomActionResult.value[0] : null;
    }
    catch (error) {
      console.log(`ERROR: Unable to fetch custom action with ClientSideComponentId ${componentId}`, error);
    }
  }
}
