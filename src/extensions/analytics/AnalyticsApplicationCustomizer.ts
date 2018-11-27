import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AnalyticsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AnalyticsApplicationCustomizer';

export interface IAnalyticsApplicationCustomizerProperties { }

export default class AnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    let trackingID: string = "UA-129865562-1";

    if (!trackingID) {
      Log.info(LOG_SOURCE, "Google Analytics Tracking ID: Not Specified");
    } else {
      Log.info(LOG_SOURCE, `Google Analytics Tracking ID: ${trackingID}`);

      //Add Google Analytics Script Tag to Page
      var gtagScript = document.createElement("script");
      gtagScript.type = "text/javascript";
      gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${trackingID}`;
      gtagScript.async = true;
      document.head.appendChild(gtagScript);

      //Invoke Google Analytics Page Tracker
      window["dataLayer"] = window["dataLayer"] || [];
      window["gtag"] = window["gtag"] || function gtag() {
        window["dataLayer"].push(arguments);
      }
      window["gtag"]('js', new Date());
      window["gtag"]('config', trackingID);
    }

    return Promise.resolve();
  }
}
