import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'CoachmarkApplicationCustomizerStrings';

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { assign } from '@uifabric/utilities';
import { CoachmarkBubble } from './CoachmarkBubble';

import { ICoachmark } from '../../service/ICoachmark';
import { CoachmarkService } from '../../service/CoachmarkService';
// import ApplicationInsightsTracking from "../../common/ApplicationInsightsTracking";

const LOG_SOURCE: string = 'CoachmarkApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICoachmarkApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  applicationInsightsKey: string;
  eventName: string;
}
const delay = ms => new Promise(res => setTimeout(res, ms));
/** A Custom Action which can be run during execution of a Client Side Application */
export default class CoachmarkApplicationCustomizer
  extends BaseApplicationCustomizer<ICoachmarkApplicationCustomizerProperties> {
    private CoachmarkPromise: Promise<ICoachmark[]>;



  @override
  public onInit(): Promise<void> {debugger;
    let MessageSourceWebUrl: string = this.context.pageContext.web.absoluteUrl;
    alert(MessageSourceWebUrl); 
    console.log(MessageSourceWebUrl);
    this.CoachmarkPromise = new CoachmarkService().getCoachmark(this.context.spHttpClient, MessageSourceWebUrl, this.context.pageContext.web.id);
    this.CoachmarkPromise.then((item: ICoachmark[]) => {
      if(item.length > 0) { //check if there is any coachmark
        let showCoach: HTMLElement = document.body.appendChild(document.createElement(`div`));
        const element: React.ReactElement<{}> = React.createElement(CoachmarkBubble
          , assign({
            context: this.context,
            data: item,
            pageRelativeURL: location.pathname.replace(this.context.pageContext.web.absoluteUrl,""),
            userEmail: this.context.pageContext.user.email,
            strings: strings,
            eventName: 'Coachmark',
            serviceProvider: new CoachmarkService()
            
          })
        );
        // applicationInsightsKey: this.getAppInsightsKey(),
        // applicationInsightsProvider: new ApplicationInsightsTracking()
        setTimeout(() => {ReactDom.render(element, showCoach); }, 5000);
        //Render Coachmark
      }
    });
    return Promise.resolve();
  }

  private getAppInsightsKey(): string {
    if (window && 'hostname' in window.location) {
      if (window.location.hostname.toLowerCase() === 'office365developer.sharepoint.com') {
        return ''; //PROD
      } else {
        return ''; //TEST
      }
    }
     return null;
  }
}