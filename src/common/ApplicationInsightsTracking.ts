import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import IApplicationInsights from './IApplicationInsights';


export default class ApplicationInsightsTracking implements IApplicationInsights {
    /**
     * Logs the searched link in Azure application Insights
     * 
     * @param applicationInsightsKey Instrumentation key for Application Insights resource
     * @param eventType Name for tracked evenet
     * @param elementID Element for which Coachmark is displayed
     * @param userEmail Email ID of the person whoe is searching
     * @param pageURL Page from which Coachmark activity is noted
     * @param eventName Event name for application insights 
     */
    public TRACKAPPLICATIONINSIGHTSLOG(applicationInsightsKey: string, elementTitle: string, elementID: string, userEmail: string, eventName: string): void {
        try {
            const appInsightsKey: string = applicationInsightsKey;
            const appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: appInsightsKey,
                    disableExceptionTracking: true,
                    disableAjaxTracking: true
                }
            });
            appInsights.loadAppInsights();
            appInsights.setAuthenticatedUserContext(userEmail);

            appInsights.trackEvent({
                name: eventName,
                properties: {
                    Title: elementTitle,
                    ElementID: elementID
                }
            });
            appInsights.flush();
        } catch (error) { /** */ }
    }
}