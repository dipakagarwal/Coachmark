export default interface IApplicationInsights {
    TRACKAPPLICATIONINSIGHTSLOG(applicationInsightsKey: string, elementTitle: string, elementID: string, userEmail: string, eventName: string);
}