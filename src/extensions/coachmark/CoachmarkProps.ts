import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { ICoachmark } from "../../service/ICoachmark";

export interface ICoachmarkProps {
    context: ApplicationCustomizerContext;
    applicationInsightsKey: string;
    eventName: string;
    data: ICoachmark[];
    serviceProvider: any;
    applicationInsightsProvider: any;
    pageRelativeURL: string;
}