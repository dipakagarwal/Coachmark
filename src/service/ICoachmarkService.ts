import { Guid } from "@microsoft/sp-core-library";

export interface ICoachmarkService {
    getCoachmark(sPHttpClient: any, baseUrl: string, webId: Guid | any): any;
    acknowledgeCoachmark(id: number, webId: Guid | any, version: number): void;
}