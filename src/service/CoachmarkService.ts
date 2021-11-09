import { Guid } from "@microsoft/sp-core-library";
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

import { ICoachmark } from "./ICoachmark";
import { ICoachmarkService } from "./ICoachmarkService";

interface ICoachmarkStatus {
    Id: number;
    Ack: Date;
}

interface ICoachmarkCache {
    Loaded?: Date;
    Coachmark: ICoachmark[];
    CoachmarkStatuses: ICoachmarkStatus[];
}

/** Returns items from the Coachmark list and caches the results */
export class CoachmarkService implements ICoachmarkService {
    private static readonly storageKeyBase: string = 'spfxCoach'; //Key used for localStorage
    private static readonly getFromListAlways: boolean = true; //Useful for testing

    //***********************
    //Public Methods
    //***********************


    /* Retrieves Coachmark that should be displayed for the given user */
    public getCoachmark(spHttpClient: SPHttpClient, baseUrl: string, webId: Guid): Promise<ICoachmark[]> {
        return new Promise<ICoachmark[]>((resolve: (Coachmark: ICoachmark[]) => void, reject: (error: string) => void): void => {
            CoachmarkService.ensureCoachmark(spHttpClient, baseUrl, webId)
            .then((Coachmark: ICoachmark[]): void => {
                resolve(Coachmark);
            }).catch((error: string): void => {
                reject(error);
            });
        });
    }

    /** Stores the date/time a Coachmark was acknowledged, used to control what shows on the next refresh
     * @param {number} id - The list ID of the Coachmark to acknowledge
    */

    public acknowledgeCoachmark(id: number, webId: Guid): void {
        const cachedData: ICoachmarkCache = CoachmarkService.retrieveCache(webId);

        // Check if status already exists, and if so update it
        // otherwise, add a new status for the id
        const index: number = CoachmarkService.indexOfCoachmarkStatusById(id, cachedData.CoachmarkStatuses);
        if(index >= 0) {
            cachedData.CoachmarkStatuses[index].Ack = new Date();
        } else {
            cachedData.CoachmarkStatuses.push({
                Id: id,
                Ack: new Date()
            });
        }
        CoachmarkService.storeCache(cachedData, webId);
    }


    //***********************
    //localStorage Management
    //***********************

    private static webStorageKey(webId: Guid): string {
        return `${CoachmarkService.storageKeyBase}_${webId}`;
    }

    /** Rehydrates spfxCoach data from localStorage (or creates a new empty set) */
    private static retrieveCache(webId: Guid): ICoachmarkCache {
        //Pull data from local storage if available and we previously cached it
        let cachedData: ICoachmarkCache = localStorage ? JSON.parse(localStorage.getItem(this.webStorageKey(webId))) : undefined;
        if(cachedData) {
            cachedData.Loaded = new Date(cachedData.Loaded.valueOf()); //Rehydates data from JSON (serializes to string)
        } else {
            //Initialize a new, empty object
            cachedData = {
                Coachmark: [],
                CoachmarkStatuses: []
            };
        }
        return cachedData;
    }

    /** Serializes spfxCoach data into localStorage */
    private static storeCache(cachedData: ICoachmarkCache, webId: Guid): void {
        //Cache the data in localStorage when possible
        if(localStorage) {
            localStorage.setItem(this.webStorageKey(webId), JSON.stringify(cachedData));
        }
    }


    //*********************
    //Coachmark Retrieval
    //*********************

    /** Retrieves Coachmark from either cache or the list depending on the cache's freshness */
    private static ensureCoachmark(spHttpClient: SPHttpClient, baseUrl: string, webId: Guid): Promise<ICoachmark[]> {
        return new Promise<ICoachmark[]>((resolve: (Coachmark: ICoachmark[]) => void, reject: (error: any) => void): void => {
            
            let cachedData: ICoachmarkCache = CoachmarkService.retrieveCache(webId);

            if(cachedData.Loaded) {
                //True Cache found, cache if it is stale
                // anything older than 2 minutes will be considered stale
                let now: Date = new Date();
                let staleTime: Date = new Date(now.getTime() + -2 * 60000);

                if(cachedData.Loaded > staleTime && !CoachmarkService.getFromListAlways) {
                    //console.log('Pulled Coachmark from localStorage');
                    resolve(CoachmarkService.reduceCoachmark(cachedData));
                    return;
                }
            }

            if((window as any).spfxCoachmarkLoadingData) {
                //Coachmark are already being loaded! Briefly wait and try again
                window.setTimeout((): void => {
                    CoachmarkService.ensureCoachmark(spHttpClient, baseUrl, webId)
                    .then((Coachmark: ICoachmark[]): void => {
                        resolve(Coachmark);
                    });
                }, 100);
            } else {
                //Set a loading flag to prevent multiple data queries from firing
                // this will be important, should there be multiple consumers of the servies on a single page
                (window as any).spfxCoachmarkLoadingData = true;

                //Coachmark need to be loaded, so let's go get them!
                CoachmarkService.getCoachmarkFromList(spHttpClient, baseUrl)
                    .then((Coachmark: ICoachmark[]): void => {
                        //console.log('Pulled Coachmark from the list');
                        cachedData.Coachmark = Coachmark;
                        cachedData.Loaded = new Date();
                        cachedData = CoachmarkService.processCache(cachedData);

                        //Update the cache
                        CoachmarkService.storeCache(cachedData, webId);

                        //Clear the loading flag
                        (window as any).spfxCoachmarkLoadingData = false;

                        //Give them some Coachmark!
                        resolve(CoachmarkService.reduceCoachmark(cachedData));
                    }).catch((error: any): void => {
                        reject(error);
                    });
            }
        });
    }

    //Breaking up the URL like this isn't necessary, but can be easier to update
    private static readonly apiEndPoint: string = "_api/web/lists/getbytitle('Coachmarks')/items";
    private static readonly select: string = "Id,Title,Description,Frequency,Enabled,ElementID,PageList";
    private static readonly orderby: string = "StartDate asc";

    /** Pulls the active Coachmark entries directly from underlying list */
    private static getCoachmarkFromList(spHttpClient: SPHttpClient, baseUrl: string): Promise<ICoachmark[]> {
        //Coachmark are only shown during their schduled window
        let now: string = new Date().toISOString();
        let filter: string = `(StartDate le datetime'${now}') and (EndDate ge datetime'${now}')`;

        return spHttpClient.get(`${baseUrl}/${CoachmarkService.apiEndPoint}?$select=${CoachmarkService.select}&$filter=${filter}&$orderby=${CoachmarkService.orderby}&$top=5000`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse): Promise<{ value: ICoachmark[] }> => {
                if(!response.ok) {
                    //Failed requests don't automatically throw exceptions which
                    // can be problematic for chained promise, so we throw one
                    throw `Unable to get items: ${response.status} (${response.statusText})`;
                }
                return response.json();
            })
            .then((results: { value: ICoachmark[] }) => {
                //Clean up extra properties
                // Even when your interface only defines certain properties, SP sends many
                // extra properties that you may or may not care about (we don't)
                // (this isn't strictly necessary but makes the cache much cleaner)
                let Coachmark: ICoachmark[] = [];
                for (let v of results.value) {
                    Coachmark.push({
                        Title: v.Title,
                        Id: v.Id,
                        Frequency: v.Frequency,
                        Enabled: v.Enabled,
                        Description: v.Description,
                        ElementID: v.ElementID,
                        Visible: true,
                        PageList: v.PageList
                    });
                }
                return Coachmark;
            });
    }


    //************************
    //Helper Functions
    //************************

    /**Helper function to return the index of an ICoachmarkStatus object by the Id property */
    private static indexOfCoachmarkStatusById(Id: number, CoachmarkStatuses: ICoachmarkStatus[]): number {
        for (let i: number = 0; i < CoachmarkStatuses.length; i++) {
            if(CoachmarkStatuses[i].Id == Id) {
                return i;
            }
        }
        return -1;
    }

    //** Helper function to clean up Coachmark statuses by removing old Coachmark */
    private static processCache(cachedData: ICoachmarkCache): ICoachmarkCache {
        //Setup a temporary array of Ids (makes the filtering easier)
        let activeIds: number[] = [];
        for(let Coachmark of cachedData.Coachmark) {
            activeIds.push(Coachmark.Id);
        }

        //only keep the status info for Coachmark that still matter (active)
        cachedData.CoachmarkStatuses = cachedData.CoachmarkStatuses.filter((value: ICoachmarkStatus): boolean => {
            return activeIds.indexOf(value.Id) >= 0;
        });

        return cachedData;
    }

    /** Adjusts the coachmark to display based on what the user has already acknowledged and Coachmark's frequency value */
    private static reduceCoachmark(cachedData: ICoachmarkCache): ICoachmark[] {
        return cachedData.Coachmark.filter((Coachmark: ICoachmark): boolean => {
            if(!Coachmark.Enabled) {
                //Disabled Coachmark are still queried so that their status isn't lost
                // however, they shouldn't be displayed
                return false;
            }

            let tsIndex: number = CoachmarkService.indexOfCoachmarkStatusById(Coachmark.Id, cachedData.CoachmarkStatuses);
            if (tsIndex >= 0) {
                let lastShown: Date = new Date(cachedData.CoachmarkStatuses[tsIndex].Ack.valueOf());
                switch (Coachmark.Frequency) {
                    case 'Once':
                        //Already shown
                        return false;
                    case 'Always':
                        return true;
                    default:
                        //Default behaviour is Once per day
                        let now: Date = new Date();
                        if (now.getFullYear() !== lastShown.getFullYear()
                            || now.getMonth() !== lastShown.getMonth()
                            || now.getDay() !== lastShown.getDay()) {
                            //Last shown on a different day, so show it!
                            return true;
                        } else {
                            //Already shown today
                            return false;
                        }
                }
            } else {
                //No previous status means it needs to be shown
                return true;
            }
        });
    }




}