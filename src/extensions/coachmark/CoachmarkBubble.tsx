import * as React from 'react';
import { Coachmark, TeachingBubbleContent } from 'office-ui-fabric-react';
import { ICoachmarkProps } from './CoachmarkProps';
import { initializeIcons } from 'office-ui-fabric-react';
import { ICoachmark } from '../../service/ICoachmark';
initializeIcons();

export interface ITeachingBubbleState {
    CoachState: ICoachmark[]; 
}


export class CoachmarkBubble extends React.Component<ICoachmarkProps, ITeachingBubbleState> {

    constructor(props: ICoachmarkProps) {
        super(props);
        this.close = this.close.bind(this);

        this.state = {
            CoachState: this.props.data
        };
    }

    public render(): JSX.Element {
        const { CoachState } = this.state;
        return (
            <div>
                {CoachState.map((elem) => {
                    return (elem.Visible && (elem.PageList === null || elem.PageList.split(";").filter((eachPage) => eachPage === this.props.pageRelativeURL).length > 0) ?
                            <div key={elem.Id}>
                                <Coachmark 
                                    target={document.querySelector("[" + elem.ElementID + "]") as HTMLElement}
                                >
                                    <TeachingBubbleContent
                                        onDismiss={() => this.close(elem.ElementID, elem.Id, elem.Title)}
                                        hasCloseIcon={true}
                                        closeButtonAriaLabel="Close"
                                        headline={elem.Title}
                                    >
                                        {elem.Description}
                                    </TeachingBubbleContent>
                                </Coachmark>
                            </div>
                        : null);
                })}
            </div>
        );
    }

    //Dismiss particular coachmark
    private close(element, id, title): any {
        const local = this.state.CoachState;
        local.some((arrayItem) => {
            if(arrayItem.ElementID == element) {
                arrayItem.Visible = false;
                return true;
            }
        });
        this.setState({
            CoachState: local
        });
        /* istanbul ignore next */
        // if(this.props.context !== null) {
        //     this.props.serviceProvider.acknowledgeCoachmark(id, this.props.context.pageContext.web.id); //Update Coachmark Acknowledgement
        //     this.props.applicationInsightsProvider.TRACKAPPLICATIONINSIGHTSLOG(this.props.applicationInsightsKey, title, element, this.props.context.pageContext.user.email, this.props.eventName);
        // }
        
    }

}