import * as React from "react";
import {
    DocumentCard,
    DocumentCardDetails,
    Text,
    format as formatString,
    IDocumentCardStyles,
    mergeStyles,
    mergeStyleSets,
    Stack,
} from "office-ui-fabric-react";

import { Person } from "@microsoft/mgt-react";

import {
    format,
    parseISO
} from "date-fns";

import { currentSiteTheme } from "../../../common/constants";
import jquery from "jquery";

import { IDiscussionBoardCardsProps } from "./IDiscussionBoardCardsProps";
import { IDiscussionBoardCardsState } from "./IDiscussionBoardCardsState";

const tileStyle: IDocumentCardStyles = {
    root: {
        minHeight: 173,
        boxShadow: "0 5px 15px rgba(50, 50, 90, .1)",
    },
};

const componentClasses = mergeStyleSets({
    eventDay: mergeStyles({
        position: "relative",
        left: -5,
        paddingLeft: 20,
        verticalAlign: "center",
        borderLeftWidth: 4,
        borderLeftColor: currentSiteTheme.theme.themePrimary,
        borderLeftStyle: "solid",
    }),
});


export class DiscussionBoardCards extends React.Component<IDiscussionBoardCardsProps, IDiscussionBoardCardsState> {

    constructor(props: IDiscussionBoardCardsProps) {
        super(props);
        console.log(props);
    }
    public async componentDidMount() {
    }
    public componentDidUpdate(prevProps: IDiscussionBoardCardsProps, prevState: IDiscussionBoardCardsState): void {
    }

    //
    public render(): React.ReactElement<IDiscussionBoardCardsProps> {
        const discussionBoardcss = {
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(min(100%, 220px), 1fr))",
            gridGap: "1rem"
        };
        return (
            <div style={discussionBoardcss}>
                {
                    this.props.itens.results.map((item: any) => {
                        return (
                            <DocumentCard styles={tileStyle} data-interception="false" onClickTarget="_blank" onClickHref={'discussionboard.aspx?pid=' + item.Id}>
                                <Stack
                                    horizontal
                                    horizontalAlign="start"
                                    verticalAlign="center"
                                    tokens={{ childrenGap: 10 }}
                                    styles={{ root: { marginTop: 20 } }}
                                >
                                    <div className={componentClasses.eventDay}>
                                        <Person
                                            userId="vinaunder@devvina.onmicrosoft.com"
                                            showPresence={true}
                                        ></Person>
                                    </div>
                                    <Stack>
                                        <Text
                                            styles={{ root: { fontWeight: 600 } }}
                                            variant="mediumPlus"
                                            block
                                            nowrap
                                        >
                                            {item.Title}
                                        </Text>
                                        <Text variant="small">
                                            {formatString(
                                                "by {0}",
                                                "Vinicius Costa"
                                            )}
                                        </Text>
                                    </Stack>
                                </Stack>
                                <DocumentCardDetails
                                    styles={{ root: { height: "60px" } }}
                                >
                                    <Text
                                        styles={{ root: { fontWeight: 600, margin: "20px" } }}
                                        variant="mediumPlus"
                                        block
                                        nowrap
                                        title={item.Body.replace(/<[^>]+>/g, '')}
                                    >
                                        {item.Body.replace(/<[^>]+>/g, '')}
                                    </Text>
                                </DocumentCardDetails>
                                <Stack styles={{ root: { position: "relative", padding: 20 } }}>
                                    <Text variant="small">
                                        {formatString(
                                            "Last Reply at {0}",
                                            format(parseISO(item.DiscussionLastUpdated), "p")
                                        )}
                                    </Text>
                                </Stack>
                                <Stack
                                    tokens={{ childrenGap: 5 }}
                                    horizontal
                                    horizontalAlign="start"
                                    wrap
                                    styles={{ root: { paddingLeft: 20, paddingRight: 20 } }}
                                >
                                    {/* <MgtPersons attendees={event.attendees} /> */}
                                </Stack>
                            </DocumentCard>
                        );
                    })
                }
            </div>
        );
    }
}
export default DiscussionBoardCards;