import * as React from "react";
import {
  Text,
  Customizer,
  Label,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
} from "office-ui-fabric-react";
import { useList } from "../../../hooks/useLists";
import { useUsers } from "../../../hooks/useUsers";

import { IReactDiscussionBoardProps } from "./IReactDiscussionBoardProps";
import { IReactDiscussionBoardState } from "./IReactDiscussionBoardState";

import { DiscussionBoardCards } from "./DiscussionBoardCards";

const { allItens } = useList();
const { currentUser } = useUsers();


export default class ReactDiscussionBoard extends React.Component<IReactDiscussionBoardProps, IReactDiscussionBoardState> {
  private _discussionItens: any[] = [];
  constructor(props: IReactDiscussionBoardProps) {
    super(props);
    this.state = {
      isLoading: false,
      hasError: false,
      errorMessage: "",
      DiscussionBoards: [],
    };
  }

  public componentDidMount(): void {
    this.GetDiscussionBoards();
  }

  public componentDidUpdate(prevProps: IReactDiscussionBoardProps, prevState: IReactDiscussionBoardState): void {

  }
  // Render
  public render(): React.ReactElement<IReactDiscussionBoardProps> {
    let _center: any = !this.state.isLoading ? "center" : "";
    return (
      <div style={{ textAlign: _center }} >
        <div>
          <Text variant="xLarge">{this.props.description}</Text>
          {
            !this.state.isLoading ?
              <div>
                <Stack horizontal horizontalAlign="center">
                  <Spinner size={SpinnerSize.medium}></Spinner>
                </Stack>
              </div>
              :
              <DiscussionBoardCards itens={this.state.DiscussionBoards}/>
          }
        </div>
      </div>
    );
  }

  private async GetDiscussionBoards() {
    const _currentUser = await currentUser();
    const listItems = await allItens(
      this.props.listname,
      "Recentes",
      20,
      _currentUser.Id
    );
    if (listItems && listItems.results.length > 0) {
      this._discussionItens = listItems;
    }

    this.setState(
      {
        isLoading: true,
        hasError: false,
        errorMessage: "",
        DiscussionBoards: this._discussionItens,
      });
  }
}