import * as React from "react";
import {
  Customizer,
  INavLink,
  INavLinkGroup,
  INavStyles,
  Label,
  Link,
  mergeStyleSets,
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

const { allItens } = useList();
const { currentUser } = useUsers();

export const ReactDiscussionBoard: React.FunctionComponent<IReactDiscussionBoardProps> = (
  props: IReactDiscussionBoardProps
) => {
  const [state, setState] = React.useState<IReactDiscussionBoardState>({
    isLoading: false,
    hasError: false,
    errorMessage: "",
    listName: props.listname,
    siteurl: props.siteurl,
    itens: [],
  });

  const stateRef = React.useRef(state); // Use to access state on eventListenners

  React.useEffect(() => {
    (async () => {
      if (!props.siteurl || !props.listname) {
        return;
      }
      try {
        let _discussionItens: any = [];
        stateRef.current = {
          ...stateRef.current,
          isLoading: true,
          itens: _discussionItens,
        };
        setState(stateRef.current);
        const _currentUser = await currentUser();
        _discussionItens = await allItens(
          props.siteurl,
          props.listname,
          "Recentes",
          20,
          _currentUser.Id
        );
        stateRef.current = {
          ...stateRef.current,
          hasError: false,
          errorMessage: "",
          isLoading: false,
          itens: _discussionItens,
        };
        setState(stateRef.current);
      } catch (error) {
        stateRef.current = {
          ...stateRef.current,
          hasError: true,
          errorMessage: error.message,
        };
        setState(stateRef.current);
      }
    })();
  }, [props.siteurl, props.listname]);

  if (state.hasError) {
    return (
      <>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {state.errorMessage}
        </MessageBar>
      </>
    );
  }

  return (
    <>
      <Customizer settings={{ siteurl: props.siteurl }}>
        {state.isLoading ? (
          <Stack horizontal horizontalAlign="center">
            <Spinner size={SpinnerSize.medium}></Spinner>
          </Stack>
        ) : (
          <>
            <Stack
              horizontalAlign="space-between"
              horizontal
              tokens={{ childrenGap: 10 }}
              style={{ width: "100%" }}
            >
              <div>{props.description}</div>
            </Stack>
            {state.itens?.length === 0 ? (
              <Label>{state.itens}</Label>
            ) : (
              <Label>{state.itens}</Label>
            )}
          </>
        )}
      </Customizer>
    </>
  );
};
