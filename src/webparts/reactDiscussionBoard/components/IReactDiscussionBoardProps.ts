import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IReactDiscussionBoardProps {
  description: string;
  siteurl: string;
  listname:string;
  context: WebPartContext;
}
