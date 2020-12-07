import { Web } from "@pnp/sp/webs";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/files";

import _ from "lodash";
// import moment from 'moment';

import { IItemAddResult } from "@pnp/sp/items";
import { MoveOperations } from "@pnp/sp/files";

export const useList = () => {
  // Run on useList hook
  (async () => {})();

  const addReply = async (
    siteurl: string,
    listname: string,
    fileref: string,
    itemid: number,
    msg: string
  ): Promise<boolean> => {
    const w = Web(siteurl);
    await w.lists
      .getByTitle(listname)
      .items.add({
        Body: msg, //message Body
        FileSystemObjectType: 0, //setto 0 to make sure Mesage Item
        ContentTypeId: "0x01070074BFFEE8358D3846943B526051CA8FA1", //set Message content type
        FileRef: fileref,
        FileDirRef: fileref,
        ParentItemID: itemid, //set Discussion item (topic) Id
      })
      .then(async (retorno: IItemAddResult) => {
        //move o item
        await w.lists
          .getByTitle(listname)
          .items.select("*,FileRef,FileDirRef") // FileRef is a discussion folder path
          .filter(`startswith(ContentTypeId, '0x0107')`)
          .orderBy("DiscussionLastUpdated", true)
          .getAll()
          .then(async (itens) => {
            var _dt = _.filter(itens, { Id: retorno.data.Id });
            var fileUrl = _dt[0].FileRef;
            var fileDirRef = _dt[0].FileDirRef;
            var moveFileUrl = fileUrl.replace(fileDirRef, fileref);
            w.getFileByServerRelativePath(fileUrl)
              .moveTo(moveFileUrl, MoveOperations.Overwrite)
              .then(() => {
                return true;
              })
              .catch((e) => {
                console.log(e);
                return false;
              });
          });
      });
    return true;
  };

  const addItem = async (
    siteurl: string,
    listname: string,
    titulo: string,
    msg: string
  ): Promise<boolean> => {
    const w = Web(siteurl);
    let retorno = await w.lists
      .getByTitle(listname)
      .items.add({
        Title: titulo,
        Body: msg, //message Body
        ContentTypeId: "0x012002007F48573909070A4A8B3F9444A54F2EC3", //set Message content type
      })
      .then(() => {
        return true;
      })
      .catch((e) => {
        console.log(e);
        return false;
      });
    return retorno;
  };

  const allItens = async (
    siteurl: string,
    listname: string,
    tipo: string,
    t: number,
    currentuser: number
  ): Promise<any> => {
    const w = Web(siteurl);
    let filter;
    if (tipo == "Recentes") {
      filter = `startswith(ContentTypeId, '0x0120')`;
    } else if (tipo == "Minhas Perguntas") {
      filter =
        `startswith(ContentTypeId, '0x0120') and AuthorId eq ` + currentuser;
    } else if (tipo == "Por Responder") {
      filter = `startswith(ContentTypeId, '0x0120') and IsFeatured eq 0`;
    }
    let retorno = await w.lists
      .getByTitle(listname)
      .items.select("*,FileRef") // FileRef is a discussion folder path
      .filter(filter)
      .orderBy("DiscussionLastUpdated", false)
      .top(t)
      .getPaged();
    return retorno;
  };

  const getItemByID = async (
    siteurl: string,
    listname: string,
    iditem: number
  ): Promise<any> => {
    let datasource: any;
    const w = Web(siteurl);
    await w.lists
      .getByTitle(listname)
      .items.select("*,FileRef") // FileRef is a discussion folder path
      .filter(`startswith(ContentTypeId, '0x0120')`)
      .orderBy("DiscussionLastUpdated", true)
      .getAll()
      .then(async (itens) => {
        let datasourceItens = _.filter(itens, { Id: this.iditem });
        let dataSourceByDefinition: any = await w.lists
          .getByTitle(this.listname)
          .items.filter(`FileDirRef eq '${datasourceItens[0].FileRef}'`)
          .get();
        datasourceItens[0].push({
          replies: dataSourceByDefinition,
        });
        return datasourceItens;
      });
  };

  // Return functions
  return {
    addReply,
    allItens,
    addItem,
    getItemByID,
  };
};
