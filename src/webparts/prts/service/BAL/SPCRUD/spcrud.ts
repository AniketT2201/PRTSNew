import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { IPatelEngProps } from "../../components/IPatelEngProps";
import { IPrtsProps } from "../../../components/IPrtsProps";
import SPCRUDOPS from "../../DAL/spcrudops";
 
export interface ISPCRUD {
    getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: IPrtsProps): Promise<any>;
    getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: IPrtsProps): Promise<any>;
    insertData(listName: string, data: any, props: IPrtsProps): Promise<any>;
    updateData(listName: string, itemId: number, data: any, props: IPrtsProps): Promise<any>;
    deleteData(listName: string, itemId: number, props: IPrtsProps): Promise<any>;
    getListInfo(listName: string, props: IPrtsProps): Promise<any>;
    getListData(listName: string, columnsToRetrieve: string, props: IPrtsProps): Promise<any>;
    batchInsert(listName: string, data: any, props: IPrtsProps): Promise<any>;
    batchUpdate(listName: string, data: any, props: IPrtsProps): Promise<any>;
    batchDelete(listName: string, data: any, props: IPrtsProps): Promise<any>;
    createFolder(listName: string, folderName: string, props: IPrtsProps):Promise<any>;
    uploadFile(folderServerRelativeUrl: string, file: File, props: IPrtsProps): Promise<any>;
    deleteFile(fileServerRelativeUrl: string, props: IPrtsProps): Promise<any>;
    currentProfile(props: IPrtsProps): Promise<any>;
    //currentUserProfile(props: IDeviationuatProps): Promise<any>;
    getLoggedInSiteGroups(props: IPrtsProps): Promise<any>;
    getAllSiteGroups(props: IPrtsProps): Promise<any>;
    getTopData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, top: number, props: IPrtsProps): Promise<any>;
    addAttchmentInList(attFiles: File, listName: string, itemId: number, fileName: string, props: IPrtsProps): Promise<any>;
}

export default async function USESPCRUD(): Promise<ISPCRUD> {
    const spCrudOps = await SPCRUDOPS();
    return {
        getData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: IPrtsProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        getRootData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: IPrtsProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        insertData: async (listName: string, data: any, props: IPrtsProps) => {
            return await spCrudOps.insertData(listName, data, props);
        },
        updateData: async (listName: string, itemId: number, data: any, props: IPrtsProps) => {
            return await spCrudOps.updateData(listName, itemId, data, props);
        },
        deleteData: async (listName: string, itemId: number, props: IPrtsProps) => {
            return await spCrudOps.deleteData(listName, itemId, props);
        },
        getListInfo: async (listName: string, props: IPrtsProps) => {
            return await spCrudOps.getListInfo(listName, props);
        },
        getListData: async (listName: string, columnsToRetrieve: string, props: IPrtsProps) => {
            return await spCrudOps.getListData(listName, columnsToRetrieve, props);
        },
        batchInsert: async (listName: string, data: any, props: IPrtsProps) => {
            return await spCrudOps.batchInsert(listName, data, props);
        },
        batchUpdate: async (listName: string, data: any, props: IPrtsProps) => {
            return await spCrudOps.batchUpdate(listName, data, props);
        },
        batchDelete: async (listName: string, data: any, props: IPrtsProps) => {
            return await spCrudOps.batchDelete(listName, data, props);
        },
        createFolder: async (listName: string, folderName: string, props: IPrtsProps) => {
            return await spCrudOps.createFolder(listName, folderName, props);
        },
        uploadFile: async (folderServerRelativeUrl: string, file: File, props: IPrtsProps) => {
            return await spCrudOps.uploadFile(folderServerRelativeUrl, file, props);
        },
        deleteFile: async (fileServerRelativeUrl: string, props: IPrtsProps) => {
            return await spCrudOps.deleteFile(fileServerRelativeUrl, props);
        },
        currentProfile: async (props: IPrtsProps) => {
            return await spCrudOps.currentProfile(props);
        },
        // const currentUserProfile = async (props: IDeviationuatProps) => {
          
        //    // const queryUrl = "https://etgworld.sharepoint.com/sites/UAT_BPM/_api/web/currentuser/groups";
            
        //     const result: any = await (await spCrudOps).currentUserProfile( props);
        //     return result;
        // };
        getLoggedInSiteGroups: async (props: IPrtsProps) => {
            return await spCrudOps.getLoggedInSiteGroups(props);
        },
        getAllSiteGroups: async (props: IPrtsProps) => {
            return await spCrudOps.getAllSiteGroups(props);
        },
        getTopData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean }, top: number, props: IPrtsProps) => {
            return await spCrudOps.getTopData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, top, props);
        },
        addAttchmentInList: async (attFiles: File, listName: string, itemId: number, fileName: string, props: IPrtsProps) => {
            return await spCrudOps.addAttchmentInList(attFiles, listName, itemId, fileName, props);
        }
    };
}