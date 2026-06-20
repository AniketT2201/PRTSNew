import { IPrtsProps } from "../../../components/IPrtsProps";
import SPCRUDOPS from "../../DAL/spcrudops";

export interface MasterPagesRequestsOps {
    getPRTSResourceData(sorting: any, props: any, filter: string): Promise<any[]>;
    getMFGShopSelectionData(sorting: any, props: any, filter: string): Promise<any[]>;
    getIssueCategoryData(sorting: any, props: any, filter: string): Promise<any[]>;
    getCommodityData(sorting: any, props: any, filter: string): Promise<any[]>;
    getBuildTypeData(sorting: any, props: any, filter: string): Promise<any[]>;
}
export default function MasterPagesRequestsOps() {
    const spCrudOps = SPCRUDOPS();
    const getPRTSResourceData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getData(
            "PRTSResource",
            "*,Title,Author/Title,AttachmentFiles",
            "Author,AttachmentFiles",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({ 
            ...item,
            ID: item.ID,
            Title: item.Title,
            Created:item?.Created,
            Author: item?.Author.Title || "",
            AttachmentFiles: item.AttachmentFiles ?? null,
        }));
    };
    const getMFGShopSelectionData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getRootData(
            "MFGShopList",
            "*,Title,Author/Title,AttachmentFiles",
            "Author,AttachmentFiles",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({
            ...item,
            ID: item.ID,
            Title: item.Title,
            Created:item?.Created,
            Author: item?.Author.Title || "",
            AttachmentFiles: item.AttachmentFiles ?? null,
        }));
    };
    const getIssueCategoryData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getRootData(
            "IssueCategoryList",
            "*,Title,Author/Title,AttachmentFiles",
            "Author,AttachmentFiles",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({
            ...item,
            ID: item.ID,
            Title: item.Title,
            Created:item?.Created,
            Author: item?.Author.Title || "",
            AttachmentFiles: item.AttachmentFiles ?? null,
        }));
    };
    const getCommodityData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getRootData(
            "CommodityList",
            "*,Title,Author/Title,CommodityHead/Id,CommodityHead/Title,CommodityHead/EMail,CommodityLead/Id,CommodityLead/Title,CommodityLead/EMail,AttachmentFiles",
            "Author,AttachmentFiles,CommodityHead,CommodityLead",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({
            ...item,
            ID: item.ID,
            Title: item.Title,
            Created:item?.Created,
            Author: item?.Author.Title || "",
            CommodityHeadId: item.CommodityHead?.Id || null,
            CommodityHeadTitle: item.CommodityHead?.Title || null,
            CommodityHeadEmail: item.CommodityHead?.EMail || null,
            CommodityLeadId: item.CommodityLead?.Id || null,
            CommodityLeadTitle: item.CommodityLead?.Title || null,
            CommodityLeadEmail: item.CommodityLead?.EMail || null,
            AttachmentFiles: item.AttachmentFiles ?? null,
        }));
    };
    const getBuildTypeData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getRootData(
            "BuildTypeList",
            "*,Title,Author/Title,AttachmentFiles",
            "Author,AttachmentFiles",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({
            ...item,
            ID: item.ID,
            Title: item.Title,
            Created:item?.Created,
            Author: item?.Author.Title || "",
            AttachmentFiles: item.AttachmentFiles ?? null,
        }));
    };

    return {
        getPRTSResourceData,
        getMFGShopSelectionData,
        getIssueCategoryData,
        getCommodityData,
        getBuildTypeData
    };
}