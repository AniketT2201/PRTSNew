import { IIAS } from "../../INTERFACE/IPRTS";
import { IPrtsProps } from "../../../components/IPrtsProps";
import SPCRUDOPS from "../../DAL/spcrudops";
export interface IIITRequestsOps {
    getIIASData(props: IIAS): Promise<IIAS>;
    getIASDatafilter(props: IIAS): Promise<IIAS>;
}

export default function IASRequestsOps() {
    const spCrudOps = SPCRUDOPS();

    // const getIIASData = async (strFilter: string, sorting: any,props: IItProps): Promise<IIAS[]> => {
    const getIIASData = async (sorting: any, props: IPrtsProps, filter:string): Promise<IIAS[]> => {
        return await (await spCrudOps).getData("PRTSList"
            , "*,ReqNo,Status,Stage,NA/ID,NA/Title,NA/EMail,DA/ID,DA/Title,DA/EMail,AnalysisDetails,Summary,InitDepartment,LastAction,IssueDescription,Author/Title,AttachmentFiles,Initiator/Id,Initiator/Title,InitiatorEmpId"
            , "NA,DA,Author,AttachmentFiles,Initiator"
            , filter
            , sorting
            , props).then(results => {
                let brr: Array<IIAS> = new Array<IIAS>();
                results.map((item: {
                    ID: number;
                    IssueDescription:any;
                    ReqNo?: any;
                    SAPNo:any; 
                    Initiator:any;  
                    Status:any;  
                    Stage:any; 
                    NA:any; 
                    NATitle:any;
                    NAEmail:any;
                    DA:any;
                    AnalysisDetails:any;
                    Summary:any;
                  //  WF:any;
                   // HeaderName:any;
                   // MovementType:any;
                    InitDepartment:any;
                    CostCenter:any;
                    UniquePartImpact:any;
                    GrossValue:any;
                    NetValue:any;
                    LastAction:any;
                    CostCenterDescription:any; 
                    Author:any;
                    Title:any; 
                    Created:any;  
                    NextApproverEmpID:any; 
                    DelegateApproverEmpID:any;
                    InitiatorEmpId:any;
                    InventoryAttachment:any;
                }) => {
                    brr.push({
                    ID:item.ID,
                    ReqNo:item?.ReqNo?? null,
                    IssueTitle:item?.Title??null,
                    SAPNo:item?.SAPNo?? null,
                    Initiator:item?.Initiator?.Title?? null,
                    Status:item?.Status?? null,
                    Stage:item?.Stage?? null,
                    NAID:item?.NA?.ID?? null,  
                    NATitle:item?.NA?.Title?? null,
                    NAEmail:item?.NA?.EMail?? null,
                    DAID:item?.DA?.ID?? null,  
                    DATitle:item?.DA?.Title?? null,
                    DAEmail:item?.DA?.EMail?? null,
                    AnalysisDetails:item?.AnalysisDetails?? null,
                    Summary:item?.Summary?? null,
                    //WF:item?.WF?? null,
                   // HeaderName:item?.HeaderName?? null,
                  //  MovementType:item?.MovementType?? null,
                    InitDepartment:item?.InitDepartment?? null,
                    CostCenter:item?.CostCenter?? null,
                    UniquePartImpact:item?.UniquePartImpact?? null,
                    GrossValue:item?.GrossValue?? null,
                    NetValue:item?.NetValue?? null,
                    LastAction:item?.LastAction?? null,
                    CostCenterDescription:item?.CostCenterDescription?? null,
                    AttachmentFiles:null,
                    Title:item?.Title,
                    Author:item?.Author.Title??null,
                    Created:item?.Created??null,
                    NextApproverEmpID:item?.NextApproverEmpID??null,
                    InitiatorEmpId:item?.InitiatorEmpId??null,
                    DelegateApproverEmpID:item?.DelegateApproverEmpID??null,
                    InventoryAttachment:item?.InventoryAttachment??null
                    });
                });
                return brr;
                }
            );
    };

    const getIASDatafilter = async (ArtId: string | number, props: IPrtsProps): Promise<IIAS[]> => {
        return await (await spCrudOps).getData("PRTSList"
            , "*,ReqNo,Status,Stage,NA/ID,NA/Title,NA/EMail,DA/ID,DA/Title,DA/EMail,AnalysisDetails,Summary,InitDepartment,LastAction,IssueDescription,Author/Title,AttachmentFiles,Initiator/Id,Initiator/email,IssueDescription"
            , "NA,DA,Author,AttachmentFiles,Initiator"
            , "Id eq '" + ArtId + "'"
            // , sorting,
            , { column: 'Order0', isAscending: true },
            props).then(results => {
                let brr: Array<IIAS> = new Array<IIAS>();
                results.map((item: {
                    ID: number;
                    ReqNo?: any;
                    IssueDescription?:any;
                    SAPNo:any; 
                    Initiator:any;  
                    Status:any;  
                    Stage:any; 
                    NA:any; 
                    NATitle:any;
                    NAEmail:any;
                    DA:any;
                    AnalysisDetails:any;
                    Summary:any;
                    // WF:any;
                    // HeaderName:any;
                    // MovementType:any;
                    InitDepartment:any;
                    CostCenter:any;
                    UniquePartImpact:any;
                    GrossValue:any;
                    NetValue:any;
                    LastAction:any;
                    CostCenterDescription:any; 
                    AttachmentFiles:any; 
                    Title:any;  
                    Author:any;   
                    Created:any; 
                    NextApproverEmpID:any; 
                    DelegateApproverEmpID:any;
                    InitiatorEmpId:any;  
                    InventoryAttachment:any;                                    
                }) => {
                    brr.push({
                    ID:item.ID,
                    ReqNo:item?.ReqNo?? null,
                    // SAPNo:item?.SAPNo?? null,
                    Initiator:item?.Initiator?.Title?? null,
                    Status:item?.Status?? null,
                    Stage:item?.Stage?? null,
                    NAID:item?.NA?.ID?? null,  
                    NATitle:item?.NA?.Title?? null,
                    NAEmail:item?.NA?.EMail?? null,
                    DAID:item?.DA?.ID?? null,  
                    DATitle:item?.DA?.Title?? null,
                    DAEmail:item?.DA?.EMail?? null,
                    AnalysisDetails:item?.AnalysisDetails?? null,
                    Summary:item?.Summary?? null,
                    // WF:item?.WF?? null,
                    // HeaderName:item?.HeaderName?? null,
                    // MovementType:item?.MovementType?? null,
                    InitDepartment:item?.InitDepartment?? null,
                    CostCenter:item?.CostCenter?? null,
                    UniquePartImpact:item?.UniquePartImpact?? null,
                    GrossValue:item?.GrossValue?? null,
                    NetValue:item?.NetValue?? null,
                    LastAction:item?.LastAction?? null,
                    CostCenterDescription:item?.CostCenterDescription?? null,
                    AttachmentFiles:item?.AttachmentFiles?? null,
                    Title:item?.Title,
                    Author:item?.Author.Title??null,
                    Created:item?.Created??null,
                    NextApproverEmpID:item?.NextApproverEmpID??null,
                    InitiatorEmpId:item?.InitiatorEmpId??null,
                    DelegateApproverEmpID:item?.DelegateApproverEmpID??null,
                    InventoryAttachment:item?.InventoryAttachment??null,
                    IssueTitle:item?.Title??null,
                    });
                });
                return brr;
            }
            );
    };

    return {
        getIIASData, getIASDatafilter
    };
}