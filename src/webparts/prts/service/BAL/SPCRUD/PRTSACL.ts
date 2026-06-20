import { IPrtsProps } from "../../../components/IPrtsProps";
import SPCRUDOPS from "../../DAL/spcrudops";

export interface IPRTSACLRequestsOps {
    getIPRTSACLData(strFilter: string, sorting: any, props: any): Promise<any>;
}
export default function IPRTSACLRequestsOps() {
    const spCrudOps = SPCRUDOPS();
    const getIPRTSACLData = async (sorting: any, props: IPrtsProps, filter: string): Promise<any[]> => {
        const results = await (await spCrudOps).getData(
            "PRTSACL",
            "*,Title,UserName/Id,UserName/Title,UserName/EMail,Role,EmployeeID,Status",
            "UserName",
            filter,
            sorting,
            props
        );
        return results.map((item: any): any => ({
            ...item,
            ID: item.ID,
            Title: item.Title,
            UserNameId: item.UserName?.ID || null,
            UserName: item.UserName?.Title || "",
            UserNameEmail: item.UserName?.EMail || "",
            Role: item.Role || "",
            EmployeeID: item.EmployeeID || "",
            Status: item.Status || ""
        }));
    };

    return {
        getIPRTSACLData
    };
}