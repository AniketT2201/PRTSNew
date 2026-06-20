import { IEmployeeProfile } from "../../INTERFACE/IEmployeeProfile";
import { IPrtsProps } from "../../../components/IPrtsProps";
import SPCRUDOPS from "../../DAL/spcrudops";
export interface IIITRequestsOps {    
    getEmployeeProfile(props: IEmployeeProfile): Promise<IEmployeeProfile>;
}

export default function IEmployeeProfileops() {
    const spCrudOps = SPCRUDOPS();    

    const getEmployeeProfile = async (ArtId: string | number, props: IPrtsProps): Promise<IEmployeeProfile[]> => {
        return await (await spCrudOps).getRootData("UserMaster"
            , 'EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID'
            , 'FullName,DirectManagerName,OfficeCity,DepartmentCode'
            , `FullName/EMail eq '`+ArtId+`' and EmployeeStatus eq 'Active'`
            , { column: 'ID', isAscending: true }            
            ,props).then(results => {
                let brr: Array<IEmployeeProfile> = new Array<IEmployeeProfile>();
                results.map((item: {
                    EmployeeId:number                                        
                }) => {
                    brr.push({
                    EmployeeID:item.EmployeeId
                    });
                });
                return brr;
            }
            );
    };

    return {
        getEmployeeProfile
    };
}