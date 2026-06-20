import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { Formik, Form, Field, FormikProps } from "formik";
import type { IPrtsProps } from '../IPrtsProps';
import { IUtilities } from '../../service/BAL/SPCRUD/utilities';
import SPCRUDOPS from "../../service/DAL/spcrudops";
import USESPCRUD, { ISPCRUD } from '../../service/BAL/SPCRUD/spcrud';
import { IPersonaProps } from 'office-ui-fabric-react';
//Date
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react';
//PlantCodeMaster
import { IMDR } from '../../service/INTERFACE/IMDR';
import IASRequestsOps from '../../service/BAL/SPCRUD/PRTS';
import Table from '../Pages/Table';
//Date
import { format } from 'date-fns';
import { ISPCRUDOPS } from '../../service/DAL/spcrudops';
import '../Pages/CSS/NewRequest.scss';
//Template
import renderTemplateTable from '../../service/BAL/SPCRUD/Template'
//Excel
import * as XLSX from "xlsx";
import TableToExcel from '@linways/table-to-excel';
import { useHistory } from 'react-router-dom';
import * as yup from 'yup';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import IEmployeeProfileops from '../../service/BAL/SPCRUD/EmployeeProfile';
import IDelegateApproverops from '../../service/BAL/SPCRUD/DelegateApprover';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUpload, faTrash } from "@fortawesome/free-solid-svg-icons";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

const MRI: IDropdownOption[] = [
    { key: 'Yes', text: 'Yes' },
    { key: 'No', text: 'No' },
];

interface FormValues {
    requesterName: string,
    reqDepartment: any,
    reqDate: any,
    approvalNoteNo: string,
    sapNo: string,
    uniquePartImpact: any,
    grossValue: any,
    netValue: any,
    movementType: any,
    costCenter: any,
    movementReason: any,
    EmployeeID: any,
    CarLine: any,
    ReasonCategory: any
}

interface ExcelRecord {
    [key: string]: string;
}

type WorkflowStep = {
    type: string;
    user: string;
    email: string;
    required: boolean;
    EmpID: string;
};

type ApprFlow = {
    user: string;
    type: string;
    required: boolean;
}

interface IRow {
    key: number;
    partNumber: string;
    description: string;
    supplier: string;
    qty: number;
    value: number;
    amount: number;
    remarks: string;
}

export interface TableRef {
    getData: () => IRow[];
    resetData: () => void;
    setData: (data: IRow[]) => void;
}

export const Draft: React.FC<IPrtsProps> = (props: IPrtsProps) => {
    const formikRef = useRef<FormikProps<FormValues>>(null);
    const initialvalue = {
        requesterName: '',
        reqDepartment: '',
        reqDate: '',
        approvalNoteNo: '',
        sapNo: '',
        uniquePartImpact: '',
        grossValue: '',
        netValue: '',
        movementType: '',
        costCenter: '',
        movementReason: '',
        EmployeeID: '',
        CarLine: '',
        ReasonCategory: ''
    }

    interface IRow {
        key: number;
        partNumber: string;
        description: string;
        supplier: string;
        qty: number;
        value: number;
        amount: number;
        remarks: string;
    }

    let spCrudObj: ISPCRUD;
    const history = useHistory();
    //MASTER LIST  
    const [CostCenterdata, setCostCenterdata] = useState([]);//Costcenter list data  
    const [EmployeeData, setEmployeeData] = useState([]); //Employee Department from Employee Profile
    //MAIN LIST
    const [MDRData, setMDRData] = useState([]); //MDR data as per request  
    const [visibleButtons, setVisibleButtons] = useState([]);//Handle Button Visibility        
    const [BindingWorkflow, setWorkflow] = useState<WorkflowStep[]>([]);
    const [workflowJSX, setWorkflowJSX] = useState(null);
    const [userWF, setuserWF] = useState<any[]>([]);
    const [attachments, setAttachments] = useState<File[]>([]);
    const [loading, setLoading] = useState(false);
    const [Buttondisable, setButtondisable] = useState(true);
    const [Maincontainer, setMaincontainer] = useState(null);
    const [sharedData, setSharedData] = useState([]);
    const [ReasonCategorydata, setReasonCategory] = useState([]);
    const [TotalAmountUp, setTotalAmountUp] = React.useState<any>();
    const [Summary, setSummary] = useState([]);
    const [summaryRows, setSummaryRows] = useState([]);
    const [ItemID, setItemID] = useState<number>();
    const [bindedattachments, setbindedattachments] = useState([]);
    const [Carline, setCarLine] = useState<any[]>([]);
    const [missingdata, setmissingdata] = useState(false);
    const [subbtn, setsubbtn] = useState(false);


    // const [SiteWiseApproval, setSiteWiseApproval] = useState([]);
    const tableRef = useRef<TableRef>(null);

    useEffect(() => {
        const data = tableRef.current?.getData() || [];

        const totalAmount = data.reduce((sum, item) => {
            return sum + Number(item.amount || 0);
        }, 0);

        console.log("Total Amount:", totalAmount);
        setTotalAmountUp(totalAmount);

        let updateworkflow = BindingWorkflow;
        for (var i in updateworkflow) {
            if (updateworkflow[i].type == "StaffHead" && formikRef.current.values.reqDepartment != "SCM" && formikRef.current.values.reqDepartment != "CEM" && formikRef.current.values.reqDepartment != "PDC") {
                if (totalAmount >= 20000)
                    updateworkflow[i].required = true;
                else
                    updateworkflow[i].required = false;
                break;
            }
        }
        setWorkflow(updateworkflow);
        displayWorkflow();

    }, [sharedData]);

    //Global Variables  
    let Stage = useRef(0);
    let ExternalApproverData = useRef([]);
    let Copyupdateworkflow = useRef<any[]>([]);
    let SiteWiseApproval = useRef<any[]>([]);
    //for Formik
    function getFieldProps(formik: FormikProps<any>, field: string) {
        return { ...formik.getFieldProps(field), errorMessage: formik.errors[field] as string };
    }

    //fetchdata
    const fetchData = async () => {
        try {
            await GetSiteWiseApproval();
            await GetCostCenterdata();
            await GetReasonCategorydata();
            await GetExternalApproverData();
            await GetCarlineData();
            await GetUserDetails();
            const hash = window.location.hash;
            const queryString = hash.split("?")[1];
            const queryParams = new URLSearchParams(queryString);
            const urlsearch = queryParams.get("ItemId");
            await Binddata(urlsearch, '');
        } catch (error) {
            console.error("Failed to fetch ACL data:", error);
        }
    };

    const formatDate = (dateString) => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-GB'); // This gives dd/mm/yyyy format
    };

    async function Binddata(urlsearch, Parameterdata) {
        const itemdata = await IASRequestsOps().getIASDatafilter(urlsearch, props);
        const item = itemdata[0];
        const tmp = JSON.parse(item.Summary);
        setSummary(JSON.parse(item.Summary));
        setSummaryRows(tmp);
        formikRef.current?.setFieldValue('reqID', item.Title);
        // formikRef.current?.setFieldValue('requesterName', item.EmpName);
        formikRef.current?.setFieldValue('reqDepartment', item.InitDepartment);
        // formikRef.current?.setFieldValue('reqDate', formatDate(item.Date));
        formikRef.current?.setFieldValue('EmployeeId', item.EmpID);
        formikRef.current?.setFieldValue('costCenter', item.CostCenter);
        // formikRef.current?.setFieldValue('CarLine', item.CarLine);
        // formikRef.current?.setFieldValue('movementReason', item.Reason);
        // formikRef.current?.setFieldValue('ReasonCategory', item.ReasonCategory);
        formikRef.current?.setFieldValue('Status', item.Status);
        setItemID(item.ID);
        setWorkflow(JSON.parse(item.ApproverList));
        let attachment = item.AttachmentFiles.map(att => ({
            name: att.FileName,
            url: att.ServerRelativeUrl,
        }));
        setbindedattachments(attachment);
        const mappedData = JSON.parse(item.Items).map((item, index) => ({
            key: item.c0,
            partNumber: item.c1,
            description: item.c2,
            supplier: item.c3,
            qty: item.c4,
            value: item.c5,
            amount: item.c6,
            remarks: item.c7
        }));
        setSharedData(mappedData);
        tableRef.current?.setData(mappedData);
    }

    //onload 
    useEffect(() => {
        fetchData();
    }, []);

    //formating Date
    function formatDateTime(date) {
        const pad = (n) => n.toString().padStart(2, '0');
        return `${pad(date.getDate())}-${pad(date.getMonth() + 1)}-${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
    }

    //List Data of CostCenter
    async function GetCostCenterdata() {
        const spCrudOps = await SPCRUDOPS();
        const CostCenterdata = await spCrudOps.getRootData(
            'CostCenter',
            'Description,Title,ID,CCOwner/Id,CCOwner/Title,CCOwner/EMail,FinApprover/Id,FinApprover/Title,FinApprover/EMail',
            'CCOwner,FinApprover',
            '',
            { column: 'ID', isAscending: true },
            props
        );

        setCostCenterdata(CostCenterdata);
    }

    //List Data of CostCenter
    async function GetReasonCategorydata() {
        const spCrudOps = await SPCRUDOPS();
        const fetchReasonCategory = await spCrudOps.getData(
            'ReasonCategory',
            'Reason',
            '',
            '',
            { column: 'ID', isAscending: true },
            props
        );

        setReasonCategory(fetchReasonCategory);
    }

    //List Data of Site Wise Approval Level
    async function GetSiteWiseApproval() {
        const spCrudOps = await SPCRUDOPS();
        const SiteWiseApprovalData = await spCrudOps.getRootData(
            'SiteWiseApproval',
            'Title,Level',
            '',
            `Title eq 'MDR'`,
            { column: 'ID', isAscending: true },
            props
        );

        //setSiteWiseApproval(SiteWiseApproval);
        SiteWiseApproval.current = SiteWiseApprovalData;
    }

    //List Data of External Approver
    async function GetExternalApproverData() {
        const spCrudOps = await SPCRUDOPS();
        const ExternalApprData = await spCrudOps.getData(
            'External_Approver',
            'ID,Title,MaterialController/Title,MaterialController/EMail,WHController/Title,WHController/EMail,InventoryController/Title,InventoryController/EMail,SCMHead/Title,SCMHead/EMail,LogisticsWH/Title,LogisticsWH/EMail,Finance/Title,Finance/EMail',
            'MaterialController,WHController,InventoryController,SCMHead,LogisticsWH,Finance',
            '',
            { column: 'ID', isAscending: true },
            props
        );

        //setExternalApproverData(ExternalApprData);
        ExternalApproverData.current = ExternalApprData

    }

    async function GetCarlineData() {
        const spCrudOps = await SPCRUDOPS();
        const carlineData = await spCrudOps.getRootData(
            'CarLine',
            'CarLine',
            '',
            '',
            { column: 'ID', isAscending: true },
            props
        );

        const uniqueCarLines = Array.from(
            new Map(carlineData.map(item => [item.CarLine, item])).values()
        );

        setCarLine(uniqueCarLines);
    }

    const handleClose = () => {
        const lastActive = sessionStorage.getItem('sidebarFrom');
        if (lastActive) {
            history.push(lastActive);
        } else {
            history.push('/InitiatorLanding'); // Fallback route if none found
        }
    };

    const getuserdata = async (mail) => {
        sp.setup({
            sp: {
                baseUrl: props.currentSPContext.pageContext.web.absoluteUrl
            },
        });

        const result = await sp.web.ensureUser(`i:0#.f|membership|` + mail);
        console.log(result);
        return result;
    }

    const validateFormAndTable = async () => {
        const formikValues = formikRef.current?.values;
        const tableData = tableRef.current?.getData() || [];

        const errors: { [key: string]: string } = {};

        if (!formikValues?.CarLine) errors.CarLine = 'Car Line is Missing';
        if (!formikValues?.costCenter) errors.costCenter = 'Cost Center is Missing';
        if (!formikValues?.movementReason) errors.movementReason = 'Description/Remarks is Missing';
        if (!formikValues?.ReasonCategory) errors.movementReason = 'Reason Category is Missing';

        if (tableData.length === 0) {
            errors.table = 'Please enter the Part Details';
        } else {
            const rowErrors = tableData.map((row, index) => {
                const missingFields = [];
                if (!row.partNumber) missingFields.push('Part Number');
                if (!row.description) missingFields.push('Description');
                if (!row.qty || row.qty <= 0) missingFields.push('Qty');
                if (!row.value || row.value <= 0) missingFields.push('Value');
                return missingFields.length > 0
                    ? `Row ${index + 1}: Missing ${missingFields.join(', ')}`
                    : null;
            }).filter(Boolean);

            if (rowErrors.length > 0) {
                errors.table = rowErrors.join('\n');
            }
        }

        if (Object.keys(errors).length > 0) {
            console.log('Validation errors:', errors);
            alert(Object.values(errors).join('\n'));
            return false;
        }

        return true;
    };

    async function submitRequest(status, TableData) {
        setLoading(true);
        setsubbtn(true);

        try {
            const valid = await validateFormAndTable();
            if (!valid) {
                setLoading(false);
                setsubbtn(false);
                return;
            }

            const data = TableData;
            const mappedData = data.map(item => ({
                c0: item.key,
                c1: item.partNumber,
                c2: item.description,
                c3: item.supplier,
                c4: item.qty,
                c5: item.value,
                c6: item.amount,
                c7: item.remarks
            }));
            console.log("Submitted data:", mappedData);

            const [empData, initiatorId] = await Promise.all([
                getuserdata(props.userEmail),
                GetEmployeeID(props.userEmail)
            ]);
            const empName = empData.data.Id;

            const updatedStatus = status === "Draft" ? "Draft" : "Pending Approval";
            const stage = status === "Draft" ? "0" : "1";

            // Determine which approver to use (index 1 or 2)
            const approverIndex = BindingWorkflow[1]?.required ? 1 : 2;
            const approverEmail = BindingWorkflow[approverIndex].email;

            const approverData = await getuserdata(approverEmail);
            const naId = approverData.data.Id;
            const naEmployeeId = await GetEmployeeID(approverEmail);

            const delegateData = await IDelegateApproverops().getDelegateApprover(approverEmail, props);
            const delegateApproverEmpId = delegateData?.[0]?.DelegateToEmpID || "";
            const daId = delegateData?.[0]?.DelegateToId || 0;

            const fields = {
                EmpID: initiatorId,
                EmpNameId: empName,
                Department: formikRef.current.values.reqDepartment,
                ApproverList: JSON.stringify(BindingWorkflow),
                Status: updatedStatus,
                Stage: stage,
                CarLine: formikRef.current.values.CarLine,
                Reason: formikRef.current.values.movementReason,
                Items: JSON.stringify(mappedData),
                CostCenter: formikRef.current.values.costCenter,
                ReasonCategory: formikRef.current.values.ReasonCategory,
                //NAId: naId,
                //NextApproverEmpID: naEmployeeId,
                //DAId: daId,
                //DelegateApproverEmpID: delegateApproverEmpId,
                //Summary: JSON.stringify(summaryList)
            };

            // Safely construct summary
            let summaryList = Array.isArray(Summary) ? [...Summary] : [];
            if (status !== 'Draft') {
                summaryList.push({
                    c1: props.userDisplayName,
                    c2: approverData.data.Title,
                    c3: formatDateTime(new Date()),
                    c4: "Request Submitted",
                    c5: ""
                });

                fields["NAId"] = naId;
                fields["NextApproverEmpID"] = naEmployeeId;
                fields["DAId"] = daId;
                fields["DelegateApproverEmpID"] = delegateApproverEmpId;
                fields["Summary"] = JSON.stringify(summaryList);
            }

            console.log("Fields to update:", fields);

            const spCrudObj = await USESPCRUD();

            // Step 1: Upload attachments (if any)
            if (ItemID && Array.isArray(attachments) && attachments.length > 0) {
                for (const file of attachments) {
                    try {
                        await spCrudObj.addAttchmentInList(file, "MDR_List", ItemID, file.name, props);
                        console.log(`Attachment ${file.name} uploaded.`);
                    } catch (error) {
                        console.error(`Failed to upload attachment ${file.name}:`, error);
                        alert(`Failed to upload attachment ${file.name}`);
                        setButtondisable(true);
                        return;
                    }
                }
            }

            // Step 2: Update list item
            await spCrudObj.updateData("MDR_List", ItemID, fields, props);
            if (status === 'Draft') {
                alert(`Request ` + formikRef.current.values.requesterName + ` has been Saved as Draft`);
            }
            else {
                alert(`Request ` + formikRef.current.values.requesterName + ` has been Submitted`);
            }
            //alert(`Request ${formikRef.current.values.requesterName} has been Submitted Successfully`);
            history.push("/InitiatorLanding");
        } catch (error) {
            console.error("submitRequest error:", error);
            alert("An error occurred while submitting the request. Please try again.");
        } finally {
            setLoading(false);
            setsubbtn(false);
        }
    }

    async function EmployeeProfile(Email) {
        const spCrudOps = await SPCRUDOPS();
        const EmployeeProfiledata = await spCrudOps.getRootData(
            'UserMaster',
            'EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID',
            'FullName,DirectManagerName,OfficeCity,DepartmentCode',
            `FullName/EMail eq '` + Email + `'`,
            { column: 'ID', isAscending: true },
            props
        );
        setEmployeeData(EmployeeProfiledata);
        return EmployeeProfiledata;
    }

    async function GetLeadershipApprovers(department) {
        try {
            const spCrudOps = await SPCRUDOPS();

            const departmentData = await spCrudOps.getRootData(
                'LeadershipApprovers',
                'Department/Id,Department/Department,User/Id,User/Title,User/EMail,Role',
                'Department,User',
                `Department/Department eq '${department}' and Role eq 'StaffHead'`,
                { column: 'ID', isAscending: true },
                props
            );

            setEmployeeData(departmentData);
            return departmentData;
        } catch (error) {
            console.error('Error fetching leadership approvers:', error);
            setEmployeeData([]);
            return [];
        }
    }

    async function GetUserDetails() {
        let item = await EmployeeProfile(props.userEmail);
        try {
            if (item.length > 0 && (SiteWiseApproval.current[0].Level != null && SiteWiseApproval.current[0].Level != undefined && SiteWiseApproval.current[0].Level != '')) {
                formikRef.current?.setFieldValue('EmployeeId', item[0].EmployeeId);
                const today = new Date();
                const formattedDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
                formikRef.current?.setFieldValue('reqDate', formattedDate);
                formikRef.current?.setFieldValue('reqDepartment', item[0].DepartmentCode.Department);
                let LeadershipAppData = await GetLeadershipApprovers(item[0].DepartmentCode.Department);

                //Copyupdateworkflow.current.push(JSON.parse('{"user":"' + item[0].FullName.Title + '","type":"initiator","required":true,"email":' + item[0].FullName.EMail +'}'))
                Copyupdateworkflow.current.push({
                    user: item[0].FullName.Title,
                    type: "initiator",
                    required: true,
                    email: item[0].FullName.EMail,
                });
                var cntApprover = SiteWiseApproval.current[0].Level;
                if (cntApprover && item[0].EmployeeId === props.EmployeeId[0].EmployeeID) {
                    showButtons('.btn-init');
                }

                let nextmanager;
                let test;

                const othermanager = () => {
                    nextmanager = item[0].DirectManagerName.EMail;
                    test = {
                        user: item[0].DirectManagerName.Title,
                        type: "Manager",
                        required: true,
                        email: item[0].DirectManagerName.EMail
                    };
                };

                for (let i = 1; i <= parseInt(cntApprover); i++) {
                    const department = item[0].DepartmentCode.Department;

                    // Department is NOT PDC or SCM → use fallback and exit loop
                    if (department !== 'PDC' && department !== 'SCM') {
                        othermanager();
                        Copyupdateworkflow.current.push(test);
                        break;
                    }

                    // First iteration for PDC/SCM → set up initial manager
                    if (i === 1) {
                        nextmanager = item[0].DirectManagerName.EMail;
                        test = {
                            user: item[0].DirectManagerName.Title,
                            type: "Manager",
                            required: true,
                            email: item[0].DirectManagerName.EMail
                        };
                    }

                    // Subsequent iterations → fetch next manager in chain
                    else {
                        const currentEmployeeData = await EmployeeProfile(nextmanager);
                        const directManager = currentEmployeeData[0].DirectManagerName;
                        const departmentMatch = currentEmployeeData[0].DepartmentCode.Department === department;

                        nextmanager = directManager.EMail;

                        if (!departmentMatch) {
                            break;
                        }

                        test = {
                            user: directManager.Title,
                            type: `Manager${i}`,
                            required: true,
                            email: directManager.EMail
                        };
                    }

                    // Push test only if it's defined
                    if (test) {
                        Copyupdateworkflow.current.push(test);
                    }
                }


                Copyupdateworkflow.current.push({
                    user: LeadershipAppData[0].User.Title,
                    type: "StaffHead",
                    required: false,
                    email: LeadershipAppData[0].User.EMail
                });
                //Copyupdateworkflow.current[Copyupdateworkflow.current.length - 1].type = "StaffHead";
                //Copyupdateworkflow.current[Copyupdateworkflow.current.length - 1].required = false;

                await GetExternalApprover(item[0].DepartmentCode.Department, Copyupdateworkflow);
            }
            else {
                //$(".MainContainer").html("<h1 style='text-align:center'>Missing Master Data.<br>Please contact administrator!!</h1>");
                let wf = (
                    <React.Fragment>
                        <h1 style={{ textAlign: 'center', color: 'white' }}>
                            Missing Master Data Please Contact IT Team
                        </h1>
                    </React.Fragment>
                );

                setWorkflowJSX(wf);
            }
        }
        catch (error) {
            console.log(error);
            setmissingdata(true);
            let wf = (
                <React.Fragment>
                    <h1 style={{ textAlign: 'center', color: 'white' }}>
                        Missing Master Data Please Contact IT Team
                    </h1>
                </React.Fragment>
            );

            setWorkflowJSX(wf);
        }
    }

    async function GetExternalApprover(department, wf) {
        let updatedwf = [];
        if (wf.current.length > 0) {
            for (var i = 0; i < wf.current.length; i++) {
                updatedwf.push(wf.current[i]);
            }
        }
        let FilterExternalApproverData;
        if (department == "CEM" || department == "SCM" || department == "PDC") {
            FilterExternalApproverData = ExternalApproverData.current.filter(test => test.Title === department);
        }
        else {
            FilterExternalApproverData = ExternalApproverData.current.filter(test => test.Title === 'OTHERS');
        }

        if (FilterExternalApproverData.length === 1) {
            const data = FilterExternalApproverData[0];

            updatedwf.push({
                user: data.MaterialController.Title,
                type: "MC",
                required: true,
                email: data.MaterialController.EMail
            });

            if (department === 'SCM') {
                updatedwf.push({
                    user: data.SCMHead.Title,
                    type: "PDC",
                    required: true,
                    email: data.SCMHead.EMail
                });
            }

            updatedwf.push({
                user: data.WHController.Title,
                type: "WHC",
                required: true,
                email: data.WHController.EMail
            });

            updatedwf.push({
                user: data.InventoryController.Title,
                type: "IC",
                required: true,
                email: data.InventoryController.EMail
            });

            if (department !== 'SCM') {
                updatedwf.push({
                    user: data.SCMHead.Title,
                    type: "SCM",
                    required: true,
                    email: data.SCMHead.EMail
                });

                if (department !== 'PDC' && department !== 'SCM') {
                    updatedwf.push({
                        user: data.Finance.Title,
                        type: "FIN",
                        required: true,
                        email: data.Finance.EMail
                    });
                }

                updatedwf.push({
                    user: data.LogisticsWH.Title,
                    type: "WH",
                    required: true,
                    email: data.LogisticsWH.EMail
                });
            }

            setuserWF(updatedwf);
            setWorkflow(updatedwf);
        }
        else {
            //$(".MainContainer").html("<h1 style='text-align:center'>Missing external approver details for the requester department.<br>Please contact administrator!!</h1>");
            let wf = (
                <React.Fragment>
                    <h1 style={{ textAlign: 'center' }}>
                        Missing external approver details for the requester department.<br />
                        Please contact administrator!!
                    </h1>
                </React.Fragment>
            );
            setMaincontainer(wf);
        }
    }

    // Get Employee ID with error handling
    async function GetEmployeeID(Email: string): Promise<string | null> {
        try {
            const spCrudOps = await SPCRUDOPS();

            const EmployeeProfiledata = await spCrudOps.getRootData(
                'UserMaster',
                'EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID',
                'FullName,DirectManagerName,OfficeCity,DepartmentCode',
                `FullName/EMail eq '${Email}' and EmployeeStatus eq 'Active'`,
                { column: 'ID', isAscending: true },
                props
            );

            if (!EmployeeProfiledata || EmployeeProfiledata.length === 0) {
                console.warn("Employee ID not found for email:", Email);
                return null;
            }

            if (EmployeeProfiledata.length > 1) {
                console.warn("Multiple active employees found with the same email:", Email);
                return null;
            }

            const empId = EmployeeProfiledata[0]?.EmployeeId;
            if (!empId) {
                console.warn("EmployeeId field missing for email:", Email);
                return null;
            }

            return empId;

        } catch (error) {
            console.error("Error fetching Employee ID for " + Email + ":", error);
            return null;
        }
    }

    const handleAddAttachments = () => {
        const input = document.createElement("input");
        input.type = "file";
        input.multiple = true; // ✅ allow multiple files

        input.onchange = (e) => {
            const target = e.target as HTMLInputElement;
            const newFiles = Array.from(target.files || []);

            // Avoid duplicates
            const filtered = newFiles.filter(
                newFile =>
                    !attachments.some(
                        existing => existing.name === newFile.name && existing.size === newFile.size
                    )
            );

            setAttachments(prev => [...prev, ...filtered]);
        };

        input.click();
    };

    const handleDeleteAttachment = (index: number) => {
        const fileName = attachments[index];
        const confirmDelete = window.confirm(`Are you sure you want to delete the attachment as it will be deleted permanently: ${fileName}?`);
        if (!confirmDelete) return; // User cancelled
        setAttachments(prev => prev.filter((_, i) => i !== index));
    };

    const DeleteAttachment = async (index: number) => {
        const fileName = bindedattachments[index].url;

        const confirmDelete = window.confirm(`Are you sure you want to delete the attachment as it will be deleted permanently: ${bindedattachments[index].name}?`);
        if (!confirmDelete) return; // User cancelled

        try {
            const spCrudObj = await USESPCRUD();
            const result = await spCrudObj.deleteFile(fileName, props);

            if (result) {
                // Success: Remove from UI
                setbindedattachments(prev => prev.filter((_, i) => i !== index));
            } else {
                console.error("Delete failed", result);
                alert("Failed to delete attachment. Please try again.");
            }
        } catch (error) {
            console.error("Delete error:", error);
            alert("An error occurred while deleting the attachment.");
        }
    };


    //List Data of MDR using Id
    const GetMDRData = async (id) => {
        const MDRColl = await IASRequestsOps().getIASDatafilter(id, props);
        setMDRData(MDRColl);
    }


    //Button Visibility
    const showButtons = (arr) => {
        let sg = "";
        let btns = [".btn-draft", ".btn-init", ".btn-withdrawn", ".btn-approver"];
        let filtered = btns.filter((btn) => arr.includes(btn));
        setVisibleButtons(filtered);
    };

    useEffect(() => {
        if (BindingWorkflow.length > 0) {
            displayWorkflow();
        }
    }, [BindingWorkflow, Stage]);

    {/* Add arrow if not the last element */ }
    {/* Only count required items for arrow placement */ }
    const displayWorkflow = () => {
        let wf = [];

        BindingWorkflow.forEach((m, i) => {
            if (m.required === true) {
                const isActive = i === Stage.current ? 'activeApprover' : 'overrideStage';
                wf.push(
                    <React.Fragment key={i}>
                        <ul className="main-menu">
                            <li className={`${m.type} ${isActive}`.trim()}>
                                {m.user}
                            </li>
                        </ul>
                    </React.Fragment>
                );
            }
        });

        setWorkflowJSX(wf);
    };

    return (
        <Formik initialValues={initialvalue} innerRef={formikRef} onSubmit={() => {}}>
            <Form onKeyDown={(e) => {
                const target = e.target as HTMLElement;
                if (e.key === 'Enter' && target.tagName !== 'TEXTAREA') {
                    e.preventDefault();
                }
            }}>
                {loading ? (
                    <div className="loading-overlay">
                        <div className="loading-content">
                            <svg
                                className="loading-spinner"
                                xmlns="http://www.w3.org/2000/svg"
                                fill="none"
                                viewBox="0 0 24 24"
                            >
                                <circle
                                    className="opacity-25"
                                    cx="12"
                                    cy="12"
                                    r="10"
                                    stroke="currentColor"
                                    strokeWidth="4"
                                />
                                <path
                                    className="opacity-75"
                                    fill="currentColor"
                                    d="M4 12a8 8 0 018-8v8H4z"
                                />
                            </svg>
                            <p className="text-white text-lg">Please wait, loading data...</p>
                        </div>
                    </div>
                ) : (
                    <>
                        <div className="container p-0" >
                            <div className="header">
                                <div className="left-banner">
                                    {/* <img src={`${props.currentSPContext.pageContext.web.absoluteUrl}/SiteAssets/Custom/imgs/MG-Motor-Logo.png`} alt="" className="hexagon" /> */}
                                    <div className="logo-text">
                                        <h2>Request Form</h2>
                                    </div>
                                </div>
                                {/* Add other header elements here if needed */}
                            </div>
                            <div id="mainContainer">
                                <div id="tablemain" style={{ overflowX: "auto" }}>
                                    <table className="table table-bordered">
                                        <colgroup>
                                            {[...Array(12)].map((_, i) => (
                                                <col key={i} style={{ width: '8.33%' }} />
                                            ))}
                                        </colgroup>
                                        {/* <tr className="wf-tr">
                    <td colSpan={12} className="wf-padding">
                      <div className="displayWF">{workflowJSX}</div>
                      <div className="displayWFdelegated hidden"></div>
                    </td>
                    </tr> */}
                                        <thead>
                                            <tr className="wf-tr p-0">
                                                <td colSpan={12} className="wf-padding p-0">
                                                    <div className="displayWF">{workflowJSX}</div>
                                                    <div className="displayWFdelegated hidden"></div>
                                                </td>
                                            </tr>

                                            <tr>
                                                <td colSpan={12} className="bg-darkgray p-0">
                                                    <div className="button-bar">
                                                        {Buttondisable && (
                                                            <>
                                                                {visibleButtons.includes(".btn-init") && (
                                                                    <>
                                                                        <button
                                                                            className="btn btn-warning btn-init"
                                                                            type="button"
                                                                            onClick={() => submitRequest('Draft', tableRef.current?.getData())}
                                                                            style={{ backgroundColor: '#030397' }}
                                                                        >
                                                                            <i className="fa fa-save"></i> Save
                                                                        </button>

                                                                        <button
                                                                            className="btn btn-warning btn-init"
                                                                            type="button"
                                                                            onClick={() => submitRequest('Submit', tableRef.current?.getData())}
                                                                        >
                                                                            <i className="fa fa-mail-forward"></i> Submit
                                                                        </button>
                                                                    </>
                                                                )}

                                                                {visibleButtons.includes(".btn-withdrawn") && (
                                                                    <a className="btn btn-warning btn-withdrawn" type="button">
                                                                        <i className="fa fa-times"></i> Withdraw
                                                                    </a>
                                                                )}

                                                                {visibleButtons.includes(".btn-approver") && (
                                                                    <>
                                                                        <a className="btn btn-warning btn-approver" type="button">
                                                                            <i className="fa fa-check"></i> Approve
                                                                        </a>
                                                                        <a className="btn btn-warning btn-approver" type="button">
                                                                            <i className="fa fa-times"></i> Reject
                                                                        </a>
                                                                        <a className="btn btn-warning btn-approver" type="button">
                                                                            <i className="fa fa-undo"></i> Rework
                                                                        </a>
                                                                        <a className="btn btn-warning btn-approver" type="button">
                                                                            <i className="fa fa-comments"></i> Comment
                                                                        </a>
                                                                    </>
                                                                )}

                                                                {visibleButtons.includes(".btn-forward") && (
                                                                    <a className="btn btn-warning btn-approver btn-forward" type="button">
                                                                        <i className="fa fa-forward"></i> Forward
                                                                    </a>
                                                                )}

                                                                {visibleButtons.includes(".btn-print") && (
                                                                    <a className="btn btn-warning btn-print" type="button">
                                                                        <i className="fa fa-print"></i> Print Preview
                                                                    </a>
                                                                )}
                                                            </>
                                                        )}
                                                        <a className="btn btn-warning btn-approver btn-forward"
                                                            onClick={handleClose} type="button"
                                                        >
                                                            <i className="fa fa-forward"></i> Close
                                                        </a>

                                                        <div className="requestStatus">
                                                            <span>Status: <span className="displayStatus"><Field name="Status" readOnly style={{border:"none"}}/></span></span>
                                                        </div>
                                                    </div>
                                                </td>
                                            </tr>
                                        </thead>
                                        <tbody className='tbodylabel'>
                                            <tr>
                                                <td colSpan={3}>
                                                    <label>Application Track No.</label>
                                                    <Field name="reqID" readOnly className="form-control" />
                                                </td>
                                                <td colSpan={3}>
                                                    <label>Employee Name.</label>
                                                    <Field name="requesterName" readOnly className="form-control" />
                                                </td>
                                                <td colSpan={3}>
                                                    <label>Consuming Department</label>
                                                    <Field name="reqDepartment" readOnly className="form-control" />
                                                </td>
                                                <td colSpan={3}>
                                                    <label>Date</label>
                                                    <Field name="reqDate" readOnly className="form-control" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colSpan={3}>
                                                    <label><span style={{ color: 'red' }}>*</span> Cost Center No.</label>
                                                    <Field as="select" name="costCenter" className="form-control">
                                                        <option value="">Select</option>
                                                        {CostCenterdata?.map((Vend) => (
                                                            <option key={Vend.ID} value={Vend.Title + ' - ' + Vend?.Description}>
                                                                {Vend.Title} - {Vend?.Description}
                                                            </option>
                                                        ))}
                                                    </Field>
                                                </td>
                                                <td colSpan={3}>
                                                    <label><span style={{ color: 'red' }}>*</span> Reason Category</label>
                                                    <Field as="select" name="ReasonCategory" className="form-control">
                                                        <option value="">Select</option>
                                                        {ReasonCategorydata?.map((Vend) => (
                                                            <option key={Vend.Reason} value={Vend.Reason}>
                                                                {Vend.Reason}
                                                            </option>
                                                        ))}
                                                    </Field>
                                                </td>
                                                <td colSpan={3}>
                                                    <label>Employee ID</label>
                                                    <Field name="EmployeeId" readOnly className="form-control" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <th colSpan={3} className="text-right">
                                                    <span style={{ color: 'red' }}>*</span> CarLine</th>
                                                <td colSpan={6}>
                                                    <Field as="select" name="CarLine" className="form-control">
                                                        <option value="">Select</option>
                                                        {Carline?.map((Vend) => (
                                                            <option key={Vend.CarLine} value={Vend.CarLine}>
                                                                {Vend.CarLine}
                                                            </option>
                                                        ))}
                                                    </Field>
                                                </td>
                                            </tr>
                                            <tr>
                                                <th colSpan={3} className="text-right"><span style={{ color: 'red' }}>*</span> Description/Remarks </th>
                                                <td colSpan={9}>
                                                    <Field as="textarea" rows={4} name="movementReason" className="form-control large-textarea" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colSpan={12}>
                                                    <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
                                                        <span className="h4 m-0">Approval Note from Buyer</span>
                                                        <button
                                                            className="btn btn-warning btn-attachment btn-init"
                                                            type="button"
                                                            onClick={handleAddAttachments}
                                                            title="Add Attachment"
                                                            aria-label="Add Attachment"
                                                        >
                                                            <FontAwesomeIcon icon={faUpload} />
                                                        </button>
                                                    </div>

                                                    <div className="attachment-list">
                                                        {attachments.map((file, index) => (
                                                            <div key={index} className="attachment-item d-flex align-items-center gap-2">
                                                                <span>{file.name}</span>
                                                                <button
                                                                    type="button"
                                                                    className="btn btn-sm btn-outline-danger"
                                                                    onClick={() => handleDeleteAttachment(index)}
                                                                >
                                                                    ❌
                                                                </button>
                                                            </div>
                                                        ))}
                                                    </div>

                                                    <div className="attachment-list-2 mt-2">
                                                        {bindedattachments.map((file, index) => (
                                                            <div key={index} className="attachment-item d-flex align-items-center gap-2">
                                                                <span>{file.name}</span>
                                                                <button
                                                                    type="button"
                                                                    className="btn btn-sm btn-outline-danger"
                                                                    onClick={() => DeleteAttachment(index)}
                                                                >
                                                                    ❌
                                                                </button>
                                                            </div>
                                                        ))}
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colSpan={12}>
                                                    <Table
                                                        ref={tableRef}
                                                        initialData={sharedData}
                                                        onDataChange={(updatedData) => setSharedData(updatedData)} />
                                                </td>
                                            </tr>

                                            <tr>
                                                <td colSpan={12}>
                                                    <div className="texth5">Summary</div>
                                                    <table className="table table-bordered" id="summaryDataTable">
                                                        <colgroup>
                                                            <col style={{ width: '15%' }} />
                                                            <col style={{ width: '15%' }} />
                                                            <col style={{ width: '15%' }} />
                                                            <col style={{ width: '15%' }} />
                                                            <col style={{ width: '40%' }} />
                                                        </colgroup>
                                                        <thead>
                                                            <tr>
                                                                <th>User Name</th>
                                                                <th>Next Approver Name</th>
                                                                <th>Action Date</th>
                                                                <th>Action Remarks</th>
                                                                <th>Remarks</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {summaryRows?.map((row, idx) => (
                                                                <tr key={idx}>
                                                                    <td>{row.c1}</td>
                                                                    <td>{row.c2}</td>
                                                                    <td>{row.c3}</td>
                                                                    <td>{row.c4}</td>
                                                                    <td>{row.c5}</td>
                                                                </tr>
                                                            ))}
                                                        </tbody>
                                                        <tbody></tbody>
                                                    </table>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                    </>)}
            </Form>
        </Formik>
    );
};
