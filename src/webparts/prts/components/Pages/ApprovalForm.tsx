import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useLocation } from "react-router-dom";
import { Formik, Form, Field, FormikProps } from "formik";
import type { IPrtsProps } from '../IPrtsProps';
import { IUtilities } from '../../service/BAL/SPCRUD/utilities';
import SPCRUDOPS from "../../service/DAL/spcrudops";
import USESPCRUD, { ISPCRUD } from '../../service/BAL/SPCRUD/spcrud';
import { IPersonaProps } from 'office-ui-fabric-react';
import TableToExcel from '@linways/table-to-excel';
//Date
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react';
//PlantCodeMaster
import { IMDR } from '../../service/INTERFACE/IMDR';
import IASRequestsOps from '../../service/BAL/SPCRUD/PRTS';
//Date
import { format } from 'date-fns';
import { ISPCRUDOPS } from '../../service/DAL/spcrudops';
import '../Pages/CSS/NewRequest.scss';

//Template
import renderTemplateTable from '../../service/BAL/SPCRUD/Template'
//Excel
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { useHistory } from 'react-router-dom';
import * as yup from 'yup';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import Table from '../Pages/Table';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faArrowRight, faUpload } from '@fortawesome/free-solid-svg-icons';
import IEmployeeProfileops from '../../service/BAL/SPCRUD/EmployeeProfile';
import IDelegateApproverops from '../../service/BAL/SPCRUD/DelegateApprover';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import '../Pages/CSS/ApprovalForm.scss';

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
  Comment: String,
  forwardinguser: any
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

export const ApprovalForm: React.FC<IPrtsProps> = (props: IPrtsProps) => {
  const inputDataTableRef = useRef(null);
  const inputDataTable = useRef(null);
  const rowCountRef = useRef(null);
  const [tableBodyHTML, setTableBodyHTML] = useState("");
  const location = useLocation();
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
    Status: '',
    Comment: '',
    forwardinguser: ''
  }


  let spCrudObj: ISPCRUD;
  const history = useHistory();
  //MASTER LIST  
  const [CostCenterdata, setCostCenterdata] = useState([]);//Costcenter list data  

  //MAIN LIST      
  const [BindingWorkflow, setWorkflow] = useState<WorkflowStep[]>([]);
  const [workflowJSX, setWorkflowJSX] = useState(null);
  const [loading, setLoading] = useState(true);
  const [summaryRows, setSummaryRows] = useState([]);
  const [AttachmentFiles, setAttachmentFiles] = useState([]);
  const [Summary, setSummary] = useState([]);
  const [ItemID, setItemID] = useState<number>();
  const [showForwardNew, setForwardNew] = useState(false);
  const [showForward, setShowForward] = useState(false);
  const [actionType, setActionType] = useState<"Reject" | "Rework" | "Withdraw" | "Comment" | null>(null);
  const [ForwardingUser, setForwardingUser] = useState([]);
  const [Buttondisabled, setbuttondisabled] = useState(true);
  const [ItemData, setItemData] = useState([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [MCdata, setMCdata] = useState([]);
  const [WHCdata, setWHCdata] = useState([]);
  const [materialcontroller, setmaterialcontroller] = useState(false);
  const [warehousecontroller, setwarehousecontroller] = useState(false);
  let VisibleButtons = useRef<any[]>([]);
  const rowsPerPage = 10;
  const totalPages = Math.ceil(ItemData?.length / rowsPerPage) || 1;
  const printRef = useRef<HTMLDivElement>(null);

  const handleActionClick = (type: "Reject" | "Rework" | "Withdraw" | "Comment") => {
    setActionType(type);
    setShowForward(true);
  };

  async function ok() {
    setLoading(true);

    const comment = formikRef.current?.values?.Comment?.trim();

    if (!comment) {
      alert("Please enter comments");
      setLoading(false);
      return;
    }
    console.log(actionType, formikRef.current.values.Comment);
    if (actionType === 'Reject') {
      updateFieldBeforeSubmit(99, actionType, actionType, formikRef.current.values.Comment, "0", "0", 0, 0, "");
    }
    else if (actionType === 'Rework') {
      updateFieldBeforeSubmit(0, actionType, actionType, formikRef.current.values.Comment, "0", "0", 0, 0, "");
    }
    else if (actionType === 'Withdraw') {
      updateFieldBeforeSubmit(98, actionType, actionType, formikRef.current.values.Comment, "0", "0", 0, 0, "");
    }
    else if (actionType === 'Comment') {
      let summary = Summary;
      let nextUser;
      const formattedDate = format(new Date(), "dd-MM-yyyy HH:mm");
      let z = updateSummary(props.userDisplayName, nextUser, formattedDate, actionType, formikRef.current.values.Comment);
      summary.push(z);
      let summarynew = JSON.stringify(summary);
      setSummaryRows(JSON.parse(summarynew));
      let fields = {
        Summary: summarynew,
      };
      console.log(fields);
      spCrudObj = await USESPCRUD();

      await spCrudObj.updateData('MDR_List', ItemID, fields, props).then(async (result) => {
        formikRef.current?.setFieldValue('Comment', '')
        setShowForward(false);
        setLoading(false);
      });
    }
    const queryParams = new URLSearchParams(location.search);
    const urlsearch = queryParams.get("ItemId");
    if (location.pathname.includes("ApprovalForm")) {
      await GetMCdata();
      await GetWarehousedata();
      await Binddata(urlsearch, '');
      setLoading(false);
    }
    const tableBody = document.querySelector("#templateTable tbody");
    if (tableBody) {
      tableBody.innerHTML = tableBodyHTML;
    }
  };

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
        alert("Employee ID not found for email:" + Email);
        return null;
      }

      if (EmployeeProfiledata.length > 1) {
        alert("Multiple active employees found with the same email:" + Email);
        return null;
      }

      const empId = EmployeeProfiledata[0]?.EmployeeId;
      if (!empId) {
        alert("EmployeeId field missing for email:" + Email);
        return null;
      }

      return empId;

    } catch (error) {
      console.error("Error fetching Employee ID for " + Email + ":", error);
      return null;
    }
  }

  async function Forward_Update() {
    setLoading(true);
    const email = formikRef.current.values.forwardinguser;
    let usermcdata
    if (materialcontroller === true) {
      usermcdata = MCdata.filter(u => u.UserName.EMail === email)[0].UserName.Title;
    }
    else if (warehousecontroller === true) {
      usermcdata = WHCdata.filter(u => u.UserName.EMail === email)[0].UserName.Title;
    }

    if (email == "") {
      alert("Please select user from the list");
      return;
    }
    let EmployeeID = await GetEmployeeID(email);
    if (EmployeeID != null && EmployeeID != undefined && EmployeeID != '') {
      let wf = BindingWorkflow.filter(m => m.required === true);
      wf[Stage.current].user = usermcdata
      wf[Stage.current].email = email
      wf[Stage.current].EmpID = EmployeeID
      setWorkflow(wf);

      let NA = await (await getuserdata(email)).data.Id;
      let DAId = 0;
      let DelegateDataNAID = await IDelegateApproverops().getDelegateApprover(email, props);
      let NextApproverEmpID = await IEmployeeProfileops().getEmployeeProfile(email, props);
      let DelegateApproverEmpID = "0";
      let tda = DelegateDataNAID;
      console.log(tda);
      if (tda.length > 0) {
        DelegateApproverEmpID = tda[0].DelegateToEmpID;
        DAId = tda[0].DelegateToId;
      }

      updateFieldBeforeSubmit(Stage.current, "Forwarded", "Pending Approval", '', NA, DAId, NextApproverEmpID, DelegateApproverEmpID, wf);
    }
    else {
      formikRef.current?.setFieldValue('forwardinguser', '');
      setLoading(false);
    }
  }

  //Global Variables  
  let InitialStage = useRef(0);
  let Stage = useRef(0);
  let NA = 0;


  async function updateFieldBeforeSubmit(stage, actionTaken, status, reason, NAId, DAId, NextApproverEmpID, DelegateApproverEmpID, newworkflow) {
    let wf = [];
    if (BindingWorkflow) {
      wf = JSON.parse(JSON.stringify(BindingWorkflow));
    }
    let summary = Summary;
    let nextUser;
    if (status !== 'Reject' && status !== 'Withdraw' && status !== 'Rework' && status !== 'Approved and Closed') {
      let requiredwf = wf.filter(m => m.required === true);
      nextUser = (Stage.current != 10) ? requiredwf[stage].user : "";
    }
    const formattedDate = format(new Date(), "dd-MM-yyyy HH:mm");
    let z = updateSummary(props.userDisplayName, nextUser, formattedDate, actionTaken, reason);
    summary.push(z);
    let summarynew = JSON.stringify(summary);
    let fields = {
      NAId: NAId,
      DAId: DAId,
      NextApproverEmpID: NextApproverEmpID[0]?.EmployeeID ?? "0",
      DelegateApproverEmpID: DelegateApproverEmpID.toString(),
      Status: status,
      Stage: stage,
      Summary: summarynew,
      //LastAction: new Date()
    };

    if (newworkflow) {
      fields["ApproverList"] = JSON.stringify(newworkflow)
    }

    console.log(fields);
    spCrudObj = await USESPCRUD();

    try {
      await spCrudObj.updateData('MDR_List', ItemID, fields, props);
      if (actionTaken === 'Forwarded') {
        alert("Request successfully submitted for Forwarded");
      }
      else if (actionTaken === 'Reject') {
        alert("Request successfully submitted for Rejected");
      }
      else if (actionTaken === 'Withdraw') {
        alert("Request successfully submitted for Withdraw");
      }
      else if (actionTaken === 'Rework') {
        alert("Request successfully submitted for Rework");
      }
      else {
        alert("Request Submitted Successfully");
      }
      handleClose();
    } catch (error) {
      Stage.current = InitialStage.current;
      console.error("Failed to submit request:", error);
      alert("Failed to submit request.");
      setLoading(false);
      setbuttondisabled(false);
    }
    finally {
      setLoading(false);
    }

  }

  //Approve Request
  async function reqApproved() {
    setLoading(true);
    setbuttondisabled(true);
    let wf = BindingWorkflow.filter((item) => { return item.required == true });
    let status;
    if ((Stage.current + 1) == wf.length) {
      status = 'Approved and Closed';
    }
    else {
      status = 'Pending Approval';
    }

    if (status == "Approved and Closed") {
      updateFieldBeforeSubmit(100, "Approved", status, formikRef.current.values.Comment, 0, 0, 0, 0, "");
    }
    else {
      for (let i = (Stage.current + 1); i < wf.length; i++) {
        if (wf[i].required == true) {

          let Stage = i;
          let DelegateApproverEmpID = "";
          let DAId = 0;
          NA = await (await getuserdata(wf[i].email)).data.Id;
          let NextApproverEmpID = await IEmployeeProfileops().getEmployeeProfile(wf[i].email, props);
          let DelegateDataNAID = await IDelegateApproverops().getDelegateApprover(wf[i].email, props)
          let tda = DelegateDataNAID;
          console.log(tda);
          if (tda.length > 0) {
            DelegateApproverEmpID = tda[0].DelegateToEmpID;
            DAId = tda[0].DelegateToId;
          }
          updateFieldBeforeSubmit(Stage, "Approved", status, formikRef.current.values.Comment, NA, DAId, NextApproverEmpID, DelegateApproverEmpID, "");
          break;
        }
      }
    }
  }

  //get user data
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

  //Button Visibility
  const showButtons = (arr) => {
    let sg = "";

    let btns = [".btn-draft", ".btn-init", ".btn-withdrawn", ".btn-approver", ".btn-forward"];
    let filtered = btns.filter((btn) => arr.includes(btn));

    VisibleButtons.current = filtered;
  };

  //Summary Update
  const updateSummary = (c1: any, c2: any, c3: any, c4: any, c5: any): { c1: any, c2: any, c3: any, c4: any, c5: any } => {
    let z: { c1: any, c2: any, c3: any, c4: any, c5: any } = {
      c1,
      c2,
      c3,
      c4,
      c5
    };
    return z;
  };

  //List Data of CostCenter
  async function GetMCdata() {
    const spCrudOps = await SPCRUDOPS();
    const MCdata = await spCrudOps.getData(
      'MaterialControllers',
      'UserName/Title,UserName/EMail',
      'UserName',
      '',
      { column: 'ID', isAscending: true },
      props
    );

    setMCdata(MCdata);
  }

  //List Data of Warehouse Controller 
  async function GetWarehousedata() {
    const spCrudOps = await SPCRUDOPS();
    const MCdata = await spCrudOps.getData(
      'WarehouseControllers',
      'UserName/Title,UserName/EMail',
      'UserName',
      '',
      { column: 'ID', isAscending: true },
      props
    );

    setWHCdata(MCdata);
  }


  const formatDate = (dateString) => {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-GB'); // This gives dd/mm/yyyy format
  };


  function searchInTable(searchBoxId: string, tableId: string) {
    const input = document.getElementById(searchBoxId) as HTMLInputElement;
    const filter = input.value.toLowerCase();
    const table = document.getElementById(tableId) as HTMLTableElement;
    const tr = table.getElementsByTagName("tr");

    for (let i = 1; i < tr.length; i++) {
      const rowText = tr[i].innerText.toLowerCase();
      tr[i].style.display = rowText.includes(filter) ? "" : "none";
    }
  }

  function exportReportToExcel(selector: string, fileName: string) {
    const table = document.querySelector(selector) as HTMLElement;
    if (table) {
      TableToExcel.convert(table, {
        name: `${fileName}.xlsx`,
        sheet: { name: 'Sheet 1' }
      });
    }
  }

  const handleClose = () => {
    const lastActive = sessionStorage.getItem('sidebarFrom');
    if (lastActive) {
      history.push(lastActive);
    } else {
      history.push('/InitiatorLanding'); // Fallback route if none found
    }
  };

  async function Binddata(urlsearch, Parameterdata) {
    const itemdata = await IASRequestsOps().getIASDatafilter(urlsearch, props);
    const item = itemdata[0];

    if (!item) return;

    const isCurrentApprover =
      props.EmployeeId[0].EmployeeID === item.NextApproverEmpID ||
      props.EmployeeId[0].EmployeeID === item.DelegateApproverEmpID;

    const isEmployee = props.EmployeeId[0].EmployeeID === item.EmpID;
    let visibleButtons = [".btn-init"];

    if (isCurrentApprover) {
      const approverType = JSON.parse(item.ApproverList).filter(m => m.required === true)[item.Stage].type;
      visibleButtons = [".btn-approver", ".btn-init"];
      if (approverType === "MC" || approverType === "WHC") {
        visibleButtons.push(".btn-forward");
        if (approverType === "MC") {
          setmaterialcontroller(true);
        }
        else if (approverType === "WHC") {
          setwarehousecontroller(true);
        }
      }
    } else if (
      isEmployee &&
      item.Status === "Pending Approval"
    ) {
      visibleButtons = [".btn-withdrawn", ".btn-init"];
    }

    if (isEmployee || props.Appadmin || props.SysAdmin) {
      visibleButtons.push(".btn-print");
    }

    showButtons(visibleButtons);


    // showButtons(visibleButtons);
    VisibleButtons.current = visibleButtons;
    // Bind form values
    Stage.current = item.Stage;
    const summary = JSON.parse(item.Summary);
    setSummary(summary);
    setSummaryRows(summary);

    formikRef.current?.setFieldValue('reqID', item.Title);
    // formikRef.current?.setFieldValue('requesterName', item.EmpName);
    formikRef.current?.setFieldValue('reqDepartment', item.InitDepartment);
    // formikRef.current?.setFieldValue('reqDate', formatDate(item.Date));
    formikRef.current?.setFieldValue('EmployeeId', item.EmpID);
    formikRef.current?.setFieldValue('costCenter', item.CostCenter);
    // formikRef.current?.setFieldValue('CarLine', item.CarLine);
    // formikRef.current?.setFieldValue('movementReason', item.Reason);
    formikRef.current?.setFieldValue('Status', item.Status);
    // formikRef.current?.setFieldValue('ReasonCategory', item.ReasonCategory);

    setItemID(item.ID);
    setWorkflow(JSON.parse(item.ApproverList));

    const attachment = item.AttachmentFiles.map(att => ({
      name: att.FileName,
      url: att.ServerRelativeUrl,
    }));
    setAttachmentFiles(attachment);

    setItemData(JSON.parse(item.Items));
  }

  useEffect(() => {
    const lastActive = sessionStorage.getItem('sidebarFrom');
    console.log('Last active sidebar tab:', lastActive);
  }, [tableBodyHTML]);

  //onload 
  useEffect(() => {
    const runStepByStep = async () => {
      let employeeId = props.EmployeeId[0].EmployeeID;
      const queryParams = new URLSearchParams(location.search);
      const urlsearch = queryParams.get("ItemId");
      if (location.pathname.includes("ApprovalForm")) {
        await GetMCdata();
        await GetWarehousedata();
        //await GetMovementflow();
        await Binddata(urlsearch, '');
        setLoading(false);
      }
    }
    runStepByStep();
  }, []);

  useEffect(() => {
    if (BindingWorkflow.length > 0) {
      displayWorkflow();
    }
  }, [BindingWorkflow, Stage]);

  {/* Add arrow if not the last element */ }
  {/* Only count required items for arrow placement */ }
  // const displayWorkflow = () => {
  //   let wf = [];

  //   BindingWorkflow.forEach((m, i) => {
  //     if (m.required === true) {
  //       const isActive = i === Stage.current ? 'activeApprover' : '';
  //       wf.push(
  //         <React.Fragment key={i}>
  //           <ul className="main-menu">
  //             <li className={`${m.type} ${isActive}`.trim()}>
  //               {m.user}
  //             </li>
  //           </ul>
  //         </React.Fragment>
  //       );
  //     }
  //   });

  //   setWorkflowJSX(wf);
  // };

  const displayWorkflow = () => {
    const wf: JSX.Element[] = [];

    const _wf = BindingWorkflow.filter((item) => item.required === true);
    let isActive;
    let notActive = false;
    _wf.forEach((m, i) => {
      //if (m.required === true) {
      if (notActive === false && Stage.current !== 99) {
        if (Stage.current === i) {
          isActive = 'activeApprover';
          notActive = true;
        }
        else {
          isActive = 'beforeactiveApprover';
        }
      }
      else {
        isActive = 'overrideStage';
      }

      wf.push(
        <ul className="main-menu">
          <li className={`${m.type} ${isActive}`.trim()}>
            {m.user}
          </li>
        </ul>
      );
      //}
    });

    setWorkflowJSX(wf);
  };

  const handlePrint = () => {
    const printContents = printRef.current.innerHTML;
    const originalContents = document.body.innerHTML;

    document.body.innerHTML = printContents;
    window.print();
    document.body.innerHTML = originalContents;
    window.location.reload(); // Optional: to re-render app cleanly
  };

  const exportToExcel = (data) => {
    // Convert data into worksheet
    const worksheet = XLSX.utils.json_to_sheet(
      data.map((row, idx) => ({
        "Sr. No": row.c0,
        "Part Number": row.c1,
        "Part Description": row.c2,
        Supplier: row.c3,
        Qty: row.c4,
        Value: row.c5,
        Amount: row.c6,
        Remarks: row.c7,
      }))
    );

    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");

    // Export as Excel file
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const dataBlob = new Blob([excelBuffer], { type: "application/octet-stream" });
    saveAs(dataBlob, "SummaryData.xlsx");
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
                    <h2>Approval Form</h2>
                  </div>
                </div>
                {/* Add other header elements here if needed */}
              </div>
              <div id="mainContainer">
                <div id="tablemain">
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
                            {Buttondisabled && (
                              <>
                                {VisibleButtons.current.includes(".btn-approver") && (
                                  <>
                                    <button className="btn btn-success btn-approver" type="button" onClick={reqApproved}>
                                      <i className="fa fa-check"></i> Approve
                                    </button>

                                    <a className="btn btn-danger btn-approver" type="button" onClick={() => handleActionClick("Reject")}>
                                      <i className="fa fa-times"></i> Reject
                                    </a>

                                    <a className="btn btn-warning btn-approver" type="button" onClick={() => handleActionClick("Rework")}>
                                      <i className="fa fa-undo"></i> Rework
                                    </a>

                                    <a className="btn btn-warning btn-approver" type="button" onClick={() => handleActionClick("Comment")}>
                                      <i className="fa fa-comments"></i> Comment
                                    </a>

                                  </>
                                )}

                                {VisibleButtons.current.includes(".btn-forward") && (
                                  <a className="btn btn-warning btn-approver btn-forward" type="button" onClick={() => setForwardNew(true)}>
                                    <i className="fa fa-forward"></i> Forward
                                  </a>
                                )}

                              </>
                            )}
                            <a className="btn btn-warning btn-approver btn-forward"
                              type="button" onClick={handleClose}
                            >
                              <i className="fa fa-forward"></i> Close
                            </a>
                            {VisibleButtons.current.includes(".btn-print") && (
                              <a
                                className="btn btn-primary d-print-none"
                                type="button" onClick={handlePrint}
                              >
                                Print Form
                              </a>
                            )}
                            <div className="requestStatus">
                              <span>Status: </span>
                              <span className="displayStatus"><Field name="Status" readOnly style={{border:"none"}}/></span>
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
                          <label>Employee ID</label>
                          <Field name="EmployeeId" readOnly className="form-control" />
                        </td>
                        <td colSpan={6}>
                          <label><span style={{ color: 'red' }}>*</span> Cost Center No.</label>
                          <Field name="costCenter" readOnly className="form-control" />
                        </td>
                        <td colSpan={3}>
                          <label><span style={{ color: 'red' }}>*</span> Reason Category</label>
                          <Field name="ReasonCategory" readOnly className="form-control" />
                        </td>
                      </tr>
                      <tr>
                        <th colSpan={3}><span style={{ color: 'red' }}>*</span> CarLine</th>
                        <td colSpan={9}>
                          <Field rows={4} name="CarLine" readOnly className="form-control" />
                        </td>
                      </tr>
                      <tr>
                        <th colSpan={3}><span style={{ color: 'red' }}>*</span> Description/Remark </th>
                        <td colSpan={9}>
                          <Field as="textarea" rows={4} readOnly name="movementReason" className="form-control large-textarea" />
                        </td>
                      </tr>
                      <tr>
                        <td colSpan={3}>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                            <span className="h4">Approval Note from Buyer</span>
                            <button className="btn btn-warning btn-attachment hidden btn-init" type="button">
                              <i className="fa fa-paperclip"></i>
                            </button>
                          </div>
                        </td>
                        <td colSpan={9}>
                          {AttachmentFiles?.length > 0 ? (
                            AttachmentFiles.map((file, idx) => (
                              <div key={idx}>
                                <a
                                  href={`${props.currentSPContext.pageContext.web.absoluteUrl
                                    .split('/')
                                    .slice(0, 3)
                                    .join('/')}${file.url}`}
                                  target="_blank"
                                  rel="noreferrer"
                                >
                                  {file.name}
                                </a>
                              </div>
                            ))
                          ) : (
                            <span>No attachments</span>
                          )}
                        </td>
                      </tr>

                      <button className="btn btn-warning export-btn" type="button" onClick={() => exportToExcel(ItemData)} style={{ marginLeft: '10px', marginTop: '10px', marginBottom: '10px' }}>Export Data</button>
                      <tr>
                        <td colSpan={12}>
                          <div style={{ height: "50vh", overflowY: "auto", width: "100%" }}>
                            <table className="table table-bordered" id="summaryDataTable">
                              <colgroup>
                                <col style={{ width: '10%' }} />
                                <col style={{ width: '15%' }} />
                                <col style={{ width: '20%' }} />
                                <col style={{ width: '15%' }} />
                                <col style={{ width: '10%' }} />
                                <col style={{ width: '10%' }} />
                                <col style={{ width: '10%' }} />
                                <col style={{ width: '10%' }} />
                              </colgroup>

                              <thead>
                                <tr>
                                  <th>Sr. No</th>
                                  <th>Part Number</th>
                                  <th>Part Description</th>
                                  <th>Supplier</th>
                                  <th>Qty</th>
                                  <th>Value</th>
                                  <th>Amount</th>
                                  <th>Remarks</th>
                                </tr>
                              </thead>

                              <tbody>
                                {ItemData?.slice((currentPage - 1) * rowsPerPage, currentPage * rowsPerPage).map((row, idx) => (
                                  <tr key={idx}>
                                    <td>{row.c0}</td>
                                    <td>{row.c1}</td>
                                    <td>{row.c2}</td>
                                    <td>{row.c3}</td>
                                    <td>{row.c4}</td>
                                    <td>{row.c5}</td>
                                    <td>{row.c6}</td>
                                    <td>{row.c7}</td>
                                  </tr>
                                ))}
                              </tbody>

                              {/* ✅ FOOTER row for Total */}
                              <tfoot>
                                <tr>
                                  <td colSpan={5}></td>
                                  <td className="fw-bold text-end">Total</td>
                                  <td className="fw-bold text-start">
                                    {ItemData?.reduce((sum, row) => sum + parseFloat(row?.c6 || 0), 0).toFixed(2)}
                                  </td>
                                </tr>
                              </tfoot>
                            </table>
                          </div>

                          {/* ✅ PAGINATION below table, centered */}
                          <div className="d-flex justify-content align-items-center gap-3 mt-2">
                            <button
                              className="btn btn-warning"
                              onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
                              disabled={currentPage === 1}
                            >
                              Previous
                            </button>

                            <span className="fw-bold">
                              Page {currentPage} of {Math.ceil(ItemData?.length / rowsPerPage) || 1}
                            </span>

                            <button
                              className="btn btn-warning"
                              onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
                              disabled={currentPage === totalPages}
                            >
                              Next
                            </button>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colSpan={12}>
                          <div className="h5">Summary</div>
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
                                  <td>{row.from || row.c1}</td>
                                  <td>{row.c2}</td>
                                  <td>{row.dt || row.c3}</td>
                                  <td>{row.status || row.c4}</td>
                                  <td>{row.remark || row.c5}</td>
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
            </div >
          </>)}

        {showForward && (
          <>
            <div
              className={`modal fade show d-block`}
              id="Remarks_Container"
              tabIndex={-1}
              role="dialog"
              aria-hidden="false"
            >
              <div className="modal-dialog modal-lg">
                <div className="modal-content">
                  <div className="modal-body">
                    <span className="h4 required">Add Comments *</span>
                    <span className="h4 remarksForTitle" />
                    {/* <textarea rows={5} className="form-control" id="userRemarks" /> */}
                    <Field as="textarea" rows={5} name="Comment" className="form-control" />
                  </div>
                  <div className="modal-footer">
                    <button
                      className="btn btn-default"
                      onClick={() => {
                        console.log("Submit clicked");
                        ok();
                        setShowForward(false);
                      }}
                    >
                      OK
                    </button>
                    <button
                      className="btn btn-default"
                      onClick={() => {
                        setShowForward(false);
                        formikRef.current?.setFieldValue('Comment', '');
                      }}
                    >
                      Cancel
                    </button>
                  </div>
                </div>
              </div>
            </div>
            <div className="modal-backdrop fade show"></div>
          </>
        )}

        {showForwardNew && (
          <>
            <div
              className="modal fade show d-block"
              tabIndex={-1}
              role="dialog"
              aria-hidden="false"
            >
              <div className="modal-dialog modal-lg">
                <div className="modal-content">
                  {materialcontroller && (
                    <div className="modal-body">
                      <span className="h4 required">Forwarding User</span>
                      <Field as="select" name="forwardinguser" className="form-control mt-2">
                        <option value="">Select</option>
                        {MCdata?.map((Vend) => (
                          <option key={Vend.UserName.Title} value={Vend.UserName.EMail}>
                            {Vend.UserName.Title}
                          </option>
                        ))}
                      </Field>
                    </div>
                  )}

                  {warehousecontroller && (
                    <div className="modal-body">
                      <span className="h4 required">Forwarding User</span>
                      <Field as="select" name="forwardinguser" className="form-control mt-2">
                        <option value="">Select</option>
                        {WHCdata?.map((Vend) => (
                          <option key={Vend.UserName.Title} value={Vend.UserName.EMail}>
                            {Vend.UserName.Title}
                          </option>
                        ))}
                      </Field>
                    </div>
                  )}

                  <div className="modal-footer">
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={() => {
                        console.log("Submit clicked");
                        Forward_Update(); // make sure this uses Formik's values if needed
                        setForwardNew(false);
                      }}
                    >
                      OK
                    </button>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={() => setForwardNew(false)}
                    >
                      Cancel
                    </button>
                  </div>
                </div>
              </div>
            </div>
            <div className="modal-backdrop fade show" />
          </>
        )}

        <div ref={printRef} className="print-area">

          {/* Top Header */}
          <div className="header" style={{ display: "flex", alignItems: "center", marginBottom: "20px" }}>
            <div className="logo" style={{ marginRight: "20px" }}>
              <img
                src="../SiteAssets/Custom/images/MG-Motor-Logo.png"
                alt=""
                style={{ height: "46px", width: "46px" }}
              />
            </div>
            <div className="headerTitle">
              <div className="companyTitle"
                style={{
                  fontSize: "20px",
                  fontWeight: "900",
                  color: "#000",
                  fontFamily: "Arial, Helvetica, sans-serif"
                }}
              >
                JSW MG Motor India Pvt. Ltd.
              </div>
              <div className="formTitle"
                style={{
                  fontSize: "20px",
                  fontWeight: "900",
                  color: "#000",
                  fontFamily: "Arial, Helvetica, sans-serif"
                }}
              >
                Material Demand Request
              </div>
            </div>
          </div>

          {/* Status Display
          <div style={{ float: "right", marginBottom: "10px" }}>
            <b>Status:</b>&nbsp;
            <Field
              name="reqID"
              readOnly
              style={{
                border: 'none',
                background: 'transparent',
                fontWeight: 'bold'
              }}
            />
          </div> */}

          {/* Form Table */}
          <table className="table table-bordered">
            <tbody>
              <tr>
                <td colSpan={3}>
                  <label>Application Track No.</label>
                  <Field name="reqID" readOnly className="form-control" />
                </td>
                <td colSpan={3}>
                  <label>Employee Name</label>
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
                  <label>Employee ID</label>
                  <Field name="EmployeeId" readOnly className="form-control" />
                </td>
              </tr>

              <tr>
                <td colSpan={3}>
                  <label>Employee ID</label>
                  <Field name="EmployeeId" className="form-control" />
                </td>
                <td colSpan={6}>
                  <label>Cost Center No.</label>
                  <Field name="costCenter" className="form-control" />
                </td>
              </tr>

              <tr>
                <th colSpan={3}>
                  <span style={{ color: "red" }}>*</span> CarLine
                </th>
                <td colSpan={9}>
                  <Field name="CarLine" className="form-control" />
                </td>
              </tr>

              <tr>
                <th colSpan={3}>
                  <span style={{ color: "red" }}>*</span> Reason
                </th>
                <td colSpan={9}>
                  {/* <Field as="textarea" name="movementReason" rows={4} className="form-control" /> */}
                  <Field name="movementReason">
                    {({ field }) => (
                      <div>{field.value || "-"}</div>
                    )}
                  </Field>
                </td>
              </tr>


              {/* Summary Table */}
              <tr>
                <td colSpan={12}>
                  <div style={{ pageBreakInside: "avoid", breakInside: "avoid" }}>
                    <table
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        border: "1px solid black",
                      }}
                    >
                      <colgroup>
                        <col style={{ width: "10%" }} />
                        <col style={{ width: "15%" }} />
                        <col style={{ width: "20%" }} />
                        <col style={{ width: "15%" }} />
                        <col style={{ width: "10%" }} />
                        <col style={{ width: "10%" }} />
                        <col style={{ width: "10%" }} />
                        <col style={{ width: "10%" }} />
                      </colgroup>
                      <thead>
                        <tr>
                          {["Sr. No", "Part Number", "Part Description", "Supplier", "Qty", "Value", "Amount", "Remarks"].map(
                            (header, idx) => (
                              <th
                                key={idx}
                                style={{
                                  border: "1px solid black",
                                  padding: "4px",
                                  backgroundColor: "#f2f2f2",
                                  textAlign: "center",
                                }}
                              >
                                {header}
                              </th>
                            )
                          )}
                        </tr>
                      </thead>
                      <tbody>
                        {ItemData?.length ? (
                          <>
                            {ItemData.map((row, idx) => (
                              <tr key={idx}>
                                {Object.values(row).map((cell, i) => (
                                  <td
                                    key={i}
                                    style={{ border: "1px solid black", padding: "4px", textAlign: i >= 4 && i <= 6 ? "right" : "left" }}
                                  >
                                    {cell}
                                  </td>
                                ))}
                              </tr>
                            ))}
                            <tr>
                              <td colSpan={5} style={{ border: "1px solid black" }}></td>
                              <td style={{ border: "1px solid black", fontWeight: "bold", textAlign: "right" }}>Total</td>
                              <td style={{ border: "1px solid black", fontWeight: "bold", textAlign: "right" }}>
                                {ItemData?.reduce((sum, row) => sum + parseFloat(row?.c6 || 0), 0).toFixed(2)}
                              </td>
                              <td style={{ border: "1px solid black" }}></td>
                            </tr>
                          </>
                        ) : (
                          <tr>
                            <td colSpan={8} style={{ border: "1px solid black", textAlign: "center", padding: "8px" }}>
                              No data available
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </td>
              </tr>

              {/* Approval History Table */}
              <tr>
                <td colSpan={12}>
                  <div className="h5">Summary</div>
                  <table className="table table-bordered">
                    <colgroup>
                      <col style={{ width: "15%" }} />
                      <col style={{ width: "15%" }} />
                      <col style={{ width: "15%" }} />
                      <col style={{ width: "15%" }} />
                      <col style={{ width: "40%" }} />
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
                      {summaryRows?.length ? (
                        summaryRows.map((row, idx) => (
                          <tr key={idx}>
                            <td>{row.from || row.c1}</td>
                            <td>{row.c2}</td>
                            <td>{row.dt || row.c3}</td>
                            <td>{row.status || row.c4}</td>
                            <td>{row.remark || row.c5}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={5}>No approval history available</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>

          {/* Security Stamp */}
          <div
            style={{
              marginTop: "40px",
              display: "flex",
              justifyContent: "flex-start",
              pageBreakInside: "avoid",
              breakInside: "avoid",
              breakBefore: "auto"
            }}
          >
            <div
              style={{
                border: "1px solid black",
                width: "150px",
                height: "150px",
                textAlign: "center",
                paddingTop: "20px",
                fontWeight: "bold",
                fontSize: "14px"
              }}
            >
              Security Stamp<br />
              Material Out From Store
            </div>
          </div>
        </div>

      </Form >
    </Formik >
  );
};
