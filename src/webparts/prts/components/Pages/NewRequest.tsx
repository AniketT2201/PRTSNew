import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useLocation , useHistory} from "react-router-dom";
import { Formik, Form, Field, FormikProps, FormikHelpers } from "formik";
import type { IPrtsProps } from '../IPrtsProps';
import { IUtilities } from '../../service/BAL/SPCRUD/utilities';
import SPCRUDOPS from "../../service/DAL/spcrudops";
import USESPCRUD, { ISPCRUD } from '../../service/BAL/SPCRUD/spcrud';
import { IPersonaProps } from 'office-ui-fabric-react';
//Date
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react';
//PlantCodeMaster
import { IIAS } from '../../service/INTERFACE/IPRTS';
import PrtsRequestsOps from '../../service/BAL/SPCRUD/PRTS';
//Date
import { format } from 'date-fns';
import { ISPCRUDOPS } from '../../service/DAL/spcrudops';
import '../Pages/CSS/NewRequest.scss';
//Template
import renderTemplateTable from '../../service/BAL/SPCRUD/Template'
//Excel
import * as XLSX from "xlsx";
import TableToExcel from '@linways/table-to-excel';
// import { useHistory } from 'react-router-dom';
import * as yup from 'yup';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import IEmployeeProfileops from '../../service/BAL/SPCRUD/EmployeeProfile';
import IDelegateApproverops from '../../service/BAL/SPCRUD/DelegateApprover';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUpload, faTrash } from "@fortawesome/free-solid-svg-icons";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Modal, Button, Tabs, Tab } from "react-bootstrap";
import { PRTSService } from '../../service/BAL/SPCRUD/PRTSService';
import 'bootstrap/dist/css/bootstrap.min.css';


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
  movementReason: any
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

export const NewRequest: React.FC<IPrtsProps> = (props: IPrtsProps) => {
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
    movementReason: ''
  }

  let spCrudObj: ISPCRUD;
  const history = useHistory();
  const fileInputRef = useRef<HTMLInputElement>(null);
  //MASTER LIST
  const [MovementDropdown, setMovementDropdown] = useState([]);//Movementflow list data
  const [CostCenterdata, setCostCenterdata] = useState([]);//Costcenter list data
  const [ParameterDetails, setParameterDetails] = useState([]);//Parameter Data
  const [EmployeeData, setEmployeeData] = useState([]); //Employee Department from Employee Profile
  //MAIN LIST
  const [PrtsData, setPrtsData] = useState([]); //Prts data as per request
  const [rid, setrid] = React.useState<any>();//itemid
  const [visibleButtons, setVisibleButtons] = useState([]);//Handle Button Visibility        
  const [BindingWorkflow, setWorkflow] = useState<WorkflowStep[]>([]);
  const [workflowJSX, setWorkflowJSX] = useState(null);
  const [ApprovalNoteNo, setApprovalNoteNo] = useState<string>("");
  const [attachments, setAttachments] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [Buttondisable, setButtondisable] = useState(true);


  const [agency, setAgency] = useState('');
  const [status, setStatus] = useState('');
  const [remarks, setRemarks] = useState('');
  const [activeTab, setActiveTab] = useState("D3_Tab1");
  const [showModal, setShowModal] = useState(false);
  // const [remarks, setRemarks] = useState("");
  // Placeholder array for agencies
   const location = useLocation();
  const navigate = useHistory();
const [activeKey, setActiveKey] = useState<string>("tab1");



 const agencies = ['Agency 1', 'Agency 2', 'Agency 3']; // Replace with real data
  const handleUpdate = () => {
    // Add validation or submission logic here
    console.log({ agency, status, remarks });
  };
  //Global Variables
  let Details = useRef("");
  let Stage = useRef(0);
  let Summary = useRef("");
  let newworkflow = useRef<WorkflowStep[]>([]);
  let uploadedFileKey = useRef<string[]>([]);
  let updateInitiatordata = useRef<any[]>([]);
  let DelegateData = useRef([]);
  //for Formik
  function getFieldProps(formik: FormikProps<any>, field: string) {
    return { ...formik.getFieldProps(field), errorMessage: formik.errors[field] as string };
  }

  //List Data of MovementFlow
  async function GetMovementflow() {
    const spCrudOps = await SPCRUDOPS();
    const Momentflowdata = await spCrudOps.getData(
      'MovementFlow',
      '*,ApprovalNoteDescription,Title,ID',
      '',
      '',
      { column: 'ID', isAscending: true },
      props
    );
    setMovementDropdown(Momentflowdata);
    getParameterDetails();
  }

  //List Data of Parameter List
  async function getParameterDetails() {
    const spCrudOps = await SPCRUDOPS();
    const Parameterdata = await spCrudOps.getData(
      'Parameters',
      'Id,Title,Details',
      '',
      '',
      { column: 'ID', isAscending: true },
      props
    );

    setParameterDetails(Parameterdata);
  }

  //List Data of CostCenter
  // async function GetCostCenterdata() {
  //   const spCrudOps = await SPCRUDOPS();
  //   const CostCenterdata = await spCrudOps.getRootData(
  //     'CostCenter',
  //     'Description,Title,ID,CCOwner/Id,CCOwner/Title,CCOwner/EMail,FinApprover/Id,FinApprover/Title,FinApprover/EMail',
  //     'CCOwner,FinApprover',
  //     '',
  //     { column: 'ID', isAscending: true },
  //     props
  //   );

  //   setCostCenterdata(CostCenterdata);
  // }

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
    setAttachments(prev => prev.filter((_, i) => i !== index));
  };

  //List Data of Prts using Id
  const GetPrtsData = async (id) => {
    const PrtsColl = await PrtsRequestsOps().getIASDatafilter(id, props);
    setPrtsData(PrtsColl);
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
    // if (PrtsData && PrtsData.stage !== undefined) {
    //   sg = Number(PrtsData.stage) || "";
    // }

    let btns = [".btn-draft", ".btn-init", ".btn-withdrawn", ".btn-approver"];
    let filtered = btns.filter((btn) => arr.includes(btn));

    // Remove .btn-forward if stage is 1
    // if (sg === 1) {
    //   filtered = filtered.filter((btn) => btn !== ".btn-forward");
    // }

    setVisibleButtons(filtered);
  };

  //draft
  // const createdraft = () => {
  //   let z = updateSummary(props.userDisplayName, "", format(new Date(), "dd-MM-yyyy HH:mm"), "Request Created", "");
  //   let summary = [z];
  //   let defaultWF =
  //     `[{"type":"Initiator","user":"` + props.userDisplayName + `}","email":"` + props.userEmail + `","required":true,"department":"` + EmployeeDept + `"},
  //   {"type":"Mgr","user":"`+ props.userDisplayName + `","email":"` + props.userEmail + `","required":true},
  //   {"type":"CCOwner","user":"","email":"","required":false},
  //   {"type":"CCApprover","user":"","email":"","required":false},
  //   {"type":"MaterialPlanner","user":"","email":"","required":false},
  //   {"type":"Warehouse","user":"","email":"","required":false},
  //   {"type":"InventoryMgr","user":"","email":"","required":false},
  //   {"type":"SCMHead","user":"","email":"","required":false},
  //   {"type":"Finance","user":"","email":"","required":false},
  //   {"type":"Inventory","user":"","email":"","required":false}]`.replaceAll("\t", "").replaceAll("\n", "");

  //   let fields = {
  //     'Title': 'Draft',
  //     'Summary': JSON.stringify(summary),
  //     'WF': defaultWF
  //   }
  //   NewSaveData("Prts_List", fields, ['home']);
  // }

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

  //NewSaveData
  async function NewSaveData(listName, fields, actions) {
    spCrudObj = await USESPCRUD();
    await spCrudObj.insertData(listName, fields, props).then(async (brrInsertResult) => {
      console.log(`New Item Created ` + brrInsertResult.data.Id + ` details submitted successfully`);

      actions.forEach((n) => {
        if (n == "home") {
          props.ItemID = brrInsertResult.data.Id;
          fetchData();
        }
        // else { eval(n); }
      });
    });
  }

  //Updating Fields



  //*******************************************************************//
  //*******************Excel Functions*********************************//

  //Import Template and Opening File Explorer to attach file
  async function importTemplate() {
    let mt = formikRef.current.values.movementType;
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
    if (mt != "") {
      await loadTemplate();
      let excelfile = document.getElementById('excelfile');
      // excelfile.click();
      setTimeout(() => {
        excelfile.click();
      }, 2000);
    }
    else {
      alert("Please select movement type, and apply the template to import");
    }
  }

  let retryCount = 0;
  async function costcenterChange() {
    const movementType = formikRef.current?.values.movementType;
    // let confirm = window.confirm("If data is uploaded Imported Template and data will Reset. Do you want to proceed");
    // if (confirm) {
    await movementchange();
    formikRef.current.setFieldValue('uniquePartImpact', '');
    formikRef.current.setFieldValue('grossValue', '');
    formikRef.current.setFieldValue('netValue', '');
    setWorkflow([]);
    setWorkflowJSX([]);
    if (movementType) {
      updateWF(movementType);
    }
    //}
  }

  async function movementchange() {
    const tableEl = document.getElementById("inputDataTable");
    const rowCountEl = document.getElementById("inputDataTableRowCount");

    if (!tableEl || !rowCountEl) {
      if (retryCount < 10) {
        console.warn("Elements not found. Retrying...");
        retryCount++;
        setTimeout(movementchange, 100);
      } else {
        console.error("Max retries reached. Aborting movementchange.");
      }
      return;
    }

    tableEl.innerHTML = '';
    rowCountEl.innerHTML = '';
    Details.current = '';
    retryCount = 0; // Reset for next time
  }

  //Loading Template on the Screen 
  async function loadTemplate() {
    Details.current = '';
    let movementNo = formikRef.current.values.movementType;
    if (movementNo != "") {
      let template = JSON.parse(ParameterDetails.filter(o => o.Title == "MovementTemplate")[0].Details).filter(m => m.movement == movementNo)[0]?.template || "";
      if (template != "") {
        let table = await renderTemplateTable(template);
        const container = document.getElementById("inputDataTable");
        if (container) {
          container.innerHTML = table;
        } else {
          console.error("Div with id='input' not found.");
        }
      }
    }
  }

  //After attaching of file below are the functions
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();
    const regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;

    if (!regex.test(fileName)) {
      console.log("Please upload a valid Excel file!");
      return;
    }

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json: ExcelRecord[] = XLSX.utils.sheet_to_json(sheet, { raw: false });

      if (json.length > 0) {
        isValidExcelUpload(json);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  //Save data of Excel to list and displaying it on screen
  const isValidExcelUpload = async (exceljson: any[]) => {
    try {
      if (!exceljson || exceljson.length === 0) {
        alert("Excel file is empty.");
        return;
      }

      let Updatedexceljson = exceljson.sort((a, b) => Number(a["Sr.No."]) - Number(b["Sr.No."]));
      const movementNo = formikRef.current.values.movementType;
      const costCenter = formikRef.current.values.costCenter;

      // Get template and field names
      const movementTemplate = ParameterDetails.find(p => p.Title === "MovementTemplate");
      const movementTemplateFields = ParameterDetails.find(p => p.Title === "MovementTemplateFields");

      const template = JSON.parse(movementTemplate?.Details || "[]").find((m: any) => m.movement === movementNo)?.template;
      const fieldNames = JSON.parse(movementTemplateFields?.Details || "[]").find((m: any) => m.template === template)?.fields;

      uploadedFileKey.current = Object.keys(exceljson[0]);
      const uploadedFileKeyString = JSON.stringify(uploadedFileKey.current).replace(/"/g, '');

      if (uploadedFileKeyString !== fieldNames) {
        alert("Invalid template file uploaded.");
        return;
      }

      //Validate cost center for specific movement types
      if (["201", "202"].includes(movementNo)) {
        if (!costCenter) {
          alert("Cost Center is mandatory for movements [201, 202]");
          if (fileInputRef.current) {
            fileInputRef.current.value = '';
          }
          return;
        }

        const uploadedCostCenters = [...new Set(exceljson.map(o => o["Cost Center"]?.toUpperCase()))];
        const selectedCostCenter = costCenter.toUpperCase();

        if (uploadedCostCenters.length > 1 || uploadedCostCenters[0] !== selectedCostCenter) {
          alert("Invalid Cost Center found in uploaded file.");
          return;
        }
      }

      //Prepare workflow before processing table
      const updatedworkflow = await updateWF(movementNo);
      if (!updatedworkflow) {
        alert("Failed to update workflow.");
        return;
      }
      setWorkflow(updatedworkflow);
      newworkflow.current = updatedworkflow;

      //Process Excel data into table
      let details: any[] = [];
      let tr = "";

      Updatedexceljson.forEach((row) => {
        let z: any = {};
        tr += "<tr>";
        uploadedFileKey.current.forEach((key, j) => {
          let cell = (row[key] || "").replace(/\n/g, "").replace(/\t/g, "");
          tr += `<td>${cell}</td>`;
          z[`c${j}`] = cell;
        });
        tr += "</tr>";
        details.push(z);
      });

      const tableBody = document.querySelector("#templateTable tbody");
      if (tableBody) tableBody.innerHTML = tr;

      const rowCount = document.getElementById("inputDataTableRowCount");
      if (rowCount) {
        rowCount.innerHTML = `<div class="p-2 bg-secondary text-white font-weight-bold">Total Rows : ${details.length}</div>`;
      }

      formikRef.current?.setFieldValue('uniquePartImpact', details.length.toString());

      if (movementNo === "102") {
        calculateReversalValueColumns(details);
      } else if (
        ["201", "202", "309", "310", "261", "262", "551", "552", "541", "542", "501", "502", "701", "702", "COGI Correction"].includes(movementNo)
      ) {
        calculateQtyPriceColumns(details);
      }

    } catch (err: any) {
      console.error("Workflow update failed:", err);
      alert(`Error occurred while preparing workflow: ${err.message}`);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };


  const calculateReversalValueColumns = (details) => {
    Details.current = JSON.stringify(details);
    const table = document.getElementById("templateTable") as HTMLTableElement;
    const dataAttr = table?.dataset;
    if (!dataAttr) return;

    const totalCol = parseInt(dataAttr.totalcol || "0");
    const qtyCol = parseInt(dataAttr.qtycol || "0");
    if (!totalCol) return;

    let sum = 0;
    let grossVal = 0;
    let totalQuantity = 0;

    const rows = table.tBodies[0]?.rows;
    if (!rows) return;

    Array.from(rows).forEach((row) => {
      const cell = row.cells[totalCol - 1];
      const value = parseFloat(cell.textContent?.trim() || "0");
      const qty = parseFloat(row.cells[qtyCol - 1]?.textContent?.trim() || "0");

      if (value < 0) {
        cell.classList.add("val-red");
      }
      totalQuantity += qty;
      sum += value;
      grossVal += Math.abs(value);
    });

    // formikRef.current?.setFieldValue('grossValue', grossVal.toFixed(2));
    // formikRef.current?.setFieldValue('netValue', sum.toFixed(2));

    const roundToNearestTens = (num: number) => {
      return Math.round(num / 10) * 10;
    };

    const formatNumber = (num: number) => {
      return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    const grossRounded = roundToNearestTens(grossVal);
    const netRounded = roundToNearestTens(sum);

    formikRef.current?.setFieldValue('grossValue', formatNumber(grossRounded));
    formikRef.current?.setFieldValue('netValue', formatNumber(netRounded));

    const footerCellQty = table.tFoot?.rows[0]?.cells[parseInt(dataAttr.qtycol || "0") - 1];
    if (footerCellQty) {
      footerCellQty.textContent = totalQuantity.toString();
    }

    const footerCell = table.tFoot?.rows[0]?.cells[parseInt(dataAttr.totalcol || "0") - 1];
    if (footerCell) {
      footerCell.textContent = sum.toFixed(2);
    }
  };

  const calculateQtyPriceColumns = (details) => {
    let exceldata = details
    const table = document.getElementById("templateTable") as HTMLTableElement;
    const dataAttr = table?.dataset;
    if (!dataAttr) return;

    const qtyCol = parseInt(dataAttr.qtycol || "0");
    const priceCol = parseInt(dataAttr.pricecol || "0");
    const totalCol = parseInt(dataAttr.totalcol || "0");
    if (!qtyCol || !priceCol || !totalCol) return;

    let sum = 0;
    let grossVal = 0;
    let totalQuantity = 0;

    const rows = table.tBodies[0]?.rows;
    if (!rows) return;

    Array.from(rows).forEach((row, index) => {
      const qty = parseFloat(row.cells[qtyCol - 1]?.textContent?.trim() || "0");
      const price = parseFloat(row.cells[priceCol - 1]?.textContent?.trim() || "0");
      const total = qty * price || 0;
      const key = `c${priceCol}`;
      exceldata[index][key] = total.toString();

      const totalCell = row.cells[totalCol - 1];
      totalCell.textContent = total.toFixed(2);

      if (total < 0) {
        totalCell.classList.add("val-red");
      }
      totalQuantity += qty;
      sum += total;
      grossVal += Math.abs(total);
    });

    const roundToNearestTens = (num: number) => {
      return Math.round(num / 10) * 10;
    };

    const formatNumber = (num: number) => {
      return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    const grossRounded = roundToNearestTens(grossVal);
    const netRounded = roundToNearestTens(sum);

    formikRef.current?.setFieldValue('grossValue', formatNumber(grossRounded));
    formikRef.current?.setFieldValue('netValue', formatNumber(netRounded));

    const footerCellQty = table.tFoot?.rows[0]?.cells[parseInt(dataAttr.qtycol || "0") - 1];
    if (footerCellQty) {
      footerCellQty.textContent = totalQuantity.toString();
    }

    const footerCell = table.tFoot?.rows[0]?.cells[parseInt(dataAttr.totalcol || "0") - 1];
    if (footerCell) {
      footerCell.textContent = sum.toFixed(2);
    }

    Details.current = JSON.stringify(exceldata);
  };
  //*****************************Excel functions End******************/
  //******************************************************************/


  const updateWF = async (movementNo: string): Promise<WorkflowStep[]> => {
    let nccowner = "CC Owner";
    let eccowner = "";

    let nccapprover = "CC Approver";
    let eccapprover = "";

    const selectedCostCenter = formikRef.current.values.costCenter;

    if (["201", "202", "551", "552"].includes(movementNo) && selectedCostCenter) {
      const ccItem = CostCenterdata.find((o) => o.Title.trim() === selectedCostCenter);
      nccowner = ccItem?.CCOwner?.Title || "Undefined";
      eccowner = ccItem?.CCOwner?.EMail || "Undefined";
      nccapprover = ccItem?.FinApprover?.Title || "Undefined";
      eccapprover = ccItem?.FinApprover?.EMail || "Undefined";
    }

    const mf = MovementDropdown.find((o) => o.Title === movementNo);
    const wf: WorkflowStep[] = [...BindingWorkflow];
    if (!mf) return wf;

    const getParamUser = (key: string) =>
      JSON.parse(ParameterDetails.find((o) => o.Title === key)?.Details || "{}");

    // Fetch and validate employee IDs
    const initiatorEmpID = await GetEmployeeID(props.userEmail);
    if (!initiatorEmpID) throw new Error(`Failed to fetch EmpID for Initiator: ${props.userEmail}`);

    const mgrEmpID = await GetEmployeeID(EmployeeData[0].DirectManagerName.EMail);
    if (!mgrEmpID) throw new Error(`Failed to fetch EmpID for Manager: ${EmployeeData[0].DirectManagerName.EMail}`);

    let ccOwnerEmpID;
    let ccApproverEmpID;
    if (["201", "202", "551", "552"].includes(movementNo) && selectedCostCenter) {
      ccOwnerEmpID = await GetEmployeeID(eccowner);
      if (!ccOwnerEmpID) throw new Error(`Failed to fetch EmpID for CC Owner: ${eccowner}`);

      const ccApproverEmpID = await GetEmployeeID(eccapprover);
      if (!ccApproverEmpID) throw new Error(`Failed to fetch EmpID for CC Approver: ${eccapprover}`);
    }

    wf[0] = { type: "Initiator", user: props.userDisplayName, email: props.userEmail, required: true, EmpID: initiatorEmpID };
    wf[1] = {
      type: "Mgr",
      user: EmployeeData[0].DirectManagerName.Title,
      email: EmployeeData[0].DirectManagerName.EMail,
      required: true,
      EmpID: mgrEmpID
    };
    wf[2] = { type: "CCOwner", user: nccowner, email: eccowner, required: !!mf.CostCenter, EmpID: ccOwnerEmpID };
    wf[3] = { type: "CCApprover", user: nccapprover, email: eccapprover, required: !!mf.CostCenter, EmpID: ccApproverEmpID };

    const steps = ["MaterialPlanner", "Warehouse", "InventoryManager", "SCMHead", "Finance", "Inventory"];
    for (let i = 0; i < steps.length; i++) {
      const step = steps[i];
      const user = getParamUser(step);
      let empID;
      if (!!mf[step] === true) {
        empID = await GetEmployeeID(user.email);

        if (!empID) throw new Error(`Failed to fetch EmpID for ${step}: ${user.email}`);
      }
      wf[i + 4] = {
        type: step,
        user: user.user || "",
        email: user.email || "",
        required: !!mf[step],
        EmpID: empID
      };
    }

    return wf;
  };

  //List Data of User Master
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

  //fetchdata
  const fetchData = async () => { 
    try {
      let Initiatordata = await EmployeeProfile(props.userEmail);
      updateInitiatordata.current = Initiatordata;
      setEmployeeData(Initiatordata);
     // await GetCostCenterdata();
      // await GetMovementflow();
      // await getParameterDetails();
      showButtons([".btn-init"]);
      formikRef.current?.setFieldValue('requesterName', props.userDisplayName);
      if (Initiatordata.length > 0) {
        formikRef.current?.setFieldValue('reqDepartment', Initiatordata[0].DepartmentCode.Department);
      }
    } catch (error) {
      console.error("Failed to fetch ACL data:", error);
    }
  };

  //onload 
  useEffect(() => {
    fetchData();
  }, []);

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
        const isActive = i === Stage.current ? 'activeApprover' : '';
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

  const handleAssign = () => {
    // TODO: Replace with actual submit logic
    console.log("Champion assigned");
  };

  

  async function updateFieldBeforeSubmit(stage, actionTaken, status, reason, approvalnotenumber) {
    try {
      let summary = [];
      let wf;
      if (Summary) {
        summary = JSON.parse(Summary.current || "[]");
      }
      if (BindingWorkflow) {
        wf = JSON.parse(JSON.stringify(BindingWorkflow));
      }

      let NA = [], DA = [];
      let NAId = 0, DAId = 0;
      let newstatus = "", newstage = 0;
      let NextApproverEmpID = "", DelegateApproverEmpID;

      if (stage == 0) {
        newstatus = status;
        newstage = 1;
        NA.push(await getuserdata(wf[1].email));
        NAId = NA[0].data.Id;
        NextApproverEmpID = await GetEmployeeID(wf[1].email);
        let DelegateDataNAID = await IDelegateApproverops().getDelegateApprover(wf[1].email, props);
        if (Array.isArray(DelegateDataNAID) && DelegateDataNAID.length > 0) {
          DelegateApproverEmpID = DelegateDataNAID[0].DelegateToEmpID;
          DAId = DelegateDataNAID[0].DelegateToId;
        }
      }

      const formattedDate = format(new Date(), "dd-MM-yyyy HH:mm");
      let nextUser = (Stage.current != 10) ? wf[Stage.current].user : "";
      let z = updateSummary(props.userDisplayName, nextUser, formattedDate, actionTaken, reason);
      summary.push(z);
      let summarynew = JSON.stringify(summary);
      const fields = {
        "NAId": NAId,
        "DAId": DAId,
        "NextApproverEmpID": NextApproverEmpID,
        "InitiatorEmpId": updateInitiatordata.current[0].EmployeeId,
        "DelegateApproverEmpID": DelegateApproverEmpID ?? "",
        "Status": newstatus,
        "Stage": newstage,
        "Summary": summarynew,
        "LastAction": new Date(),
        "MovementType": formikRef.current.values.movementType,
        "HeaderName": JSON.stringify(uploadedFileKey.current),
        "Details": Details.current,
        "WF": JSON.stringify(newworkflow.current),
        "CostCenter": formikRef.current.values.costCenter,
        "UniquePartImpact": formikRef.current.values.uniquePartImpact,
        "GrossValue": formikRef.current.values.grossValue,
        "NetValue": formikRef.current.values.netValue,
        "Department": formikRef.current.values.reqDepartment
      };

      if (stage === 0) {
        fields["ApprovalNoteNo"] = approvalnotenumber;
        fields["MovementReason"] = formikRef.current.values.movementReason;
      }

      if (actionTaken === "Cost Center Change") {
        fields["CostCenter"] = formikRef.current.values.costCenter;
      }

      if (status === "Approved and Closed") {
        fields["SAPNo"] = formikRef.current.values.sapNo;
      }

      setLoading(true);
      setButtondisable(false);

      const spCrudObj = await USESPCRUD();

      // Step 1: Insert with only Title to get the ID
      const insertTempResult = await spCrudObj.insertData('Prts_List', { Title: 'Temp' }, props);
      const rid = insertTempResult.data.Id;
      const finalTitle = `Prts/${format(new Date(), 'yyyy')}/${rid.toString().padStart(6, '0')}`;

      // Step 2: Upload attachments first
      if (attachments.length > 0) {
        for (const file of attachments) {
          try {
            await spCrudObj.addAttchmentInList(file, 'Prts_List', rid, file.name, props);
            console.log(`Attachment ${file.name} uploaded.`);
          } catch (error) {
            console.error(`Failed to upload attachment ${file.name}:`, error);
            alert(`Failed to upload attachment ${file.name}`);
            setButtondisable(true);
            setLoading(false);
            return;
          }
        }
      }

      // Step 3: Update fields with final data (including real Title)
      fields["Title"] = finalTitle;
      await spCrudObj.updateData('Prts_List', rid, fields, props);

      alert(`Request ${finalTitle} Submitted Successfully`);
      history.push('/InitiatorLanding');

    } catch (error) {
      console.error('Submission failed:', error);
      alert('An error occurred during submission.');
      setButtondisable(true);
    } finally {
      setLoading(false);
    }
  }

  function downloadTemplate() {
    const mt = formikRef.current?.values?.movementType;

    if (mt && mt.trim() !== "") {
      loadTemplate();

      // Use a more reliable async mechanism (optional but better than setTimeout)
      setTimeout(() => {
        const table = document.querySelector('#templateTable') as HTMLTableElement | null;

        if (table) {
          const tfoot = table.querySelector('tfoot');
          if (tfoot) tfoot.remove();

          exportReportToExcel('#templateTable', `${mt}-MovementTemplate`);
        }
      }, 2000);
    } else {
      alert("Please select movement type, and click 'Apply' button before downloading the template.");
    }
  }


  function exportData() {
    const an = formikRef.current?.values?.approvalNoteNo?.trim().replaceAll("/", "-") || "export";
    const mt = formikRef.current?.values?.movementType || "Movement";

    if (Details.current.length > 0) {
      const fileName = `${mt}-${an}`;
      exportReportToExcel('#templateTable', fileName);
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
  async function submitRequest(formstatus) {
    setButtondisable(false);
    const values = formikRef.current.values;
    const movementType = values.movementType;
    const missingFields = [];

    // Field name to label map
    const fieldLabels = {
      uniquePartImpact: "Unique Part Impact",
      grossValue: "Gross Value",
      netValue: "Net Value",
      movementType: "Movement Type",
      costCenter: "Cost Center",
      movementReason: "Movement Reason"
    };

    // General mandatory fields
    const baseFields = ["movementType"];
    baseFields.forEach(field => {
      if (!values[field]) missingFields.push(fieldLabels[field]);
    });

    // Special case: Interface
    if (movementType === "Interface") {
      ["uniquePartImpact", "movementType"].forEach(field => {
        if (!values[field] && !missingFields.includes(fieldLabels[field])) {
          missingFields.push(fieldLabels[field]);
        }
      });
    }

    // Cost Center check
  

    
   

    // Submit or show alert
    if (missingFields.length === 0) {
      if (formstatus === 'Draft') {
        SaveData('Draft', 0);
      } else {
        const approvalNoteNumber = await generateNoteNumber();
        updateFieldBeforeSubmit(0, "Request Submitted", "Pending for Approval", "Submitted", approvalNoteNumber);
      }
    } else {
      setButtondisable(true);
      alert("Cannot submit the request. Missing mandatory fields:\n- " + missingFields.join("\n- "));
    }
  }

  async function SaveData(NewStatus, NewStage) {
    let fields =
    {
      "Status": NewStatus,
      "Stage": NewStage,
      "LastAction": new Date(),
      "MovementType": formikRef.current.values.movementType,
      "HeaderName": JSON.stringify(uploadedFileKey.current),
      "Details": Details.current,
      "WF": JSON.stringify(newworkflow.current),
      "CostCenter": formikRef.current.values.costCenter,
      "UniquePartImpact": formikRef.current.values.uniquePartImpact,
      "GrossValue": formikRef.current.values.grossValue,
      "NetValue": formikRef.current.values.netValue,
      "InitiatorEmpId": props.EmployeeId[0].EmployeeID,
      "Department": formikRef.current.values.reqDepartment,
      "MovementReason": formikRef.current.values.movementReason,
    };
    spCrudObj = await USESPCRUD();
    await spCrudObj.insertData('Prts_List', fields, props).then(async (brrInsertResult) => {
      setLoading(true);
      console.log('Item inserted:', brrInsertResult);
      console.log('Item inserted:', brrInsertResult);
      console.log(brrInsertResult);
      let rid = brrInsertResult.data.Id;
      let Title = `Prts/${format(new Date(), 'yyyy')}/${rid.toString().padStart(6, 0)}`;
      let Newfields = { 'Title': Title }
      await spCrudObj.updateData('Prts_List', brrInsertResult.data.Id, Newfields, props).then(async (result) => {
        const itemId = rid;
        if (itemId && attachments.length > 0) {
          for (const file of attachments) {
            try {
              await spCrudObj.addAttchmentInList(file, 'Prts_List', itemId, file.name, props);
              console.log(`Attachment ${file.name} uploaded.`);
            } catch (error) {
              console.error(`Failed to upload attachment ${file.name}:`, error);
              alert(`Failed to upload attachment ${file.name}:`);
              setLoading(false);
              return false;
            }
          }
          alert(`Request ` + Title + ` has been Saved as Draft`);
          setLoading(false);
        }
        else {
          alert(`Request ` + Title + ` has been Saved as Draft`);
          setLoading(false);
        }
      });
      history.push('/InitiatorLanding');
    });
  }

  function getCurrentFinancialYear() {
    var financial_year = "";
    var today = new Date();
    if ((today.getMonth() + 1) <= 3) {
      financial_year = (today.getFullYear() - 1).toString().substring(2, 4) + "-" + today.getFullYear().toString().substring(2, 4)
    } else {
      financial_year = today.getFullYear().toString().substring(2, 4) + "-" + (today.getFullYear() + 1).toString().substring(2, 4)
    }
    return financial_year;

  }

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

  const handleClose = () => {
    const lastActive = sessionStorage.getItem('sidebarFrom');
    if (lastActive) {
      history.push(lastActive);
    } else {
      history.push('/InitiatorLanding'); // Fallback route if none found
    }
  };

  async function generateNoteNumber() {
    let filterSRNO = ParameterDetails.filter((m) => m.Title === 'MovementWiseSerialNo');

    if (!filterSRNO.length || !filterSRNO[0].Details) {
      console.error("MovementWiseSerialNo details not found.");
      return;
    }

    let tmp = JSON.parse(filterSRNO[0].Details);
    let mm = formikRef?.current?.values?.movementType;

    let idx = tmp.findIndex(o => o.movement.indexOf(mm) >= 0);
    if (idx === -1) {
      console.error("Movement type not found in serial number config.");
      return;
    }

    let no = Number(tmp[idx].srno) + 1;
    let srno = Number(tmp[idx].default) + Number(tmp[idx].srno);
    let Approvalnoteno = `FY/${getCurrentFinancialYear()}/MG${srno.toString().padStart(6, '0')}`
    setApprovalNoteNo(Approvalnoteno);
    tmp[idx].srno = no;
    let Fields = JSON.stringify(tmp);
    console.log(Fields);
    try {
      const spCrudObj = await USESPCRUD();
      await spCrudObj.updateData("Parameters", filterSRNO[0].Id, Fields, props);
      console.log("Parameter Data has been updated successfully");

    } catch (error) {
      console.error("Error updating parameter data:", error);
    }
    return Approvalnoteno;
  }

  function handlesubmit(values: FormValues, formikHelpers: FormikHelpers<FormValues>): void | Promise<any> {
    throw new Error('Function not implemented.');
  }

  function set_InitContainer(): void {
    throw new Error('Function not implemented.');
  }
//--------------------------------------------------------Tab contain---------------------------------------------------------
  useEffect(() => {
    const hash = location.hash.replace("#", ""); 
    if (hash && hash !== "InitiatorLanding") {
      setActiveKey(hash);
    }
  }, [location]);

  // Change tab & update URL hash without reloading
  const handleSelect = (k: string | null) => {
    if (k) {
      setActiveKey(k);
      navigate.push(`#${k}`, { replace: true });
    }
  };
//-----------------------------------------------------------------------------------------------------------------------------
  return (
    <Formik initialValues={initialvalue} onSubmit={handlesubmit} innerRef={formikRef}>
      <Form>
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
          <div className='container'>
            <div className="row btn-Row" >
              <div className="col-sm-9" style={{ marginTop: '10px' }} id="btnBar">
                <a
                  style={{ cursor: 'pointer', display: 'inline' }}
                  className="btn btn-default"
                  id="btnClose"
                >
                  Back
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'inline' }}
                  className="btn btn-default"
                  id="btnCreateDraft"
                >
                  Create Draft
                </a>

                <a
                  style={{ cursor: 'pointer' }}
                  className="btn btn-default btnNonDraft"
                  id="btnSubmit"
                >
                  Submit
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft"
                  id="btnWithDrawn"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Withdrawn
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft"
                  id="btnPrint"
                >
                  Print
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnInitClose"
                  id="btnCloseIssue"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Close Issue
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnInitD"
                  id="btnReturnBackToPITMember"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Return Back to PIT Member
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnInitD"
                  id="btnForwardToNextPITMember"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Forward to next PIT Member
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverD"
                  id="btnProcessWithIssueCloseD"
                >
                  Process with Issue Close
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverD"
                  id="btnProcessWithIssueOpen"
                >
                  Process with Issue Open
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverD"
                  id="btnSendBackToPreviousStage"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Send Back to Previous Stage
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverD"
                  id="btnForwardAtD7"
                  data-bs-toggle="modal"
                  data-bs-target="#ForwardAtD7Container"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Assign To Champion
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverNT"
                  id="btnProcessWithIssueCloseNT"
                >
                  Process with Issue Close
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnApproverNT"
                  id="btnBackToIntiator"
                  data-bs-toggle="modal"
                  data-bs-target="#RemarksContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Send Back to Initiator (Issue Open)
                </a>

                <a
                  style={{ cursor: 'pointer', display: 'none' }}
                  className="btn btn-default btnNonDraft btnInitNT"
                  id="btnAssignIssueToAnoterUser"
                  data-bs-toggle="modal"
                  data-bs-target="#AssignToAnoterUserContainer"
                  data-backdrop="static"
                  data-keyboard="false"
                >
                  Re-Assign Issue
                </a>
              </div>

              <div className="col-sm-2 request" style={{ marginTop: '10px' }}>
                <span className="statusLbl" style={{ fontSize: '12px' }}>
                  Status: <span id="docStatus"></span>
                </span>
              </div>

              <div className="col-sm-1">
                <span
                  id="CH_Status"
                  className="badge badge-light"
                  style={{ fontSize: '20px', marginTop: '8px' }}
                ></span>
              </div>
            </div>
            {/* rEMARK sECTION */}
            <div className="row" style={{display:'none'}}>
              <div className="container">
                {/* Bootstrap Modal */}
                <div className="modal fade" id="RemarksContainer" role="dialog" aria-hidden="true">
                  <div className="modal-dialog">
                    {/* Modal content */}
                    <div className="modal-content">
                      <div className="modal-header">
                        <button type="button" className="close" data-bs-dismiss="modal" aria-label="Close">
                          ×
                        </button>
                        <h4 className="modal-title">Remarks</h4>
                      </div>

                      <div className="modal-body">
                        {/* Row 1 */}
                        <div className="row top-buffer">
                          <div className="col-sm-12">
                            <label htmlFor="mCommonRemarks">
                              <span className="required">*</span> Remarks -{' '}
                              <span id="RemarksTitle"></span>
                            </label>
                            <textarea
                              rows={3}
                              className="form-control restriceTabAndDoubleString"
                              id="mCommonRemarks"
                              maxLength={250}
                            ></textarea>
                          </div>
                        </div>
                      </div>

                      <div className="modal-footer">
                        <div>
                          <button
                            id="btnRemarks"
                            type="button"
                            className="btn btn-default"
                            onClick={() => {
                              // Add validation or logic here
                              console.log('Update button clicked');
                            }}
                          >
                            Update
                          </button>
                          <button
                            type="button"
                            className="btn btn-default"
                            data-bs-dismiss="modal"
                          >
                            Close
                          </button>
                        </div>
                      </div>
                    </div>
                    {/* End modal-content */}
                  </div>
                  {/* End modal-dialog */}
                </div>
                {/* End modal fade */}
              </div>
            </div>
            <div className="row">
              <div className="container">
                <div className="modal fade" id="AssignToAnoterUserContainer" role="dialog" aria-hidden="true">
                  <div className="modal-dialog">
                    <div className="modal-content">

                      {/* Modal Header */}
                      <div className="modal-header">
                        <button type="button" className="close" data-bs-dismiss="modal" aria-label="Close">
                          ×
                        </button>
                        <h4 className="modal-title">Issue Assign To Another User</h4>
                      </div>

                      {/* Modal Body */}
                      <div className="modal-body">
                        {/* Row 1 */}
                        <div className="row top-buffer">
                          <div className="col-sm-5">
                            <label htmlFor="newNonTechDept">
                              <span className="required">*</span> Agency
                            </label>
                            <select
                              id="newNonTechDept"
                              className="form-control"
                              value={agency}
                              onChange={(e) => setAgency(e.target.value)}
                            >
                              <option value="">-- Select Agency --</option>
                              {agencies.map((item, index) => (
                                <option key={index} value={item}>{item}</option>
                              ))}
                            </select>
                          </div>

                          <div className="col-sm-5">
                            <label htmlFor="manualCH_Status">
                              <span className="required">*</span> Change Status
                            </label>
                            <select
                              id="manualCH_Status"
                              className="form-control"
                              value={status}
                              onChange={(e) => setStatus(e.target.value)}
                            >
                              <option value="">-- Select Status --</option>
                              <option value="1/6">1/6</option>
                              <option value="2/6">2/6</option>
                              <option value="3/6">3/6</option>
                              <option value="4/6">4/6</option>
                            </select>
                          </div>
                        </div>

                        {/* Row 2 */}
                        <div className="row top-buffer">
                          <div className="col-sm-12">
                            <span className="glyphicon glyphicon-user"></span>
                            <label htmlFor="peoplePickerUser2">
                              <span className="required">*</span> Part quality issue / Other issue
                            </label>
                            <br />
                            <div id="peoplePickerUser2">
                              Loading people picker...
                            </div>
                          </div>
                        </div>

                        {/* Row 3 */}
                        <div className="row top-buffer">
                          <div className="col-sm-12">
                            <label htmlFor="mNTRemarks">
                              <span className="required">*</span> Remarks
                            </label>
                            <textarea
                              id="mNTRemarks"
                              rows={3}
                              maxLength={250}
                              className="form-control restriceTabAndDoubleString"
                              value={remarks}
                              onChange={(e) => setRemarks(e.target.value)}
                            />
                          </div>
                        </div>
                      </div>

                      {/* Modal Footer */}
                      <div className="modal-footer">
                        <button
                          id="btnRemarks"
                          type="button"
                          className="btn btn-default"
                          onClick={handleUpdate}
                        >
                          Update
                        </button>
                        <button type="button" className="btn btn-default" data-bs-dismiss="modal">
                          Close
                        </button>
                      </div>

                    </div>
                  </div>
                </div>
              </div>
            </div>
            <div className="row">
              <div className="container">
                <div className="modal fade" id="ForwardAtD7Container" role="dialog" aria-hidden="true">
                  <div className="modal-dialog">
                    <div className="modal-content">

                      {/* Modal Header */}
                      <div className="modal-header">
                        <button type="button" className="close" data-bs-dismiss="modal" aria-label="Close">
                          ×
                        </button>
                        <h4 className="modal-title">Assign Issue To Champion</h4>
                      </div>

                      {/* Modal Body */}
                      <div className="modal-body">
                        <div className="row top-buffer">
                          <div className="col-sm-12">
                            <span className="glyphicon glyphicon-user"></span>
                            <label htmlFor="peoplePickerUser3">
                              <span className="required">*</span> Select Champion Name
                            </label>
                            <br />
                            <div id="peoplePickerUser3">
                              Loading people picker...
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Modal Footer */}
                      <div className="modal-footer">
                        <div>
                          <button
                            id="btnRemarks"
                            type="button"
                            className="btn btn-default"
                            onClick={handleAssign}
                          >
                            Assign
                          </button>
                          <button
                            type="button"
                            className="btn btn-default"
                            data-bs-dismiss="modal"
                          >
                            Close
                          </button>
                        </div>
                      </div>

                    </div>
                  </div>
                </div>
              </div>
            </div>
            <div>
              {/* Workflow Bar */}
              <div className="row">
                <div className="workFlowBar">
                  <span id="GUI_WorkFlow"></span>
                </div>
              </div>

              {/* Problem Description and Req. Date */}
              <div className="row marginTop10">
                <div className="col-sm-1 caption">Problem Description</div>
                <div className="col-sm-8">
                  <input
                    type="text"
                    id="vTitle"
                    readOnly
                    style={{ width: '100%' }}
                  />
                </div>
                <div className="col-sm-1 caption">Req. Date</div>
                <div className="col-sm-2">
                  <input
                    type="text"
                    id="vReqDate"
                    readOnly
                    style={{ width: '100%' }}
                  />
                </div>
              </div>

              {/* Req. No */}
              <div className="row marginTop03">
                <div className="col-sm-9"></div>
                <div className="col-sm-1 caption">Req. No</div>
                <div className="col-sm-2">
                  <input
                    type="text"
                    id="vReqNo"
                    readOnly
                    style={{ width: '100%' }}
                  />
                </div>
              </div>
            </div>
            <div className="row marginTop10" id="FormRow">
              <ul className="nav nav-tabs" id="myTab" role="tablist">
                <li id="li-Basic" className="active">
                  <a data-toggle="tab" href="#Tab1" role="tab">BASE INFORMATION</a>
                </li>
                <li id="li-NT">
                  <a data-toggle="tab" href="#Tab2" role="tab">TECHNICAL ISSUE</a>
                </li>
                <li id="li-7D1">
                  <a data-toggle="tab" href="#Tab3" role="tab">D1</a>
                </li>
                <li id="li-7D2">
                  <a data-toggle="tab" href="#Tab4" role="tab">D2</a>
                </li>
                <li id="li-7D3">
                  <a data-toggle="tab" href="#Tab5" role="tab">D3</a>
                </li>
                <li id="li-7D4">
                  <a data-toggle="tab" href="#Tab6" role="tab">D4</a>
                </li>
                <li id="li-7D5">
                  <a data-toggle="tab" href="#Tab7" role="tab">D5</a>
                </li>
                <li id="li-7D6">
                  <a data-toggle="tab" href="#Tab8" role="tab">D6</a>
                </li>
                <li id="li-7D7">
                  <a data-toggle="tab" href="#Tab9" role="tab">D7</a>
                </li>
                <li>
                  <a data-toggle="tab" href="#Tab10" role="tab">SUMMARY</a>
                </li>
              </ul>
              <div className="tab-content">
                <div  id="Tab1" className="tab-pane fade in active">
                  <div className="row" style={{ textAlign: "right" }}>
                    <a
                      id="btnBasic"
                      className="btn btn-primary"
                      style={{ marginRight: "20px", marginTop: "3px" }}
                      data-target="#initContainer"
                      data-toggle="modal"
                      data-backdrop="static"
                      data-keyboard="false"
                      onClick={() => set_InitContainer()} // Replace with your actual function
                    >
                      Edit
                    </a>
                  </div>

                  <div className="container">
                    <div className="modal fade" id="initContainer" role="dialog">
                      <div className="modal-dialog modal-lg">
                        <div className="modal-content">
                          <div className="modal-header">
                            <button type="button" className="close" data-dismiss="modal">
                              &times;
                            </button>
                            <h4 className="modal-title">Issue Details</h4>
                          </div>

                          <div className="modal-body">
                            <div className="row top-buffer">
                              <div className="col-sm-12">
                                <label htmlFor="mTitle">
                                  <span className="required">*</span>Problem Description
                                </label>
                                <input
                                  className="form-control restriceTabAndDoubleString"
                                  type="text"
                                  id="mTitle"
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-3">
                                <label htmlFor="mPartName">
                                  <span className="required">*</span>Part Name
                                </label>
                                <input
                                  className="form-control restriceTabAndDoubleString"
                                  type="text"
                                  id="mPartName"
                                />
                              </div>
                              <div className="col-sm-3">
                                <label htmlFor="mPartNo">
                                  <span className="required">*</span>Part Number
                                </label>
                                <input
                                  className="form-control restriceTabAndDoubleString"
                                  type="text"
                                  id="mPartNo"
                                />
                              </div>
                              <div className="col-sm-3">
                                <label htmlFor="mPartSupplier">
                                  <span className="required">*</span>Supplier
                                </label>
                                <input
                                  className="form-control restriceTabAndDoubleString"
                                  type="text"
                                  id="mPartSupplier"
                                />
                              </div>
                            </div>

                            {/* Continue building rows here... */}
                          </div>

                          <div className="modal-footer">
                            <div>
                              <button
                                id="btnUpdate_initContainer"
                                type="button"
                                className="btn btn-default"
                              >
                                Update
                              </button>
                              <button
                                type="button"
                                className="btn btn-default"
                                data-dismiss="modal"
                              >
                                Close
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                  <table className="table table-bordered marginTop10" style={{ width: "100%", padding: "3px" }}>
                    <thead>
                      <tr>
                        <th>PRTS Source</th>
                        <th>Model</th>
                        <th>Initiator Name</th>
                        <th>Initiator Department</th>
                        <th>VIN Number</th>
                        <th>MFG Shop Selection</th>
                        <th>Issue Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td><input className="txtFullWidth" type="text" id="vPRTSSource" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vProjectCode" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vInitName" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vInitDept" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vIssueVINNo" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vMFGShop" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vIssueStatus" readOnly /></td>
                      </tr>
                      <tr>
                        <th>Build Type</th>
                        <th colSpan={2}>Part Name</th>
                        <th>Part No</th>
                        <th colSpan={3}>Supplier</th>
                      </tr>
                      <tr>
                        <td><input className="txtFullWidth" type="text" id="vBuildType" readOnly /></td>
                        <td colSpan={2}><input className="txtFullWidth" type="text" id="vPartName" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vPartNo" readOnly /></td>
                        <td colSpan={3}><input className="txtFullWidth" type="text" id="vPartSupplier" readOnly /></td>
                      </tr>
                      <tr>
                        <th colSpan={4} style={{ textAlign: "left" }}>Problem Definition</th>
                        <th colSpan={2} style={{ textAlign: "left" }}>
                          <span className="Basic_AfterSubmit">Attachment</span>
                          <div className="files Basic_BeforeSubmit" id="attachFilesHolder1">
                            <label htmlFor="GeneralAttachmentFile" style={{ marginLeft: "20px" }}>
                              [ Attachments <span className="glyphicon glyphicon-paperclip"></span> ]
                              <span className="glyphicon glyphicon-arrow-left"></span> Click
                            </label>
                            <input id="GeneralAttachmentFile" type="file" style={{ display: "none" }} />
                          </div>
                        </th>
                        <th>Engine Type</th>
                      </tr>
                      <tr>
                        <td colSpan={4}>
                          <textarea rows={3} id="vIssueDescription" readOnly />
                        </td>
                        <td colSpan={2}>
                          <div>
                            <img
                              src="../../Style Library/Custom/images/wait-circle.gif"
                              alt="waitMsg"
                              id="GeneralimgWaitFileUpload"
                              style={{ display: "none" }}
                            />
                          </div>
                          <div id="GeneralAttachmentFileList"></div>
                          <div id="GeneralUploadedFileList"></div>
                        </td>
                        <td><input className="txtFullWidth" type="text" id="vEngineType" readOnly /></td>
                      </tr>
                      <tr>
                        <th>Issue Category</th>
                        <th>Severity</th>
                        <th>Qty Affected</th>
                        <th>Variant Affected</th>
                        <th>Is Repeated</th>
                        <th>Issue Ref. No</th>
                        <th>Commodity</th>
                      </tr>
                      <tr>
                        <td><input className="txtFullWidth" type="text" id="vCategory" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vSeverity" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vQtyAffected" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vVariantAffected" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vIsRepeated" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vRefNo" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vCommodity" readOnly /></td>
                      </tr>
                      <tr id="attrRow">
                        <th>Purging Attachment</th>
                        <td style={{ textAlign: "left" }}>
                          <span id="vPurgingAttachment"></span>
                        </td>
                      </tr>
                      <tr>
                        <td colSpan={10}>
                          {/* You can include the attachment sub-table as a separate React component here for clarity */}
                          {/* <AttachmentTable /> */}
                        </td>
                      </tr>
                      <tr>
                        <th colSpan={5}>Analysis details</th>
                        <th>Is Root Cause Found</th>
                        <th>Require 7D Process</th>
                      </tr>
                      <tr>
                        <td colSpan={5}><textarea rows={3} id="vAnalysis" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vRootCauseFound" readOnly /></td>
                        <td><input className="txtFullWidth" type="text" id="vIs7D" readOnly /></td>
                      </tr>
                    </tbody>
                  </table>

                </div>
                <div className="tab-pane fade" id="Tab2">
                  <ul className="nav nav-tabs marginTop03" id="Sub_NT_Tab">
                    <li className="active">
                      <a data-toggle="tab" href="#NT_Tab1">
                        Active
                      </a>
                    </li>
                    <li>
                      <a data-toggle="tab" href="#NT_Tab2">
                        History
                      </a>
                    </li>
                  </ul>

                  <div className="tab-content">
                    <div className="tab-pane fade in active" id="NT_Tab1">
                      <div className="row" style={{ textAlign: "right" }}>
                        <a
                          id="btnNT"
                          className="btn btn-primary"
                          style={{ marginRight: "20px", marginTop: "3px" }}
                          data-target="#NonTechnicalContainer"
                          data-toggle="modal"
                          data-backdrop="static"
                          data-keyboard="false"
                        >
                          Edit
                        </a>
                      </div>

                      <div className="container">
                        <div className="modal fade" id="NonTechnicalContainer" role="dialog">
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button type="button" className="close" data-dismiss="modal">
                                  &times;
                                </button>
                                <h4 className="modal-title">Technical Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mNTAnalysis">
                                      <span className="required">*</span>
                                      Analysis Details
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="mNTAnalysis"
                                    ></textarea>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-3">
                                    <label htmlFor="mNTRootCauseFound">
                                      <span className="required">*</span>
                                      Is Root Cause Found?
                                    </label>
                                    <select className="form-control" id="mNTRootCauseFound">
                                      <option value="-1"></option>
                                      <option value="Yes">Yes</option>
                                      <option value="No">No</option>
                                    </select>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mNTICA_Details">
                                      <span className="required conditiona2Chk">*</span>
                                      ICA Action Implementation Details
                                    </label>
                                    <input
                                      type="text"
                                      id="mNTICA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mNTICA_VIN">
                                      <span className="required conditiona2Chk">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mNTICA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mNTPCA_Details">
                                      <span className="required conditiona2Chk">*</span>
                                      PCA Action
                                    </label>
                                    <input
                                      type="text"
                                      id="mNTPCA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mNTPCA_VIN">
                                      <span className="required conditiona2Chk">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mNTPCA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mNT_Remarks" id="lblRemarks">
                                      <span className="required conditiona3Chk">*</span>
                                      Remarks (if Root Cause Not Found)
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="mNT_Remarks"
                                    ></textarea>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="txtRootCause" id="lblRootCause">
                                      <span className="required conditionaRtcChk">*</span>
                                      Root Cause
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="txtRootCause"
                                    ></textarea>
                                  </div>
                                </div>
                              </div>
                              <div className="modal-footer">
                                <button
                                  id="btnUpdate_NonTechnicalContainer"
                                  type="button"
                                  className="btn btn-default"
                                >
                                  Update
                                </button>
                                <button
                                  type="button"
                                  className="btn btn-default"
                                  data-dismiss="modal"
                                >
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div id="NT_Tab2" className="tab-pane fade">
                      <div className="marginTop10">
                        <table id="NT_HistoryTable">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Req_Date</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>


                <div id="Tab3" className="tab-pane fade">
                  <ul className="nav nav-tabs marginTop03" id="Sub_D1_Tab">
                    <li className="active">
                      <a data-toggle="tab" href="#D1_Tab1">
                        Active
                      </a>
                    </li>
                    <li>
                      <a data-toggle="tab" href="#D1_Tab2">
                        History
                      </a>
                    </li>
                  </ul>

                  <div className="tab-content">
                    <div id="D1_Tab1" className="tab-pane fade in active">
                      <div className="row" style={{ textAlign: "right" }}>
                        <a
                          id="btnD1"
                          className="btn btn-primary"
                          style={{ marginRight: "20px", marginTop: "3px" }}
                          data-target="#D1Container"
                          data-toggle="modal"
                          data-backdrop="static"
                          data-keyboard="false"
                        >
                          Edit
                        </a>
                      </div>

                      <div className="container">
                        <div className="modal fade" id="D1Container" role="dialog">
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button type="button" className="close" data-dismiss="modal">
                                  &times;
                                </button>
                                <h4 className="modal-title">Diamond 1 - Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="D1AnalysisAttachmentFile">
                                      <span className="required">*</span>
                                      Analysis Details [ Attachments <span className="glyphicon glyphicon-paperclip"></span> ]
                                      <span className="glyphicon glyphicon-arrow-left"></span> Click
                                    </label>
                                    <input id="D1AnalysisAttachmentFile" type="file" style={{ display: "none" }} />
                                    <div id="D1AnalysisAttachmentFileList"></div>
                                    <div id="D1AnalysisUploadedFileList"></div>
                                    <div>
                                      <img
                                        src="../../Style Library/Custom/images/wait-circle.gif"
                                        alt="waitMsg"
                                        id="D1AnalysisimgWaitFileUpload"
                                        style={{ display: "none" }}
                                      />
                                    </div>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-3">
                                    <label htmlFor="mD1RootCauseFound">
                                      <span className="required">*</span>
                                      Is Root Cause Found?
                                    </label>
                                    <select className="form-control" id="mD1RootCauseFound">
                                      <option value="-1"></option>
                                      <option value="Yes">Yes</option>
                                      <option value="No">No</option>
                                    </select>
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD1ActionStatus">
                                      <span className="required chkActionStatus">*</span>
                                      Action Status
                                    </label>
                                    <select className="form-control" id="mD1ActionStatus">
                                      <option value="-1"></option>
                                      <option value="Will_Implemented">Action Will Be Implemented</option>
                                      <option value="Implemented">Action Implemented</option>
                                    </select>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD1ICA_Details">
                                      <span className="required conditiona2Chk">*</span>
                                      ICA Action Implementation Details
                                    </label>
                                    <input type="text" id="mD1ICA_Details" className="form-control restriceTabAndDoubleString" maxLength={250} />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD1ICA_VIN">
                                      <span className="required conditiona2Chk">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input type="text" id="mD1ICA_VIN" className="form-control restriceTabAndDoubleString" maxLength={250} />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD1PCA_Details">
                                      <span className="required chkPCA">*</span>
                                      PCA Action
                                    </label>
                                    <input type="text" id="mD1PCA_Details" className="form-control restriceTabAndDoubleString" maxLength={250} />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD1PCA_VIN">
                                      <span className="required chkPCA">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input type="text" id="mD1PCA_VIN" className="form-control restriceTabAndDoubleString" maxLength={250} />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD1_Remarks">
                                      <span className="required conditiona3Chk">*</span>
                                      Remarks (if Root Cause Not Found)
                                    </label>
                                    <textarea rows={3} className="form-control restriceTabAndDoubleString" id="mD1_Remarks"></textarea>
                                  </div>
                                </div>
                              </div>

                              <div className="modal-footer">
                                <button id="btnUpdate_D1Container" type="button" className="btn btn-default">
                                  Update
                                </button>
                                <button type="button" className="btn btn-default" data-dismiss="modal">
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Summary Table */}
                      <table className="table table-bordered marginTop10" style={{ width: "100%", padding: "3px" }}>
                        <thead>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD1AssignTo" readOnly /></td>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD1AssignDT" readOnly /></td>
                            <td colSpan={3}><input className="txtFullWidth" type="text" id="vD1ActionStatus" readOnly /></td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div id="D1_Tab2" className="tab-pane fade">
                      <div className="marginTop10">
                        <table id="D1_HistoryTable">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="tab-pane fade" id="Tab4">
                  <ul className="nav nav-tabs marginTop03" id="Sub_D2_Tab">
                    <li className="active">
                      <a data-toggle="tab" href="#D2_Tab1">Active</a>
                    </li>
                    <li>
                      <a data-toggle="tab" href="#D2_Tab2">History</a>
                    </li>
                  </ul>

                  <div className="tab-content">
                    <div id="D2_Tab1" className="tab-pane fade in active">
                      <div className="row" style={{ textAlign: "right" }}>
                        <a
                          id="btnD2"
                          className="btn btn-primary"
                          style={{ marginRight: "20px", marginTop: "3px" }}
                          data-target="#D2Container"
                          data-toggle="modal"
                          data-backdrop="static"
                          data-keyboard="false"
                        >
                          Edit
                        </a>
                      </div>

                      <div className="container">
                        <div className="modal fade" id="D2Container" role="dialog">
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button type="button" className="close" data-dismiss="modal">
                                  &times;
                                </button>
                                <h4 className="modal-title">Diamond 2 - Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <div className="files" id="attachFilesHolder3">
                                      <label htmlFor="D2AnalysisAttachmentFile" style={{ marginLeft: 0 }}>
                                        <span className="required">*</span>
                                        Analysis Details [ Attachments <span className="glyphicon glyphicon-paperclip"></span>]
                                        <span className="glyphicon glyphicon-arrow-left"></span>
                                        Click
                                      </label>
                                      <input id="D2AnalysisAttachmentFile" type="file" style={{ display: "none" }} />
                                    </div>

                                    <div id="D2AnalysisAttachmentFileList"></div>
                                    <div id="D2AnalysisUploadedFileList"></div>

                                    <div>
                                      <img
                                        src="../../Style Library/Custom/images/wait-circle.gif"
                                        alt="waitMsg"
                                        id="D2AnalysisimgWaitFileUpload"
                                        style={{ display: "none" }}
                                      />
                                    </div>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-3">
                                    <label htmlFor="mD2RootCauseFound">
                                      <span className="required">*</span>
                                      Is Root Cause Found?
                                    </label>
                                    <select className="form-control" id="mD2RootCauseFound">
                                      <option value="-1"></option>
                                      <option value="Yes">Yes</option>
                                      <option value="No">No</option>
                                    </select>
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD2ActionStatus">
                                      <span className="required chkActionStatus">*</span>
                                      Action Status
                                    </label>
                                    <select className="form-control" id="mD2ActionStatus">
                                      <option value="-1"></option>
                                      <option value="Will_Implemented">Action Will Be Implemented</option>
                                      <option value="Implemented">Action Implemented</option>
                                    </select>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD2ICA_Details">
                                      <span className="required conditiona2Chk">*</span>
                                      ICA Action Implementation Details
                                    </label>
                                    <input
                                      type="text"
                                      id="mD2ICA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD2ICA_VIN">
                                      <span className="required conditiona2Chk">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mD2ICA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD2PCA_Details">
                                      <span className="required chkPCA">*</span>
                                      PCA Action
                                    </label>
                                    <input
                                      type="text"
                                      id="mD2PCA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD2PCA_VIN">
                                      <span className="required chkPCA">*</span>
                                      VIN Cut Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mD2PCA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD2_Remarks">
                                      <span className="required conditiona3Chk">*</span>
                                      Remarks (if Root Cause Not Found)
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="mD2_Remarks"
                                    ></textarea>
                                  </div>
                                </div>
                              </div>
                              <div className="modal-footer">
                                <button
                                  id="btnUpdate_D2Container"
                                  type="button"
                                  className="btn btn-default"
                                >
                                  Update
                                </button>
                                <button type="button" className="btn btn-default" data-dismiss="modal">
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Table display */}
                      <table className="table table-bordered marginTop10" style={{ width: "100%", padding: "3px" }}>
                        <tbody>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                          <tr>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD2AssignTo" readOnly /></td>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD2AssignDT" readOnly /></td>
                            <td colSpan={3}><input className="txtFullWidth" type="text" id="vD2ActionStatus" readOnly /></td>
                          </tr>
                          <tr>
                            <th colSpan={6}>Analysis details [<span id="D2Template"></span>]</th>
                            <th>Is Root Cause Found</th>
                          </tr>
                          <tr>
                            <td colSpan={6}><div id="vD2Analysis"></div></td>
                            <td><input className="txtFullWidth" type="text" id="vD2RootCauseFound" readOnly /></td>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <th colSpan={5}>ICA action Implementation details</th>
                            <th colSpan={2}>VIN Cut off </th>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <td colSpan={5}><textarea rows={3} id="vD2ICA_Details" readOnly></textarea></td>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD2ICA_VIN" readOnly /></td>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <th colSpan={5}>PCA action plan</th>
                            <th colSpan={2}>VIN Cut off </th>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <td colSpan={5}><textarea rows={3} id="vD2PCA_Details" readOnly></textarea></td>
                            <td colSpan={2}><input className="txtFullWidth" type="text" id="vD2PCA_VIN" readOnly /></td>
                          </tr>
                          <tr>
                            <th colSpan={4}>Remarks</th>
                            <th colSpan={3}>
                              <span className="D2_AfterSubmit">Attachment</span>
                              <div className="files D2_BeforeSubmit" id="attachFilesHolder4">
                                <label htmlFor="D2AttachmentFile" style={{ marginLeft: "20px" }}>
                                  [ Attachments <span className="glyphicon glyphicon-paperclip"></span> ]
                                  <span className="glyphicon glyphicon-arrow-left"></span>
                                  Click
                                </label>
                                <input id="D2AttachmentFile" type="file" style={{ display: "none" }} />
                              </div>
                            </th>
                          </tr>
                          <tr>
                            <td colSpan={4}><textarea rows={3} id="vD2_Remarks" readOnly></textarea></td>
                            <td colSpan={3}>
                              <div>
                                <img
                                  src="../../Style Library/Custom/images/wait-circle.gif"
                                  alt="waitMsg"
                                  id="D2imgWaitFileUpload"
                                  style={{ display: "none" }}
                                />
                              </div>
                              <div id="D2AttachmentFileList"></div>
                              <div id="D2UploadedFileList"></div>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div id="D2_Tab2" className="tab-pane fade">
                      <div className="marginTop10">
                        <table id="D2_HistoryTable">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
                {/*  end of tab 4 */}
                <div id="Tab5" className="tab-pane fade">
                  {/* Tabs */}
                  <ul className="nav nav-tabs marginTop03" id="Sub_D3_Tab">
                    <li className={activeTab === "D3_Tab1" ? "active" : ""}>
                      <a
                        href="#D3_Tab1"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D3_Tab1");
                        }}
                      >
                        Active
                      </a>
                    </li>
                    <li className={activeTab === "D3_Tab2" ? "active" : ""}>
                      <a
                        href="#D3_Tab2"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D3_Tab2");
                        }}
                      >
                        History
                      </a>
                    </li>
                  </ul>

                  <div className="tab-content">
                    {/* ---------------- D3_Tab1 ---------------- */}
                    <div
                      id="D3_Tab1"
                      className={`tab-pane fade ${activeTab === "D3_Tab1" ? "in active" : ""}`}
                    >
                      <div className="row" style={{ textAlign: "right" }}>
                        <button
                          id="btnD3"
                          className="btn btn-primary"
                          style={{ marginRight: "20px", marginTop: "3px" }}
                          onClick={() => setShowModal(true)}
                        >
                          Edit
                        </button>
                      </div>

                      {/* Modal */}
                      {showModal && (
                        <div className="modal fade in" style={{ display: "block" }}>
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button
                                  type="button"
                                  className="close"
                                  onClick={() => setShowModal(false)}
                                >
                                  ×
                                </button>
                                <h4 className="modal-title">Diamond 3 - Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                {/* Example: Remarks field */}
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD3_Remarks">
                                      <span className="required">*</span> Remarks
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control"
                                      id="mD3_Remarks"
                                      value={remarks}
                                      onChange={(e) => setRemarks(e.target.value)}
                                    />
                                  </div>
                                </div>
                              </div>
                              <div className="modal-footer">
                                <button
                                  id="btnUpdate_D3Container"
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => {
                                    console.log("Save remarks:", remarks);
                                    setShowModal(false);
                                  }}
                                >
                                  Update
                                </button>
                                <button
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => setShowModal(false)}
                                >
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Example Table */}
                      <table className="table table-bordered marginTop10" style={{ width: "100%", padding: "3px" }}>
                        <thead>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td colSpan={2}><input className="txtFullWidth" type="text" readOnly value="John Doe" /></td>
                            <td colSpan={2}><input className="txtFullWidth" type="text" readOnly value="2025-09-22" /></td>
                            <td colSpan={3}><input className="txtFullWidth" type="text" readOnly value="Open" /></td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* ---------------- D3_Tab2 ---------------- */}
                    <div
                      id="D3_Tab2"
                      className={`tab-pane fade ${activeTab === "D3_Tab2" ? "in active" : ""}`}
                    >
                      <div className="marginTop10">
                        <table id="D3_HistoryTable" className="table table-bordered">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td>1</td>
                              <td>ABC Agency</td>
                              <td>Jane Smith</td>
                              <td>Analysis Sample</td>
                              <td>Yes</td>
                              <td>Implemented</td>
                              <td>12345</td>
                              <td>Planned</td>
                              <td>67890</td>
                              <td>Checked</td>
                              <td>File.pdf</td>
                              <td>2025-09-21 14:30</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
                {/* end of 5 tab */}
                <div id="Tab6" className="tab-pane fade">
                  {/* Tabs */}
                  <ul className="nav nav-tabs marginTop03" id="Sub_D4_Tab">
                    <li className={activeTab === "D4_Tab1" ? "active" : ""}>
                      <a
                        href="#D4_Tab1"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D4_Tab1");
                        }}
                      >
                        Active
                      </a>
                    </li>
                    <li className={activeTab === "D4_Tab2" ? "active" : ""}>
                      <a
                        href="#D4_Tab2"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D4_Tab2");
                        }}
                      >
                        History
                      </a>
                    </li>
                  </ul>

                  <div className="tab-content">
                    {/* ---------------- D4_Tab1 ---------------- */}
                    <div
                      id="D4_Tab1"
                      className={`tab-pane fade ${activeTab === "D4_Tab1" ? "in active" : ""}`}
                    >
                      <div className="row" style={{ textAlign: "right" }}>
                        <button
                          id="btnD4"
                          className="btn btn-primary"
                          style={{ marginRight: "20px", marginTop: "3px" }}
                          onClick={() => setShowModal(true)}
                        >
                          Edit
                        </button>
                      </div>

                      {/* Modal */}
                      {showModal && (
                        <div className="modal fade in" style={{ display: "block" }}>
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button
                                  type="button"
                                  className="close"
                                  onClick={() => setShowModal(false)}
                                >
                                  ×
                                </button>
                                <h4 className="modal-title">Diamond 4 - Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                {/* Example: Remarks field */}
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD4_Remarks">
                                      <span className="required">*</span> Remarks
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control"
                                      id="mD4_Remarks"
                                      value={remarks}
                                      onChange={(e) => setRemarks(e.target.value)}
                                    />
                                  </div>
                                </div>
                              </div>
                              <div className="modal-footer">
                                <button
                                  id="btnUpdate_D4Container"
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => {
                                    console.log("Save remarks:", remarks);
                                    setShowModal(false);
                                  }}
                                >
                                  Update
                                </button>
                                <button
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => setShowModal(false)}
                                >
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Example Table */}
                      <table className="table table-bordered marginTop10" style={{ width: "100%", padding: "3px" }}>
                        <thead>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly value="John Doe" />
                            </td>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly value="2025-09-22" />
                            </td>
                            <td colSpan={3}>
                              <input className="txtFullWidth" type="text" readOnly value="Open" />
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* ---------------- D4_Tab2 ---------------- */}
                    <div
                      id="D4_Tab2"
                      className={`tab-pane fade ${activeTab === "D4_Tab2" ? "in active" : ""}`}
                    >
                      <div className="marginTop10">
                        <table id="D4_HistoryTable" className="table table-bordered">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td>1</td>
                              <td>XYZ Agency</td>
                              <td>Michael Brown</td>
                              <td>Sample analysis</td>
                              <td>No</td>
                              <td>Planned</td>
                              <td>11111</td>
                              <td>Pending</td>
                              <td>22222</td>
                              <td>In progress</td>
                              <td>doc.pdf</td>
                              <td>2025-09-20 09:15</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
                {/* endof tab6 */}
                <div id="Tab7" className="tab-pane fade">
                  {/* Tabs */}
                  <ul className="nav nav-tabs marginTop03" id="Sub_D5_Tab">
                    <li className={activeTab === "D5_Tab1" ? "active" : ""}>
                      <a
                        href="#D5_Tab1"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D5_Tab1");
                        }}
                      >
                        Active
                      </a>
                    </li>
                    <li className={activeTab === "D5_Tab2" ? "active" : ""}>
                      <a
                        href="#D5_Tab2"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D5_Tab2");
                        }}
                      >
                        History
                      </a>
                    </li>
                  </ul>

                  {/* Tab Content */}
                  <div className="tab-content">
                    {/* Tab 1 */}
                    <div
                      id="D5_Tab1"
                      className={`tab-pane fade ${activeTab === "D5_Tab1" ? "in active" : ""}`}
                    >
                      <div className="row" style={{ textAlign: "right" }}>
                        <a
                          id="btnD5"
                          className="btn btn-primary"
                          style={{ marginRight: 20, marginTop: 3 }}
                          data-target="#D5Container"
                          data-toggle="modal"
                          data-backdrop="static"
                          data-keyboard="false"
                        >
                          Edit
                        </a>
                      </div>

                      {/* Modal */}
                      <div className="container">
                        <div className="modal fade" id="D5Container" role="dialog">
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button type="button" className="close" data-dismiss="modal">
                                  ×
                                </button>
                                <h4 className="modal-title">Diamond 5 - Issue Details</h4>
                              </div>
                              <div className="modal-body">
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD5Analysis">
                                      <span className="required">*</span> Analysis Details
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="mD5Analysis"
                                    ></textarea>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-3">
                                    <label htmlFor="mD5RootCauseFound">
                                      <span className="required">*</span> Is Root Cause Found?
                                    </label>
                                    <select className="form-control" id="mD5RootCauseFound">
                                      <option value="-1"></option>
                                      <option value="Yes">Yes</option>
                                      <option value="No">No</option>
                                    </select>
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD5ActionStatus">
                                      <span className="required chkActionStatus">*</span>{" "}
                                      Action Status
                                    </label>
                                    <select className="form-control" id="mD5ActionStatus">
                                      <option value="-1"></option>
                                      <option value="Will_Implemented">
                                        Action Will Be Implemented
                                      </option>
                                      <option value="Implemented">Action Implemented</option>
                                    </select>
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD5ICA_Details">
                                      <span className="required conditiona2Chk">*</span> ICA
                                      Action Implementation Details
                                    </label>
                                    <input
                                      type="text"
                                      id="mD5ICA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD5ICA_VIN">
                                      <span className="required conditiona2Chk">*</span> VIN Cut
                                      Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mD5ICA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-8">
                                    <label htmlFor="mD5PCA_Details">
                                      <span className="required chkPCA">*</span> PCA Action
                                    </label>
                                    <input
                                      type="text"
                                      id="mD5PCA_Details"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                  <div className="col-sm-4">
                                    <label htmlFor="mD5PCA_VIN">
                                      <span className="required chkPCA">*</span> VIN Cut Off
                                    </label>
                                    <input
                                      type="text"
                                      id="mD5PCA_VIN"
                                      className="form-control restriceTabAndDoubleString"
                                      maxLength={250}
                                    />
                                  </div>
                                </div>

                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD5_Remarks">
                                      <span className="required conditiona3Chk">*</span> Remarks
                                      (if Root Cause Not Found)
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control restriceTabAndDoubleString"
                                      id="mD5_Remarks"
                                    ></textarea>
                                  </div>
                                </div>
                              </div>

                              <div className="modal-footer">
                                <div>
                                  <button
                                    id="btnUpdate_D5Container"
                                    type="button"
                                    className="btn btn-default"
                                  >
                                    Update
                                  </button>
                                  <button
                                    type="button"
                                    className="btn btn-default"
                                    data-dismiss="modal"
                                  >
                                    Close
                                  </button>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Table */}
                      <table
                        className="table table-bordered marginTop10"
                        style={{ width: "100%", padding: "3px" }}
                      >
                        <tbody>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                          <tr>
                            <td colSpan={2}>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5AssignTo"
                                readOnly
                              />
                            </td>
                            <td colSpan={2}>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5AssignDT"
                                readOnly
                              />
                            </td>
                            <td colSpan={3}>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5ActionStatus"
                                readOnly
                              />
                            </td>
                          </tr>
                          <tr>
                            <th colSpan={6}>Analysis details</th>
                            <th>Is Root Cause Found</th>
                          </tr>
                          <tr>
                            <td colSpan={6}>
                              <textarea rows={3} id="vD5Analysis" readOnly></textarea>
                            </td>
                            <td>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5RootCauseFound"
                                readOnly
                              />
                            </td>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <th colSpan={5}>ICA action Implementation details</th>
                            <th colSpan={2}>VIN Cut off</th>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <td colSpan={5}>
                              <textarea rows={3} id="vD5ICA_Details" readOnly></textarea>
                            </td>
                            <td colSpan={2}>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5ICA_VIN"
                                readOnly
                              />
                            </td>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <th colSpan={5}>PCA action plan</th>
                            <th colSpan={2}>VIN Cut off</th>
                          </tr>
                          <tr className="Hide_If_NT_RootCauseNotFound">
                            <td colSpan={5}>
                              <textarea rows={3} id="vD5PCA_Details" readOnly></textarea>
                            </td>
                            <td colSpan={2}>
                              <input
                                className="txtFullWidth"
                                type="text"
                                id="vD5PCA_VIN"
                                readOnly
                              />
                            </td>
                          </tr>

                          <tr>
                            <th colSpan={4}>Remarks</th>
                            <th colSpan={3}>
                              <span className="D5_AfterSubmit">Attachment</span>
                              <div className="files D5_BeforeSubmit" id="attachFilesHolder4">
                                <label htmlFor="D5AttachmentFile" style={{ marginLeft: 20 }}>
                                  [ Attachments{" "}
                                  <span className="glyphicon glyphicon-paperclip"></span> ]
                                  <span className="glyphicon glyphicon-arrow-left"></span>{" "}
                                  Click
                                </label>
                                <input
                                  id="D5AttachmentFile"
                                  type="file"
                                  style={{ display: "none" }}
                                />
                              </div>
                            </th>
                          </tr>
                          <tr>
                            <td colSpan={4}>
                              <textarea rows={3} id="vD5_Remarks" readOnly></textarea>
                            </td>
                            <td colSpan={3}>
                              <div>
                                <img
                                  src="../../Style Library/Custom/images/wait-circle.gif"
                                  alt="waitMsg"
                                  id="D5imgWaitFileUpload"
                                  style={{ display: "none" }}
                                />
                              </div>
                              <div id="D5AttachmentFileList"></div>
                              <div id="D5UploadedFileList"></div>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    {/* Tab 2 */}
                    <div
                      id="D5_Tab2"
                      className={`tab-pane fade ${activeTab === "D5_Tab2" ? "in active" : ""}`}
                    >
                      <div className="marginTop10">
                        <table id="D5_HistoryTable">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>

                {/* end of tab 7 */}

                <div id="Tab8" className="tab-pane fade">
                  {/* Sub Tabs */}
                  <Tabs
                    id="Sub_D6_Tab"
                    activeKey={activeTab}
                    onSelect={(k) => setActiveTab(k)}
                    className="marginTop03"
                  >
                    <Tab eventKey="active" title="Active">
                      <div className="row" style={{ textAlign: "right" }}>
                        <Button
                          id="btnD6"
                          className="btn btn-primary"
                          style={{ marginRight: 20, marginTop: 3 }}
                          onClick={() => setShowModal(true)}
                        >
                          Edit
                        </Button>
                      </div>

                      {/* Modal */}
                      <Modal
                        show={showModal}
                        onHide={() => setShowModal(false)}
                        size="lg"
                        backdrop="static"
                        keyboard={false}
                      >
                        <Modal.Header closeButton>
                          <Modal.Title>Diamond 6 - Issue Details</Modal.Title>
                        </Modal.Header>
                        <Modal.Body>
                          <div className="row top-buffer">
                            <div className="col-sm-12">
                              <label>
                                <span className="required">*</span> Analysis Details
                              </label>
                              <textarea rows={3} className="form-control"></textarea>
                            </div>
                          </div>

                          <div className="row top-buffer">
                            <div className="col-sm-3">
                              <label>
                                <span className="required">*</span> Is Root Cause Found?
                              </label>
                              <select className="form-control">
                                <option value="-1"></option>
                                <option value="Yes">Yes</option>
                                <option value="No">No</option>
                              </select>
                            </div>
                            <div className="col-sm-4">
                              <label>
                                <span className="required chkActionStatus">*</span> Action
                                Status
                              </label>
                              <select className="form-control">
                                <option value="-1"></option>
                                <option value="Will_Implemented">
                                  Action Will Be Implemented
                                </option>
                                <option value="Implemented">Action Implemented</option>
                              </select>
                            </div>
                          </div>

                          <div className="row top-buffer">
                            <div className="col-sm-8">
                              <label>
                                <span className="required conditiona2Chk">*</span> ICA Action
                                Implementation Details
                              </label>
                              <input type="text" className="form-control" maxLength={250} />
                            </div>
                            <div className="col-sm-4">
                              <label>
                                <span className="required conditiona2Chk">*</span> VIN Cut
                                Off
                              </label>
                              <input type="text" className="form-control" maxLength={250} />
                            </div>
                          </div>

                          <div className="row top-buffer">
                            <div className="col-sm-8">
                              <label>
                                <span className="required chkPCA">*</span> PCA Action
                              </label>
                              <input type="text" className="form-control" maxLength={250} />
                            </div>
                            <div className="col-sm-4">
                              <label>
                                <span className="required chkPCA">*</span> VIN Cut Off
                              </label>
                              <input type="text" className="form-control" maxLength={250} />
                            </div>
                          </div>

                          <div className="row top-buffer">
                            <div className="col-sm-12">
                              <label>
                                <span className="required conditiona3Chk">*</span> Remarks
                                (if Root Cause Not Found)
                              </label>
                              <textarea rows={3} className="form-control"></textarea>
                            </div>
                          </div>
                        </Modal.Body>
                        <Modal.Footer>
                          <Button
                            id="btnUpdate_D6Container"
                            className="btn btn-default"
                            onClick={() => setShowModal(false)}
                          >
                            Update
                          </Button>
                          <Button
                            className="btn btn-default"
                            onClick={() => setShowModal(false)}
                          >
                            Close
                          </Button>
                        </Modal.Footer>
                      </Modal>

                      {/* Table */}
                      <table
                        className="table table-bordered marginTop10"
                        style={{ width: "100%", padding: "3px" }}
                      >
                        <thead>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                            <td colSpan={3}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                          </tr>
                          <tr>
                            <th colSpan={6}>Analysis details</th>
                            <th>Is Root Cause Found</th>
                          </tr>
                          <tr>
                            <td colSpan={6}>
                              <textarea rows={3} readOnly></textarea>
                            </td>
                            <td>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                          </tr>
                          <tr>
                            <th colSpan={5}>ICA action Implementation details</th>
                            <th colSpan={2}>VIN Cut off</th>
                          </tr>
                          <tr>
                            <td colSpan={5}>
                              <textarea rows={3} readOnly></textarea>
                            </td>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                          </tr>
                          <tr>
                            <th colSpan={5}>PCA action plan</th>
                            <th colSpan={2}>VIN Cut off</th>
                          </tr>
                          <tr>
                            <td colSpan={5}>
                              <textarea rows={3} readOnly></textarea>
                            </td>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                          </tr>
                          <tr>
                            <th colSpan={3}>Remarks</th>
                            <th colSpan={3}>
                              <span className="D6_AfterSubmit">Attachment</span>
                              <div className="files D6_BeforeSubmit">
                                <label style={{ marginLeft: 20 }}>
                                  [ Attachments 📎 ] ⬅ Click
                                </label>
                                <input type="file" style={{ display: "none" }} />
                              </div>
                            </th>
                          </tr>
                          <tr>
                            <td colSpan={4}>
                              <textarea rows={3} readOnly></textarea>
                            </td>
                            <td colSpan={3}>
                              <div>
                                <img
                                  src="../../Style Library/Custom/images/wait-circle.gif"
                                  alt="waitMsg"
                                  style={{ display: "none" }}
                                />
                              </div>
                              <div id="D6AttachmentFileList"></div>
                              <div id="D6UploadedFileList"></div>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </Tab>

                    <Tab eventKey="history" title="History">
                      <div className="marginTop10">
                        <table id="D6_HistoryTable" className="table">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </Tab>
                  </Tabs>
                </div>
                <div id="Tab9" className="tab-pane fade">
                  {/* Tabs */}
                  <ul className="nav nav-tabs marginTop03" id="Sub_D7_Tab">
                    <li className={activeTab === "D7_Tab1" ? "active" : ""}>
                      <a
                        href="#D7_Tab1"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D7_Tab1");
                        }}
                      >
                        Active
                      </a>
                    </li>
                    <li className={activeTab === "D7_Tab2" ? "active" : ""}>
                      <a
                        href="#D7_Tab2"
                        onClick={(e) => {
                          e.preventDefault();
                          setActiveTab("D7_Tab2");
                        }}
                      >
                        History
                      </a>
                    </li>
                  </ul>

                  {/* Tab content */}
                  <div className="tab-content">
                    {/* --- TAB 1 --- */}
                    <div
                      id="D7_Tab1"
                      className={`tab-pane fade ${activeTab === "D7_Tab1" ? "in active" : ""}`}
                    >
                      <div className="row" style={{ textAlign: "right" }}>
                        <button
                          id="btnD7"
                          className="btn btn-primary"
                          style={{ marginRight: 20, marginTop: 3 }}
                          onClick={() => setShowModal(true)}
                        >
                          Edit
                        </button>
                      </div>

                      {/* Modal */}
                      {showModal && (
                        <div className="modal fade in" style={{ display: "block" }}>
                          <div className="modal-dialog modal-lg">
                            <div className="modal-content">
                              <div className="modal-header">
                                <button
                                  type="button"
                                  className="close"
                                  onClick={() => setShowModal(false)}
                                >
                                  ×
                                </button>
                                <h4 className="modal-title">Diamond 7 - Issue Details</h4>
                              </div>

                              <div className="modal-body">
                                {/* Example row */}
                                <div className="row top-buffer">
                                  <div className="col-sm-12">
                                    <label htmlFor="mD7Analysis">
                                      <span className="required">*</span> Analysis Details
                                    </label>
                                    <textarea
                                      rows={3}
                                      className="form-control"
                                      id="mD7Analysis"
                                    ></textarea>
                                  </div>
                                </div>

                                {/* Add other form rows here same as your HTML */}
                              </div>

                              <div className="modal-footer">
                                <button
                                  id="btnUpdate_D7Container"
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => {
                                    // handle save logic here
                                    setShowModal(false);
                                  }}
                                >
                                  Update
                                </button>
                                <button
                                  type="button"
                                  className="btn btn-default"
                                  onClick={() => setShowModal(false)}
                                >
                                  Close
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Table (kept same as HTML, only class -> className) */}
                      <table
                        className="table table-bordered marginTop10"
                        style={{ width: "100%", padding: "3px" }}
                      >
                        <thead>
                          <tr>
                            <th colSpan={2}>Issue Assign To</th>
                            <th colSpan={2}>Issue Assign Date</th>
                            <th colSpan={3}>Action Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                            <td colSpan={2}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                            <td colSpan={3}>
                              <input className="txtFullWidth" type="text" readOnly />
                            </td>
                          </tr>
                          {/* Continue converting rest of your table the same way */}
                        </tbody>
                      </table>
                    </div>

                    {/* --- TAB 2 --- */}
                    <div
                      id="D7_Tab2"
                      className={`tab-pane fade ${activeTab === "D7_Tab2" ? "in active" : ""}`}
                    >
                      <div className="marginTop10">
                        <table id="D7_HistoryTable" className="table">
                          <thead>
                            <tr>
                              <th></th>
                              <th>Agency</th>
                              <th>UserName</th>
                              <th>Analysis Details</th>
                              <th>Found Root Cause</th>
                              <th>ICA Action</th>
                              <th>ICA VIN Cut OFF</th>
                              <th>PCA Action</th>
                              <th>PCA VIN Cut OFF</th>
                              <th>Remarks</th>
                              <th>Attachment</th>
                              <th>Action DateTime</th>
                            </tr>
                          </thead>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
                <div id="Tab10" className="tab-pane fade">
                  <div style={{ textAlign: "left", marginTop: "10px" }}>
                    <table
                      id="eSummaryDataTable"
                      className="table table-bordered"
                      style={{ width: "100%", tableLayout: "fixed" }}
                    >
                      <colgroup>
                        <col style={{ width: "15%" }} />
                        <col style={{ width: "15%" }} />
                        <col style={{ width: "10%" }} />
                        <col style={{ width: "20%" }} />
                        <col style={{ width: "40%" }} />
                      </colgroup>
                      <thead>
                        <tr>
                          <th>Initiator/Approver</th>
                          <th>Forwarded To</th>
                          <th>Action Date</th>
                          <th>Action</th>
                          <th>Action Remarks</th>
                        </tr>
                      </thead>
                      <tbody>
                        {/* You can map rows dynamically here */}
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>





            </div>
          </div>


        )}
      </Form>
    </Formik>
  );
};


