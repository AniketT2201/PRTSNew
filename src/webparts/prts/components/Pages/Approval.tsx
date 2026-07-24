import * as React from 'react';
import { useState, useEffect, createContext, useContext } from 'react';
import { IPrtsProps } from '../IPrtsProps';
import ButtonBar from './ButtonBar';
import { useHistory, useParams } from 'react-router-dom';
import RemarksModal from './RemarksModal';
import AssignToAnotherUserModal from './AssigntoAnotherUser';
import BaseInfoTab from './Tabfirst';
import { sp } from "@pnp/sp";
import { format } from "date-fns";
// import Tab2TechnicalIssue from "./TechnicalIssueData";
import Tab2TechnicalIssueFull from './TechnicalIssueData';
import Tab10Summary from './NewTabSummary';
import SPCRUDOPS from '../../service/DAL/spcrudops';
import IEmployeeProfileops from '../../service/BAL/SPCRUD/EmployeeProfile';
import '../../components/Pages/CSS/NewPage.scss';
const AppContext = createContext(null);

const AppProvider = ({ children }) => {
  const [currentUser, setCurrentUser] = useState(null);
  const [approvalWF, setApprovalWF] = useState(null);

  useEffect(() => {
    //window.location.reload();
    setCurrentUser({ name: 'John Doe', roles: ['Approver', 'Requester'] });
    setApprovalWF({ status: 'In Progress' });
  }, []);

  return <AppContext.Provider value={{ currentUser, approvalWF }}>{children}</AppContext.Provider>;
};


const useAppContext = () => useContext(AppContext);

const Header = () => (
  // <div className="row mainTitle">
  //   <div className="col-xs-9">
  //     <span style={{ fontSize: '32px' }}>Problem Resolution Tracking System</span>
  //   </div>
  // </div>
  <div className="header">
    <div className="left-banner">
      <div className="logo-text">
        <h2>Problem Resolution Tracking System</h2>
      </div>
    </div>
  </div>
);
declare const _spPageContextInfo: {
  webAbsoluteUrl: string;
  webServerRelativeUrl: string;
  userDisplayName: string;
};
interface TechIssueData {
  assignDate: string;
  issueAssignTo: string;
  agencyName: string;
  mNTAnalysis: string;
  mNTRootCauseFound: string;
  mNTICA_Details: string;
  mNTICA_VIN: string;
  mNTPCA_Details: string;
  mNTPCA_VIN: string;
  mNT_Remarks: string;
  mNT_RootCause: string;
}

const formatDateTime = (date: Date) => date.toISOString();

const Approval = (props: IPrtsProps) => {
  //  const params = useParams()
  const { RequestId } = useParams<{ RequestId: string }>();
  const [dataLoaded, setDataLoaded] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState("New Request");
  const [ChStatus, setCHStatus] = useState('');
  const [jsonSummary, setJsonSummary] = useState<Array<{ c1: string; c2: string; c3: string; c4: string; c5: string }>>([]);
  const [isRemarksOpen, setRemarksOpen] = useState(false);
  const [remarksTitle, setRemarksTitle] = useState('');
  const [remarksFor, setRemarksFor] = useState('');
  const [isAssignUserOpen, setAssignUserOpen] = useState(false);
  const [requestNumber, setRequestNumber] = useState("");
  const [requestDate] = useState(`${new Date()}`);
  const [problemDescription, setProblemDescription] = useState("");
  const [activeTab, setActiveTab] = useState('basic');
  const [editMode, setEditMode] = useState(false);
  const [reqId, setRequestId] = useState<string | undefined>(undefined);
  const [canEdit, setCanEdit] = useState(false);
  const [showAttachmentTable, setShowAttachmentTable] = useState(false);
  const [currentTab, setCurrentTab] = useState("Basic");
  const [visibleTabs, setVisibleTabs] = useState({});
  const [canEditTabs, setCanEditTabs] = useState({});
  const [commodityselected, setcommodityselected] = useState<string>('');
  const [severityLevel, setSeverityLevel] = useState<string>('');
  const [baseInfoData, setBaseInfoData] = React.useState({
    mIssueCategory: '',
    mRepeatedIssue: '',
    mInitDept: "",
    mInitName: "",
    mIssueStatus: "",
    mIs7D: "",
    mRootCauseFound: "",
    mAnalysis: "",
    mPurgingAttachment: '',
    mTitle: "",
    mPartName: "",
    mPartNo: "",
    mPartSupplier: "",
    mPRTSSource: "",
    mProjectCode: "",
    mIssueVINNo: "",
    mMFGShop: "",
    mIssueDescription: "",
    mCategory: "",
    mSeverity: "",
    mQtyAffected: "",
    mVariantAffected: "",
    mEngineType: "",
    mIsRepeated: "",
    mRefNo: "",
    mCommodity: "",
    mBuildType: "",
    mAgency: "",
    mPartQualityIssue: [],
    mPartSupplierSource: ''
  });
  const [activeTechData, setActiveTechData] = useState<TechIssueData>({
    assignDate: "",
    issueAssignTo: "",
    agencyName: "",
    mNTAnalysis: "",
    mNTRootCauseFound: "",
    mNTICA_Details: "",
    mNTICA_VIN: "",
    mNTPCA_Details: "",
    mNTPCA_VIN: "",
    mNT_Remarks: "",
    mNT_RootCause: ""
  });
  const [historyTechData, setHistoryTechData] = useState<any[]>([]);

  // Button states
  const [Stage, setStage] = useState<number | null>(null);
  const [CurrUser, setCurrUser] = useState(props.userDisplayName);
  const [NextApprover, setNextApprover] = useState("");
  const [DeleApprover, setDeleApprover] = useState("");
  const [InitName, setInitName] = useState("");
  const [initnameEmail, setinitnameEmail] = useState("");
  const [Is7D, setIs7D] = useState("");
  const [IsRootCauseFound, setIsRootCauseFound] = useState("");
  const [SelectedTab, setSelectedTab] = useState("");
  const [d1JsonArray, setD1JsonArray] = useState<any[]>([]);
  const [d2JsonArray, setD2JsonArray] = useState<any[]>([]);
  const [d3JsonArray, setD3JsonArray] = useState<any[]>([]);
  const [d4JsonArray, setD4JsonArray] = useState<any[]>([]);
  const [d5JsonArray, setD5JsonArray] = useState<any[]>([]);
  const [d6JsonArray, setD6JsonArray] = useState<any[]>([]);
  const [d7JsonArray, setD7JsonArray] = useState<any[]>([]);

const [userDisplayName, setUserDisplayName] = useState(props.userDisplayName);
const [CHstatusselected, setCHStatusSelected] = useState("");
const [nonTechnicalIssueData, setNonTechnicaIssueData] = useState("");
  //---------------------------------------------------------------------------------------

  //_______________________________________________________________________________________
  // const [activeTab, setActiveTab] = React.useState('basic');

  const [buttons, setButtons] = useState({
    btnClose: true,
    btnCreateDraft: true,
    btnSubmit: false,
    btnWithDrawn: false,
    btnPrint: false,
    btnCloseIssue: false,
    btnProcessWithIssueCloseNT: false,
    btnBackToInitiator: false,
    btnAssignIssueToAnoterUser: false,
    btnsubmitforreview: false,
    btnReject: false,
    btnWrongIssueAssign: false,
    btnclickRework: false
  });

  const defaultButtonsState = {
    btnClose: true,
    btnCreateDraft: false,
    btnSubmit: false,
    btnWithDrawn: false,
    btnPrint: false,
    btnCloseIssue: false,
    btnProcessWithIssueCloseNT: false,
    btnBackToInitiator: false,
    btnAssignIssueToAnoterUser: false,
    btnsubmitforreview: false,
    btnReject: false,
    btnWrongIssueAssign: false,
    btnclickRework: false
  };

  sp.setup({
    spfxContext: props.currentSPContext
  });


//   const resetAllStateForNewRequest = () => {
//   setRequestId(undefined);
//   setRequestNumber("");
//   setProblemDescription("");

//   // Base Info
//   setBaseInfoData({
//     mIssueCategory: '',
//     mRepeatedIssue: '',
//     mInitDept: "",
//     mInitName: userDisplayName || "",
//     mIssueStatus: "Open",
//     mIs7D: "",
//     mRootCauseFound: "",
//     mAnalysis: "",
//     mPurgingAttachment: '',
//     mTitle: "",
//     mPartName: "",
//     mPartNo: "",
//     mPartSupplier: "",
//     mPRTSSource: "",
//     mProjectCode: "",
//     mIssueVINNo: "",
//     mMFGShop: "",
//     mIssueDescription: "",
//     mCategory: "",
//     mSeverity: "",
//     mQtyAffected: "",
//     mVariantAffected: "",
//     mEngineType: "",
//     mIsRepeated: "",
//     mRefNo: "",
//     mCommodity: "",
//     mBuildType: "",
//     mAgency: "",
//     mPartQualityIssue: [],
//     mPartSupplierSource: ''
//   });

//   // Technical
//   setActiveTechData({
//     assignDate: "",
//     issueAssignTo: "",
//     agencyName: "",
//     mNTAnalysis: "",
//     mNTRootCauseFound: "",
//     mNTICA_Details: "",
//     mNTICA_VIN: "",
//     mNTPCA_Details: "",
//     mNTPCA_VIN: "",
//     mNT_Remarks: "",
//     mNT_RootCause: ""
//   });


//   // Reset UI
//   setActiveTab("basic");
//   setCurrentTab("Basic");
//   setStage(null);
//   setCHStatus("");
//   setStatus("");
// };

useEffect(() => {
  if (props.Reader) {
    setButtons(defaultButtonsState);
  }
}, [props.Reader]);

useEffect(() => {
  // ✏️ EDIT REQUEST → LOAD DATA
  if (RequestId) {
    setRequestId(RequestId);
    setDataLoaded(false);
    handleBaseInfoGetNewPage(RequestId);
  }
}, [RequestId]);

  // -----------------------------
  // NOTE: this function now returns { createdId } if a new item was created.
  // This allows BaseInfoTab to call onSave() and then upload its pending attachments.
  // -----------------------------
const checkReferencePresentInList = async (refNo: string) => {
  if (!refNo) return false;

  const items = await sp.web.lists
    .getByTitle("PRTSList")
    .items
    .filter(`ReqNo eq '${refNo}'`)
    .top(1)();

  return items.length > 0;
};




//   const handleBaseInfoSave = async (formState: any) => {
//     let newItemId;
    
//      const cleanedForm = { ...formState };
//       Object.keys(cleanedForm).forEach((key) => {
//         if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
//           cleanedForm[key] = "";
//         }
//       });
      
//       if(!problemDescription?.trim()){
//       alert("Problem Description or Title are mandatory fields.");
//       return;
//       }
//       // Reference Request Number validation

//       if (cleanedForm.mRepeatedIssue === "Yes") {

//         const refNo = cleanedForm.mRefNo?.trim();

//         if (!refNo) {
//           alert("Reference Request Number is mandatory for Repeated Issue.");
//           return;
//         }

//         const exists = await checkReferencePresentInList(refNo);
 
//         if (!exists) {
//           alert(`Reference Request No '${refNo}' does not exist in the system.`);
//           return; // ❌ STOP SAVE
//         }
//       }
//     try {
//       setLoading(true);
//       const EmployeeId = await GetUserDepartment(props, props.userEmail);
//       // const userDisplayName = props.userDisplayName;
//       updateSummary(userDisplayName, "", formatDateTime(new Date()), "Request Created", "");
//       const summary = JSON.stringify(jsonSummary);
//       const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
//       const addResult = await sp.web.lists.getByTitle("PRTSList").items.add({ Status: "Open", Summary: summary, InitiatorId: userId, Stage: 0, InitDepartment: EmployeeId?.DepartmentCode?.Department, InitiatorEmpId: props.EmployeeId[0].EmployeeID });
//        newItemId = addResult.data.Id;
//       const reqNo = newItemId.toString().padStart(5, "0");
//       const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
//       setRequestNumber(fullReqNo);
//       await sp.web.lists.getByTitle("PRTSList").items.getById(newItemId).update({ ReqNo: fullReqNo });
//       //       await appendSummaryAndPersist(newItemId!, {
//       //   c1: props.userDisplayName,
//       //   c2: "",
//       //   c3: formatDateTime(new Date()),
//       //   c4: "Request Created",
//       //   c5: ""
//       // }); 
   
//     } catch (error) {
//       console.error("Error creating request:", error);
//       alert("Error creating request: " + error.message);
//     }
//     try {
     
//       const userId = props.currentSPContext.pageContext.legacyPageContext.userId;

//       // Prepare Non-Technical JSON
//       const Agency = cleanedForm.mAgency;
//       const userName = cleanedForm.mPartQualityIssue?.Name || "";

//       const today = new Date();
//       const formattedDate =
//         today.getFullYear() +
//         "-" +
//         String(today.getMonth() + 1).padStart(2, "0") +
//         "-" +
//         String(today.getDate()).padStart(2, "0");

//       const nonTechJson = [
//         {
//           c1: Agency,
//           c2: userName,
//           c3: formattedDate
//         }
//       ];
//       const finalNonTechJson = JSON.stringify(nonTechJson);
//       const NEXTaPPOVEReMPLOEEID = await IEmployeeProfileops().getEmployeeProfile((cleanedForm.mPartQualityIssue?.Email||cleanedForm.diamondUsers[0].Email), props)
//       const updateData: any = {
//         Title: problemDescription || cleanedForm.mTitle || "",
//         PartName: cleanedForm.mPartName,
//         PartNumbe: cleanedForm.mPartNo,
//         SupplierName: cleanedForm.mPartSupplier,
//         PRTSSource: cleanedForm.mPRTSSource,
//         ProjectCode: cleanedForm.mProjectCode,
//         BuildType: cleanedForm.mBuildType,
//         VINNo: cleanedForm.mIssueVINNo,
//         MFGShopSelection: cleanedForm.mMFGShop,
//         IssueDescription: cleanedForm.mIssueDescription,
//         IssueCategory: cleanedForm.mIssueCategory,
//         Severity: cleanedForm.mSeverity,
//         QtyAffected: cleanedForm.mQtyAffected,
//         VariantAffected: cleanedForm.mVariantAffected,
//         RepeatedIssue: cleanedForm.mRepeatedIssue,
//         RefReqNo: cleanedForm.mRefNo,
//         Commodity: cleanedForm.mCommodity,
//         SupplierSource: cleanedForm.mPartSupplierSource,
//         EngineType: cleanedForm.mEngineType,
//         InitDepartment: cleanedForm.mInitDept,
//         InitiatorId: userId,
//         SelectedTab: "Basic",
//         NonTechnical_IssueData: finalNonTechJson,

//         NextApproverEmpID: NEXTaPPOVEReMPLOEEID[0]?.EmployeeID || null,
//         AnalysisDetails: cleanedForm.mAnalysis,
//         Is7DRequired: cleanedForm.mIs7D,
//         IsRootCauseFound: cleanedForm.mRootCauseFound,
//         PurgingAttachment:cleanedForm.selectedFiles[0]?.name || ""
//       };

//       // Same logic as before for CH_Status and ApproverList
//       // (D1-D7 disciplines removed: these branches no longer clear D1_IssueData..D7_IssueData)
//       if (cleanedForm.mRootCauseFound === "Yes") {
//         updateData.CH_Status = "1/6";
//         updateData.ApproverList = `${userName}`;
//         updateData.NAId = cleanedForm.mPartQualityIssue?.Id || null;
//       } else if (cleanedForm.mIs7D === "No") {
//         updateData.CH_Status = "1/6";
//         updateData.NonTechnical_IssueData = finalNonTechJson;
//         updateData.ApproverList = `${userName}`;
//         updateData.NAId = cleanedForm.mPartQualityIssue?.Id || null;

//       }

//       // If reqId exists -> update; else create and return createdId (so BaseInfoTab can upload pending files)
//       if (newItemId) {
//         await sp.web.lists.getByTitle("PRTSList").items.getById((newItemId)).update(updateData);
//          await uploadAttachments(newItemId, cleanedForm.selectedFiles);
//         updateSummary(
//           props.userDisplayName,
//           "",
//           format(new Date(), "yyyy-MM-dd"),
//           "Details updated",
//           ""
//         );
//         alert("Details updated successfully");
//               history.push(`/InitiatorLandingedit/${newItemId}`);
//  // existing item updated: no createdId to return
//       } else {
//         // Create a new item
//         const addResult = await sp.web.lists.getByTitle("PRTSList").items.add(updateData);
//         const createdId = addResult.data?.Id ?? (addResult as any).id ?? null;
//         if (createdId) {
//           // set reqId and generate ReqNo as you did in handleCreateDraft
//           const reqNo = String(createdId).padStart(5, "0");
//           const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
//           await sp.web.lists.getByTitle("PRTSList").items.getById(createdId).update({ ReqNo: fullReqNo });


//           setRequestNumber(fullReqNo);
//           setRequestId(String(createdId));
//           updateSummary(
//             props.userDisplayName,
//             "",
//             format(new Date(), "yyyy-MM-dd"),
//             "Details created",
//             ""
//           );
//           alert("Details created successfully");
//           // IMPORTANT: return createdId so child (BaseInfoTab) can upload pending files
//           return { createdId };
//         } else {
//           alert("Item created but ID not returned by PnPJS — attachments cannot be uploaded automatically.");
//           return;
//         }
//       }
//     } catch (error) {
//       console.log("SAVE ERROR:", error);
//       alert("Error saving details: " + (error.message ? error.message : error));
//       throw error;
//     } finally {
//       setLoading(false);
//     }
//   };

// const handleBaseInfoSave = async (formState: any, reqId?: string | number) => {
//     let newItemId: number | undefined = reqId ? Number(reqId) : undefined;
    
//     const cleanedForm = { ...formState };
//     Object.keys(cleanedForm).forEach((key) => {
//       if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
//         cleanedForm[key] = "";
//       }
//     });
//     const removeEmptyFields = (obj: any) => {
//       return Object.fromEntries(
//         Object.entries(obj).filter(
//           ([_, value]) =>
//             value !== undefined &&
//             value !== null &&
//             value !== ""
//         )
//       );
//     };

//     if (!problemDescription?.trim()) {
//       alert("Problem Description or Title are mandatory fields.");
//       return;
//     }
//     // Reference Request Number validation

//     if (cleanedForm.mRepeatedIssue === "Yes") {

//       const refNo = cleanedForm.mRefNo?.trim();

//       if (!refNo) {
//         alert("Reference Request Number is mandatory for Repeated Issue.");
//         return;
//       }

//       const exists = await checkReferencePresentInList(refNo);

//       if (!exists) {
//         alert(`Reference Request No '${refNo}' does not exist in the system.`);
//         return; // ❌ STOP SAVE
//       }
//     }

//     // Create SharePoint item ONLY if this is a brand new request (no existing draft id passed in)
//     if (!newItemId) {
//       try {
//         setLoading(true);
//         const EmployeeId = await GetUserDepartment(props, props.userEmail);
//         // const userDisplayName = props.userDisplayName;
//         const newEntry = {
//           c1: props.userDisplayName,
//           c2: "",
//           c3: formatDateTime(new Date()),
//           c4: "Request Submitted",
//           c5: "Technical Issue"
//         }
//         const updatedSummary = [newEntry];
//         setJsonSummary(updatedSummary);
//         const summary = JSON.stringify(updatedSummary);
//         const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
//         const addResult = await sp.web.lists.getByTitle("PRTSList").items.add({ Status: "Open", Summary: summary, InitiatorId: userId, Stage: 0, InitDepartment: EmployeeId?.DepartmentCode?.Department, InitiatorEmpId: props.EmployeeId[0].EmployeeID });
//         newItemId = addResult.data.Id;
//         const reqNo = newItemId.toString().padStart(5, "0");
//         const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
//         setRequestNumber(fullReqNo);
//         await sp.web.lists.getByTitle("PRTSList").items.getById(newItemId).update({ ReqNo: fullReqNo });

//       } catch (error) {
//         console.error("Error creating request:", error);
//         alert("Error creating request: " + error.message);
//         setLoading(false);
//         return; // stop execution — do not continue with an undefined newItemId
//       }
//     } else {
//       // Existing draft - skip creation, just show loading state
//       setLoading(true);
//     }

//     // Single source of truth for the item id from here on — no duplicate create/update branching
//     const itemId = newItemId!;

//     try {

//       const userId = props.currentSPContext.pageContext.legacyPageContext.userId;

//       // Prepare Non-Technical JSON
//       const Agency = cleanedForm.mAgency;
//       const initiatorName = cleanedForm.mPartQualityIssue?.Name || "";

//       const today = new Date();
//       const formattedDate =
//         today.getFullYear() +
//         "-" +
//         String(today.getMonth() + 1).padStart(2, "0") +
//         "-" +
//         String(today.getDate()).padStart(2, "0");

//       const nonTechJson = [
//         {
//           c1: Agency,
//           c2: initiatorName,
//           c3: formattedDate
//         }
//       ];
//       const finalNonTechJson = JSON.stringify(nonTechJson);
//       const NEXTaPPOVEReMPLOEEID = await IEmployeeProfileops().getEmployeeProfile((cleanedForm.mPartQualityIssue?.Email || null), props)
//       const updateData: any = {
//         Title: problemDescription || cleanedForm.mTitle || "",
//         PartName: cleanedForm.mPartName,
//         PartNumbe: cleanedForm.mPartNo,
//         SupplierName: cleanedForm.mPartSupplier,
//         PRTSSource: cleanedForm.mPRTSSource,
//         ProjectCode: cleanedForm.mProjectCode,
//         BuildType: cleanedForm.mBuildType,
//         VINNo: cleanedForm.mIssueVINNo,
//         MFGShopSelection: cleanedForm.mMFGShop,
//         IssueDescription: cleanedForm.mIssueDescription,
//         IssueCategory: cleanedForm.mIssueCategory,
//         Severity: cleanedForm.mSeverity,
//         QtyAffected: cleanedForm.mQtyAffected,
//         VariantAffected: cleanedForm.mVariantAffected,
//         RepeatedIssue: cleanedForm.mRepeatedIssue,
//         RefReqNo: cleanedForm.mRefNo,
//         Commodity: cleanedForm.mCommodity,
//         SupplierSource: cleanedForm.mPartSupplierSource,
//         EngineType: cleanedForm.mEngineType,
//         InitDepartment: cleanedForm.mInitDept,
//         InitiatorId: userId,
//         SelectedTab: "Basic",
//         NonTechnical_IssueData: finalNonTechJson,

//         NextApproverEmpID: NEXTaPPOVEReMPLOEEID[0]?.EmployeeID || null,
//         AnalysisDetails: cleanedForm.mAnalysis,
//         // Is7DRequired: cleanedForm.mIs7D,
//         IsRootCauseFound: cleanedForm.mRootCauseFound,
//         PurgingAttachment:cleanedForm.selectedFiles[0]?.name || ""
//       };

//       // Same logic as before for CH_Status and ApproverList
//       // (D1-D7 disciplines removed: these branches no longer clear D1_IssueData..D7_IssueData)
//       if (cleanedForm.mRootCauseFound === "Yes") {
//         updateData.CH_Status = "1/6";
//         updateData.ApproverList = `${initiatorName}`;
//         updateData.NAId = cleanedForm.mPartQualityIssue?.Id || null;
//       } else if (cleanedForm.mIs7D === "No") {
//         updateData.CH_Status = "1/6";
//         updateData.NonTechnical_IssueData = finalNonTechJson;
//         updateData.ApproverList = `${initiatorName}`;
//         updateData.NAId = cleanedForm.mPartQualityIssue?.Id || null;

//       }

//       // Save Base Information fields — single update call, used for BOTH new and existing requests
//       await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update(updateData);

//       // Single upload block - runs for BOTH new and existing requests
//       await uploadAttachments(itemId, cleanedForm.selectedFiles);

//       // ===== Submit logic (merged from handleonSubmitRequest) =====

//       const submitData: any = {};

//       // Status, Stage and SelectedTab must always be updated on Submit,
//       // regardless of whether an approver/agency was assigned
//       submitData.Status = "In process - Tech";
//       submitData.Stage = 2;
//       submitData.SelectedTab = "NT";
//       const item = await sp.web.lists
//         .getByTitle("PRTSList")
//         .items.getById(newItemId)
//         .select("Summary")();
      
//       const existingSummary = item.Summary
//         ? JSON.parse(item.Summary)
//         : [];
//       const newEntry = {
//         c1: props.userDisplayName,
//         c2: initiatorName ? initiatorName : initiatorName,
//         c3: formatDateTime(new Date()),
//         c4: "Request Submitted",
//         c5: "Technical Issue"
//       };
//       existingSummary.push(newEntry);
//       setJsonSummary(existingSummary);
//       submitData.Summary = JSON.stringify(existingSummary);

//       // Only approver assignment + delegation logic stays conditional
//       if (activeTechData && activeTechData.issueAssignTo) {
//         const t = [
//           {
//             c1: activeTechData.agencyName,
//             c2: activeTechData.issueAssignTo,
//             c3: formatDate(new Date()),
//             c4: activeTechData.mNTAnalysis,
//             c5: activeTechData.mNTRootCauseFound,
//             c6: activeTechData.mNTICA_Details,
//             c7: activeTechData.mNTICA_VIN,
//             c8: activeTechData.mNTPCA_Details,
//             c9: activeTechData.mNTPCA_VIN,
//             c10: activeTechData.mNT_Remarks,
//             c13: activeTechData.mNT_RootCause
//           }
//         ];
//         const assignTo = activeTechData.issueAssignTo;
//         submitData.NAId = await getUserId(assignTo);
//         const delegated = await checkDelegation(assignTo);
//         submitData.DAId = delegated ? await getUserId(delegated) : null;
//         submitData.NonTechnical_IssueData = JSON.stringify(t);
//       }
      
//       // Always update summary, LastAction, EmailSendFlag
//       submitData.LastAction = new Date();
//       submitData.EmailSendFlag = 1;
//       //submitData.ReqNo = requestNumber;

//       // Single submit update - the only SharePoint update for submit-related fields
//       await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update(submitData);
//       alert("Request submitted successfully");
//       history.push("/");

//     } catch (error) {
//       console.log("SAVE ERROR:", error);
//       alert("Error submitting details: " + (error.message ? error.message : error));
//       throw error;
//     } finally {
//       setLoading(false);
//     }
//   };

// const handleBaseInfoSaveDraft = async (formState: any) => {
//   try {
//   setLoading(true);

//   const cleanedForm = { ...formState };

//   Object.keys(cleanedForm).forEach((key) => {
//     if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
//       cleanedForm[key] = "";
//     }
//   });

//   // Validation
//   if (!problemDescription?.trim()) {
//     alert("Problem Description or Title is mandatory.");
//     return;
//   }

//   if (cleanedForm.mRepeatedIssue === "Yes") {
//     const refNo = cleanedForm.mRefNo?.trim();

//     if (!refNo) {
//       alert("Reference Request Number is mandatory for Repeated Issue.");
//       return;
//     }

//     const exists = await checkReferencePresentInList(refNo);

//     if (!exists) {
//       alert(`Reference Request No '${refNo}' does not exist in the system.`);
//       return;
//     }
//   }

//   const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
//   // Prepare Non-Technical JSON
//   const Agency = cleanedForm.mAgency;
//   const initiatorName = cleanedForm.mPartQualityIssue?.Name || "";

//   const today = new Date();
//   const formattedDate =
//     today.getFullYear() +
//     "-" +
//     String(today.getMonth() + 1).padStart(2, "0") +
//     "-" +
//     String(today.getDate()).padStart(2, "0");

//   const nonTechJson = [
//     {
//       c1: Agency,
//       c2: initiatorName,
//       c3: formattedDate
//     }
//   ];
//   const finalNonTechJson = JSON.stringify(nonTechJson);

//   let itemId = reqId ? Number(reqId) : null;
//   let isNewDraft = false;

//   // ==========================
//   // CREATE NEW DRAFT
//   // ==========================
//   if (!itemId) {
//     isNewDraft = true;

//     const employee = await GetUserDepartment(
//       props,
//       props.userEmail
//     );

//     const addResult = await sp.web.lists
//       .getByTitle("PRTSList")
//       .items.add({
//         Status: "Draft",
//         Stage: 0,
//         SelectedTab: "Basic",
//         InitiatorId: userId,
//         InitDepartment: employee?.DepartmentCode?.Department,
//         InitiatorEmpId: props.EmployeeId?.[0]?.EmployeeID || ""
//       });

//     itemId = addResult.data.Id;

//     const reqNo = String(itemId).padStart(5, "0");
//     const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;

//     await sp.web.lists
//       .getByTitle("PRTSList")
//       .items.getById(itemId)
//       .update({
//         ReqNo: fullReqNo
//       });

//     setRequestNumber(fullReqNo);
//     setRequestId(String(itemId));
//   }

//   // ==========================
//   // SUMMARY
//   // ==========================
//   // If your updateSummary updates jsonSummary state asynchronously,
//   // consider replacing this with appendSummaryAndPersist().
//   const newEntry = {
//     c1: props.userDisplayName,
//     c2: "",
//     c3: formatDateTime(new Date()),
//     c4: isNewDraft ? "Draft Created" : "Draft Created",
//     c5: ""
//   }
//   const updatedSummary = [newEntry];
//   setJsonSummary(updatedSummary);
//   const summary = JSON.stringify(updatedSummary);

//   // ==========================
//   // UPDATE DRAFT DATA
//   // ==========================
//   const updateData: any = {
//     Title: problemDescription || cleanedForm.mTitle || "",
//     PartName: cleanedForm.mPartName,
//     PartNumbe: cleanedForm.mPartNo,
//     SupplierName: cleanedForm.mPartSupplier,
//     PRTSSource: cleanedForm.mPRTSSource,
//     ProjectCode: cleanedForm.mProjectCode,
//     BuildType: cleanedForm.mBuildType,
//     VINNo: cleanedForm.mIssueVINNo,
//     MFGShopSelection: cleanedForm.mMFGShop,
//     IssueDescription: cleanedForm.mIssueDescription,
//     IssueCategory: cleanedForm.mIssueCategory,
//     Severity: cleanedForm.mSeverity, 
//     QtyAffected: cleanedForm.mQtyAffected,
//     VariantAffected: cleanedForm.mVariantAffected,
//     RepeatedIssue: cleanedForm.mRepeatedIssue,
//     RefReqNo: cleanedForm.mRefNo,
//     Commodity: cleanedForm.mCommodity,
//     SupplierSource: cleanedForm.mPartSupplierSource,
//     EngineType: cleanedForm.mEngineType,
//     InitDepartment: cleanedForm.mInitDept,
//     InitiatorId: userId,

//     AnalysisDetails: cleanedForm.mAnalysis,
//     Is7DRequired: cleanedForm.mIs7D,
//     IsRootCauseFound: cleanedForm.mRootCauseFound,

//     PurgingAttachment: cleanedForm.selectedFiles?.[0]?.name || "",

//     SelectedTab: "Basic",
//     NonTechnical_IssueData: finalNonTechJson,

//     // Draft fields
//     Status: "Draft",
//     Stage: 0,
//     EmailSendFlag: 0,
//     LastAction: new Date(),

//     Summary: summary
//   };

//   await sp.web.lists
//     .getByTitle("PRTSList")
//     .items.getById(itemId)
//     .update(updateData);

//   // ==========================
//   // ATTACHMENTS
//   // ==========================
//   if (cleanedForm.selectedFiles && cleanedForm.selectedFiles.length > 0) {
//     await uploadAttachments(
//       itemId,
//       cleanedForm.selectedFiles
//     );
//   }

//   alert(
//     isNewDraft
//       ? "Draft created successfully"
//       : "Draft updated successfully"
//   );

//   // Redirect only on first save
//   // if (isNewDraft) {
//   //   history.push(`/InitiatorLandingedit/${itemId}`);
//   // }
//   history.push('/');

//   } catch (error: any) {
//     console.error("Draft Save Error:", error);
//     alert("Error saving draft: " + (error?.message || error));
//   } finally {
//     setLoading(false);
//   }
// };

const handleBaseInfoSave = async (formState: any, reqId?: string | number) => {
  let newItemId: number | undefined = reqId ? Number(reqId) : undefined;
 
  const cleanedForm = { ...formState };
  Object.keys(cleanedForm).forEach((key) => {
    if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
      cleanedForm[key] = "";
    }
  });
  const removeEmptyFields = (obj: any) => {
    return Object.fromEntries(
      Object.entries(obj).filter(
        ([_, value]) => value !== undefined && value !== null && value !== ""
      )
    );
  };
 
  if (!problemDescription?.trim()) {
    alert("Problem Description or Title are mandatory fields.");
    return;
  }
  // Reference Request Number validation
  if (cleanedForm.mRepeatedIssue === "Yes") {
    const refNo = cleanedForm.mRefNo?.trim();
    if (!refNo) {
      alert("Reference Request Number is mandatory for Repeated Issue.");
      return;
    }
    const exists = await checkReferencePresentInList(refNo);
    if (!exists) {
      alert(`Reference Request No '${refNo}' does not exist in the system.`);
      return; // ❌ STOP SAVE
    }
  }
 
  // Create SharePoint item ONLY if this is a brand new request (no existing draft id passed in)
  if (!newItemId) {
    try {
      setLoading(true);
      const EmployeeId = await GetUserDepartment(props, props.userEmail);
      const newEntry = {
        c1: props.userDisplayName,
        c2: "",
        c3: formatDateTime(new Date()),
        c4: "Request Submitted",
        c5: "Technical Issue"
      };
      const updatedSummary = [newEntry];
      setJsonSummary(updatedSummary);
      const summary = JSON.stringify(updatedSummary);
      const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
      const addResult = await sp.web.lists.getByTitle("PRTSList").items.add({ Status: "Open", Summary: summary, InitiatorId: userId, Stage: 0, InitDepartment: EmployeeId?.DepartmentCode?.Department, InitiatorEmpId: props.EmployeeId[0].EmployeeID });
      newItemId = addResult.data.Id;
      const reqNo = newItemId.toString().padStart(5, "0");
      const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
      setRequestNumber(fullReqNo);
      await sp.web.lists.getByTitle("PRTSList").items.getById(newItemId).update({ ReqNo: fullReqNo });
    } catch (error) {
      console.error("Error creating request:", error);
      alert("Error creating request: " + error.message);
      setLoading(false);
      return; // stop execution — do not continue with an undefined newItemId
    }
  } else {
    // Existing draft - skip creation, just show loading state
    setLoading(true);
  }
 
  // Single source of truth for the item id from here on — no duplicate create/update branching
  const itemId = newItemId!;
 
  try {
    const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
 
    // Prepare Non-Technical JSON
    const Agency = cleanedForm.mAgency;
    const initiatorName = cleanedForm.mPartQualityIssue?.Name || "";
 
    const today = new Date();
    const formattedDate =
      today.getFullYear() +
      "-" +
      String(today.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(today.getDate()).padStart(2, "0");
 
    const nonTechJson = [{ c1: Agency, c2: initiatorName, c3: formattedDate }];
    const finalNonTechJson = JSON.stringify(nonTechJson);
    const NEXTaPPOVEReMPLOEEID = await IEmployeeProfileops().getEmployeeProfile((cleanedForm.mPartQualityIssue?.Email || null), props);
 
    const updateData: any = {
      Title: problemDescription || cleanedForm.mTitle || "",
      PartName: cleanedForm.mPartName,
      PartNumbe: cleanedForm.mPartNo,
      SupplierName: cleanedForm.mPartSupplier,
      PRTSSource: cleanedForm.mPRTSSource,
      ProjectCode: cleanedForm.mProjectCode,
      BuildType: cleanedForm.mBuildType,
      VINNo: cleanedForm.mIssueVINNo,
      MFGShopSelection: cleanedForm.mMFGShop,
      IssueDescription: cleanedForm.mIssueDescription,
      IssueCategory: cleanedForm.mIssueCategory,
      Severity: cleanedForm.mSeverity,
      QtyAffected: cleanedForm.mQtyAffected,
      VariantAffected: cleanedForm.mVariantAffected,
      RepeatedIssue: cleanedForm.mRepeatedIssue,
      RefReqNo: cleanedForm.mRefNo,
      Commodity: cleanedForm.mCommodity,
      SupplierSource: cleanedForm.mPartSupplierSource,
      EngineType: cleanedForm.mEngineType,
      InitDepartment: cleanedForm.mInitDept,
      InitiatorId: userId,
      SelectedTab: "Basic",
      NonTechnical_IssueData: finalNonTechJson,
 
      // FIX #2: NAId / NextApproverEmpID set as BASE fields so NextApprover
      // is always persisted, regardless of which CH_Status branch fires below.
      NAId: cleanedForm.mPartQualityIssue?.Id || null,
      NextApproverEmpID: NEXTaPPOVEReMPLOEEID[0]?.EmployeeID || null,
 
      AnalysisDetails: cleanedForm.mAnalysis,
      IsRootCauseFound: cleanedForm.mRootCauseFound,
      PurgingAttachment: cleanedForm.selectedFiles[0]?.name || ""
    };
 
    // Same logic as before for CH_Status and ApproverList
    // (D1-D7 disciplines removed: these branches no longer clear D1_IssueData..D7_IssueData)
    if (cleanedForm.mRootCauseFound === "Yes") {
      updateData.CH_Status = "1/6";
      updateData.ApproverList = `${initiatorName}`;
    } else if (cleanedForm.mIs7D === "No") {
      updateData.CH_Status = "1/6";
      updateData.NonTechnical_IssueData = finalNonTechJson;
      updateData.ApproverList = `${initiatorName}`;
    }
 
    // FIX #1: strip empty/undefined/null fields before merging into an
    // EXISTING item so untouched form fields don't blank out saved data.
    await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update(removeEmptyFields(updateData));
 
    // Single upload block - runs for BOTH new and existing requests
    await uploadAttachments(itemId, cleanedForm.selectedFiles);
 
    // ===== Submit logic (merged from handleonSubmitRequest) =====
    const submitData: any = {};
    submitData.Status = "In process - Tech";
    submitData.Stage = 2;
    submitData.SelectedTab = "NT";
 
    const item = await sp.web.lists.getByTitle("PRTSList").items.getById(newItemId).select("Summary")();
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: initiatorName ? initiatorName : initiatorName,
      c3: formatDateTime(new Date()),
      c4: "Request Submitted",
      c5: "Technical Issue"
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    submitData.Summary = JSON.stringify(existingSummary);
 
    if (activeTechData && activeTechData.issueAssignTo) {
      const t = [{
        c1: activeTechData.agencyName,
        c2: activeTechData.issueAssignTo,
        c3: formatDate(new Date()),
        c4: activeTechData.mNTAnalysis,
        c5: activeTechData.mNTRootCauseFound,
        c6: activeTechData.mNTICA_Details,
        c7: activeTechData.mNTICA_VIN,
        c8: activeTechData.mNTPCA_Details,
        c9: activeTechData.mNTPCA_VIN,
        c10: activeTechData.mNT_Remarks,
        c13: activeTechData.mNT_RootCause
      }];
      const assignTo = activeTechData.issueAssignTo;
      submitData.NAId = await getUserId(assignTo);
      const delegated = await checkDelegation(assignTo);
      submitData.DAId = delegated ? await getUserId(delegated) : null;
      submitData.NonTechnical_IssueData = JSON.stringify(t);
    }
 
    submitData.LastAction = new Date();
    submitData.EmailSendFlag = 1;
 
    await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update(submitData);
    alert("Request submitted successfully");
    history.push("/");
  } catch (error) {
    console.log("SAVE ERROR:", error);
    alert("Error submitting details: " + (error.message ? error.message : error));
    throw error;
  } finally {
    setLoading(false);
  }
};
 
const handleBaseInfoSaveDraft = async (formState: any) => {
  try {
    setLoading(true);
 
    const cleanedForm = { ...formState };
    Object.keys(cleanedForm).forEach((key) => {
      if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
        cleanedForm[key] = "";
      }
    });
    const removeEmptyFields = (obj: any) => {
      return Object.fromEntries(
        Object.entries(obj).filter(
          ([_, value]) => value !== undefined && value !== null && value !== ""
        )
      );
    };
 
    // Validation
    if (!problemDescription?.trim()) {
      alert("Problem Description or Title is mandatory.");
      return;
    }
    if (cleanedForm.mRepeatedIssue === "Yes") {
      const refNo = cleanedForm.mRefNo?.trim();
      if (!refNo) {
        alert("Reference Request Number is mandatory for Repeated Issue.");
        return;
      }
      const exists = await checkReferencePresentInList(refNo);
      if (!exists) {
        alert(`Reference Request No '${refNo}' does not exist in the system.`);
        return;
      }
    }
 
    const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
    const Agency = cleanedForm.mAgency;
    const initiatorName = cleanedForm.mPartQualityIssue?.Name || "";
 
    const today = new Date();
    const formattedDate =
      today.getFullYear() +
      "-" +
      String(today.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(today.getDate()).padStart(2, "0");
 
    const nonTechJson = [{ c1: Agency, c2: initiatorName, c3: formattedDate }];
    const finalNonTechJson = JSON.stringify(nonTechJson);
 
    // FIX #2: resolve the next approver's employee record for drafts too,
    // same as handleBaseInfoSave does — previously this was never computed here.
    const NEXTaPPOVEReMPLOEEID = await IEmployeeProfileops().getEmployeeProfile((cleanedForm.mPartQualityIssue?.Email || null), props);
 
    let itemId = reqId ? Number(reqId) : null;
    let isNewDraft = false;
 
    // ==========================
    // CREATE NEW DRAFT
    // ==========================
    if (!itemId) {
      isNewDraft = true;
      const employee = await GetUserDepartment(props, props.userEmail);
 
      const addResult = await sp.web.lists.getByTitle("PRTSList").items.add({
        Status: "Draft",
        Stage: 0,
        SelectedTab: "Basic",
        InitiatorId: userId,
        InitDepartment: employee?.DepartmentCode?.Department,
        InitiatorEmpId: props.EmployeeId?.[0]?.EmployeeID || ""
      });
 
      itemId = addResult.data.Id;
      const reqNo = String(itemId).padStart(5, "0");
      const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
 
      await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update({ ReqNo: fullReqNo });
 
      setRequestNumber(fullReqNo);
      setRequestId(String(itemId));
    }
 
    // ==========================
    // SUMMARY
    // ==========================
    const newEntry = {
      c1: props.userDisplayName,
      c2: "",
      c3: formatDateTime(new Date()),
      c4: isNewDraft ? "Draft Created" : "Draft Created",
      c5: ""
    };
    const updatedSummary = [newEntry];
    setJsonSummary(updatedSummary);
    const summary = JSON.stringify(updatedSummary);
 
    // ==========================
    // UPDATE DRAFT DATA
    // ==========================
    const updateData: any = {
      Title: problemDescription || cleanedForm.mTitle || "",
      PartName: cleanedForm.mPartName,
      PartNumbe: cleanedForm.mPartNo,
      SupplierName: cleanedForm.mPartSupplier,
      PRTSSource: cleanedForm.mPRTSSource,
      ProjectCode: cleanedForm.mProjectCode,
      BuildType: cleanedForm.mBuildType,
      VINNo: cleanedForm.mIssueVINNo,
      MFGShopSelection: cleanedForm.mMFGShop,
      IssueDescription: cleanedForm.mIssueDescription,
      IssueCategory: cleanedForm.mIssueCategory,
      Severity: cleanedForm.mSeverity,
      QtyAffected: cleanedForm.mQtyAffected,
      VariantAffected: cleanedForm.mVariantAffected,
      RepeatedIssue: cleanedForm.mRepeatedIssue,
      RefReqNo: cleanedForm.mRefNo,
      Commodity: cleanedForm.mCommodity,
      SupplierSource: cleanedForm.mPartSupplierSource,
      EngineType: cleanedForm.mEngineType,
      InitDepartment: cleanedForm.mInitDept,
      InitiatorId: userId,
 
      // FIX #2: NextApprover now actually gets persisted for drafts.
      NAId: cleanedForm.mPartQualityIssue?.Id || null,
      NextApproverEmpID: NEXTaPPOVEReMPLOEEID[0]?.EmployeeID || null,
      ApproverList: initiatorName ? `${initiatorName}` : undefined,
 
      AnalysisDetails: cleanedForm.mAnalysis,
      Is7DRequired: cleanedForm.mIs7D,
      IsRootCauseFound: cleanedForm.mRootCauseFound,
 
      PurgingAttachment: cleanedForm.selectedFiles?.[0]?.name || "",
 
      SelectedTab: "Basic",
      NonTechnical_IssueData: finalNonTechJson,
 
      // Draft fields
      Status: "Draft",
      Stage: 0,
      EmailSendFlag: 0,
      LastAction: new Date(),
 
      Summary: summary
    };
 
    // FIX #1: strip empty/undefined/null fields before merging so partially
    // filled draft saves don't blank out previously saved values.
    await sp.web.lists.getByTitle("PRTSList").items.getById(itemId).update(removeEmptyFields(updateData));
 
    // ==========================
    // ATTACHMENTS
    // ==========================
    if (cleanedForm.selectedFiles && cleanedForm.selectedFiles.length > 0) {
      await uploadAttachments(itemId, cleanedForm.selectedFiles);
    }
 
    alert(isNewDraft ? "Draft created successfully" : "Draft updated successfully");
    history.push('/');
  } catch (error: any) {
    console.error("Draft Save Error:", error);
    alert("Error saving draft: " + (error?.message || error));
  } finally {
    setLoading(false);
  }
};


const uploadAttachments = async (itemId: number, files?: File[]) => {
  if (!files || files.length === 0) return;

  const item = sp.web.lists
    .getByTitle("PRTSList")
    .items.getById(itemId);

  for (const file of files) {
    await item.attachmentFiles.add(file.name, file);
  }
};


  useEffect(() => {
    if (RequestId) {
      setRequestId(RequestId);
    }
     setVisibleTabs({
      basic: true,
      technical: false,
      summary: true
    });
    handleBaseInfoGetNewPage(RequestId);

  }, [RequestId]);

  
  const updateSummary = (c1: string, c2: string, c3: string, c4: string, c5: string) => {
    const newEntry = { c1, c2, c3, c4, c5 };
    setJsonSummary(prevSummary => [...prevSummary, newEntry]);
  };
  const getDigest = async () => {
    const res = await fetch(`${props.context.pageContext._site.absoluteUrl}/PRTS/_api/contextinfo`, { method: 'POST', headers: { Accept: 'application/json;odata=verbose' } });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  };
  const history = useHistory();

  const handleClose = () => {
    history.push('/');
  };
  async function GetUserDepartment(props, Email) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('UserMaster', '*,FullName/EMail,DepartmentCode/Department', 'FullName,DepartmentCode', `FullName/EMail eq '${Email}'`, { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.DepartmentCode, item])).values());
    const first = uniqueModels[0] as any;
    // setFormData(prev => ({ ...prev, mInitDept: first?.DepartmentCode?.Department || "" }));
    return first;
  }

  useEffect(() => {
    if (!RequestId) return;
    if (Stage === null || InitName === undefined) return;

    const norm = (s: string) => s?.trim().toLowerCase() || "";

    const isInitiator = norm(CurrUser) === norm(InitName);

    const isApprover =
      norm(CurrUser) === norm(NextApprover) ||
      norm(CurrUser) === norm(DeleApprover);

    if (!dataLoaded) {
      console.log("Waiting for data...");
      return;
    }

    const show = (list: string) => {
      const all = { ...defaultButtonsState };

      if (list !== "AllHide") {
        list
          .split(",")
          .map((x) => x.trim())
          .forEach((id) => {
            if ((all as any).hasOwnProperty(id)) {
              (all as any)[id] = true;
            }
          });
      }

      setButtons(all);
    };

    console.log("PRTS Button Logic", {
      Stage,
      CHstatusselected,
      CurrUser,
      InitName,
      NextApprover,
      DeleApprover,
      isInitiator,
      isApprover,
    });

    // Default
    show("AllHide");

    // New Request
    if (!RequestId) {
      if (isInitiator) {
        show("btnClose");
      }
      return;
    }

    // Existing Request in Draft
    if (status?.trim() === "Draft") {
      if (isInitiator) {
        show("btnClose");
      }
      return;
    }

    switch (CHstatusselected) {
      /* 1/6 Agency user when updating technical issue */
      case "1/6":
        const nonTechData = nonTechnicalIssueData
          ? JSON.parse(nonTechnicalIssueData)
          : [];

        const isEqualToCurrentUser = nonTechData.some(
          item =>
            item.c2?.trim().toLowerCase() ===
            props.userDisplayName?.trim().toLowerCase()
        );

        // if (isEqualToCurrentUser) {
        //   show("btnSubmit");
        // }

        // if (isApprover) {
        //   show("btnSubmit");
        // }

        break;

      /*** 2/6* Agency User submits for review* Wrong Issue Assigned*/
      case "2/6":
        if (isApprover) {
          show("btnsubmitforreview");
        }

        // Initiator can reassign issue
        if (isInitiator) {
          show(
            "btnsubmitforreview,btnWrongIssueAssign,btnAssignIssueToAnoterUser"
          );
        }
        break;

      /*** 3/6* Commodity Lead*/
      case "3/6":
        if (isApprover) {
          show(
            "btnsubmitforreview,btnWrongIssueAssign,btnclickRework"
          );
        }
        break;

      /*** 4/6* Initiator review*/
      case "4/6":
        if (isInitiator) {
          show("btnsubmitforreview,btnAssignIssueToAnoterUser,btnWithDrawn,btnPrint");
        }
        break;

      /**
       * 5/6
       * Manager review
       */
      case "5/6":
        if (isApprover) {
          show("btnCloseIssue,btnReject");
        }
        break;

      /**
       * 6/6
       * Closed
       */
      case "6/6":
        show("AllHide");
        break;

      default:
        show("AllHide");
        break;
    }
  }, [dataLoaded,RequestId,Stage,CurrUser,InitName,NextApprover,DeleApprover,CHstatusselected,]);
  const CheckValidSubmit = () => {
    const hasNT =
      activeTechData &&
      (activeTechData.issueAssignTo?.trim() !== "" ||
        activeTechData.agencyName?.trim() !== "");
    if (!hasNT) {
      alert("Please provide 'Basic Information' before you submit the request.");
      return false;
    }
    return true;
  };
  const formatDate = (date: Date) => {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  };
  const getUserId = async (loginOrDisplayName: string): Promise<number> => {
    try {
      const ensure = await sp.web.ensureUser(loginOrDisplayName);
      return ensure.data.Id;
    } catch (err) {
      console.error("User resolution failed:", err);
      throw new Error("Cannot find user: " + loginOrDisplayName);
    }
  };
  const checkDelegation = async (approverName: string): Promise<string | null> => {
    try {
      const items = await sp.web.lists
        .getByTitle("DelegationMaster")
        .items.filter(`Approver eq '${approverName}' and IsActive eq 1`)
        .select("DelegatedTo")
        .get();
      if (items.length > 0) {
        return items[0].DelegatedTo;
      }
      return null;
    } catch (e) {
      console.error("Delegation lookup failed:", e);
      return null;
    }
  };

  const handleonSubmitRequest = async (reqIdParam: string | undefined) => {
    try {
      console.log("Submit clicked");
      setLoading(true);
      // const isValidSubmit = CheckValidSubmit();
      // if (!isValidSubmit) return;
      if (!reqIdParam) {
        alert("Request ID missing.");
        return;
      }

      let updateObject: any = {};

      // Non-technical
      if (activeTechData && activeTechData.issueAssignTo) {
        const t = [
          {
            c1: activeTechData.agencyName,
            c2: activeTechData.issueAssignTo,
            c3: formatDate(new Date()),
            c4: activeTechData.mNTAnalysis,
            c5: activeTechData.mNTRootCauseFound,
            c6: activeTechData.mNTICA_Details,
            c7: activeTechData.mNTICA_VIN,
            c8: activeTechData.mNTPCA_Details,
            c9: activeTechData.mNTPCA_VIN,
            c10: activeTechData.mNT_Remarks,
            c13: activeTechData.mNT_RootCause
          }
        ];
        const assignTo = activeTechData.issueAssignTo;
        updateObject.NAId = await getUserId(assignTo);
        const delegated = await checkDelegation(assignTo);
        updateObject.DAId = delegated ? await getUserId(delegated) : null;
        updateObject.NonTechnical_IssueData = JSON.stringify(t);
        updateObject.SelectedTab = "NT";
        updateObject.Status = "In process - Tech";
        updateObject.Stage = 2;
        // updateSummary(
        //   props.userDisplayName,
        //   delegated ? `${assignTo} (${delegated})` : assignTo,
        //   formatDateTime(new Date()),
        //   "Request Submitted",
        //   "Technical Issue"
        // );
        const newEntry = {
        c1: props.userDisplayName,
        c2: assignTo ? assignTo : assignTo,
        c3: formatDateTime(new Date()),
        c4: "Request Submitted",
        c5: "Technical Issue"
      };
      const updatedSummary = [newEntry];
      setJsonSummary(updatedSummary);

      updateObject.Summary = JSON.stringify(updatedSummary);
      } 
      // Always update summary, LastAction, EmailSendFlag, ReqNo
      updateObject.LastAction = new Date();
      updateObject.EmailSendFlag = 1;
        updateObject.Stage = 2;
      //updateObject.ReqNo = requestNumber;
      await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqIdParam)).update(updateObject);
      alert("Request submitted successfully");
      history.push("/");
    } catch (err: any) {
      alert("Submit error: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  };

  const handleonWithdrawn = () => {
    openRemarksModal(4);
  };
  const handleonPrint = () => { 
    /* TODO: implement printing */
    window.print();
  };
  const handleonCloseIssue = () => {
    openRemarksModal(6);
  };
  const handleonReturnBackToPITMember = () => {
    openRemarksModal(1);
  };
  const handleonForwardToNextPITMember = () => {
    openRemarksModal(2);
  };

  // Validates the current in-flight container before "Process With Issue Close".
  // D1-D7 discipline tabs have been removed; only the Technical ("NT") container
  // carries a root-cause validation requirement now.
  const CheckValidProcessWithIssueClose = (container: string): boolean => {
    if (container === "NT") {
      if (!activeTechData) {
        alert("Missing Data for NT");
        return false;
      }
      if (activeTechData.mNTRootCauseFound !== "Yes") {
        setActiveTab("technical");
        alert('Change [Is Root Cause Found] = "Yes".\nIn Technical Tab');
        return false;
      }
      return true;
    }
    return true;
  };

  const handleProcessWithIssueClose = async () => {
    try {
      setLoading(true);
      console.log("Process With Issue Close clicked");
      const isValid = CheckValidProcessWithIssueClose(currentTab);
      if (!isValid) {
        alert("Please complete required fields before processing.");
        return; 
      }
      if (!reqId) {
        alert("Request ID missing");
        return;
      }
      const updateObj: any = {};
      const initiatorName = InitName;
      updateObj.NAId = await getUserId(initiatorName);
      const delegated = await checkDelegation(initiatorName);
      updateObj.DAId = delegated ? await getUserId(delegated) : null;
      updateObj.Stage = 1;
      updateObj.CH_Status = "5/6";
      updateObj.SelectedTab = "Basic";
      updateObj.Status = "In process - Initiator";
      updateObj.IssueResolveRequestBy = currentTab;

      // NOTE: previously you read DOM fields and AttachmentList variables that didn't exist.
      // We set safe defaults here; if you are storing attachment lists in parent state, replace the empty strings with those variables.
      updateObj.SOSJESValue = "";
      updateObj.ControlPlanValue = "";
      updateObj.PFMEAValue = "";
      updateObj.KaizenValue = "";
      updateObj.QualityAlertValue = "";

      // Also set attachment fields — currently blank; wire to actual attachment lists if needed
      updateObj.SOSJESAttachment = "";
      updateObj.ControlPlan = "";
      updateObj.PFMEAAttachment = "";
      updateObj.KaizenAttachment = "";
      updateObj.QualityAlertAttachment = "";

      const newEntry = {
        c1: props.userDisplayName,
        c2: initiatorName ? initiatorName : initiatorName,
        c3: formatDateTime(new Date()),
        c4: "Issue Resolved",
        c5: "Sent To Initiator for Review and Close"
      };
      const updatedSummary = [newEntry];
      setJsonSummary(updatedSummary);

      updateObj.Summary = JSON.stringify(updatedSummary);
      updateObj.LastAction = new Date();
      updateObj.EmailSendFlag = 1;

      await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update(updateObj);

      console.log("Process With Issue Close saved");

      // Previously you uploaded purging docs using globals (fileData). That implementation relied on globals
      // and would crash. BaseInfoTab now manages attachments itself. If you still need to upload additional
      // files from parent, implement the uploadPurgingDocs function properly and supply the required data.
      // For now we call a no-op stub to maintain parity with previous code that conditionally called uploadPurgingDocs().
      if (typeof uploadPurgingDocs === "function") {
        try {
          await uploadPurgingDocs();
        } catch (err) {
          console.warn("uploadPurgingDocs failed or is not implemented:", err);
        }
      }

      alert("Issue processed successfully");
      history.push("/");
    } catch (err: any) {
      alert("Error: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  };

  // Safe stub for uploadPurgingDocs — previous implementation relied on undefined globals.
  // TODO: Replace this with real upload logic if needed. If using BaseInfoTab attachments, this stub can be removed.
  const uploadPurgingDocs = async () => {
    console.warn("uploadPurgingDocs(): not implemented in parent. BaseInfoTab handles attachments now.");
    return;
  };

  const handleonSendBackToPreviousStage = () => { openRemarksModal(3); };
  const handleonAssignToChampion = () => { /* TODO */ };
  const handleonProcessWithIssueCloseNT = async () => { 
    try {
      console.log("Process With Issue Close clicked");
      setLoading(true);
       const isValid = 
      activeTechData &&
      (activeTechData.issueAssignTo?.trim() !== "" ||
        activeTechData.agencyName?.trim() !== "");
      if (!isValid) {
        alert("Please complete required fields before processing.");
        return;
      }
      if (!reqId) {
        alert("Request ID missing");
        return;
      }
      const updateObj: any = {};
      const initiatorName = InitName;
      updateObj.NAId = await getUserId(initiatorName);
      const delegated = await checkDelegation(initiatorName);
      updateObj.DAId = delegated ? await getUserId(delegated) : null;
      updateObj.Stage = 1;
      updateObj.CH_Status = "5/6";
      updateObj.SelectedTab = "Basic";
      updateObj.Status = "In process - Initiator";
      updateObj.IssueResolveRequestBy = currentTab;

      // NOTE: previously you read DOM fields and AttachmentList variables that didn't exist.
      // We set safe defaults here; if you are storing attachment lists in parent state, replace the empty strings with those variables.
      updateObj.SOSJESValue = "";
      updateObj.ControlPlanValue = "";
      updateObj.PFMEAValue = "";
      updateObj.KaizenValue = "";
      updateObj.QualityAlertValue = "";

      // Also set attachment fields — currently blank; wire to actual attachment lists if needed
      updateObj.SOSJESAttachment = "";
      updateObj.ControlPlan = "";
      updateObj.PFMEAAttachment = "";
      updateObj.KaizenAttachment = "";
      updateObj.QualityAlertAttachment = "";

      const item = await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(reqId))
        .select("Summary")();
  
      const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
      const newEntry = {
        c1: props.userDisplayName,
        c2: initiatorName ? initiatorName : initiatorName,
        c3: formatDateTime(new Date()),
        c4: "Issue Resolved",
        c5: "Sent To Initiator for Review and Close"
      };
      existingSummary.push(newEntry);
      setJsonSummary(existingSummary);
      updateObj.Summary = JSON.stringify(existingSummary);

      updateObj.LastAction = new Date();
      updateObj.EmailSendFlag = 1;

      await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update(updateObj);

      console.log("Process With Issue Close saved");

      // Previously you uploaded purging docs using globals (fileData). That implementation relied on globals
      // and would crash. BaseInfoTab now manages attachments itself. If you still need to upload additional
      // files from parent, implement the uploadPurgingDocs function properly and supply the required data.
      // For now we call a no-op stub to maintain parity with previous code that conditionally called uploadPurgingDocs().
      if (typeof uploadPurgingDocs === "function") {
        try {
          await uploadPurgingDocs();
        } catch (err) {
          console.warn("uploadPurgingDocs failed or is not implemented:", err);
        }
      }

      alert("Issue processed successfully");
      history.push("/");
    } catch (err: any) {
      alert("Error: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  
   };
  const handleonBackToInitiator = () => { openRemarksModal(5); };
  const handleonReassignIssue = () =>  setAssignUserOpen(true);
  const closeRemarksModal = () => setRemarksOpen(false);

  const onRemarksUpdate = (remarks: string) => {
    console.log('Updated remarks for', remarksFor, ':', remarks);
    setRemarksOpen(false);
  };
  const remarksMap = {
    1: { key: 'ReturnBackToPITMember', title: 'Return Back to PIT Member' },
    2: { key: 'ForwardToNextPITMember', title: 'Forward to Next PIT Member' },
    3: { key: 'SendBackToPreviousStage', title: 'Send Back to Previous Stage' },
    4: { key: 'WithDrawn', title: 'Withdrawn' },
    5: { key: 'SendBackToInitiator', title: 'Send Back to Initiator' },
    6: { key: 'CloseIssue', title: 'Close Issue' },
  };
  const openRemarksModal = (btnId: number) => {
    const remark = remarksMap[btnId];
    if (remark) {
      setRemarksFor(remark.key);
      setRemarksTitle(remark.title);
      setRemarksOpen(true);
    }
  };

  const handleRemarksUpdate = (remarks: string) => {
    if (!remarks.trim()) {
      alert('Remarks cannot be blank.');
      return;
    }
    switch (remarksFor) {
      case 'SendBackToPreviousStage':
        sendBackToPreviousStage();
        break;
      case 'WithDrawn':
        withdrawnRequest();
        break;
      case 'SendBackToInitiator':
        sendBackToInitiator();
        break;
      case 'CloseIssue':
        closeIssue(RequestId, remarks);
        break;
      default:
        break;
    }
    closeRemarksModal();
  };

  const closeIssue = async (reqIdParam: string | undefined, remarks: string) => {
    try {
      setLoading(true);
      if (!reqIdParam) {
        alert("Request ID missing");
        return;
      }
  
      const updateObj: any = {
        NAId: null,
        DAId: null,
        Stage: 3,
        Status: "Close",
        CH_Status: "6/6",
        IssueResolveRequestBy: "",
        LastAction: new Date(),
        EmailSendFlag: 1
      };
  
      // SUMMARY (fetch current + append, so history isn't lost)
      const item = await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(reqIdParam))
        .select("Summary")();
  
      const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
      const newEntry = {
        c1: props.userDisplayName,
        c2: "",
        c3: formatDateTime(new Date()),
        c4: "Issue Closed",
        c5: remarks
      };
      existingSummary.push(newEntry);
      setJsonSummary(existingSummary);
      updateObj.Summary = JSON.stringify(existingSummary);
  
      await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(reqIdParam))
        .update(updateObj);
  
      console.log("Data Saved.");
      history.push("/");
    } catch (error: any) {
      alert("Error saving issue closing: " + (error?.message || error) + ". Refresh and try again, or contact Administrator.");
    } finally {
      setLoading(false);
    }
  };

  const openAssignUserModal = () => setAssignUserOpen(true);
  const closeAssignUserModal = () => setAssignUserOpen(false);
const handleAssignSubmit = async (data: { agency: string; status: string; user: string; remarks: string }) => {
  try {
    setLoading(true);
    const { agency, status, user, remarks } = data;
 
    if (!reqId) {
      alert("Request ID missing");
      return;
    }
 
    // -------------------------
    // 1️⃣ Resolve Next Approver (NA)
    // -------------------------
    const nextApproverId = await getUserId(user);
 
    // -------------------------
    // 2️⃣ Check Delegation (DA)
    // -------------------------
    const delegatedName = await checkDelegation(user);
    const delegatedId = delegatedName ? await getUserId(delegatedName) : null;
 
    // -------------------------
    // 3️⃣ Fetch current Summary + NonTechnical_IssueData from SharePoint
    //    (single source of truth, instead of relying on possibly-stale
    //    historyTechData/jsonSummary React state)
    // -------------------------
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary", "NonTechnical_IssueData")();
 
    const today = formatDate(new Date());
 
    const oldNT = item.NonTechnical_IssueData ? JSON.parse(item.NonTechnical_IssueData) : [];
 
    // Add new row (c1 = agency, c2 = user, c3 = date)
    const newRow = {
      c1: agency,
      c2: user,
      c3: today
    };
 
    const updatedNT = [...oldNT, newRow];
    setHistoryTechData(updatedNT);
 
    // -------------------------
    // 4️⃣ Build Approver List (Initiator + New User)
    // -------------------------
    const approverList = `${InitName};${user}`;
 
    // -------------------------
    // 5️⃣ Build update object for SharePoint
    // -------------------------
    const updateObj: any = {
      NAId: nextApproverId,
      DAId: delegatedId,
      Stage: 2,
      Status: "In process - Tech",
      SelectedTab: "NT",
      CH_Status: status,
 
      NonTechnical_IssueData: JSON.stringify(updatedNT),
 
      ApproverList: approverList,
 
      LastAction: new Date(),
      EmailSendFlag: 1
    };
 
    // -------------------------
    // 6️⃣ Add Summary entry (fetched above, appended here)
    // -------------------------
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: user ? user : user,
      c3: formatDateTime(new Date()),
      c4: "Re-Assign",
      c5: remarks
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);
 
    // -------------------------
    // 7️⃣ Update SharePoint item
    // -------------------------
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);
 
    alert("Issue Re-Assigned Successfully");
 
    closeAssignUserModal();
    history.push("/");
 
  } catch (err: any) {
    alert("Re-Assign Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
};

const appendSummaryAndPersist = async (
  reqIdParam: string,
  entry: { c1: string; c2: string; c3: string; c4: string; c5: string },
  extraUpdate?: any
) => {
  const nextSummary = [...jsonSummary, entry];

  // update state
  setJsonSummary(nextSummary);

  // persist to SharePoint
  await sp.web.lists
    .getByTitle("PRTSList")
    .items.getById(Number(reqIdParam))
    .update({
      Summary: JSON.stringify(nextSummary),
      LastAction: new Date(),
      EmailSendFlag: 1,
      ...(extraUpdate || {})
    });
};


  const formatDateForInput = (dateString: string) => {
    const date = new Date(dateString);
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const day = ('0' + date.getDate()).slice(-2);
    return `${year}-${month}-${day}`;
  };

  const applyStagePermissions = (StatusParam: any, StageParam: any, CurrUserParam: any, NextApproverParam: any, DeleApproverParam: any, InitNameParam: any, TabSelected: any) => {
    const StatusTab = ["basic", "technical"];
    const selectedIndex = StatusTab.indexOf((TabSelected || "basic").toLowerCase());
    let visible: any = {};
    let editable: any = {};
    StatusTab.forEach(t => { visible[t] = true; editable[t] = false; });
    // Hide Technical tab in Draft status
    if (StatusParam === "Draft") {
      visible.technical = false;
    }
    if (StageParam > 0 && CurrUserParam === InitNameParam && CurrUserParam === NextApproverParam) {
      StatusTab.forEach(t => editable[t] = false);
    } else if (CurrUserParam !== NextApproverParam && CurrUserParam !== DeleApproverParam) {
      StatusTab.forEach(t => editable[t] = false);
    } else {
      StatusTab.forEach((t, i) => { editable[t] = i === selectedIndex; });
    }
    setVisibleTabs(visible);
    setCanEditTabs(editable);
  };

  const handleBaseInfoGetNewPage = async (reqIdParam: string | undefined) => {
    try {
      setLoading(true);
      if (!reqIdParam) return;
      const item = await sp.web.lists.getByTitle("PRTSList").items
        .getById(Number(reqIdParam))
        .select("Title", "Severity", "RefReqNo", "Commodity", "SupplierSource", "InitDepartment", "IsRootCauseFound", "Is7DRequired", "AnalysisDetails", "IssueStatus", "NonTechnical_IssueData", "Initiator/Title","Initiator/EMail", "Initiator/Id", "Stage", "NA/Id", "NA/Title", "DA/Id", "DA/Title", "SelectedTab", "ReqNo", "CH_Status", "Status")
        .expand("Initiator", "NA", "DA")
        .get();

      const StageVal = item.Stage;
      const CurrUserVal = props.userDisplayName;
      const NextApproverVal = item?.NA?.Title || "";
      const DeleApproverVal = item?.DA?.Title || "";
      const InitNameVal = item?.Initiator?.Title || "";
      const TabSelected = item.SelectedTab;
      setcommodityselected(item.Commodity || "");
      setSeverityLevel(item.Severity || "");
      setCHStatusSelected(item.CH_Status || "");
      setStage(StageVal);
      setNextApprover(NextApproverVal);
      setDeleApprover(DeleApproverVal);
      setInitName(InitNameVal);
      setinitnameEmail(item?.Initiator?.EMail || "");
      setIs7D(item.Is7DRequired || "");
      setIsRootCauseFound(item.IsRootCauseFound || "");
      setSelectedTab(item.SelectedTab || "");
      setCHStatus(item.CH_Status);
      setStatus(item.Status);
      setNonTechnicaIssueData(item.NonTechnical_IssueData);
      applyStagePermissions(item.Status, StageVal, CurrUserVal, NextApproverVal, DeleApproverVal, InitNameVal, TabSelected);
      setRequestNumber(item.ReqNo);
      setProblemDescription(item.Title);
      updateTabsBasedOnConditions(item.IsRootCauseFound || "Select", item.Is7DRequired || "Select", item.NonTechnical_IssueData || "");
      console.log("Before setDataLoaded");

setDataLoaded(true);

console.log("After setDataLoaded");
    } catch (error) {
      console.log(error);
      if (error instanceof Error) {
        console.error(error.message);
        console.error(error.stack);
      }
    } finally {
      setLoading(false);
    }
  };

  const tabs = [
    { id: 'basic', label: 'BASE INFORMATION' },
    { id: 'technical', label: 'TECHNICAL ISSUE' },
    { id: 'summary', label: 'SUMMARY' },
  ];

  function handleTechnicalIssueSave(data: TechIssueData, updatedJson?: any[]): void {
    setActiveTechData(data);
    if (updatedJson) setHistoryTechData(updatedJson);
  }

  
  const updateTabsBasedOnConditions = ( rootCauseFound: string,is7D: string, nonTechIssueData?: string) => {

    // Rule 2: 7D = No and root cause = No
    const nonTechData = nonTechIssueData ? JSON.parse(nonTechIssueData) : [];
    const isEqualToCurrentUser = nonTechData.some(item => 
      item.c2?.trim().toLowerCase() === props.userDisplayName?.trim().toLowerCase());
    
    if (isEqualToCurrentUser) {
      setVisibleTabs({
        basic: true,
        technical: true,
        summary: true
      });
    }

    // Rule 3: Root Cause Found Yes → Only Basic + Technical + Summary
    if (rootCauseFound === "Yes" && !isEqualToCurrentUser) {
      setVisibleTabs({
        basic: true,
        technical: false,
        summary: true
      });
      return;
    }
  };

  // NEW: Sync form data from child to parent (this prevents reset)
const handleBaseInfoChange = (updatedData: any) => {
  setBaseInfoData(prev => ({ ...prev, ...updatedData }));
};

async function sendBackToPreviousStage() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 1,
      Status: "Sent Back to Previous Stage",
      SelectedTab: currentTab.toUpperCase(),
      CH_Status: "2/6"
    };

    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: "",
      c3: formatDateTime(new Date()),
      c4: "Sent Back to Previous Stage",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);

    //await updateRequest(reqId, updateObj, "Sent Back to Previous Stage");
    alert("Sent Back to Previous Stage successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

const validateProcessWithIssueOpen = (tab: string): boolean => {
  const json = getDxJson(tab);

  if (!json || json.length === 0) {
    alert("No data entered in the current tab.");
    return false;
  }

  const last = json[json.length - 1];

 if (!last.c3 || last.c3.trim() === "") {
  alert("Please fill analysis (c3) before forwarding.");
  setActiveTab(tab.toLowerCase());
  return false;
}


  return true;
};

const getDxJson = (tab: string): any[] => {
  switch (tab) {
    case "D1": return d1JsonArray;
    case "D2": return d2JsonArray;
    case "D3": return d3JsonArray;
    case "D4": return d4JsonArray;
    case "D5": return d5JsonArray;
    case "D6": return d6JsonArray;
    case "D7": return d7JsonArray;
    default: return [];
  }
};
const handleonProcessWithIssueOpen = async () => {
  try {
    setLoading(true);
    const tab = currentTab.toUpperCase();

    if (!validateProcessWithIssueOpen(tab)) return;

    // Determine next tab
    const nextNum = Number(tab.replace("D", "")) + 1;
    if (nextNum > 7) {
      alert("D7 is the last stage!");
      return;
    }
    const nextTab = `D${nextNum}`;

    // Get JSON of current tab
    const json = getDxJson(tab);
    const last = json[json.length - 1];

    // Resolve approver from JSON c1
    const nextApproverName = last.c1;
    const nextApproverId = await getUserId(nextApproverName);

    const delegatedName = await checkDelegation(nextApproverName);
    const delegatedId = delegatedName ? await getUserId(delegatedName) : null;

    last.c2 = formatDate(new Date()); // Set date
    json[json.length - 1] = last;

    const updateObj: any = {
      NAId: nextApproverId,
      DAId: delegatedId,
      [`${nextTab}_IssueData`]: JSON.stringify(json),
      SelectedTab: nextTab,
      Status: `In process- ${nextTab}`,
      LastAction: new Date(),
      EmailSendFlag: 1,
    };


    const newEntry = {
      c1: props.userDisplayName,
      c2: nextApproverName ? nextApproverName : nextApproverName,
      c3: formatDateTime(new Date()),
      c4: "Open Issue Forwarded",
      c5: ""
    };
    const updatedSummary = [newEntry];
    setJsonSummary(updatedSummary);

    updateObj.Summary = JSON.stringify(updatedSummary);

    await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update(updateObj);

    alert("Forwarded Successfully!");
    history.push("/");

  } catch (e: any) {
    alert("Error: " + (e?.message || e));
  } finally {
    setLoading(false);
  }
};

async function withdrawnRequest() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 3,
      Status: "Withdrawn",
      CH_Status: "0/6",
      NAId: null,
      NextApproverEmpID: "",
      DAId: null
    };
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: "",
      c3: formatDateTime(new Date()),
      c4: "Request Withdrawn",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);

    //await updateRequest(reqId, updateObj, "Request Withdrawn");
    alert("Request withdrawn successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

async function sendBackToInitiator() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 1,
      Status: "Returned to Initiator",
      SelectedTab: "Basic",
      CH_Status: "1/6"
    };
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: InitName ? InitName : InitName,
      c3: formatDateTime(new Date()),
      c4: "Sent Back to Initiator",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);

    //await updateRequest(reqId, updateObj, "Sent Back to Initiator");
    alert("Returned to Initiator successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}
const updateRequest = async (reqIdParam: string, updateObject: any, summaryMessage: string) => {
  const nextSummary = [
    ...jsonSummary,
    {
      c1: props.userDisplayName,
      c2: '',
      c3: formatDateTime(new Date()),
      c4: summaryMessage,
      c5: ''
    }
  ];

  updateObject.Summary = JSON.stringify(nextSummary);
  updateObject.LastAction = new Date();
  updateObject.EmailSendFlag = 1;

  await sp.web.lists
    .getByTitle("PRTSList")
    .items.getById(Number(reqIdParam))
    .update(updateObject);
};
const getCommodityApprovers = async (commodity: string) => {
  if (!commodity) return null;

  const spCrudOps = await SPCRUDOPS();
  const data = await spCrudOps.getRootData(
    "CommodityList",
    "*,CommodityLead/Id,CommodityLead/EMail,CommodityHead/Title,CommodityHead/Id,CommodityHead/EMail",
    "CommodityLead,CommodityHead",
    `Title eq '${commodity}'`,
    { column: "ID", isAscending: true },
    props
  );

  if (!data.length) return null;

  return {
    lead: {
      id: data[0].CommodityLead?.Id,
      email: data[0].CommodityLead?.EMail,
      Title: data[0].CommodityLead?.Title,
      // EmployeeId: data[0].EmployeeId
    },
    head: {
      id: data[0].CommodityHead?.Id,
      email: data[0].CommodityHead?.EMail,
      Title: data[0].CommodityHead?.Title
    }
  };
};

const submitforreview = async () => {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }
 
    const approvers = await getCommodityApprovers(commodityselected);
    if (!approvers) {
      alert("Commodity approvers not configured");
      return;
    }
 
    const updateObj: any = {};
    const severity = severityLevel;
 
    // Track who the request is moving to and why, for the Summary entry below
    let nextApproverName = "";
    let summaryMessage = "";
 
    // =====================================================
    // 1️⃣ AGENCY → COMMODITY LEAD
    // =====================================================
    if (CHstatusselected === "2/6") {
      updateObj.Stage = 3;
      updateObj.CH_Status = "3/6";
      updateObj.NAId = approvers.lead.id;
      updateObj.NextApproverEmpID =
        await GetApproverEmployeeId(approvers.lead.email);
      updateObj.Status = "Pending with Commodity Lead";
      nextApproverName = approvers.lead.Title || "";
      summaryMessage = "Submitted for Review - Forwarded to Commodity Lead";
    }
 
    // =====================================================
    // 2️⃣ COMMODITY LEAD → COMMODITY HEAD (Severity = 50)
    // =====================================================
    else if (CHstatusselected === "3/6" && severity === "50" && approvers.head.Title != NextApprover) {
      updateObj.Stage = 3;
      updateObj.CH_Status = "3/6";
      updateObj.NAId = approvers.head.id;
      updateObj.NextApproverEmpID =
        await GetApproverEmployeeId(approvers.head.email);
      updateObj.Status = "Pending with Commodity Head";
      nextApproverName = approvers.head.Title || "";
      summaryMessage = "Submitted for Review - Forwarded to Commodity Head";
    }
 
    // =====================================================
    // 3️⃣ COMMODITY LEAD → INITIATOR (Severity ≠ 50)
    // =====================================================
    else if (CHstatusselected === "3/6") {
      updateObj.Stage = 4;
      updateObj.CH_Status = "4/6";
      updateObj.NAId = await getUserId(InitName);
      updateObj.NextApproverEmpID =
        await GetApproverEmployeeId(initnameEmail);
      updateObj.Status = "Pending with Initiator";
      nextApproverName = InitName || "";
      summaryMessage = "Submitted for Review - Forwarded to Initiator";
    }
 
    // =====================================================
    // 4️⃣ INITIATOR → MANAGER
    // =====================================================
    else if (CHstatusselected === "4/6") {
      const manager = await getInitiatorManagerId();
 
      updateObj.Stage = 5;
      updateObj.CH_Status = "5/6";
      updateObj.NAId = manager.Id;
      updateObj.NextApproverEmpID =
        await GetApproverEmployeeId(manager.EMail);
      updateObj.Status = "Pending with Manager";
      nextApproverName = manager.Title || "";
      summaryMessage = "Submitted for Review - Forwarded to Manager";
    }
 
    // =====================================================
    // COMMON FIELDS
    // =====================================================
    updateObj.SelectedTab = "NT";
    updateObj.LastAction = new Date();
    updateObj.EmailSendFlag = 1;
 
    // =====================================================
    // SUMMARY (fetch current + append, so history isn't lost)
    // =====================================================
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: nextApproverName,
      c3: formatDateTime(new Date()),
      c4: summaryMessage || "Submitted for Review",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);
 
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);
 
    alert("Submitted successfully");
    history.push("/");
 
  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
};

const getInitiatorManagerId = async () =>{
  const spCrudOps = await SPCRUDOPS();
  const data = await spCrudOps.getRootData(
    "UserMaster",
    "*,DirectManagerName/Id,DirectManagerName/Title,DirectManagerName/EMail",
    "DirectManagerName",
    `CompanyEmailId eq '${initnameEmail}'`,
    { column: "ID", isAscending: true },
    props
  );
  return data.length ? data[0].DirectManagerName : null;
}
const GetApproverEmployeeId = async (ApproverEmail) =>{
  const spCrudOps = await SPCRUDOPS();
  const data = await spCrudOps.getRootData(
    "UserMaster",
    "EmployeeId,DirectManagerName/Id,DirectManagerName/Title",
    "DirectManagerName",
    `CompanyEmailId eq '${ApproverEmail}'`,
    { column: "ID", isAscending: true },
    props
  );
  return data.length ? data[0].EmployeeId: null;
}
const rejectbutton = async () => {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    // 🔹 Only manager can reject at closure stage
    if (CHstatusselected !== "5/6") {
      alert("Reject not allowed at this stage");
      return;
    }

    const approvers = await getCommodityApprovers(commodityselected);
    if (!approvers) {
      alert("Commodity approvers not configured");
      return;
    }

    const updateObj: any = {};

    // ============================
    // Decide who gets rework
    // ============================
    let nextApproverId: number | null = null;
    let nextApproverName = "";

   
      nextApproverId = approvers.lead.id;
      nextApproverName = "Commodity Lead";
    
    // ============================
    // Update workflow
    // ============================
    updateObj.Stage = 4;
    updateObj.CH_Status = "3/6";
    updateObj.NAId = nextApproverId;
    updateObj.NextApproverEmpID = await GetApproverEmployeeId(approvers.lead.email);
    updateObj.DAId = null;
    updateObj.Status = "Rework Requested";
    updateObj.SelectedTab = "NT";
    updateObj.LastAction = new Date();
    updateObj.EmailSendFlag = 1;

    // ============================
    // Summary entry
    // ============================
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: nextApproverName ? nextApproverName : nextApproverName,
      c3: formatDateTime(new Date()),
      c4: "Closure Rejected",
      c5: "Request sent back for rework"
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);

    alert("Closure rejected and sent for rework");
    history.push("/");

  } catch (err: any) {
    alert("Reject error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
};
async function WrongIssueAssign() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 2,
      Status: "Pending with Agency User",
      SelectedTab: "Basic",
      CH_Status: "2/6",
      LastAction: new Date(),
      EmailSendFlag: 1
    };
    // assign back to agency user
    updateObj.NAId = await getUserId(InitName);;
    updateObj.NextApproverEmpID = await GetApproverEmployeeId(initnameEmail);
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: InitName ? InitName : InitName,
      c3: formatDateTime(new Date()),
      c4: "Wrong Issue Assigned",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    //await updateRequest(reqId, updateObj, "Wrong Issue Assigned");
    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);

    alert("Returned to Initiator successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

async function handleonClickRework() {
    try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 2,
      Status: "In process - Tech",
      SelectedTab: currentTab.toUpperCase(),
      CH_Status: "2/6"
    };
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .select("Summary")();
 
    const existingSummary = item.Summary ? JSON.parse(item.Summary) : [];
    const newEntry = {
      c1: props.userDisplayName,
      c2: InitName ? InitName : InitName,
      c3: formatDateTime(new Date()),
      c4: "Sent Back to Previous Stage",
      c5: ""
    };
    existingSummary.push(newEntry);
    setJsonSummary(existingSummary);
    updateObj.Summary = JSON.stringify(existingSummary);

    // ============================
    // Persist
    // ============================
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .update(updateObj);
    //await updateRequest(reqId, updateObj, "Sent Back to Previous Stage");
    alert("Sent Back to Previous Stage successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

const buildWorkflow = () => {
  const wf: JSX.Element[] = [];

  // Initiator
  wf.push(
    <li
      key="initiator"
      className={Stage > 0 ? "beforeactiveApprover" : "activeApprover"}
      title="Initiator"
    >
      {InitName || props.userDisplayName}
    </li>
  );

  // Next Approver
  if (NextApprover) {
    wf.push(
      <li
        key="next"
        className="activeApprover"
        title="Current Approver"
      >
        {NextApprover}
      </li>
    );
  }

  return wf;
};

  return (
    <AppProvider>
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
      <div className="container p-0">
        <Header />
        <div className="displayWF">
          <ul className="main-menu mb-0">
            {buildWorkflow()}
          </ul>
        </div>
        <div style={{ backgroundColor: "#fff", paddingTop: '15px' }}>
          <ButtonBar
            buttons={buttons}
            onClose={handleClose}
            //onCreateDraft={handleCreateDraft(formState)}
            onSubmitRequest={() => handleonSubmitRequest(reqId)}
            onWithdrawn={handleonWithdrawn}
            onPrint={handleonPrint}
            onCloseIssue={handleonCloseIssue}
            onReturnBackToPITMember={handleonReturnBackToPITMember}
            onForwardToNextPITMember={handleonForwardToNextPITMember}
            onProcessWithIssueClose={handleProcessWithIssueClose}
            onProcessWithIssueOpen={handleonProcessWithIssueOpen}
            onSendBackToPreviousStage={handleonSendBackToPreviousStage}
            onAssignToChampion={handleonAssignToChampion}
            onProcessWithIssueCloseNT={handleonProcessWithIssueCloseNT}
            onBackToInitiator={handleonBackToInitiator}
            onReassignIssue={handleonReassignIssue}
            onSubmitForReview={submitforreview}
            onRejectClick={rejectbutton}
            onWrongIssueAssign={WrongIssueAssign}
            onClickRework={handleonClickRework}
            status={status}
            chStatus={ChStatus}
          />

          <RemarksModal isOpen={isRemarksOpen} onClose={closeRemarksModal} onUpdate={handleRemarksUpdate} remarksTitle={remarksTitle} />
          <AssignToAnotherUserModal
              isOpen={isAssignUserOpen}
              onClose={closeAssignUserModal}
              onSubmit={handleAssignSubmit}
              // agencies={agencies}
              context={props.currentSPContext} 
              parentProps={props} 
          />
          {/* Description section */}
          <div >
            <div id='showApprovalFlow'>

            </div>
            <div className="row marginTop10">
              <div className="col-sm-2 caption fontcolor">Problem Description 
                <i
                  className="fa fa-exclamation-circle"
                  style={{ color: "#d60000", marginLeft: "4px" }}
                />
              </div>
              <div className="col-sm-7">
                <input type="text" id="vTitle" style={{ width: '100%', marginTop: '6px' }} value={problemDescription} onChange={(e) => setProblemDescription(e.target.value)} />
              </div>
              <div className="col-sm-1 caption">Req. Date</div>
              <div className='col-sm-2'>
                <input type="date" id="vReqDate" readOnly style={{ width: '100%', marginTop: '6px' }} value={formatDateForInput(requestDate)} />
              </div>
            </div>
            <div className="row marginTop03">
              <div className="col-sm-9"></div>
              <div className="col-sm-1 caption">Req. No</div>
              <div className="col-sm-2">
                <input type="text" id="vReqNo" readOnly style={{ width: '100%' }} value={requestNumber} />
              </div>
            </div>
          </div>

          {/* Tabs */}
          <>
            <ul className="nav nav-tabs" id="myTab" style={{ display: 'flex', flexDirection: 'row', gap: '80px', margin: '15px' }}>
              {tabs.map(tab =>
                visibleTabs[tab.id] !== false && (
                  <li key={tab.id} className={activeTab === tab.id ? 'active' : ''} style={{ cursor: 'pointer' }}
                    onClick={() => { setActiveTab(tab.id); setCurrentTab(tab.id); }}>
                    <a>{tab.label}</a>
                  </li>
                )
              )}
            </ul>
          </>
        </div>
        <div className="tab-content">
          {visibleTabs["basic"] !== false && activeTab === "basic" && (
            <BaseInfoTab
              data={baseInfoData}
              onSave={handleBaseInfoSave}
              onDraft={handleBaseInfoSaveDraft}
              onFormChange={updateTabsBasedOnConditions} 
              onChange={handleBaseInfoChange}
              IPrtsProps={props} />
          )}

          {visibleTabs["technical"] !== false && activeTab === "technical" && (
            <Tab2TechnicalIssueFull props={props} reqId={reqId} activeData={activeTechData} onSave={handleTechnicalIssueSave} />
          )}

          {visibleTabs["summary"] !== false && activeTab === "summary" && (
            <Tab10Summary />
          )}
        </div>
      </div>
    )}
    </AppProvider>
  );
};

export default Approval;