
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
import { string } from 'yup';
// import Tab2TechnicalIssue from "./TechnicalIssueData";
import Tab3D1 from "./NewTabD1";
import Tab4D2 from './NewTabD2';
import Tab5D3 from './NewTabD3';
import Tab6D4 from './NewTabD4';
import Tab7D5 from './NewTabD5';
import Tab8D6 from './NewTabD6';
import Tab9D7 from './NewTabD7';
import Tab2TechnicalIssueFull from './TechnicalIssueData';
import Tab10Summary from './NewTabSummary';
import SPCRUDOPS from '../../service/DAL/spcrudops';
import IEmployeeProfileops from '../../service/BAL/SPCRUD/EmployeeProfile';
import '../../components/Pages/CSS/NewPage.scss';
import { Async, IDropdownOption } from '@fluentui/react';
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
interface D1Data {
  mD1RootCauseFound: string;
  mD1ActionStatus: string;
  mD1ICA_Details: string;
  mD1ICA_VIN: string;
  mD1PCA_Details: string;
  mD1PCA_VIN: string;
  mD1_Remarks: string;
  // add any other required properties here
}
interface D1SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D1Data;
  latestJson: any[]; 
}
interface D2Data {
  mD2RootCauseFound: string;
  mD2ActionStatus: string;
  mD2ICA_Details: string;
  mD2ICA_VIN: string;
  mD2PCA_Details: string;
  mD2PCA_VIN: string;
  mD2_Remarks: string;
  vD2AssignTo?: string;
  vD2AssignDT?: string;
  vD2ActionStatus?: string;
  mD2AnalysisDetails?: string;
  // add any other required properties here
}
interface D2SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D2Data;
  latestJson: any[]; 
}
interface D3Data {
  mD3RootCauseFound: string;
  mD3ActionStatus: string;
  mD3ICA_Details: string;
  mD3ICA_VIN: string;
  mD3PCA_Details: string;
  mD3PCA_VIN: string;
  mD3_Remarks: string;
  vD3AssignTo?: string;
  vD3AssignDT?: string;
  vD3ActionStatus?: string;
  mD3AnalysisDetails?: string;
}
interface D3SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D3Data;
  latestJson: any[]; 
}
interface D4Data {
  mD4RootCauseFound: string;
  mD4ActionStatus: string;
  mD4ICA_Details: string;
  mD4ICA_VIN: string;
  mD4PCA_Details: string;
  mD4PCA_VIN: string;
  mD4_Remarks: string;
  vD4AssignTo?: string;
  vD4AssignDT?: string;
  vD4ActionStatus?: string;
  mD4AnalysisDetails?: string;
}
interface D4SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D4Data;
  latestJson: any[]; 
}
interface D5Data {
  mD5RootCauseFound: string;
  mD5ActionStatus: string;
  mD5ICA_Details: string;
  mD5ICA_VIN: string;
  mD5PCA_Details: string;
  mD5PCA_VIN: string;
  mD5_Remarks: string;
  vD5AssignTo?: string;
  vD5AssignDT?: string;
  vD5ActionStatus?: string;
  mD5AnalysisDetails?: string;
}
interface D5SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D5Data;
  latestJson: any[]; 
}
interface D6Data {
  mD6RootCauseFound: string;
  mD6ActionStatus: string;
  mD6ICA_Details: string;
  mD6ICA_VIN: string;
  mD6PCA_Details: string;
  mD6PCA_VIN: string;
  mD6_Remarks: string;
  vD6AssignTo?: string;
  vD6AssignDT?: string;
  vD6ActionStatus?: string;
  mD6AnalysisDetails?: string;
}
interface D6SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D6Data;
  latestJson: any[]; 
}

interface D7Data {
  mD7RootCauseFound: string;
  mD7ActionStatus: string;
  mD7ICA_Details: string;
  mD7ICA_VIN: string;
  mD7PCA_Details: string;
  mD7PCA_VIN: string;
  mD7_Remarks: string;
  vD7AssignTo?: string;
  vD7AssignDT?: string;
  vD7ActionStatus?: string;
  mD7AnalysisDetails?: string;
}
interface D7SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D7Data;
  latestJson: any[]; 
}
const formatDateTime = (date: Date) => date.toISOString();

const ProblemResolutionTrackingSystem = (props: IPrtsProps) => {
  //  const params = useParams()
  const { RequestId } = useParams<{ RequestId: string }>();
  const [dataLoaded, setDataLoaded] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');
  const [ChStatus, setCHStatus] = useState('');
  const [jsonSummary, setJsonSummary] = useState<Array<{ c1: string; c2: string; c3: string; c4: string; c5: string }>>([]);
  const [isRemarksOpen, setRemarksOpen] = useState(false);
  const [remarksTitle, setRemarksTitle] = useState('');
  const [remarksFor, setRemarksFor] = useState('');
  const [isAssignUserOpen, setAssignUserOpen] = useState(false);
  const [agencies, setAgencies] = useState<{ Title: string }[]>([]);
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
  const [d1JsonArray, setD1JsonArray] = useState<any[]>([]);
  const [d2JsonArray, setD2JsonArray] = useState<any[]>([]);
  const [d3JsonArray, setD3JsonArray] = useState<any[]>([]);
  const [d4JsonArray, setD4JsonArray] = useState<any[]>([]);
  const [d5JsonArray, setD5JsonArray] = useState<any[]>([]);
  const [d6JsonArray, setD6JsonArray] = useState<any[]>([]);
  const [d7JsonArray, setD7JsonArray] = useState<any[]>([]);
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
  const [d1HistoryData, setD1HistoryData] = React.useState([]);
  const [d2HistoryData, setD2HistoryData] = React.useState([]);
  const [d3HistoryData, setD3HistoryData] = React.useState([]);
  const [d4HistoryData, setD4HistoryData] = React.useState([]);
  const [d5HistoryData, setD5HistoryData] = React.useState([]);
  const [d6HistoryData, setD6HistoryData] = React.useState([]);
  const [d7HistoryData, setD7HistoryData] = React.useState([]);

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
  const [d1Json, setD1Json] = useState<any[]>([]);
const [d2Json, setD2Json] = useState<any[]>([]);
const [d3Json, setD3Json] = useState<any[]>([]);
const [d4Json, setD4Json] = useState<any[]>([]);
const [d5Json, setD5Json] = useState<any[]>([]);
const [d6Json, setD6Json] = useState<any[]>([]);
const [d7Json, setD7Json] = useState<any[]>([]);
const [userDisplayName, setUserDisplayName] = useState(props.userDisplayName);
const [CHstatusselected, setCHStatusSelected] = useState("");
const [nonTechnicalIssueData, setNonTechnicaIssueData] = useState("");
  //---------------------------------------------------------------------------------------

  //_______________________________________________________________________________________
  // const [activeTab, setActiveTab] = React.useState('basic');
  const [d1ActiveData, setD1ActiveData] = React.useState({
    mD1RootCauseFound: '',
    mD1ActionStatus: '',
    mD1ICA_Details: '',
    mD1ICA_VIN: '',
    mD1PCA_Details: '',
    mD1PCA_VIN: '',
    mD1_Remarks: '',
    vD1AssignTo: '',
    vD1AssignDT: '',
    vD1ActionStatus: '',
  });

  const [d2ActiveData, setD2ActiveData] = React.useState({
    mD2RootCauseFound: "",
    mD2ActionStatus: "",
    mD2ICA_Details: "",
    mD2ICA_VIN: "",
    mD2PCA_Details: "",
    mD2PCA_VIN: "",
    mD2_Remarks: "",
    vD2AssignTo: "",
    vD2AssignDT: "",
    vD2ActionStatus: "",
    mD2AnalysisDetails: ""
  });

  const [d3ActiveData, setD3ActiveData] = React.useState({
    mD3RootCauseFound: '',
    mD3ActionStatus: '',
    mD3ICA_Details: '',
    mD3ICA_VIN: '',
    mD3PCA_Details: '',
    mD3PCA_VIN: '',
    mD3_Remarks: '',
    vD3AssignTo: '',
    vD3AssignDT: '',
    vD3ActionStatus: '',
  });

  const [d4ActiveData, setD4ActiveData] = React.useState({
    mD4RootCauseFound: '',
    mD4ActionStatus: '',
    mD4ICA_Details: '',
    mD4ICA_VIN: '',
    mD4PCA_Details: '',
    mD4PCA_VIN: '',
    mD4_Remarks: '',
    vD4AssignTo: '',
    vD4AssignDT: '',
    vD4ActionStatus: '',
  });

  const [d5ActiveData, setD5ActiveData] = React.useState({
    mD5RootCauseFound: '',
    mD5ActionStatus: '',
    mD5ICA_Details: '',
    mD5ICA_VIN: '',
    mD5PCA_Details: '',
    mD5PCA_VIN: '',
    mD5_Remarks: '',
    vD5AssignTo: '',
    vD5AssignDT: '',
    vD5ActionStatus: '',
  });

  const [d6ActiveData, setD6ActiveData] = React.useState({
    mD6RootCauseFound: '',
    mD6ActionStatus: '',
    mD6ICA_Details: '',
    mD6ICA_VIN: '',
    mD6PCA_Details: '',
    mD6PCA_VIN: '',
    mD6_Remarks: '',
    vD6AssignTo: '',
    vD6AssignDT: '',
    vD6ActionStatus: '',
  });
  const [d7ActiveData, setD7ActiveData] = React.useState({
    mD7RootCauseFound: '',
    mD7ActionStatus: '',
    mD7ICA_Details: '',
    mD7ICA_VIN: '',
    mD7PCA_Details: '',
    mD7PCA_VIN: '',
    mD7_Remarks: '',
    vD7AssignTo: '',
    vD7AssignDT: '',
    vD7ActionStatus: '',
  });

  const [buttons, setButtons] = useState({
    btnClose: true,
    btnCreateDraft: true,
    btnSubmit: true,
    btnWithDrawn: false,
    btnPrint: false,
    btnCloseIssue: false,
    btnReturnBackToPITMember: false,
    btnForwardToNextPITMember: false,
    btnProcessWithIssueCloseD: false,
    btnProcessWithIssueOpen: false,
    btnSendBackToPreviousStage: false,
    btnForwardAtD7: false,
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
    btnReturnBackToPITMember: false,
    btnForwardToNextPITMember: false,
    btnProcessWithIssueCloseD: false,
    btnProcessWithIssueOpen: false,
    btnSendBackToPreviousStage: false,
    btnForwardAtD7: false,
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


  const resetAllStateForNewRequest = () => {
  setRequestId(undefined);
  setRequestNumber("");
  setProblemDescription("");

  // Base Info
  setBaseInfoData({
    mIssueCategory: '',
    mRepeatedIssue: '',
    mInitDept: "",
    mInitName: userDisplayName || "",
    mIssueStatus: "Open",
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

  // Technical
  setActiveTechData({
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

  // Clear histories
  setHistoryTechData([]);
  setD1JsonArray([]);
  setD2JsonArray([]);
  setD3JsonArray([]);
  setD4JsonArray([]);
  setD5JsonArray([]);
  setD6JsonArray([]);
  setD7JsonArray([]);

  // Reset UI
  setActiveTab("basic");
  setCurrentTab("Basic");
  setStage(null);
  setCHStatus("");
  setStatus("");
};

useEffect(() => {
  // if (!RequestId) {
  //   // 🧹 NEW REQUEST → RESET EVERYTHING
  //   resetAllStateForNewRequest();
  //   setUserDisplayName(props.userDisplayName)
  //   return;
  // }

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
      
//       if(problemDescription == ""){
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
//       const fullReqNo = `PRTS/${new Date().getFullYear()} /${reqNo}`;
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
//         PurgingAttachment:cleanedForm.selectedFiles[0].name || ""
//       };

//       // Same logic as before for CH_Status and ApproverList
//       if (cleanedForm.mRootCauseFound === "Yes") {
//         updateData.CH_Status = "1/6";
//         updateData.D1_IssueData = "";
//         updateData.D2_IssueData = "";
//         updateData.D3_IssueData = "";
//         updateData.D4_IssueData = "";
//         updateData.D5_IssueData = "";
//         updateData.D6_IssueData = "";
//         updateData.D7_IssueData = "";
//         updateData.ApproverList = `${userName}`;
//         updateData.NAId = cleanedForm.mPartQualityIssue?.Id || null;
//       } else if (cleanedForm.mIs7D === "Yes") {
//         updateData.CH_Status = "1/6";
//         updateData.D1_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[0].Name, c2: formattedDate }]);
//         updateData.D2_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[1].Name , c2: formattedDate }]);
//         updateData.D3_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[2].Name, c2: formattedDate }]);
//         updateData.D4_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[3].Name, c2: formattedDate }]);
//         updateData.D5_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[4].Name, c2: formattedDate }]);
//         updateData.D6_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[5].Name, c2: formattedDate }]);
//         updateData.D7_IssueData = JSON.stringify([{ c1: cleanedForm.diamondUsers[6].Name, c2: formattedDate }]);
//         updateData.NonTechnical_IssueData = "";
//         updateData.NAId = cleanedForm.diamondUsers[0]?.Id || null;
//         let AL = "";
//         for (let i = 1; i <= 7; i++) {
//           AL += ";" + (cleanedForm.diamondUsers[i]?.Name ?? "");
//         }
//         updateData.ApproverList = AL;
//       } else if (cleanedForm.mIs7D === "No") {
//         updateData.CH_Status = "1/6";
//         updateData.NonTechnical_IssueData = finalNonTechJson;
//         updateData.D1_IssueData = "";
//         updateData.D2_IssueData = "";
//         updateData.D3_IssueData = "";
//         updateData.D4_IssueData = "";
//         updateData.D5_IssueData = "";
//         updateData.D6_IssueData = "";
//         updateData.D7_IssueData = "";
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

const handleBaseInfoSave = async (formState: any) => {
  try {
  setLoading(true);

  const cleanedForm = { ...formState };

  Object.keys(cleanedForm).forEach((key) => {
    if (cleanedForm[key] === "-1" || cleanedForm[key] === "?") {
      cleanedForm[key] = "";
    }
  });

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
      alert(`Reference Request No '${refNo}' does not exist.`);
      return;
    }
  }

  const userId =
    props.currentSPContext.pageContext.legacyPageContext.userId;

  let itemId = reqId ? Number(reqId) : null;

  // CREATE NEW DRAFT
  if (!itemId) {
    const employee = await GetUserDepartment(
      props,
      props.userEmail
    );

    updateSummary(
      props.userDisplayName,
      "",
      formatDateTime(new Date()),
      "Draft Created",
      ""
    );

    const addResult = await sp.web.lists
      .getByTitle("PRTSList")
      .items.add({
        Status: "Draft",
        Stage: 0,
        InitiatorId: userId,
        InitDepartment:
          employee?.DepartmentCode?.Department,
        InitiatorEmpId:
          props.EmployeeId[0].EmployeeID
      });

    itemId = addResult.data.Id;

    const reqNo = String(itemId).padStart(5, "0");
    const fullReqNo =
      `PRTS/${new Date().getFullYear()}/${reqNo}`;

    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(itemId)
      .update({
        ReqNo: fullReqNo
      });

    setRequestNumber(fullReqNo);
    setRequestId(String(itemId));
  }

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

    AnalysisDetails: cleanedForm.mAnalysis,
    Is7DRequired: cleanedForm.mIs7D,
    IsRootCauseFound: cleanedForm.mRootCauseFound,

    PurgingAttachment: cleanedForm.selectedFiles?.[0]?.name || "",

    SelectedTab: "Basic",

    // IMPORTANT
    Status: "Draft",
    Stage: 0
  };

  await sp.web.lists
    .getByTitle("PRTSList")
    .items.getById(itemId)
    .update(updateData);

  if (
    cleanedForm.selectedFiles &&
    cleanedForm.selectedFiles.length > 0
  ) {
    await uploadAttachments(
      itemId,
      cleanedForm.selectedFiles
    );
  }

  updateSummary(
    props.userDisplayName,
    "",
    formatDateTime(new Date()),
    "Draft Saved",
    ""
  );

  alert("Draft saved successfully");
  history.push(`/InitiatorLandingedit/${itemId}`);

  } catch (error: any) {
   console.error(error);
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
      d1: false,
      d2: false,
      d3: false,
      d4: false,
      d5: false,
      d6: false,
      d7: false,
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
    history.push('/')
  };
  async function GetUserDepartment(props, Email) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('UserMaster', '*,FullName/EMail,DepartmentCode/Department', 'FullName,DepartmentCode', `FullName/EMail eq '${Email}'`, { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.DepartmentCode, item])).values());
    const first = uniqueModels[0] as any;
    // setFormData(prev => ({ ...prev, mInitDept: first?.DepartmentCode?.Department || "" }));
    return first;
  }

  const handleCreateDraft = async () => {
    try {
      const EmployeeId = await GetUserDepartment(props, props.userEmail);
      const userDisplayName = props.userDisplayName;
      updateSummary(userDisplayName, "", formatDateTime(new Date()), "Request Created", "");
      const summary = JSON.stringify(jsonSummary);
      const userId = props.currentSPContext.pageContext.legacyPageContext.userId;
      const addResult = await sp.web.lists.getByTitle("PRTSList").items.add({ Status: "Open", Summary: summary, InitiatorId: userId, Stage: 0, InitDepartment: EmployeeId?.DepartmentCode?.Department, InitiatorEmpId: props.EmployeeId[0].EmployeeID });
      const newItemId = addResult.data.Id;
      const reqNo = newItemId.toString().padStart(5, "0");
      const fullReqNo = `PRTS/${new Date().getFullYear()}/${reqNo}`;
      setRequestNumber(fullReqNo);
      await sp.web.lists.getByTitle("PRTSList").items.getById(newItemId).update({ ReqNo: fullReqNo });
      history.push(`/InitiatorLandingedit/${newItemId}`);
    } catch (error) {
      console.error("Error creating request:", error);
      alert("Error creating request: " + error.message);
    }
  };

  // useEffect(() => {
  //   if (!RequestId) return;
  //   if (Stage === null || InitName === undefined) return;

  // // Wait until data is loaded
  // if (NextApprover === "" && DeleApprover === "" && InitName === "") {
  //   console.log("Waiting for approver data...");
  //   return;
  // }

  // const norm = (s: string) => s?.trim().toLowerCase() || "";

  // const isInitiator = norm(CurrUser) === norm(InitName);
  // const isApprover =
  //   norm(CurrUser) === norm(NextApprover) ||
  //   norm(CurrUser) === norm(DeleApprover);

  // console.log("Button Logic Triggered With:");
  // console.log({ Stage, CurrUser, NextApprover, DeleApprover, InitName, isInitiator, isApprover });


  //   const show = (list: string) => {
  //     const all = { ...defaultButtonsState };
  //     if (list !== "AllHide") {
  //       list.split(",").map(x => x.trim()).forEach(id => {
  //         if ((all as any).hasOwnProperty(id)) {
  //           (all as any)[id] = true;
  //         }
  //       });
  //     }
  //     setButtons(all);
  //   };

  //   // Stage and permission logic (kept your logic)
  //   if (Stage === 0 && isInitiator) {
  //     show("btnSubmit");
  //     return;
  //   }
  //   if (Stage === 1 && isInitiator) {
  //     if (Is7D === "No") {
  //       show("btnWithDrawn,btnPrint,btnCloseIssue,btnAssignIssueToAnoterUser");
  //     } else if (Is7D === "Yes") {
  //       show("btnWithDrawn,btnPrint,btnCloseIssue,btnReturnBackToPITMember,btnForwardToNextPITMember");
  //     } else {
  //       show("btnWithDrawn,btnPrint,btnCloseIssue,btnAssignIssueToAnoterUser");
  //     }
  //     return;
  //   }
  //   if (Stage === 2 && isInitiator) {
  //     show("btnWithDrawn,btnPrint");
  //     return;
  //   }
  //   if (Stage === 2 && isApprover) {
  //     if (Is7D === "No") {
  //       show("btnsubmitforreview,btnBackToInitiator");
  //       return;
  //     }
  //     if (Is7D === "Yes") {
  //       if (currentTab === "d7" || SelectedTab === "D7") {
  //         show("btnProcessWithIssueCloseD,btnForwardAtD7,btnSendBackToPreviousStage");
  //       } else {
  //         show("btnProcessWithIssueCloseD,btnProcessWithIssueOpen,btnSendBackToPreviousStage");
  //       }
  //       return;
  //     }
  //     // if (IsRootCauseFound === "Yes") {
  //     //   show("btnProcessWithIssueCloseNT,btnBackToInitiator");
  //     //   return;
  //     // }
  //   }
  //   if (Stage === 2 && isApprover && CHstatusselected === "2/6") {
  //   show("btnsubmitforreview,btnWrongIssueAssign"); // Submit for Review → 3/6
  //   return;
  // }
  // if (Stage === 3 && isApprover && CHstatusselected === "3/6") {
  //   show("btnsubmitforreview,btnWrongIssueAssign,btnclickRework"); // Submit for Review → 4/6 / 5/6
  //   return;
  // }
  //  if (Stage === 4 && isApprover && CHstatusselected === "4/6") {
  //   show("btnWithDrawn,btnPrint,btnsubmitforreview"); // Submit for Review → 4/6 / 5/6
  //   return;
  // }
  // if (Stage === 5 && isApprover && CHstatusselected === "5/6") {
  //   show("btnWithDrawn,btnPrint,btnCloseIssue,btnReject"); // Submit for Review → 4/6 / 5/6
  //   return;
  // }

  //   if (Stage === 3 && isInitiator) {
  //     show("btnPrint");
  //     return;
  //   }
  //   show("AllHide");
  //   // eslint-disable-next-line react-hooks/exhaustive-deps
  // }, [Stage, CurrUser, NextApprover, DeleApprover, InitName, Is7D, IsRootCauseFound, currentTab, SelectedTab,CHstatusselected]);

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
        show("btnSubmit,btnClose");
      }
      return;
    }

    // Existing Request in Draft
    if (status?.trim() === "Draft") {
      if (isInitiator) {
        show("btnSubmit,btnClose");
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

        if (isEqualToCurrentUser) {
          show("btnSubmit");
        }

        if (isApprover) {
          show("btnSubmit");
        }

        break;

      /*** 2/6* Agency User submits for review* Wrong Issue Assigned*/
      case "2/6":
        if (isApprover) {
          show("btnsubmitforreview,btnWrongIssueAssign");
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
          show("btnsubmitforreview");
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
    const hasD1 = d1JsonArray && d1JsonArray.length > 0;
    if (!hasNT && !hasD1) {
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
      const hasNT =
        activeTechData &&
        activeTechData.issueAssignTo;

      const hasD1 =
        d1JsonArray &&
        d1JsonArray.length > 0;

      if (!hasNT && !hasD1) {
        alert(
          'Please provide "Basic Information" before you submit the request.'
        );
        return;
      }
      let updateObject: any = {};
      await appendSummaryAndPersist(reqIdParam, {
        c1: props.userDisplayName,
        c2: "",
        c3: formatDateTime(new Date()),
        c4: "Request Submitted",
        c5: currentTab === "NT" ? "Technical Issue" : "7 Diamond Process"
      }, {
        Stage: 2,
        Status: "In process"
      });

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
        updateSummary(
          props.userDisplayName,
          delegated ? `${assignTo} (${delegated})` : assignTo,
          formatDateTime(new Date()),
          "Request Submitted",
          "Technical Issue"
        );
        updateObject.CH_Status = "1/6";

        updateObject.ApproverList = `${props.userDisplayName};${assignTo}`;

        const emp = await IEmployeeProfileops().getEmployeeProfile(assignTo, props);

        updateObject.NextApproverEmpID = emp?.[0]?.EmployeeID || null;

        updateObject.D1_IssueData = "";
        updateObject.D2_IssueData = "";
        updateObject.D3_IssueData = "";
        updateObject.D4_IssueData = "";
        updateObject.D5_IssueData = "";
        updateObject.D6_IssueData = "";
        updateObject.D7_IssueData = "";

      } else if (d1JsonArray && d1JsonArray.length > 0) {
        const t = [...d1JsonArray];
        const assignTo = t[0].c1;
        t[0].c2 = formatDate(new Date());
        updateObject.NAId = await getUserId(assignTo);
        const delegated = await checkDelegation(assignTo);
        updateObject.DAId = delegated ? await getUserId(delegated) : null;
        updateObject.D1_IssueData = JSON.stringify(t);
        updateObject.SelectedTab = "D1";
        updateObject.Status = "In process - D1";
        updateObject.Stage = 2;
        updateSummary(
          props.userDisplayName,
          delegated ? `${assignTo} (${delegated})` : assignTo,
          formatDateTime(new Date()),
          "Request Submitted",
          "7 Diamond Process Required"
        );
      }
      // Always update summary, LastAction, EmailSendFlag, ReqNo
      updateObject.Summary = JSON.stringify(jsonSummary);
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

  const CheckValidProcessWithIssueClose = (container: string): boolean => {
    let t: any = null;
    if (container === "NT") {
      if (!activeTechData) {
        alert("Missing Data for NT");
        return false;
      }
      t = activeTechData;
      if (t.mNTRootCauseFound !== "Yes") {
        setActiveTab("technical");
        alert('Change [Is Root Cause Found] = "Yes".\nIn Technical Tab');
        return false;
      }
      return true;
    }
    const jsonMap: any = {
      D1: d1JsonArray,
      D2: d2JsonArray,
      D3: d3JsonArray,
      D4: d4JsonArray,
      D5: d5JsonArray,
      D6: d6JsonArray,
      D7: d7JsonArray
    };
    const arr = jsonMap[container];
    if (!arr || arr.length === 0) {
      alert("Missing data for " + container);
      return false;
    }
    t = arr[arr.length - 1];
    if (t.c4 !== "Yes") {
      setActiveTab(container.toLowerCase());
      alert(`Change [Is Root Cause Found] = "Yes".\nIn ${container} Tab`);
      return false;
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

      updateSummary(
        props.userDisplayName,
        delegated ? `${initiatorName} (${delegated})` : initiatorName,
        formatDateTime(new Date()),
        "Issue Resolved",
        "Sent To Initiator for Review and Close"
      );

      updateObj.Summary = JSON.stringify(jsonSummary);
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

    updateSummary(
      props.userDisplayName,
      delegatedName ? `${nextApproverName} (${delegatedName})` : nextApproverName,
      formatDateTime(new Date()),
      "Open Issue Forwarded",
      ""
    );

    updateObj.Summary = JSON.stringify(jsonSummary);

    await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update(updateObj);

    alert("Forwarded Successfully!");
    history.push("/");

  } catch (e: any) {
    alert("Error: " + (e?.message || e));
  } finally {
    setLoading(false);
  }
};
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
// function handleD1Update(result: DxSaveResult) {
//   setD1Json(result.latestJson);   // ⭐ NEW
// }

// function handleD2Update(result: DxSaveResult) { setD2Json(result.latestJson); }
// function handleD3Update(result: DxSaveResult) { setD3Json(result.latestJson); }
// function handleD4Update(result: DxSaveResult) { setD4Json(result.latestJson); }
// function handleD5Update(result: DxSaveResult) { setD5Json(result.latestJson); }
// function handleD6Update(result: DxSaveResult) { setD6Json(result.latestJson); }
// function handleD7Update(result: DxSaveResult) { setD7Json(result.latestJson); }

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

      updateSummary(
        props.userDisplayName,
        delegated ? `${initiatorName} (${delegated})` : initiatorName,
        formatDateTime(new Date()),
        "Issue Resolved",
        "Sent To Initiator for Review and Close"
      );

      updateObj.Summary = JSON.stringify(jsonSummary);
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
      case 'ReturnBackToPITMember':
        returnBackToPITMember();
        break;
      case 'SendBackToPreviousStage':
        sendBackToPreviousStage();
        break;
      case 'ForwardToNextPITMember':
        forwardToNextPITMember();
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
      const siteUrl = props.context.pageContext._site.absoluteUrl;
      const requestNumber = parseInt(reqIdParam);
      const endpoint = `${siteUrl}/PRTS/_api/web/lists/getbytitle('PRTSList')/items(${requestNumber})`;
      const digest = await getDigest();
      const nextJsonSummary = [
        ...jsonSummary,
        {
          c1: props.userDisplayName,
          c2: '',
          c3: formatDateTime(new Date()),
          c4: 'Issue Closed',
          c5: remarks,
        },
      ];
      const updateData = {
        __metadata: { type: 'SP.Data.PRTSListListItem' },
        NAId: null,
        DAId: null,
        Stage: 3,
        Status: 'Close',
        CH_Status: '6/6',
        IssueResolveRequestBy: '',
        LastAction: new Date(),
        EmailSendFlag: 1,
        Summary: JSON.stringify(nextJsonSummary),
      };
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
        },
        body: JSON.stringify(updateData),
      });
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error.message.value);
      }
      await appendSummaryAndPersist(reqIdParam!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Issue Closed",
  c5: remarks
}, {
  Stage: 3,
  Status: "Close",
  CH_Status: "6/6",
  NAId: null,
  DAId: null
});

      console.log('Data Saved.');
      history.push('/');
    } catch (error: any) {
      alert('Error saving issue closing: ' + (error?.message || error) + '. Refresh and try again, or contact Administrator.');
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
    // 3️⃣ Create NonTechnical JSON update (same as Angular)
    // -------------------------
    const today = formatDate(new Date());

    // Load previous NT data (historyTechData)
    const oldNT = Array.isArray(historyTechData) ? [...historyTechData] : [];

    // Add new row (c1 = agency, c2 = user, c3 = date)
    const newRow = {
      c1: agency,
      c2: user,
      c3: today
    };

    const updatedNT = [...oldNT, newRow];

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
    // 6️⃣ Add Summary entry
    // -------------------------
    updateSummary(
      props.userDisplayName,
      delegatedName ? `${user} (${delegatedName})` : user,
      formatDateTime(new Date()),
      "Re-Assign",
      remarks
    );

    updateObj.Summary = JSON.stringify(jsonSummary);

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

  const applyStagePermissions = (StageParam: any, CurrUserParam: any, NextApproverParam: any, DeleApproverParam: any, InitNameParam: any, TabSelected: any) => {
    const StatusTab = ["basic", "technical", "d1", "d2", "d3", "d4", "d5", "d6", "d7"];
    const selectedIndex = StatusTab.indexOf((TabSelected || "basic").toLowerCase());
    let visible: any = {};
    let editable: any = {};
    StatusTab.forEach(t => { visible[t] = true; editable[t] = false; });
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
        .select("Title", "Severity", "RefReqNo", "Commodity", "SupplierSource", "InitDepartment", "IsRootCauseFound", "Is7DRequired", "AnalysisDetails", "IssueStatus", "NonTechnical_IssueData", "Initiator/Title","Initiator/EMail", "Initiator/Id", "Stage", "NA/Id", "NA/Title", "DA/Id", "DA/Title", "SelectedTab", "ReqNo", "D1_IssueData", "CH_Status", "Status")
        .expand("Initiator", "NA", "DA")
        .get();

      const StageVal = item.Stage;
      const CurrUserVal = props.userDisplayName;
      const NextApproverVal = item?.NA?.Title || "";
      const DeleApproverVal = item?.DA?.Title || "";
      const InitNameVal = item?.Initiator?.Title || "";
      const TabSelected = item.SelectedTab;
      const d1RawJson = item.D1_IssueData;
      let parsedD1: any[] = [];
      try { if (d1RawJson) parsedD1 = JSON.parse(d1RawJson); } catch (e) { console.error("Invalid JSON in D1_IssueData:", e); }
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
      setD1JsonArray(parsedD1);
      setNonTechnicaIssueData(item.NonTechnical_IssueData);
      applyStagePermissions(StageVal, CurrUserVal, NextApproverVal, DeleApproverVal, InitNameVal, TabSelected);
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
    { id: 'd1', label: 'D1' },
    { id: 'd2', label: 'D2' },
    { id: 'd3', label: 'D3' },
    { id: 'd4', label: 'D4' },
    { id: 'd5', label: 'D5' },
    { id: 'd6', label: 'D6' },
    { id: 'd7', label: 'D7' },
    { id: 'summary', label: 'SUMMARY' },
  ];

  function handleTechnicalIssueSave(data: TechIssueData, updatedJson?: any[]): void {
    setActiveTechData(data);
    if (updatedJson) setHistoryTechData(updatedJson);
  }

  // D1..D7 handlers unchanged except using reqId from state
  async function handleD1Update(result: D1SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD1JsonArray(result.updatedArray);
      setD1Json(result.updatedArray);
      setD1ActiveData(prev => ({ ...prev, ...result.savedFields }));
      if (reqId) {
        await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update({ D1_IssueData: result.jsonString, CH_Status: result.chStatus });
      } else { console.warn("reqId not set — skipping D1 SP update."); }
      await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D1",
  c5: result.latestJson[0].c9 || ""
});

      alert("D1 updated successfully");
    } catch (err: any) {
      alert("Failed to update D1: " + (err?.message || err));
      console.error(err);
    } finally {
      setLoading(false);
    }
  }

  async function handleD2Update(result: D2SaveResult): Promise<void> {
    try {
      setLoading(true);
      // update local JSON array state
      setD2JsonArray(result.updatedArray);
      setD2Json(result.updatedArray); 
      // merge saved editable fields with existing readonly fields in parent
      setD2ActiveData(prev => ({
        ...prev,
        ...result.savedFields, // updates mD2* fields but keeps vD2* fields
      }));

      // persist to SharePoint
      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D2_IssueData: result.jsonString,
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D2 SP update.");
      }
        await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D2",
  c5: result.latestJson[0].c9 || ""
});
      alert("D2 updated successfully");
    } catch (err: any) {
      console.error("Error saving D2:", err);
      alert("Failed to update D2: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  }

  async function handleD3Update(result: D3SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD3JsonArray(result.updatedArray);
 setD3Json(result.latestJson); 
      setD3ActiveData(prev => ({
        ...prev,
        ...result.savedFields,
      }));

      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D3_IssueData: result.jsonString,   // ✅ Correct field
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D3 SP update.");
      }
  await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D3",
  c5: result.latestJson[0].c9 || ""
});
      alert("D3 updated successfully");   // ✅ Correct alert
    } catch (err: any) {
      console.error("Error saving D3:", err);
      alert("Failed to update D3: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  }

  async function handleD4Update(result: D4SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD4JsonArray(result.updatedArray);
      setD4Json(result.latestJson);
      setD4ActiveData(prev => ({
        ...prev,
        ...result.savedFields,
      }));

      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D4_IssueData: result.jsonString,   // ✅ Correct field
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D4 SP update.");
      }
  await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D4",
  c5: result.latestJson[0].c9 || ""
});
      alert("D4 updated successfully");   // ✅ Correct alert
    } catch (err: any) {
      console.error("Error saving D4:", err);
      alert("Failed to update D4: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  }

  async function handleD5Update(result: D5SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD5JsonArray(result.updatedArray);
      setD5Json(result.latestJson);
      setD5ActiveData(prev => ({
        ...prev,
        ...result.savedFields,
      }));

      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D5_IssueData: result.jsonString,   // ✅ Correct field
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D5 SP update.");
      }
  await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D5",
  c5: result.latestJson[0].c9 || ""
});
      alert("D5 updated successfully");   // ✅ Correct alert
    } catch (err: any) {
      console.error("Error saving D5:", err);
      alert("Failed to update D5: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  }
  async function handleD6Update(result: D6SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD6JsonArray(result.updatedArray);
      setD6Json(result.latestJson); 
      setD6ActiveData(prev => ({
        ...prev,
        ...result.savedFields,
      }));

      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D6_IssueData: result.jsonString,   // ✅ Correct field
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D6 SP update.");
      }
  await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D6",
  c5: result.latestJson[0].c9 || ""
});
      alert("D6 updated successfully");   // ✅ Correct alert
    } catch (err: any) {
      console.error("Error saving D6:", err);
      alert("Failed to update D6: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
  }
  async function handleD7Update(result: D7SaveResult): Promise<void> {
    try {
      setLoading(true);
      setD7JsonArray(result.updatedArray);
      setD7Json(result.latestJson); 
      setD7ActiveData(prev => ({
        ...prev,
        ...result.savedFields,
      }));

      if (reqId) {
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            D7_IssueData: result.jsonString,   // ✅ Correct field
            CH_Status: result.chStatus,
          });
      } else {
        console.warn("reqId not set — skipping D7 SP update.");
      }
        await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Updated D7",
  c5: result.latestJson[0].c9 || ""
});

      alert("D7 updated successfully");   // ✅ Correct alert
    } catch (err: any) {
      console.error("Error saving D7:", err);
      alert("Failed to update D7: " + (err?.message || err));
    } finally {
      setLoading(false);
    }
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
        d1: false,
        d2: false,
        d3: false,
        d4: false,
        d5: false,
        d6: false,
        d7: false,
        summary: true
      });
    }

    // Rule 3: Root Cause Found Yes → Only Basic + Technical + Summary
    if (rootCauseFound === "Yes" && !isEqualToCurrentUser) {
      setVisibleTabs({
        basic: true,
        technical: false,
        d1: false,
        d2: false,
        d3: false,
        d4: false,
        d5: false,
        d6: false,
        d7: false,
        summary: true
      });
      return;
    }

    // Rule 1: 7D = Yes → Show D1–D7 only
    if (is7D === "Yes") {
      setVisibleTabs({
        basic: true,
        technical: false,
        d1: true,
        d2: true,
        d3: true,
        d4: true,
        d5: true,
        d6: true,
        d7: true,
        summary: true
      });
      return;
    }
  };
async function returnBackToPITMember() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const updateObj: any = {
      Stage: 1,
      Status: "Returned to PIT Member",
      SelectedTab: "D1",
      CH_Status: "2/6"
    };

    await updateRequest(reqId, updateObj, "Returned Back to PIT Member");
    alert("Returned Back to PIT Member successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

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

    await updateRequest(reqId, updateObj, "Sent Back to Previous Stage");
    alert("Sent Back to Previous Stage successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

async function forwardToNextPITMember() {
  try {
    setLoading(true);
    if (!reqId) {
      alert("Request ID missing");
      return;
    }

    const nextStageTabMap: any = {
      d1: "D2",
      d2: "D3",
      d3: "D4",
      d4: "D5",
      d5: "D6",
      d6: "D7",
      d7: "D7"
    };

    const nextTab = nextStageTabMap[currentTab] || "D2";

    const updateObj: any = {
      Stage: 2,
      Status: `In process - ${nextTab}`,
      SelectedTab: nextTab,
      CH_Status: "3/6"
    };

    await updateRequest(reqId, updateObj, "Forwarded To Next PIT Member");
    await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Forwarded to Next PIT Member",
  c5: ""
}, {
  SelectedTab: nextTab,
  Stage: 2,
  Status: `In process - ${nextTab}`
});

    alert("Forwarded to Next PIT Member successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}

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
      CH_Status: "0/6"
    };
await appendSummaryAndPersist(reqId!, {
  c1: props.userDisplayName,
  c2: "",
  c3: formatDateTime(new Date()),
  c4: "Request Withdrawn",
  c5: ""}, {
  Stage: 3,
  Status: "Withdrawn",
  CH_Status: "0/6"
});

    await updateRequest(reqId, updateObj, "Request Withdrawn");
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

    await updateRequest(reqId, updateObj, "Sent Back to Initiator");
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
    }

    // =====================================================
    // COMMON FIELDS
    // =====================================================
    updateObj.SelectedTab = "NT";
    updateObj.LastAction = new Date();
    updateObj.EmailSendFlag = 1;

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
    updateSummary(
      props.userDisplayName,
      nextApproverName,
      formatDateTime(new Date()),
      "Closure Rejected",
      "Request sent back for rework"
    );

    updateObj.Summary = JSON.stringify(jsonSummary);

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
    // updateObj.NAId = agencyUserId;
    // updateObj.NextApproverEmpID = agencyUserEmpId;

    await updateRequest(reqId, updateObj, "Wrong Issue Assigned");
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

    await updateRequest(reqId, updateObj, "Sent Back to Previous Stage");
    alert("Sent Back to Previous Stage successfully");
    history.push("/");

  } catch (err: any) {
    alert("Error: " + (err?.message || err));
  } finally {
    setLoading(false);
  }
}
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
        <div style={{ backgroundColor: "#fff", paddingTop: '15px' }}>
          <ButtonBar
            buttons={buttons}
            onClose={handleClose}
            // onCreateDraft={handleCreateDraft}
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
              <div className="col-sm-1 caption">Problem Description</div>
              <div className="col-sm-8">
                <input type="text" id="vTitle" style={{ width: '100%' }} value={problemDescription} onChange={(e) => setProblemDescription(e.target.value)} />
              </div>
              <div className="col-sm-1 caption">Req. Date</div>
              <div className='col-sm-2'>
                <input type="date" id="vReqDate" readOnly style={{ width: '100%' }} value={formatDateForInput(requestDate)} />
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
              onFormChange={updateTabsBasedOnConditions} 
              IPrtsProps={props} />
          )}

          {visibleTabs["technical"] !== false && activeTab === "technical" && (
            <Tab2TechnicalIssueFull reqId={reqId} activeData={activeTechData} onSave={handleTechnicalIssueSave} />
          )}

          {visibleTabs["d1"] !== false && activeTab === "d1" && (
            <Tab3D1 historyData={d1HistoryData} onSave={handleD1Update} />
          )}

          {visibleTabs["d2"] !== false && activeTab === "d2" && (
            <Tab4D2 activeData={d2ActiveData} existingJsonArray={d2JsonArray} onSave={handleD2Update} />
          )}

          {visibleTabs["d3"] !== false && activeTab === "d3" && (
            <Tab5D3 activeData={d3ActiveData} historyData={d3HistoryData} existingJsonArray={d3JsonArray} onSave={handleD3Update} />
          )}

          {visibleTabs["d4"] !== false && activeTab === "d4" && (
            <Tab6D4 activeData={d4ActiveData} historyData={d4HistoryData} onSave={handleD4Update} />
          )}

          {visibleTabs["d5"] !== false && activeTab === "d5" && (
            <Tab7D5 activeData={d5ActiveData} historyData={d5HistoryData} existingJsonArray={d5JsonArray} onSave={handleD5Update} />
          )}

          {visibleTabs["d6"] !== false && activeTab === "d6" && (
            <Tab8D6 activeData={d6ActiveData} historyData={d6HistoryData} existingJsonArray={d6JsonArray} onSave={handleD6Update} />
          )}

          {visibleTabs["d7"] !== false && activeTab === "d7" && (
            <Tab9D7 activeData={d7ActiveData} historyData={d7HistoryData} onSave={handleD7Update} />
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

export default ProblemResolutionTrackingSystem;

/* Helper no-op implementations for functions you left unimplemented.
   Replace with your real implementations or integrate them with your codebase.
*/
function findDelegation(singleUser: any, delegatedApprovers: any) {
  throw new Error('Function not implemented.');
}

