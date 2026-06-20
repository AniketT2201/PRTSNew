import * as React from 'react';
import { useState, useEffect } from "react";
import { IPrtsProps } from '../IPrtsProps';
import { PRTSService } from '../../service/BAL/SPCRUD/PRTSService';
import "../Pages/CSS/BaseInfoTab.scss";
import SPCRUDOPS from '../../service/DAL/spcrudops';
import { IDropdownOption } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp/presets/all';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useParams } from 'react-router-dom';

interface DiamondUser {
  Id: number;
  Name: string;
  Email?: string;
}

interface BaseInfoData {
  mIssueCategory?: string | number | readonly string[];
  mRepeatedIssue?: string | number | readonly string[];
  mInitDept?: string;
  mInitName?: string;
  mInitEmail?: string; // added per request
  mIssueStatus?: string;
  mIs7D?: string;
  mRootCauseFound?: string;
  mAnalysis?: string;
  mPurgingAttachment?: React.ReactNode;
  mIssueAttachment?: React.ReactNode;
  mTitle?: string;
  mPartName?: string;
  mPartNo?: string;
  mPartSupplier?: string;
  mPRTSSource?: string;
  mProjectCode?: string;
  mIssueVINNo?: string;
  mMFGShop?: string;
  mIssueDescription?: string;
  mCategory?: string;
  mSeverity?: string;
  mQtyAffected?: string;
  mVariantAffected?: string;
  mEngineType?: string;
  mIsRepeated?: string;
  mRefNo?: string;
  mCommodity?: string;
  mBuildType?: string;
  mAgency?: string;
  mPartQualityIssue?: any; // PeoplePicker result stored as { Id, LoginName, Name }
  mPartSupplierSource?: string;

  // internal (optional)
  diamondUsers?: DiamondUser[];
}

interface BaseInfoProps {
  data: BaseInfoData;
  onSave: (data: BaseInfoData) => void;
  onFormChange: (rootCause: string, is7D: string) => void;  // <--- ADD THIS
  IPrtsProps: IPrtsProps;
}

const BaseInfoTab = React.forwardRef((props: BaseInfoProps, ref) => {
  const { data, onSave, onFormChange, IPrtsProps } = props;
  const { RequestId } = useParams<{ RequestId: string }>();
  const [formData, setFormData] = useState<BaseInfoData>(data || {});
  const [modeldata, setModelData] = useState<any[]>([]);
  const [mfgShopOptions, setMfgShopOptions] = useState<any[]>([]);
  const [currentUserEmail, setCurrentUserEmail] = useState<string>("");
  const [engineTypeOptions, setEngineTypeOptions] = useState<any[]>([]);
  const [commodityOption, setCommodityOption] = useState<any[]>([]);
  const [varientDataOptions, setVarientDataOptions] = useState<any[]>([]);
  const [agencyOptions, setAgencyOptions] = useState<any[]>([]);
  const [prtsResourceOptions, setprtsResourceOptions] = useState<any[]>([]);
  const [buildTypeOptions, setBuildTypeOptions] = useState<any[]>([]);
  const [issueCategoryOption, setIssueCategoryOption] = useState<any[]>([]);
  const [reqId, setRequestId] = useState<string | undefined>(undefined);
  const [diamondUsers, setDiamondUsers] = useState<DiamondUser[]>([]);
  const [loadingSave, setLoadingSave] = useState(false);

  const [isCreated, setIsCreated] = useState<boolean>(false);
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedIssueFiles, setSelectedIssueFiles] = useState<File[]>([]);

  const [isApprover, setIsApprover] = useState<boolean>(false);
const [isEditMode, setIsEditMode] = useState<boolean>(false);
const [attachmentMatrix, setAttachmentMatrix] = useState({
  sosjes: "No",
  control: "No",
  pfmea: "No",
  kaizen: "No",
  qualityAlert: "No"
});

const [matrixFiles, setMatrixFiles] = useState<{
  sosjes?: File[];
  control?: File[];
  pfmea?: File[];
  kaizen?: File[];
  qualityAlert?: File[];
}>({});
useEffect(() => {
  if (!RequestId) {
              setIsCreated(false);

  }
}, [RequestId]);


  useEffect(() => { setFormData(prev => ({ ...prev, ...data })); }, [data]);



  useEffect(() => {
    GetIssueCategory(IPrtsProps);
    GetPRTSResource(IPrtsProps);
    GetBuildType(IPrtsProps);
    GetModels(IPrtsProps);
    GetMfgShopOptions(IPrtsProps);
    GetCommodityList(IPrtsProps);
    GetAgencyData(IPrtsProps);

    if (RequestId) {
      (async () => {
        const d = await handleBaseInfoGet(RequestId);
        if (d) setFormData(prev => ({ ...prev, ...d }));
      })();
    }

    PRTSService.GetSPListItems(props.IPrtsProps.currentSPContext.pageContext.web.absoluteUrl, "", "ProgramCode");
  }, []);

  useEffect(() => {
    setFormData(prev => ({ ...prev, mIssueStatus: prev.mIssueStatus || "Open" }));

    const loadInitiator = async () => {
      const user = IPrtsProps.context.pageContext.user;
      setFormData(prev => ({ ...prev, mInitName: user.displayName, mInitEmail: user.email }));
      setCurrentUserEmail(user.email);
      GetUserDepartment(IPrtsProps, user.email).catch(console.error);
    };

    loadInitiator();
  }, []);
  // When criteria met, load diamond users
  useEffect(() => {
    const shouldFetch =
      formData.mRootCauseFound === "No" &&
      formData.mIs7D === "Yes" &&
      !!formData.mMFGShop &&
      formData.mMFGShop !== "Select";

    if (!shouldFetch) {
      setDiamondUsers([]);
      setFormData(prev => ({ ...prev, diamondUsers: [] }));
      return;
    }

    (async () => {
      const users = await loadDiamondUsers(formData.mMFGShop || "");
      setDiamondUsers(users);
      setFormData(prev => ({ ...prev, diamondUsers: users }));

      // send to parent (explicitly include diamondUsers)
      // onSave({ ...formData, diamondUsers: users });
    })();
  }, [formData.mRootCauseFound, formData.mIs7D, formData.mMFGShop]);

  const handleChange = (event: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { id, value } = event.target;
    setFormData(prev => ({ ...prev, [id]: value }));
  };

  const validateFields = () => {
    const requiredFields = [
      { key: "mIssueCategory", label: "Issue Category" },
      { key: "mRepeatedIssue", label: "Repeated Issue" },
      { key: "mInitDept", label: "Initiator Department" },
      { key: "mInitName", label: "Initiator Name" },
      { key: "mInitEmail", label: "Initiator Email" },
      { key: "mIssueStatus", label: "Issue Status" },
      { key: "mRootCauseFound", label: "Root Cause Found" },
      { key: "mAnalysis", label: "Analysis" },
      { key: "mPartName", label: "Part Name" },
      { key: "mPartNo", label: "Part No" },
      { key: "mPartSupplier", label: "Part Supplier" },
      { key: "mPRTSSource", label: "PRTS Source" },
      { key: "mProjectCode", label: "Project Code" },
      { key: "mIssueVINNo", label: "Issue VIN No" },
      { key: "mMFGShop", label: "MFG Shop" },
      { key: "mIssueDescription", label: "Issue Description" },
      // { key: "mCategory", label: "Category" },
      { key: "mSeverity", label: "Severity" },
      { key: "mQtyAffected", label: "Qty Affected" },
      { key: "mVariantAffected", label: "Variant Affected" },
      { key: "mEngineType", label: "Engine Type" },
      // { key: "mIsRepeated", label: "Is Repeated" },
      // { key: "mRefNo", label: "Reference No" },
      { key: "mCommodity", label: "Commodity" },
      { key: "mBuildType", label: "Build Type" },
      // { key: "mPartSupplierSource", label: "Part Supplier Source" },
    ];

    for (const field of requiredFields) {
      if (!formData[field.key as keyof typeof formData]) {
        alert(`${field.label} is required`);
        return false;
      }
    }
    // conditional validations
    // Case 1: if Root Cause Found = Yes then Agency + Part Quality Issue required
    if (formData.mRootCauseFound === "Yes") {
      if (!formData.mAgency) {
        alert("Agency is required");
        return false;
      }
      if (!formData.mPartQualityIssue) {
        alert("Part Quality Issue is required");
        return false;
      }
    }

    // Case 2: if Root Cause Found = No then Is7D required
    if (formData.mRootCauseFound === "No") {
      if (!formData.mIs7D) {
        alert("Is 7D is required");
        return false;
      }

      // Case 3: if Root Cause Found = No and Is7D = No then Agency + Part Quality Issue required
      if (formData.mIs7D === "No") {
        if (!formData.mAgency) {
          alert("Agency is required");
          return false;
        }
        if (!formData.mPartQualityIssue) {
          alert("Part Quality Issue is required");
          return false;
        }
      }
    }
    return true;
  };

  const handleSave = async () => {
    if (!validateFields()) return;
    try {
      setLoadingSave(true);

      // push basic form to parent first
      // onSave(formData);

      // if (!RequestId) {
      //   alert('RequestId missing - save to SP skipped.');
      //   return;
      // }

      // Decide workflow based on mIs7D
      const itemUpdate: any = {};
      const formattedDate = formatDate(new Date());

      if (formData.mIs7D === 'Yes') {
        // ensure diamondUsers loaded
        const users = diamondUsers.length ? diamondUsers : await loadDiamondUsers(formData.mMFGShop || "");

        // map to D1..D7 values (use Name)
        const userNames = users.map(u => u.Name || "");

        itemUpdate.D1_IssueData = JSON.stringify([{ c1: userNames[0] || "", c2: formattedDate }]);
        itemUpdate.D2_IssueData = JSON.stringify([{ c1: userNames[1] || "", c2: formattedDate }]);
        itemUpdate.D3_IssueData = JSON.stringify([{ c1: userNames[2] || "", c2: formattedDate }]);
        itemUpdate.D4_IssueData = JSON.stringify([{ c1: userNames[3] || "", c2: formattedDate }]);
        itemUpdate.D5_IssueData = JSON.stringify([{ c1: userNames[4] || "", c2: formattedDate }]);
        itemUpdate.D6_IssueData = JSON.stringify([{ c1: userNames[5] || "", c2: formattedDate }]);
        itemUpdate.D7_IssueData = JSON.stringify([{ c1: userNames[6] || "", c2: formattedDate }]);
        // clear NonTechnical_IssueData
        itemUpdate.NonTechnical_IssueData = "";

        // Build ApproverList: initiator;user1;user2;...
        const AL = [formData.mInitName || ""].concat(userNames.filter(Boolean)).join(';');
        itemUpdate.ApproverList = AL;

        // optional: set CH_Status
        itemUpdate.CH_Status = '1/6';

      } else {
        // Non-technical path (mIs7D === 'No')
        // Expect PeoplePicker stored in mPartQualityIssue (Id/Name/LoginName)
        const user = formData.mPartQualityIssue;
        const agency = formData.mAgency || "";
        const userDisplay = (user && (user.Name || user.DisplayText || user.Title)) || "";

        const nt = [{ c1: agency, c2: userDisplay, c3: formattedDate }];
        itemUpdate.NonTechnical_IssueData = JSON.stringify(nt);

        // clear D1..D7
        itemUpdate.D1_IssueData = "";
        itemUpdate.D2_IssueData = "";
        itemUpdate.D3_IssueData = "";
        itemUpdate.D4_IssueData = "";
        itemUpdate.D5_IssueData = "";
        itemUpdate.D6_IssueData = "";
        itemUpdate.D7_IssueData = "";

        // Approver list: initiator;selectedUser
        const AL = [formData.mInitName || "", userDisplay].filter(Boolean).join(';');
        itemUpdate.ApproverList = AL;

        itemUpdate.CH_Status = '2/6';
      }

      // Persist to SharePoint
      // await sp.web.lists.getByTitle('PRTSList').items.getById(Number(RequestId)).update(itemUpdate);

      // update local state and notify parent
      const updatedForm = { ...formData, selectedFiles, selectedIssueFiles };
      onSave(updatedForm);

      // alert('Saved successfully.');
    } catch (err: any) {
      console.error('Save error', err);
      alert('Error while saving: ' + (err?.message || err));
    } finally {
      setLoadingSave(false);
    }
  };

  const getUserDetailsByName = async (displayName: string) => {
    const results = await sp.web.siteUsers.filter(`Title eq '${displayName}'`).get();
    if (results.length > 0) return results[0];
    return null;
  };

  useEffect(() => {
  if (RequestId && formData.mPurgingAttachment !== undefined) {
    console.log("URL:", window.location.href);
    console.log("RequestId:", RequestId);
    loadAttachments(Number(RequestId));
  }
}, [RequestId, formData.mPurgingAttachment]);


const loadAttachments = async (itemId: number) => {
  try {
    const files = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(itemId)
      .attachmentFiles();

    setAttachments(files);
  } catch (error) {
    console.error("Error loading attachments", error);
  }
};


const deleteAttachment = async (fileName: string) => {
  if (!RequestId) return;

  const confirmDelete = confirm(`Delete attachment "${fileName}"?`);
  if (!confirmDelete) return;

  await sp.web.lists
    .getByTitle("PRTSList")
    .items.getById(Number(RequestId))
    .attachmentFiles.getByName(fileName)
    .delete();

  await loadAttachments(Number(RequestId));
};

  //------------- Get Data -----------------------------
  
  const handleBaseInfoGet = async (reqId: string): Promise<BaseInfoData | null> => {
    try {
      const item = await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).select(
        "Title",
        "PartName",
        "PartNumbe",
        "Status",
        "SupplierName",
        "PRTSSource",
        "ProjectCode",
        "BuildType",
        "VINNo",
        "MFGShopSelection",
        "IssueDescription",
        "IssueCategory",
        "Severity",
        "QtyAffected",
        "VariantAffected",
        "EngineType",
        "RepeatedIssue",
        "RefReqNo",
        "Commodity",
        "SupplierSource",
        "InitDepartment",
        "IsRootCauseFound",
        "Is7DRequired",
        "AnalysisDetails",
        "IssueStatus",
        "NonTechnical_IssueData",
        "Initiator/Title",
        "Initiator/Id",
        "PurgingAttachment",
        "SOSJESValue",
        "ControlPlanValue",
        "PFMEAValue",
        "KaizenValue",
        "QualityAlertValue"
      ).expand("Initiator").get();
      if (item.Status !== "Draft") {
         setIsCreated(true)
      }
      let userName = '';
      let agencyName = '';
      try {
        const agencyDetails = JSON.parse(item.NonTechnical_IssueData || '[]');
        if (agencyDetails && agencyDetails.length > 0) agencyName = agencyDetails[0].c1 || '';
        if (agencyDetails && agencyDetails.length > 0) userName = agencyDetails[0].c2 || '';
      } catch {
        agencyName = '';
        userName = '';
      }

      const getuserdata = getUserDetailsByName(userName);
      const getuserEmail = (await getuserdata)?.Email || '';

      const formState: BaseInfoData = {
        mTitle: item.Title || "",
        mPartName: item.PartName || "",
        mPartNo: item.PartNumbe || "",
        mPartSupplier: item.SupplierName || "",
        mPRTSSource: item.PRTSSource || "",
        mProjectCode: item.ProjectCode || "",
        mBuildType: item.BuildType || "",
        mIssueVINNo: item.VINNo || "",
        mMFGShop: item.MFGShopSelection || "",
        mIssueDescription: item.IssueDescription || "",
        mIssueCategory: item.IssueCategory || "",
        mSeverity: item.Severity || "",
        mQtyAffected: item.QtyAffected || "",
        mVariantAffected: item.VariantAffected || "",
        mEngineType: item.EngineType || "",
        mIsRepeated: item.RepeatedIssue || "",
        mRefNo: item.RefReqNo || "",
        mCommodity: item.Commodity || "",
        mPartSupplierSource: item.SupplierSource || "",
        mInitDept: item.InitDepartment || "",
        mAgency: agencyName || "",
        mPartQualityIssue: getuserEmail || "",
        mRootCauseFound: item.IsRootCauseFound || "Select",
        mIs7D: item.Is7DRequired || "Select",
        mAnalysis: item.AnalysisDetails || "",
        mInitName: item.Initiator?.Title || "",
        mIssueStatus: item.IssueStatus || "Open",
        mPurgingAttachment: item.PurgingAttachment || ""
      };
setAttachmentMatrix(prev => ({
  ...prev,
  sosjes: item.SOSJESValue || "No"
}));
setAttachmentMatrix(prev => ({
  ...prev,
  control: item.ControlPlanValue || "No"
}));
setAttachmentMatrix(prev => ({
  ...prev,
  pfmea: item.PFMEAValue || "No"
}));
setAttachmentMatrix(prev => ({
  ...prev,
  kaizen: item.KaizenValue || "No"
}));
setAttachmentMatrix(prev => ({
  ...prev,
  qualityAlert: item.QualityAlertValue || "No"
}));
      GetEngineTypedata(IPrtsProps, item.ProjectCode);
      return formState;

    } catch (error) {
      console.log("GET ERROR:", error);
      alert("Error loading data: " + (error.message ? error.message : error));
      return null;
    }
  };
//-------Get Attachments by Prefix -----
const getAttachmentsByPrefix = (prefix: string) => {
  return attachments.filter(file =>
    file.FileName?.toLowerCase().startsWith(prefix.toLowerCase())
  );
};
//---------handle upload Modified Matrix Files -----
const uploadAttachment = async (file: File, prefix: string) => {
  if (!RequestId) {
    alert("Please save the request first.");
    return;
  }

  const fileName = `${prefix}_${file.name}`;

  try {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .attachmentFiles.add(fileName, file);

    // reload attachments after upload
    await loadAttachments(Number(RequestId));
  } catch (error) {
    console.error("Upload failed", error);
    alert("Failed to upload attachment");
  }
};



  //------ Get Models -----
  async function GetModels(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('CarLine', 'Model,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Model, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Model }));
    setModelData(ddVehicleModels);
  }

  async function GetPRTSResource(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getData('PRTSResource', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setprtsResourceOptions(ddVehicleModels);
  }

  async function GetBuildType(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('BuildTypeList', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setBuildTypeOptions(ddVehicleModels);
  }

  async function GetAgencyData(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('PRTS_NonTechnical_Issue_Agencys', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setAgencyOptions(ddVehicleModels);
  }

  async function GetIssueCategory(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('IssueCategoryList', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setIssueCategoryOption(ddVehicleModels);
  }

  async function GetCommodityList(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('CommodityList', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setCommodityOption(ddVehicleModels);
  }

  async function GetvarientAffected(props, Model) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('CarLine', 'Model,ID,Variant', '', `Model eq '${Model}'`, { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Variant, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Variant }));
    setVarientDataOptions(ddVehicleModels);
  }

  async function GetEngineTypedata(props, event) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('CarLine', 'Model,ID,Engine', '', `Model eq '${event}'`, { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Engine, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Engine }));
    setEngineTypeOptions(ddVehicleModels);
    GetvarientAffected(IPrtsProps, event);
  }

  async function GetUserDepartment(props, Email) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('UserMaster', '*,FullName/EMail,DepartmentCode/Department', 'FullName,DepartmentCode', `FullName/EMail eq '${Email}'`, { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.DepartmentCode, item])).values());
    const first = uniqueModels[0] as any;
    setFormData(prev => ({ ...prev, mInitDept: first?.DepartmentCode?.Department || "" }));
  }

  async function GetMfgShopOptions(props) {
    const spCrudOps = await SPCRUDOPS();
    const Modeldata = await spCrudOps.getRootData('MFGShopList', 'Title,ID', '', '', { column: 'ID', isAscending: true }, props);
    const uniqueModels = Array.from(new Map(Modeldata.map((item: any) => [item.Title, item])).values());
    const ddVehicleModels: IDropdownOption[] = uniqueModels.map((item: any) => ({ key: item.ID, text: item.Title }));
    setMfgShopOptions(ddVehicleModels);
  }

  const loadDiamondUsers = async (mfgShop: string) => {
    try {
      // Using SP directly (SPCRUDOPS could be used as well)
      const spCrudOps = await SPCRUDOPS();
      const items = await spCrudOps.getRootData(
        'PRTS_7D_Matrix',
        ' UserDiamond1/Id,UserDiamond1/Title ,UserDiamond1/EMail ,UserDiamond2/Id, UserDiamond2/Title, UserDiamond2/EMail,UserDiamond3/Id,UserDiamond3/Title,UserDiamond3/EMail,UserDiamond4/Id,UserDiamond4/Title,UserDiamond4/EMail,UserDiamond5/Id,UserDiamond5/Title,UserDiamond5/EMail,UserDiamond6/Id,UserDiamond6/Title,UserDiamond6/EMail, UserDiamond7/Id,UserDiamond7/Title,UserDiamond7/EMail',
        'UserDiamond1,UserDiamond2,UserDiamond3,UserDiamond4,UserDiamond5,UserDiamond6, UserDiamond7',
        `Title eq '${mfgShop}'`,
        { column: 'ID', isAscending: true },
        IPrtsProps      // <-- FIXED
      );

      if (!items || items.length === 0) return [];
      const row = items[0] as any;
      const candidates = [
        row.UserDiamond1,
        row.UserDiamond2,
        row.UserDiamond3,
        row.UserDiamond4,
        row.UserDiamond5,
        row.UserDiamond6,
        row.UserDiamond7,
      ].filter(Boolean);

      return candidates.map((u: any) => ({ Id: u.Id, Name: u.Title, Email: u.EMail }));
    } catch (err) {
      console.error('Error loading diamond users', err);
      return [];
    }
  };

  const formatDate = (d: Date) => {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };
const uploadMatrixAttachment = async (
  key: keyof typeof attachmentMatrix,
  files: FileList | null
) => {
  if (!RequestId || !files || files.length === 0) return;

  const columnMap = {
    sosjes: {
      valueCol: "SOSJESValue",
      attachmentCol: "SOSJESAttachment",
      prefix: "SOSJESAttachment"
    },
    control: {
      valueCol: "ControlPlanValue",
      attachmentCol: "ControlPlan",
      prefix: "ControlPlan"
    },
    pfmea: {
      valueCol: "PFMEAValue",
      attachmentCol: "PFMEAAttachment",
      prefix: "PFMEAAttachment"
    },
    kaizen: {
      valueCol: "KaizenValue",
      attachmentCol: "KaizenAttachment",
      prefix: "KaizenAttachment"
    },
    qualityAlert: {
      valueCol: "QualityAlertValue",
      attachmentCol: "QualityAlertAttachment",
      prefix: "QualityAlertAttachment"
    }
  };

  const config = columnMap[key];
  const dropdownValue = attachmentMatrix[key];

  try {
    const uploadedFileNames: string[] = [];

    for (const file of Array.from(files)) {
      const fileName = `${config.prefix}_${file.name}`;

      await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(RequestId))
        .attachmentFiles.add(fileName, file);

      uploadedFileNames.push(fileName);
    }

    // 🔹 Update dropdown value + attachment file names
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        [config.valueCol]: dropdownValue,
        [config.attachmentCol]: uploadedFileNames.join("; ")
      });

    await loadAttachments(Number(RequestId));
  } catch (error) {
    console.error("Matrix upload failed", error);
    alert("Upload failed");
  }
};

const getPurgingAttachments = () => {
  const purgingValue = formData.mPurgingAttachment
    ?.toString()
    .trim()
    .toLowerCase();

  if (!purgingValue) return [];

  return attachments.filter(file =>
    file.FileName?.toLowerCase().includes(purgingValue)
  );
};

const getIssueAttachments = () => {
  const issueValue = formData.mIssueAttachment
    ?.toString()
    .trim()
    .toLowerCase();

  if (!issueValue) return [];

  return attachments.filter(file => 
    file.FileName?.toLowerCase().includes(issueValue)
  );
};


  React.useImperativeHandle(ref, () => ({
    saveBaseInfo: handleSave
  }));

  const downloadFile = (fileUrl: string) => {
    const link = document.createElement("a");
    link.href = fileUrl;
    link.download = fileUrl.split("/").pop() || "";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

 
  return (
    <div className="prts-container">

      {!isCreated && (
        <div style={{ textAlign: "right", marginBottom: 10 }}>
          <button
            className="btn btn-success"
            onClick={handleSave}
            disabled={loadingSave}
          >
            {loadingSave ? 'Saving...' : 'Save'}
          </button>
        </div>
      )}


      {/* Section A */}
      <div className="form-section">
        <div className="section-header">A. Basic Details</div>
        <div className="section-body">

          <div className="form-row">
            <label><span className="required">*</span>PRTS Source</label>
            <select disabled={isCreated} name="" id="mPRTSSource" value={formData.mPRTSSource || ''} onChange={handleChange} >
              <option value="Select">Select</option>
              {prtsResourceOptions.map(resource => (
                <option key={resource.key || resource.ID} value={resource.text}>{resource.text}</option>
              ))}
            </select>



            <label><span className="required">*</span>Initiator Name</label>
            <input id="mInitName" value={formData.mInitName || ''} onChange={handleChange} disabled />

            <label><span className="required">*</span>Initiator Department</label>
            <input id="mInitDept" value={formData.mInitDept || ''} onChange={handleChange} disabled />
          </div>

          <div className="form-row">


            <label><span className="required">*</span>MFG Shop Selection</label>
            <select id="mMFGShop" disabled={isCreated} value={formData.mMFGShop || ''} onChange={handleChange}>
              <option value="">Select</option>
              {mfgShopOptions?.map(shop => (
                <option key={shop.key} value={shop.text}>{shop.text}</option>
              ))}
            </select>

            <label><span className="required">*</span>Issue Status</label>
            <input id="mIssueStatus" value={formData.mIssueStatus || ''} onChange={handleChange} disabled />
          </div>

        </div>
      </div>

      {/* Section B */}
      <div className="form-section">
        <div className="section-header">B. VIN Details</div>
        <div className="section-body">

          <div className="form-row">
            <label><span className="required">*</span>VIN Number</label>
            <input disabled={isCreated} id="mIssueVINNo" value={formData.mIssueVINNo || ''} onChange={handleChange} />
            <label><span className="required">*</span>Build Type</label>
            <select disabled={isCreated} name="mBuildType" id="mBuildType" value={formData.mBuildType || ''} onChange={handleChange}>
              <option value="Select">Select</option>
              {buildTypeOptions.map(buildType => (
                <option key={buildType.key || buildType.ID} value={buildType.text}>{buildType.text}</option>
              ))}
            </select>

            <label><span className="required">*</span>Part Name</label>
            <input disabled={isCreated} id="mPartName" value={formData.mPartName || ''} onChange={handleChange} />

            <label><span className="required">*</span>Part No</label>
            <input disabled={isCreated} id="mPartNo" value={formData.mPartNo || ''} onChange={handleChange} />
          </div>

          <div className="form-row">
            <label><span className="required">*</span>Supplier Source</label>
            <select disabled={isCreated} name="mPartSupplierSource" id="mPartSupplierSource" value={formData.mPartSupplierSource || ''} onChange={handleChange} >
              <option value="Select">Select</option>
              <option value="KD">KD</option>
              <option value="LC">LC</option>
            </select>

            <label><span className="required">*</span>Supplier Name</label>
            <input disabled={isCreated} id="mPartSupplier" value={formData.mPartSupplier || ''} onChange={handleChange} />

            <label><span className="required">*</span>Model</label>
            <select disabled={isCreated} id="mProjectCode" value={formData.mProjectCode || ''} onChange={(e) => { const selectedModel = e.target.value; handleChange(e); GetEngineTypedata(IPrtsProps, selectedModel); }} >
              <option value="">Select</option>
              {modeldata.map(model => (
                <option key={model.key} value={model.text}>{model.text}</option>
              ))}
            </select>
            <label><span className="required">*</span>Engine Type</label>
            <select disabled={isCreated} id="mEngineType" value={formData.mEngineType || ''} onChange={handleChange}  >
              <option value="">Select</option>
              {engineTypeOptions.map(EngineType => (
                <option key={EngineType.key} value={EngineType.text}>{EngineType.text}</option>
              ))}
            </select>
          </div>

        </div>
      </div>

      {/* Section C */}
      <div className="form-section">
        <div className="section-header">C. Problem Details</div>
        <div className="section-body">

          <div className="form-row">
            <label><span className="required">*</span>Problem Definition</label>
            <textarea disabled={isCreated} id="mIssueDescription" rows={3} value={formData.mIssueDescription || ''} onChange={handleChange} />
          </div>

          <div className="form-row">
            <label><span className="required">*</span>Issue Category</label>
            <select disabled={isCreated} name="mIssueCategory" id="mIssueCategory" value={formData.mIssueCategory || ''} onChange={handleChange}>
              <option value="Select">Select</option>
              {issueCategoryOption.map(category => (
                <option key={category.key} value={category.text}>{category.text}</option>
              ))}
            </select>

            <label><span className="required">*</span>Severity</label>
            <select disabled={isCreated} value={formData.mSeverity || ''} id="mSeverity" onChange={handleChange}>
              <option value="-1">Select</option>
              <option value="50">50</option>
              <option value="20">20</option>
              <option value="10">10</option>
              <option value="5">5</option>
            </select>

            <label><span className="required">*</span>Commodity</label>
            <select name="mCommodity" disabled={isCreated} id="mCommodity" value={formData.mCommodity || ''} onChange={handleChange}>
              <option value="Select">Select</option>
              {commodityOption.map(Commodity => (
                <option key={Commodity.key} value={Commodity.text}>{Commodity.text}</option>
              ))}
            </select>
          </div>

          <div className="form-row">
            <label><span className="required">*</span>Qty Affected</label>
            <input type='number' id="mQtyAffected" disabled={isCreated} value={formData.mQtyAffected ?? ''} onChange={handleChange} />

            <label><span className="required">*</span>Variant Affected</label>
            <select id="mVariantAffected" disabled={isCreated} value={formData.mVariantAffected || ''} onChange={handleChange} >
              <option value="Select">Select</option>
              {varientDataOptions.map(varient => (
                <option key={varient.key} value={varient.text}>{varient.text}</option>
              ))}
            </select>

            <label><span className="required">*</span>Repeated Issue</label>
            <select name="mRepeatedIssue" id="mRepeatedIssue" disabled={isCreated} value={formData.mRepeatedIssue || ''} onChange={handleChange} >
              <option value="Select">Select</option>
              <option value="Yes">Yes</option>
              <option value="No">No</option>
            </select>
          </div>

          <div className="form-row">
            <label>{formData.mRepeatedIssue === "Yes" && (<span className="required">*</span>)}Issue Ref No</label>
            <input id="mRefNo" disabled={isCreated} value={formData.mRefNo || ''} onChange={handleChange} />
          </div>

          <div className="form-row" style={{ width: '100% !important' }}>
            <label><span className="required">*</span>Analysis</label>
            <textarea id="mAnalysis" disabled={isCreated} rows={3} value={formData.mAnalysis || ''} onChange={handleChange} />
          </div>

          <div className="form-row">
            <label><span className="required">*</span>Root Cause Found</label>
            <select name="mRootCauseFound" id="mRootCauseFound" value={formData.mRootCauseFound || 'Select'}
              disabled={isCreated}
              onChange={(e) => {
                const value = e.target.value;

                setFormData(prev => {
                  const updated = { ...prev, mRootCauseFound: value };
                  onFormChange(updated.mRootCauseFound, updated.mIs7D || "Select");
                  return updated;
                });

                handleChange(e);
              }}>
              <option value="Select">Select</option>
              <option value="Yes">Yes</option>
              <option value="No">No</option>
            </select>

            {/* Show 7D Process only when Root Cause is NO */}
            {formData.mRootCauseFound === "No" && (
              <>
                <label><span className="required">*</span>Require 7D Process</label>
                <select name="mIs7D" id="mIs7D" value={formData.mIs7D || 'Select'} disabled={isCreated}
                  onChange={(e) => {
                    const value = e.target.value;

                    setFormData(prev => {
                      const updated = { ...prev, mIs7D: value };
                      onFormChange(updated.mRootCauseFound || "Select", updated.mIs7D);
                      return updated;
                    });

                    handleChange(e);
                  }} >
                  <option value="Select">Select</option>
                  <option value="Yes">Yes</option>
                  <option value="No">No</option>
                </select>
              </>
            )}
          </div>

          {/* Agency + Part Quality Issue Conditions */}
          {(
            (formData.mRootCauseFound === "Yes") ||
            (formData.mRootCauseFound === "No" && formData.mIs7D === "No")
          ) && (
              <div className="form-row">
                <label><span className="required">*</span>Agency user</label>
                <PeoplePicker
                  disabled={isCreated}
                  context={IPrtsProps.currentSPContext}
                  personSelectionLimit={1}
                  showtooltip={true}
                  tooltipDirectional={1}
                  required={false}
                  //disabled={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={500}
                  defaultSelectedUsers={[formData.mPartQualityIssue]}
                  onChange={async (items: any[]) => {
                    if (items.length === 0) {
                      setFormData(prev => ({ ...prev, mPartQualityIssue: null }));
                      return;
                    }

                    const loginName = items[0].loginName;
                    const spUser = await sp.web.ensureUser(loginName);

                    setFormData(prev => ({
                      ...prev,
                      mPartQualityIssue: {
                        Id: spUser.data.Id,
                        LoginName: loginName,
                        Name: items[0].text,
                        Email: items[0].secondaryText
                      }
                    }));
                  }}
                />

                <label><span className="required">*</span>Agency</label>
                <select name="mAgency" id="mAgency" value={formData.mAgency || ''} disabled={isCreated} onChange={handleChange}>
                  <option value="Select">Select</option>
                  {agencyOptions.map(agency => (
                    <option key={agency.key} value={agency.text}>{agency.text}</option>
                  ))}
                </select>
              </div>
            )}

          <div className="form-row">
            <label><span className="required">*</span>Download Purging Template</label>
            <button
              className="btn btn-sm btn-danger"
              onClick={() => 
                downloadFile(
                  `${props.IPrtsProps.currentSPContext.pageContext.web.absoluteUrl}/Template/PurgingReportFormat.pdf`
                )
              }
            >
              Download Template
            </button>
          </div>
          <div className="form-row">
            <label><span className="required">*</span>Purging Attachment</label>
            {formData.mPurgingAttachment || (
              <input
                type="file"
                className="form-control"
                multiple
                disabled={isCreated}
                accept=".pdf"
                onChange={(e) => {
                  if (e.target.files) {
                    setSelectedFiles(Array.from(e.target.files));
                  }
                }}
              />
            )}
          </div>
        {getPurgingAttachments().length > 0 && (
  <ul>
    {getPurgingAttachments().map((file) => (
      <li key={file.FileName} className="d-flex align-items-center gap-2">
        <a
          href={file.ServerRelativeUrl}
          target="_blank"
          rel="noreferrer"
        >
          {file.FileName}
        </a>

        <button
          className="btn btn-sm btn-danger"
          onClick={() => deleteAttachment(file.FileName)}
        >
          ❌
        </button>
      </li>
    ))}
  </ul>
)}

<div className="form-row">
            <label><span className="required">*</span>Issue Attachment</label>
            {formData.mIssueAttachment || (
              <input
                type="file"
                className="form-control"
                multiple
                disabled={isCreated}
                accept=".pdf,.doc,.docx,.jpg,.jpeg,.png,.gif,.bmp,.webp"
                onChange={(e) => {
                  if (e.target.files) {
                    setSelectedIssueFiles(Array.from(e.target.files));
                  }
                }}
              />
            )}
          </div>
        {getIssueAttachments().length > 0 && (
  <ul>
    {getIssueAttachments().map((file) => (
      <li key={file.FileName} className="d-flex align-items-center gap-2">
        <a
          href={file.ServerRelativeUrl}
          target="_blank"
          rel="noreferrer"
        >
          {file.FileName}
        </a>

        <button
          className="btn btn-sm btn-danger"
          onClick={() => deleteAttachment(file.FileName)}
        >
          ❌
        </button>
      </li>
    ))}
  </ul>
)}

{isCreated && (
  <table className="table table-bordered" id="tabelAttachment">
    <colgroup>
      <col width="20%" />
      <col width="20%" />
      <col width="20%" />
      <col width="20%" />
      <col width="20%" />
    </colgroup>

    <thead>
      <tr>
        <th>SOS/JES</th>
        <th>Control Plan</th>
        <th>PFMEA</th>
        <th>Kaizen</th>
        <th>Quality Alert</th>
      </tr>
    </thead>

    <tbody>
      <tr>
        {/* SOS/JES */}
        <td>
          <select
            className="form-control"
            value={attachmentMatrix.sosjes}
onChange={async (e) => {
  const value = e.target.value;

  setAttachmentMatrix(prev => ({ ...prev, sosjes: value }));

  if (RequestId) {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        SOSJESValue: value
      });
  }
}}
          >
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>

          {attachmentMatrix.sosjes === "Yes" && (
            <input
              type="file"
              className="form-control mt-2"
              multiple
              accept=".pdf,.ppt,.pptx,.xls,.xlsx"
              onChange={(e) =>
 uploadMatrixAttachment("sosjes", e.target.files)
              }
            />
          )}
          {getAttachmentsByPrefix("SOSJESAttachment").map(file => (
  <div key={file.FileName} className="d-flex align-items-center">
    <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer">
      {file.FileName.replace("SOSJESAttachment_", "")}
    </a>
    <button
      className="btn btn-sm btn-danger"
      onClick={() => deleteAttachment(file.FileName)}
    >
      ❌
    </button>
  </div>
  
))}
        </td>

        {/* Control Plan */}
        <td>
          <select
            className="form-control"
            value={attachmentMatrix.control}
            // onChange={(e) =>
            //   setAttachmentMatrix({ ...attachmentMatrix, control: e.target.value })
            // }
            onChange={async (e) => {
  const value = e.target.value;

  setAttachmentMatrix(prev => ({ ...prev, control: value }));

  if (RequestId) {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        ControlPlanValue: value
      });
  }
}}

          >
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>

          {attachmentMatrix.control === "Yes" && (
            <input
              type="file"
              className="form-control mt-2"
              accept=".pdf,.ppt,.pptx,.xls,.xlsx"
              multiple
              onChange={(e) =>
uploadMatrixAttachment("control", e.target.files)   
           }
            />
          )}
          {getAttachmentsByPrefix("ControlPlan").map(file => (
  <div key={file.FileName} className="d-flex align-items-center">
    <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer">
      {file.FileName.replace("ControlPlan", "")}
    </a>
    <button
      className="btn btn-sm btn-danger"
      onClick={() => deleteAttachment(file.FileName)}
    >
      ❌
    </button>
  </div>
))}
        </td>

        {/* PFMEA */}
        <td>
          <select
            className="form-control"
            value={attachmentMatrix.pfmea}
            // onChange={(e) =>
            //   setAttachmentMatrix({ ...attachmentMatrix, pfmea: e.target.value })
            // }
            onChange={async (e) => {
  const value = e.target.value;

  setAttachmentMatrix(prev => ({ ...prev, pfmea: value }));

  if (RequestId) {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        PFMEAValue: value
      });
  }
}}
          >
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>

          {attachmentMatrix.pfmea === "Yes" && (
            <input
              type="file"
              className="form-control mt-2"
              multiple
              accept=".pdf,.ppt,.pptx,.xls,.xlsx"
              onChange={(e) =>
uploadMatrixAttachment("pfmea", e.target.files)
              }
            />
          )}
          {getAttachmentsByPrefix("PFMEAAttachment").map(file => (
  <div key={file.FileName} className="d-flex align-items-center">
    <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer">
      {file.FileName.replace("PFMEAAttachment_", "")}
    </a>
    <button
      className="btn btn-sm btn-danger"
      onClick={() => deleteAttachment(file.FileName)}
    >
      ❌
    </button>
  </div>
))}
        </td>

        {/* Kaizen */}
        <td>
          <select
            className="form-control"
            value={attachmentMatrix.kaizen}
            // onChange={(e) =>
            //   setAttachmentMatrix({ ...attachmentMatrix, kaizen: e.target.value })
            // }
            onChange={async (e) => {
  const value = e.target.value;

  setAttachmentMatrix(prev => ({ ...prev, kaizen: value }));

  if (RequestId) {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        KaizenValue: value
      });
  }
}}
          >
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>

          {attachmentMatrix.kaizen === "Yes" && (
            <input
              type="file"
              className="form-control mt-2"
              multiple
              accept=".pdf,.ppt,.pptx,.xls,.xlsx"
              onChange={(e) =>
uploadMatrixAttachment("kaizen", e.target.files)  
            }
            />
          )}
          {getAttachmentsByPrefix("KaizenAttachment").map(file => (
  <div key={file.FileName} className="d-flex align-items-center">
    <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer">
      {file.FileName.replace("KaizenAttachment_", "")}
    </a>
    <button
      className="btn btn-sm btn-danger"
      onClick={() => deleteAttachment(file.FileName)}
    >
      ❌
    </button>
  </div>
))}
        </td>

        {/* Quality Alert */}
        <td>
          <select
            className="form-control"
            value={attachmentMatrix.qualityAlert}
            // onChange={(e) =>
            //   setAttachmentMatrix({
            //     ...attachmentMatrix,
            //     qualityAlert: e.target.value
            //   })
            // }
            onChange={async (e) => {
  const value = e.target.value;

  setAttachmentMatrix(prev => ({ ...prev, qualityAlert: value }));

  if (RequestId) {
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        QualityAlertValue: value
      });
  }
}}
          >
            <option value="No">No</option>
            <option value="Yes">Yes</option>
          </select>

          {attachmentMatrix.qualityAlert === "Yes" && (
            <input
              type="file"
              className="form-control mt-2"
              multiple
              accept=".pdf,.ppt,.pptx,.xls,.xlsx"
              onChange={(e) =>
uploadMatrixAttachment("qualityAlert", e.target.files)              }
            />
          )}
          {getAttachmentsByPrefix("QualityAlertAttachment").map(file => (
  <div key={file.FileName} className="d-flex align-items-center">
    <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer">
      {file.FileName.replace("QualityAlertAttachment_", "")}
    </a>
    <button
      className="btn btn-sm btn-danger"
      onClick={() => deleteAttachment(file.FileName)}
    >
      ❌
    </button>
  </div>
))}
        </td>
      </tr>
    </tbody>
  </table>
)}


        </div>
      </div>

    </div>
  );
});

export default BaseInfoTab;
