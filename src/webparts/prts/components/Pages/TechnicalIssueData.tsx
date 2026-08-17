// Tab2TechnicalIssueFull.tsx
import * as React from "react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { useLocation, useParams } from "react-router-dom";
import { sp } from "@pnp/sp/presets/all";
import { IPrtsProps } from '../IPrtsProps';
import './CSS/NewPage.scss'


interface TechIssueData {
  assignDate?: string;
  issueAssignTo?: string;
  agencyName?: string;
  mNTAnalysis?: string;
  mNTRootCauseFound?: string; // 'Yes' | 'No' | ''
  mNTICA_Details?: string;
  mNTICA_VIN?: string;
  mNTPCA_Details?: string;
  mNTPCA_VIN?: string;
  mNT_Remarks?: string;
  mNT_RootCause?: string;
  // helpers
  attachmentHTML?: string;
  actionDate?: string;
}

interface Tab2Props {
  props?: IPrtsProps;
  reqId?: string;              // <-- add this
  activeData?: TechIssueData | null;
  historyData?: any[];
  onSave?: (data: TechIssueData, updatedJsonArray?: any[]) => void;
}


const EMPTY_TECH: TechIssueData = {
  agencyName: "",
  issueAssignTo: "",
  assignDate: "",
  mNTAnalysis: "",
  mNTRootCauseFound: "",
  mNTICA_Details: "",
  mNTICA_VIN: "",
  mNTPCA_Details: "",
  mNTPCA_VIN: "",
  mNT_Remarks: "",
  mNT_RootCause: "",
  attachmentHTML: "",
  actionDate: "",
};

const formatDate = (d: Date) => {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
};

const isNotBlank = (v: any) =>
  v !== null &&
  v !== undefined &&
  String(v).trim() !== "" &&
  String(v).trim() !== "-1" &&
  String(v).trim() !== "?";

const safeParseJsonToArray = (raw: any): any[] => {
  if (raw === null || raw === undefined) return [];
  if (Array.isArray(raw)) return raw;
  if (typeof raw !== "string") {
    // Try to coerce objects to array
    try {
      return raw ? [raw] : [];
    } catch {
      return [];
    }
  }
  const trimmed = raw.trim();
  if (!trimmed) return [];
  try {
    const parsed = JSON.parse(trimmed);
    if (Array.isArray(parsed)) return parsed;
    // Sometimes SP stores a single object
    return parsed ? [parsed] : [];
  } catch (err) {
    console.warn("safeParseJsonToArray: JSON parse failed", err);
    return [];
  }
};

const cRowToTech = (r: any): TechIssueData => {
  if (!r) return { ...EMPTY_TECH };
  return {
    agencyName: r.c1 ?? "",
    issueAssignTo: r.c2 ?? "",
    assignDate: r.c3 ?? "",
    mNTAnalysis: r.c4 ?? "",
    mNTRootCauseFound: r.c5 ?? "",
    mNTICA_Details: r.c6 ?? "",
    mNTICA_VIN: r.c7 ?? "",
    mNTPCA_Details: r.c8 ?? "",
    mNTPCA_VIN: r.c9 ?? "",
    mNT_Remarks: r.c10 ?? "",
    attachmentHTML: r.c11 ?? "",
    actionDate: r.c12 ?? "",
    mNT_RootCause: r.c13 ?? "",
  };
};

// Build an HTML fragment for attachments by merging existingHtml and new files (just names)
const buildAttachmentHTMLWithLinks = (
  existingHtml: string | undefined,
  urls: string[]
) => {
  const existing = existingHtml || "";

  const html = urls
    .map(u => {
      const name = u.split("/").pop();
      return `
        <div class="nt-attach-item">
          <a href="${u}" target="_blank" rel="noreferrer">${name}</a>
        </div>`;
    })
    .join("");

  return existing + html;
};

const escapeHtml = (unsafe: string) =>
  unsafe
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");

const Tab2TechnicalIssueFull: React.FC<Tab2Props> = ({ props, reqId: propReqId, activeData, onSave }) => {
const { RequestId: routeRequestId } = useParams<{ RequestId?: string }>();
  // prefer explicit prop, then route param, then empty
  const location = useLocation();
  const isFromAllReqDash = new URLSearchParams(location.search).get("from") === "AllReqDash";
  const reqId = (propReqId ?? routeRequestId ?? "").toString();
  const [formData, setFormData] = useState<TechIssueData>(EMPTY_TECH);
  const [editingData, setEditingData] = useState<TechIssueData>(EMPTY_TECH);
  const [isEditing, setIsEditing] = useState(false);
  const [activeSubTab, setActiveSubTab] = useState<"Active" | "History">("Active");
  const [storedJsonArray, setStoredJsonArray] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [IsRootCauseFound, setIsRootCauseFound] = useState(false);
  const [isCreated, setIsCreated] = useState<boolean>(false);
  const [attachments, setAttachments] = useState<any[]>([]);
  const [Editor, setEditor] = useState<{ Title: string; Id: number } | null>(null);
  const [CHSStatus, setCHSStatus] = useState<any>("");
  const [nextApprover, setNextApprover] = useState<any>("");
  const [attachmentMatrix, setAttachmentMatrix] = useState({
    sosjes: "No",
    control: "No",
    pfmea: "No",
    kaizen: "No",
    qualityAlert: "No"
  });
  // Fetch NonTechnical_IssueData from SharePoint when RequestId changes
  // 1) Make fetchItemData stable

  const getUserDetailsByName = async (displayName: string) => {
    const results = await sp.web.siteUsers.filter(`Title eq '${displayName}'`).get();
    if (results.length > 0) return results[0];
    return null;
  };

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
    if (!reqId) return;
  
    const confirmDelete = confirm(`Delete attachment "${fileName}"?`);
    if (!confirmDelete) return;
  
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .attachmentFiles.getByName(fileName)
      .delete();
  
    await loadAttachments(Number(reqId));
  };

  const handleBaseInfoGet = async (reqId: string): Promise<any | null> => {
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
        const formState = {
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
          mIs7D: item.Is7DRequired || "No",
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
        return formState;
  
      } catch (error) {
        console.log("GET ERROR:", error);
        alert("Error loading data: " + (error.message ? error.message : error));
        return null;
      }
    };

    const uploadMatrixAttachment = async (
      key: keyof typeof attachmentMatrix,
      files: FileList | null
    ) => {
      if (!reqId || !files || files.length === 0) return;
    
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
            .items.getById(Number(reqId))
            .attachmentFiles.add(fileName, file);
    
          uploadedFileNames.push(fileName);
        }
    
        // 🔹 Update dropdown value + attachment file names
        await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(reqId))
          .update({
            [config.valueCol]: dropdownValue,
            [config.attachmentCol]: uploadedFileNames.join("; ")
          });
    
        await loadAttachments(Number(reqId));
      } catch (error) {
        console.error("Matrix upload failed", error);
        alert("Upload failed");
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
     if (!reqId) {
       alert("Please save the request first.");
       return;
     }
   
     const fileName = `${prefix}_${file.name}`;
   
     try {
       await sp.web.lists
         .getByTitle("PRTSList")
         .items.getById(Number(reqId))
         .attachmentFiles.add(fileName, file);
   
       // reload attachments after upload
       await loadAttachments(Number(reqId));
     } catch (error) {
       console.error("Upload failed", error);
       alert("Failed to upload attachment");
     }
   };

const fetchItemData = useCallback(async (rid: string) => {
  if (!rid) return;
  // setLoading(true);
  setError(null);
  try {
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(rid))
      .select("NonTechnical_IssueData, Editor/Title, Editor/Id, CH_Status, NA/Title, NA/Id")
      .expand("Editor, NA")
      .get();

    const raw = item?.NonTechnical_IssueData ?? "";
    const parsed = safeParseJsonToArray(raw);
    setEditor(item?.Editor ?? null);
    setCHSStatus(item?.CH_Status ?? "");
    setNextApprover(item?.NA?.Title ?? "");
    setStoredJsonArray(parsed);

    if (parsed.length > 0) {
      const last = parsed[parsed.length - 1];
      const mapped = cRowToTech(last);
      if(mapped.mNTRootCauseFound === "Yes"){
        setIsRootCauseFound(true);
      }
      setFormData(mapped);
      setEditingData(mapped);
    } else {
      const initial = activeData ?? EMPTY_TECH;
      setFormData(initial);
      setEditingData(initial);
    }
  } catch (err: any) {
    console.error("fetchItemData error", err);
    setError("Failed to load technical data.");

    const initial = activeData ?? EMPTY_TECH;
    setFormData(initial);
    setEditingData(initial);
    setStoredJsonArray([]);
  } finally {
    setLoading(false);
  }
}, [activeData, EMPTY_TECH]); 

useEffect(() => {
  if (reqId) {
    fetchItemData(reqId);
    handleBaseInfoGet(reqId);
    loadAttachments(Number(reqId));
  } else {
    setStoredJsonArray([]);
    setFormData(activeData ?? EMPTY_TECH);
    setEditingData(activeData ?? EMPTY_TECH);
    handleBaseInfoGet(reqId);
    loadAttachments(Number(reqId));
  }
}, [reqId, activeData, EMPTY_TECH]); 

  // handle field change (works for input/select/textarea)
  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const id = (e.target.id || e.target.getAttribute("name") || "").toString();
    const value = e.target.value;
    if (!id) return;
    setEditingData((prev) => ({ ...prev, [id]: value } as TechIssueData));
  };

  const validateNTIssueDetails = (data: TechIssueData) => {
    let msg = "";

    const txtfield = [
      { val: data.mNTAnalysis, label: "Analysis Details" },
      { val: data.mNTRootCauseFound, label: "Is Root Cause Found" },
    ];
    txtfield.forEach((f) => {
      if (!isNotBlank(f.val)) msg += `<li>${f.label}</li>`;
    });

    if (String(data.mNTRootCauseFound) === "Yes") {
      const conds = [
        { val: data.mNTICA_Details, label: "ICA Details" },
        { val: data.mNTICA_VIN, label: "ICA VIN Cut Off" },
        { val: data.mNTPCA_Details, label: "PCA Details" },
        { val: data.mNTPCA_VIN, label: "PCA VIN Cut Off" },
      ];
      conds.forEach((c) => {
        if (!isNotBlank(c.val)) msg += `<li>${c.label}</li>`;
      });
      if (!isNotBlank(data.mNT_RootCause)) msg += `<li>Provide Root Cause, As Root Cause Found</li>`;
    }

  // if (String(data.mNTRootCauseFound) === "No" && !isNotBlank(data.mNT_Remarks)) {
  //   msg += `<li>Provide Remarks, As Root Cause Not Found</li>`;
  // }

    if (isNotBlank(msg)) {
      const textMsg = msg.replace(/<li>/g, "• ").replace(/<\/li>/g, "\n");
      alert(`Missing Data in Mandatory Fields:\n\n${textMsg}`);
      return false;
    }
    return true;
  };

  const handleStartEdit = () => {
    setEditingData({
      ...EMPTY_TECH,
      agencyName: formData.agencyName,
      issueAssignTo: formData.issueAssignTo,
      assignDate: formData.assignDate,
    });
    setIsEditing(true);
  };

  const handleCancel = () => {
    setEditingData(formData);
    setIsEditing(false);
    setSelectedFiles([]);
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;
    const arr = Array.from(files);
    setSelectedFiles((prev) => [...prev, ...arr]);
    
  };

  const removeSelectedFile = (index: number) => {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
  };
const uploadAttachmentsToList = async (
  reqId: string,
  files: File[]
): Promise<string[]> => {
  if (!reqId || files.length === 0) return [];

  const urls: string[] = [];

  for (const file of files) {
    const res = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(reqId))
      .attachmentFiles.add(file.name, file);

    // ✅ Cast required because typings are incomplete
    const serverUrl = (res.data as any).ServerRelativeUrl;
    urls.push(serverUrl);
  }

  return urls;
};

  // Save handler: append new row to history and update list item in SP
  const handleSave = useCallback(
    async (ev?: React.MouseEvent) => {
      if (ev) ev.preventDefault();
      try {
        setLoading(true);
        const isAssignedToCurrentUser =
          editingData.issueAssignTo?.trim().toLowerCase() === props.userDisplayName?.trim().toLowerCase();

        if (isAssignedToCurrentUser) {
          const requiredFields = [
            "agencyName",
            "issueAssignTo",
            "assignDate",
            "mNTAnalysis",
            "mNTRootCauseFound",
            "mNTICA_Details",
            "mNTICA_VIN",
            "mNTPCA_Details",
            "mNTPCA_VIN",
          ];

          const missingFields = requiredFields.filter(
            (field) => !editingData[field]?.toString().trim()
          );

          if (missingFields.length > 0) {
            alert("Please fill all mandatory fields before submitting.");
            return;
          }
        }
        if (!validateNTIssueDetails(editingData)) return;

        if (!reqId) {
          alert("RequestId not found. Cannot save to SharePoint.");
          return;
        }

        // setLoading(true);

        // clone current stored json array
        let current: any[] = Array.isArray(storedJsonArray) ? JSON.parse(JSON.stringify(storedJsonArray)) : [];

        const Actiondt = formatDate(new Date());
        const remarks = String(editingData.mNTRootCauseFound) === "No" ? editingData.mNT_Remarks || "" : "";
        const rootCauseDetails = String(editingData.mNTRootCauseFound) === "Yes" ? editingData.mNT_RootCause || "" : "";

        const lastIndex = current.length - 1;
        const existingAttachment = lastIndex >= 0 && isNotBlank(current[lastIndex]?.c11) ? current[lastIndex].c11 : "";
        const uploadedUrls =
          selectedFiles.length > 0
            ? await uploadAttachmentsToList(reqId, selectedFiles)
            : [];
            const attachmentHtml = buildAttachmentHTMLWithLinks(
          existingAttachment,
          uploadedUrls
        );
        const newC = {
          c1: editingData.agencyName || "",
          c2: nextApprover || editingData.issueAssignTo || "",
          c3: editingData.assignDate || "",
          c4: editingData.mNTAnalysis || "",
          c5: editingData.mNTRootCauseFound || "",
          c6: editingData.mNTICA_Details || "",
          c7: editingData.mNTICA_VIN || "",
          c8: editingData.mNTPCA_Details || "",
          c9: editingData.mNTPCA_VIN || "",
          c10: remarks,
          c11: attachmentHtml,
          c12: Actiondt,
          c13: rootCauseDetails,
        };
        // Add the new record to history
        current.push(newC);

        await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update({
          NonTechnical_IssueData: JSON.stringify(current),
          CH_Status: CHSStatus === "1/6" ? "2/6" : CHSStatus
        });

        // update local UI state
        const mapped = cRowToTech(newC);
        setFormData(mapped);
        setEditingData(mapped);
        setStoredJsonArray(current);
        setIsEditing(false);
        setSelectedFiles([]);

        if (onSave) onSave(mapped, current);
        alert("Details saved successfully.");
        window.location.reload();
      } catch (err: any) {
        console.error("Save error", err);
        alert("Error saving details: " + (err?.message || err));
      } finally {
        setLoading(false);
      }
    },
    [editingData, selectedFiles, storedJsonArray, reqId, onSave]
  );

  const showIcaPca = (val?: string) => String(val || "").toLowerCase() === "yes";

  const renderedHistory = useMemo(() => {
    if (!Array.isArray(storedJsonArray) || storedJsonArray.length === 0) return null;
    return storedJsonArray.map((r, idx) => (
      <tr key={idx}>
        <td>{idx + 1}</td>
        <td>{r.c1}</td>
        <td>{r.c2}</td>
        <td>{r.c3}</td>
        <td style={{ whiteSpace: "pre-wrap" }}>{r.c4}</td>
        <td>{r.c5}</td>
        <td>{r.c6}</td>
        <td>{r.c7}</td>
        <td>{r.c8}</td>
        <td>{r.c9}</td>
        <td>{r.c10}</td>
        <td>
          <div dangerouslySetInnerHTML={{ __html: r.c11 || "" }} />
        </td>
        <td>{r.c12}</td>
      </tr>
    ));
  }, [storedJsonArray]);
  // console.log("Next Approver :", nextApprover);
  // console.log("Current User  :", props.userDisplayName);
  // console.log(
  //   nextApprover?.trim().toLowerCase() ===
  //   props.userDisplayName?.trim().toLowerCase()
  // );

  return (
    <div>
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
        <div id="Tab2" className="">
        
          {error && <div className="alert alert-danger">{error}</div>}
          
          <ul className="nav nav-tabs marginTop03 justifyActiveContent" id="Sub_NT_Tab">
            <li className={activeSubTab === "Active" ? "active" : ""}>
              <a onClick={() => setActiveSubTab("Active")}>Active</a>
            </li>
            <li className={activeSubTab === "History" ? "active" : ""}>
              <a onClick={() => setActiveSubTab("History")}>History</a>
            </li>
          </ul>

          <div className="tab-content">
            {activeSubTab === "Active" && (
              <div id="NT_Tab1" className="tab-pane in active">
                {!isFromAllReqDash && (
                  <div className="row justifycontentsavebutton">
                    {!isEditing ? (
                      (nextApprover?.trim().toLowerCase() === props.userDisplayName?.trim().toLowerCase()) && (
                        <button
                          id="btnNT"
                          className="btn btn-primary"
                          style={{ marginRight: 20, marginTop: 3, width: "70px" }}
                          onClick={handleStartEdit}
                        >
                          Edit
                        </button>
                      )
                    ) : (
                      <>
                        <button
                          className="btn btn-success"
                          style={{ marginRight: 8, marginTop: 3 }}
                          onClick={handleSave}
                        >
                          Update
                        </button>
                        <button
                          className="btn btn-default"
                          style={{ marginRight: 20, marginTop: 3 }}
                          onClick={handleCancel}
                        >
                          Close
                        </button>
                      </>
                    )}
                  </div>
                )}

                {/* Edit Modal */}
                {isEditing && (
                  <>
                    {/* BACKDROP */}
                    <div className="modal-backdrop fade show"></div>

                    {/* MODAL */}
                    <div
                      className="modal show d-block"
                      tabIndex={-1}
                      style={{ display: "block", margin: "3rem" }}
                    >
                      <div
                        className="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable"
                        style={{ transform: "translate(0, -10px)" }}
                      >
                        <div className="modal-content" style={{ maxHeight: "80vh" }}>

                          {/* HEADER */}
                          <div className="modal-header">
                            <h5 className="modal-title">
                              Technical Issue Details
                            </h5>

                            <button
                              type="button"
                              className="btn-close"
                              onClick={handleCancel}
                            />
                          </div>

                          {/* BODY */}
                          <div
                            className="modal-body"
                            style={{ maxHeight: "70vh", overflowY: "auto" }}
                          >

                            {/* ---------- KEEP YOUR EXISTING FORM HERE ---------- */}

                            <div className="row top-buffer">
                              <div className="col-sm-12">
                                <label htmlFor="mNTAnalysis">
                                  <span className="required">*</span> Analysis Details
                                </label>
                                <textarea
                                  id="mNTAnalysis"
                                  className="form-control"
                                  rows={3}
                                  value={editingData.mNTAnalysis || ""}
                                  onChange={handleChange}
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-3">
                                <label htmlFor="mNTRootCauseFound">
                                  <span className="required">*</span> Is Root Cause Found?
                                </label>
                                <select
                                  id="mNTRootCauseFound"
                                  className="form-control"
                                  value={editingData.mNTRootCauseFound || ""}
                                  onChange={handleChange}
                                >
                                  <option value="">Select</option>
                                  <option value="Yes">Yes</option>
                                  <option value="No">No</option>
                                </select>
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-8">
                                <label
                                  htmlFor="mNTICA_Details"
                                  className={
                                    showIcaPca(editingData.mNTRootCauseFound)
                                      ? "required"
                                      : ""
                                  }
                                >
                                  ICA Action Implementation Details
                                </label>

                                <input
                                  id="mNTICA_Details"
                                  className="form-control"
                                  maxLength={250}
                                  value={editingData.mNTICA_Details || ""}
                                  onChange={handleChange}
                                />
                              </div>

                              <div className="col-sm-4">
                                <label htmlFor="mNTICA_VIN"><span className="required">*</span>VIN Cut Off</label>

                                <input
                                  type="date"
                                  id="mNTICA_VIN"
                                  className="form-control"
                                  value={editingData.mNTICA_VIN || ""}
                                  onChange={handleChange}
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-8">
                                <label htmlFor="mNTPCA_Details"><span className="required">*</span>PCA Action</label>

                                <input
                                  id="mNTPCA_Details"
                                  className="form-control"
                                  maxLength={250}
                                  value={editingData.mNTPCA_Details || ""}
                                  onChange={handleChange}
                                />
                              </div>

                              <div className="col-sm-4">
                                <label htmlFor="mNTPCA_VIN"><span className="required">*</span>VIN Cut Off</label>

                                <input
                                  type="date"
                                  id="mNTPCA_VIN"
                                  className="form-control"
                                  value={editingData.mNTPCA_VIN || ""}
                                  onChange={handleChange}
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-12">
                                <label htmlFor="mNT_Remarks"><span className="required">*</span>
                                  Remarks (if Root Cause Not Found)
                                </label>

                                <textarea
                                  id="mNT_Remarks"
                                  className="form-control"
                                  rows={3}
                                  value={editingData.mNT_Remarks || ""}
                                  onChange={handleChange}
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-12">
                                <label htmlFor="mNT_RootCause"><span className="required">*</span>Root Cause</label>

                                <textarea
                                  id="mNT_RootCause"
                                  className="form-control"
                                  rows={3}
                                  value={editingData.mNT_RootCause || ""}
                                  onChange={handleChange}
                                />
                              </div>
                            </div>

                            <div className="row top-buffer">
                              <div className="col-sm-12">
                                <label><span className="required">*</span>Attachment (names only)</label>

                                <input
                                  id="NTAttachmentFile"
                                  type="file"
                                  multiple
                                  onChange={handleFileSelect}
                                />

                                {selectedFiles.length > 0 && (
                                  <div
                                    className="selected-files-list"
                                    style={{ marginTop: 8 }}
                                  >
                                    <strong>Files to add:</strong>

                                    <ul>
                                      {selectedFiles.map((f, idx) => (
                                        <li key={idx}>
                                          {f.name} ({Math.round(f.size / 1024)} KB)

                                          <button
                                            className="btn btn-link btn-sm"
                                            style={{ marginLeft: 8 }}
                                            onClick={() => removeSelectedFile(idx)}
                                          >
                                            Remove
                                          </button>
                                        </li>
                                      ))}
                                    </ul>
                                  </div>
                                )}
                              </div>
                            </div>

                            {/* Matrix Attachments */}
                            <div className="row top-buffer">
                              <div className="col-sm-12">
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
                                
                                  if (reqId) {
                                    await sp.web.lists
                                      .getByTitle("PRTSList")
                                      .items.getById(Number(reqId))
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
                                
                                  if (reqId) {
                                    await sp.web.lists
                                      .getByTitle("PRTSList")
                                      .items.getById(Number(reqId))
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
                                
                                  if (reqId) {
                                    await sp.web.lists
                                      .getByTitle("PRTSList")
                                      .items.getById(Number(reqId))
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
                                
                                  if (reqId) {
                                    await sp.web.lists
                                      .getByTitle("PRTSList")
                                      .items.getById(Number(reqId))
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
                                
                                  if (reqId) {
                                    await sp.web.lists
                                      .getByTitle("PRTSList")
                                      .items.getById(Number(reqId))
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

                          {/* FOOTER */}
                          <div className="modal-footer">
                            <button
                              className="btn btn-primary"
                              onClick={handleSave}
                              style={{ padding: "6px 26px", borderRadius: "4px" }}
                            >
                              Update
                            </button>

                            <button
                              className="btn btn-secondary"
                              onClick={handleCancel}
                              style={{ padding: "6px 26px", borderRadius: "4px" }}
                            >
                              Close
                            </button>
                          </div>

                        </div>
                      </div>
                    </div>
                  </>
                )}

                {/* Display readonly info table here */}
                <table className="table table-bordered marginTop10" style={{ width: "100%", padding: 3, marginTop: 16 }}>
                  <thead>
                    <tr>
                      <th colSpan={2}>Agency Name</th>
                      <th colSpan={2}>Issue Assign To</th>
                      <th colSpan={3}>Assign Date</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td colSpan={2}>
                        <input readOnly className="txtFullWidth form-control" type="text" id="vNTAgencyName" value={nextApprover} />
                      </td>
                      <td colSpan={2}>
                        <input readOnly className="txtFullWidth form-control" type="text" id="vNTAssignTo" value={formData.issueAssignTo || ""} />
                      </td>
                      <td colSpan={3}>
                        <input readOnly className="txtFullWidth form-control" type="text" id="vNTAssignDT" value={formData.assignDate || ""} />
                      </td>
                    </tr>

                    <tr>
                      <th colSpan={6}>Analysis details</th>
                      <th>Is Root Cause Found</th>
                    </tr>

                    <tr>
                      <td colSpan={6}>
                        <textarea readOnly rows={3} id="vNTAnalysis" className="form-control" value={formData.mNTAnalysis || ""} />
                      </td>
                      <td>
                        <input readOnly className="txtFullWidth form-control" type="text" id="vNTRootCauseFound" value={formData.mNTRootCauseFound || ""} />
                      </td>
                    </tr>

                    {showIcaPca(formData.mNTRootCauseFound) && (
                      <>
                        <tr className="Hide_If_NT_RootCauseNotFound">
                          <th colSpan={5}>ICA action Implementation details</th>
                          <th colSpan={2}>VIN Cut off</th>
                        </tr>
                        <tr className="Hide_If_NT_RootCauseNotFound">
                          <td colSpan={5}>
                            <textarea readOnly rows={3} id="vNTICA_Details" className="form-control" value={formData.mNTICA_Details || ""} />
                          </td>
                          <td colSpan={2}>
                            <input readOnly className="txtFullWidth form-control" type="text" id="vNTICA_VIN" value={formData.mNTICA_VIN || ""} />
                          </td>
                        </tr>

                        <tr className="Hide_If_NT_RootCauseNotFound">
                          <th colSpan={5}>PCA action plan</th>
                          <th colSpan={2}>VIN Cut off</th>
                        </tr>
                        <tr className="Hide_If_NT_RootCauseNotFound">
                          <td colSpan={5}>
                            <textarea readOnly rows={3} id="vNTPCA_Details" className="form-control" value={formData.mNTPCA_Details || ""} />
                          </td>
                          <td colSpan={2}>
                            <input readOnly className="txtFullWidth form-control" type="text" id="vNTPCA_VIN" value={formData.mNTPCA_VIN || ""} />
                          </td>
                        </tr>
                      </>
                    )}

                    <tr>
                      <th colSpan={4}>Remarks</th>
                      <th colSpan={3}>
                        <span className="NT_AfterSubmit">Attachment</span>
                      </th>
                    </tr>

                    <tr>
                      <td colSpan={4}>
                        <textarea readOnly rows={3} id="vNT_Remarks" className="form-control" value={formData.mNT_Remarks || ""} />
                      </td>
                      <td colSpan={3}>
                        <div id="NTAttachmentFileList" dangerouslySetInnerHTML={{ __html: formData.attachmentHTML || "" }} />
                      </td>
                    </tr>

                    <tr>
                      <th colSpan={4}>Root Cause</th>
                    </tr>

                    <tr>
                      <td colSpan={4}>
                        <textarea readOnly rows={3} id="vNT_RootCauseDetails" className="form-control" value={formData.mNT_RootCause || ""} />
                      </td>
                    </tr>
                    
                  </tbody>
                </table>
              </div>
            )}

            {activeSubTab === "History" && (
              <div id="NT_Tab2" className="">
                <div className="marginTop10">
                  <table id="NT_HistoryTable" className="table table-striped marginTop10" style={{ width: "100%" }}>
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Agency</th>
                        <th>UserName</th>
                        <th>Req_Date</th>
                        <th>Analysis Details</th>
                        <th>Found Root Cause</th>
                        <th>ICA Action</th>
                        <th>ICA VIN Cut OFF</th>
                        <th>PCA Action</th>
                        <th>PCA VIN CUT OFF</th>
                        <th>Remarks</th>
                        <th>Attachment</th>
                        <th>Action DateTime</th>
                      </tr>
                    </thead>
                    <tbody>{renderedHistory ?? <tr><td colSpan={13} className="text-center">No history records</td></tr>}</tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {loading && (
            <div style={{ position: "fixed", left: 8, bottom: 8, background: "black", padding: 8, border: "1px solid #ddd", borderRadius: 4 }}>
              Loading...
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default Tab2TechnicalIssueFull;
