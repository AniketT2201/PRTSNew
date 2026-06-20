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
  reqId?: string;              // <-- add this
  activeData?: TechIssueData | null;
  historyData?: any[];
  onSave?: (data: TechIssueData, updatedJsonArray?: any[]) => void;
  props?: IPrtsProps;
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

const Tab2TechnicalIssueFull: React.FC<Tab2Props> = ({ reqId: propReqId, activeData, onSave, props }) => {
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
  // Fetch NonTechnical_IssueData from SharePoint when RequestId changes
  // 1) Make fetchItemData stable
const fetchItemData = useCallback(async (rid: string) => {
  if (!rid) return;
  // setLoading(true);
  setError(null);
  try {
    const item = await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(rid))
      .select("NonTechnical_IssueData")
      .get();

    const raw = item?.NonTechnical_IssueData ?? "";
    const parsed = safeParseJsonToArray(raw);

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
  } else {
    setStoredJsonArray([]);
    setFormData(activeData ?? EMPTY_TECH);
    setEditingData(activeData ?? EMPTY_TECH);
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
    setEditingData(formData);
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
          c2: editingData.issueAssignTo || "",
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

        // append and persist
        current.push(newC);

        await sp.web.lists.getByTitle("PRTSList").items.getById(Number(reqId)).update({
          NonTechnical_IssueData: JSON.stringify(current),
          CH_Status: "2/6"
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

  return (
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
                  <button id="btnNT" className="btn btn-primary" style={{ marginRight: 20, marginTop: 3 ,width:'70px'}} onClick={handleStartEdit}>
                    Edit
                  </button>
                ) : (
                  <>
                    <button className="btn btn-success" style={{ marginRight: 8, marginTop: 3 }} onClick={handleSave}>
                      Update
                    </button>
                    <button className="btn btn-default" style={{ marginRight: 20, marginTop: 3 }} onClick={handleCancel}>
                      Close
                    </button>
                  </>
                )}
              </div>
            )}

            {/* Edit Modal */}
            {isEditing && (
              <div className="modal fade show d-block" role="dialog" aria-modal="true" tabIndex={-1} style={{ backgroundColor: "rgba(0,0,0,0.5)" }}>
                <div className="modal-dialog modal-lg" role="document">
                  <div className="modal-content">
                    <div className="modal-header">
                      <button type="button" className="close" onClick={handleCancel} aria-label="Close">
                        &times;
                      </button>
                      <h4 className="modal-title">Technical Issue Details</h4>
                    </div>

                    <div className="modal-body">
                      <div className="row top-buffer">
                        <div className="col-sm-12">
                          <label htmlFor="mNTAnalysis">
                            <span className="required">*</span> Analysis Details
                          </label>
                          <textarea id="mNTAnalysis" className="form-control" rows={3} value={editingData.mNTAnalysis || ""} onChange={handleChange} />
                        </div>
                      </div>

                      <div className="row top-buffer">
                        <div className="col-sm-3">
                          <label htmlFor="mNTRootCauseFound">
                            <span className="required">*</span> Is Root Cause Found?
                          </label>
                          <select id="mNTRootCauseFound" className="form-control" value={editingData.mNTRootCauseFound || ""} onChange={handleChange}>
                            <option value="">Select</option>
                            <option value="Yes">Yes</option>
                            <option value="No">No</option>
                          </select>
                        </div>
                      </div>

                      <div className="row top-buffer">
                        <div className="col-sm-8">
                          <label htmlFor="mNTICA_Details" className={showIcaPca(editingData.mNTRootCauseFound) ? "required" : ""}>
                            ICA Action Implementation Details
                          </label>
                          <input id="mNTICA_Details" className="form-control" maxLength={250} value={editingData.mNTICA_Details || ""} onChange={handleChange} />
                        </div>
                        <div className="col-sm-4">
                          <label htmlFor="mNTICA_VIN">VIN Cut Off</label>
<input
  type="date"
  id="mNTICA_VIN"
  className="form-control"
  value={editingData.mNTICA_VIN || ""}
  onChange={handleChange}
/>                        </div>
                      </div>

                      <div className="row top-buffer">
                        <div className="col-sm-8">
                          <label htmlFor="mNTPCA_Details">PCA Action</label>
                          <input id="mNTPCA_Details" className="form-control" maxLength={250} value={editingData.mNTPCA_Details || ""} onChange={handleChange} />
                        </div>
                        <div className="col-sm-4">
                          <label htmlFor="mNTPCA_VIN">VIN Cut Off</label>
                          <input type="date" id="mNTPCA_VIN" className="form-control" value={editingData.mNTPCA_VIN || ""} onChange={handleChange} />
                        </div>
                      </div>

                      <div className="row top-buffer">
                        <div className="col-sm-12">
                          <label htmlFor="mNT_Remarks">Remarks (if Root Cause Not Found)</label>
                          <textarea id="mNT_Remarks" className="form-control" rows={3} value={editingData.mNT_Remarks || ""} onChange={handleChange} />
                        </div>
                      </div>

                      <div className="row top-buffer">
                        <div className="col-sm-12">
                          <label htmlFor="mNT_RootCause">Root Cause</label>
                          <textarea id="mNT_RootCause" className="form-control" rows={3} value={editingData.mNT_RootCause || ""} onChange={handleChange} />
                        </div>
                      </div>

                      {/* Attachment control */}
                      <div className="row top-buffer">
                        <div className="col-sm-12">
                          <label>Attachment (names only)</label>
                          <input id="NTAttachmentFile" type="file" multiple onChange={handleFileSelect} />
                          {selectedFiles.length > 0 && (
                            <div className="selected-files-list" style={{ marginTop: 8 }}>
                              <strong>Files to add:</strong>
                              <ul>
                                {selectedFiles.map((f, idx) => (
                                  <li key={idx}>
                                    {f.name} ({Math.round(f.size / 1024)} KB)
                                    <button style={{ marginLeft: 8 }} className="btn btn-link btn-xs" onClick={() => removeSelectedFile(idx)}>
                                      Remove
                                    </button>
                                  </li>
                                ))}
                              </ul>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="modal-footer">
                      <button id="btnUpdate_NonTechnicalContainer" type="button" className="btn btn-default" onClick={handleSave}>
                        Update
                      </button>
                      <button type="button" className="btn btn-default" onClick={handleCancel}>
                        Close
                      </button>
                    </div>
                  </div>
                </div>
              </div>
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
                    <input readOnly className="txtFullWidth form-control" type="text" id="vNTAgencyName" value={formData.agencyName || ""} />
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
  );
};

export default Tab2TechnicalIssueFull;
