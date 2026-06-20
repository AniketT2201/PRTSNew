// Tab6D4.tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import "./CSS/Tab3D2.scss"; // reuse same styles as other tabs

export interface D4Data {
  vD4AssignTo?: string;
  vD4AssignDT?: string;
  vD4ActionStatus?: string;

  mD4AnalysisDetails?: string;
  mD4RootCauseFound?: string;
  mD4ActionStatus?: string;
  mD4ICA_Details?: string;
  mD4ICA_VIN?: string;
  mD4PCA_Details?: string;
  mD4PCA_VIN?: string;
  mD4_Remarks?: string;
  mD4AttachmentName?: string;
}

interface D4HistoryRow {
  agency: string;
  userName: string;
  analysisDetails: string;
  foundRootCause: string;
  icaAction: string;
  icaVINCutOff: string;
  pcaAction: string;
  pcaVINCutOff: string;
  remarks: string;
  attachment: string;
  actionDateTime: string;
}

interface Tab6D4Props {
  activeData?: D4Data;
  historyData?: D4HistoryRow[];
  existingJsonArray?: any[]; // optional preloaded array
  onSave?: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D4Data;
     latestJson: any[];
  }) => void;
}

const defaultEmptyEntry = () => ({
  c1: "",
  c2: "",
  c3: "",
  c4: "",
  c5: "",
  c6: "",
  c7: "",
  c8: "",
  c9: "",
  c10: "",
  c11: "",
  c12: "",
});

function formatDateForAngularLike(d: Date) {
  const pad = (n: number) => (n < 10 ? "0" + n : n);
  return (
    pad(d.getDate()) +
    "/" +
    pad(d.getMonth() + 1) +
    "/" +
    d.getFullYear() +
    " " +
    pad(d.getHours()) +
    ":" +
    pad(d.getMinutes()) +
    ":" +
    pad(d.getSeconds())
  );
}

const Tab6D4: React.FC<Tab6D4Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  // single source of truth
  const [form, setForm] = useState<D4Data>({
    mD4AnalysisDetails: "",
    mD4RootCauseFound: "",
    mD4ActionStatus: "",
    mD4ICA_Details: "",
    mD4ICA_VIN: "",
    mD4PCA_Details: "",
    mD4PCA_VIN: "",
    mD4_Remarks: "",
    mD4AttachmentName: "",
    ...activeData,
  });

  const [jsonArray, setJsonArray] = useState<any[]>(
    existingJsonArray && existingJsonArray.length > 0
      ? JSON.parse(JSON.stringify(existingJsonArray))
      : [defaultEmptyEntry()]
  );
  const [activeSubTab, setActiveSubTab] = useState<"Active" | "History">("Active");
  const [editOpen, setEditOpen] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});

  // sync activeData changes
  useEffect(() => {
    if (activeData) setForm((p) => ({ ...p, ...activeData }));
  }, [activeData]);

  // if parent passed existingJsonArray, derive form values from last element
  useEffect(() => {
    if (existingJsonArray && existingJsonArray.length) {
      const last = existingJsonArray[existingJsonArray.length - 1] || {};
      setJsonArray(JSON.parse(JSON.stringify(existingJsonArray)));
      setForm((prev) => ({
        ...prev,
        vD4AssignTo: last.c1 || prev.vD4AssignTo,
        vD4AssignDT: last.c2 || prev.vD4AssignDT,
        vD4ActionStatus: last.c12 || prev.vD4ActionStatus,
        mD4AnalysisDetails: last.c3 || prev.mD4AnalysisDetails,
        mD4AttachmentName: last.c10 || prev.mD4AttachmentName,
        mD4RootCauseFound: last.c4 || prev.mD4RootCauseFound,
        mD4ICA_Details: last.c5 || prev.mD4ICA_Details,
        mD4ICA_VIN: last.c6 || prev.mD4ICA_VIN,
        mD4PCA_Details: last.c7 || prev.mD4PCA_Details,
        mD4PCA_VIN: last.c8 || prev.mD4PCA_VIN,
        mD4_Remarks: last.c9 || prev.mD4_Remarks,
        mD4ActionStatus: last.c12 || prev.mD4ActionStatus,
      }));
    }
  }, [existingJsonArray]);

  // load from SP if RequestId provided
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      try {
        const item = await sp.web.lists.getByTitle("PRTSList").items.getById(Number(RequestId)).select("D4_IssueData", "CH_Status").get();
        let parsed: any[] = [];
        if (item && item.D4_IssueData) {
          try {
            parsed = JSON.parse(item.D4_IssueData);
          } catch (e) {
            console.warn("D4_IssueData JSON parse failed", e);
            parsed = [];
          }
        }

        if (parsed.length) {
          setJsonArray(parsed);
          const last = parsed[parsed.length - 1] || {};
          setForm((prev) => ({
            ...prev,
            vD4AssignTo: last.c1 || prev.vD4AssignTo,
            vD4AssignDT: last.c2 || prev.vD4AssignDT,
            vD4ActionStatus: last.c12 || prev.vD4ActionStatus,
            mD4AnalysisDetails: last.c3 || prev.mD4AnalysisDetails,
            mD4AttachmentName: last.c10 || prev.mD4AttachmentName,
            mD4RootCauseFound: last.c4 || prev.mD4RootCauseFound,
            mD4ICA_Details: last.c5 || prev.mD4ICA_Details,
            mD4ICA_VIN: last.c6 || prev.mD4ICA_VIN,
            mD4PCA_Details: last.c7 || prev.mD4PCA_Details,
            mD4PCA_VIN: last.c8 || prev.mD4PCA_VIN,
            mD4_Remarks: last.c9 || prev.mD4_Remarks,
            mD4ActionStatus: last.c12 || prev.mD4ActionStatus,
          }));
        }
      } catch (err) {
        console.error("Error fetching D4 data", err);
      }
    };

    load();
  }, [RequestId]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { id, value } = e.target;
    setForm((p) => ({ ...p, [id]: value }));
    setErrors((s) => ({ ...s, [id]: "" }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files && e.target.files[0];
    if (f) {
      setForm((p) => ({ ...p, mD4AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD4AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  // Validation same as D3
  const isValid_IssueDetails_D4 = () => {
    const e: Record<string, string> = {};
    // Either attachment OR analysis details
    if (!isNotBlank(form.mD4AttachmentName) && !isNotBlank(form.mD4AnalysisDetails)) {
      e.mD4AttachmentName = "Analysis attachment is required (upload file) or enter analysis details.";
    }

    if (!isNotBlank(form.mD4RootCauseFound)) {
      e.mD4RootCauseFound = "Please select whether root cause is found.";
    }

    if (!isNotBlank(form.mD4ActionStatus) || form.mD4ActionStatus === "-1") {
      e.mD4ActionStatus = "Please select Action Status.";
    }

    if (form.mD4RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD4ICA_Details)) e.mD4ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD4ICA_VIN)) e.mD4ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD4PCA_Details)) e.mD4PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD4PCA_VIN)) e.mD4PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD4_Remarks)) e.mD4_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  // same CH_Status logic as other tabs
  const computeCHStatus = (current: D4Data) => {
    let ActionStatus = current.mD4ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD4RootCauseFound === "Yes" && !isNotBlank(current.mD4PCA_Details)) {
      s = "2/6";
    } else if (current.mD4RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

 const handleSave = async () => {
  if (!isValid_IssueDetails_D4()) return;

  try {
    let updated = [...jsonArray];

    // Ensure at least one row exists
    if (updated.length === 0) updated.push(defaultEmptyEntry());

    // If RootCauseFound = Yes → always add NEW row (Angular behavior)
    if (form.mD4RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    // Correct mapping (Angular)
    const analysisText = form.mD4AnalysisDetails || "";
    const attachmentName = form.mD4AttachmentName || "";

    const rootCause = form.mD4RootCauseFound || "";
    const icaDetails = form.mD4ICA_Details || "";
    const icaVIN = form.mD4ICA_VIN || "";
    const pcaDetails = form.mD4PCA_Details || "";
    const pcaVIN = form.mD4PCA_VIN || "";
    const remarks = form.mD4_Remarks || "";

    const { s, ActionStatus } = computeCHStatus(form);

    // preserve existing assignment fields
    const prevC1 = updated[idx]?.c1 || "";
    const prevC2 = updated[idx]?.c2 || "";

    const dt = formatDateForAngularLike(new Date());

    // FINAL Angular-accurate JSON object
    updated[idx] = {
      c1: prevC1,
      c2: prevC2,
      c3: analysisText,     // TEXT ONLY
      c4: rootCause,
      c5: icaDetails,
      c6: icaVIN,
      c7: pcaDetails,
      c8: pcaVIN,
      c9: remarks,
      c10: attachmentName,  // FILE ONLY
      c11: dt,
      c12: ActionStatus,
    };

    setJsonArray(updated);
    setForm((p) => ({ ...p, mD4AttachmentName: attachmentName }));

    // If RequestId missing → local save only
    if (!RequestId) {
      console.warn("No RequestId…saving only locally.");
      onSave?.({
        updatedArray: updated,
        jsonString: JSON.stringify(updated),
        chStatus: s,
        savedFields: { ...form, mD4AttachmentName: attachmentName },
        latestJson: updated,
      });
      setEditOpen(false);
      return;
    }

    // SAVE TO SHAREPOINT
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        CH_Status: s,
        D4_IssueData: JSON.stringify(updated),
      });

    // Callback to parent
    onSave?.({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD4AttachmentName: attachmentName },
      latestJson: updated,
    });

    setEditOpen(false);
  } catch (err) {
    console.error("Save D4 failed", err);
    alert("Failed to save D4 data: " + (err?.message || err));
  }
};

  const Err = ({ id }: { id: string }) => (errors[id] ? <div className="field-error">{errors[id]}</div> : null);
 const derivedHistory: D4HistoryRow[] =
    jsonArray.length > 1
      ? jsonArray.slice(0, jsonArray.length - 1).map((r, i) => ({
        agency: "",                 // not stored in D2 JSON
        userName: r.c1,
        analysisDetails: r.c3,
        foundRootCause: r.c4,
        icaAction: r.c5,
        icaVINCutOff: r.c6,
        pcaAction: r.c7,
        pcaVINCutOff: r.c8,
        remarks: r.c9,
        attachment: r.c10,
        actionDateTime: r.c11,
      }))
      : [];
  return (
    <>
      <ul className="nav nav-tabs marginTop03 justifyActiveContent">
        <li className={activeSubTab === "Active" ? "active" : ""}>
          <a onClick={() => setActiveSubTab("Active")}>Active</a>
        </li>
        <li className={activeSubTab === "History" ? "active" : ""}>
          <a onClick={() => setActiveSubTab("History")}>History</a>
        </li>
      </ul>


      <div className="tab-content">
        {/* <div id="D4_Tab1" className="tab-pane in active"> */}
         <div
          id="D4_Tab1"
          className={`tab-pane ${activeSubTab === "Active" ? "in active" : "fade"}`}
        >
          <div  className="d-flex justify-content-end mb-3" style={{ textAlign: "right" }}>
            <button className="btn btn-primary" onClick={() => setEditOpen(true)} style={{ marginRight: 20, marginTop: 3 }}>
              Edit
            </button>
          </div>

          <div className="section-card">
            <div className="section-headerD2">A. Issue Details</div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Issue Assign To</label>
                <input readOnly className="form-control" value={form.vD4AssignTo || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Assign Date</label>
                <input readOnly className="form-control" value={form.vD4AssignDT || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Action Status</label>
                <input readOnly className="form-control" value={form.vD4ActionStatus || form.mD4ActionStatus || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">B. Analysis Details</div>

            <div className="row section-row">
              <div className="col-sm-12">
                <label className="form-label">Analysis Attachment</label>
                <input readOnly className="form-control" value={form.mD4AttachmentName || form.mD4AnalysisDetails || ""} />
              </div>
            </div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Is Root Cause Found</label>
                <input readOnly className="form-control" value={form.mD4RootCauseFound || ""} />
              </div>

              <div className="col-sm-8">
                <label className="form-label">Remarks</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD4_Remarks || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">C. ICA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">ICA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD4ICA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">ICA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD4ICA_VIN || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">D. PCA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">PCA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD4PCA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">PCA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD4PCA_VIN || ""} />
              </div>
            </div>
          </div>
        </div>

        {/* History Tab uses historyData prop */}
        {/* <div id="D4_Tab2" className="tab-pane fade"> */}
         <div
        id="D4_Tab2"
        className={`tab-pane ${activeSubTab === "History" ? "in active" : "fade"}`}
      >
          <div className="marginTop10">
            <table className="table table-striped">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Agency</th>
                  <th>UserName</th>
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
              <tbody>
                {derivedHistory && derivedHistory.length ? (
                  derivedHistory.map((row, idx) => (
                    <tr key={idx}>
                      <td>{idx + 1}</td>
                      <td>{row.agency}</td>
                      <td>{row.userName}</td>
                      <td>{row.analysisDetails}</td>
                      <td>{row.foundRootCause}</td>
                      <td>{row.icaAction}</td>
                      <td>{row.icaVINCutOff}</td>
                      <td>{row.pcaAction}</td>
                      <td>{row.pcaVINCutOff}</td>
                      <td>{row.remarks}</td>
                      <td>{row.attachment}</td>
                      <td>{row.actionDateTime}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={12} style={{ textAlign: "center" }}>
                      No history available
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Edit modal */}
      {editOpen && (
        <div className="modal fade show d-block modal-overlay" tabIndex={-1} role="dialog" aria-modal="true">
          <div className="modal-dialog modal-lg" role="document">
            <div className="modal-content modal-card">
              <div className="modal-header modal-header-red">
                <h4 className="modal-title">Diamond 4 - Issue Details</h4>
                <button type="button" className="close" onClick={() => setEditOpen(false)}>
                  &times;
                </button>
              </div>

              <div className="modal-body">
                <div className="section-card small">
                  <div className="section-headerD2 small">A. Issue Details</div>
                  <div className="row">
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Issue Assign To</label>
                      <input id="vD4AssignTo" className="form-control" value={form.vD4AssignTo || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Assign Date</label>
                      <input id="vD4AssignDT" className="form-control" value={form.vD4AssignDT || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4">
                      <label className="form-label">Action Status</label>
                      <select id="mD4ActionStatus" className="form-control" value={form.mD4ActionStatus || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD4ActionStatus" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">B. Analysis Details</div>
                  <div className="row">
                    <div className="col-sm-12">
                      <label className="form-label">Analysis Attachment</label>
                      <div className="upload-row">
                        <input id="D4AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                        <div className="uploaded-filename">{form.mD4AttachmentName ? form.mD4AttachmentName : <em>No file selected</em>}</div>
                      </div>
                      <Err id="mD4AttachmentName" />
                    </div>
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <div className="col-sm-4">
                      <label className="form-label">Is Root Cause Found?</label>
                      <select id="mD4RootCauseFound" className="form-control" value={form.mD4RootCauseFound || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD4RootCauseFound" />
                    </div>

                    <div className="col-sm-8">
                      <label className="form-label">Remarks</label>
                      <textarea id="mD4_Remarks" className="form-control" rows={2} value={form.mD4_Remarks || ""} onChange={handleChange} />
                      <Err id="mD4_Remarks" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">C. ICA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">ICA Details</label>
                      <textarea id="mD4ICA_Details" className="form-control" rows={2} value={form.mD4ICA_Details || ""} onChange={handleChange} />
                      <Err id="mD4ICA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">ICA VIN Cutoff</label>
                      <input type="date" id="mD4ICA_VIN" className="form-control" value={form.mD4ICA_VIN || ""} onChange={handleChange} />
                      <Err id="mD4ICA_VIN" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">D. PCA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">PCA Details</label>
                      <textarea id="mD4PCA_Details" className="form-control" rows={2} value={form.mD4PCA_Details || ""} onChange={handleChange} />
                      <Err id="mD4PCA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">PCA VIN Cutoff</label>
                      <input type="date" id="mD4PCA_VIN" className="form-control" value={form.mD4PCA_VIN || ""} onChange={handleChange} />
                      <Err id="mD4PCA_VIN" />
                    </div>
                  </div>
                </div>
              </div>

              <div className="modal-footer modal-footer-spaced">
                <button id="btnUpdate_D4Container" type="button" className="btn btn-primary" onClick={handleSave}>
                  Update
                </button>
                <button type="button" className="btn btn-default" onClick={() => setEditOpen(false)}>
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default Tab6D4;
