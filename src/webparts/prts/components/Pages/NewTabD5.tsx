// Tab7D5.tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import "./CSS/Tab3D2.scss"; // reuse same styles as other tabs

export interface D5Data {
  vD5AssignTo?: string;
  vD5AssignDT?: string;
  vD5ActionStatus?: string;

  mD5AnalysisDetails?: string;
  mD5RootCauseFound?: string;
  mD5ActionStatus?: string;
  mD5ICA_Details?: string;
  mD5ICA_VIN?: string;
  mD5PCA_Details?: string;
  mD5PCA_VIN?: string;
  mD5_Remarks?: string;
  mD5AttachmentName?: string;
}

interface D5HistoryRow {
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

interface Tab7D5Props {
  activeData?: D5Data;
  historyData?: D5HistoryRow[];
  existingJsonArray?: any[]; // optional preloaded array
  onSave?: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D5Data;
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

const Tab7D5: React.FC<Tab7D5Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  // single source of truth
  const [form, setForm] = useState<D5Data>({
    mD5AnalysisDetails: "",
    mD5RootCauseFound: "",
    mD5ActionStatus: "",
    mD5ICA_Details: "",
    mD5ICA_VIN: "",
    mD5PCA_Details: "",
    mD5PCA_VIN: "",
    mD5_Remarks: "",
    mD5AttachmentName: "",
    ...activeData,
  });
  const [activeSubTab, setActiveSubTab] = useState<"Active" | "History">("Active");
  const [jsonArray, setJsonArray] = useState<any[]>(
    existingJsonArray && existingJsonArray.length > 0
      ? JSON.parse(JSON.stringify(existingJsonArray))
      : [defaultEmptyEntry()]
  );

  const [editOpen, setEditOpen] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});

  // sync parent activeData
  useEffect(() => {
    if (activeData) setForm((p) => ({ ...p, ...activeData }));
  }, [activeData]);

  // derive from existingJsonArray passed by parent (if any)
  useEffect(() => {
    if (existingJsonArray && existingJsonArray.length) {
      const last = existingJsonArray[existingJsonArray.length - 1] || {};
      setJsonArray(JSON.parse(JSON.stringify(existingJsonArray)));
      setForm((prev) => ({
        ...prev,
        vD5AssignTo: last.c1 || prev.vD5AssignTo,
        vD5AssignDT: last.c2 || prev.vD5AssignDT,
        vD5ActionStatus: last.c12 || prev.vD5ActionStatus,
        mD5AnalysisDetails: last.c3 || prev.mD5AnalysisDetails,
        mD5AttachmentName: last.c10 || prev.mD5AttachmentName,
        mD5RootCauseFound: last.c4 || prev.mD5RootCauseFound,
        mD5ICA_Details: last.c5 || prev.mD5ICA_Details,
        mD5ICA_VIN: last.c6 || prev.mD5ICA_VIN,
        mD5PCA_Details: last.c7 || prev.mD5PCA_Details,
        mD5PCA_VIN: last.c8 || prev.mD5PCA_VIN,
        mD5_Remarks: last.c9 || prev.mD5_Remarks,
        mD5ActionStatus: last.c12 || prev.mD5ActionStatus,
      }));
    }
  }, [existingJsonArray]);

  // load from SharePoint when RequestId present
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      try {
        const item = await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(RequestId))
          .select("D5_IssueData", "CH_Status")
          .get();

        let parsed: any[] = [];
        if (item && item.D5_IssueData) {
          try {
            parsed = JSON.parse(item.D5_IssueData);
          } catch (e) {
            console.warn("D5_IssueData JSON parse failed", e);
            parsed = [];
          }
        }

        if (parsed.length) {
          setJsonArray(parsed);
          const last = parsed[parsed.length - 1] || {};
          setForm((prev) => ({
            ...prev,
            vD5AssignTo: last.c1 || prev.vD5AssignTo,
            vD5AssignDT: last.c2 || prev.vD5AssignDT,
            vD5ActionStatus: last.c12 || prev.vD5ActionStatus,
            mD5AnalysisDetails: last.c3 || prev.mD5AnalysisDetails,
            mD5AttachmentName: last.c10 || prev.mD5AttachmentName,
            mD5RootCauseFound: last.c4 || prev.mD5RootCauseFound,
            mD5ICA_Details: last.c5 || prev.mD5ICA_Details,
            mD5ICA_VIN: last.c6 || prev.mD5ICA_VIN,
            mD5PCA_Details: last.c7 || prev.mD5PCA_Details,
            mD5PCA_VIN: last.c8 || prev.mD5PCA_VIN,
            mD5_Remarks: last.c9 || prev.mD5_Remarks,
            mD5ActionStatus: last.c12 || prev.mD5ActionStatus,
          }));
        }
      } catch (err) {
        console.error("Error fetching D5 data", err);
      }
    };

    load();
  }, [RequestId]);

  const handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>
  ) => {
    const { id, value } = e.target;
    setForm((p) => ({ ...p, [id]: value }));
    setErrors((s) => ({ ...s, [id]: "" }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files && e.target.files[0];
    if (f) {
      setForm((p) => ({ ...p, mD5AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD5AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  // Validation — same rules as D4 / D3
  const isValid_IssueDetails_D5 = () => {
    const e: Record<string, string> = {};
    // Either attachment OR analysis details
    if (!isNotBlank(form.mD5AttachmentName) && !isNotBlank(form.mD5AnalysisDetails)) {
      e.mD5AttachmentName =
        "Analysis attachment is required (upload file) or enter analysis details.";
    }

    if (!isNotBlank(form.mD5RootCauseFound)) {
      e.mD5RootCauseFound = "Please select whether root cause is found.";
    }

    if (!isNotBlank(form.mD5ActionStatus) || form.mD5ActionStatus === "-1") {
      e.mD5ActionStatus = "Please select Action Status.";
    }

    if (form.mD5RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD5ICA_Details)) e.mD5ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD5ICA_VIN)) e.mD5ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD5PCA_Details)) e.mD5PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD5PCA_VIN)) e.mD5PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD5_Remarks)) e.mD5_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  // CH_Status computation
  const computeCHStatus = (current: D5Data) => {
    let ActionStatus = current.mD5ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD5RootCauseFound === "Yes" && !isNotBlank(current.mD5PCA_Details)) {
      s = "2/6";
    } else if (current.mD5RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

const handleSave = async () => {
  if (!isValid_IssueDetails_D5()) return;

  try {
    let updated = [...jsonArray];

    // Ensure at least one row exists
    if (updated.length === 0) updated.push(defaultEmptyEntry());

    // RootCauseFound = "Yes" → ALWAYS ADD NEW ROW (Angular logic)
    if (form.mD5RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    // TEXT vs FILE mapping (Angular)
    const analysisText = form.mD5AnalysisDetails || "";
    const attachmentName = form.mD5AttachmentName || "";

    const rootCause = form.mD5RootCauseFound || "";
    const icaDetails = form.mD5ICA_Details || "";
    const icaVIN = form.mD5ICA_VIN || "";
    const pcaDetails = form.mD5PCA_Details || "";
    const pcaVIN = form.mD5PCA_VIN || "";
    const remarks = form.mD5_Remarks || "";

    const { s, ActionStatus } = computeCHStatus(form);

    const prevC1 = updated[idx]?.c1 || "";
    const prevC2 = updated[idx]?.c2 || "";
    const dt = formatDateForAngularLike(new Date());

    // FINAL Angular-accurate JSON
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
    setForm((p) => ({ ...p, mD5AttachmentName: attachmentName }));

    // No RequestId → local save
    if (!RequestId) {
      onSave?.({
        updatedArray: updated,
        jsonString: JSON.stringify(updated),
        chStatus: s,
        savedFields: { ...form, mD5AttachmentName: attachmentName },
        latestJson: updated
      });
      setEditOpen(false);
      return;
    }

    // Save to SharePoint
    await sp.web.lists
      .getByTitle("PRTSList")
      .items.getById(Number(RequestId))
      .update({
        CH_Status: s,
        D5_IssueData: JSON.stringify(updated),
      });

    // Parent callback
    onSave?.({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD5AttachmentName: attachmentName },
      latestJson: updated
    });

    setEditOpen(false);

  } catch (err) {
    console.error("Save D5 failed", err);
    alert("Failed to save D5 data: " + (err?.message || err));
  }
};


  const Err = ({ id }: { id: string }) => (errors[id] ? <div className="field-error">{errors[id]}</div> : null);
 const derivedHistory: D5HistoryRow[] =
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
        {/* <div id="D5_Tab1" className="tab-pane in active"> */}
         <div
          id="D5_Tab1"
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
                <input readOnly className="form-control" value={form.vD5AssignTo || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Assign Date</label>
                <input readOnly className="form-control" value={form.vD5AssignDT || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Action Status</label>
                <input readOnly className="form-control" value={form.vD5ActionStatus || form.mD5ActionStatus || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">B. Analysis Details</div>

            <div className="row section-row">
              <div className="col-sm-12">
                <label className="form-label">Analysis Attachment</label>
                <input readOnly className="form-control" value={form.mD5AttachmentName || form.mD5AnalysisDetails || ""} />
              </div>
            </div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Is Root Cause Found</label>
                <input readOnly className="form-control" value={form.mD5RootCauseFound || ""} />
              </div>

              <div className="col-sm-8">
                <label className="form-label">Remarks</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD5_Remarks || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">C. ICA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">ICA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD5ICA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">ICA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD5ICA_VIN || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">D. PCA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">PCA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD5PCA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">PCA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD5PCA_VIN || ""} />
              </div>
            </div>
          </div>
        </div>

        {/* History Tab */}
        {/* <div id="D5_Tab2" className="tab-pane fade"> */}
         <div
        id="D5_Tab2"
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
                  <th>ICA VIN CUT OFF</th>
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
                <h4 className="modal-title">Diamond 5 - Issue Details</h4>
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
                      <input id="vD5AssignTo" className="form-control" value={form.vD5AssignTo || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Assign Date</label>
                      <input id="vD5AssignDT" className="form-control" value={form.vD5AssignDT || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4">
                      <label className="form-label">Action Status</label>
                      <select id="mD5ActionStatus" className="form-control" value={form.mD5ActionStatus || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD5ActionStatus" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">B. Analysis Details</div>
                  <div className="row">
                    <div className="col-sm-12">
                      <label className="form-label">Analysis Attachment</label>
                      <div className="upload-row">
                        <input id="D5AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                        <div className="uploaded-filename">{form.mD5AttachmentName ? form.mD5AttachmentName : <em>No file selected</em>}</div>
                      </div>
                      <Err id="mD5AttachmentName" />
                    </div>
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <div className="col-sm-4">
                      <label className="form-label">Is Root Cause Found?</label>
                      <select id="mD5RootCauseFound" className="form-control" value={form.mD5RootCauseFound || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD5RootCauseFound" />
                    </div>

                    <div className="col-sm-8">
                      <label className="form-label">Remarks</label>
                      <textarea id="mD5_Remarks" className="form-control" rows={2} value={form.mD5_Remarks || ""} onChange={handleChange} />
                      <Err id="mD5_Remarks" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">C. ICA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">ICA Details</label>
                      <textarea id="mD5ICA_Details" className="form-control" rows={2} value={form.mD5ICA_Details || ""} onChange={handleChange} />
                      <Err id="mD5ICA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">ICA VIN Cutoff</label>
                      <input type="date" id="mD5ICA_VIN" className="form-control" value={form.mD5ICA_VIN || ""} onChange={handleChange} />
                      <Err id="mD5ICA_VIN" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">D. PCA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">PCA Details</label>
                      <textarea id="mD5PCA_Details" className="form-control" rows={2} value={form.mD5PCA_Details || ""} onChange={handleChange} />
                      <Err id="mD5PCA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">PCA VIN Cutoff</label>
                      <input type="date" id="mD5PCA_VIN" className="form-control" value={form.mD5PCA_VIN || ""} onChange={handleChange} />
                      <Err id="mD5PCA_VIN" />
                    </div>
                  </div>
                </div>
              </div>

              <div className="modal-footer modal-footer-spaced">
                <button id="btnUpdate_D5Container" type="button" className="btn btn-primary" onClick={handleSave}>
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

export default Tab7D5;
