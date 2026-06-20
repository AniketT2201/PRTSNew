// Tab9D7.tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import "./CSS/Tab3D2.scss"; // reuse same styles as other tabs

export interface D7Data {
  vD7AssignTo?: string;
  vD7AssignDT?: string;
  vD7ActionStatus?: string;

  mD7AnalysisDetails?: string;
  mD7RootCauseFound?: string;
  mD7ActionStatus?: string;
  mD7ICA_Details?: string;
  mD7ICA_VIN?: string;
  mD7PCA_Details?: string;
  mD7PCA_VIN?: string;
  mD7_Remarks?: string;
  mD7AttachmentName?: string;
}

interface D7HistoryRow {
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

interface Tab9D7Props {
  activeData?: D7Data;
  historyData?: D7HistoryRow[];
  existingJsonArray?: any[]; // optional preloaded array
  onSave?: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D7Data;
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

// sample uploaded image path from conversation assets (used as UI placeholder/preview)
const exampleLocalImage = "/mnt/data/ba239be0-6401-4043-8057-381a70485c6c.png";

const Tab9D7: React.FC<Tab9D7Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  const [form, setForm] = useState<D7Data>({
    mD7AnalysisDetails: "",
    mD7RootCauseFound: "",
    mD7ActionStatus: "",
    mD7ICA_Details: "",
    mD7ICA_VIN: "",
    mD7PCA_Details: "",
    mD7PCA_VIN: "",
    mD7_Remarks: "",
    mD7AttachmentName: "",
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

  // sync when parent activeData changes
  useEffect(() => {
    if (activeData) setForm((p) => ({ ...p, ...activeData }));
  }, [activeData]);

  // derive from existingJsonArray if given
  useEffect(() => {
    if (existingJsonArray && existingJsonArray.length) {
      const last = existingJsonArray[existingJsonArray.length - 1] || {};
      setJsonArray(JSON.parse(JSON.stringify(existingJsonArray)));
      setForm((prev) => ({
        ...prev,
        vD7AssignTo: last.c1 || prev.vD7AssignTo,
        vD7AssignDT: last.c2 || prev.vD7AssignDT,
        vD7ActionStatus: last.c12 || prev.vD7ActionStatus,
        mD7AnalysisDetails: last.c3 || prev.mD7AnalysisDetails,
        mD7AttachmentName: last.c10 || prev.mD7AttachmentName,
        mD7RootCauseFound: last.c4 || prev.mD7RootCauseFound,
        mD7ICA_Details: last.c5 || prev.mD7ICA_Details,
        mD7ICA_VIN: last.c6 || prev.mD7ICA_VIN,
        mD7PCA_Details: last.c7 || prev.mD7PCA_Details,
        mD7PCA_VIN: last.c8 || prev.mD7PCA_VIN,
        mD7_Remarks: last.c9 || prev.mD7_Remarks,
        mD7ActionStatus: last.c12 || prev.mD7ActionStatus,
      }));
    }
  }, [existingJsonArray]);

  // fetch from SharePoint when RequestId present
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      try {
        const item = await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(RequestId))
          .select("D7_IssueData", "CH_Status")
          .get();

        let parsed: any[] = [];
        if (item && item.D7_IssueData) {
          try {
            parsed = JSON.parse(item.D7_IssueData);
          } catch (e) {
            console.warn("D7_IssueData JSON parse failed", e);
            parsed = [];
          }
        }

        if (parsed.length) {
          setJsonArray(parsed);
          const last = parsed[parsed.length - 1] || {};
          setForm((prev) => ({
            ...prev,
            vD7AssignTo: last.c1 || prev.vD7AssignTo,
            vD7AssignDT: last.c2 || prev.vD7AssignDT,
            vD7ActionStatus: last.c12 || prev.vD7ActionStatus,
            mD7AnalysisDetails: last.c3 || prev.mD7AnalysisDetails,
            mD7AttachmentName: last.c10 || prev.mD7AttachmentName,
            mD7RootCauseFound: last.c4 || prev.mD7RootCauseFound,
            mD7ICA_Details: last.c5 || prev.mD7ICA_Details,
            mD7ICA_VIN: last.c6 || prev.mD7ICA_VIN,
            mD7PCA_Details: last.c7 || prev.mD7PCA_Details,
            mD7PCA_VIN: last.c8 || prev.mD7PCA_VIN,
            mD7_Remarks: last.c9 || prev.mD7_Remarks,
            mD7ActionStatus: last.c12 || prev.mD7ActionStatus,
          }));
        }
      } catch (err) {
        console.error("Error fetching D7 data", err);
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
      // For parity with D4/D5/D6: store filename in JSON (c10). Add actual upload logic if needed.
      setForm((p) => ({ ...p, mD7AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD7AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  const isValid_IssueDetails_D7 = () => {
    const e: Record<string, string> = {};

    if (!isNotBlank(form.mD7AttachmentName) && !isNotBlank(form.mD7AnalysisDetails)) {
      e.mD7AttachmentName = "Analysis attachment is required (upload file) or enter analysis details.";
    }

    if (!isNotBlank(form.mD7RootCauseFound)) {
      e.mD7RootCauseFound = "Please select whether root cause is found.";
    }

    if (!isNotBlank(form.mD7ActionStatus) || form.mD7ActionStatus === "-1") {
      e.mD7ActionStatus = "Please select Action Status.";
    }

    if (form.mD7RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD7ICA_Details)) e.mD7ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD7ICA_VIN)) e.mD7ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD7PCA_Details)) e.mD7PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD7PCA_VIN)) e.mD7PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD7_Remarks)) e.mD7_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const computeCHStatus = (current: D7Data) => {
    let ActionStatus = current.mD7ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD7RootCauseFound === "Yes" && !isNotBlank(current.mD7PCA_Details)) {
      s = "2/6";
    } else if (current.mD7RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

 const handleSave = async () => {
  if (!isValid_IssueDetails_D7()) return;

  try {
    let updated = [...jsonArray];

    // Ensure at least one entry exists
    if (updated.length === 0) updated.push(defaultEmptyEntry());

    // Angular rule:
    // When RootCauseFound = "Yes" → ALWAYS append a new EMPTY row
    if (form.mD7RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    // Correct JSON mapping
    const analysisText = form.mD7AnalysisDetails || "";
    const attachmentName = form.mD7AttachmentName || "";

    const rootCause = form.mD7RootCauseFound || "";
    const icaDetails = form.mD7ICA_Details || "";
    const icaVIN = form.mD7ICA_VIN || "";
    const pcaDetails = form.mD7PCA_Details || "";
    const pcaVIN = form.mD7PCA_VIN || "";
    const remarks = form.mD7_Remarks || "";

    const { s, ActionStatus } = computeCHStatus(form);

    const prevC1 = updated[idx]?.c1 || "";
    const prevC2 = updated[idx]?.c2 || "";

    const dt = formatDateForAngularLike(new Date());

    // FINAL correct JSON structure for D7
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
      c10: attachmentName,  // FILE NAME ONLY
      c11: dt,
      c12: ActionStatus,
    };

    // Update UI
    setJsonArray(updated);
    setForm((prev) => ({ ...prev, mD7AttachmentName: attachmentName }));

    // If no RequestId → return result to parent
    if (!RequestId) {
      onSave?.({
        updatedArray: updated,
        jsonString: JSON.stringify(updated),
        chStatus: s,
        savedFields: { ...form, mD7AttachmentName: attachmentName },
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
        D7_IssueData: JSON.stringify(updated),
      });

    // Callback to parent
    onSave?.({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD7AttachmentName: attachmentName },
      latestJson: updated
    });

    setEditOpen(false);

  } catch (err) {
    console.error("Save D7 failed", err);
    alert("Failed to save D7 data: " + (err?.message || err));
  }
};


  const Err = ({ id }: { id: string }) => (errors[id] ? <div className="field-error">{errors[id]}</div> : null);
 const derivedHistory: D7HistoryRow[] =
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
        {/* <div id="D7_Tab1" className="tab-pane in active"> */}
          <div
          id="D7_Tab1"
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
                <input readOnly className="form-control" value={form.vD7AssignTo || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Assign Date</label>
                <input readOnly className="form-control" value={form.vD7AssignDT || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Action Status</label>
                <input readOnly className="form-control" value={form.vD7ActionStatus || form.mD7ActionStatus || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">B. Analysis Details</div>

            <div className="row section-row">
              <div className="col-sm-12">
                <label className="form-label">Analysis Attachment</label>
                <input readOnly className="form-control" value={form.mD7AttachmentName || form.mD7AnalysisDetails || ""} />
                {/* preview example using the uploaded local image path (if you want to show preview) */}
                <div style={{ marginTop: 8 }}>
                  <img src={exampleLocalImage} alt="preview" style={{ maxWidth: 120, opacity: 0.85 }} />
                </div>
              </div>
            </div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Is Root Cause Found</label>
                <input readOnly className="form-control" value={form.mD7RootCauseFound || ""} />
              </div>

              <div className="col-sm-8">
                <label className="form-label">Remarks</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD7_Remarks || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">C. ICA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">ICA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD7ICA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">ICA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD7ICA_VIN || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">D. PCA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">PCA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD7PCA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">PCA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD7PCA_VIN || ""} />
              </div>
            </div>
          </div>
        </div>

        {/* History Tab */}
        {/* <div id="D7_Tab2" className="tab-pane fade"> */}
           <div
        id="D7_Tab2"
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
                <h4 className="modal-title">Diamond 7 - Issue Details</h4>
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
                      <input id="vD7AssignTo" className="form-control" value={form.vD7AssignTo || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Assign Date</label>
                      <input id="vD7AssignDT" className="form-control" value={form.vD7AssignDT || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4">
                      <label className="form-label">Action Status</label>
                      <select id="mD7ActionStatus" className="form-control" value={form.mD7ActionStatus || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD7ActionStatus" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">B. Analysis Details</div>
                  <div className="row">
                    <div className="col-sm-12">
                      <label className="form-label">Analysis Attachment</label>
                      <div className="upload-row">
                        <input id="D7AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                        <div className="uploaded-filename">{form.mD7AttachmentName ? form.mD7AttachmentName : <em>No file selected</em>}</div>
                      </div>
                      <Err id="mD7AttachmentName" />
                    </div>
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <div className="col-sm-4">
                      <label className="form-label">Is Root Cause Found?</label>
                      <select id="mD7RootCauseFound" className="form-control" value={form.mD7RootCauseFound || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD7RootCauseFound" />
                    </div>

                    <div className="col-sm-8">
                      <label className="form-label">Remarks</label>
                      <textarea id="mD7_Remarks" className="form-control" rows={2} value={form.mD7_Remarks || ""} onChange={handleChange} />
                      <Err id="mD7_Remarks" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">C. ICA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">ICA Details</label>
                      <textarea id="mD7ICA_Details" className="form-control" rows={2} value={form.mD7ICA_Details || ""} onChange={handleChange} />
                      <Err id="mD7ICA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">ICA VIN Cutoff</label>
                      <input id="mD7ICA_VIN" className="form-control" value={form.mD7ICA_VIN || ""} onChange={handleChange} />
                      <Err id="mD7ICA_VIN" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">D. PCA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">PCA Details</label>
                      <textarea id="mD7PCA_Details" className="form-control" rows={2} value={form.mD7PCA_Details || ""} onChange={handleChange} />
                      <Err id="mD7PCA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">PCA VIN Cutoff</label>
                      <input id="mD7PCA_VIN" className="form-control" value={form.mD7PCA_VIN || ""} onChange={handleChange} />
                      <Err id="mD7PCA_VIN" />
                    </div>
                  </div>
                </div>
              </div>

              <div className="modal-footer modal-footer-spaced">
                <button id="btnUpdate_D7Container" type="button" className="btn btn-primary" onClick={handleSave}>
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

export default Tab9D7;
