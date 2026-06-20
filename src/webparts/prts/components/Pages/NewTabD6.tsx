// Tab8D6.tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import "./CSS/Tab3D2.scss"; // reuse same styles

export interface D6Data {
  vD6AssignTo?: string;
  vD6AssignDT?: string;
  vD6ActionStatus?: string;

  mD6AnalysisDetails?: string;
  mD6RootCauseFound?: string;
  mD6ActionStatus?: string;
  mD6ICA_Details?: string;
  mD6ICA_VIN?: string;
  mD6PCA_Details?: string;
  mD6PCA_VIN?: string;
  mD6_Remarks?: string;
  mD6AttachmentName?: string;
}

interface D6HistoryRow {
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

interface Tab8D6Props {
  activeData?: D6Data;
  historyData?: D6HistoryRow[];
  existingJsonArray?: any[]; // optional preloaded array
  onSave?: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D6Data;
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

const Tab8D6: React.FC<Tab8D6Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  const [form, setForm] = useState<D6Data>({
    mD6AnalysisDetails: "",
    mD6RootCauseFound: "",
    mD6ActionStatus: "",
    mD6ICA_Details: "",
    mD6ICA_VIN: "",
    mD6PCA_Details: "",
    mD6PCA_VIN: "",
    mD6_Remarks: "",
    mD6AttachmentName: "",
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

  // sync when parent activeData changes
  useEffect(() => {
    if (activeData) setForm((p) => ({ ...p, ...activeData }));
  }, [activeData]);

  // derive from existingJsonArray if parent passed it
  useEffect(() => {
    if (existingJsonArray && existingJsonArray.length) {
      const last = existingJsonArray[existingJsonArray.length - 1] || {};
      setJsonArray(JSON.parse(JSON.stringify(existingJsonArray)));
      setForm((prev) => ({
        ...prev,
        vD6AssignTo: last.c1 || prev.vD6AssignTo,
        vD6AssignDT: last.c2 || prev.vD6AssignDT,
        vD6ActionStatus: last.c12 || prev.vD6ActionStatus,
        mD6AnalysisDetails: last.c3 || prev.mD6AnalysisDetails,
        mD6AttachmentName: last.c10 || prev.mD6AttachmentName,
        mD6RootCauseFound: last.c4 || prev.mD6RootCauseFound,
        mD6ICA_Details: last.c5 || prev.mD6ICA_Details,
        mD6ICA_VIN: last.c6 || prev.mD6ICA_VIN,
        mD6PCA_Details: last.c7 || prev.mD6PCA_Details,
        mD6PCA_VIN: last.c8 || prev.mD6PCA_VIN,
        mD6_Remarks: last.c9 || prev.mD6_Remarks,
        mD6ActionStatus: last.c12 || prev.mD6ActionStatus,
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
          .select("D6_IssueData", "CH_Status")
          .get();

        let parsed: any[] = [];
        if (item && item.D6_IssueData) {
          try {
            parsed = JSON.parse(item.D6_IssueData);
          } catch (e) {
            console.warn("D6_IssueData JSON parse failed", e);
            parsed = [];
          }
        }

        if (parsed.length) {
          setJsonArray(parsed);
          const last = parsed[parsed.length - 1] || {};
          setForm((prev) => ({
            ...prev,
            vD6AssignTo: last.c1 || prev.vD6AssignTo,
            vD6AssignDT: last.c2 || prev.vD6AssignDT,
            vD6ActionStatus: last.c12 || prev.vD6ActionStatus,
            mD6AnalysisDetails: last.c3 || prev.mD6AnalysisDetails,
            mD6AttachmentName: last.c10 || prev.mD6AttachmentName,
            mD6RootCauseFound: last.c4 || prev.mD6RootCauseFound,
            mD6ICA_Details: last.c5 || prev.mD6ICA_Details,
            mD6ICA_VIN: last.c6 || prev.mD6ICA_VIN,
            mD6PCA_Details: last.c7 || prev.mD6PCA_Details,
            mD6PCA_VIN: last.c8 || prev.mD6PCA_VIN,
            mD6_Remarks: last.c9 || prev.mD6_Remarks,
            mD6ActionStatus: last.c12 || prev.mD6ActionStatus,
          }));
        }
      } catch (err) {
        console.error("Error fetching D6 data", err);
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
      // NOTE: this only stores file name in JSON like other tabs.
      // If you want to upload to SP library, add upload logic here.
      setForm((p) => ({ ...p, mD6AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD6AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  // Validation (either attachment OR analysis required)
  const isValid_IssueDetails_D6 = () => {
    const e: Record<string, string> = {};

    if (!isNotBlank(form.mD6AttachmentName) && !isNotBlank(form.mD6AnalysisDetails)) {
      e.mD6AttachmentName = "Analysis attachment is required (upload file) or enter analysis details.";
    }

    if (!isNotBlank(form.mD6RootCauseFound)) {
      e.mD6RootCauseFound = "Please select whether root cause is found.";
    }

    if (!isNotBlank(form.mD6ActionStatus) || form.mD6ActionStatus === "-1") {
      e.mD6ActionStatus = "Please select Action Status.";
    }

    if (form.mD6RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD6ICA_Details)) e.mD6ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD6ICA_VIN)) e.mD6ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD6PCA_Details)) e.mD6PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD6PCA_VIN)) e.mD6PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD6_Remarks)) e.mD6_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  // Compute CH_Status
  const computeCHStatus = (current: D6Data) => {
    let ActionStatus = current.mD6ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD6RootCauseFound === "Yes" && !isNotBlank(current.mD6PCA_Details)) {
      s = "2/6";
    } else if (current.mD6RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

 const handleSave = async () => {
  if (!isValid_IssueDetails_D6()) return;

  try {
    let updated = [...jsonArray];

    // Ensure at least one row exists
    if (updated.length === 0) updated.push(defaultEmptyEntry());

    // Angular Logic:
    // If RootCauseFound = "Yes" → ALWAYS PUSH NEW EMPTY ROW
    if (form.mD6RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    // Correct Angular JSON rule:
    const analysisText = form.mD6AnalysisDetails || "";
    const attachmentName = form.mD6AttachmentName || "";

    const rootCause = form.mD6RootCauseFound || "";
    const icaDetails = form.mD6ICA_Details || "";
    const icaVIN = form.mD6ICA_VIN || "";
    const pcaDetails = form.mD6PCA_Details || "";
    const pcaVIN = form.mD6PCA_VIN || "";
    const remarks = form.mD6_Remarks || "";

    const { s, ActionStatus } = computeCHStatus(form);

    const prevC1 = updated[idx]?.c1 || "";
    const prevC2 = updated[idx]?.c2 || "";

    const dt = formatDateForAngularLike(new Date());

    // FINAL & CORRECTED JSON structure
    updated[idx] = {
      c1: prevC1,
      c2: prevC2,
      c3: analysisText,     // Text only
      c4: rootCause,
      c5: icaDetails,
      c6: icaVIN,
      c7: pcaDetails,
      c8: pcaVIN,
      c9: remarks,
      c10: attachmentName,  // File only
      c11: dt,
      c12: ActionStatus,
    };

    setJsonArray(updated);
    setForm((p) => ({ ...p, mD6AttachmentName: attachmentName }));

    // If no RequestId → only local save
    if (!RequestId) {
      onSave?.({
        updatedArray: updated,
        jsonString: JSON.stringify(updated),
        chStatus: s,
        savedFields: { ...form, mD6AttachmentName: attachmentName },
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
        D6_IssueData: JSON.stringify(updated),
      });

    onSave?.({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD6AttachmentName: attachmentName },
      latestJson: updated
    });

    setEditOpen(false);

  } catch (err) {
    console.error("Save D6 failed", err);
    alert("Failed to save D6 data: " + (err?.message || err));
  }
};


  const Err = ({ id }: { id: string }) => (errors[id] ? <div className="field-error">{errors[id]}</div> : null);
 const derivedHistory: D6HistoryRow[] =
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
        {/* <div id="D6_Tab1" className="tab-pane in active"> */}
          <div
          id="D6_Tab1"
          className={`tab-pane ${activeSubTab === "Active" ? "in active" : "fade"}`}
        >
          <div  className="d-flex justify-content-end mb-3" style={{ textAlign: "right" }}>
            <button
              className="btn btn-primary"
              onClick={() => setEditOpen(true)}
              style={{ marginRight: 20, marginTop: 3 }}
            >
              Edit
            </button>
          </div>

          <div className="section-card">
            <div className="section-headerD2">A. Issue Details</div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Issue Assign To</label>
                <input readOnly className="form-control" value={form.vD6AssignTo || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Assign Date</label>
                <input readOnly className="form-control" value={form.vD6AssignDT || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Action Status</label>
                <input readOnly className="form-control" value={form.vD6ActionStatus || form.mD6ActionStatus || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">B. Analysis Details</div>

            <div className="row section-row">
              <div className="col-sm-12">
                <label className="form-label">Analysis Attachment</label>
                <input readOnly className="form-control" value={form.mD6AttachmentName || form.mD6AnalysisDetails || ""} />
              </div>
            </div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Is Root Cause Found</label>
                <input readOnly className="form-control" value={form.mD6RootCauseFound || ""} />
              </div>

              <div className="col-sm-8">
                <label className="form-label">Remarks</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD6_Remarks || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">C. ICA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">ICA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD6ICA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">ICA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD6ICA_VIN || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">D. PCA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">PCA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form.mD6PCA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">PCA VIN Cutoff</label>
                <input readOnly className="form-control" value={form.mD6PCA_VIN || ""} />
              </div>
            </div>
          </div>
        </div>

        {/* History Tab */}
        {/* <div id="D6_Tab2" className="tab-pane fade"> */}
           <div
        id="D6_Tab2"
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
                <h4 className="modal-title">Diamond 6 - Issue Details</h4>
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
                      <input id="vD6AssignTo" className="form-control" value={form.vD6AssignTo || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Assign Date</label>
                      <input id="vD6AssignDT" className="form-control" value={form.vD6AssignDT || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4">
                      <label className="form-label">Action Status</label>
                      <select id="mD6ActionStatus" className="form-control" value={form.mD6ActionStatus || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD6ActionStatus" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">B. Analysis Details</div>
                  <div className="row">
                    <div className="col-sm-12">
                      <label className="form-label">Analysis Attachment</label>
                      <div className="upload-row">
                        <input id="D6AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                        <div className="uploaded-filename">{form.mD6AttachmentName ? form.mD6AttachmentName : <em>No file selected</em>}</div>
                      </div>
                      <Err id="mD6AttachmentName" />
                    </div>
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <div className="col-sm-4">
                      <label className="form-label">Is Root Cause Found?</label>
                      <select id="mD6RootCauseFound" className="form-control" value={form.mD6RootCauseFound || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD6RootCauseFound" />
                    </div>

                    <div className="col-sm-8">
                      <label className="form-label">Remarks</label>
                      <textarea id="mD6_Remarks" className="form-control" rows={2} value={form.mD6_Remarks || ""} onChange={handleChange} />
                      <Err id="mD6_Remarks" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">C. ICA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">ICA Details</label>
                      <textarea id="mD6ICA_Details" className="form-control" rows={2} value={form.mD6ICA_Details || ""} onChange={handleChange} />
                      <Err id="mD6ICA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">ICA VIN Cutoff</label>
                      <input id="mD6ICA_VIN" className="form-control" value={form.mD6ICA_VIN || ""} onChange={handleChange} />
                      <Err id="mD6ICA_VIN" />
                    </div>
                  </div>
                </div>

                <div className="section-card small">
                  <div className="section-headerD2 small">D. PCA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">PCA Details</label>
                      <textarea id="mD6PCA_Details" className="form-control" rows={2} value={form.mD6PCA_Details || ""} onChange={handleChange} />
                      <Err id="mD6PCA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">PCA VIN Cutoff</label>
                      <input id="mD6PCA_VIN" className="form-control" value={form.mD6PCA_VIN || ""} onChange={handleChange} />
                      <Err id="mD6PCA_VIN" />
                    </div>
                  </div>
                </div>
              </div>

              <div className="modal-footer modal-footer-spaced">
                <button id="btnUpdate_D6Container" type="button" className="btn btn-primary" onClick={handleSave}>
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

export default Tab8D6;
