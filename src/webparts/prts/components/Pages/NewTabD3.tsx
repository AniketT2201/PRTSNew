// Tab5D3.tsx
import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import "./CSS/Tab3D2.scss";

export interface D3Data {
  vD3AssignTo?: string;
  vD3AssignDT?: string;
  vD3ActionStatus?: string;

  mD3AnalysisDetails?: string;
  mD3RootCauseFound?: string;
  mD3ActionStatus?: string;
  mD3ICA_Details?: string;
  mD3ICA_VIN?: string;
  mD3PCA_Details?: string;
  mD3PCA_VIN?: string;
  mD3_Remarks?: string;
  mD3AttachmentName?: string;
}

interface D3HistoryRow {
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

interface Tab5D3Props {
  activeData?: D3Data;
  historyData?: D3HistoryRow[];
  existingJsonArray?: any[]; // optional preloaded array
  onSave?: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D3Data;
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

const Tab5D3: React.FC<Tab5D3Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  // `form` is single source of truth for UI
  const [form, setForm] = useState<D3Data>({
    mD3AnalysisDetails: "",
    mD3RootCauseFound: "",
    mD3ActionStatus: "",
    mD3ICA_Details: "",
    mD3ICA_VIN: "",
    mD3PCA_Details: "",
    mD3PCA_VIN: "",
    mD3_Remarks: "",
    mD3AttachmentName: "",
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

  // keep form in sync when parent activeData changes
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
        vD3AssignTo: last.c1 || prev.vD3AssignTo,
        vD3AssignDT: last.c2 || prev.vD3AssignDT,
        vD3ActionStatus: last.c12 || prev.vD3ActionStatus,
        mD3AnalysisDetails: last.c3 || prev.mD3AnalysisDetails,
        mD3AttachmentName: last.c10 || prev.mD3AttachmentName,
        mD3RootCauseFound: last.c4 || prev.mD3RootCauseFound,
        mD3ICA_Details: last.c5 || prev.mD3ICA_Details,
        mD3ICA_VIN: last.c6 || prev.mD3ICA_VIN,
        mD3PCA_Details: last.c7 || prev.mD3PCA_Details,
        mD3PCA_VIN: last.c8 || prev.mD3PCA_VIN,
        mD3_Remarks: last.c9 || prev.mD3_Remarks,
        mD3ActionStatus: last.c12 || prev.mD3ActionStatus,
      }));
    }
  }, [existingJsonArray]);

  // Fetch from SP when RequestId present
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      try {
        const item = await sp.web.lists.getByTitle("PRTSList").items.getById(Number(RequestId)).select("D3_IssueData", "CH_Status").get();
        let parsed: any[] = [];
        if (item && item.D3_IssueData) {
          try {
            parsed = JSON.parse(item.D3_IssueData);
          } catch (e) {
            console.warn("D3_IssueData JSON parse failed", e);
            parsed = [];
          }
        }

        if (parsed.length) {
          setJsonArray(parsed);
          const last = parsed[parsed.length - 1] || {};
          setForm((prev) => ({
            ...prev,
            vD3AssignTo: last.c1 || prev.vD3AssignTo,
            vD3AssignDT: last.c2 || prev.vD3AssignDT,
            vD3ActionStatus: last.c12 || prev.vD3ActionStatus,
            mD3AnalysisDetails: last.c3 || prev.mD3AnalysisDetails,
            mD3AttachmentName: last.c10 || prev.mD3AttachmentName,
            mD3RootCauseFound: last.c4 || prev.mD3RootCauseFound,
            mD3ICA_Details: last.c5 || prev.mD3ICA_Details,
            mD3ICA_VIN: last.c6 || prev.mD3ICA_VIN,
            mD3PCA_Details: last.c7 || prev.mD3PCA_Details,
            mD3PCA_VIN: last.c8 || prev.mD3PCA_VIN,
            mD3_Remarks: last.c9 || prev.mD3_Remarks,
            mD3ActionStatus: last.c12 || prev.mD3ActionStatus,
          }));
        }
      } catch (err) {
        console.error("Error fetching D3 data", err);
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
      setForm((p) => ({ ...p, mD3AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD3AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  // Validation equivalent to your Angular isValid_IssueDetails for D3
  const isValid_IssueDetails_D3 = () => {
    const e: Record<string, string> = {};
    // Either attachment OR analysis details (keeps parity with your D2/D3 requirement)
    if (!isNotBlank(form.mD3AttachmentName) && !isNotBlank(form.mD3AnalysisDetails)) {
      e.mD3AttachmentName = "Analysis attachment is required (upload file) or enter analysis details.";
    }

    if (!isNotBlank(form.mD3RootCauseFound)) {
      e.mD3RootCauseFound = "Please select whether root cause is found.";
    }

    if (!isNotBlank(form.mD3ActionStatus) || form.mD3ActionStatus === "-1") {
      e.mD3ActionStatus = "Please select Action Status.";
    }

    if (form.mD3RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD3ICA_Details)) e.mD3ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD3ICA_VIN)) e.mD3ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD3PCA_Details)) e.mD3PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD3PCA_VIN)) e.mD3PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD3_Remarks)) e.mD3_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  // Compute CH_Status like in Angular Save_Container
  const computeCHStatus = (current: D3Data) => {
    let ActionStatus = current.mD3ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD3RootCauseFound === "Yes" && !isNotBlank(current.mD3PCA_Details)) {
      s = "2/6";
    } else if (current.mD3RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

  // Save function equivalent to Save_Container('D3')
  const handleSave = async () => {
    if (!isValid_IssueDetails_D3()) return;

    let updated = [...jsonArray];

    // If no rows exist, add one
    if (updated.length === 0) {
      updated.push(defaultEmptyEntry());
    }

    // If root cause found = Yes → NEW ROW MUST BE ADDED
    if (form.mD3RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    const analysisText = form.mD3AnalysisDetails || "";
    const attachmentName = form.mD3AttachmentName || "";

    const rootCause = form.mD3RootCauseFound || "";
    const icaDetails = form.mD3ICA_Details || "";
    const icaVIN = form.mD3ICA_VIN || "";
    const pcaDetails = form.mD3PCA_Details || "";
    const pcaVIN = form.mD3PCA_VIN || "";
    const remarks = form.mD3_Remarks || "";

    // Angular equivalent calculation
    const { s, ActionStatus } = computeCHStatus(form);

    // Preserve original assigned user & date
    const prevC1 = updated[0]?.c1 || "";
    const prevC2 = updated[0]?.c2 || "";

    const dt = formatDateForAngularLike(new Date());

    // Correct mapping — matches Angular perfectly
    updated[idx] = {
      c1: prevC1,
      c2: prevC2,
      c3: analysisText,      // Analysis text only
      c4: rootCause,
      c5: icaDetails,
      c6: icaVIN,
      c7: pcaDetails,
      c8: pcaVIN,
      c9: remarks,
      c10: attachmentName,   // Only file name
      c11: dt,
      c12: ActionStatus,
    };

    setJsonArray(updated);
    setForm((p) => ({ ...p, mD3AttachmentName: attachmentName }));

    // Save to SP only if RequestId exists
    if (RequestId) {
      await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(RequestId))
        .update({
          CH_Status: s,
          D3_IssueData: JSON.stringify(updated),
        });
    }

    onSave?.({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD3AttachmentName: attachmentName },
      latestJson: updated,
    });

    setEditOpen(false);
  };


  const Err = ({ id }: { id: string }) => (errors[id] ? <div className="field-error">{errors[id]}</div> : null);
  const derivedHistory: D3HistoryRow[] =
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
        {/* <div id="D3_Tab1" className="tab-pane in active"> */}
        <div
          id="D3_Tab1"
          className={`tab-pane ${activeSubTab === "Active" ? "in active" : "fade"}`}
        >
        <div className="d-flex justify-content-end mb-3" style={{ textAlign: "right" }}>
          <button className="btn btn-primary" onClick={() => setEditOpen(true)} style={{ marginRight: 20, marginTop: 3 }}>
            Edit
          </button>
        </div>

        {/* READ-ONLY DISPLAY (using `form`) */}
        <div className="section-card">
          <div className="section-headerD2">A. Issue Details</div>

          <div className="row section-row">
            <div className="col-sm-4">
              <label className="form-label">Issue Assign To</label>
              <input readOnly className="form-control" value={form.vD3AssignTo || ""} />
            </div>

            <div className="col-sm-4">
              <label className="form-label">Assign Date</label>
              <input readOnly className="form-control" value={form.vD3AssignDT || ""} />
            </div>

            <div className="col-sm-4">
              <label className="form-label">Action Status</label>
              <input readOnly className="form-control" value={form.vD3ActionStatus || form.mD3ActionStatus || ""} />
            </div>
          </div>
        </div>

        <div className="section-card">
          <div className="section-headerD2">B. Analysis Details</div>

          <div className="row section-row">
            <div className="col-sm-12">
              <label className="form-label">Analysis Attachment</label>
              <input readOnly className="form-control" value={form.mD3AttachmentName || form.mD3AnalysisDetails || ""} />
            </div>
          </div>

          <div className="row section-row">
            <div className="col-sm-4">
              <label className="form-label">Is Root Cause Found</label>
              <input readOnly className="form-control" value={form.mD3RootCauseFound || ""} />
            </div>

            <div className="col-sm-8">
              <label className="form-label">Remarks</label>
              <textarea readOnly rows={2} className="form-control" value={form.mD3_Remarks || ""} />
            </div>
          </div>
        </div>

        <div className="section-card">
          <div className="section-headerD2">C. ICA Action Details</div>

          <div className="row section-row">
            <div className="col-sm-8">
              <label className="form-label">ICA Details</label>
              <textarea readOnly rows={2} className="form-control" value={form.mD3ICA_Details || ""} />
            </div>

            <div className="col-sm-4">
              <label className="form-label">ICA VIN Cutoff</label>
              <input readOnly className="form-control" value={form.mD3ICA_VIN || ""} />
            </div>
          </div>
        </div>

        <div className="section-card">
          <div className="section-headerD2">D. PCA Action Details</div>

          <div className="row section-row">
            <div className="col-sm-8">
              <label className="form-label">PCA Details</label>
              <textarea readOnly rows={2} className="form-control" value={form.mD3PCA_Details || ""} />
            </div>

            <div className="col-sm-4">
              <label className="form-label">PCA VIN Cutoff</label>
              <input readOnly className="form-control" value={form.mD3PCA_VIN || ""} />
            </div>
          </div>
        </div>
      </div>

      {/* History Tab uses historyData prop */}
      {/* <div id="D3_Tab2" className="tab-pane fade"> */}
      <div
        id="D3_Tab2"
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
                <th>PCA VIN Cut OFF</th>
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
    </div >

      {/* Edit modal */ }
  {
    editOpen && (
      <div className="modal fade show d-block modal-overlay" tabIndex={-1} role="dialog" aria-modal="true">
        <div className="modal-dialog modal-lg" role="document">
          <div className="modal-content modal-card">
            <div className="modal-header modal-header-red">
              <h4 className="modal-title">Diamond 3 - Issue Details</h4>
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
                    <input id="vD3AssignTo" className="form-control" value={form.vD3AssignTo || ""} onChange={handleChange} />
                  </div>
                  <div className="col-sm-4" style={{ display: "none" }}>
                    <label className="form-label">Assign Date</label>
                    <input id="vD3AssignDT" className="form-control" value={form.vD3AssignDT || ""} onChange={handleChange} />
                  </div>
                  <div className="col-sm-4">
                    <label className="form-label">Action Status</label>
                    <select id="mD3ActionStatus" className="form-control" value={form.mD3ActionStatus || ""} onChange={handleChange}>
                      <option value=""></option>
                      <option value="Will_Implemented">Action Will Be Implemented</option>
                      <option value="Implemented">Action Implemented</option>
                    </select>
                    <Err id="mD3ActionStatus" />
                  </div>
                </div>
              </div>

              <div className="section-card small">
                <div className="section-headerD2 small">B. Analysis Details</div>
                <div className="row">
                  <div className="col-sm-12">
                    <label className="form-label">Analysis Attachment</label>
                    <div className="upload-row">
                      <input id="D3AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                      <div className="uploaded-filename">{form.mD3AttachmentName ? form.mD3AttachmentName : <em>No file selected</em>}</div>
                    </div>
                    <Err id="mD3AttachmentName" />
                  </div>
                </div>

                <div className="row" style={{ marginTop: 12 }}>
                  <div className="col-sm-4">
                    <label className="form-label">Is Root Cause Found?</label>
                    <select id="mD3RootCauseFound" className="form-control" value={form.mD3RootCauseFound || ""} onChange={handleChange}>
                      <option value=""></option>
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>
                    </select>
                    <Err id="mD3RootCauseFound" />
                  </div>

                  <div className="col-sm-8">
                    <label className="form-label">Remarks</label>
                    <textarea id="mD3_Remarks" className="form-control" rows={2} value={form.mD3_Remarks || ""} onChange={handleChange} />
                    <Err id="mD3_Remarks" />
                  </div>
                </div>
              </div>

              <div className="section-card small">
                <div className="section-headerD2 small">C. ICA Action Details</div>
                <div className="row">
                  <div className="col-sm-8">
                    <label className="form-label">ICA Details</label>
                    <textarea id="mD3ICA_Details" className="form-control" rows={2} value={form.mD3ICA_Details || ""} onChange={handleChange} />
                    <Err id="mD3ICA_Details" />
                  </div>

                  <div className="col-sm-4">
                    <label className="form-label">ICA VIN Cutoff</label>
                    <input type="date" id="mD3ICA_VIN" className="form-control" value={form.mD3ICA_VIN || ""} onChange={handleChange} />
                    <Err id="mD3ICA_VIN" />
                  </div>
                </div>
              </div>

              <div className="section-card small">
                <div className="section-headerD2 small">D. PCA Action Details</div>
                <div className="row">
                  <div className="col-sm-8">
                    <label className="form-label">PCA Details</label>
                    <textarea id="mD3PCA_Details" className="form-control" rows={2} value={form.mD3PCA_Details || ""} onChange={handleChange} />
                    <Err id="mD3PCA_Details" />
                  </div>

                  <div className="col-sm-4">
                    <label className="form-label">PCA VIN Cutoff</label>
                    <input type="date" id="mD3PCA_VIN" className="form-control" value={form.mD3PCA_VIN || ""} onChange={handleChange} />
                    <Err id="mD3PCA_VIN" />
                  </div>
                </div>
              </div>
            </div>

            <div className="modal-footer modal-footer-spaced">
              <button id="btnUpdate_D3Container" type="button" className="btn btn-primary" onClick={handleSave}>
                Update
              </button>
              <button type="button" className="btn btn-default" onClick={() => setEditOpen(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
    )
  }
    </>
  );
};

export default Tab5D3;
