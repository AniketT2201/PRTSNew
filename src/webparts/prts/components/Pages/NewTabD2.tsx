import * as React from "react";
import { useEffect, useState } from "react";
import "./CSS/Tab3D2.scss"
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import { Form } from "formik";

// NOTE: uploaded design reference image (local path)
// Reference image: /mnt/data/0bfc68b3-a8ef-45b3-8e7f-ac05c21dcfa5.png

export interface D2Data {
  vD2AssignTo?: string;
  vD2AssignDT?: string;
  vD2ActionStatus?: string;

  mD2AnalysisDetails?: string;
  mD2RootCauseFound?: string;
  mD2ActionStatus?: string;
  mD2ICA_Details?: string;
  mD2ICA_VIN?: string;
  mD2PCA_Details?: string;
  mD2PCA_VIN?: string;
  mD2_Remarks?: string;
  mD2AttachmentName?: string;
}

interface D2HistoryRow {
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

interface Tab4D2Props {
  activeData?: D2Data;
  historyData?: D2HistoryRow[];
  existingJsonArray?: any[];
  onSave: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D2Data;
    latestJson: any[];
  }) => void;
}
interface D2SaveResult {
  updatedArray: any[];
  jsonString: string;
  chStatus: string;
  savedFields: D2Data;
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

const Tab4D2: React.FC<Tab4D2Props> = ({
  activeData,
  historyData = [],
  existingJsonArray,
  onSave,
}) => {
  // local form state (used for edit modal and as fallback display)
  const [form, setForm] = useState<D2Data>({
    mD2AnalysisDetails: "",
    mD2RootCauseFound: "",
    mD2ActionStatus: "",
    mD2ICA_Details: "",
    mD2ICA_VIN: "",
    mD2PCA_Details: "",
    mD2PCA_VIN: "",
    mD2_Remarks: "",
    mD2AttachmentName: "",
    ...activeData,
  });
  const [activeSubTab, setActiveSubTab] = useState<"Active" | "History">("Active");
  const [editOpen, setEditOpen] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [jsonArray, setJsonArray] = useState<any[]>(
    existingJsonArray && existingJsonArray.length > 0
      ? JSON.parse(JSON.stringify(existingJsonArray))
      : [defaultEmptyEntry()]
  );

  const { RequestId } = useParams<{ RequestId: string }>();

  // keep form in sync when parent activeData prop changes
  useEffect(() => {
    setForm((prev) => ({ ...prev, ...(activeData || {}) }));
  }, [activeData]);

  // keep jsonArray in sync if existingJsonArray prop changes
  useEffect(() => {
    if (existingJsonArray && existingJsonArray.length > 0) {
      setJsonArray(JSON.parse(JSON.stringify(existingJsonArray)));
    }
  }, [existingJsonArray]);

  // -------------------GET Function (loads from SharePoint when RequestId present) -------------------
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      const data = await getD2Data(RequestId);
      if (data) {
        setForm(data.activeData);
        setJsonArray(data.jsonArray && data.jsonArray.length ? data.jsonArray : [defaultEmptyEntry()]);
      }
    };

    load();
  }, [RequestId]);

  const getD2Data = async (ReqId: string) => {
    try {
      const item = await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(ReqId))
        .select("D2_IssueData")
        .get();

      let parsed: any[] = [];

      if (item && item.D2_IssueData) {
        try {
          parsed = JSON.parse(item.D2_IssueData);
        } catch (e) {
          console.error("JSON parse error for D2_IssueData", e);
          parsed = [];
        }
      }

      // last row holds current values normally
      const last = parsed && parsed.length ? parsed[parsed.length - 1] : defaultEmptyEntry();

      // Create display/active data from last JSON entry (assign fields usually c1/c2 in last element)
      const activeDataFromJson: D2Data = {
        vD2AssignTo: last.c1 || "",
        vD2AssignDT: last.c2 || "",
        vD2ActionStatus: last.c12 || "",

        mD2AnalysisDetails: last.c3 || "",
        mD2AttachmentName: last.c10 || "",
        mD2RootCauseFound: last.c4 || "",
        mD2ICA_Details: last.c5 || "",
        mD2ICA_VIN: last.c6 || "",
        mD2PCA_Details: last.c7 || "",
        mD2PCA_VIN: last.c8 || "",
        mD2_Remarks: last.c9 || "",
        mD2ActionStatus: last.c12 || "",
      };

      return { activeData: activeDataFromJson, jsonArray: parsed };
    } catch (err) {
      console.error("GET ERROR (D2):", err);
      return null;
    }
  };
  // __________________________________________________________________________

  const handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>
  ) => {
    const id = e.target.id;
    const val = e.target.value;
    setForm((p) => ({ ...p, [id]: val }));
    setErrors((s) => ({ ...s, [id]: "" }));
  };

  // file chooser — we store only filename (like Angular app did)
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files && e.target.files[0];
    if (f) {
      setForm((p) => ({ ...p, mD2AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD2AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  const validateForD2 = (): boolean => {
    const e: Record<string, string> = {};
    if (!isNotBlank(form.mD2AttachmentName) && !isNotBlank(form.mD2AnalysisDetails)) {
      e.mD2AttachmentName = "Analysis attachment is required (upload a file).";
    }
    if (!isNotBlank(form.mD2RootCauseFound)) {
      e.mD2RootCauseFound = "Please select whether root cause is found.";
    }
    if (!isNotBlank(form.mD2ActionStatus) || form.mD2ActionStatus === "-1") {
      e.mD2ActionStatus = "Please select Action Status.";
    }
    if (form.mD2RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD2ICA_Details)) e.mD2ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD2ICA_VIN)) e.mD2ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD2PCA_Details)) e.mD2PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD2PCA_VIN)) e.mD2PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD2_Remarks)) e.mD2_Remarks = "Remarks required when root cause is not found.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const computeCHStatus = (current: D2Data) => {
    let ActionStatus = current.mD2ActionStatus || "";
    if (ActionStatus === "-1") ActionStatus = "";
    let s = "1/6";
    if (current.mD2RootCauseFound === "Yes" && !isNotBlank(current.mD2PCA_Details)) {
      s = "2/6";
    } else if (current.mD2RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }
    return { s, ActionStatus };
  };

  const handleSave = () => {
    if (!validateForD2()) return;

    const analysisText = form.mD2AnalysisDetails || "";
    const attachmentName = form.mD2AttachmentName || "";

    let updated = [...jsonArray];

    // No rows? add new empty
    if (updated.length === 0) {
      updated.push(defaultEmptyEntry());
    }

    // Append NEW row if root cause FOUND (matches Angular)
    if (form.mD2RootCauseFound === "Yes") {
      updated.push(defaultEmptyEntry());
    }

    const idx = updated.length - 1;

    const rootCause = form.mD2RootCauseFound || "";
    const icaDetails = form.mD2ICA_Details || "";
    const icaVIN = form.mD2ICA_VIN || "";
    const pcaDetails = form.mD2PCA_Details || "";
    const pcaVIN = form.mD2PCA_VIN || "";
    const remarks = form.mD2_Remarks || "";

    const { s, ActionStatus } = computeCHStatus(form);

    const prevC1 = updated[0]?.c1 || "";
    const prevC2 = updated[0]?.c2 || "";
    const dt = formatDateForAngularLike(new Date());

    // Correct mapping (same as D1)
    updated[idx] = {
      c1: prevC1,
      c2: prevC2,
      c3: analysisText,        // analysis TEXT
      c4: rootCause,
      c5: icaDetails,
      c6: icaVIN,
      c7: pcaDetails,
      c8: pcaVIN,
      c9: remarks,
      c10: attachmentName,     // ONLY file name
      c11: dt,
      c12: ActionStatus,
    };

    setJsonArray(updated);
    setEditOpen(false);

    onSave({
      updatedArray: updated,
      jsonString: JSON.stringify(updated),
      chStatus: s,
      savedFields: { ...form, mD2AttachmentName: attachmentName },
      latestJson: updated
    });
  };


  const Err = ({ id }: { id: string }) =>
    errors[id] ? <div className="field-error">{errors[id]}</div> : null;

  // Use activeData prop if provided, otherwise fallback to locally-loaded form
  const display: D2Data = (activeData && Object.keys(activeData).length > 0) ? activeData : form;
const derivedHistory: D2HistoryRow[] =
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
        <div
          id="D2_Tab1"
          className={`tab-pane ${activeSubTab === "Active" ? "in active" : "fade"}`}
        >
          <div className="d-flex justify-content-end mb-3" style={{ textAlign: "right" }}>
            <button className="btn btn-primary" onClick={() => setEditOpen(true)}>
              Edit
            </button>
          </div>

          {/* READ-ONLY DISPLAY — use `display` so either prop or fetched data shows */}
          <div className="section-card">
            <div className="section-headerD2">A. Issue Details</div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Issue Assign To</label>
                <input readOnly className="form-control" value={form?.vD2AssignTo || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Assign Date</label>
                <input readOnly className="form-control" value={form?.vD2AssignDT || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">Action Status</label>
                <input readOnly className="form-control" value={form?.vD2ActionStatus || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">B. Analysis Details</div>

            <div className="row section-row">
              <div className="col-sm-12">
                <label className="form-label">Analysis Attachment</label>
                <input
                  readOnly
                  className="form-control"
                  value={form?.mD2AttachmentName || jsonArray[jsonArray.length - 1]?.c10 || ""}
                />
              </div>
            </div>

            <div className="row section-row">
              <div className="col-sm-4">
                <label className="form-label">Is Root Cause Found</label>
                <input readOnly className="form-control" value={form?.mD2RootCauseFound || ""} />
              </div>

              <div className="col-sm-8">
                <label className="form-label">Remarks</label>
                <textarea readOnly rows={2} className="form-control" value={form?.mD2_Remarks || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">C. ICA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">ICA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form?.mD2ICA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">ICA VIN Cutoff</label>
                <input readOnly className="form-control" value={form?.mD2ICA_VIN || ""} />
              </div>
            </div>
          </div>

          <div className="section-card">
            <div className="section-headerD2">D. PCA Action Details</div>

            <div className="row section-row">
              <div className="col-sm-8">
                <label className="form-label">PCA Details</label>
                <textarea readOnly rows={2} className="form-control" value={form?.mD2PCA_Details || ""} />
              </div>

              <div className="col-sm-4">
                <label className="form-label">PCA VIN Cutoff</label>
                <input readOnly className="form-control" value={form?.mD2PCA_VIN || ""} />
              </div>
            </div>
          </div>
        </div>

        {/* History Tab */}
<div
  id="D2_Tab2"
  className={`tab-pane ${activeSubTab === "History" ? "in active" : "fade"}`}
>
          <div className="marginTop10">
            <table id="D2_HistoryTable" className="table table-striped">
              <thead>
                <tr>
                  <th>#</th><th>Agency</th><th>UserName</th><th>Analysis Details</th><th>Found Root Cause</th><th>ICA Action</th><th>ICA VIN Cut OFF</th><th>PCA Action</th><th>PCA VIN CUT OFF</th><th>Remarks</th><th>Attachment</th><th>Action DateTime</th>
                </tr>
              </thead>
              <tbody>
                {derivedHistory && derivedHistory.length ? (
                  derivedHistory.map((r, idx) => (
                    <tr key={idx}>
                      <td>{idx + 1}</td>
                      <td>{r.agency}</td>
                      <td>{r.userName}</td>
                      <td>{r.analysisDetails}</td>
                      <td>{r.foundRootCause}</td>
                      <td>{r.icaAction}</td>
                      <td>{r.icaVINCutOff}</td>
                      <td>{r.pcaAction}</td>
                      <td>{r.pcaVINCutOff}</td>
                      <td>{r.remarks}</td>
                      <td>{r.attachment}</td>
                      <td>{r.actionDateTime}</td>
                    </tr>
                  ))
                ) : (
                  <tr><td colSpan={12} style={{ textAlign: "center" }}>No history available</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* EDIT MODAL (styled to match screenshot; uses native file chooser) */}
      {editOpen && (
        <div className="modal fade show d-block modal-overlay" tabIndex={-1} role="dialog" aria-modal="true">
          <div className="modal-dialog modal-lg" role="document">
            <div className="modal-content modal-card">
              <div className="modal-header modal-header-red">
                <h4 className="modal-title">Diamond 2 - Issue Details</h4>
                <button type="button" className="close" onClick={() => setEditOpen(false)}>&times;</button>
              </div>

              <div className="modal-body">
                {/* A. Issue Details (compact in modal) */}
                <div className="section-card small">
                  <div className="section-headerD2 small">A. Issue Details</div>
                  <div className="row">
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Issue Assign To</label>
                      <input id="vD2AssignTo" className="form-control" value={form.vD2AssignTo || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4" style={{ display: "none" }}>
                      <label className="form-label">Assign Date</label>
                      <input id="vD2AssignDT" className="form-control" value={form.vD2AssignDT || ""} onChange={handleChange} />
                    </div>
                    <div className="col-sm-4">
                      <label className="form-label">Action Status</label>
                      <select id="mD2ActionStatus" className="form-control" value={form.mD2ActionStatus || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD2ActionStatus" />
                    </div>
                  </div>
                </div>

                {/* B. Analysis Details */}
                <div className="section-card small">
                  <div className="section-headerD2 small">B. Analysis Details</div>
                  <div className="row">
                    <div className="col-sm-12">
                      <label className="form-label">Analysis Attachment</label>
                      <div className="upload-row">
                        <input id="D2AnalysisAttachmentFile" type="file" onChange={handleFileChange} />
                        <div className="uploaded-filename">{form.mD2AttachmentName ? form.mD2AttachmentName : <em>No file selected</em>}</div>
                      </div>
                      <Err id="mD2AttachmentName" />
                    </div>
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <div className="col-sm-4">
                      <label className="form-label">Is Root Cause Found?</label>
                      <select id="mD2RootCauseFound" className="form-control" value={form.mD2RootCauseFound || ""} onChange={handleChange}>
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD2RootCauseFound" />
                    </div>

                    <div className="col-sm-8">
                      <label className="form-label">Remarks</label>
                      <textarea id="mD2_Remarks" className="form-control" rows={2} value={form.mD2_Remarks || ""} onChange={handleChange} />
                      <Err id="mD2_Remarks" />
                    </div>
                  </div>
                </div>

                {/* C. ICA */}
                <div className="section-card small">
                  <div className="section-headerD2 small">C. ICA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">ICA Details</label>
                      <textarea id="mD2ICA_Details" className="form-control" rows={2} value={form.mD2ICA_Details || ""} onChange={handleChange} />
                      <Err id="mD2ICA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">ICA VIN Cutoff</label>
                      <input type="date" id="mD2ICA_VIN" className="form-control" value={form.mD2ICA_VIN || ""} onChange={handleChange} />
                      <Err id="mD2ICA_VIN" />
                    </div>
                  </div>
                </div>

                {/* D. PCA */}
                <div className="section-card small">
                  <div className="section-headerD2 small">D. PCA Action Details</div>
                  <div className="row">
                    <div className="col-sm-8">
                      <label className="form-label">PCA Details</label>
                      <textarea id="mD2PCA_Details" className="form-control" rows={2} value={form.mD2PCA_Details || ""} onChange={handleChange} />
                      <Err id="mD2PCA_Details" />
                    </div>

                    <div className="col-sm-4">
                      <label className="form-label">PCA VIN Cutoff</label>
                      <input type="date" id="mD2PCA_VIN" className="form-control" value={form.mD2PCA_VIN || ""} onChange={handleChange} />
                      <Err id="mD2PCA_VIN" />
                    </div>
                  </div>
                </div>
              </div>

              <div className="modal-footer modal-footer-spaced">
                <button id="btnUpdate_D2Container" type="button" className="btn btn-primary" onClick={handleSave}>
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

export default Tab4D2;
