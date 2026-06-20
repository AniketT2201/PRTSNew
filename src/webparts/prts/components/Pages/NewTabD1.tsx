import * as React from "react";
import { useEffect, useState } from "react";
import "./CSS/Tab3D1.scss";
import { useParams } from "react-router-dom";
import { sp } from "@pnp/sp/presets/all";

interface D1Data {
  vD1AssignTo?: string;
  vD1AssignDT?: string;
  vD1ActionStatus?: string;

  mD1AnalysisDetails?: string;
  mD1RootCauseFound?: string;
  mD1ActionStatus?: string;
  mD1ICA_Details?: string;
  mD1ICA_VIN?: string;
  mD1PCA_Details?: string;
  mD1PCA_VIN?: string;
  mD1_Remarks?: string;
  mD1AttachmentName?: string;
}

interface D1HistoryRow {
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

interface Tab3D1Props {
  historyData?: D1HistoryRow[];

  onSave: (result: {
    updatedArray: any[];
    jsonString: string;
    chStatus: string;
    savedFields: D1Data;
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

const Tab3D1D1: React.FC<Tab3D1Props> = ({ historyData = [], onSave }) => {
  const { RequestId } = useParams<{ RequestId: string }>();

  const [form, setForm] = useState<D1Data>({});
  const [editOpen, setEditOpen] = useState(false);
  const [activeSubTab, setActiveSubTab] = useState<"Active" | "History">("Active");
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [jsonArray, setJsonArray] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  // ▬▬▬▬▬▬▬▬▬▬▬ GET DATA FROM SHAREPOINT ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
  useEffect(() => {
    if (!RequestId) return;

    const load = async () => {
      const data = await getD1Data(RequestId);

      if (data) {
        setForm(data.activeData);
        setJsonArray(data.jsonArray);
      }
    };

    load();
  }, [RequestId]);

  const getD1Data = async (ReqId: string) => {
    try {
      const item = await sp.web.lists
        .getByTitle("PRTSList")
        .items.getById(Number(ReqId))
        .select("D1_IssueData")
        .get();

      let jsonArray: any[] = [];

      if (item.D1_IssueData) {
        try {
          jsonArray = JSON.parse(item.D1_IssueData);
        } catch (e) {
          console.error("JSON parse error", e);
        }
      }

      const last = jsonArray.length ? jsonArray[jsonArray.length - 1] : {};

      const activeData: D1Data = {
        vD1AssignTo: last.c1 || "",
        vD1AssignDT: last.c2 || "",
        vD1ActionStatus: last.c12 || "",

        mD1AnalysisDetails: last.c3 || "",
        mD1AttachmentName: last.c10 || "",
        mD1RootCauseFound: last.c4 || "",
        mD1ICA_Details: last.c5 || "",
        mD1ICA_VIN: last.c6 || "",
        mD1PCA_Details: last.c7 || "",
        mD1PCA_VIN: last.c8 || "",
        mD1_Remarks: last.c9 || "",
        mD1ActionStatus: last.c12 || "",
      };

      return { activeData, jsonArray };
    } catch (err) {
      console.error("GET ERROR:", err);
      return null;
    }
  };

  // ▬▬▬▬▬▬▬▬▬▬▬ FORM CHANGE HANDLERS ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
  const handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>
  ) => {
    const id = e.target.id;
    const val = e.target.value;
    setForm((p) => ({ ...p, [id]: val }));
    setErrors((s) => ({ ...s, [id]: "" }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files && e.target.files[0];
    setSelectedFiles(f ? [f] : []);
    if (f) {
      setForm((p) => ({ ...p, mD1AttachmentName: f.name }));
      setErrors((s) => ({ ...s, mD1AttachmentName: "" }));
    }
  };

  const isNotBlank = (v?: string) => !!v && v.toString().trim().length > 0;

  // ▬▬▬▬▬▬▬▬▬▬▬ VALIDATION ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
  const validateForD1 = (): boolean => {
    const e: Record<string, string> = {};

    if (!isNotBlank(form.mD1AttachmentName) && !isNotBlank(form.mD1AnalysisDetails)) {
      e.mD1AttachmentName = "Analysis attachment is required.";
    }

    if (!isNotBlank(form.mD1RootCauseFound)) {
      e.mD1RootCauseFound = "Please select root cause status.";
    }

    if (!isNotBlank(form.mD1ActionStatus)) {
      e.mD1ActionStatus = "Please select action status.";
    }

    if (form.mD1RootCauseFound === "Yes") {
      if (!isNotBlank(form.mD1ICA_Details)) e.mD1ICA_Details = "ICA details required.";
      if (!isNotBlank(form.mD1ICA_VIN)) e.mD1ICA_VIN = "ICA VIN required.";
      if (!isNotBlank(form.mD1PCA_Details)) e.mD1PCA_Details = "PCA details required.";
      if (!isNotBlank(form.mD1PCA_VIN)) e.mD1PCA_VIN = "PCA VIN required.";
    } else {
      if (!isNotBlank(form.mD1_Remarks)) e.mD1_Remarks = "Remarks required.";
    }

    setErrors(e);
    return Object.keys(e).length === 0;
  };

  // ▬▬▬▬▬▬▬▬▬▬▬ CH STATUS LOGIC ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
  const computeCHStatus = (current: D1Data) => {
    let ActionStatus = current.mD1ActionStatus || "";
    let s = "1/6";

    if (current.mD1RootCauseFound === "Yes" && !isNotBlank(current.mD1PCA_Details)) {
      s = "2/6";
    } else if (current.mD1RootCauseFound === "Yes") {
      if (ActionStatus === "Will_Implemented") s = "3/6";
      else if (ActionStatus === "Implemented") s = "4/6";
    }

    return { s, ActionStatus };
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

  // ▬▬▬▬▬▬▬▬▬▬▬ SAVE ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
 const handleSave = async () => {
  if (!validateForD1()) return;

  const analysisText = form.mD1AnalysisDetails || "";
  const attachmentName = form.mD1AttachmentName || "";

  let updated = [...jsonArray];

  // Create default row if no data exists
  if (updated.length === 0) updated.push(defaultEmptyEntry());

  // APPEND NEW ENTRY WHEN ROOT CAUSE FOUND = Yes
  if (form.mD1RootCauseFound === "Yes") {
    updated.push(defaultEmptyEntry());
  }
        let current: any[] = Array.isArray(jsonArray) ? JSON.parse(JSON.stringify(jsonArray)) : [];
  const idx = updated.length - 1;

  const { s, ActionStatus } = computeCHStatus(form);
      const lastIndex = current.length - 1;
        const existingAttachment = lastIndex >= 0 && isNotBlank(current[lastIndex]?.c11) ? current[lastIndex].c11 : "";
const uploadedUrls =
  selectedFiles.length > 0
    ? await uploadAttachmentsToList(RequestId, selectedFiles)
    : [];
    const attachmentHtml = buildAttachmentHTMLWithLinks(
  existingAttachment,
  uploadedUrls
);
  updated[idx] = {
    c1: jsonArray[0]?.c1 || "",
    c2: jsonArray[0].c2 || "",
    c3: attachmentHtml,
    c4: form.mD1RootCauseFound || "",
    c5: form.mD1ICA_Details || "",
    c6: form.mD1ICA_VIN || "",
    c7: form.mD1PCA_Details || "",
    c8: form.mD1PCA_VIN || "",
    c9: form.mD1_Remarks || "",
    c10: attachmentHtml,
    c11: formatDateForAngularLike(new Date()),
    c12: ActionStatus,
  };

  setJsonArray(updated);

  onSave({
    updatedArray: updated,
    jsonString: JSON.stringify(updated),
    chStatus: s,
    savedFields: form,
    latestJson: updated
  });

  setEditOpen(false);
};


  const Err = ({ id }: { id: string }) =>
    errors[id] ? <div style={{ color: "red" }}>{errors[id]}</div> : null;

  // ▬▬▬▬▬▬▬▬▬▬▬ UI ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
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
        {/* ACTIVE TAB */}
        <div
          id="D1_Tab1"
          className={`tab-pane ${activeSubTab === "Active" ? "in active" : "fade"}`}
        >          <div className="d-flex justify-content-end mb-3">
            <button className="btn btn-primary" onClick={() => setEditOpen(true)}>
              Edit
            </button>
          </div>

          {/* SECTION A */}
          <div className="section-card">
            <div className="section-headerD1">A. Issue Details</div>

            <div className="row g-3">
              <div className="col-md-4">
                <label>Issue Assign To</label>
                <input className="form-control" readOnly value={form?.vD1AssignTo || ""} />
              </div>

              <div className="col-md-4">
                <label>Assign Date</label>
                <input className="form-control" readOnly value={form?.vD1AssignDT || ""} />
              </div>

              <div className="col-md-4">
                <label>Action Status</label>
                <input className="form-control" readOnly value={form?.vD1ActionStatus || ""} />
              </div>
            </div>
          </div>

          {/* SECTION B */}
          <div className="section-card">
            <div className="section-headerD1">B. Analysis Details</div>

            <div className="row g-3">
              <div className="col-md-12">
                <label>Analysis Attachment</label>
                <div className="form-control">
                  {jsonArray?.[jsonArray.length - 1]?.c3 || "-"}
                </div>
              </div>

              <div className="col-md-4">
                <label>Is Root Cause Found</label>
                <input className="form-control" readOnly value={form?.mD1RootCauseFound || ""} />
              </div>

              <div className="col-md-8">
                <label>Remarks</label>
                <textarea className="form-control" readOnly value={form?.mD1_Remarks || ""} />
              </div>
            </div>
          </div>

          {/* SECTION C */}
          <div className="section-card">
            <div className="section-headerD1">C. ICA Action Details</div>

            <div className="row g-3">
              <div className="col-md-8">
                <label>ICA Details</label>
                <textarea className="form-control" readOnly value={form?.mD1ICA_Details || ""} />
              </div>

              <div className="col-md-4">
                <label>ICA VIN Cutoff</label>
                <input className="form-control" readOnly value={form?.mD1ICA_VIN || ""} />
              </div>
            </div>
          </div>

          {/* SECTION D */}
          <div className="section-card">
            <div className="section-headerD1">D. PCA Action Details</div>

            <div className="row g-3">
              <div className="col-md-8">
                <label>PCA Details</label>
                <textarea className="form-control" readOnly value={form?.mD1PCA_Details || ""} />
              </div>

              <div className="col-md-4">
                <label>PCA VIN Cutoff</label>
                <input className="form-control" readOnly value={form?.mD1PCA_VIN || ""} />
              </div>
            </div>
          </div>

          {/* SECTION E */}
          <div className="section-card">
            <div className="section-headerD1">E. Attachments</div>

            <div className="col-md-6">
              <label>Uploaded File</label>
              <div className="form-control">
                {jsonArray?.[jsonArray.length - 1]?.c10 || "-"}
              </div>
            </div>
          </div>
        </div>

        {/* HISTORY TAB */}
        <div
          id="D1_Tab2"
          className={`tab-pane ${activeSubTab === "History" ? "in active" : "fade"}`}
        >          <div className="card p-3 shadow-sm mt-3">
            <table className="table table-striped table-bordered">
              <thead className="table-dark">
                <tr>
                  <th>#</th>
                  <th>Agency</th>
                  <th>User</th>
                  <th>Analysis</th>
                  <th>Root Cause</th>
                  <th>ICA</th>
                  <th>ICA VIN</th>
                  <th>PCA</th>
                  <th>PCA VIN</th>
                  <th>Remarks</th>
                  <th>Attachment</th>
                  <th>Date</th>
                </tr>
              </thead>

              <tbody>
  {jsonArray.length ? (
    jsonArray.map((r, i) => (
      <tr key={i}>
        <td>{i + 1}</td>
        <td>{r.c1}</td>
        <td>{r.c2}</td>
        <td>{r.c3}</td>
        <td>{r.c4}</td>
        <td>{r.c5}</td>
        <td>{r.c6}</td>
        <td>{r.c7}</td>
        <td>{r.c8}</td>
        <td>{r.c9}</td>

        {/* ✅ Attachment */}
        <td
          dangerouslySetInnerHTML={{
            __html: r.c11 || "-"
          }}
        />

        <td>{r.c12}</td>
      </tr>
    ))
  ) : (
    <tr>
      <td colSpan={12} className="text-center">
        No history available
      </td>
    </tr>
  )}
</tbody>

            </table>
          </div>
        </div>
      </div>

      {/* EDIT MODAL */}
      {editOpen && (
        <div className="modal fade show d-block" style={{ backgroundColor: "rgba(0,0,0,0.45)" }}>
          <div className="modal-dialog modal-lg">
            <div className="modal-content">

              <div className="modal-header">
                <h4 className="modal-title">Diamond 1 - Edit Issue Details</h4>
                <button type="button" className="close" onClick={() => setEditOpen(false)}>
                  &times;
                </button>
              </div>

              <div className="modal-body">
                {/* Analysis */}
                <div className="section-card mb-3">
                  <div className="section-headerD1">Analysis Details</div>

                  <div className="mb-3">
                    <label>Analysis Attachment *</label>
                    <input
                      id="D1AnalysisAttachmentFile"
                      type="file"
                      className="form-control"
                      onChange={handleFileChange}
                    />
                    {form.mD1AttachmentName && (
                      <div className="mt-2 text-success">
                        Uploaded: <strong>{form.mD1AttachmentName}</strong>
                      </div>
                    )}
                    <Err id="mD1AttachmentName" />
                  </div>
                </div>

                {/* ROOT CAUSE */}
                <div className="section-card mb-3">
                  <div className="section-headerD1">Root Cause & Action</div>

                  <div className="row g-3">
                    <div className="col-md-4">
                      <label>Is Root Cause Found *</label>
                      <select
                        id="mD1RootCauseFound"
                        className="form-control"
                        value={form.mD1RootCauseFound || ""}
                        onChange={handleChange}
                      >
                        <option value=""></option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                      </select>
                      <Err id="mD1RootCauseFound" />
                    </div>

                    <div className="col-md-4">
                      <label>Action Status *</label>
                      <select
                        id="mD1ActionStatus"
                        className="form-control"
                        value={form.mD1ActionStatus || ""}
                        onChange={handleChange}
                      >
                        <option value=""></option>
                        <option value="Will_Implemented">Action Will Be Implemented</option>
                        <option value="Implemented">Action Implemented</option>
                      </select>
                      <Err id="mD1ActionStatus" />
                    </div>
                  </div>
                </div>

                {/* ICA */}
                <div className="section-card mb-3">
                  <div className="section-headerD1">ICA Action Implementation</div>

                  <div className="row g-3">
                    <div className="col-md-8">
                      <label>ICA Details</label>
                      <input
                        id="mD1ICA_Details"
                        className="form-control"
                        value={form.mD1ICA_Details || ""}
                        onChange={handleChange}
                      />
                      <Err id="mD1ICA_Details" />
                    </div>

                    <div className="col-md-4">
                      <label>VIN Cut Off</label>
                      <input
                        id="mD1ICA_VIN"
                        type="date"
                        className="form-control"
                        value={form.mD1ICA_VIN || ""}
                        onChange={handleChange}
                      />
                      <Err id="mD1ICA_VIN" />
                    </div>
                  </div>
                </div>

                {/* PCA */}
                <div className="section-card mb-3">
                  <div className="section-headerD1">PCA Action</div>

                  <div className="row g-3">
                    <div className="col-md-8">
                      <label>PCA Details</label>
                      <input
                        id="mD1PCA_Details"
                        className="form-control"
                        value={form.mD1PCA_Details || ""}
                        onChange={handleChange}
                      />
                      <Err id="mD1PCA_Details" />
                    </div>

                    <div className="col-md-4">
                      <label>VIN Cut Off</label>
                      <input
                        id="mD1PCA_VIN"
                        type="date"
                        className="form-control"
                        value={form.mD1PCA_VIN || ""}
                        onChange={handleChange}
                      />
                      <Err id="mD1PCA_VIN" />
                    </div>
                  </div>
                </div>

                {/* Remarks */}
                <div className="section-card mb-3">
                  <div className="section-headerD1">Remarks</div>

                  <textarea
                    id="mD1_Remarks"
                    className="form-control"
                    rows={3}
                    value={form.mD1_Remarks || ""}
                    onChange={handleChange}
                  />
                  <Err id="mD1_Remarks" />
                </div>
              </div>

              <div className="modal-footer">
                <button className="btn btn-success" onClick={handleSave}>
                  Update
                </button>
                <button className="btn btn-secondary" onClick={() => setEditOpen(false)}>
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

export default Tab3D1D1;
