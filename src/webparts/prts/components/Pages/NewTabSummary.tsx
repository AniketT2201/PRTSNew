import * as React from "react";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useParams } from "react-router-dom";
import { formatDate } from "./MaterialMaster";

interface SummaryRow {
  c1: string; // Initiator/Approver
  c2: string; // Forwarded To
  c3: string; // Action Date
  c4: string; // Action
  c5: string; // Action Remarks
}

interface Tab10SummaryProps {
  summaryData?: SummaryRow[];  // optional, parent can pass array
}

const Tab10Summary: React.FC<Tab10SummaryProps> = ({ summaryData }) => {
  const { RequestId } = useParams<{ RequestId: string }>();
  const [rows, setRows] = useState<SummaryRow[]>([]);

  useEffect(() => {
    // If parent provides summary data → use it
    if (summaryData && summaryData.length > 0) {
      setRows(summaryData);
      return;
    }

    // Otherwise fetch from SharePoint PRTSList
    if (!RequestId) return;

    const load = async () => {
      try {
        const item = await sp.web.lists
          .getByTitle("PRTSList")
          .items.getById(Number(RequestId))
          .select("Summary")
          .get();

        let arr: SummaryRow[] = [];

        if (item.Summary) {
          try {
            arr = JSON.parse(item.Summary);
          } catch (e) {
            console.error("Error parsing Summary JSON", e);
            arr = [];
          }
        }

        setRows(arr || []);
      } catch (err) {
        console.error("Failed loading Summary", err);
        setRows([]);
      }
    };

    load();
  }, [RequestId, summaryData]);

  return (
    <div id="Tab10" className="tab-pane in active fade-margin">
      <div style={{ textAlign: "left", marginTop: 10 }}>
        <table
          id="eSummaryDataTable"
          className="table table-bordered"
          style={{ width: "100%", tableLayout: "fixed" }}
        >
          <colgroup>
            <col style={{ width: "15%" }} />
            <col style={{ width: "15%" }} />
            <col style={{ width: "10%" }} />
            <col style={{ width: "20%" }} />
            <col style={{ width: "40%" }} />
          </colgroup>

          <thead>
            <tr>
              <th>Initiator/Approver</th>
              <th>Forwarded To</th>
              <th>Action Date</th>
              <th>Action</th>
              <th>Action Remarks</th>
            </tr>
          </thead>

          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={5} style={{ textAlign: "center" }}>
                  No summary records found
                </td>
              </tr>
            ) : (
              rows.map((r, index) => (
                <tr key={index}>
                  <td>{r.c1}</td>
                  <td>{r.c2}</td>
                  <td>{formatDate(r.c3)}</td>
                  <td>{r.c4}</td>
                  <td>{r.c5}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default Tab10Summary;
