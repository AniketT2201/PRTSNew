import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SPCRUDOPS from '../../service/DAL/spcrudops';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface AssignToAnotherUserModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSubmit: (assignment: {
    agency: string;
     
    user: string;
    remarks: string;
  }) => void;
  context: WebPartContext;   // For PeoplePicker
  parentProps: any;          // <<< IMPORTANT for SPCRUDOPS
}

const AssignToAnotherUserModal: React.FC<AssignToAnotherUserModalProps> = ({
  isOpen,
  onClose,
  onSubmit,
  context,
  parentProps
}) => {

  const [agency, setAgency] = useState('');
  const [agenciesList, setAgenciesList] = useState<any[]>([]);
  const [status, setStatus] = useState('');
  const [selectedUser, setSelectedUser] = useState<any>(null);
  const [remarks, setRemarks] = useState('');

  /** Load agencies when modal opens */
  useEffect(() => {
    if (isOpen) {
      GetAgencyData();
    }
  }, [isOpen]);

  /** Fetch Agency List */
  async function GetAgencyData() {
    try {
      const spCrudOps = await SPCRUDOPS();

      const Modeldata = await spCrudOps.getRootData(
        "PRTS_NonTechnical_Issue_Agencys",
        "Title,ID",
        "",
        "",
        { column: "ID", isAscending: true },
        parentProps              // <<< FIXED: correct object passed
      );

      const uniqueModels = Array.from(
        new Map(Modeldata.map((item: any) => [item.Title, item])).values()
      );

      const options = uniqueModels.map((item: any) => ({
        key: item.ID,
        text: item.Title
      }));

      setAgenciesList(options);
      setAgency(options[0]?.text || "");

    } catch (err) {
      console.error("Error loading agency data:", err);
    }
  }

  /** Submit Form */
  const handleSubmit = () => {
    if (!agency  || !selectedUser || !remarks.trim()) {
      alert("Please fill in all required fields including user and remarks.");
      return;
    }

    onSubmit({
      agency,
      user: selectedUser?.text || "",
      remarks
    });
  };

  if (!isOpen) return null;

  return (
    <div className="modal fade show d-block" style={{ backgroundColor: "rgba(0,0,0,0.5)" }}>
      <div className="modal-dialog">
        <div className="modal-content">

          <div className="modal-header">
            <button type="button" className="close" onClick={onClose}>&times;</button>
            <h4 className="modal-title">Issue Assign To Another User</h4>
          </div>

          <div className="modal-body">

            {/* Agency Dropdown */}
            <div className="row top-buffer">
              <div className="col-sm-5">
                <label><span className="required">*</span>Agency</label>
                <select className="form-control" value={agency}
                  onChange={(e) => setAgency(e.target.value)}>
                  {agenciesList.map((item) => (
                    <option key={item.key} value={item.text}>{item.text}</option>
                  ))}
                </select>
              </div>

              {/* Status Dropdown */}
              {/* <div className="col-sm-5">
                <label><span className="required">*</span>Change Status</label>
                <select className="form-control" value={status}
                  onChange={(e) => setStatus(e.target.value)}>
                  <option value=""></option>
                  <option value="1/6">1/6</option>
                  <option value="2/6">2/6</option>
                  <option value="3/6">3/6</option>
                  <option value="4/6">4/6</option>
                </select>
              </div> */}
            </div>

            {/* People Picker */}
            <div className="row top-buffer">
              <div className="col-sm-12">
                <label><span className="required">*</span>Select User</label>

                <PeoplePicker
                  context={parentProps.context}
                  titleText="Assign To"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={true}
                  onChange={(items) => setSelectedUser(items[0] || null)}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={500}
                />
              </div>
            </div>

            {/* Remarks */}
            <div className="row top-buffer">
              <div className="col-sm-12">
                <label><span className="required">*</span>Remarks</label>
                <textarea
                  className="form-control"
                  rows={3}
                  maxLength={250}
                  value={remarks}
                  onChange={(e) => setRemarks(e.target.value)}
                />
              </div>
            </div>

          </div>

          {/* Footer */}
          <div className="modal-footer">
            <button type="button" className="btn btn-default" onClick={handleSubmit}>Update</button>
            <button type="button" className="btn btn-default" onClick={onClose}>Close</button>
          </div>

        </div>
      </div>
    </div>
  );
};

export default AssignToAnotherUserModal;
