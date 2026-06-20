import * as React from 'react';
import { useState } from 'react';

interface RemarksModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUpdate: (remarks: string) => void;
  remarksTitle?: string;
}

const RemarksModal: React.FC<RemarksModalProps> = ({ isOpen, onClose, onUpdate, remarksTitle }) => {
  const [remarks, setRemarks] = useState('');

  if (!isOpen) return null; // Modal hidden if not open



    function handleUpdate(): void {
        onUpdate(remarks);
    }

  return (
    <div
      className="modal fade show d-block"
      role="dialog"
      aria-modal="true"
      tabIndex={-1}
      style={{ backgroundColor: 'rgba(0,0,0,0.5)' }}
    >
      <div className="modal-dialog" role="document">
        <div className="modal-content">
          <div className="modal-header">
            <button type="button" className="close" onClick={onClose} aria-label="Close">
              &times;
            </button>
            <h4 className="modal-title">Remarks</h4>
          </div>
          <div className="modal-body">
            <div className="row top-buffer">
              <div className="col-sm-12">
                <label htmlFor="mCommonRemarks">
                  <span className="required">*</span> Remarks - <span>{remarksTitle}</span>
                </label>
                <textarea
                  id="mCommonRemarks"
                  maxLength={250}
                  rows={3}
                  className="form-control restriceTabAndDoubleString"
                  value={remarks}
                  onChange={(e) => setRemarks(e.target.value)}
                />
              </div>
            </div>
          </div>
          <div className="modal-footer">
            <div>
              <button id="btnRemarks" type="button" className="btn btn-default" onClick={handleUpdate}>
                Update
              </button>
              <button type="button" className="btn btn-default" onClick={onClose}>
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default RemarksModal;
