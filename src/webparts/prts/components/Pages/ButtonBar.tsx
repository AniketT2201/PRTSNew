import * as React from 'react';

interface ButtonBarProps {
  buttons: any;

  onClose: () => void;
  // onCreateDraft: () => void;
  onSubmitRequest: (reqId: any) => void | Promise<void>;
  onWithdrawn: (code: number, msg: string) => void;
  onPrint: () => void;
  onCloseIssue: (code: number, msg: string) => void;
  onReturnBackToPITMember: (code: number, msg: string) => void;
  onForwardToNextPITMember: (code: number, msg: string) => void;
  onProcessWithIssueClose: () => void;
  onProcessWithIssueOpen: () => void;
  onSendBackToPreviousStage: (code: number, msg: string) => void;
  onAssignToChampion: () => void;
  onProcessWithIssueCloseNT: () => void;
  onBackToInitiator: (code: number, msg: string) => void;
  onReassignIssue: () => void;
  onSubmitForReview: () => void;
  onRejectClick: () => void;
  onWrongIssueAssign: () => void;
onClickRework: () => void;
  status: string;
  chStatus: string;
}

const ButtonBar: React.FC<ButtonBarProps> = ({
  buttons,
  onClose,
  // onCreateDraft,
  onSubmitRequest,
  onWithdrawn,
  onPrint,
  onCloseIssue,
  onReturnBackToPITMember,
  onForwardToNextPITMember,
  onProcessWithIssueClose,
  onProcessWithIssueOpen,
  onSendBackToPreviousStage,
  onAssignToChampion,
  onProcessWithIssueCloseNT,
  onBackToInitiator,
  onReassignIssue,
  onSubmitForReview,
  onRejectClick,
  onWrongIssueAssign,
  onClickRework,
  status,
  chStatus,
}) => {

  return (
    <div className="btn-Row" style={{paddingBottom:'15px'}}>
      <div className="mainbttns" id="btnBar">

        {buttons.btnClose && (
          <a className="btn btn-default btnss" id="btnClose" onClick={onClose}>Back</a>
        )}

        {/* {buttons.btnCreateDraft && (
          <a className="btn btn-default btnssone" id="btnCreateDraft" onClick={onCreateDraft}>Create Draft</a>
        )} */}

        {buttons.btnSubmit && (
          <a className="btn btn-default btnNonDraft btnsstwo" id="btnSubmit"
            onClick={onSubmitRequest}>Submit</a>
        )}

        {buttons.btnWithDrawn && (
          <a className="btn btn-default btnNonDraft btnssthree" id="btnWithDrawn"
             onClick={() => onWithdrawn(4, "To Withdrawn")}>
            Withdrawn
          </a>
        )}

        {buttons.btnPrint && (
          <a className="btn btn-default btnNonDraft btnssfour" id="btnPrint" onClick={onPrint}>Print</a>
        )}

        {buttons.btnCloseIssue && (
          <a className="btn btn-default btnNonDraft btnssfive" id="btnCloseIssue"
             onClick={() => onCloseIssue(6, "on Issue Closing")}>
            Close Issue
          </a>
        )}

        {buttons.btnReturnBackToPITMember && (
          <a className="btn btn-default btnNonDraft btnsssix" id="btnReturnBackToPITMember"
             onClick={() => onReturnBackToPITMember(1, "For Sending Back To PIT Member")}>
            Return Back to PIT Member
          </a>
        )}

        {buttons.btnForwardToNextPITMember && (
          <a className="btn btn-default btnNonDraft btnssseven" id="btnForwardToNextPITMember"
             onClick={() => onForwardToNextPITMember(2, "For Sending To Next PIT Member")}>
            Forward to next PIT Member
          </a>
        )}

        {buttons.btnProcessWithIssueCloseD && (
          <a className="btn btn-default btnNonDraft btnsseight" id="btnProcessWithIssueCloseD"
             onClick={onProcessWithIssueClose}>
            Process with Issue Close
          </a>
        )}

        {buttons.btnProcessWithIssueOpen && (
          <a className="btn btn-default btnNonDraft btnssnine" id="btnProcessWithIssueOpen"
             onClick={onProcessWithIssueOpen}>
            Process with Issue Open
          </a>
        )}

        {buttons.btnSendBackToPreviousStage && (
          <a className="btn btn-default btnNonDraft btnssten" id="btnSendBackToPreviousStage"
             onClick={() => onSendBackToPreviousStage(3, "For Sending Back to Previous Member")}>
            Send Back to Previous Stage
          </a>
        )}

        {buttons.btnForwardAtD7 && (
          <a className="btn btn-default btnNonDraft btnsseleven" id="btnForwardAtD7"
             onClick={onAssignToChampion}>
            Assign To Champion
          </a>
        )}

        {buttons.btnProcessWithIssueCloseNT && (
          <a className="btn btn-default btnNonDraft btnsstwelve" id="btnProcessWithIssueCloseNT"
             onClick={onProcessWithIssueCloseNT}>
            Process with Issue Close
          </a>
        )}

        {buttons.btnBackToInitiator && (
          <a className="btn btn-default btnNonDraft btnssthirteen" id="btnBackToIntiator"
             onClick={() => onBackToInitiator(5, "For Sending Back to Initiator")}>
            Send Back to Initiator (Issue Open)
          </a>
        )}

        {buttons.btnAssignIssueToAnoterUser && (
          <a className="btn btn-default btnNonDraft btnssfourteen" id="btnAssignIssueToAnoterUser"
             onClick={onReassignIssue}>
            Re-Assign Issue
          </a>
        )}
   {buttons.btnsubmitforreview && (
          <a className="btn btn-default btnNonDraft btnssfourteen" id="btnsubmitforreview"
             onClick={onSubmitForReview}>
            Submit for Review
          </a>
        )}
        {buttons.btnReject && (
          <a className="btn btn-default btnNonDraft btnssfourteen" id="btnReject"
             onClick={onRejectClick}>
            Reject 
          </a>
        )}
         {buttons.btnWrongIssueAssign && (
          <a className="btn btn-default btnNonDraft btnssfourteen" id="btnWrongIssueAssign"
             onClick={onWrongIssueAssign}>
            Wrong Issue Assigned 
          </a>
        )}
         {buttons.btnclickRework && (
          <a className="btn btn-default btnNonDraft btnssfourteen" id="btnclickRework"
             onClick={onClickRework}>
            Rework
          </a>
        )}
      </div>
      <div className="request">
        <span>Status: {status}</span>
         <span className="badge badge-light" style={{color:'black'}}>{chStatus}</span>
      </div>

      {/* <div className="col-sm-1">
        <span className="badge badge-light" style={{color:'black'}}>{chStatus}</span>
      </div> */}
    </div>
  );
};

export default ButtonBar;
