import * as React from 'react';
import './Ias.scss';
import type { IPrtsProps } from './IPrtsProps';
import InitiatorLanding from './Pages/InitiatorLandingPage';
import { NewRequest } from './Pages/NewRequest';
import { Dashboard } from './Pages/Dashboard';
import { MyReqDash } from './Pages/MyReqDashboard';
import { ApprovalForm } from './Pages/ApprovalForm';
import { AllReqDash } from './Pages/AllReqDashboard';
import { Print } from './Pages/PrintRequest';
import { Draft } from './Pages/Draft';
import { CostCenter } from './Pages/CostCenter';
import { MaterialUser } from './Pages/MaterialMaster';
import { WarehouseUser } from './Pages/WarehouseMaster';
import PartwiseReport from './Pages/PartwiseReport';
import ForwardingUser from './Pages/ForwardingUser';
import Sidebar from '../components/Pages/Sidebar';
// import { HashRouter as Router, Switch, Route } from 'react-router-dom';
import { BrowserRouter as Router, Switch, Route, HashRouter } from 'react-router-dom';
import ProblemResolutionTrackingSystem from './Pages/NewPage';
import BuildType from './Pages/Master Pages/BuildType';
import PRTSResource from './Pages/Master Pages/PRTSResource';
import Commodity from './Pages/Master Pages/Commodity';
import IssueCategory from './Pages/Master Pages/IssueCategory';
import MFGShopSelection from './Pages/Master Pages/MFGShopSelection';
import Approval from './Pages/Approval';

const Drr: React.FC<IPrtsProps> = (props) => {
  const { hasTeamsContext, EmployeeId, Maintenance } = props;

  const hasAccess = EmployeeId.length > 0;

  const readerAccessDenied = () => (
    <div
      className="d-flex justify-content-center align-items-center"
      style={{ height: "100vh" }}
    >
      <h3>No access. Please contact the IT team.</h3>
    </div>
  );


  if (props.Maintenance === true) {
    return (
      <div
        className="d-flex justify-content-center align-items-center"
        style={{ height: "100vh", width: "100vw", overflow: "hidden" }}
      >
        <img
          src="../SiteAssets/Custom/imgs/Maintenance.png"
          alt="Maintenance Mode"
          style={{
            width: "100%",
            height: "100%",
            objectFit: "cover"
          }}
        />
      </div>
    );
  }
  else if (!hasAccess) {
    return (
      <div className="d-flex justify-content-center align-items-center" style={{ height: '100vh' }}>
        <h3>No access. Please contact the IT team.</h3>
      </div>
    );
  }

  return (
    <div>
      <HashRouter>
        {/* <Router> */}
          <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>
            <Sidebar {...props} />
            <div className="main">
              <Switch>
                <Route
                  path="/InitiatorLanding"
                  render={() => 
                    props.Reader === true
                      ? readerAccessDenied()
                      : <ProblemResolutionTrackingSystem {...props} />
                  }
                />
                <Route
                  path="/InitiatorLandingedit/:RequestId"
                  render={() => 
                    props.Reader === true
                      ? readerAccessDenied()
                      : <Approval {...props} />
                  }
                />
                <Route
                  exact
                  path="/"
                  render={() => <InitiatorLanding {...props} />}
                />
                <Route
                  exact
                  path="/Dashboard"
                  render={() => <Dashboard {...props} />}
                />
                <Route
                  exact
                  path="/ApprovalForm"
                  render={() => <ApprovalForm {...props} />}
                />
                <Route
                  exact
                  path="/MyReqDash"
                  render={() => <MyReqDash {...props} />}
                />
                <Route
                  exact
                  path="/AllReqDash"
                  render={() => <AllReqDash {...props} />}
                />
                <Route
                  exact
                  path="/Draft"
                  render={() => <Draft {...props} />}
                />
                <Route
                  exact
                  path="/CostCenter"
                  render={() => <CostCenter {...props} />}
                />
                <Route
                  exact
                  path="/PartwiseReport"
                  render={() => <PartwiseReport {...props} />}
                />
                <Route
                  exact
                  path="/ForwardingUser"
                  render={() => <ForwardingUser {...props} />}
                />
                <Route
                  exact
                  path="/Print"
                  render={() => <Print {...props} />}
                />
                <Route
                  exact
                  path="/Material"
                  render={() => <MaterialUser {...props} />}
                />
                <Route
                  exact
                  path="/Warehouse"
                  render={() => <WarehouseUser {...props} />}
                />
                {/* Master Pages Routes */}
                <Route
                  exact
                  path="/BuildType"
                  render={() => <BuildType {...props} />}
                />
                <Route
                  exact
                  path="/PRTSResource"
                  render={() => <PRTSResource {...props} />}
                />
                <Route
                  exact
                  path="/Commodity"
                  render={() => <Commodity {...props} />}
                />
                <Route
                  exact
                  path="/IssueCategory"
                  render={() => <IssueCategory {...props} />}
                />
                <Route
                  exact
                  path="/MFGShop"
                  render={() => <MFGShopSelection {...props} />}
                />
              </Switch>
            </div>
          </div>
        {/* </Router> */}
      </HashRouter>
    </div>

  );
};

export default Drr;

