import * as React from 'react';
import { useHistory, useLocation } from 'react-router-dom';
import '../Pages/CSS/Sidebar.scss';
import type { IPrtsProps } from '../IPrtsProps';
import '@fortawesome/fontawesome-free/css/all.min.css';
import SPCRUDOPS from '../../service/DAL/spcrudops';
import logo from '../../assets/MG-Motor-Logo.png'
type LocationState = {
  from?: string;
};

const Sidebar = (props: IPrtsProps) => {
  const history = useHistory();
  const location = useLocation();

  const [IASACL, setIASACL] = React.useState<any[]>([]);
  const [AppAdmin, setAppAdmin] = React.useState(false);
  const [Admin, setAdmin] = React.useState(false);
  const [Editor, setEditor] = React.useState(false);

  const username = props.userDisplayName;

  const activeByOverride = React.useMemo(() => {
    if (location.pathname.startsWith('/ApprovalForm')) {
      const state = location.state as LocationState | undefined;
      const memoryState = state?.from || sessionStorage.getItem('sidebarFrom');
      return memoryState;
    }
    return null;
  }, [location]);

  React.useEffect(() => {
    if (location.pathname.startsWith('/ApprovalForm')) {
      const state = location.state as LocationState | undefined;
      if (state?.from) {
        sessionStorage.setItem('sidebarFrom', state.from);
      }
    } else {
      sessionStorage.removeItem('sidebarFrom');
    }
    GetIAS_ACL();
  }, [location]);

async function GetIAS_ACL() {
  const spCrudOps = await SPCRUDOPS();

  const Momentflowdata = await spCrudOps.getData(
    'PRTSACL',
    'ID,UserName/Title,UserName/EMail,Role,EmployeeID,Title',
    'UserName',
    '',
    { column: 'ID', isAscending: true },
    props
  );

  const Userfiltereddata = Momentflowdata.filter(
    (item) =>
      item.UserName?.EMail?.trim().toLowerCase() === props.userEmail?.trim().toLowerCase() &&
      item.EmployeeID === props.EmployeeId[0].EmployeeID
  );

  const roles = Userfiltereddata.map(x => x.Role);
  const titles = Userfiltereddata.map(x => x.Title);

  setEditor(roles.includes('Editor'));
  setAdmin(titles.includes('SysAdmin'));
  setAppAdmin(titles.includes('AppAdmin'));

  setIASACL(Userfiltereddata);
}

  const getActiveClass = (key: string) => {
    const currentPath = location.pathname;

    if (activeByOverride) {
      return activeByOverride === key ? 'active' : '';
    }

    switch (key) {
      case '/InitiatorLanding':
        return currentPath === '/InitiatorLanding' ? 'active' : '';
      case '/':
        return currentPath === '/' ? 'active' : '';
      case '/Action':
        return ['/MyReqDash', '/Dashboard'].some((path) => currentPath.includes(path)) ? 'active' : '';
      case '/Report':
        return ['/PartwiseReport', '/AllReqDash'].some((path) => currentPath.includes(path)) ? 'active' : '';
      case '/Configure':
        return ['/Warehouse', '/Material'].some((path) => currentPath.includes(path)) ? 'active' : '';  
      default:
        return '';
    }
  };

  return (
    <div className="sidebar">
      <div className="sidehead">
        <img
          src={logo}
          alt="MG Motor Logo"
        />
        <h2 className="logo">JSW MGI</h2>
      </div>

      <div className="sidehead-user">
        <div>
          <i className="fas fa-user" style={{ marginLeft: '20px' }}></i>&nbsp;
          {username}
        </div>
      </div>

      <ul className="nav">
        <li className="nav-item">
          <a
            className={`nav-link ${getActiveClass('/')}`}
            onClick={() => history.push('/')}
          >
            <i className="fas fa-home" style={{ marginRight: '8px' }}></i>Home
          </a>
        </li>

        {Editor && (
          <li className="nav-item">
            <a
              className={`nav-link ${getActiveClass('/InitiatorLanding')}`}
              onClick={() => history.push('/InitiatorLanding')}
            >
              <i className="fas fa-plus-circle" style={{ marginRight: '8px' }}></i>New Request
            </a>
          </li>
        )}  

        <li className={`nav-item has-submenu ${getActiveClass('/Action')}`}>
          <div className="nav-link">
            <i className="fa fa-bolt" aria-hidden="true"></i> Action
          </div>
          <ul className="sub-menu">
            <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/Dashboard')}`}
                onClick={() => history.push('/Dashboard')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>Pending Action
              </a>
            </li>
            <li>
              <a
                className={`nav-link ${getActiveClass('/MyReqDash')}`}
                onClick={() => history.push('/MyReqDash')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>My Request
              </a>
            </li>
          </ul>
        </li>

        <li className={`nav-item has-submenu ${getActiveClass('/Report')}`}>
          <div className="nav-link">
            <i className="fa fa-chart-bar" aria-hidden="true"></i> Report
          </div>
          <ul className="sub-menu">
            {/* <li>
              <a
                className={`nav-link ${getActiveClass('/PartwiseReport')}`}
                onClick={() => history.push('/PartwiseReport')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>By Part Item
              </a>
            </li> */}
            <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/AllReqDash')}`}
                onClick={() => history.push('/AllReqDash')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>All Request
              </a>
            </li>
          </ul>
        </li>

        {Admin && (
          <>
          <li className="nav-item has-submenu settings">
            <div className="nav-link">
              <i className="fa fa-gears" aria-hidden="true"></i> Admin
            </div>
            <ul className="sub-menu">
              <li><span className="sub-menu-title">&nbsp;Application List</span></li> 
              <li>
                <a
                  href="../../PRTS/Lists/PRTSACL/AllItems.aspx"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                   <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>  ACL
                </a>
              </li>
              <li>
                <a
                  href="../../PRTS/Lists/PRTSList/AllItems.aspx"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                    <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i> PRTS List
                </a>
              </li>
              {/* <li>
                <a
                  href="../../prts/Lists/External_Approver/AllItems.aspx"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                     <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i> External Approver
                </a>
              </li> */}
              <li>
                <a
                  href="../../PRTS/_layouts/15/viewlsts.aspx?view=14"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                    <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i> Site Content
                </a>
              </li>
            </ul>
          </li>
          <li className={`nav-item has-submenu ${getActiveClass('Configure')}`} style={{display:'none'}}>
            <div className="nav-link">
              <i className="fa fa-gears" aria-hidden="true"></i> Configure
            </div>
            <ul className="sub-menu configuresection">
            </ul>
          </li>
          </>
        )}
        {(AppAdmin || Admin) && (
          <>
          <li className="nav-item has-submenu settings">
            <div className="nav-link">
              <i className="fa fa-gears" aria-hidden="true"></i> Master (App Admin)
            </div>
            <ul className="sub-menu">
              <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/PRTSResource')}`}
                onClick={() => history.push('/PRTSResource')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>PRTS Resource
              </a>
            </li>
            <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/IssueCategory')}`}
                onClick={() => history.push('/IssueCategory')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>Issue Category
              </a>
            </li>
            <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/Commodity')}`}
                onClick={() => history.push('/Commodity')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>Commodity
              </a>
            </li>
            <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/MFGShop')}`}
                onClick={() => history.push('/MFGShop')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>MFG Shop Selection
              </a>
            </li>
            {/* <li className="nav-item">
              <a
                className={`nav-link ${getActiveClass('/BuildType')}`}
                onClick={() => history.push('/BuildType')}
              >
                <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>Build Type
              </a>
            </li> */}
            </ul>
          </li>
          <li className={`nav-item has-submenu ${getActiveClass('Configure')}`} style={{display:'none'}}>
            <div className="nav-link">
              <i className="fa fa-gears" aria-hidden="true"></i> Configure
            </div>
            <ul className="sub-menu configuresection">
              {/* <li>
                <a
                  className={getActiveClass('/Warehouse')}
                  onClick={() => history.push('/Warehouse')}
                >
                   <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>  Warehouse Controller Users
                </a>
              </li> */}
              {/* <li>
                <a
                  className={getActiveClass('/Material')}
                  onClick={() => history.push('/Material')}
                >
                   <i className="fas fa-tasks" style={{ marginRight: '8px' }}></i>  Material Controller Users
                </a>
              </li> */}
            </ul>
          </li>
          </>
        )}
      </ul>
    </div>
  );
};

export default Sidebar;
