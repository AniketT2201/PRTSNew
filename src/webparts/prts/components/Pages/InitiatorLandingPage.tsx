import * as React from 'react';
import { useEffect } from 'react';
import type { IPrtsProps } from '../IPrtsProps';
import '../Pages/CSS/Landing.scss';
import SPCRUDOPS from "../../service/DAL/spcrudops";
import { useHistory } from 'react-router-dom';

const InitiatorLanding: React.FC<IPrtsProps> = (props: IPrtsProps) => {
  const history = useHistory();
  const [isEditor, setIsEditor] = React.useState(false);
  const [isAppAdmin, setIsAppAdmin] = React.useState(false);
  const [isSysAdmin, setIsSysAdmin] = React.useState(false);

  const openPage = (url: string): void => {
    console.log(`Open page: ${url}`);
    // Example: window.location.href = url;
  };

  const loadFramePage = (url: string): void => {
    console.log(`Load frame page: ${url}`);
    // Example: dynamically load content
  };

  const displayModel = (type: number): void => {
    const selector = type === 1 ? '.displayMovementTypes' : '.displayMovementWiseFlow';
    const modal = document.querySelector(selector);
    modal?.classList.remove('hidden');
  };

  const closeModal = (selector: string): void => {
    document.querySelector(selector)?.classList.add('hidden');
  };

  // Load initial data
  useEffect(() => {
    const fetchData = async () => {
      try {
        const spCrudOps = await SPCRUDOPS();
        const data = await spCrudOps.getData(
          'PRTSACL',
          'ID,Title,UserName/Title,UserName/EMail,Role',
          'UserName',
          '',
          { column: 'ID', isAscending: true },
          props
        );
        const filteredData = data.filter(
          item => item.UserName?.EMail?.trim().toLowerCase() === props.userEmail?.trim().toLowerCase()
        );
        if (filteredData) {
          setIsEditor(filteredData[0].Role === "Editor");

          if (filteredData[0].Title === "SysAdmin") {
            setIsSysAdmin(true);
            setIsAppAdmin(true);
          } else if (filteredData[0].Title === "AppAdmin") {
            setIsAppAdmin(true);
          }
        }
      } catch (error) {
        console.error("Failed to fetch ACL data:", error);
      }
    };

    fetchData();
  }, []);




  return (
    <>
      <div id="header"></div>

      {/* Main Content */}
      <div id="contentPage" className="container-fluid p-0">
        <div className="help-info mt-4 mb-4">
          <h1>Welcome to Problem Resolving Tracking System</h1>
          {/* <p></p>
          <h1>About</h1>
          <p>This is an approval form for the material requirement from warehouse for various purposes.</p>

          <h1>Workflow</h1>
          <ul>
            <li>
              For User of SCM<br />
              <span>Initiator</span> →
              <span>Manager</span> →
              <span>Manager2</span> →
              <span>Material Controller</span> →
              <span>PDC Head</span> →
              <span>WH Controller</span> →
              <span>Inventory Controller</span>               
            </li>
            <li>
              For User of PDC<br />
              <span>Initiator</span> →
              <span>Manager</span> →
              <span>Manager2</span> →
              <span>Material Controller</span> →              
              <span>WH Controller</span> →
              <span>Inventory Controller</span> →
              <span>SCM Head</span> →
              <span>Warehouse</span>          
            </li>
            <li>
              For Other Departments<br />
              <b>Condition 1: when amount is ≥ 20k</b><br />
              <span>Initiator</span> →
              <span>Manager</span> →
              <span>Department Staff Head</span> →
              <span>Material Controller</span> →
              <span>WH Controller</span> →
              <span>Inventory Controller</span> →
              <span>SCM Head</span> →
              <span>Finance</span> →
              <span>WareHouse</span>
              <br />
              <b>Condition 2: when amount is &lt; 20k</b><br />
              <span>Initiator</span> →
              <span>Manager</span> →
              <span>Material Controller</span> →
              <span>WH Controller</span> →
              <span>Inventory Controller</span> →
              <span>SCM Head</span> →
              <span>Finance</span> →
              <span>WareHouse</span>
            </li>            
          </ul> */}
        </div>
      </div>
    </>
  );
};

export default InitiatorLanding;
