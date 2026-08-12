import * as React from "react";
import { useState, useEffect, useRef } from "react";
import type { IPrtsProps } from "../../IPrtsProps";
import SPCRUDOPS from "../../../service/DAL/spcrudops";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useHistory } from 'react-router-dom';
import "../CSS/NewRequest.scss";
import * as XLSX from 'xlsx';
import { sp } from '@pnp/sp/presets/all';
import { Web } from '@pnp/sp/presets/all';
import { formatAmount } from "../../../service/BAL/SPCRUD/Helper";
import MasterPagesRequestsOps from "../../../service/BAL/SPCRUD/MasterPages";
import IPRTSACLRequestsOps from "../../../service/BAL/SPCRUD/PRTSACL";

export const Commodity: React.FC<IPrtsProps> = (props: IPrtsProps) => {

    const [searchTerm, setSearchTerm] = useState("");
    const [CommodityData, setCommodityData] = useState<any[]>([]);
    const [aclData, setAclData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [popupVisible, setPopupVisible] = useState(false);
    const [loading, setLoading] = useState(true);
    const [recordsPerPage, setRecordsPerPage] = useState(10); // default 10 records per page
    const [originalCommodityTitle, setOriginalCommodityTitle] = useState("");
    const [isEdit, setIsEdit] = useState(false);
    const [selectedId, setSelectedId] = useState<number | null>(null);
    const [headPickerKey, setHeadPickerKey] = useState(0);
    const [leadPickerKey, setLeadPickerKey] = useState(0);

    const [vendorForm, setVendorForm] = useState({
        Title: "",
        CommodityHeadId: null,
        CommodityHeadEmail: "",
        CommodityLeadId: null,
        CommodityLeadEmail: ""
    });

    const [columnFilters, setColumnFilters] = useState({
        Title: "",
        CommodityHeadId: null,
        CommodityLeadId: null
    });

    const resetFilters = () => {
        setColumnFilters({
            Title: "",
            CommodityHeadId: null,
            CommodityLeadId: null
        });
        setSearchTerm("");
    };
    // Smooth‑scroll anchor
    const pageTopRef = useRef<HTMLDivElement | null>(null);
    const itemsPerPage = recordsPerPage;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);
    const history = useHistory();

    const formatDate = (dateString: string) => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-GB'); // dd/mm/yyyy
    };

    const GetCommodityData = async () => {
        let employeeProfile = await EmployeeProfile(props.userEmail);
        try {
            setLoading(true);
            const Commodity = await MasterPagesRequestsOps().getCommodityData(
                { column: "ID", isAscending: false },
                props,
                ''
            );
        
            //console.log('Commodity data: ', Commodity);// Debug log
            setCommodityData(Commodity);
            setFilteredData(Commodity);
        } catch (err: any) {
            console.error(err.message);
            alert("Error fetching Commodity data. Please try again.");
        } finally {
            setLoading(false);
        }
    };

    const GetACLData = async () => {
        try {
            setLoading(true);
            const ACLData = await IPRTSACLRequestsOps().getIPRTSACLData(
                { column: "ID", isAscending: true },
                props,
                `UserName/EMail ne null and EmployeeID ne null and Status eq 'Active'`
            )
            setAclData(ACLData);
        } catch (err: any) {
            console.error(err.message);
            alert("Error fetching ACL data. Please try again.");
        } finally {
            setLoading(false);
        }
    }

    useEffect(() => {
        GetCommodityData();
        GetACLData();
    }, []);

    //Filter Search based on each column 
    useEffect(() => {
        let filtered = CommodityData;
        Object.keys(columnFilters).forEach((key) => {
            const value = columnFilters[key]?.toString().toLowerCase() || "";
            if (value) {
                filtered = filtered.filter((item) => {
                    if (!item[key]) return false;
                    if (key === "Created") {
                        return formatDate(item[key]).toLowerCase().includes(value);
                    }
                    return item[key].toString().toLowerCase().includes(value);
                });
            }
        });

        setFilteredData(filtered);
        setCurrentPage(1);
    }, [columnFilters]);

    //filter based on search
    useEffect(() => {
        if (!searchTerm) {
            setFilteredData(CommodityData);
        } else {
            const lowerSearch = searchTerm.toLowerCase();
            const filtered = CommodityData.filter(item =>
                (item.Title || "").toLowerCase().includes(lowerSearch) 
            );
            setFilteredData(filtered);
            setCurrentPage(1);
        }
    }, [searchTerm, CommodityData]);  // ✅ added CommodityData

    const handleColumnFilterChange = (key: string, value: string) => {
        setColumnFilters(prev => ({ ...prev, [key]: value }));
    };


    //List Data of User Master
    async function EmployeeProfile(Email: string) {
        const spCrudOps = await SPCRUDOPS();
        return spCrudOps.getRootData(
            'UserMaster',
            'EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID',
            'FullName,DirectManagerName,OfficeCity,DepartmentCode',
            `FullName/EMail eq '${Email}'`,
            { column: 'ID', isAscending: true },
            props
        );
    }

    const GetEmployeeID = async (Email: string): Promise<string | null> => {
        try {
        const spCrudOps = await SPCRUDOPS();

        const EmployeeProfiledata = await spCrudOps.getRootData(
            'UserMaster',
            'EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID',
            'FullName,DirectManagerName,OfficeCity,DepartmentCode',
            `FullName/EMail eq '${Email}' and EmployeeStatus eq 'Active'`,
            { column: 'ID', isAscending: true },
            props
        );

        if (!EmployeeProfiledata || EmployeeProfiledata.length === 0) {
            console.warn("Employee ID not found for email:", Email);
            return null;
        }

        if (EmployeeProfiledata.length > 1) {
            console.warn("Multiple active employees found with the same email:", Email);
            return null;
        }

        const empId = EmployeeProfiledata[0]?.EmployeeId;
        if (!empId) {
            console.warn("EmployeeId field missing for email:", Email);
            return null;
        }

        return empId;

        } catch (error) {
        console.error("Error fetching Employee ID for " + Email + ":", error);
        return null;
        }
    }

    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
            /* Smooth scroll to top of table */
            pageTopRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    };

    const paginatedData = filteredData.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage
    );

    const exportToExcel = () => {
        // Always export all filtered data (ignore pagination)
        const dataToExport = [...filteredData].sort((a, b) => b.ID - a.ID);

        if (dataToExport.length === 0) {
            alert("No records found to export.");
            return;
        }

        // Map fields to clean column labels
        const exportData = dataToExport.map((item) => ({
            "Commodity Name": item.Title
        }));

        // Create sheet + workbook
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Commodity");

        // Save file with today’s date
        const today = new Date().toISOString().slice(0, 10); // yyyy-mm-dd
        XLSX.writeFile(workbook, `Commodity_${today}.xlsx`);
    };

    const openAddPopup = () => {
    setIsEdit(false);
    setSelectedId(null);
    setVendorForm({
        Title: "",
        CommodityHeadId: null,
        CommodityHeadEmail: "",
        CommodityLeadId: null,
        CommodityLeadEmail: ""
    });
    setPopupVisible(true);
    };

    const openEditPopup = (item) => {
    setIsEdit(true);
    setSelectedId(item.ID);
    // store original Commodity Name
    setOriginalCommodityTitle(item.Title);
    setVendorForm({
        Title: item.Title,
        CommodityHeadId: item.CommodityHeadId || null,
        CommodityHeadEmail: item.CommodityHeadEmail || "",
        CommodityLeadId: item.CommodityLeadId || null,
        CommodityLeadEmail: item.CommodityLeadEmail || ""
    });
    setPopupVisible(true);
    };

    const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setVendorForm(prev => ({ ...prev, [name]: value }));
    };

    const validateForm = () => {
      if (!vendorForm.Title) {
        alert("Commodity Name is required");
        return false;
      }
      if (!vendorForm.CommodityHeadId) {
        alert("Commodity Head is required");
        return false;
      }
      if (!vendorForm.CommodityLeadId) {
        alert("Commodity Lead is required");
        return false;
      }

      return true;
    };

    const update = async () => {
        const spCrudObj = await SPCRUDOPS();
        if (!validateForm()) return;
        setLoading(true);
        try {
            // 🔹 Check duplicate CommodityName
            // if (vendorForm.Title !== originalCommodityTitle) {
              const title = vendorForm.Title.trim();
              const vendorData = await MasterPagesRequestsOps().getCommodityData(
                  { column: "ID", isAscending: true },
                  props,
                  `Title eq '${title.replace(/'/g, "''")}' and CommodityHeadId eq ${vendorForm.CommodityHeadId} and CommodityLeadId eq ${vendorForm.CommodityLeadId}`
              );

              // 🔹 Ignore same record when editing
              // const duplicateVendor = vendorData.filter(
              //     item => Number(item.ID) !== Number(selectedId)
              // );

              if (vendorData.length > 0) {
                  alert("Commodity Name already exists!");
                  setLoading(false);
                  return;
              }
            // }  
            const payload = {
                Title: vendorForm.Title,
                CommodityHeadId: vendorForm.CommodityHeadId,
                CommodityLeadId: vendorForm.CommodityLeadId
            };
            if (isEdit && selectedId) {
            await spCrudObj.updateRootData('CommodityList', selectedId, payload, props);
            alert("Commodity details updated successfully!");
            } else {
            await spCrudObj.insertRootData('CommodityList', payload, props);
            alert("Commodity details added successfully!");
            }

            setPopupVisible(false);
            GetCommodityData(); // refresh grid
        } catch (err: any) {
            console.error(err.message);
            alert("Error saving Commodity details. Please try again.");
        } finally {
            setLoading(false); 
        }
    };




    return (
        <div className="min-h-screen bg-gray-100">
            {/* Anchor for scroll-to-top */}
            <div ref={pageTopRef} />

            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>Commodity Master Dashboard</h2>
                    </div>
                </div>
            </div>

            <main className="Main-Dash">
                {loading ? (
                    <div className="loading-overlay">
                        <div className="loading-content">
                        <svg
                            className="loading-spinner"
                            xmlns="http://www.w3.org/2000/svg"
                            fill="none"
                            viewBox="0 0 24 24"
                        >
                            <circle
                            className="opacity-25"
                            cx="12"
                            cy="12"
                            r="10"
                            stroke="currentColor"
                            strokeWidth="4"
                            />
                            <path
                            className="opacity-75"
                            fill="currentColor"
                            d="M4 12a8 8 0 018-8v8H4z"
                            />
                        </svg>
                        <p className="text-white text-lg">Please wait, loading data...</p>
                        </div>
                    </div>
                ) : (
                    <>
                        {/* Search Bar */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '10px 0' }}>
                          <div className="flex flex-col">
                            <input
                                type="text"
                                placeholder="Search..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="w-64 px-4 py-2 text-sm border-gray-300 rounded-full dashboard-sha focus:outline-none focus:ring-2 focus:ring-red-500"
                                style={{ width: "250px", margin: "10px 10px 10px 0px" }}
                            />
                          </div>
                            {filteredData.length > 0 && (
                                <button className="btn btn-warning export-btn" type="button" onClick={exportToExcel} style={{ marginLeft: "10px" }}>
                                    Export Data
                                </button>
                            )}
                            <i
                                className="fa fa-refresh cursor-pointer text-xl text-gray-700 hover:text-black"
                                onClick={resetFilters}
                                title="Reset Filters"
                                style={{ paddingLeft: "10px" }}
                            ></i>
                            <button className="btn btn-warning export-btn" onClick={openAddPopup} style={{ marginLeft: "10px" }}>
                                <i className="fa fa-mail-forward"></i>Add Commodity
                            </button>
                        </div>
                        {/* Table */}
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll max-h-[65vh] overflow-y-auto">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead style={{ backgroundColor: "#ce0b0e", position: "sticky", top: "0px" }} className="text-white">
                                        <tr>
                                            <th className="px-4 py-2"></th>
                                            <th className="px-4 py-2">Commodity Name</th>
                                            <th className="px-4 py-2">Commodity Head</th>
                                            <th className="px-4 py-2">Commodity Lead</th>
                                        </tr>
                                        <tr className="bg-gray-100 text-black">
                                            <th></th>
                                            {["Title", "CommodityHeadTitle", "CommodityLeadTitle"].map((col) => (
                                                <th key={col} className="px-4 py-1">
                                                    <input
                                                        type="text"
                                                        value={columnFilters[col]}
                                                        onChange={(e) => handleColumnFilterChange(col, e.target.value)}
                                                        className="w-full px-2 py-1 text-xs border border-gray-300 rounded"
                                                        placeholder="Search"
                                                        style={{ width: "140px" }}
                                                    />
                                                </th>
                                            ))}
                                            <th></th>
                                            <th></th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {[...filteredData].sort((a, b) => b.ID - a.ID).slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage)
                                            .map((item, index) => (
                                                <tr key={index} className="border-t">
                                                    <td className="px-4 py-2"><button className="btn btn-warning export-btn" onClick={() => openEditPopup(item)}><i className="fa fa-edit"></i>Edit</button></td>
                                                    <td className="px-4 py-2">{item.Title}</td>
                                                    <td className="px-4 py-2">{item.CommodityHeadTitle}</td>
                                                    <td className="px-4 py-2">{item.CommodityLeadTitle}</td>

                                                </tr>
                                            ))}
                                    </tbody>
                                </table>
                            </div>

                            {/* Pagination */}
                            <div className="flex justify-center mt-6 overflow-x-auto">
                                <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-orange rounded shadow">
                                    <button
                                        onClick={() => handlePageChange(currentPage - 1)}
                                        disabled={currentPage === 1}
                                        style={{ backgroundColor: "orange", color: "black", opacity: currentPage === 1 ? 0.5 : 1 }}
                                        className="px-3 py-1 border rounded"
                                    >
                                        Previous
                                    </button>

                                    {currentPage > 3 && (
                                        <>
                                            <button
                                                onClick={() => handlePageChange(1)}
                                                style={{ backgroundColor: "orange", color: "black" }}
                                                className="px-3 py-1 border rounded"
                                            >
                                                1
                                            </button>
                                            <span className="px-2">...</span>
                                        </>
                                    )}

                                    {Array.from({ length: totalPages }, (_, i) => i + 1)
                                        .filter((page) => Math.abs(page - currentPage) <= 2)
                                        .map((page) => (
                                            <button
                                                key={page}
                                                onClick={() => handlePageChange(page)}
                                                style={{ backgroundColor: currentPage === page ? "yellow" : "orange", color: "black", fontWeight: currentPage === page ? "bold" : "normal" }}
                                                className="px-3 py-1 border rounded"
                                            >
                                                {page}
                                            </button>
                                        ))}

                                    {currentPage < totalPages - 2 && (
                                        <>
                                            <span className="px-2">...</span>
                                            <button
                                                onClick={() => handlePageChange(totalPages)}
                                                style={{ backgroundColor: "orange", color: "black" }}
                                                className="px-3 py-1 border rounded"
                                            >
                                                {totalPages}
                                            </button>
                                        </>
                                    )}

                                    <button
                                        onClick={() => handlePageChange(currentPage + 1)}
                                        disabled={currentPage === totalPages}
                                        style={{ backgroundColor: "orange", color: "black", opacity: currentPage === totalPages ? 0.5 : 1 }}
                                        className="px-3 py-1 border rounded"
                                    >
                                        Next
                                    </button>
                                </div>
                            </div>
                        </div>
                        {popupVisible && (
                            <>
                                <div className="modal fade show d-block" tabIndex={-1} role="dialog" aria-hidden="false">
                                    <div className="modal-dialog modal-lg">
                                        <div className="modal-content">
                                            <div className="modal-body">
                                                <h4 className="modal-title">{isEdit ? "Edit Commodity" : "Add Commodity"}</h4>
                                                <table className="table table-bordered">
                                                    <colgroup>
                                                        <col style={{ width: '30%' }} />
                                                        <col style={{ width: '70%' }} />
                                                    </colgroup>
                                                    <tbody>
                                                        <tr>
                                                            <td>Commodity Name:</td>
                                                            <td>
                                                           <input
                                                                name="Title"
                                                                className="form-control"
                                                                value={vendorForm.Title}
                                                                onChange={handleInputChange}
                                                            />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Commodity Head:</td>
                                                            <td>
                                                                <PeoplePicker
                                                                    key={headPickerKey}
                                                                    context={props.currentSPContext}
                                                                    personSelectionLimit={1}
                                                                    principalTypes={[PrincipalType.User]}
                                                                    resolveDelay={500}
                                                                    defaultSelectedUsers={vendorForm.CommodityHeadEmail ? [vendorForm.CommodityHeadEmail] : []}
                                                                    onChange={async (users: any[]) => {
                                                                        if (!users || users.length === 0) {
                                                                            setVendorForm({ ...vendorForm, CommodityHeadId: null, CommodityHeadEmail: "" });
                                                                            return;
                                                                        }
                                                                        try {
                                                                            const email = users[0].secondaryText || users[0].loginName;
                                                                            const checkACLData = aclData.find((acl) => acl.UserNameEmail?.toLowerCase() === email.toLowerCase());
                                                                            if (!checkACLData) {
                                                                                alert("Selected user does not exist in ACL. \nPlease Contact IT team.");
                                                                                setVendorForm({ ...vendorForm, CommodityHeadId: null, CommodityHeadEmail: "" });
                                                                                setHeadPickerKey(prev => prev + 1); // reset PeoplePicker
                                                                                return;
                                                                            }
                                                                            const empId = await GetEmployeeID(email);
                                                                            if (!empId) {
                                                                                alert("Selected user does not exist in User Master. \nPlease Contact IT team.");
                                                                                setVendorForm({ ...vendorForm, CommodityHeadId: null, CommodityHeadEmail: "" });
                                                                                setHeadPickerKey(prev => prev + 1); // reset PeoplePicker
                                                                                return;
                                                                            }
                                                                            const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
                                                                            const user = await web.ensureUser(users[0].loginName);
                                                                            setVendorForm({ ...vendorForm, CommodityHeadId: user.data.Id || null, CommodityHeadEmail: email });
                                                                        } catch (error: any) {
                                                                            console.error("Error ensuring user for Id:", error);
                                                                        }
                                                                    }}
                                                                />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>Commodity Lead:</td>
                                                            <td>
                                                                <PeoplePicker
                                                                    key={leadPickerKey}
                                                                    context={props.currentSPContext}
                                                                    personSelectionLimit={1}
                                                                    principalTypes={[PrincipalType.User]}
                                                                    resolveDelay={500}
                                                                    defaultSelectedUsers={vendorForm.CommodityLeadEmail ? [vendorForm.CommodityLeadEmail] : []}
                                                                    onChange={async (users: any[]) => {
                                                                        if (!users || users.length === 0) {
                                                                            setVendorForm({ ...vendorForm, CommodityLeadId: null, CommodityLeadEmail: "" });
                                                                            return;
                                                                        }
                                                                        try {
                                                                            const email = users[0].secondaryText || users[0].loginName;
                                                                            const checkACLData = aclData.find((acl) => acl.UserNameEmail?.toLowerCase() === email.toLowerCase());
                                                                            if (!checkACLData) {
                                                                                alert("Selected user does not exist in ACL. \nPlease Contact IT team.");
                                                                                setVendorForm({ ...vendorForm, CommodityHeadId: null, CommodityHeadEmail: "" });
                                                                                setHeadPickerKey(prev => prev + 1); // reset PeoplePicker
                                                                                return;
                                                                            }
                                                                            const empId = await GetEmployeeID(email);
                                                                            if (!empId) {
                                                                                alert("Selected user does not exist in User Master. \nPlease Contact IT team.");
                                                                                setVendorForm({ ...vendorForm, CommodityLeadId: null, CommodityLeadEmail: "" });
                                                                                setLeadPickerKey(prev => prev + 1); // reset PeoplePicker
                                                                                return;
                                                                            }
                                                                            const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
                                                                            const user = await web.ensureUser(users[0].loginName);
                                                                            setVendorForm({ ...vendorForm, CommodityLeadId: user.data.Id || null, CommodityLeadEmail: email });
                                                                        } catch (error: any) {
                                                                            console.error("Error ensuring user for Id:", error);
                                                                        }
                                                                    }}
                                                                />
                                                            </td>
                                                        </tr>
                                                        
                                                    </tbody>
                                                </table>
                                            </div>
                                            <div className="modal-footer">
                                                <button type="button" className="btn btn-primary" onClick={update}>{isEdit ? "Update" : "Submit"}</button>
                                                <button type="button" className="btn btn-secondary" onClick={() => setPopupVisible(false)}>Cancel</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div className="modal-backdrop fade show" />
                            </>
                        )}
                    </>
                )}
            </main>
        </div>
    );

};

export default Commodity;
