import * as React from "react";
import { useState, useEffect, useRef } from "react";
import type { IPrtsProps } from "../../IPrtsProps";
import SPCRUDOPS from "../../../service/DAL/spcrudops";
import { useHistory } from 'react-router-dom';
import "../CSS/NewRequest.scss";
import * as XLSX from 'xlsx';
import { formatAmount } from "../../../service/BAL/SPCRUD/Helper";
import MasterPagesRequestsOps from "../../../service/BAL/SPCRUD/MasterPages";

export const IssueCategory: React.FC<IPrtsProps> = (props: IPrtsProps) => {

    const [searchTerm, setSearchTerm] = useState("");
    const [IssueCategoryData, setIssueCategoryData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [popupVisible, setPopupVisible] = useState(false);
    const [loading, setLoading] = useState(true);
    const [recordsPerPage, setRecordsPerPage] = useState(10); // default 10 records per page
    const [originalIssueCategoryTitle, setOriginalIssueCategoryTitle] = useState("");
    const [isEdit, setIsEdit] = useState(false);
    const [selectedId, setSelectedId] = useState<number | null>(null);

    const [vendorForm, setVendorForm] = useState({
    Title: ""
    });

    const [columnFilters, setColumnFilters] = useState({
        Title: ""
    });

    const resetFilters = () => {
        setColumnFilters({
            Title: ""
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

    const GetIssueCategoryData = async () => {
        let employeeProfile = await EmployeeProfile(props.userEmail);
        try {
            setLoading(true);
            const IssueCategory = await MasterPagesRequestsOps().getIssueCategoryData(
                { column: "ID", isAscending: false },
                props,
                ''
            );
        
            //console.log('Issue Category data: ', IssueCategory);// Debug log
            setIssueCategoryData(IssueCategory);
            setFilteredData(IssueCategory);
        } catch (err: any) {
            console.error(err.message);
            alert("Error fetching Issue Category data. Please try again.");
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        GetIssueCategoryData();
    }, []);

    //Filter Search based on each column 
    useEffect(() => {
        let filtered = IssueCategoryData;
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
            setFilteredData(IssueCategoryData);
        } else {
            const lowerSearch = searchTerm.toLowerCase();
            const filtered = IssueCategoryData.filter(item =>
                (item.Title || "").toLowerCase().includes(lowerSearch) 
            );
            setFilteredData(filtered);
            setCurrentPage(1);
        }
    }, [searchTerm, IssueCategoryData]);  // ✅ added IssueCategoryData

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
            "Issue Category Name": item.Title
        }));

        // Create sheet + workbook
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "IssueCategory");

        // Save file with today’s date
        const today = new Date().toISOString().slice(0, 10); // yyyy-mm-dd
        XLSX.writeFile(workbook, `IssueCategory${today}.xlsx`);
    };

    const openAddPopup = () => {
    setIsEdit(false);
    setSelectedId(null);
    setVendorForm({
        Title: ""
    });
    setPopupVisible(true);
    };

    const openEditPopup = (item) => {
    setIsEdit(true);
    setSelectedId(item.ID);
    // store original Issue Category Name
    setOriginalIssueCategoryTitle(item.Title);
    setVendorForm({
        Title: item.Title
    });
    setPopupVisible(true);
    };

    const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setVendorForm(prev => ({ ...prev, [name]: value }));
    };

    const validateForm = () => {
      if (!vendorForm.Title) {
        alert("Issue Category Name is required");
        return false;
      }

      return true;
    };

    const update = async () => {
        const spCrudObj = await SPCRUDOPS();
        if (!validateForm()) return;
        setLoading(true);
        try {
            // 🔹 Check duplicate IssueCategoryName
            // if (vendorForm.Title !== originalIssueCategoryTitle) {
              const title = vendorForm.Title.trim();
              const vendorData = await MasterPagesRequestsOps().getIssueCategoryData(
                  { column: "ID", isAscending: true },
                  props,
                  `Title eq '${title.replace(/'/g, "''")}'`
              );

              // 🔹 Ignore same record when editing
              // const duplicateVendor = vendorData.filter(
              //     item => Number(item.ID) !== Number(selectedId)
              // );

              if (vendorData.length > 0) {
                  alert("Issue Category Name already exists!");
                  setLoading(false);
                  return;
              }
            // }  
            if (isEdit && selectedId) {
            await spCrudObj.updateRootData('IssueCategoryList', selectedId, vendorForm, props);
            alert("Issue Category details updated successfully!");
            } else {
            await spCrudObj.insertRootData('IssueCategoryList', vendorForm, props);
            alert("Issue Category details added successfully!");
            }

            setPopupVisible(false);
            GetIssueCategoryData(); // refresh grid
        } catch (err: any) {
            console.error(err.message);
            alert("Error saving Issue Category details. Please try again.");
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
                        <h2>Issue Category Master Dashboard</h2>
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
                                <i className="fa fa-mail-forward"></i>Add Issue Category
                            </button>
                        </div>
                        {/* Table */}
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll max-h-[65vh] overflow-y-auto">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead style={{ backgroundColor: "#ce0b0e", position: "sticky", top: "0px" }} className="text-white">
                                        <tr>
                                            <th className="px-4 py-2"></th>
                                            <th className="px-4 py-2">Issue Category Name</th>
                                        </tr>
                                        <tr className="bg-gray-100 text-black">
                                            <th></th>
                                            {["Title"].map((col) => (
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
                                                <h4 className="modal-title">{isEdit ? "Edit Issue Category" : "Add Issue Category"}</h4>
                                                <table className="table table-bordered">
                                                    <colgroup>
                                                        <col style={{ width: '30%' }} />
                                                        <col style={{ width: '70%' }} />
                                                    </colgroup>
                                                    <tbody>
                                                        <tr>
                                                            <td>Issue Category Name:</td>
                                                            <td>
                                                           <input
                                                                name="Title"
                                                                className="form-control"
                                                                value={vendorForm.Title}
                                                                onChange={handleInputChange}
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

export default IssueCategory;
