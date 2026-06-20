import * as React from "react";
import { useState, useEffect, useRef } from "react";
import type { IPrtsProps } from '../IPrtsProps';
import IASRequestsOps from "../../service/BAL/SPCRUD/PRTS";
import SPCRUDOPS from "../../service/DAL/spcrudops";
import { useHistory } from 'react-router-dom';
import "../Pages/CSS/NewRequest.scss";

export const Print: React.FC<IPrtsProps> = (props: IPrtsProps) => {

    const [searchTerm, setSearchTerm] = useState("");
    const [MDRData, setMDRData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);
    const [showFilterPopup, setShowFilterPopup] = useState(false);
    const [MovementDropdown, setMovementDropdown] = useState<any[]>([]);
    const [filterInputs, setFilterInputs] = useState({
        ageing: "",
        movementType: "",
        approvalNoteYear: "",
    });

    const [columnFilters, setColumnFilters] = useState({
        ID: "",
        Title: "",
        EmpID: "",
        EmpName: "",
        Department: "",
        CostCenter: "",
        CarLine: "",
        Items: "",
        Reason: "",
        NA: "",
        DA: "",
        Status: ""
    });

     const resetFilters = () => {
        setColumnFilters({
            ID: "",
            Title: "",
            EmpID: "",
            EmpName: "",
            Department: "",
            CostCenter: "",
            CarLine: "",
            Items: "",
            Reason: "",
            NA: "",
            DA: "",
            Status: ""
        });
        setSearchTerm("");
    };
    // Smooth‑scroll anchor
    const pageTopRef = useRef<HTMLDivElement | null>(null);
    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);

    const history = useHistory();

    // const handleTitleClick = (id) => {
    //     history.push({
    //         pathname: '/ApprovalForm',
    //         search: `?ItemId=${id}`,
    //         state: { from: '/AllReqDash' }
    //     }
    //     );
    // };

    const handleTitleClick = (id) => {
        // Save dashboard state to sessionStorage
        sessionStorage.setItem(
            'dashboardState',
            JSON.stringify({
                searchTerm,
                columnFilters,
                currentPage,
                filterInputs,
            })
        );

        sessionStorage.setItem('sidebarFrom', '/AllReqDash');

        history.push({
            pathname: '/ApprovalForm',
            search: `?ItemId=${id}`
        });
    };

    async function GetMovementflow() {
        const spCrudOps = await SPCRUDOPS();
        const Momentflowdata = await spCrudOps.getData(
            'MovementFlow',
            '*,ApprovalNoteDescription,Title,ID',
            '',
            '',
            { column: 'ID', isAscending: true },
            props
        );
        setMovementDropdown(Momentflowdata);
    }
    const dateDifference = (fromDt: Date, toDt: Date) => {
        const diff = new Date(toDt.getTime() - fromDt.getTime());
        const days = diff.getTime() / 1000 / 60 / 60 / 24;
        return days.toFixed(0);
    };

    const formatDate = (dateString: string) => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-GB'); // dd/mm/yyyy
    };

    const GetMDRData = async () => {
        await EmployeeProfile(props.userEmail);
        setLoading(true);
        const MDRColl = await IASRequestsOps().getIIASData(
            { column: "ID", isAscending: false },
            props,
            ''
        );
        let MDRCollFilter = MDRColl.filter((test) => test.Status != "Draft")
        setMDRData(MDRCollFilter);
        setFilteredData(MDRCollFilter);
        setLoading(false);
    };

    const applyAdvancedFiltersPageload = (inputs = filterInputs) => {
        // Use "inputs" (either passed-in or current state)
        let filtered = MDRData;

        if (inputs.ageing) {
            filtered = filtered.filter(item => {
                const daysStr = item.Status === "Pending for Approval"
                    ? dateDifference(new Date(item.Created), new Date())
                    : "0";

                const days = parseInt(daysStr || "0", 10); // Ensure it's a number
                return days === parseInt(inputs.ageing, 10);
            });
        }

        if (inputs.movementType) {
            filtered = filtered.filter((item) => item.MovementType === inputs.movementType);
        }

        if (inputs.approvalNoteYear) {
            filtered = filtered.filter((item) =>
                item.ApprovalNoteYear?.toString().includes(inputs.approvalNoteYear)
            );
        }

        setFilteredData(filtered);
    };

    /****************/

    useEffect(() => {
        GetMDRData();
        GetMovementflow();
    }, []);

    useEffect(() => {
        if (MDRData.length > 0) {
            const savedState = sessionStorage.getItem('dashboardState');
            if (savedState) {
                const { searchTerm, columnFilters, currentPage, filterInputs } = JSON.parse(savedState);
                setSearchTerm(searchTerm || '');
                setColumnFilters(columnFilters || {});
                setCurrentPage(currentPage || 1);
                setFilterInputs(filterInputs || {});
                applyAdvancedFiltersPageload(filterInputs || {});
                sessionStorage.removeItem('dashboardState'); // optional cleanup
            }
        }
    }, [MDRData]);

    //Filter Search based on each column 
    useEffect(() => {
        let filtered = MDRData;
        Object.keys(columnFilters).forEach((key) => {
            const value = columnFilters[key].toLowerCase();
            if (value) {
                filtered = filtered.filter((item) => {
                    // Handle missing fields safely
                    if (["AgeingCurrent", "AgeingCreate"].includes(key)) {
                        let days = "0";

                        if (key === "AgeingCurrent") {
                            days = item.Status === "Pending for Approval"
                                ? dateDifference(new Date(item.LastAction), new Date())
                                : "0";
                        }

                        if (key === "AgeingCreate") {
                            days = item.Status === "Pending for Approval"
                                ? dateDifference(new Date(item.Created), new Date())
                                : item.Status === "Draft"
                                    ? "0"
                                    : dateDifference(new Date(item.Created), new Date(item.LastAction));
                        }

                        return days.toLowerCase().includes(value);
                    }

                    if (!item[key]) return false;

                    if (key === "Created") {
                        const formatted = formatDate(item[key]).toLowerCase(); // dd/mm/yyyy
                        return formatted.includes(value);
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
            setFilteredData(MDRData);
        } else {
            const lowerSearch = searchTerm.toLowerCase();
            const filtered = MDRData.filter(item =>
                item.ID?.toString().toLowerCase().includes(lowerSearch) ||
                item.Title?.toLowerCase().includes(lowerSearch) ||
                item.EmpID?.toLowerCase().includes(lowerSearch) ||
                item.EmpName?.toLowerCase().includes(lowerSearch) ||
                item.Department?.toLowerCase().includes(lowerSearch) ||
                item.CostCenter?.toLowerCase().includes(lowerSearch) ||
                item.CarLine?.toLowerCase().includes(lowerSearch) ||
                item.Items?.toLowerCase().includes(lowerSearch) ||
                item.Reason?.toLowerCase().includes(lowerSearch) ||
                item.NATitle?.toLowerCase().includes(lowerSearch) ||
                item.DATitle?.toLowerCase().includes(lowerSearch) ||
                item.Status?.toLowerCase().includes(lowerSearch) ||
                item.Total?.toString().toLowerCase().includes(lowerSearch)
            );
            setFilteredData(filtered);
            setCurrentPage(1);
        }
    }, [searchTerm]);

    const handleColumnFilterChange = (key: string, value: string) => {
        setColumnFilters(prev => ({ ...prev, [key]: value }));
    };

    /**********filter Agening,Movement Vise,Approval Note Year */
    const applyAdvancedFilters = () => {
        const { ageing, movementType, approvalNoteYear } = filterInputs;

        let filtered = [...MDRData];

        // Ageing Filter (ensure both are numbers)
        if (ageing) {
            filtered = filtered.filter(item => {
                const daysStr = item.Status === "Pending for Approval"
                    ? dateDifference(new Date(item.Created), new Date())
                    : "0";

                const days = parseInt(daysStr || "0", 10); // Ensure it's a number
                return days === parseInt(ageing, 10);
            });
        }

        // Movement Type Filter
        if (movementType) {
            filtered = filtered.filter(item =>
                item.MovementType?.toLowerCase() === movementType.toLowerCase()
            );
        }

        // Approval Note Year Filter (partial match on Title)
        if (approvalNoteYear) {
            const searchTerm = approvalNoteYear.toLowerCase();
            filtered = filtered.filter(item =>
                item.ApprovalNoteNo?.toLowerCase().includes(searchTerm)
            );
        }

        setFilteredData(filtered);
        setCurrentPage(1);
    };
    /********* */


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

    /*** ──────────────────────────
     * RENDER
     * ────────────────────────── */
    return (
        <div className="min-h-screen bg-gray-100">
            {/* Anchor for scroll-to-top */}
            <div ref={pageTopRef} />

            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>All Request Dashboard</h2>
                    </div>
                </div>
            </div>

            <main className="p-6">
                {loading ? (
                    <div className="loading-overlay">
                        <div className="loading-content">
                            <svg className="loading-spinner" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
                            </svg>
                            <p className="text-white text-lg">Please wait, loading data...</p>
                        </div>
                    </div>
                ) : (
                    <>
                        {/* Search Bar */}
                        <div className="flex items-center gap-4 mb-4">
                            <input
                                type="text"
                                placeholder="Search..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="w-64 px-4 py-2 text-sm border border-gray-300 rounded-full shadow focus:outline-none focus:ring-2 focus:ring-red-500"
                                style={{ width: "250px", marginRight: "20px", marginTop: "20px" }}
                            />                            
                            <button className="btn btn-warning export-btn" type="button" onClick={() => setShowFilterPopup(true)}>Filter</button>                                                            
                            <i
                                className="fa fa-refresh cursor-pointer text-xl text-gray-700 hover:text-black"
                                onClick={resetFilters}
                                title="Reset Filters"
                                style={{paddingLeft:"10px"}}
                            ></i>
                        </div>
                        {/* Table */}
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll max-h-[65vh] overflow-y-auto">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead style={{ backgroundColor: "#ce0b0e" }} className="text-white">
                                        <tr>
                                            <th className="px-4 py-2">Req No</th>
                                            <th className="px-4 py-2">Emp No</th>
                                            <th className="px-4 py-2">Emp Name</th>
                                            <th className="px-4 py-2">Department</th>
                                            <th className="px-4 py-2">Cost Center</th>
                                            <th className="px-4 py-2">CarLine</th>
                                            <th className="px-4 py-2">Total Amt.</th>
                                            <th className="px-4 py-2">Reason</th>
                                            <th className="px-4 py-2">Next Approver</th>
                                            <th className="px-4 py-2">Delegated Approver</th>
                                            <th className="px-4 py-2">Status</th>
                                        </tr>
                                        <tr className="bg-gray-100 text-black">
                                            {["Title", "EmpID", "EmpName", "Department", "CostCenter", "CarLine", "Total", "Reason", "NATitle", "DATitle","Status"].map((col) => (
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
                                                <td className="px-4 py-2"><a onClick={(e) => { e.preventDefault(); handleTitleClick(item.ID); }} href="#" className="text-blue-600 hover:text-blue-800 underline">{item.Title}</a></td>
                                                <td className="px-4 py-2">{item.EmpID}</td>
                                                <td className="px-4 py-2">{item.EmpName}</td>
                                                <td className="px-4 py-2">{item.Department}</td>
                                                <td className="px-4 py-2">{item.CostCenter}</td>
                                                <td className="px-4 py-2">{item.CarLine}</td>
                                                <td className="px-4 py-2">{item.Total.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                                                <td className="px-4 py-2">{item.Reason}</td>
                                                <td className="px-4 py-2">{item.NATitle}</td>
                                                <td className="px-4 py-2">{item.DATitle}</td>
                                                <td className="px-4 py-2">{item.Status}</td>
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

                                    {showFilterPopup && (
                                        <>
                                            <div
                                                className="modal fade show d-block"
                                                tabIndex={-1}
                                                role="dialog"
                                                aria-hidden="false"
                                            >
                                                <div className="modal-dialog modal-lg">
                                                    <div className="modal-content">
                                                        <div className="modal-body">
                                                            {/* Header */}
                                                            <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
                                                                <span className="h4">Advanced Filters</span>
                                                            </div>

                                                            {/* Filter Form */}
                                                            <form className="mt-3">
                                                                {/* Ageing */}
                                                                <div className="form-group mb-3">
                                                                    <label>Ageing (from Create Date)</label>
                                                                    <input
                                                                        type="number"
                                                                        className="form-control"
                                                                        value={filterInputs.ageing || ""}
                                                                        onChange={(e) =>
                                                                            setFilterInputs({ ...filterInputs, ageing: e.target.value })
                                                                        }
                                                                        placeholder="Enter number of days"
                                                                    />
                                                                </div>

                                                                {/* Movement Type */}
                                                                <div className="form-group mb-3">
                                                                    <label>Movement Type</label>
                                                                    <select
                                                                        className="select.form-control.mt-2"
                                                                        value={filterInputs.movementType || ""}
                                                                        onChange={(e) =>
                                                                            setFilterInputs({
                                                                                ...filterInputs,
                                                                                movementType: e.target.value,
                                                                            })
                                                                        }
                                                                    >
                                                                        <option value="">Select</option>
                                                                        {MovementDropdown?.map((Vend) => (
                                                                            <option key={Vend.ID} value={Vend.Title}>
                                                                                {Vend.ApprovalNoteDescription}
                                                                            </option>
                                                                        ))}
                                                                    </select>
                                                                </div>

                                                                {/* Approval Note Year */}
                                                                <div className="form-group mb-3">
                                                                    <label>Approval Note Year</label>
                                                                    <input
                                                                        type="text"
                                                                        className="form-control"
                                                                        value={filterInputs.approvalNoteYear || ""}
                                                                        onChange={(e) =>
                                                                            setFilterInputs({
                                                                                ...filterInputs,
                                                                                approvalNoteYear: e.target.value,
                                                                            })
                                                                        }
                                                                        placeholder="Enter year or keyword"
                                                                    />
                                                                </div>
                                                            </form>
                                                        </div>

                                                        {/* Footer Buttons */}
                                                        <div className="modal-footer">
                                                            <button
                                                                type="button"
                                                                className="btn btn-primary"
                                                                onClick={() => {
                                                                    applyAdvancedFilters();
                                                                    setShowFilterPopup(false);
                                                                }}
                                                            >
                                                                Apply Filters
                                                            </button>
                                                            <button
                                                                type="button"
                                                                className="btn btn-secondary"
                                                                onClick={() => setShowFilterPopup(false)}
                                                            >
                                                                Cancel
                                                            </button>
                                                            <button
                                                                onClick={() => {
                                                                    setFilterInputs({ ageing: "", movementType: "", approvalNoteYear: "" });
                                                                    setFilteredData(MDRData);
                                                                }}
                                                                className="ml-2 px-3 py-1 text-sm bg-gray-200 rounded hover:bg-gray-300"
                                                            >
                                                                Clear Filters
                                                            </button>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="modal-backdrop fade show" />
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
                    </>
                )}
            </main>
        </div>
    );

};

export default Print;
