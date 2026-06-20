import * as React from "react";
import { useState, useEffect } from "react";
import type { IPrtsProps } from "../IPrtsProps";
import IASRequestsOps from "../../service/BAL/SPCRUD/PRTS";
import SPCRUDOPS from "../../service/DAL/spcrudops";
import { useHistory } from "react-router-dom";
import "../Pages/CSS/NewRequest.scss";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import * as moment from "moment";

export const Dashboard: React.FC<IPrtsProps> = (props: IPrtsProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [columnFilters, setColumnFilters] = useState<any>({});
    const [MDRData, setMDRData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);

    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);
    const history = useHistory();

    const handleTitleClick = (id: number) => {
        history.push({
            pathname: `/InitiatorLandingedit/${id}`,
            // search: `?ItemId=${id}`,
            state: { from: "/Dashboard" },
        });
    };

    const GetMDRData = async () => {
        await EmployeeProfile(props.userEmail);
        setLoading(true);
        const MDRColl = await IASRequestsOps().getIIASData(
            { column: "Modified", isAscending: true },
            props,
            ``
        );

        let MDRCol2 = MDRColl.filter(
            (m: any) =>
                m.NextApproverEmpID === props.EmployeeId[0].EmployeeID ||
                m.DelegateApproverEmpID === props.EmployeeId[0].EmployeeID
        );
        setMDRData(MDRCol2);
        setFilteredData(MDRCol2);
        setLoading(false);
    };

    const resetFilters = () => {
        setColumnFilters({
            ReqNo: "",
            Initiator: "",
            InitDepartment: "",
            Status: "",
            IssueTitle: "",
            // CarLine: "",
            // Total: "",
            // Reason: "",
            NATitle: "",
            DATitle: "",
            LastAction: ""
        });
        setSearchTerm("");
        setCurrentPage(1);
        setFilteredData(MDRData); // restore full dataset
    };

    useEffect(() => {
        GetMDRData();
    }, []);

    // 🔎 Global + Column Filtering
    useEffect(() => {
        let data = [...MDRData];

        // global search
        if (searchTerm) {
            const lowerSearch = searchTerm.toLowerCase();
            data = data.filter((item: any) =>
                item.ID?.toString().toLowerCase().includes(lowerSearch) ||
                item.ReqNo?.toLowerCase().includes(lowerSearch) ||
                item.Initiator?.toLowerCase().includes(lowerSearch) ||
                item.InitDepartment?.toLowerCase().includes(lowerSearch) ||
                item.Status?.toLowerCase().includes(lowerSearch) ||
                item.IssueTitle?.toLowerCase().includes(lowerSearch) ||
                item.NATitle?.toLowerCase().includes(lowerSearch) ||
                item.DATitle?.toLowerCase().includes(lowerSearch) ||
                (item.LastAction &&
                    moment(new Date(item.LastAction))
                        .format("DD-MMM-YYYY")
                        .toLowerCase()
                        .includes(lowerSearch))
            );
        }

        // column filters
        Object.keys(columnFilters).forEach((col) => {
            if (columnFilters[col]) {
                const value = columnFilters[col].toLowerCase();
                data = data.filter((item: any) =>
                    item[col]?.toString().toLowerCase().includes(value)
                );
            }
        });

        setFilteredData(data);
        setCurrentPage(1);
    }, [searchTerm, columnFilters, MDRData]);

    async function EmployeeProfile(Email: string) {
        const spCrudOps = await SPCRUDOPS();
        return spCrudOps.getRootData(
            "UserMaster",
            "EmployeeId,Id,FullName/Title,FullName/ID,FullName/EMail,DirectManagerName/Title,DirectManagerName/ID,DirectManagerName/EMail,OfficeCity/CompanyLocation,OfficeCity/ID,DepartmentCode/Department,DepartmentCode/ID",
            "FullName,DirectManagerName,OfficeCity,DepartmentCode",
            `FullName/EMail eq '${Email}'`,
            { column: "ID", isAscending: true },
            props
        );
    }

    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
        }
    };

    const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

    const paginatedData = sortedData.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage
    );

    // ✅ Export filtered data to Excel
    const exportToExcel = () => {
        if (!filteredData || filteredData.length === 0) {
            alert("No data to export!");
            return;
        }

        const exportData = filteredData.map((item: any) => ({
            "Req No": item.ReqNo,
            "Initiator Name": item.Initiator,
            "Department": item.InitDepartment,
            // "Department": item.Department,
            "Status": item.Status,
            "Issue Title": item.IssueTitle,
            "Next Approver": item.NATitle,
            "Delegated Approver": item.DATitle,
            "Last Action Taken": item.LastAction ? moment(new Date(item.LastAction)).format("DD-MMM-YYYY") : ""
        }));

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "FilteredData");

        const excelBuffer = XLSX.write(workbook, {
            bookType: "xlsx",
            type: "array",
        });
        const data = new Blob([excelBuffer], {
            type: "application/octet-stream",
        });
        saveAs(
            data,
            `Pending_Requests_${new Date().toISOString().slice(0, 10)}.xlsx`
        );
    };

    return (
        <div className="min-h-screen bg-gray-100">
            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>Pending Request Dashboard</h2>
                    </div>
                </div>
            </div>

            <main className="p-6">
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
                            <p className="text-white text-lg">
                                Please wait, loading data...
                            </p>
                        </div>
                    </div>
                ) : (
                    <>
                        {/* 🔎 Search + Export */}
                        <div className="flex items-center gap-4 mb-4">
                            <input
                                type="text"
                                placeholder="Search..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="w-64 px-4 py-2 text-sm border border-gray-300 rounded-full shadow focus:outline-none focus:ring-2 focus:ring-red-500"
                                style={{ width: "250px", marginRight: "20px", marginTop: "20px" }}
                            />

                            {/* Export Button */}
                            <button
                                onClick={exportToExcel}
                                className="btn btn-warning export-btn"
                                style={{ marginBottom: "5px" }}
                            >
                                Export to Excel
                            </button>

                            <i
                                className="fa fa-refresh cursor-pointer text-xl text-gray-700 hover:text-black"
                                onClick={resetFilters}
                                title="Reset Filters"
                                style={{ paddingLeft: "10px" }}
                            ></i>

                        </div>
                        {/* 📊 Table */}
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead
                                        style={{ backgroundColor: "#ce0b0e" }}
                                        className="text-white"
                                    >
                                        <tr>
                                            {/* <th className="px-4 py-2">Req No</th>
                                            <th className="px-4 py-2">Emp No</th>
                                            <th className="px-4 py-2">Emp Name</th>
                                            <th className="px-4 py-2">Department</th>
                                            <th className="px-4 py-2">Cost Center</th>
                                            <th className="px-4 py-2">CarLine</th>
                                            <th className="px-4 py-2">Total Amt.</th>
                                            <th className="px-4 py-2">Reason</th>
                                            <th className="px-4 py-2">Next Approver</th>
                                            <th className="px-4 py-2">Delegated Approver</th> */}
                                            <th className="px-4 py-2">Req No</th>
                                            <th className="px-4 py-2">Initiator Name</th>
                                            <th className="px-4 py-2">Department</th>
                                            <th className="px-4 py-2">Status</th>
                                            <th className="px-4 py-2">Issue Title</th>
                                            <th className="px-4 py-2">Next Approver</th>
                                            <th className="px-4 py-2">Delegated Approver</th>
                                            <th className="px-4 py-2">Last Action Taken</th>
                                        </tr>
                                        {/* 🔽 Column Filters */}
                                        <tr className="bg-gray-100">
                                            {["ReqNo", "Initiator", "InitDepartment", "Status", "IssueTitle", "NATitle", "DATitle", "LastAction"].map((col) => (
                                                <th key={col} className="px-2 py-1">
                                                    <input
                                                        type="text"
                                                        placeholder="Filter"
                                                        value={columnFilters[col] || ""}
                                                        onChange={(e) =>
                                                            setColumnFilters({
                                                                ...columnFilters,
                                                                [col]: e.target.value,
                                                            })
                                                        }
                                                        className="w-full px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-red-400"
                                                    />
                                                </th>
                                            ))}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {paginatedData.map((item, index) => (
                                            <tr key={index} className="border-t">
                                                <td className="px-4 py-2">
                                                    <a
                                                        onClick={(e) => {
                                                            e.preventDefault();
                                                            handleTitleClick(item.ID);
                                                        }}
                                                        href="#"
                                                        className="text-blue-600 hover:text-blue-800 underline"
                                                    >
                                                        {item.ReqNo}
                                                    </a>
                                                </td>
                                                {/* <td className="px-4 py-2">{item.EmpID}</td>
                                                <td className="px-4 py-2">{item.EmpName}</td>
                                                <td className="px-4 py-2">{item.Department}</td>
                                                <td className="px-4 py-2">{item.CostCenter}</td>
                                                <td className="px-4 py-2">{item.CarLine}</td>
                                                <td className="px-4 py-2">{item.Total}</td>
                                                <td className="px-4 py-2">{item.Reason}</td>
                                                <td className="px-4 py-2">{item.NATitle}</td>
                                                <td className="px-4 py-2">{item.DATitle}</td> */}
                                                <td className="px-4 py-2">{item.Initiator}</td>
                                                <td className="px-4 py-2">{item.InitDepartment}</td>
                                                <td className="px-4 py-2">{item.Status}</td>
                                                <td className="px-4 py-2">{item.IssueTitle}</td>
                                                <td className="px-4 py-2">{item.NATitle}</td>
                                                <td className="px-4 py-2">{item.DATitle}</td>
                                                <td className="px-4 py-2">{item.LastAction ? moment(new Date(item.LastAction)).format("DD-MMM-YYYY") : ""}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            {/* 📄 Pagination */}
                            <div className="flex justify-center mt-6 overflow-x-auto">
                                <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-orange rounded shadow">
                                    <button
                                        onClick={() =>
                                            handlePageChange(currentPage - 1)
                                        }
                                        disabled={currentPage === 1}
                                        style={{
                                            backgroundColor: "orange",
                                            color: "black",
                                            opacity: currentPage === 1 ? 0.5 : 1,
                                        }}
                                        className="px-3 py-1 border rounded"
                                    >
                                        Previous
                                    </button>

                                    {currentPage > 3 && (
                                        <>
                                            <button
                                                onClick={() => handlePageChange(1)}
                                                style={{
                                                    backgroundColor: "orange",
                                                    color: "black",
                                                }}
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
                                                style={{
                                                    backgroundColor:
                                                        currentPage === page
                                                            ? "yellow"
                                                            : "orange",
                                                    color: "black",
                                                    fontWeight:
                                                        currentPage === page
                                                            ? "bold"
                                                            : "normal",
                                                }}
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
                                                style={{
                                                    backgroundColor: "orange",
                                                    color: "black",
                                                }}
                                                className="px-3 py-1 border rounded"
                                            >
                                                {totalPages}
                                            </button>
                                        </>
                                    )}

                                    <button
                                        onClick={() => handlePageChange(currentPage + 1)}
                                        disabled={currentPage === totalPages}
                                        style={{
                                            backgroundColor: "orange",
                                            color: "black",
                                            opacity: currentPage === totalPages ? 0.5 : 1,
                                        }}
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

export default Dashboard;
