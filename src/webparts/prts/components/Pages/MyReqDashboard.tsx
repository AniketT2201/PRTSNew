import * as React from "react";
import { useState, useEffect } from "react";
import type { IPrtsProps } from '../IPrtsProps';
import IASRequestsOps from "../../service/BAL/SPCRUD/PRTS";
import SPCRUDOPS from "../../service/DAL/spcrudops";
import { useHistory } from 'react-router-dom';
import "../Pages/CSS/NewRequest.scss";
import '@fortawesome/fontawesome-free/css/all.min.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import * as moment from "moment";
export const MyReqDash: React.FC<IPrtsProps> = (props: IPrtsProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [MDRData, setMDRData] = useState([]);
    const [filteredData, setFilteredData] = useState([]);
    const [loading, setLoading] = useState(true);
    const [columnFilters, setColumnFilters] = useState({
       ID: "",
            ReqNo: "",
            Initiator: "",
            InitDepartment: "",
            Status: "",
            IssueTitle: "",
            NATitle: "",
            DATitle: "",
            LastAction: ""
    });

    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const history = useHistory();

    const totalPages = Math.ceil(filteredData.length / itemsPerPage);

    const handleTitleClick = (id, status) => {
        const path = "/InitiatorLandingedit/"+`${id}`;
        history.push({ pathname: path });
    };

    const dateDifference = (fromDt: Date, toDt: Date) => {
        let diff: any = new Date(toDt.getTime() - fromDt.getTime());
        let days = diff / 1000 / 60 / 60 / 24;
        return days.toFixed(0);
    };

    const formatDate = (dateString) => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-GB');
    };

    const GetMDRData = async () => {
        setLoading(true);
        const MDRColl = await IASRequestsOps().getIIASData(
            { column: "ID", isAscending: true },
            props,
            ''
        );
        //`EmpID eq ` + props.EmployeeId[0].EmployeeID
        let filtereddata = MDRColl.filter((m) => m.InitiatorEmpId === props.EmployeeId[0].EmployeeID)
        setMDRData(filtereddata);
        setFilteredData(filtereddata);
        setLoading(false);
    };

    const resetFilters = () => {
        setColumnFilters({
            ID: "",
            ReqNo: "",
            Initiator: "",
            InitDepartment: "",
            Status: "",
            IssueTitle: "",
            NATitle: "",
            DATitle: "",
            LastAction: ""
        });
        setSearchTerm("");
    };

    useEffect(() => {
        if (MDRData.length === 0) {
            GetMDRData();
        }
    }, []);

    useEffect(() => {
        let filtered = [...MDRData];

        // Column Filters
        Object.keys(columnFilters).forEach((key) => {
            const value = columnFilters[key].toLowerCase();
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

        // Global Search
        if (searchTerm) {
            const lowerSearch = searchTerm.toLowerCase();
            filtered = filtered.filter(item =>
                item.ReqNo?.toLowerCase().includes(lowerSearch) ||
                item.Initiator?.toLowerCase().includes(lowerSearch) ||
                item.InitDepartment?.toLowerCase().includes(lowerSearch) ||
                // formatDate(item.Created)?.toLowerCase().includes(lowerSearch) ||
                item.Status?.toLowerCase().includes(lowerSearch) ||
                item.IssueTitle?.toLowerCase().includes(lowerSearch) ||
                item.NATitle?.toLowerCase().includes(lowerSearch) ||
                item.DATitle?.toLowerCase().includes(lowerSearch) ||
                item.LastAction?.toLowerCase().includes(lowerSearch) 
                // item.Status?.toLowerCase().includes(lowerSearch) ||
                // item.NATitle?.toLowerCase().includes(lowerSearch) ||
                // item.DATitle?.toLowerCase().includes(lowerSearch)
            );
        }

        setFilteredData(filtered);
        setCurrentPage(1);
    }, [columnFilters, searchTerm, MDRData]);    

    const handleColumnFilterChange = (key: string, value: string) => {
        setColumnFilters(prev => ({ ...prev, [key]: value }));
    };

    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
        }
    };

    const dateformate = (Date:Date) =>{

    }
    const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);
    const paginatedData = sortedData.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

    return (
        <div className="min-h-screen bg-gray-100">
            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>My Request Dashboard</h2>
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
                        <div className="flex items-center gap-4 mb-4">
                            <input
                                type="text"
                                placeholder="Search..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="w-64 px-4 py-2 text-sm border border-gray-300 rounded-full shadow focus:outline-none focus:ring-2 focus:ring-red-500"
                                style={{ width: "250px", marginRight: "20px", marginTop: "20px" }}
                            />

                            <i
                                className="fa fa-refresh cursor-pointer text-xl text-gray-700 hover:text-black"
                                onClick={resetFilters}
                                title="Reset Filters"
                            ></i>
                        </div>
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead style={{ backgroundColor: "#ce0b0e" }} className="text-white">
                                        <tr>
                                            <th className="px-4 py-2">Req No</th>
                                            <th className="px-4 py-2">Initiator Name</th>
                                            <th className="px-4 py-2">Department</th>
                                            <th className="px-4 py-2">Status</th>
                                            <th className="px-4 py-2">Issue Title</th>
                                            <th className="px-4 py-2">Next Approver</th>
                                            <th className="px-4 py-2">Delegated Approver</th>
                                            <th className="px-4 py-2">Last Action Taken</th>
                                            {/* <th className="px-4 py-2">Reason</th>
                                            <th className="px-4 py-2">Next Approver</th>
                                            <th className="px-4 py-2">Delegated Approver</th> */}
                                        </tr>
                                        <tr className="bg-gray-100 text-black">
                                            {["ReqNo", "Initiator", "InitDepartment", "Status", "Issue Title", "Next Approver", "Delegated Approver", "Last Action Taken"].map((col) => (
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
                                        {paginatedData.map((item, index) => (
                                            <tr key={index} className="border-t">
                                                {/* <td className="px-4 py-2">{item.Title}</td>*/}
                                                <td className="px-4 py-2"><a onClick={(e) => { e.preventDefault(); handleTitleClick(item.ID, item.Status); }} href="#" className="text-blue-600 hover:text-blue-800 underline">{item.ReqNo}</a></td>
                                                <td className="px-4 py-2">{item.Initiator}</td>
                                                <td className="px-4 py-2">{item.InitDepartment}</td>
                                                <td className="px-4 py-2">{item.Status}</td>
                                                <td className="px-4 py-2">{item.IssueTitle}</td>
                                                <td className="px-4 py-2">{item.NATitle}</td>
                                                <td className="px-4 py-2">{item.DATitle}</td>
                                                 <td className="px-4 py-2">{item.LastAction ? moment(new Date(item.LastAction)).format("DD-MMM-YYYY") : ""}</td>
                                                {/* <td className="px-4 py-2">{item.Total.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                                                <td className="px-4 py-2">{item.Reason}</td>
                                                <td className="px-4 py-2">{item.NATitle}</td>
                                                <td className="px-4 py-2">{item.DATitle}</td> */}
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            {/* Pagination */}
                            <div className="flex justify-center mt-6 overflow-x-auto">
                                <div className="flex space-x-2 px-4 py-2 bg-orange rounded shadow">
                                    <button onClick={() => handlePageChange(currentPage - 1)} disabled={currentPage === 1} className="px-3 py-1 border rounded" style={{ backgroundColor: "orange", color: "black", opacity: currentPage === 1 ? 0.5 : 1 }}>
                                        Previous
                                    </button>
                                    {currentPage > 3 && <><button onClick={() => handlePageChange(1)} className="px-3 py-1 border rounded" style={{ backgroundColor: "orange", color: "black" }}>1</button><span className="px-2">...</span></>}
                                    {Array.from({ length: totalPages }, (_, i) => i + 1).filter(page => Math.abs(page - currentPage) <= 2).map((page) => (
                                        <button key={page} onClick={() => handlePageChange(page)} className="px-3 py-1 border rounded" style={{ backgroundColor: currentPage === page ? "yellow" : "orange", color: "black", fontWeight: currentPage === page ? "bold" : "normal" }}>
                                            {page}
                                        </button>
                                    ))}
                                    {currentPage < totalPages - 2 && <><span className="px-2">...</span><button onClick={() => handlePageChange(totalPages)} className="px-3 py-1 border rounded" style={{ backgroundColor: "orange", color: "black" }}>{totalPages}</button></>}
                                    <button onClick={() => handlePageChange(currentPage + 1)} disabled={currentPage === totalPages} className="px-3 py-1 border rounded" style={{ backgroundColor: "orange", color: "black", opacity: currentPage === totalPages ? 0.5 : 1 }}>
                                        Next
                                    </button>
                                </div>
                            </div>
                        </div>
                    </>
                )
                }
            </main >
        </div >
    );
};

export default MyReqDash;
