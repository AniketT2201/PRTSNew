import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { format, differenceInDays } from 'date-fns';
import { Formik, Form, Field, FormikProps, useFormikContext } from "formik";
import * as XLSX from 'xlsx';
import type { IPrtsProps } from '../IPrtsProps';
import '../Pages/CSS/NewRequest.scss';
import SPCRUDOPS from "../../service/DAL/spcrudops";
import { useHistory } from 'react-router-dom';
import { ReorderingRuleMatchType } from '@pnp/sp/search';
import { RoleType } from '@pnp/sp/sharing';
import USESPCRUD, { ISPCRUD } from '../../service/BAL/SPCRUD/spcrud';
import Select from 'react-select';
import { values } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp/presets/all";

interface FormValues {
    user: any;
    Department: any;
    RoleType?: string;
    roleType: any;
}

export const formatDate = (date: Date | string): string => {
    const parsed = typeof date === 'string' ? new Date(date) : date;
    return format(parsed, 'dd/MM/yyyy');
};

export const dateDifference = (from: Date | string, to: Date | string = new Date()): string => {
    const fromDate = typeof from === 'string' ? new Date(from) : from;
    const toDate = typeof to === 'string' ? new Date(to) : to;
    return differenceInDays(toDate, fromDate).toString();
};

export const WarehouseUser: React.FC<IPrtsProps> = (props: IPrtsProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [MCData, setMCData] = useState([]);
    const [filteredData, setFilteredData] = useState([]);
    const [loading, setLoading] = useState(true);
    const [popupVisible, setPopupVisible] = useState(false);
    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const [selectedUsers, setSelectedUsers] = useState(0);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);
    const history = useHistory();

    async function GetWarehouseControllersData() {
        setLoading(true);
        const spCrudOps = await SPCRUDOPS();
        const Mtrldata = await spCrudOps.getData(
            'WarehouseControllers',
            'UserName/Title,UserName/EMail,ID',
            'UserName',
            '',
            { column: 'ID', isAscending: true },
            props
        );
        setMCData(Mtrldata);
        setLoading(false);
    }

    useEffect(() => {
        GetWarehouseControllersData();
    }, [])

    useEffect(() => {
        if (!searchTerm) {
            setFilteredData(MCData);
        } else {
            const lowerSearch = searchTerm.toLowerCase();
            const filtered = MCData.filter(item =>
                item.UserName?.Title?.toString().toLowerCase().includes(lowerSearch)
            );
            setFilteredData(filtered);
            setCurrentPage(1);
        }
    }, [searchTerm, MCData]);

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

    const handleEditClick = () => {
        setPopupVisible(true);
    };

    const _getPeoplePickerItems = async (items: any[]) => {
        if (items.length > 0) {
            const loginName = items[0].loginName;

            try {
                sp.setup({
                    sp: {
                        baseUrl: props.currentSPContext.pageContext.web.absoluteUrl
                    },
                });
                const user = await sp.web.ensureUser(loginName);
                setSelectedUsers(user.data.Id)
            } catch (err) {
                console.error("Failed to get user:", err);
            }
        }
    };

    const update = async () => {
        try {
            if (!selectedUsers || selectedUsers === 0) {
                alert("Please select a user.");
                return;
            }

            const fields = {
                UserNameId: selectedUsers
            };

            const spCrudObj = await USESPCRUD();

            const result = await spCrudObj.insertData('WarehouseControllers', fields, props);
            console.log("Insert result:", result);

            await GetWarehouseControllersData();

            alert("Data updated successfully.");
            setPopupVisible(false);
        } catch (error) {
            console.error("Error in update():", error);
            alert("An error occurred while updating data.");
        }
    };

    const DeleteUser = async (item) => {
        try {
            const spCrudObj = await USESPCRUD();

            const result = await spCrudObj.deleteData('WarehouseControllers', item, props);
            console.log("Delete result:", result);

            alert("Data deleted successfully.");
            await GetWarehouseControllersData(); // Optional: refresh list after delete
        } catch (error) {
            console.error("Error in DeleteUser():", error);
            alert("An error occurred while deleting data.");
        }
    };

    return (
        <div className="min-h-screen bg-gray-100">
            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>Configure WareHouse Controller</h2>
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
                            <p className="text-white text-lg">Please wait, loading data...</p>
                        </div>
                    </div>
                ) : (
                    <>
                        {/* Search Bar */}
                        <div
                            className="Dashboard-Search-New"
                            style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "10px", marginBottom: "10px", width: "300px" }}
                        >
                            <input
                                type="text"
                                placeholder="Search..."
                                className="w-64 px-4 py-2 text-sm border border-gray-300 rounded-full shadow focus:outline-none focus:ring-2 focus:ring-red-500"
                                value=""
                            />
                            <button
                                type="button"
                                title="Add"
                                onClick={handleEditClick}
                                style={{
                                    color: "green",
                                    fontSize: "24px",
                                    background: "transparent",
                                    border: "none",
                                    cursor: "pointer",
                                    paddingBottom: "4px",
                                }}
                            >
                                ➕
                            </button>
                        </div>

                        {/* Table */}
                        <div className="overflow-x-auto">
                            <div className="table-vert-scroll">
                                <table className="min-w-full bg-white rounded-2xl shadow-md">
                                    <thead style={{ backgroundColor: "#ce0b0e" }}
                                        className="text-white"
                                    >
                                        <tr>
                                            <th className="px-4 py-2">Action</th>
                                            <th className="px-4 py-2">User Name</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {paginatedData.map((item, index) => (
                                            <tr key={index} className="border-t">
                                                <td className="px-4 py-2"><button
                                                    onClick={() => DeleteUser(item.ID)}
                                                    title="Remove"
                                                    style={{
                                                        background: 'transparent',
                                                        border: 'none',
                                                        fontSize: '20px',
                                                        cursor: 'pointer',
                                                        padding: 0
                                                    }}
                                                >
                                                    ❌
                                                </button></td>
                                                <td className="px-4 py-2">{item.UserName.Title}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            {/* Pagination */}
                            <div className="flex justify-center mt-6 overflow-x-auto">
                                <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-orange rounded shadow">

                                    {/* Previous Button */}
                                    <button
                                        onClick={() => handlePageChange(currentPage - 1)}
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

                                    {/* First Page Shortcut */}
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

                                    {/* Main Page Numbers */}
                                    {Array.from({ length: totalPages }, (_, i) => i + 1)
                                        .filter((page) => Math.abs(page - currentPage) <= 2)
                                        .map((page) => (
                                            <button
                                                key={page}
                                                onClick={() => handlePageChange(page)}
                                                style={{
                                                    backgroundColor: currentPage === page ? "yellow" : "orange",
                                                    color: "black",
                                                    fontWeight: currentPage === page ? "bold" : "normal"
                                                }}
                                                className="px-3 py-1 border rounded"
                                            >
                                                {page}
                                            </button>
                                        ))}

                                    {/* Last Page Shortcut */}
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

                                    {/* Next Button */}
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

                        {popupVisible && (
                            <>
                                <div className="modal fade show d-block" tabIndex={-1} role="dialog" aria-hidden="false">
                                    <div className="modal-dialog modal-lg">
                                        <div className="modal-content">
                                            <div className="modal-body">
                                                <table className="table table-bordered">
                                                    <colgroup>
                                                        <col style={{ width: '30%' }} />
                                                        <col style={{ width: '70%' }} />
                                                    </colgroup>
                                                    <tbody>
                                                        <tr>
                                                            <th>Select User</th>
                                                            <td>
                                                                <PeoplePicker
                                                                    context={props.context}
                                                                    personSelectionLimit={1}
                                                                    groupName={""}
                                                                    showtooltip={true}
                                                                    onChange={_getPeoplePickerItems}
                                                                    principalTypes={[PrincipalType.User]}
                                                                    resolveDelay={1000}
                                                                />
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <div className="modal-footer">
                                                <button type="button" className="btn btn-primary" onClick={update}>OK</button>
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
        </div >
    );
};

export default WarehouseUser;