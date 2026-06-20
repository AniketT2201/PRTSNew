import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import { format, differenceInDays } from 'date-fns';
import * as XLSX from 'xlsx';
import type { IPrtsProps } from '../IPrtsProps';
import IASRequestsOps from "../../service/BAL/SPCRUD/PRTS";
import { useHistory } from 'react-router-dom';

// Helpers
export const formatDate = (date: Date | string): string => {
  if (!date) return '';
  const parsed = typeof date === 'string' ? new Date(date) : date;
  if (isNaN(parsed.getTime())) return '';
  return format(parsed, 'dd/MM/yyyy');
};

export const dateDifference = (from: Date | string, to: Date | string = new Date()): string => {
  const fromDate = typeof from === 'string' ? new Date(from) : from;
  const toDate = typeof to === 'string' ? new Date(to) : to;
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) return '0';
  return differenceInDays(toDate, fromDate).toString();
};

const PartwiseReport: React.FC<IPrtsProps> = (props: IPrtsProps) => {
  const [MDRData, setMDRData] = useState<any[]>([]);
  const [detailsData, setDetailsData] = useState<any[]>([]);
  const [partNoInput, setPartNoInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [fromDate, setFromDate] = useState('');
  const [toDate, setToDate] = useState('');

  const [columnFilters, setColumnFilters] = useState<Record<string, string>>({
    partno: '',
    Description: '',
    qty: '',
    price: '',
    Total: '',
    RequestDate: '',
    Title: '',
    Status: '',
    Initiator: '',
    NATitle: ''
  });

  const itemsPerPage = 10;
  const history = useHistory();

  // Load data once
  useEffect(() => {
    GetMDRData();
  }, []);

  const GetMDRData = async () => {
    setLoading(true);
    setError('');
    try {
      const result = await IASRequestsOps().getIIASData(
        { column: "Modified", isAscending: true },
        props,
        ``
      );
      setMDRData(result);
    } catch (err) {
      console.error('Error fetching data:', err);
      setError('Failed to fetch data.');
    } finally {
      setLoading(false);
    }
  };

  const safeJsonParse = (str: string, fallback: any) => {
    try {
      return JSON.parse(str);
    } catch {
      return fallback;
    }
  };

  const filterByPartNumber = (data: any[], partNo: string) => {
    const upperPartNo = partNo.toUpperCase();
    const details: any[] = [];

    data.forEach((d) => {
      const detailRows = safeJsonParse(d.Items, []);
      const matched = detailRows.filter((row: any) => {
        const val = row?.c1;
        const c = val !== null && val !== undefined ? String(val).toUpperCase() : '';
        return c === upperPartNo;
      });

      matched.forEach((m: any, idx: number) => {
        const item = {
          partno: m.c1 ?? '',
          Description: m.c2 ?? '',
          qty: m.c4 ?? '',
          price: m.c5 ?? '',
          Total: m.c6 ?? '',
          ID: d.ID,
          RequestDate: d.Created,
          Title: d.Title,
          Status: d.Status,
          Initiator: d.EmpName,
          NATitle: d.NATitle,
          _rowKey: `${d.ID}-${idx}` // unique key
        };
        details.push(item);
      });
    });

    return details;
  };

  // Handle search
  const handleSearch = () => {
    const trimmedPart = partNoInput.trim();
    if (!trimmedPart) {
      setDetailsData([]);
      return;
    }

    let filteredDATA = filterByPartNumber(MDRData, trimmedPart);

    // Date filters
    if (fromDate && toDate && fromDate === toDate) {
      const targetDate = new Date(fromDate).toDateString();
      filteredDATA = filteredDATA.filter(
        item => new Date(item.RequestDate).toDateString() === targetDate
      );
    } else if (fromDate || toDate) {
      const from = fromDate ? new Date(fromDate) : null;
      const to = toDate ? new Date(toDate) : null;
      filteredDATA = filteredDATA.filter(item => {
        const itemDate = new Date(item.RequestDate);
        return (!from || itemDate >= from) && (!to || itemDate <= to);
      });
    }

    setDetailsData(filteredDATA);
    setCurrentPage(1);

    // reset column filters
    setColumnFilters({
      partno: '',
      Description: '',
      qty: '',
      price: '',
      Total: '',
      RequestDate: '',
      Title: '',
      Status: '',
      Initiator: '',
      NATitle: ''
    });
  };

  // Apply column filters on top of search
  const finalFilteredData = useMemo(() => {
    if (!detailsData || detailsData.length === 0) return [];
    return detailsData.filter(item =>
      Object.entries(columnFilters).every(([key, value]) => {
        if (!value) return true;
        const itemVal = String(item[key] ?? '').toLowerCase();
        return itemVal.includes(value.toLowerCase());
      })
    );
  }, [detailsData, columnFilters]);

  // Pagination
  const totalPages = Math.max(1, Math.ceil(finalFilteredData.length / itemsPerPage));
  const paginatedData = finalFilteredData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const resetFilters = () => {
    setColumnFilters({
      partno: '',
      Description: '',
      qty: '',
      price: '',
      Total: '',
      RequestDate: '',
      Title: '',
      Status: '',
      Initiator: '',
      NATitle: ''
    });
    setPartNoInput('');
    setFromDate('');
    setToDate('');
    setDetailsData([]);
    setCurrentPage(1);
  };

  const handleTitleClick = (id: string) => {
    history.push({
      pathname: '/ApprovalForm',
      search: `?ItemId=${id}`,
      state: { from: '/PartwiseReport' }
    });
  };

  // ✅ Export uses finalFilteredData (ignores pagination)
  const exportToExcel = () => {
    if (finalFilteredData.length === 0) return;

    const formatDateForExcel = (date: any) => {
      if (!date) return "";
      const d = new Date(date);
      if (isNaN(d.getTime())) return "";
      const day = String(d.getDate()).padStart(2, "0");
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const year = d.getFullYear();
      return `${day}/${month}/${year}`;
    };

    const cleanData = finalFilteredData.map(({ _rowKey, RequestDate, ...rest }) => ({
      ...rest,
      RequestDate: formatDateForExcel(RequestDate),
    }));

    const worksheet = XLSX.utils.json_to_sheet(cleanData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "FilteredReport");
    XLSX.writeFile(workbook, "PartWiseReport.xlsx");
  };


  // ---- UI ----
  return (
    <div className="min-h-screen bg-gray-100">
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>Part Wise Report</h2>
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
            {/* Search Bar and Export Button */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '10px 0px 10px 0px' }}>
              <div className="flex flex-col">
                <label htmlFor="partNo" className="text-sm mb-1">Enter Part No</label>
                <input
                  id="partNo"
                  type="text"
                  placeholder="Enter Part Number"
                  value={partNoInput}
                  onChange={(e) => setPartNoInput(e.target.value)}
                  className="w-[250px]  h-[32px] text-sm border-gray-300 rounded dashboard-sha focus:outline-none focus:ring-2 focus:ring-red-500"
                />
              </div>
              <div className="flex flex-col">
                <label htmlFor="fromDate" className="text-sm mb-1">From Date</label>
                <input
                  id="fromDate"
                  type="date"
                  value={fromDate}
                  onChange={(e) => setFromDate(e.target.value)}
                  className="w-[250px]  h-[32px] text-sm border-gray-300 rounded dashboard-sha focus:outline-none focus:ring-2 focus:ring-red-500"
                />
              </div>
              <div className="flex flex-col">
                <label htmlFor="toDate" className="text-sm mb-1">End Date</label>
                <input
                  id="toDate"
                  type="date"
                  value={toDate}
                  onChange={(e) => setToDate(e.target.value)}
                  className="w-[250px]  h-[32px] text-sm border-gray-300 rounded dashboard-sha focus:outline-none focus:ring-2 focus:ring-red-500"
                />
              </div>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <button className="btn btn-success btn-approver" onClick={handleSearch}>
                Search
              </button>
              {finalFilteredData.length > 0 && (
                <button className="btn btn-warning export-btn" type="button" onClick={exportToExcel}>
                  Export Data
                </button>
              )}
              <i
                className="fa fa-refresh cursor-pointer text-xl text-gray-700 hover:text-black"
                onClick={resetFilters}
                title="Reset Filters"
                style={{ height: '32px', display: 'flex', alignItems: 'center' }}
              ></i>
            </div>

            {/* Table */}
            <div className="overflow-x-auto" style={{ paddingTop: "20px" }}>
              <div className="table-vert-scroll">
                <table className="min-w-full bg-white rounded-2xl shadow-md">
                  <thead style={{ backgroundColor: "#ce0b0e" }} className="text-white">
                    <tr>
                      <th className="px-4 py-2">Part Number</th>
                      <th className="px-4 py-2">Description</th>
                      <th className="px-4 py-2">Qty</th>
                      <th className="px-4 py-2">Price</th>
                      <th className="px-4 py-2">Total</th>
                      <th className="px-4 py-2">Req No</th>
                      <th className="px-4 py-2">Req Date</th>
                      <th className="px-4 py-2">Status</th>
                      <th className="px-4 py-2">Initiator</th>
                      <th className="px-4 py-2">Next Approver</th>
                    </tr>
                    <tr className="bg-gray-100 text-black">
                      {["partno", "Description", "qty", "price", "Total", "RequestDate", "Title", "Status", "Initiator", "NATitle"].map((col) => (
                        <th key={col} className="px-4 py-1">
                          <input
                            type="text"
                            value={columnFilters[col]}
                            onChange={(e) => setColumnFilters(prev => ({ ...prev, [col]: e.target.value }))}
                            className="w-full px-2 py-1 text-xs border border-gray-300 rounded"
                            placeholder="Search"
                            style={{ width: "140px" }}
                          />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedData.map((item) => (
                      <tr key={item._rowKey} className="border-t">
                        <td className="px-4 py-2">{item.partno}</td>
                        <td className="px-4 py-2">{item.Description}</td>
                        <td className="px-4 py-2">{item.qty}</td>
                        <td className="px-4 py-2">{item.price}</td>
                        <td className="px-4 py-2">{item.Total}</td>
                        <td className="px-4 py-2">
                          <a
                            href="#"
                            onClick={(e) => { e.preventDefault(); handleTitleClick(item.ID); }}
                            className="text-blue-600 hover:text-blue-800 underline"
                          >
                            {item.Title}
                          </a>
                        </td>
                        <td className="px-4 py-2">{formatDate(item.RequestDate)}</td>
                        <td className="px-4 py-2">{item.Status}</td>
                        <td className="px-4 py-2">{item.Initiator}</td>
                        <td className="px-4 py-2">{item.NATitle}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              {totalPages > 1 && (
                <div className="flex justify-center mt-6 overflow-x-auto">
                  <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-orange rounded shadow">
                    <button
                      onClick={() => handlePageChange(currentPage - 1)}
                      disabled={currentPage === 1}
                      className="px-3 py-1 border rounded"
                      style={{ backgroundColor: 'orange', opacity: currentPage === 1 ? 0.5 : 1 }}
                    >
                      Previous
                    </button>
                    {currentPage > 3 && (
                      <>
                        <button onClick={() => handlePageChange(1)} className="px-3 py-1 border rounded" style={{ backgroundColor: 'orange' }}>1</button>
                        <span className="px-2">...</span>
                      </>
                    )}
                    {Array.from({ length: totalPages }, (_, i) => i + 1)
                      .filter(page => Math.abs(page - currentPage) <= 2)
                      .map(page => (
                        <button
                          key={page}
                          onClick={() => handlePageChange(page)}
                          className="px-3 py-1 border rounded"
                          style={{
                            backgroundColor: currentPage === page ? 'yellow' : 'orange',
                            fontWeight: currentPage === page ? 'bold' : 'normal'
                          }}
                        >
                          {page}
                        </button>
                      ))}
                    {currentPage < totalPages - 2 && (
                      <>
                        <span className="px-2">...</span>
                        <button onClick={() => handlePageChange(totalPages)} className="px-3 py-1 border rounded" style={{ backgroundColor: 'orange' }}>{totalPages}</button>
                      </>
                    )}
                    <button
                      onClick={() => handlePageChange(currentPage + 1)}
                      disabled={currentPage === totalPages}
                      className="px-3 py-1 border rounded"
                      style={{ backgroundColor: 'orange', opacity: currentPage === totalPages ? 0.5 : 1 }}
                    >
                      Next
                    </button>
                  </div>
                </div>
              )}
            </div>
          </>
        )}
      </main>
    </div>
  );
};

export default PartwiseReport;
