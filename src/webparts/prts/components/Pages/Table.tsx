import * as React from 'react';
import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Label,
  Stack,
} from '@fluentui/react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from '@fluentui/react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// ---------- Interfaces ----------
interface IRow {
  key: number;
  partNumber: string;
  description: string;
  supplier: string;
  qty: number;
  value: number;
  amount: number;
  remarks: string;
}

const expectedHeaders = [
  'Part number',
  'Part Description',
  'Supplier',
  'Qty',
  'Value',
  'Amount',
  'Remarks',
];

interface TableProps {
  initialData?: IRow[];
  onDataChange?: (data: IRow[]) => void;
}

export interface TableRef {
  getData: () => IRow[];
  resetData: () => void;
  setData: (data: IRow[]) => void;
}

const Table = React.forwardRef<TableRef, TableProps>((props, ref) => {
  const { initialData = [], onDataChange } = props;

  // ---------- State ----------
  const [data, setData] = React.useState<IRow[]>([]);
  const [showPanel, setShowPanel] = React.useState(false);
  const [formErrors, setFormErrors] = React.useState<{ [key: string]: string }>({});
  const [formData, setFormData] = React.useState<Omit<IRow, 'key'>>({
    partNumber: '',
    description: '',
    supplier: '',
    qty: 0,
    value: 0,
    amount: 0,
    remarks: '',
  });

  const [currentPage, setCurrentPage] = React.useState(1);
  const pageSize = 10;
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const totalPages = Math.ceil(data.length / pageSize);
  const paginatedData = data.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // ---------- Effects ----------
  React.useEffect(() => {
    if (initialData.length > 0) {
      const cloned = initialData.map((row, idx) => ({ ...row, key: idx + 1 }));
      setData(cloned);
    }
  }, []);

  React.useImperativeHandle(ref, () => ({
    getData: () => data,
    resetData: () => setData([]),
    setData: (newData: IRow[]) => {
      const cloned = newData.map((row, idx) => ({ ...row, key: idx + 1 }));
      setData(cloned);
    },
  }));

  React.useEffect(() => {
    onDataChange?.(data);
  }, [data]);

  // ---------- Handlers ----------
  const handleAdd = () => {
    const amount = formData.qty * formData.value;

    const newRow: IRow = {
      key: data.length + 1,
      ...formData,
      amount,
    };

    setData([...data, newRow]);
    setShowPanel(false);
    setFormData({
      partNumber: '',
      description: '',
      supplier: '',
      qty: 0,
      value: 0,
      amount: 0,
      remarks: '',
    });
  };

  const handleEdit = (key: number, field: keyof IRow, value: string | number) => {
    const updated = data.map((row) => {
      if (row.key === key) {
        const updatedRow = { ...row, [field]: value };
        updatedRow.amount = Number(updatedRow.qty) * Number(updatedRow.value);
        return updatedRow;
      }
      return row;
    });
    setData(updated);
  };

  const handleDelete = (key: number) => {
    if (!window.confirm("Are you sure you want to delete this row?")) return;
    const updated = data
      .filter((row) => row.key !== key)
      .map((row, index) => ({ ...row, key: index + 1 }));
    setData(updated);
  };

  const handleDownloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([expectedHeaders]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'ImportTemplate.xlsx');
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const bstr = event.target?.result;
      if (!bstr) return;

      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as string[][];

      const headers = jsonData[0];
      const isValid = expectedHeaders.every((h, i) => h === headers[i]);
      if (!isValid) {
        alert('Uploaded Excel does not match the expected template.');
        return;
      }

      const dataRows = jsonData.slice(1).filter((row) => row.length > 0);
      const newRows: Omit<IRow, 'key'>[] = dataRows.map((row) => {
        const qty = Number(row[3]) || 0;
        const value = Number(row[4]) || 0;
        const amount = qty * value;
        return {
          partNumber: row[0] || '',
          description: row[1] || '',
          supplier: row[2] || '',
          qty,
          value,
          amount,
          remarks: row[6] || '',
        };
      });

      const combined = [...data, ...newRows].map((row, idx) => ({
        ...row,
        key: idx + 1,
      }));

      setData(combined);
      e.target.value = '';
    };

    reader.readAsBinaryString(file);
  };

  // ---------- Columns ----------
  const columns: IColumn[] = [
    {
      key: 'srNo', name: 'Sr. No', fieldName: 'key', minWidth: 40,
      onRender: (_, index) => <span>{(currentPage - 1) * pageSize + index + 1}</span>
    },
    {
      key: 'partNumber', name: 'Part Number', fieldName: 'partNumber', minWidth: 100,
      onRender: (item: IRow) => (
        <TextField value={item.partNumber}
          onChange={(e, val) => handleEdit(item.key, 'partNumber', val || '')} />
      )
    },
    {
      key: 'description', name: 'Part Description', fieldName: 'description', minWidth: 200,
      onRender: (item: IRow) => (
        <TextField multiline rows={3} value={item.description}
          onChange={(e, val) => handleEdit(item.key, 'description', val || '')} />
      )
    },
    {
      key: 'supplier', name: 'Supplier', fieldName: 'supplier', minWidth: 100,
      onRender: (item: IRow) => (
        <TextField value={item.supplier}
          onChange={(e, val) => handleEdit(item.key, 'supplier', val || '')} />
      )
    },
    {
      key: 'qty', name: 'Qty', fieldName: 'qty', minWidth: 100,
      onRender: (item: IRow) => (
        <input type="number" step={0.1} value={item.qty}
          onChange={(e) => handleEdit(item.key, "qty", parseFloat(e.target.value) || 0)}
          className="form-control" />
      )
    },
    {
      key: 'value', name: 'Value', fieldName: 'value', minWidth: 100,
      onRender: (item: IRow) => (
        <input type="number" step={0.1} value={item.value}
          onChange={(e) => handleEdit(item.key, "value", parseFloat(e.target.value) || 0)}
          className="form-control" />
      )
    },
    {
      key: 'amount', name: 'Amount', fieldName: 'amount', minWidth: 100,
      onRender: (item: IRow) => <span>{(item.qty * item.value).toFixed(2)}</span>
    },
    {
      key: 'remarks', name: 'Remarks', fieldName: 'remarks', minWidth: 150,
      onRender: (item: IRow) => (
        <TextField value={item.remarks}
          onChange={(e, val) => handleEdit(item.key, 'remarks', val || '')} />
      )
    },
    {
      key: 'delete', name: 'Delete', fieldName: '', minWidth: 60,
      onRender: (item: IRow) => (
        <i className="fa fa-trash" style={{ cursor: 'pointer', color: 'red', fontSize: 16 }}
          onClick={() => handleDelete(item.key)} />
      )
    }
  ];

  const totalAmount = data.reduce((sum, row) => sum + (row.amount || 0), 0);

  return (
    <div>
      {/* Action buttons */}
      <PrimaryButton text="Add Row" onClick={() => setShowPanel(true)} style={{ marginRight: 10, backgroundColor: "rgb(3, 120, 124)", border: "1px solid rgb(3, 120, 124)", color: "#fff" }} />
      <PrimaryButton text="Export Template" onClick={handleDownloadTemplate} style={{ marginRight: 10, backgroundColor: "rgb(3, 120, 124)", border: "1px solid rgb(3, 120, 124)", color: "#fff" }} />
      <PrimaryButton text="Import Excel" onClick={() => fileInputRef.current?.click()} style={{ marginRight: 10, backgroundColor: "rgb(3, 120, 124)", border: "1px solid rgb(3, 120, 124)", color: "#fff" }} />
      <PrimaryButton text="Reset" onClick={() => setData([])} style={{ marginRight: 10, backgroundColor: "rgb(3, 120, 124)", border: "1px solid rgb(3, 120, 124)", color: "#fff" }} />

      {/* Hidden file input */}
      <input type="file" accept=".xlsx, .xls" ref={fileInputRef} style={{ display: 'none' }} onChange={handleImportExcel} />

      {/* Modal form */}
      {showPanel && (
        <>
          <div className="modal fade show d-block m-0" tabIndex={-1} role="dialog" aria-hidden="false">
            <div className="modal-dialog modal-lg">
              <div className="modal-content">
                <div className="modal-body">
                  <span className="h4 required">Add Row</span>

                  <TextField
                    label="Part Number"
                    value={formData.partNumber}
                    onChange={(e, val) => setFormData({ ...formData, partNumber: val || '' })}
                    errorMessage={formErrors.partNumber}
                  />

                  <TextField
                    label="Part Description"
                    multiline
                    rows={3}
                    value={formData.description}
                    onChange={(e, val) => setFormData({ ...formData, description: val || '' })}
                    errorMessage={formErrors.description}
                  />

                  <TextField
                    label="Supplier"
                    value={formData.supplier}
                    onChange={(e, val) => setFormData({ ...formData, supplier: val || '' })}
                    errorMessage={formErrors.supplier}
                  />

                  {/* ✅ Qty with decimal support */}
                  <TextField
                    label="Qty"
                    value={formData.qty.toString()}
                    type="number"
                    onChange={(e, val) => {
                      const newQty = val ?? "";
                      const numQty = parseFloat(newQty);
                      setFormData({
                        ...formData,
                        qty: newQty as any, // keep string while typing
                        amount: (isNaN(numQty) ? 0 : numQty) * (parseFloat(formData.value as any) || 0),
                      });
                    }}
                    errorMessage={formErrors.qty}
                  />

                  {/* ✅ Value with decimal support */}
                  <TextField
                    label="Value"
                    value={formData.value.toString()}
                    type="number"
                    onChange={(e, val) => {
                      const newValue = val ?? "";
                      const numValue = parseFloat(newValue);
                      setFormData({
                        ...formData,
                        value: newValue as any, // keep string while typing
                        amount: (parseFloat(formData.qty as any) || 0) * (isNaN(numValue) ? 0 : numValue),
                      });
                    }}
                    errorMessage={formErrors.value}
                  />

                  <TextField
                    label="Amount"
                    value={formData.amount.toFixed(2)}
                    readOnly
                  />

                  <TextField
                    label="Remarks"
                    value={formData.remarks}
                    onChange={(e, val) => setFormData({ ...formData, remarks: val || '' })}
                  />
                </div>

                <div className="modal-footer">
                  <button
                    className="btn btn-primary"
                    onClick={() => {
                      const errors: { [key: string]: string } = {};
                      const qtyNum = parseFloat(formData.qty as any) || 0;
                      const valueNum = parseFloat(formData.value as any) || 0;

                      if (!formData.partNumber.trim()) errors.partNumber = 'Part Number is required';
                      if (!formData.description.trim()) errors.description = 'Description is required';
                      if (!formData.supplier.trim()) errors.supplier = 'Supplier is required';
                      if (qtyNum <= 0) errors.qty = 'Qty must be greater than 0';
                      if (valueNum <= 0) errors.value = 'Value must be greater than 0';

                      setFormErrors(errors);
                      if (Object.keys(errors).length > 0) return;

                      handleAdd();
                      setFormErrors({});
                    }}
                  >
                    OK
                  </button>
                  <button className="btn btn-secondary" onClick={() => setShowPanel(false)}>Cancel</button>
                </div>
              </div>
            </div>
          </div>
          <div className="modal-backdrop fade show"></div>
        </>
      )}

      {/* Data list */}
      <DetailsList
        items={paginatedData}
        columns={columns}
        setKey="editable"
        layoutMode={DetailsListLayoutMode.fixedColumns}
        selectionMode={SelectionMode.none}
      />

      {/* Footer total */}
      <div style={{ display: 'flex', marginTop: 10, fontWeight: 'bold' }}>
        <div style={{ minWidth: 40 }}></div>
        <div style={{ minWidth: 100 }}></div>
        <div style={{ minWidth: 200 }}></div>
        <div style={{ minWidth: 100 }}></div>
        <div style={{ minWidth: 80 }}></div>
        <div style={{ minWidth: 80 }}></div>
        <div style={{ minWidth: 80, textAlign: 'center' }}>Total</div>
        <div style={{ minWidth: 150, textAlign: 'center' }}>{totalAmount.toFixed(2)}</div>
        <div style={{ minWidth: 60 }}></div>
      </div>

      {/* Pagination */}
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 20, flexWrap: 'wrap' } }}>
        <DefaultButton text="Previous" onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))} disabled={currentPage === 1} />
        <Label styles={{ root: { paddingTop: 6 } }}>Page {currentPage} of {totalPages}</Label>
        <DefaultButton text="Next" onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))} disabled={currentPage === totalPages || totalPages === 0} />
      </Stack>
    </div>
  );
});

export default Table;
