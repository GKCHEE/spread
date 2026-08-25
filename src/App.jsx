import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, Download, Search, FileSpreadsheet, Filter } from 'lucide-react';

export default function App() {
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [fileName, setFileName] = useState('');
  
  // Filter States
  const [searchSupplier, setSearchSupplier] = useState('');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');

  // Handle File Upload and Parse Excel/CSV
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      const binaryString = evt.target.result;
      const workbook = XLSX.read(binaryString, { type: 'binary' });
      
      // Get first worksheet
      const worksheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[worksheetName];
      
      // Convert to JSON (array of objects)
      const parsedData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
      
      if (parsedData.length > 0) {
        // Extract column headers (ignoring the massive amount of empty trailing columns)
        const allKeys = Object.keys(parsedData[0]);
        const usefulColumns = allKeys.filter(key => !key.includes('__EMPTY'));
        
        setColumns(usefulColumns);
        setData(parsedData);
      }
    };

    reader.readAsBinaryString(file);
  };

  // Filter Logic
  const filteredData = data.filter((row) => {
    // We assume your column names are exactly as provided in your CSV: "Supplier", "Outstanding Amount"
    const supplier = row['Supplier'] || row['Supplier '] || '';
    const outAmount = parseFloat(row['Outstanding Amount'] || row['Outstanding'] || 0);

    const matchSupplier = supplier.toString().toLowerCase().includes(searchSupplier.toLowerCase());
    
    const min = parseFloat(minAmount);
    const max = parseFloat(maxAmount);
    
    const matchMin = isNaN(min) ? true : outAmount >= min;
    const matchMax = isNaN(max) ? true : outAmount <= max;

    // Filter out rows that are entirely empty or just total rows
    const isTotalRow = supplier.toString().toUpperCase().includes('TOTAL');

    return matchSupplier && matchMin && matchMax && !isTotalRow;
  });

  // Export Filtered Data back to Excel
  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "AP_Report");
    
    // Generate buffer and trigger download
    XLSX.writeFile(workbook, `AP_Filtered_Report_${new Date().getTime()}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Header */}
        <header className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4">
          <div>
            <h1 className="text-2xl font-black text-slate-900 flex items-center gap-2">
              <FileSpreadsheet className="text-blue-600 w-8 h-8" />
              Accounts Payable Portal
            </h1>
            <p className="text-slate-500 font-medium mt-1">Upload, Filter, and Export AP Excel Files securely in your browser.</p>
          </div>
          
          <div className="relative group cursor-pointer">
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv" 
              onChange={handleFileUpload} 
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className="bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-600 hover:text-white transition-colors px-6 py-3 rounded-xl font-bold flex items-center gap-2">
              <UploadCloud className="w-5 h-5" />
              {fileName ? 'Replace File' : 'Drop Excel File Here'}
            </div>
          </div>
        </header>

        {/* Workspace - Only shows if data is loaded */}
        {data.length > 0 && (
          <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
            
            {/* Sidebar Filters */}
            <aside className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 h-fit space-y-6">
              <div className="flex items-center gap-2 mb-4">
                <Filter className="w-5 h-5 text-slate-500" />
                <h2 className="font-bold text-lg">Report Filters</h2>
              </div>

              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Search Supplier Name</label>
                <div className="relative">
                  <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                  <input 
                    type="text" 
                    value={searchSupplier}
                    onChange={(e) => setSearchSupplier(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500"
                    placeholder="e.g. COWAY"
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-bold text-slate-700 mb-2">Outstanding Amount Range</label>
                <div className="flex items-center gap-2">
                  <input 
                    type="number" 
                    value={minAmount}
                    onChange={(e) => setMinAmount(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-blue-500"
                    placeholder="Min"
                  />
                  <span className="text-slate-400">-</span>
                  <input 
                    type="number" 
                    value={maxAmount}
                    onChange={(e) => setMaxAmount(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-blue-500"
                    placeholder="Max"
                  />
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100">
                <button 
                  onClick={exportToExcel}
                  className="w-full bg-emerald-500 hover:bg-emerald-600 text-white font-bold py-3 rounded-xl flex items-center justify-center gap-2 transition-colors"
                >
                  <Download className="w-5 h-5" />
                  Export Report to Excel
                </button>
              </div>
            </aside>

            {/* Data Table */}
            <main className="lg:col-span-3 bg-white p-6 rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
              <div className="flex justify-between items-center mb-6">
                <h2 className="font-bold text-lg flex items-center gap-2">
                  Dataset Viewer
                  <span className="bg-slate-100 text-slate-600 text-xs py-1 px-2 rounded-full">
                    {filteredData.length} Records found
                  </span>
                </h2>
                <div className="text-sm font-medium text-slate-500">
                  File: <span className="text-slate-800">{fileName}</span>
                </div>
              </div>

              <div className="overflow-x-auto overflow-y-auto max-h-[600px] border border-slate-200 rounded-xl relative custom-scrollbar">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                    <tr>
                      {columns.map((col, idx) => (
                        <th key={idx} className="p-4 font-bold text-slate-700 border-b border-slate-200 whitespace-nowrap">
                          {col}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, rowIdx) => (
                      <tr key={rowIdx} className="hover:bg-blue-50/50 transition-colors border-b border-slate-100 last:border-0">
                        {columns.map((col, colIdx) => (
                          <td key={colIdx} className="p-4 text-slate-600 whitespace-nowrap">
                            {row[col] !== undefined && row[col] !== null ? row[col].toString() : '-'}
                          </td>
                        ))}
                      </tr>
                    ))}
                    {filteredData.length === 0 && (
                      <tr>
                        <td colSpan={columns.length} className="p-8 text-center text-slate-400 font-medium">
                          No records match your current filters.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </main>

          </div>
        )}

        {/* Empty State Instructions */}
        {data.length === 0 && (
          <div className="bg-white border-2 border-dashed border-slate-200 rounded-2xl p-16 text-center">
            <FileSpreadsheet className="w-16 h-16 text-slate-300 mx-auto mb-4" />
            <h3 className="text-xl font-bold text-slate-700 mb-2">No File Loaded</h3>
            <p className="text-slate-500 max-w-md mx-auto">
              Please drop your monthly Accounts Payable Excel or CSV file into the upload area above to begin filtering, reviewing, and exporting reports. All processing is done safely on your own computer.
            </p>
          </div>
        )}

      </div>
    </div>
  );
}
