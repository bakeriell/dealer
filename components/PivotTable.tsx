import React, { useState } from 'react';
import { DealerData, ColumnDefinition } from '../types';
import { ChevronDown, ChevronUp, Search, Filter } from 'lucide-react';

interface PivotTableProps {
  data: DealerData[];
  columns: ColumnDefinition[];
}

export const PivotTable: React.FC<PivotTableProps> = ({ data, columns }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortCol, setSortCol] = useState<string>('Dealer Name');
  const [sortAsc, setSortAsc] = useState(true);

  // Filter columns to prioritize Dealer Name as the first fixed column
  const mainColKey = 'Dealer Name';
  const otherCols = columns.filter(c => c.key !== mainColKey && c.key !== 'id');

  // Helper to safely render cell values (Fixes React Error #31 for Date objects)
  const renderCell = (value: any) => {
    if (value === null || value === undefined) {
      return '';
    }
    if (value instanceof Date) {
      return value.toLocaleDateString();
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  };

  const filteredData = data.filter(row => 
    Object.values(row).some(val => 
      renderCell(val).toLowerCase().includes(searchTerm.toLowerCase())
    )
  );

  const sortedData = [...filteredData].sort((a, b) => {
    const valA = a[sortCol];
    const valB = b[sortCol];
    
    if (valA === valB) return 0;
    if (valA === undefined) return 1;
    if (valB === undefined) return -1;

    // Handle string/number comparison
    if (valA < valB) return sortAsc ? -1 : 1;
    if (valA > valB) return sortAsc ? 1 : -1;
    return 0;
  });

  const handleHeaderClick = (key: string) => {
    if (sortCol === key) {
      setSortAsc(!sortAsc);
    } else {
      setSortCol(key);
      setSortAsc(true);
    }
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-full">
      <div className="p-4 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-center gap-4 bg-slate-50/50">
        <h3 className="font-semibold text-slate-800 flex items-center gap-2">
          <Filter size={18} className="text-slate-500" />
          Dealer Data Grid
        </h3>
        <div className="relative w-full sm:w-72">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input 
            type="text" 
            placeholder="Search dealers, cities, data..." 
            className="w-full pl-10 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      <div className="overflow-auto custom-scrollbar flex-1">
        <table className="w-full text-left text-sm border-collapse">
          <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
            <tr>
              <th 
                className="p-4 font-semibold text-slate-700 border-b border-slate-200 min-w-[200px] cursor-pointer hover:bg-slate-100 transition-colors sticky left-0 bg-slate-50 z-20 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]"
                onClick={() => handleHeaderClick(mainColKey)}
              >
                <div className="flex items-center gap-1">
                  {mainColKey}
                  {sortCol === mainColKey && (
                    sortAsc ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                  )}
                </div>
              </th>
              {otherCols.map(col => (
                <th 
                  key={col.key}
                  className="p-4 font-semibold text-slate-700 border-b border-slate-200 whitespace-nowrap cursor-pointer hover:bg-slate-100 transition-colors"
                  onClick={() => handleHeaderClick(col.key)}
                >
                  <div className="flex items-center gap-1">
                    {col.label}
                    {sortCol === col.key && (
                      sortAsc ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                    )}
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {sortedData.length > 0 ? (
              sortedData.map((row) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors group">
                  <td className="p-4 font-medium text-slate-900 border-r border-slate-100 bg-white group-hover:bg-slate-50 sticky left-0 z-10 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                    {renderCell(row[mainColKey])}
                  </td>
                  {otherCols.map(col => (
                    <td key={col.key} className="p-4 text-slate-600 whitespace-nowrap">
                      {renderCell(row[col.key])}
                    </td>
                  ))}
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={columns.length} className="p-8 text-center text-slate-500">
                  No dealers found matching your search.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div className="p-3 border-t border-slate-200 bg-slate-50 text-xs text-slate-500 text-right">
        Showing {sortedData.length} records
      </div>
    </div>
  );
};