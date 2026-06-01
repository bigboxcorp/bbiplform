import React, { useState, useEffect } from 'react';
import { Database, AlertTriangle, RefreshCw } from 'lucide-react';

export default function ResponsesViewer({ formId }: { formId: string }) {
  const [data, setData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [formTitle, setFormTitle] = useState('Collected Responses');
  const [columns, setColumns] = useState<string[]>([]);
  const [themeColor, setThemeColor] = useState<string>('#2563eb');

  useEffect(() => {
    fetchResponses();
  }, [formId]);

  const fetchResponses = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/forms/${formId}/responses`);
      const payload = await res.json();
      
      if (!res.ok) throw new Error(payload.error || 'Failed to load responses');
      
      if (payload.columns) setColumns(payload.columns);
      if (payload.data) setData(payload.data);
      if (payload.title) setFormTitle(payload.title);
      if (payload.themeColor) setThemeColor(payload.themeColor);

    } catch(err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const formatCell = (val: any, colName: string) => {
      if (typeof val === 'number') {
          const lCol = (colName || '').toLowerCase();
          if (lCol.includes('date') || lCol.includes('time') || lCol.includes('at') || lCol.includes('submit')) {
              // Convert Excel serial date to JS Date
              // Excel epoch is Jan 1 1900. 1 = 1/1/1900
              const jsDate = new Date((val - 25569) * 86400 * 1000);
              if (!isNaN(jsDate.getTime())) {
                  return jsDate.toLocaleString();
              }
          }
      }
      if (typeof val === 'string') {
          // Check if it looks like a URL or comma separated URLs
          if (val.startsWith('http')) {
             const parts = val.split(',').map(s => s.trim());
             if (parts.length > 1) {
               return (
                   <div className="flex flex-col gap-1 w-full max-w-[200px]">
                       {parts.map((p, ix) => <a key={ix} href={p} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline truncate">Attachment {ix+1}</a>)}
                   </div>
               );
             } else {
               return <a href={val} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline truncate max-w-[200px] inline-block">Link</a>;
             }
          }
          if (val.startsWith('=HYPERLINK')) {
             const match = val.match(/"(http.*?)"/);
             if (match) {
                return <a href={match[1]} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline truncate max-w-[200px] inline-block">Link</a>;
             }
          }
      }
      return String(val ?? '');
  };

  if (loading) {
    return <div className="min-h-screen flex items-center justify-center text-slate-500 animate-pulse">Loading responses...</div>;
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4">
         <div className="bg-red-50 text-red-600 p-6 rounded-xl border border-red-200">
             <AlertTriangle className="mb-2" />
             <p className="font-bold">Error loading responses</p>
             <p className="text-sm">{error}</p>
         </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 p-6 sm:p-10">
       <div className="max-w-6xl mx-auto space-y-6">
          <div className="flex justify-between items-end border-b pb-4" style={{ borderColor: themeColor }}>
             <div>
                <h1 className="text-2xl font-bold text-slate-800">{formTitle}</h1>
                <p className="text-sm text-slate-500">Live response viewer</p>
             </div>
             <button onClick={fetchResponses} className="text-xs font-bold text-slate-600 bg-white border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-100 flex items-center gap-1.5">
                <RefreshCw size={14} /> Refresh
             </button>
          </div>

          <div className="bg-white border text-sm border-slate-200 shadow-sm rounded-xl overflow-hidden w-full overflow-x-auto">
             <table className="w-full text-left whitespace-nowrap">
                <thead>
                   <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                      {columns.map((c, i) => <th key={i} className="px-4 py-3 font-semibold">{c}</th>)}
                   </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                   {data.length === 0 ? (
                      <tr><td colSpan={columns.length || 1} className="p-8 text-center text-slate-400">No responses yet.</td></tr>
                   ) : data.map((row, i) => (
                      <tr key={i} className="hover:bg-slate-50">
                          {row.map((val: any, j: number) => (
                              <td key={j} className="px-4 py-3">
                                 {formatCell(val, columns[j])}
                              </td>
                          ))}
                      </tr>
                   ))}
                </tbody>
             </table>
          </div>
       </div>
    </div>
  );
}
