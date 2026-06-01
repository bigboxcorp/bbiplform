import React, { useEffect, useState } from 'react';
import { ExcelSaveConfig, MSTokens } from '../types';
import { fetchGraph } from '../utils/graphHelper';
import { Activity, RefreshCw } from 'lucide-react';

interface EditLogsProps {
  saveConfig: ExcelSaveConfig | null;
  tokens: MSTokens | null;
}

export default function EditLogs({ saveConfig, tokens }: EditLogsProps) {
  const [logs, setLogs] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  const loadLogs = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/logs');
      const data = await res.json();
      setLogs(data);
    } catch(e) {}
    setLoading(false);
  };

  useEffect(() => {
    loadLogs();
  }, []);

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 max-h-[80vh] flex flex-col">
      <div className="flex items-center justify-between border-b border-slate-200 pb-4 mb-4">
        <h2 className="font-semibold text-slate-800 text-base flex items-center gap-2">
           <Activity size={18} className="text-blue-500" /> System Action Logs
        </h2>
        <button onClick={loadLogs} className="text-xs flex items-center gap-1 text-slate-500 hover:text-blue-600 transition-colors">
          <RefreshCw size={12} className={loading ? 'animate-spin' : ''} /> Refresh
        </button>
      </div>
      
      <div className="flex-1 overflow-y-auto pr-2 space-y-3">
         {loading && logs.length === 0 && <div className="text-xs text-slate-400">Loading logs...</div>}
         {!loading && logs.length === 0 && <div className="text-xs text-slate-400">No logs found. Modifying the Excel table via dashboard will record logs here.</div>}
         
         {logs.map(log => (
            <div key={log.id} className="border-l-2 border-blue-500 pl-3 py-1 bg-slate-50 rounded-r-lg p-2 text-xs">
              <div className="flex items-center justify-between text-[10px] text-slate-500 font-mono mb-1">
                 <span>{new Date(log.createdAt).toLocaleString()}</span>
                 <span className="bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded font-bold uppercase">{log.action}</span>
              </div>
              <div className="text-slate-700 font-medium">{log.details}</div>
            </div>
         ))}
      </div>
    </div>
  );
}
