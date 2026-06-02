import React, { useEffect, useState } from 'react';
import { FileText, Plus, ExternalLink, Settings, Clock, CheckCircle } from 'lucide-react';
import { FormConfig, ExcelSaveConfig } from '../types';

interface SavedForm {
  id: string;
  config: FormConfig;
  excelConfig: ExcelSaveConfig | null;
  createdAt: string;
  updatedAt: string;
}

export default function FormsHome({ onEditForm, userEmail }: { onEditForm: (id: string | null, config: FormConfig, excelConfig: ExcelSaveConfig | null) => void, userEmail?: string | null }) {
  const [forms, setForms] = useState<SavedForm[]>([]);
  const [loading, setLoading] = useState(true);

  const fetchForms = () => {
    setLoading(true);
    const url = userEmail ? `/api/forms?email=${encodeURIComponent(userEmail)}` : '/api/forms';
    fetch(url)
      .then(res => res.json())
      .then(data => {
        setForms(data);
        setLoading(false);
      })
      .catch(err => {
        console.error(err);
        setLoading(false);
      });
  };

  const safeOrigin = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1') ? window.location.origin : window.location.origin.replace(/^http:\/\//i, 'https://');

  useEffect(() => {
    fetchForms();
  }, [userEmail]);

  const handleCreateNew = () => {
    onEditForm(null, {
      title: 'New Form',
      description: 'Fill details below',
      settings: { requireMicrosoftLogin: false },
      fields: [{ id: 'q1', label: 'Question 1', type: 'short_text', required: true }]
    }, null);
  };

  return (
    <div className="flex flex-col gap-6 max-w-6xl w-full mx-auto p-4 sm:p-6 lg:p-8">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center bg-white p-6 rounded-xl border border-slate-200 shadow-sm gap-4">
        <div>
          <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
            <FileText className="text-blue-600" /> My Forms
          </h2>
          <p className="text-sm text-slate-500 mt-1">Manage and create data collection forms</p>
        </div>
        
        <div className="flex flex-col sm:items-end gap-3">
          <button onClick={handleCreateNew} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg font-bold text-sm shadow-sm flex items-center justify-center gap-2 transition-all cursor-pointer">
            <Plus size={16} /> Create New Form
          </button>
        </div>
      </div>

      {loading ? (
        <div className="text-center p-8 text-slate-500 animate-pulse font-medium">Loading forms...</div>
      ) : forms.length === 0 ? (
        <div className="bg-slate-50 border border-dashed border-slate-300 rounded-xl p-12 text-center flex flex-col items-center">
          <div className="w-16 h-16 bg-white rounded-full flex items-center justify-center shadow-sm mb-4">
            <FileText className="text-slate-400 w-8 h-8" />
          </div>
          <h3 className="text-lg font-bold text-slate-700">No Forms Created Yet</h3>
          <p className="text-sm text-slate-500 mt-2 mb-6 max-w-sm">Create your first form to start collecting data directly into your Microsoft 365 tables.</p>
          <button onClick={handleCreateNew} className="bg-blue-100 text-blue-700 hover:bg-blue-200 px-6 py-2.5 rounded-lg font-bold text-sm flex items-center gap-2 transition-colors cursor-pointer border border-blue-200">
            <Plus size={16} /> Create First Form
          </button>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {forms.map(form => (
            <div key={form.id} className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col group hover:shadow-md hover:border-blue-200 transition-all cursor-pointer" onClick={() => onEditForm(form.id, form.config, form.excelConfig)}>
              <div className="p-5 flex-1">
                <div className="flex items-start justify-between mb-3">
                  <h3 className="font-bold text-slate-800 text-lg truncate pr-4">{form.config.title || 'Untitled Form'}</h3>
                  <div className={`p-1.5 rounded-md ${form.excelConfig ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'}`}>
                    {form.excelConfig ? <CheckCircle size={14} /> : <Settings size={14} />}
                  </div>
                </div>
                <p className="text-xs text-slate-500 line-clamp-2">{form.config.description || 'No description'}</p>
                
                <div className="mt-4 flex items-center gap-1.5 text-[10px] text-slate-400 font-mono">
                  <Clock size={12} /> {new Date(form.createdAt).toLocaleDateString()} at {new Date(form.createdAt).toLocaleTimeString()}
                </div>
              </div>
              
              <div className="border-t border-slate-100 p-3 bg-slate-50 flex items-center justify-between group-hover:bg-blue-50/50 transition-colors">
                 <span className="text-xs font-bold text-slate-600 group-hover:text-blue-700">Edit Form</span>
                 <a 
                   href={`${safeOrigin}/form/${form.id}`} 
                   target="_blank" 
                   rel="noopener noreferrer" 
                   onClick={e => e.stopPropagation()} 
                   className="p-1.5 text-slate-400 hover:text-blue-600 bg-white rounded shadow-xs hover:shadow-sm"
                   title="Open Live Form"
                 >
                   <ExternalLink size={14} />
                 </a>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
