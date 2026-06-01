import React, { useState } from 'react';
import { FormConfig, ExcelSaveConfig, MSTokens, FormSubmission } from '../types';
import { fetchGraph, getTableColumns, addRowToTable, getTeamDrive, fetchPublicGraph, uploadFilePublic, uploadFileSystem, getProfile } from '../utils/graphHelper';
import { 
  Send, 
  ArrowRight, 
  CheckCircle, 
  RefreshCw, 
  Download, 
  Database,
  CloudLightning,
  Monitor,
  AlertTriangle,
  Clock,
  Trash2
} from 'lucide-react';

interface FormSubmissionFormProps {
  formConfig: FormConfig;
  tokens?: MSTokens | null;
  setTokens?: (t: MSTokens) => void;
  saveConfig?: ExcelSaveConfig | null;
  submissions?: FormSubmission[];
  setSubmissions?: React.Dispatch<React.SetStateAction<FormSubmission[]>>;
  publicMode?: boolean;
  formId?: string;
}

export default function FormSubmissionForm({
  formConfig,
  tokens,
  setTokens,
  saveConfig,
  submissions = [],
  setSubmissions,
  publicMode = false,
  formId
}: FormSubmissionFormProps) {
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [validationErrors, setValidationErrors] = useState<Record<string, string>>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [errorMessage, setErrorMessage] = useState('');

  // Handle value triggers
  const handleInputChange = (fieldId: string, val: any) => {
    setFormData(prev => ({
      ...prev,
      [fieldId]: val
    }));
    // Remove validation warning once user type some input
    if (validationErrors[fieldId]) {
      setValidationErrors(prev => {
        const copy = { ...prev };
        delete copy[fieldId];
        return copy;
      });
    }
  };

  const handleCheckboxChange = (fieldId: string, option: string, checked: boolean) => {
    setFormData(prev => {
      const current = prev[fieldId] || [];
      const newArray = checked 
        ? [...current, option]
        : current.filter((v: string) => v !== option);
      return { ...prev, [fieldId]: newArray };
    });
    if (validationErrors[fieldId]) {
      setValidationErrors(prev => {
        const copy = { ...prev };
        delete copy[fieldId];
        return copy;
      });
    }
  };

  // Submit trigger
  const dataFields = formConfig.fields.filter(f => f.type !== 'section_break');

  const pages: any[][] = [];
  let curr: any[] = [];
  formConfig.fields.forEach(f => {
     if (f.type === 'section_break') {
         if (curr.length > 0) pages.push(curr);
         curr = [f];
     } else {
         curr.push(f);
     }
  });
  if (curr.length > 0) pages.push(curr);

  const [currentPage, setCurrentPage] = useState(0);

  const fieldsToRender = pages.length > 0 ? pages[currentPage] : [];

  const handleNextOrSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (publicMode && !saveConfig) {
      setSubmitStatus('error');
      setErrorMessage('Form is currently locked by the administrator and is not accepting responses.');
      return;
    }

    // Client-Side Validation for Current Page ONLY
    const errors: Record<string, string> = {};
    fieldsToRender.forEach(field => {
      if (field.type === 'section_break') return;
      const val = formData[field.id];
      const isEmpty = val === undefined || val === null || (typeof val === 'string' && val.trim() === '') || (Array.isArray(val) && val.length === 0);
      
      if (field.required && isEmpty) {
        errors[field.id] = `${field.label} is required.`;
      } else if (!isEmpty) {
        if (field.type === 'short_text' || field.type === 'long_text') {
          if (field.minLength && String(val).length < field.minLength) errors[field.id] = `Minimum ${field.minLength} characters required.`;
          if (field.maxLength && String(val).length > field.maxLength) errors[field.id] = `Maximum ${field.maxLength} characters allowed.`;
          if (field.pattern && !new RegExp(field.pattern).test(String(val))) errors[field.id] = `Invalid format.`;
        }
        if (field.type === 'number') {
           const numVal = Number(val);
           if (!isNaN(numVal)) {
             if (field.minValue !== undefined && numVal < Number(field.minValue)) errors[field.id] = `Value must be at least ${field.minValue}.`;
             if (field.maxValue !== undefined && numVal > Number(field.maxValue)) errors[field.id] = `Value cannot exceed ${field.maxValue}.`;
           }
        }
        if (field.type === 'grid_radio' && field.required && field.gridRows?.length) {
            const keys = Object.keys(val || {});
            if (keys.length < field.gridRows.length) {
                errors[field.id] = `Please answer all rows.`;
            }
        }
      }
    });

    if (Object.keys(errors).length > 0) {
      setValidationErrors(errors);
      return;
    }

    setValidationErrors({});

    if (currentPage < pages.length - 1) {
       setCurrentPage(currentPage + 1);
       window.scrollTo({ top: 0, behavior: 'smooth' });
       return;
    }

    let hasData = false;
    for (const key of Object.keys(formData)) {
      const val = formData[key];
      if (val !== undefined && val !== null && val !== '') {
         if (Array.isArray(val) && val.length === 0) continue;
         if (typeof val === 'object' && Object.keys(val).length === 0 && !(val instanceof File)) continue;
         hasData = true;
         break;
      }
    }
    if (!hasData) {
      setErrorMessage('You cannot submit an empty form. Please answer at least one question.');
      return;
    }

    setIsSubmitting(true);
    setErrorMessage('');
    
    let submissionId = `SUB-${Math.random().toString(36).substring(2, 8).toUpperCase()}`;
    try {
       const res = await fetch(`/api/forms/${formId}/generate-id`, { method: 'POST' });
       if (res.ok) {
           const data = await res.json();
           if (data.nextId) submissionId = data.nextId;
       }
    } catch(err) {
       console.warn('Could not fetch sequential ID from server', err);
    }

    const submittedAt = new Date().toISOString();

    const newSubmission: FormSubmission = {
      id: submissionId,
      submittedAt: submittedAt,
      data: { ...formData },
      status: 'pending'
    };

    try {
      const finalFormData = { ...formData };
      
      // Upload files immediately if we have access to graph
      for (const field of formConfig.fields) {
         if (field.type === 'file' && formData[field.id] && Array.isArray(formData[field.id])) {
             const files = formData[field.id] as File[];
             const uploadedLinks: string[] = [];
             for (let i = 0; i < files.length; i++) {
                 const file = files[i];
                 try {
                     let uploadedResult: any = null;
                     if (saveConfig) {
                       let driveId = saveConfig.driveId;
                       if (!driveId && saveConfig.driveItemId.includes('!')) {
                         driveId = saveConfig.driveItemId.split('!')[0];
                       }
                       if (!driveId && saveConfig.groupId && tokens && setTokens) {
                          const driveInfo = await getTeamDrive(saveConfig.groupId, tokens, setTokens);
                          driveId = driveInfo.id;
                       }
                       if (driveId) {
                         if (publicMode && formId) {
                           uploadedResult = await uploadFilePublic(
                              driveId,
                              saveConfig.channelName || 'General',
                              `${submissionId}_${file.name}`,
                              file,
                              formId
                           );
                         } else if (tokens && setTokens) {
                           uploadedResult = await uploadFileSystem(
                              driveId,
                              saveConfig.channelName || 'General',
                              `${submissionId}_${file.name}`,
                              file,
                              tokens,
                              setTokens
                           );
                         }
                       }
                     }
                     if (uploadedResult && uploadedResult.webUrl) {
                         uploadedLinks.push(uploadedResult.webUrl);
                     } else {
                         uploadedLinks.push(`[Local Placeholder: ${file.name}]`);
                     }
                 } catch (err: any) {
                     console.error('File upload failed:', err);
                     throw new Error(`Failed to upload file ${file.name}: ${err.message}`);
                 }
             }
             finalFormData[field.id] = uploadedLinks.join(', ');
         } else if (field.type === 'grid_radio' || field.type === 'grid_checkbox') {
             const valObj = formData[field.id];
             if (valObj && typeof valObj === 'object') {
                 finalFormData[field.id] = Object.entries(valObj).map(([row, checked]) => `${row}: ${Array.isArray(checked) ? checked.join('|') : checked}`).join(' ; ');
             }
         }
      }

      if (publicMode && formId && saveConfig) {
        let respondentEmail = (window as any).respondentEmail || '';
        // PUBLIC SUBMISSION VIA GRAPH PROXY
        let driveId = saveConfig.driveId;
        if (!driveId && saveConfig.driveItemId.includes('!')) {
          driveId = saveConfig.driveItemId.split('!')[0];
        }
        const fileItemId = saveConfig.driveItemId;
        const tableName = saveConfig.tableName;

        // Fetch Columns anonymously
        const endpointGetCols = `drives/${driveId}/items/${fileItemId}/workbook/tables/${tableName}/columns`;
        let tableColumns: any[] = [];
        try {
          const res = await fetchPublicGraph(endpointGetCols, formId);
          tableColumns = res.value || [];
        } catch (colErr) {
          console.warn('Could not read columns, falling back to mapping sequence');
        }

        let rowData: any[] = [];
        const currentTimestampFormatted = new Date().toLocaleString('en-US', { timeZone: 'UTC' }) + ' UTC';

        if (tableColumns.length > 0) {
          rowData = tableColumns.map((col: any) => {
            const colNameClean = (col.name || '').trim().toLowerCase();

            const mappedFieldEntry = Object.entries(saveConfig.columnsMapping).find(
              ([fId, excelColName]) => excelColName.trim().toLowerCase() === colNameClean
            );

            if (mappedFieldEntry) {
              const fId = mappedFieldEntry[0];
              if (fId === '__submission_id') return submissionId;
              if (fId === '__submitted_at') return currentTimestampFormatted;
              if (fId === 'respondent_email') return respondentEmail || 'Anonymous';
              const userVal = finalFormData[fId];
              return userVal !== undefined ? (Array.isArray(userVal) ? userVal.join(', ') : userVal) : '';
            }

            if (colNameClean.includes('submission id') || colNameClean === 'id') return submissionId;
            if (colNameClean.includes('submitted at')) return currentTimestampFormatted;
            if (respondentEmail && (colNameClean.includes('email') || colNameClean.includes('submitted by') || colNameClean.includes('respondent'))) return respondentEmail;

            const matchedField = formConfig.fields.find(f => f.label.trim().toLowerCase() === colNameClean);
            if (matchedField) {
              const val = finalFormData[matchedField.id];
              return val !== undefined ? (Array.isArray(val) ? val.join(', ') : val) : '';
            }
            return ''; 
          });
        } else {
          rowData = [
            submissionId,
            currentTimestampFormatted,
            ...dataFields.map(f => {
              const val = finalFormData[f.id];
              return val !== undefined ? (Array.isArray(val) ? val.join(', ') : val) : '';
            })
          ];
        }

        const endpointAddRow = `drives/${driveId}/items/${fileItemId}/workbook/tables/${tableName}/rows`;
        await fetchPublicGraph(endpointAddRow, formId, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ values: [rowData] })
        });
        
        newSubmission.status = 'success';
      } else if (tokens && saveConfig && setTokens) {
        // AUTHENTICATED SANDBOX SUBMISSION
        let driveId = saveConfig.driveId;
        if (!driveId) {
          if (saveConfig.driveItemId.includes('!')) {
            driveId = saveConfig.driveItemId.split('!')[0];
          } else if (saveConfig.groupId) {
             const driveInfo = await getTeamDrive(saveConfig.groupId, tokens, setTokens);
             driveId = driveInfo.id;
          } else {
             throw new Error('Missing drive ID and could not resolve it. Please reconnect your spreadsheet in the connector tab.');
          }
        }
        
        const fileItemId = saveConfig.driveItemId;
        const tableName = saveConfig.tableName;

        let tableColumns: any[] = [];
        try {
          tableColumns = await getTableColumns(driveId, fileItemId, tableName, tokens, setTokens);
        } catch (colErr: any) {
          console.warn('Could not read columns layout. Falling back to default order.', colErr);
        }

        let rowData: any[] = [];
        const currentTimestampFormatted = new Date().toLocaleString('en-US', { timeZone: 'UTC' }) + ' UTC';

        if (tableColumns.length > 0) {
          rowData = tableColumns.map((col: any) => {
            const colNameClean = (col.name || '').trim().toLowerCase();

            const mappedFieldEntry = Object.entries(saveConfig.columnsMapping).find(
              ([fId, excelColName]) => excelColName.trim().toLowerCase() === colNameClean
            );

            if (mappedFieldEntry) {
              const fId = mappedFieldEntry[0];
              if (fId === '__submission_id') return submissionId;
              if (fId === '__submitted_at') return currentTimestampFormatted;
              if (fId === 'respondent_email') return typeof (window as any).respondentEmail !== 'undefined' ? ((window as any).respondentEmail || 'Anonymous') : 'Anonymous';
              const userVal = finalFormData[fId];
              return userVal !== undefined ? (Array.isArray(userVal) ? userVal.join(', ') : userVal) : '';
            }

            if (colNameClean.includes('submission id') || colNameClean === 'id') return submissionId;
            if (colNameClean.includes('submitted at')) return currentTimestampFormatted;
            
            if (colNameClean.includes('email') || colNameClean.includes('submitted by') || colNameClean.includes('respondent')) {
               return typeof (window as any).respondentEmail !== 'undefined' ? (window as any).respondentEmail : 'Admin / Owner';
            }

            const matchedField = formConfig.fields.find(f => f.label.trim().toLowerCase() === colNameClean);
            if (matchedField) {
              const val = finalFormData[matchedField.id];
              return val !== undefined ? (Array.isArray(val) ? val.join(', ') : val) : '';
            }
            return ''; 
          });
        } else {
          rowData = [
            submissionId,
            currentTimestampFormatted,
            ...dataFields.map(f => {
              const val = finalFormData[f.id];
              return val !== undefined ? (Array.isArray(val) ? val.join(', ') : val) : '';
            })
          ];
        }

        await addRowToTable(driveId, fileItemId, tableName, [rowData], tokens, setTokens);
        newSubmission.status = 'success';
      } else {
        // LOCAL OFFLINE MODE
        await new Promise(resolve => setTimeout(resolve, 800));
        newSubmission.status = 'success';
      }

      if (setSubmissions) setSubmissions(prev => [newSubmission, ...prev]);
      setSubmitStatus('success');
      
      if (publicMode && formId && formConfig.settings?.allowMultipleSubmissions === false) {
          localStorage.setItem(`__form_submitted_${formId}`, 'true');
          const email = (window as any).respondentEmail;
          if (email) {
             try {
                await fetch(`/api/forms/${formId}/record-submission`, {
                   method: 'POST',
                   headers: {'Content-Type': 'application/json'},
                   body: JSON.stringify({ email })
                });
             } catch(e) {}
          }
      }

      setFormData({}); 
    } catch (err: any) {
      console.error('Submission failed:', err);
      newSubmission.status = 'failed';
      newSubmission.error = err.message;
      if (setSubmissions) setSubmissions(prev => [newSubmission, ...prev]);
      
      setSubmitStatus('error');
      
      let errMsg = err.message || 'Unknown error';
      if (errMsg.toLowerCase().includes('itemnotfound') || errMsg.toLowerCase().includes('not found')) {
         errMsg = 'The connected Excel table or file has been deleted, renamed, or moved. Please open the form builder and reconnect to a new table.';
      }
      
      setErrorMessage(`${errMsg}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleExportCSV = () => {
    if (submissions.length === 0) return;
    const headers = ['Submission ID', 'Submitted At', ...dataFields.map(f => f.label)];
    const rows = submissions.map(sub => {
      return [
        sub.id,
        new Date(sub.submittedAt).toLocaleString(),
        ...dataFields.map(field => {
          const val = sub.data[field.id];
          if (val === undefined || val === null) return '';
          const stringified = String(Array.isArray(val) ? val.join(', ') : val).replace(/"/g, '""');
          return stringified.includes(',') || stringified.includes('\n') || stringified.includes('"')
            ? `"${stringified}"`
            : stringified;
        })
      ];
    });

    const csvContent = [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
    const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const cleanTitle = formConfig.title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    link.setAttribute('download', `${cleanTitle}_submissions.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const clearHistoryLog = () => {
    if (window.confirm('Clear history log?')) {
      if (setSubmissions) setSubmissions([]);
    }
  };

  const themeStyle = formConfig.settings?.themeColor ? { borderColor: formConfig.settings.themeColor } : {};
  const themeBgStyle = formConfig.settings?.themeColor ? { backgroundColor: formConfig.settings.themeColor } : {};

  return (
    <div className="space-y-6" id="form-submission-container">
      {formConfig.settings?.logoUrl && (
        <div className={`flex mb-4 ${formConfig.settings.logoAlignment === 'left' ? 'justify-start' : formConfig.settings.logoAlignment === 'right' ? 'justify-end' : 'justify-center'}`}>
           <img src={formConfig.settings.logoUrl} alt="Logo" style={{ height: formConfig.settings.logoSize ? `${formConfig.settings.logoSize}px` : '64px' }} className="object-contain" />
        </div>
      )}
      
      <div className="bg-white rounded-xl shadow-sm overflow-hidden flex flex-col" style={{ borderTopWidth: '8px', ...themeStyle }}>
        {formConfig.settings?.coverUrl && (
           <img src={formConfig.settings.coverUrl} alt="Cover" className="w-full h-32 sm:h-48 object-cover" />
        )}
        <div className={`p-6 border-b border-slate-100 flex items-center justify-between transition-colors bg-white ${formConfig.settings?.headerAlignment === 'center' ? 'flex-col text-center' : formConfig.settings?.headerAlignment === 'right' ? 'flex-row-reverse text-right' : 'flex-row text-left'}`}>
          <div className={`${formConfig.settings?.headerAlignment === 'center' ? 'w-full mb-3' : ''}`}>
            <h2 className="font-bold text-slate-800 text-lg">{formConfig.title || 'Data Entry Form'}</h2>
            <p className="text-sm text-slate-500 mt-1">{formConfig.description || 'Fill out the form below.'}</p>
          </div>
          
          {!publicMode && (
            <div className="flex items-center gap-1.5 shrink-0">
              {tokens && saveConfig ? (
                <span className="text-[11px] font-bold text-emerald-600 bg-emerald-50 border border-emerald-200 px-3 py-1 rounded-full flex items-center gap-1">
                  <CloudLightning size={12} className="text-emerald-500 animate-pulse" />
                  M365 Excel Sync
                </span>
              ) : (
                <span className="text-[11px] font-bold text-slate-650 bg-slate-100 border border-slate-200 px-3 py-1 rounded-full flex items-center gap-1">
                  <Monitor size={12} className="text-slate-500" />
                  Offline Sandbox
                </span>
              )}
            </div>
          )}
        </div>

        {submitStatus === 'success' ? (
          <div className="p-8 text-center space-y-5 animate-scaleIn">
            <div className="inline-flex items-center justify-center w-14 h-14 rounded-full bg-emerald-50 text-emerald-600 border border-emerald-150">
              <CheckCircle size={28} />
            </div>
            
            <div className="max-w-md mx-auto">
              <h3 className="font-bold text-slate-800 text-lg leading-snug">{formConfig.settings?.allowMultipleSubmissions === false ? "Thank you!" : "Response Submitted Successfully"}</h3>
              {formConfig.settings?.allowMultipleSubmissions === false && (
                <p className="text-slate-500 text-sm mt-2">Your response has been recorded. You can now close this window.</p>
              )}
            </div>

            {formConfig.settings?.allowMultipleSubmissions !== false && (
              <button 
                onClick={() => setSubmitStatus('idle')}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-sm hover:shadow active:scale-95 transition-all cursor-pointer"
              >
                Submit Another Entry
              </button>
            )}
          </div>
        ) : (
          <form onSubmit={handleNextOrSubmit} className="p-8 space-y-6">
            {submitStatus === 'error' && (
              <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl flex items-start gap-3 text-rose-800 text-xs font-semibold leading-relaxed">
                <AlertTriangle size={15} className="text-rose-600 shrink-0 mt-0.5" />
                <span>{errorMessage}</span>
              </div>
            )}

            {formConfig.fields.length === 0 ? (
              <div className="text-center py-12 text-slate-400 font-semibold text-xs border border-dashed border-slate-200 rounded-xl bg-slate-50/55">
                No active form layout detected.
              </div>
            ) : (
              <div className="space-y-6">
                {fieldsToRender.map(field => {
                   if (field.type === 'section_break') {
                     return (
                        <div key={field.id} className="pt-2 pb-1 border-b border-slate-200">
                           <h2 className="text-xl font-bold text-slate-800">{field.label}</h2>
                           {field.placeholder && <p className="text-sm text-slate-500 mt-1">{field.placeholder}</p>}
                        </div>
                     );
                   }
                   return (
                  <div key={field.id} className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700 flex items-center gap-1 select-none">
                      {field.label}
                      {field.required && <span className="text-red-500">*</span>}
                    </label>

                    {/* SHORT TEXT */}
                    {field.type === 'short_text' && (
                      <input 
                        type="text"
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        placeholder={field.placeholder || `e.g. Enter ${field.label.toLowerCase()}...`}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* LONG TEXT */}
                    {field.type === 'long_text' && (
                      <textarea 
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        placeholder={field.placeholder || `Enter details here...`}
                        rows={4}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 resize-y max-h-[300px] ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* NUMERIC INPUTS */}
                    {field.type === 'number' && (
                      <input 
                        type="number"
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        placeholder={field.placeholder || `0`}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* DATE INPUTS */}
                    {field.type === 'date' && (
                      <input 
                        type="date"
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* TIME INPUTS */}
                    {field.type === 'time' && (
                      <input 
                        type="time"
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* RATING / LINEAR SCALE */}
                    {field.type === 'rating' && (
                      <div className="flex flex-wrap gap-2 mt-1">
                        {[1, 2, 3, 4, 5].map((val) => (
                           <label key={val} className="flex flex-col items-center gap-1 cursor-pointer">
                              <span className="text-[10px] text-slate-400 font-bold">{val}</span>
                              <input
                                 type="radio"
                                 name={field.id}
                                 value={val}
                                 checked={String(formData[field.id]) === String(val)}
                                 onChange={(e) => handleInputChange(field.id, e.target.value)}
                                 className="w-5 h-5 text-blue-600 border-slate-300 focus:ring-blue-500 form-radio cursor-pointer"
                              />
                           </label>
                        ))}
                      </div>
                    )}

                    {/* DROPDOWN */}
                    {field.type === 'select' && (
                      <select
                        value={formData[field.id] || ''}
                        onChange={(e) => handleInputChange(field.id, e.target.value)}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      >
                        <option value="">{field.placeholder || `Select ${field.label}...`}</option>
                        {(field.options || []).map((opt, oIdx) => (
                          <option key={oIdx} value={opt}>{opt}</option>
                        ))}
                      </select>
                    )}

                    {/* RADIO BUTTONS */}
                    {field.type === 'radio' && (
                      <div className="space-y-2 mt-1">
                        {(field.options || []).map((opt, oIdx) => (
                          <label key={oIdx} className="flex items-center gap-2 cursor-pointer group">
                            <input
                              type="radio"
                              name={field.id}
                              value={opt}
                              checked={formData[field.id] === opt}
                              onChange={(e) => handleInputChange(field.id, e.target.value)}
                              className="w-4 h-4 text-blue-600 border-slate-300 focus:ring-blue-500"
                            />
                            <span className="text-sm text-slate-700 font-medium group-hover:text-slate-900">{opt}</span>
                          </label>
                        ))}
                      </div>
                    )}

                    {/* CHECKBOXES */}
                    {field.type === 'checkbox' && (
                      <div className="space-y-2 mt-1">
                        {(field.options || []).map((opt, oIdx) => (
                          <label key={oIdx} className="flex items-center gap-2 cursor-pointer group">
                            <input
                              type="checkbox"
                              checked={(formData[field.id] || []).includes(opt)}
                              onChange={(e) => handleCheckboxChange(field.id, opt, e.target.checked)}
                              className="w-4 h-4 text-blue-600 border-slate-300 rounded focus:ring-blue-500"
                            />
                            <span className="text-sm text-slate-700 font-medium group-hover:text-slate-900">{opt}</span>
                          </label>
                        ))}
                      </div>
                    )}

                    {/* FILE UPLOAD */}
                    {field.type === 'file' && (
                      <input 
                        type="file"
                        multiple={field.fileOptions ? field.fileOptions.maxAllowed > 1 : false}
                        onChange={(e) => {
                          const files = Array.from(e.target.files || []);
                          if (field.fileOptions) {
                              const validFiles = files.filter((f: File) => f.size <= field.fileOptions!.maxSizeMB * 1024 * 1024 && (field.fileOptions!.allowedTypes && field.fileOptions!.allowedTypes.length ? field.fileOptions!.allowedTypes.some(t => f.type.includes(t) || f.name.endsWith(t)) : true));
                              if (validFiles.length !== files.length) {
                                  alert('Some files were rejected due to size or type restrictions.');
                              }
                              handleInputChange(field.id, validFiles.slice(0, field.fileOptions.maxAllowed));
                          } else {
                              handleInputChange(field.id, files);
                          }
                        }}
                        className={`w-full text-sm px-4 py-2.5 bg-white rounded-lg border focus:outline-none focus:ring-2 transition-all font-medium text-slate-800 ${
                          validationErrors[field.id] ? 'border-red-350 focus:ring-red-500 bg-red-50/20' : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                        }`}
                      />
                    )}

                    {/* GRID RADIO */}
                    {field.type === 'grid_radio' && (
                      <div className="overflow-x-auto mt-2 border border-slate-200 rounded-lg">
                        <table className="w-full text-sm text-left">
                          <thead className="bg-slate-50 text-slate-600 font-medium">
                            <tr>
                              <th className="p-3 border-b border-slate-200"></th>
                              {(field.gridCols || []).map((col, cIdx) => (
                                <th key={cIdx} className="p-3 border-b border-slate-200 text-center">{col}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {(field.gridRows || []).map((row, rIdx) => (
                              <tr key={rIdx} className="border-b border-slate-100 last:border-0 hover:bg-slate-50/50">
                                <td className="p-3 font-medium text-slate-700">{row}</td>
                                {(field.gridCols || []).map((col, cIdx) => (
                                  <td key={cIdx} className="p-3 text-center">
                                    <input 
                                      type="radio"
                                      name={`${field.id}_${rIdx}`}
                                      value={col}
                                      checked={formData[field.id] && formData[field.id][row] === col}
                                      onChange={(e) => {
                                         const cur = formData[field.id] || {};
                                         handleInputChange(field.id, { ...cur, [row]: col });
                                      }}
                                      className="w-4 h-4 text-blue-600"
                                    />
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* GRID CHECKBOX */}
                    {field.type === 'grid_checkbox' && (
                      <div className="overflow-x-auto mt-2 border border-slate-200 rounded-lg">
                        <table className="w-full text-sm text-left">
                          <thead className="bg-slate-50 text-slate-600 font-medium">
                            <tr>
                              <th className="p-3 border-b border-slate-200"></th>
                              {(field.gridCols || []).map((col, cIdx) => (
                                <th key={cIdx} className="p-3 border-b border-slate-200 text-center">{col}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {(field.gridRows || []).map((row, rIdx) => (
                              <tr key={rIdx} className="border-b border-slate-100 last:border-0 hover:bg-slate-50/50">
                                <td className="p-3 font-medium text-slate-700">{row}</td>
                                {(field.gridCols || []).map((col, cIdx) => (
                                  <td key={cIdx} className="p-3 text-center">
                                    <input 
                                      type="checkbox"
                                      checked={formData[field.id] && formData[field.id][row] && formData[field.id][row].includes(col)}
                                      onChange={(e) => {
                                         const cur = formData[field.id] || {};
                                         const rowArr = cur[row] || [];
                                         const newArr = e.target.checked ? [...rowArr, col] : rowArr.filter((c: string) => c !== col);
                                         handleInputChange(field.id, { ...cur, [row]: newArr });
                                      }}
                                      className="w-4 h-4 text-blue-600 rounded"
                                    />
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {validationErrors[field.id] && (
                      <span className="text-[11px] font-bold text-red-500 block animate-fadeIn pl-1 pt-1">{validationErrors[field.id]}</span>
                    )}
                  </div>
                  );
                })}
              </div>
            )}

            {formConfig.fields.length > 0 && (
              <div className="pt-6 flex gap-3">
                {currentPage > 0 && (
                   <button
                     type="button"
                     onClick={() => {
                        setCurrentPage(p => p - 1);
                        window.scrollTo({ top: 0, behavior: 'smooth' });
                     }}
                     className="px-6 py-3 bg-white text-slate-700 font-bold border border-slate-300 rounded-lg shadow-sm hover:bg-slate-50 transition-all cursor-pointer"
                   >
                     Back
                   </button>
                )}
                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg shadow-md hover:shadow-lg transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50 disabled:pointer-events-none active:scale-[0.98]"
                  style={themeBgStyle}
                >
                  {isSubmitting ? (
                    <>
                      <RefreshCw size={16} className="animate-spin" />
                      Submitting...
                    </>
                  ) : currentPage < pages.length - 1 ? (
                    <>
                       Next Section
                    </>
                  ) : (
                    <>
                      <Send size={15} />
                      Submit Form
                    </>
                  )}
                </button>
              </div>
            )}
          </form>
        )}
      </div>

      {!publicMode && (
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mt-6">
          <div className="p-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/75">
            <div className="flex items-center gap-2">
              <span className="p-2 bg-blue-50 text-blue-650 rounded-lg border border-blue-100">
                <Database size={15} />
              </span>
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">Session Local Logs ({submissions.length})</span>
            </div>

            <div className="flex items-center gap-2">
              {submissions.length > 0 && (
                <>
                  <button 
                    onClick={clearHistoryLog}
                    className="p-1.5 px-3 text-[11px] text-slate-500 hover:text-red-650 font-bold flex items-center gap-1 border border-slate-200 bg-white rounded-lg active:scale-95 cursor-pointer transition-all hover:bg-slate-50 shadow-xs"
                  >
                    <Trash2 size={12} /> Clear Data
                  </button>
                  <button 
                    onClick={handleExportCSV}
                    className="p-1.5 px-3 bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-bold flex items-center gap-1.5 shadow-md rounded-lg active:scale-95 cursor-pointer transition-all"
                  >
                    <Download size={12} /> Download CSV
                  </button>
                </>
              )}
            </div>
          </div>

          {submissions.length === 0 ? (
            <div className="text-center py-12 text-slate-400 font-semibold text-xs leading-relaxed bg-white">
              No local sandbox entries have been recorded yet.
            </div>
          ) : (
            <div className="overflow-x-auto max-h-[300px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 border-b border-slate-200 text-[10px] font-bold uppercase select-none">
                    <th className="p-4 pl-5">ID</th>
                    <th className="p-4">Status</th>
                    <th className="p-4">Submitted At</th>
                    {dataFields.map(field => (
                      <th key={field.id} className="p-4 truncate max-w-[120px]">{field.label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {submissions.map(sub => (
                    <tr key={sub.id} className="hover:bg-slate-50/50 text-slate-700 text-xs">
                      <td className="p-4 pl-5 font-mono font-bold text-slate-900">{sub.id}</td>
                      <td className="p-4">
                        {sub.status === 'success' ? (
                          <span className="text-[10px] font-bold text-emerald-700 bg-emerald-55 px-2.5 py-1 rounded-full border border-emerald-200">
                            SUCCESS
                          </span>
                        ) : sub.status === 'failed' ? (
                          <span className="text-[10px] font-bold text-red-700 bg-red-50 px-2.5 py-1 rounded-full border border-red-200" title={sub.error}>
                            FAILED
                          </span>
                        ) : (
                          <span className="text-[10px] font-bold text-slate-500 bg-slate-100 px-2.5 py-1 rounded-full animate-pulse border border-slate-205">
                            PENDING
                          </span>
                        )}
                      </td>
                      <td className="p-4 font-mono text-slate-500 flex items-center gap-1">
                        <Clock size={12} className="text-slate-400" />
                        {new Date(sub.submittedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                      </td>
                      {dataFields.map(f => (
                        <td key={f.id} className="p-4 truncate max-w-[140px] font-semibold text-slate-800" title={String(sub.data[f.id] || '')}>
                          {sub.data[f.id] !== undefined ? String(sub.data[f.id]) : <span className="text-slate-300">-</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
