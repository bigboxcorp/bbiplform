import React from 'react';
import { FormConfig, FormField } from '../types';
import { Plus, Trash, Type, AlignLeft, Calendar, Hash, GripVertical, Check } from 'lucide-react';

interface FormBuilderProps {
  config: FormConfig;
  onChange: (newConfig: FormConfig) => void;
}

export default function FormBuilder({ config, onChange }: FormBuilderProps) {
  const [newFieldLabel, setNewFieldLabel] = React.useState('');
  const [newFieldType, setNewFieldType] = React.useState<FormField['type']>('short_text');
  const [newFieldOptionsStr, setNewFieldOptionsStr] = React.useState('');
  const [newFieldGridRowsStr, setNewFieldGridRowsStr] = React.useState('');
  const [newFieldGridColsStr, setNewFieldGridColsStr] = React.useState('');

  // Fast auto-id from label
  const createSlug = (label: string) => {
    return label
      .toLowerCase()
      .trim()
      .replace(/[^\w\s-]/g, '')
      .replace(/[\s_-]+/g, '_')
      .replace(/^-+|-+$/g, '');
  };

  const addField = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFieldLabel.trim()) return;

    const baseId = createSlug(newFieldLabel) || `field_${Date.now()}`;
    // Deduplicate fields id
    let uniqueId = baseId;
    let counter = 1;
    while (config.fields.some(f => f.id === uniqueId)) {
      uniqueId = `${baseId}_${counter}`;
      counter++;
    }

    const needsOptions = ['select', 'radio', 'checkbox'].includes(newFieldType);
    const needsGrid = ['grid_radio', 'grid_checkbox'].includes(newFieldType);

    const parsedOptions = needsOptions && newFieldOptionsStr.trim() 
      ? newFieldOptionsStr.split(',').map(s => s.trim()).filter(Boolean) 
      : (needsOptions ? ['Option 1', 'Option 2'] : undefined);

    const parsedGridRows = needsGrid && newFieldGridRowsStr.trim()
      ? newFieldGridRowsStr.split(',').map(s => s.trim()).filter(Boolean)
      : (needsGrid ? ['Row 1', 'Row 2'] : undefined);

    const parsedGridCols = needsGrid && newFieldGridColsStr.trim()
      ? newFieldGridColsStr.split(',').map(s => s.trim()).filter(Boolean)
      : (needsGrid ? ['Col 1', 'Col 2'] : undefined);

    const newField: FormField = {
      id: uniqueId,
      label: newFieldLabel.trim(),
      type: newFieldType,
      required: false,
      placeholder: `Fill ${newFieldLabel.trim()}...`,
      options: parsedOptions,
      gridRows: parsedGridRows,
      gridCols: parsedGridCols
    };

    onChange({
      ...config,
      fields: [...config.fields, newField],
      settings: { ...config.settings, isMappingLocked: false }
    });

    setNewFieldLabel('');
    setNewFieldOptionsStr('');
    setNewFieldGridRowsStr('');
    setNewFieldGridColsStr('');
    setNewFieldType('short_text');
  };

  const removeField = (id: string) => {
    onChange({
      ...config,
      fields: config.fields.filter(f => f.id !== id),
      settings: { ...config.settings, isMappingLocked: false }
    });
  };

  const updateField = (id: string, updates: Partial<FormField>) => {
    onChange({
      ...config,
      fields: config.fields.map(f => {
        if (f.id === id) {
          const updated = { ...f, ...updates };
          // If toggled to select, fill placeholder options
          if (['select', 'radio', 'checkbox'].includes(updated.type) && !updated.options) {
            updated.options = ['Option 1', 'Option 2'];
          } else if (updated.type && !['select', 'radio', 'checkbox'].includes(updated.type)) {
            delete updated.options;
          }
          return updated;
        }
        return f;
      }),
      settings: { ...config.settings, isMappingLocked: false }
    });
  };

  // Quick preset fields
  // Removed as requested

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6" id="form-builder-component">
      {/* Header and presets */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-155 pb-5">
        <div>
          <h2 className="font-semibold text-slate-800 text-base">1. Define Form Structure</h2>
          <p className="text-xs text-slate-500 mt-0.5">Customize fields that map to your Excel columns</p>
        </div>
      </div>

      {/* Form Details inputs */}
      <div className="grid grid-cols-1 gap-4 bg-slate-50/50 p-5 rounded-xl border border-slate-200 leading-normal">
        <div className="space-y-1.5">
          <label className="text-xs font-semibold text-slate-700">Form Title</label>
          <input 
            type="text" 
            value={config.title}
            onChange={(e) => onChange({ ...config, title: e.target.value })}
            className="w-full text-sm px-4 py-2 bg-white rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-semibold text-slate-850 transition-all shadow-xs"
            placeholder="Feedback Form"
            id="form-title-input"
          />
        </div>
        <div className="space-y-1.5">
          <label className="text-xs font-semibold text-slate-700">Form Subtitle / Intro</label>
          <input 
            type="text" 
            value={config.description}
            onChange={(e) => onChange({ ...config, description: e.target.value })}
            className="w-full text-sm px-4 py-2 bg-white rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-650 transition-all shadow-xs"
            placeholder="Enter form information..."
            id="form-description-input"
          />
        </div>
      </div>

      {/* Existing Fields list */}
      <div className="space-y-3">
        <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Dynamic Fields ({config.fields.length})</h3>
        {config.fields.length === 0 ? (
          <div className="text-center py-12 rounded-xl border border-dashed border-slate-205 bg-slate-50/50">
            <span className="text-xs text-slate-450 font-medium block">No custom user fields defined yet. Choose a preset template above or add custom fields below.</span>
          </div>
        ) : (
          <div className="space-y-3">
            {config.fields.map((field, idx) => (
              <div 
                key={field.id} 
                className={`flex flex-col gap-3 p-4 rounded-xl border transition-all ${field.type === 'section_break' ? 'bg-blue-50/40 border-blue-200 border-l-4 border-l-blue-500 my-4' : 'bg-white border-slate-250 hover:border-slate-300 hover:shadow-xs'}`}
                id={`field-row-${field.id}`}
              >
                <div className="flex items-center gap-2">
                  <span className={`p-2 rounded-lg ${field.type === 'section_break' ? 'bg-blue-100 text-blue-600' : 'bg-slate-100 text-slate-500'}`}>
                    {(field.type === 'short_text' || field.type === 'long_text') && <Type size={14} />}
                    {field.type === 'number' && <Hash size={14} />}
                    {field.type === 'date' && <Calendar size={14} />}
                    {(field.type === 'select' || field.type === 'radio' || field.type === 'checkbox') && <AlignLeft size={14} />}
                    {field.type === 'section_break' && <Hash size={14} />}
                  </span>
                  
                  <div className="flex-1 min-w-0 pr-2">
                    <input 
                      type="text" 
                      value={field.label}
                      onChange={(e) => updateField(field.id, { label: e.target.value })}
                      className={`text-sm font-semibold bg-transparent hover:bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 px-2 py-1 rounded w-full border border-transparent hover:border-slate-200 transition-all ${field.type === 'section_break' ? 'text-blue-800 text-base' : 'text-slate-750'}`}
                      title={field.type === 'section_break' ? "Section Title" : "Rename field"}
                      placeholder={field.type === 'section_break' ? "Section Title..." : undefined}
                    />
                    <div className="flex items-center gap-1.5 mt-1 pl-2">
                      <select
                         value={field.type}
                         onChange={(e) => updateField(field.id, { type: e.target.value as FormField['type'] })}
                         className="text-[10px] font-semibold uppercase tracking-wider bg-transparent cursor-pointer hover:bg-slate-100 rounded px-1 -ml-1 border-none outline-none appearance-none"
                         style={field.type === 'section_break' ? {color: '#2563eb'} : {color: '#94a3b8'}}
                      >
                         <option value="short_text">SHORT TEXT</option>
                         <option value="long_text">LONG TEXT</option>
                         <option value="number">NUMBER</option>
                         <option value="date">DATE</option>
                         <option value="time">TIME</option>
                         <option value="select">DROPDOWN</option>
                         <option value="radio">SINGLE CHOICE</option>
                         <option value="checkbox">MULTIPLE CHOICE</option>
                         <option value="grid_radio">RADIO GRID</option>
                         <option value="grid_checkbox">CHECKBOX GRID</option>
                         <option value="file">FILE UPLOAD</option>
                         <option value="rating">RATING</option>
                         <option value="section_break">SECTION BREAK</option>
                      </select>
                    </div>
                  </div>

                  <div className="flex items-center gap-3 shrink-0">
                    {field.type !== 'section_break' && (
                      <label className="flex items-center gap-1.5 text-xs text-slate-600 font-semibold cursor-pointer">
                        <input 
                          type="checkbox" 
                          checked={field.required}
                          onChange={(e) => updateField(field.id, { required: e.target.checked })}
                          className="rounded text-blue-600 focus:ring-blue-500 border-slate-300 size-3.5 cursor-pointer"
                        />
                        Required
                      </label>
                    )}
                    <button 
                      onClick={() => removeField(field.id)}
                      className="p-1.5 hover:bg-red-50 text-slate-400 hover:text-red-600 rounded-lg transition-colors cursor-pointer"
                      title="Remove Field"
                    >
                      <Trash size={14} />
                    </button>
                  </div>
                </div>

                <div className="pl-10 grid grid-cols-2 md:grid-cols-4 gap-2">
                   {field.type === 'section_break' && (
                      <div className="col-span-full flex flex-col gap-2 bg-blue-50/50 p-3 rounded-lg border border-blue-100 mt-1">
                         <label className="text-[10px] uppercase font-bold text-blue-800">After this section</label>
                         <div className="flex items-center gap-2">
                           <select 
                              value={field.sectionEndAction || 'next'}
                              onChange={(e) => updateField(field.id, { sectionEndAction: e.target.value as any })}
                              className="border border-blue-200 rounded-md px-2 py-1.5 text-xs outline-none focus:border-blue-500 bg-white shadow-sm w-48 text-slate-700 cursor-pointer"
                           >
                              <option value="next">Continue to next section</option>
                              <option value="submit">Submit form</option>
                              <option value="goto_section">Go to section...</option>
                           </select>
                           {field.sectionEndAction === 'goto_section' && (
                              <select 
                                 value={field.sectionEndTarget || ''}
                                 onChange={(e) => updateField(field.id, { sectionEndTarget: e.target.value })}
                                 className="border border-blue-200 rounded-md px-2 py-1.5 text-xs outline-none focus:border-blue-500 bg-white shadow-sm w-48 text-slate-700 cursor-pointer"
                              >
                                 <option value="">-- Select section --</option>
                                 {config.fields.filter(f => f.type === 'section_break' && f.id !== field.id).map(sec => (
                                    <option key={sec.id} value={sec.id}>{sec.label}</option>
                                 ))}
                              </select>
                           )}
                         </div>
                      </div>
                   )}
                   {(field.type === 'short_text' || field.type === 'long_text') && (
                      <>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Min Length</label>
                            <input type="number" value={field.minLength || ''} onChange={(e) => updateField(field.id, { minLength: e.target.value ? parseInt(e.target.value) : undefined })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="e.g. 3" min="0" />
                         </div>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Max Length</label>
                            <input type="number" value={field.maxLength || ''} onChange={(e) => updateField(field.id, { maxLength: e.target.value ? parseInt(e.target.value) : undefined })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="e.g. 50" min="1" />
                         </div>
                         <div className="flex flex-col gap-1 col-span-2 md:col-span-2">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Regex Pattern (Optional)</label>
                            <input type="text" value={field.pattern || ''} onChange={(e) => updateField(field.id, { pattern: e.target.value || undefined })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="e.g. ^[a-z]+$" />
                         </div>
                      </>
                   )}
                   {field.type === 'number' && (
                      <>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Min Value</label>
                            <input type="number" value={field.minValue || ''} onChange={(e) => updateField(field.id, { minValue: e.target.value ? parseFloat(e.target.value) : undefined })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="e.g. 0" />
                         </div>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Max Value</label>
                            <input type="number" value={field.maxValue || ''} onChange={(e) => updateField(field.id, { maxValue: e.target.value ? parseFloat(e.target.value) : undefined })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="e.g. 100" />
                         </div>
                      </>
                   )}
                   {field.type === 'file' && (
                      <>
                         <div className="col-span-full">
                            <span className="text-[10px] bg-amber-50 text-amber-700 px-2 py-0.5 rounded font-bold border border-amber-200 inline-block mb-2">Note: To accept files, form must be connected to a Microsoft Team</span>
                         </div>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Max Allowed Files</label>
                            <input type="number" value={field.fileOptions?.maxAllowed || 1} onChange={(e) => updateField(field.id, { fileOptions: { ...(field.fileOptions || { maxSizeMB: 5 }), maxAllowed: parseInt(e.target.value || '1') } })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" min="1" max="10" />
                         </div>
                         <div className="flex flex-col gap-1">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Max Size (MB)</label>
                            <input type="number" value={field.fileOptions?.maxSizeMB || 5} onChange={(e) => updateField(field.id, { fileOptions: { ...(field.fileOptions || { maxAllowed: 1 }), maxSizeMB: parseInt(e.target.value || '5') } })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" min="1" max="100" />
                         </div>
                         <div className="flex flex-col gap-1 col-span-2">
                            <label className="text-[10px] text-slate-500 font-medium font-mono uppercase">Allowed Ext (e.g. pdf, png)</label>
                            <input type="text" value={field.fileOptions?.allowedTypes?.join(', ') || ''} onChange={(e) => updateField(field.id, { fileOptions: { ...(field.fileOptions || { maxSizeMB: 5, maxAllowed: 1 }), allowedTypes: e.target.value.split(',').map(s => s.trim()).filter(Boolean) } })} className="border border-slate-200 rounded px-2 py-1 text-xs outline-none focus:border-blue-400" placeholder="Leave empty for all" />
                         </div>
                      </>
                   )}
                </div>

                {/* If selected type is 'select', allow adding choice list items */}
                {['select', 'radio', 'checkbox'].includes(field.type) && (field.options || field.type === 'select' || field.type === 'radio' || field.type === 'checkbox') && (
                  <div className="text-xs pl-10 pr-2 mt-2 space-y-1 rounded-lg">
                    <div className="flex items-center justify-between mb-1.5">
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Choices:</span>
                    </div>
                    <div className="flex flex-col gap-2 relative">
                      {(field.options || []).map((option, oIdx) => {
                         const jump = field.logicJumps?.find(j => j.value === option);
                         return (
                        <div key={oIdx} className="flex flex-wrap items-center gap-1.5 bg-slate-50 border border-slate-200 pl-2.5 pr-1.5 py-1.5 rounded-lg font-semibold text-slate-700 text-[11px] shadow-xs">
                          <span>{option}</span>
                          <button 
                            onClick={() => {
                              const nextOptions = [...(field.options || [])];
                              nextOptions.splice(oIdx, 1);
                              updateField(field.id, { options: nextOptions });
                            }}
                            className="p-0.5 rounded-full text-slate-400 hover:text-red-500 hover:bg-white cursor-pointer mr-2"
                          >
                            ×
                          </button>
                          
                          {['select', 'radio'].includes(field.type) && (
                            <div className="flex items-center gap-1 ml-auto border-l border-slate-200 pl-3">
                               <span className="text-[9px] text-slate-400 font-bold uppercase">If selected:</span>
                               <select
                                  value={jump?.action || ''}
                                  onChange={(e) => {
                                     const newAction = e.target.value as any;
                                     let jumps = [...(field.logicJumps || [])];
                                     if (!newAction) {
                                         jumps = jumps.filter(j => j.value !== option);
                                     } else {
                                         const existing = jumps.findIndex(j => j.value === option);
                                         if (existing >= 0) {
                                            jumps[existing] = { ...jumps[existing], action: newAction };
                                         } else {
                                            jumps.push({ value: option, action: newAction });
                                         }
                                     }
                                     updateField(field.id, { logicJumps: jumps });
                                  }}
                                  className="border border-slate-200 rounded px-1.5 py-0.5 text-[10px] outline-none w-24 cursor-pointer focus:border-blue-400"
                               >
                                  <option value="">Continue</option>
                                  <option value="goto_section">Go to section</option>
                                  <option value="submit">Submit form</option>
                               </select>
                               {jump?.action === 'goto_section' && (
                                  <select
                                     value={jump.targetSectionId || ''}
                                     onChange={(e) => {
                                        let jumps = [...(field.logicJumps || [])];
                                        const existing = jumps.findIndex(j => j.value === option);
                                        if (existing >= 0) {
                                            jumps[existing] = { ...jumps[existing], targetSectionId: e.target.value };
                                        }
                                        updateField(field.id, { logicJumps: jumps });
                                     }}
                                     className="border border-slate-200 rounded px-1.5 py-0.5 text-[10px] outline-none w-28 cursor-pointer focus:border-blue-400"
                                  >
                                     <option value="">-- Choose --</option>
                                     {config.fields.filter(f => f.type === 'section_break' && f.id !== field.id).map(sec => (
                                        <option key={sec.id} value={sec.id}>{sec.label}</option>
                                     ))}
                                  </select>
                               )}
                            </div>
                          )}
                        </div>
                      )})}
                      <button
                        onClick={() => {
                          const val = prompt('Enter choice label:');
                          if (val && val.trim()) {
                            updateField(field.id, { options: [...(field.options || []), val.trim()] });
                          }
                        }}
                        className="text-[10px] px-3 py-1 border border-dashed border-blue-400 text-blue-600 hover:bg-blue-50 bg-white font-bold rounded-full cursor-pointer transition-all"
                      >
                        + Add Choice
                      </button>
                    </div>
                  </div>
                )}

                {/* Grid Rows Configuration */}
                {['grid_radio', 'grid_checkbox'].includes(field.type) && (
                  <div className="text-xs pl-10 pr-2 mt-2 space-y-3 rounded-lg">
                    {/* Rows */}
                    <div>
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1.5">Rows (Questions):</span>
                      <div className="flex flex-wrap gap-1.5">
                        {(field.gridRows || []).map((row, rIdx) => (
                          <div key={rIdx} className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 pl-2.5 pr-1.5 py-0.5 rounded-full font-semibold text-slate-700 text-[11px] shadow-xs">
                            <span>{row}</span>
                            <button onClick={() => updateField(field.id, { gridRows: (field.gridRows || []).filter((_, i) => i !== rIdx) })} className="p-0.5 rounded-full text-slate-400 hover:text-red-500">×</button>
                          </div>
                        ))}
                        <button onClick={() => { const v = prompt('Enter row text:'); if(v) updateField(field.id, { gridRows: [...(field.gridRows || []), v.trim()] }); }} className="text-[10px] px-3 py-1 border border-dashed border-slate-400 text-slate-600 hover:bg-slate-50 bg-white font-bold rounded-full cursor-pointer">+ Add Row</button>
                      </div>
                    </div>
                    {/* Cols */}
                    <div>
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1.5">Columns (Choices):</span>
                      <div className="flex flex-wrap gap-1.5">
                        {(field.gridCols || []).map((col, cIdx) => (
                          <div key={cIdx} className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 pl-2.5 pr-1.5 py-0.5 rounded-full font-semibold text-slate-700 text-[11px] shadow-xs">
                            <span>{col}</span>
                            <button onClick={() => updateField(field.id, { gridCols: (field.gridCols || []).filter((_, i) => i !== cIdx) })} className="p-0.5 rounded-full text-slate-400 hover:text-red-500">×</button>
                          </div>
                        ))}
                        <button onClick={() => { const v = prompt('Enter column label (e.g. Yes/No, 1-5):'); if(v) updateField(field.id, { gridCols: [...(field.gridCols || []), v.trim()] }); }} className="text-[10px] px-3 py-1 border border-dashed border-slate-400 text-slate-600 hover:bg-slate-50 bg-white font-bold rounded-full cursor-pointer">+ Add Column</button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Add Field interface */}
      <form onSubmit={addField} className="mt-8 bg-blue-50/40 border border-blue-200/60 rounded-xl p-5 shadow-sm" id="add-field-form">
        <div className="flex items-center gap-2 mb-4">
          <div className="bg-blue-100 text-blue-600 p-1.5 rounded-md">
            <Plus size={16} />
          </div>
          <span className="text-[13px] font-bold text-slate-700 uppercase tracking-wider block">Append New Column / Field</span>
        </div>
        <div className="flex flex-col gap-4">
           <div className="flex flex-col sm:flex-row gap-3">
             <div className="flex-1 space-y-1.5">
               <label className="text-[11px] font-semibold text-slate-500 uppercase">Field Label</label>
               <input 
                 type="text" 
                 value={newFieldLabel}
                 onChange={(e) => setNewFieldLabel(e.target.value)}
                 placeholder="e.g. Employee Name, Quantity"
                 className="w-full text-sm px-4 py-2 bg-white rounded-lg border border-slate-300 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all font-medium shadow-xs"
                 id="new-field-label-input"
               />
             </div>
             <div className="sm:w-64 space-y-1.5">
               <label className="text-[11px] font-semibold text-slate-500 uppercase">Input Type</label>
               <select
                 value={newFieldType}
                 onChange={(e) => setNewFieldType(e.target.value as FormField['type'])}
                 className="w-full text-sm px-4 py-2 bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 text-slate-700 font-semibold shadow-xs"
                 id="new-field-type-select"
               >
                 <option value="short_text">Short Text</option>
                 <option value="long_text">Paragraph / Long Text</option>
                 <option value="number">Numeric Input</option>
                 <option value="date">Date Picker</option>
                 <option value="time">Time Picker</option>
                 <option value="rating">Linear Scale (Rating)</option>
                 <option value="select">Dropdown Choice</option>
                 <option value="radio">Multiple Choice (Radio)</option>
                 <option value="checkbox">Checkboxes</option>
                 <option value="file">File Upload</option>
                 <option value="grid_radio">Grid (Radio Choice)</option>
                 <option value="grid_checkbox">Grid (Checkboxes)</option>
                 <option value="section_break">--- Section Break ---</option>
               </select>
             </div>
           </div>

           {/* Pre-fill options based on selected type */}
           {['select', 'radio', 'checkbox'].includes(newFieldType) && (
             <div className="bg-blue-50/50 border border-blue-100 p-3 rounded-lg flex flex-col gap-1.5 animate-fadeIn">
               <label className="text-[10px] uppercase font-bold text-blue-800">Choices (Comma Separated)</label>
               <input 
                 type="text" 
                 value={newFieldOptionsStr} 
                 onChange={(e) => setNewFieldOptionsStr(e.target.value)}
                 placeholder="e.g. Option A, Option B, Option C"
                 className="w-full text-xs px-3 py-2 bg-white border border-blue-200 rounded outline-none focus:border-blue-500"
               />
             </div>
           )}

           {['grid_radio', 'grid_checkbox'].includes(newFieldType) && (
             <div className="bg-blue-50/50 border border-blue-100 p-3 rounded-lg flex flex-col sm:flex-row gap-3 animate-fadeIn">
               <div className="flex-1 flex flex-col gap-1.5">
                 <label className="text-[10px] uppercase font-bold text-blue-800">Rows / Questions (Comma Separated)</label>
                 <input 
                   type="text" 
                   value={newFieldGridRowsStr} 
                   onChange={(e) => setNewFieldGridRowsStr(e.target.value)}
                   placeholder="e.g. Quality, Speed, Price"
                   className="w-full text-xs px-3 py-2 bg-white border border-blue-200 rounded outline-none focus:border-blue-500"
                 />
               </div>
               <div className="flex-1 flex flex-col gap-1.5">
                 <label className="text-[10px] uppercase font-bold text-blue-800">Columns / Choices (Comma Separated)</label>
                 <input 
                   type="text" 
                   value={newFieldGridColsStr} 
                   onChange={(e) => setNewFieldGridColsStr(e.target.value)}
                   placeholder="e.g. Good, Neutral, Bad"
                   className="w-full text-xs px-3 py-2 bg-white border border-blue-200 rounded outline-none focus:border-blue-500"
                 />
               </div>
             </div>
           )}

           <div className="flex justify-end mt-1">
             <button 
               type="submit"
               disabled={!newFieldLabel.trim()}
               className="bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold px-5 py-2.5 flex items-center justify-center gap-1.5 shadow-sm transition-all disabled:opacity-50 disabled:pointer-events-none hover:shadow active:scale-95 cursor-pointer"
               id="add-field-submit-btn"
             >
               <Plus size={14} />
               Add Field to Form
             </button>
           </div>
        </div>
      </form>
    </div>
  );
}
