import React, { useState, useEffect } from 'react';
import { FormConfig, ExcelSaveConfig, MSTokens } from './types';
import { getProfile } from './utils/graphHelper';
import FormBuilder from './components/FormBuilder';
import MicrosoftConnector from './components/MicrosoftConnector';
import PublicForm from './components/PublicForm';
import ResponsesViewer from './components/ResponsesViewer';
import EditLogs from './components/EditLogs';
import FormsHome from './components/FormsHome';
import { 
  FileSpreadsheet, 
  Settings, 
  Terminal, 
  PlayCircle, 
  CheckCircle,
  Copy,
  ExternalLink,
  Share2,
  ChevronLeft
} from 'lucide-react';

export default function App() {
  const path = window.location.pathname;

  // ROUTER: Route to Public Form
  if (path.startsWith('/form/')) {
    const formId = path.split('/')[2];
    return <PublicForm formId={formId} />;
  }

  // ROUTER: Route to View Responses
  if (path.startsWith('/responses/')) {
    const formId = path.split('/')[2];
    return <ResponsesViewer formId={formId} />;
  }

  // BUILDER APP LOGIC
  const [activeTab, setActiveTab] = useState<'home' | 'designer' | 'connector' | 'dashboard' | 'logs'>('home');
  const [activeFormId, setActiveFormId] = useState<string | null>(null);
  const [appUrl, setAppUrl] = useState('');
  const [isPublishing, setIsPublishing] = useState(false);
  const [publishedUrl, setPublishedUrl] = useState<string | null>(null);
  
  // A. Form Fields configuration
  const [formConfig, setFormConfig] = useState<FormConfig>({
    title: 'New Form',
    description: 'Please fill out all required details below.',
    settings: { requireMicrosoftLogin: false },
    fields: [
      { id: 'full_name', label: 'Full Name', type: 'short_text', required: true, placeholder: 'Enter name' }
    ]
  });

  // B. Microsoft Graph Tokens
  const [tokens, setTokens] = useState<MSTokens | null>(null);
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // C. Active Excel mapping targets
  const [saveConfig, setSaveConfig] = useState<ExcelSaveConfig | null>(null);

  // Load cached elements
  useEffect(() => {
    const storedTokens = localStorage.getItem('microsoft_tokens');
    if (storedTokens) {
      try {
        const t = JSON.parse(storedTokens);
        setTokens(t);
        getProfile(t, setTokens).then(p => {
          setUserEmail(p.userPrincipalName || p.mail || null);
        }).catch(e => { console.error('Failed to get profile:', e); });
      } catch (e) {}
    }
    fetch('/api/config').then(res => res.json()).then(data => setAppUrl(data.appUrl)).catch(() => {});
  }, []);

  const handleLogin = async () => {
    try {
      setIsLoggingIn(true);
      const clientRedirectUri = `${window.location.origin}/auth/callback`;
      const res = await fetch(`/api/auth/url?redirect_uri=${encodeURIComponent(clientRedirectUri)}`);
      if (!res.ok) throw new Error(await res.text());
      const { url } = await res.json();
      
      const authWindow = window.open(url, 'microsoft_oauth_popup', 'width=600,height=700,status=no,resizable=yes');
      if (!authWindow) {
        alert('Popup blocker active. Please allow popups for M365 Login.');
        setIsLoggingIn(false);
        return;
      }

      const handleMsg = (e: MessageEvent) => {
        const origin = e.origin;
        if (origin !== window.location.origin) return;
        if (e.data?.type === 'OAUTH_AUTH_SUCCESS') {
          const authTokens = e.data.tokens as MSTokens;
          setTokens(authTokens);
          localStorage.setItem('microsoft_tokens', JSON.stringify(authTokens));
          
          getProfile(authTokens, setTokens).then(p => {
             setUserEmail(p.userPrincipalName || p.mail || null);
          }).catch(e => { console.error('Failed to get profile from login:', e); });

          setIsLoggingIn(false);
          window.removeEventListener('message', handleMsg);
        } else if (e.data?.type === 'OAUTH_AUTH_ERROR') {
          alert(`Auth Error: ${e.data.error}`);
          setIsLoggingIn(false);
          window.removeEventListener('message', handleMsg);
        }
      };
      window.addEventListener('message', handleMsg);
    } catch (err: any) {
      alert(err.message);
      setIsLoggingIn(false);
    }
  };

  // If no tokens, force login screen
  if (!tokens) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
         <div className="bg-white p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            <div className="w-16 h-16 bg-blue-600 rounded-xl flex items-center justify-center text-white font-bold text-lg mx-auto mb-6 shadow-sm">BBIPL</div>
            <h1 className="text-2xl font-bold text-slate-800 mb-2">Welcome Back</h1>
            <p className="text-sm text-slate-500 mb-8">Please sign in with your Microsoft 365 account to access the Forms Dashboard.</p>
            <button 
              onClick={handleLogin}
              disabled={isLoggingIn}
              className="w-full bg-slate-800 hover:bg-slate-900 text-white font-semibold py-3 px-4 rounded-xl transition-colors disabled:opacity-50"
            >
              {isLoggingIn ? 'Connecting...' : 'Sign in with Microsoft'}
            </button>
         </div>
       </div>
    );
  }

  // Sync changes
  const handleFormConfigChange = (newConfig: FormConfig) => {
    setFormConfig(newConfig);
  };
  const handleTokensChange = (newTokens: MSTokens | null) => {
    setTokens(newTokens);
    if (newTokens) localStorage.setItem('microsoft_tokens', JSON.stringify(newTokens));
    else localStorage.removeItem('microsoft_tokens');
  };
  const handleSaveConfigChange = (newSaveConfig: ExcelSaveConfig | null) => {
    setSaveConfig(newSaveConfig);
  };

  const handlePublish = async () => {
    if (!saveConfig) {
      alert("Please map the form to an Excel spreadsheet in 'Microsoft 365 Integration' tab before publishing.");
      return;
    }
    if (!tokens) {
      alert("Please connect your Microsoft Account before publishing.");
      return;
    }

    setIsPublishing(true);
    try {
      const isNew = !activeFormId;
      const url = isNew ? '/api/forms' : `/api/forms/${activeFormId}`;
      const method = isNew ? 'POST' : 'PUT';

      const response = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          config: formConfig,
          excelConfig: saveConfig,
          creatorTokens: tokens,
          creatorEmail: userEmail
        })
      });
      if (!response.ok) throw new Error('Failed to publish form.');
      
      const data = await response.json();
      const finalId = isNew ? data.id : activeFormId;
      
      if (isNew) {
         setActiveFormId(finalId);
      }
      
      const pUrl = `${window.location.origin}/form/${finalId}`;
      setPublishedUrl(pUrl);
      
      if (!isNew) {
         alert('Live form updated successfully!');
      }
    } catch (err: any) {
      alert(`Publish Error: ${err.message}`);
    } finally {
      setIsPublishing(false);
    }
  };

  const handleEditForm = (id: string | null, config: FormConfig, excelConfig: ExcelSaveConfig | null) => {
    setActiveFormId(id);
    setFormConfig(config);
    setSaveConfig(excelConfig);
    if (id) {
       setPublishedUrl(`${window.location.origin}/form/${id}`);
    } else {
       setPublishedUrl(null);
    }
    setActiveTab('designer');
  };

  const handleBackToHome = () => {
    setActiveFormId(null);
    setSaveConfig(null);
    setPublishedUrl(null);
    setActiveTab('home');
  };

  if (activeTab === 'home') {
    return (
       <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans">
          <nav className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-6 sm:px-8 shadow-sm shrink-0 sticky top-0 z-50">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center text-white font-bold text-sm text-[10px]">BBIPL</div>
              <span className="font-semibold text-lg tracking-tight text-slate-800">
                Forms Dashboard <span className="text-slate-450 font-normal text-sm">v3.0</span>
              </span>
            </div>
            <div className="flex items-center gap-4">
              <div className={`flex items-center gap-2 text-xs sm:text-sm font-medium px-3.5 py-1.5 rounded-full border transition-colors ${
                tokens ? 'text-emerald-700 bg-emerald-50 border-emerald-200' : 'text-amber-700 bg-amber-50 border-amber-200'
              }`}>
                <div className={`w-2 h-2 rounded-full ${tokens ? 'bg-emerald-500' : 'bg-amber-500'}`}></div>
                {tokens ? 'M365 Linked' : 'Integration Pending'}
              </div>
            </div>
          </nav>
          <FormsHome onEditForm={handleEditForm} userEmail={userEmail} />
       </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans">
      <nav className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 sm:px-6 shadow-sm shrink-0 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <button onClick={handleBackToHome} className="p-1.5 text-slate-500 hover:bg-slate-100 hover:text-slate-800 rounded-md transition-colors mr-1">
             <ChevronLeft size={20} />
          </button>
          <div className="hidden sm:flex w-8 h-8 bg-blue-600 rounded items-center justify-center text-white font-bold text-sm text-[10px]">BBIPL</div>
          <span className="font-semibold text-base sm:text-lg tracking-tight text-slate-800 truncate max-w-[200px] sm:max-w-[400px]">
            {formConfig.title || 'Untitled Form'}
          </span>
        </div>
        
        <div className="hidden sm:flex items-center gap-4">
          <div className={`flex items-center gap-2 text-xs sm:text-sm font-medium px-3.5 py-1.5 rounded-full border transition-colors ${
            tokens && saveConfig ? 'text-emerald-700 bg-emerald-50 border-emerald-200' : 'text-amber-700 bg-amber-50 border-amber-200'
          }`}>
            <div className={`w-2 h-2 rounded-full ${tokens && saveConfig ? 'bg-emerald-500' : 'bg-amber-500'}`}></div>
            {tokens && saveConfig ? 'M365 Linked' : 'Integration Pending'}
          </div>
        </div>
      </nav>

      <div className="bg-slate-100/65 border-b border-slate-250 py-2.5 px-4 sm:px-6 shrink-0 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-1.5 bg-white p-1 rounded-xl border border-slate-200/80 shadow-xs flex-wrap">
          <button
            onClick={() => setActiveTab('designer')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold cursor-pointer transition-all ${
              activeTab === 'designer' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-800 hover:bg-slate-50'
            }`}
          >
            1. Form Builder
          </button>
          <button
            onClick={() => setActiveTab('connector')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold cursor-pointer transition-all ${
              activeTab === 'connector' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-800 hover:bg-slate-50'
            }`}
          >
            2. Settings
            {saveConfig && <span className="w-1.5 h-1.5 rounded-full bg-emerald-400"></span>}
          </button>
        </div>
        <div className="flex gap-2 w-full sm:w-auto">
           {(activeTab === 'designer' || activeTab === 'connector') && (
             <button 
                onClick={handlePublish} disabled={isPublishing || !saveConfig} 
                className="w-full sm:w-auto bg-blue-600 text-white px-4 py-2 rounded-lg font-bold text-xs hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed flex justify-center gap-1.5 items-center transition-colors cursor-pointer shadow-sm"
             >
                <CheckCircle size={14} /> {isPublishing ? 'Saving...' : (activeFormId ? 'Save & Update Form' : 'Publish New Form')}
             </button>
           )}
        </div>
      </div>

      <main className="flex-1 max-w-[1400px] w-full mx-auto p-4 sm:p-6 lg:p-8 grid grid-cols-1 lg:grid-cols-12 gap-6 overflow-hidden">
        <section className={`flex flex-col min-w-0 ${activeTab === 'designer' ? 'lg:col-span-8 pr-0 lg:pr-6 border-r border-slate-200' : 'lg:col-span-12'}`}>
          <div className="flex-1 flex flex-col">
            {activeTab === 'designer' && (
              <FormBuilder config={formConfig} onChange={handleFormConfigChange} />
            )}
            {activeTab === 'connector' && (
              <MicrosoftConnector 
                formConfig={formConfig} setFormConfig={handleFormConfigChange} tokens={tokens} setTokens={handleTokensChange}
                saveConfig={saveConfig} setSaveConfig={handleSaveConfigChange} publishedUrl={publishedUrl} formId={activeFormId}
              />
            )}
          </div>
        </section>

        {activeTab === 'designer' && (
          <aside className="lg:col-span-4 flex flex-col gap-5">
             {publishedUrl && (
               <div className="bg-emerald-50 rounded-xl border border-emerald-200 p-5 shadow-sm">
                <h3 className="text-xs font-bold text-emerald-800 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <CheckCircle size={14} /> LIVE FORM
                </h3>
                <p className="text-xs text-emerald-700 mb-4 font-medium leading-relaxed">
                  Your form is currently live. People can fill this URL, and entries will go directly to your configured Excel sheet. You can edit the form and click "Update" above.
                </p>
                <div className="flex bg-white border border-emerald-200 rounded-lg overflow-hidden shadow-xs">
                   <input type="text" readOnly value={publishedUrl} className="w-full text-xs text-slate-700 pl-3 py-2 bg-transparent focus:outline-none" />
                   <button onClick={() => { navigator.clipboard.writeText(publishedUrl); alert('Link copied!'); }} className="bg-emerald-100 hover:bg-emerald-200 px-3 text-emerald-800 transition-colors flex items-center justify-center shrink-0 border-l border-emerald-200 cursor-pointer" title="Copy">
                     <Copy size={13} />
                   </button>
                   <a href={publishedUrl} target="_blank" rel="noopener noreferrer" className="bg-emerald-100 hover:bg-emerald-200 px-3 text-emerald-800 transition-colors flex items-center justify-center shrink-0 border-l border-emerald-200 cursor-pointer" title="Open Link">
                     <ExternalLink size={13} />
                   </a>
                </div>
             </div>
           )}

          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm flex flex-col gap-4">
            <div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">Sync Mapping Status</h3>
              <div className="space-y-3 font-mono text-[11px] leading-normal">
                {saveConfig ? (
                  <div className="space-y-1 bg-slate-50 p-3 rounded-lg border border-slate-100">
                    <div className="flex items-center gap-1.5 text-slate-700 font-bold">
                      <CheckCircle className="text-emerald-500 w-3.5 h-3.5 shrink-0" />
                      <span>Mapping Synced</span>
                    </div>
                    <div className="pl-5 text-slate-500 space-y-0.5 mt-1">
                      <p>• Tab: {saveConfig.sheetName}</p>
                      <p>• Obj: {saveConfig.tableName}</p>
                      <p>• File: {saveConfig.fileName}</p>
                    </div>
                  </div>
                ) : (
                  <div className="bg-rose-50/50 p-3 rounded-lg border border-rose-100 text-rose-800">
                    <div className="font-bold flex items-center gap-1">
                      <span>⚠️ Action Required</span>
                    </div>
                    <p className="text-[10px] text-rose-700/80 font-sans mt-1 leading-relaxed">
                      Please link your Microsoft 365 Account and map to an Excel Spreadsheet before publishing the form.
                    </p>
                  </div>
                )}
              </div>
            </div>

            <div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Form Column Headers</h3>
              <div className="bg-slate-50 p-3 rounded-lg border border-slate-100 space-y-1 max-h-[160px] overflow-y-auto pr-1">
                <p className="font-mono text-[10px] text-slate-400">1. Submission ID</p>
                <p className="font-mono text-[10px] text-slate-400">2. Submitted At</p>
                {formConfig.fields.map((f, idx) => (
                  <p key={f.id} className="font-mono text-[10px] text-slate-700 truncate" title={f.label}>
                    {idx + 3}. {f.label} {f.required && <span className="text-red-500">*</span>}
                  </p>
                ))}
              </div>
            </div>
          </div>
        </aside>
      )}
      </main>
    </div>
  );
}
