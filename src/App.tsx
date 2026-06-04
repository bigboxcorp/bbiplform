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
  ChevronLeft,
  FileText
} from 'lucide-react';

export default function App() {
  const path = window.location.pathname;
  const urlParams = new URLSearchParams(window.location.search);
  const formIdQS = urlParams.get('form');
  const responsesIdQS = urlParams.get('responses');

  // ROUTER: Route to Public Form
  if (formIdQS) {
    return <PublicForm formId={formIdQS} />;
  }
  if (path.startsWith('/f/')) {
    const formId = path.split('/f/')[1].split('/')[0].split('?')[0];
    if (formId) return <PublicForm formId={formId} />;
  }
  if (path.includes('/form/')) {
    const formId = path.split('/form/')[1].split('/')[0].split('?')[0];
    if (formId) return <PublicForm formId={formId} />;
  }

  // ROUTER: Route to View Responses
  if (responsesIdQS) {
    return <ResponsesViewer formId={responsesIdQS} />;
  }
  if (path.startsWith('/r/')) {
    const formId = path.split('/r/')[1].split('/')[0].split('?')[0];
    if (formId) return <ResponsesViewer formId={formId} />;
  }
  if (path.includes('/responses/')) {
    const formId = path.split('/responses/')[1].split('/')[0].split('?')[0];
    if (formId) return <ResponsesViewer formId={formId} />;
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
          const mail = p.userPrincipalName || p.mail || null;
          setUserEmail(mail);
          if (mail) localStorage.setItem('microsoft_user_email', mail);
        }).catch(e => { console.error('Failed to get profile:', e); });
      } catch (e) {}
    }
    fetch('/api/config').then(res => res.json()).then(data => setAppUrl(data.appUrl)).catch(() => {});
  }, []);

  const [loginFailedAttempts, setLoginFailedAttempts] = useState(0);
  const [loginTimer, setLoginTimer] = useState(0);

  useEffect(() => {
    if (loginTimer > 0) {
      const t = setTimeout(() => {
         setLoginTimer(prev => prev - 1);
         if (loginTimer - 1 === 0) setLoginFailedAttempts(0);
      }, 1000);
      return () => clearTimeout(t);
    }
  }, [loginTimer]);

  const handleLogin = async () => {
    if (loginTimer > 0) return;
    try {
      setIsLoggingIn(true);
      const isLocalhost = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1');
      const origin = isLocalhost ? window.location.origin : window.location.origin.replace(/^http:\/\//i, 'https://');
      const clientRedirectUri = `${origin}/auth/callback`;
      const res = await fetch(`/api/auth/url?redirect_uri=${encodeURIComponent(clientRedirectUri)}`);
      if (!res.ok) throw new Error(await res.text());
      const { url } = await res.json();
      
      const authWindow = window.open(url, 'microsoft_oauth_popup', 'width=600,height=700,status=no,resizable=yes');
      if (!authWindow) {
        alert('Popup blocker active. Please allow popups for M365 Login.');
        setIsLoggingIn(false);
        return;
      }

      let intervalId: NodeJS.Timeout;

      const recordFailure = () => {
         const newFails = loginFailedAttempts + 1;
         setLoginFailedAttempts(newFails);
         if (newFails >= 3) {
            setLoginTimer(60);
         }
         setIsLoggingIn(false);
      };

      const handleMsg = (e: MessageEvent) => {
        const origin = e.origin;
        if (origin !== window.location.origin) return;
        if (e.data?.type === 'OAUTH_AUTH_SUCCESS') {
          clearInterval(intervalId);
          window.removeEventListener('message', handleMsg);
          const authTokens = e.data.tokens as MSTokens;
          setTokens(authTokens);
          localStorage.setItem('microsoft_tokens', JSON.stringify(authTokens));
          
          getProfile(authTokens, setTokens).then(p => {
             const mail = p.userPrincipalName || p.mail || null;
             setUserEmail(mail);
             if (mail) localStorage.setItem('microsoft_user_email', mail);
          }).catch(e => { console.error('Failed to get profile from login:', e); });

          setLoginFailedAttempts(0);
          setIsLoggingIn(false);
        } else if (e.data?.type === 'OAUTH_AUTH_ERROR') {
          clearInterval(intervalId);
          window.removeEventListener('message', handleMsg);
          recordFailure();
        }
      };
      
      window.addEventListener('message', handleMsg);
      
      intervalId = setInterval(() => {
        try {
          if (authWindow?.closed) {
              clearInterval(intervalId);
              window.removeEventListener('message', handleMsg);
              recordFailure();
          }
        } catch (e) {
          // Ignore COOP block errors for authWindow.closed
        }
      }, 1000);

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
            
            {loginTimer > 0 && (
                <div className="mb-4 text-sm text-red-600 bg-red-50 p-3 rounded-lg border border-red-100 font-semibold">
                   Too many failed attempts. Try again in {loginTimer}s
                </div>
            )}
            
            <button 
              onClick={handleLogin}
              disabled={isLoggingIn || loginTimer > 0}
              className="w-full bg-slate-800 hover:bg-slate-900 text-white font-semibold py-3 px-4 rounded-xl transition-colors disabled:opacity-50 flex items-center justify-center gap-2"
            >
              {isLoggingIn ? 'Connecting...' : 'Sign in with Microsoft'}
            </button>
            {loginFailedAttempts > 0 && loginTimer === 0 && !isLoggingIn && (
                <div className="mt-4 text-xs font-medium text-amber-600">
                   Login was closed or failed ({loginFailedAttempts}/3 attempts remain)
                </div>
            )}
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
    if (formConfig.settings?.isMappingLocked === false) {
      alert("Please Complete and Lock the Sync Mapping in 'Microsoft 365 Integration' tab before updating the form.");
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
      
      const isLocalhost = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1');
      const origin = isLocalhost ? window.location.origin : window.location.origin.replace(/^http:\/\//i, 'https://');
      const pUrl = `${origin}/f/${finalId}`;
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
      const isLocalhost = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1');
      const origin = isLocalhost ? window.location.origin : window.location.origin.replace(/^http:\/\//i, 'https://');
      const pUrl = `${origin}/f/${id}`;
      setPublishedUrl(pUrl);
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

  const handleSignoutGlobal = () => {
     setTokens(null);
     setSaveConfig(null);
     localStorage.removeItem('microsoft_tokens');
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
              <div className={`hidden sm:flex items-center gap-2 text-xs sm:text-sm font-medium px-3.5 py-1.5 rounded-full border transition-colors ${
                tokens ? 'text-emerald-700 bg-emerald-50 border-emerald-200' : 'text-amber-700 bg-amber-50 border-amber-200'
              }`}>
                <div className={`w-2 h-2 rounded-full ${tokens ? 'bg-emerald-500' : 'bg-amber-500'}`}></div>
                {tokens ? (
                  <span className="flex flex-col text-left leading-tight text-[10px]">
                     <span className="opacity-80 font-bold">M365 Database Linked</span>
                  </span>
                ) : 'Integration Pending'}
              </div>
              {tokens && (
                 <button onClick={handleSignoutGlobal} className="text-xs text-rose-600 hover:text-rose-800 font-bold px-3 py-1.5 bg-rose-50 hover:bg-rose-100 rounded-lg transition-colors border border-rose-200 cursor-pointer">
                    Logout
                 </button>
              )}
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
          <div className="flex items-center gap-2">
            <div className={`flex items-center gap-2 text-xs sm:text-sm font-medium px-3.5 py-1.5 rounded-full border transition-colors ${
              tokens ? (saveConfig && formConfig.settings?.isMappingLocked === false ? 'text-amber-700 bg-amber-50 border-amber-200' : 'text-emerald-700 bg-emerald-50 border-emerald-200') : 'text-amber-700 bg-amber-50 border-amber-200'
            }`}>
              <div className={`w-2 h-2 rounded-full ${tokens ? (saveConfig && formConfig.settings?.isMappingLocked === false ? 'bg-amber-500' : 'bg-emerald-500') : 'bg-amber-500'}`}></div>
              {tokens ? (
                  <span className="flex flex-col text-left leading-tight text-[10px]">
                     <span className={saveConfig && formConfig.settings?.isMappingLocked === false ? "opacity-100 font-bold" : "opacity-80"}>
                        {saveConfig ? (formConfig.settings?.isMappingLocked === false ? 'Re-sync Setup Required' : 'M365 Database Linked') : 'Database mapping missing'}
                     </span>
                  </span>
              ) : 'Integration Pending'}
            </div>
          </div>
          {tokens && (
               <button onClick={handleSignoutGlobal} className="text-xs text-rose-600 hover:text-rose-800 font-bold px-3 py-1.5 bg-rose-50 hover:bg-rose-100 rounded-lg transition-colors border border-rose-200">
                  Logout
               </button>
          )}
        </div>
      </nav>

      <div className="bg-slate-100/90 backdrop-blur-md border-b border-slate-200 py-2.5 px-4 sm:px-6 shrink-0 flex flex-wrap items-center justify-between gap-4 sticky top-16 z-40">
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
                onClick={handlePublish} disabled={isPublishing || !saveConfig || formConfig.settings?.isMappingLocked === false} 
                className="w-full sm:w-auto bg-blue-600 text-white px-4 py-2 rounded-lg font-bold text-xs hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed flex justify-center gap-1.5 items-center transition-colors cursor-pointer shadow-sm"
             >
                <CheckCircle size={14} /> {isPublishing ? 'Saving...' : (activeFormId ? 'Save & Update Form' : 'Publish New Form')}
             </button>
           )}
        </div>
      </div>

      <main className="flex-1 max-w-[1400px] w-full mx-auto p-4 sm:p-6 lg:p-8 grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
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
                <div className="flex bg-white border border-emerald-200 rounded-lg overflow-hidden shadow-xs mb-3">
                   <input type="text" readOnly value={publishedUrl} className="w-full text-xs text-slate-700 pl-3 py-2 bg-transparent focus:outline-none" />
                   <button onClick={() => { navigator.clipboard.writeText(publishedUrl); alert('Link copied!'); }} className="bg-emerald-100 hover:bg-emerald-200 px-3 text-emerald-800 transition-colors flex items-center justify-center shrink-0 border-l border-emerald-200 cursor-pointer" title="Copy">
                     <Copy size={13} />
                   </button>
                   <a href={publishedUrl} target="_blank" rel="noopener noreferrer" className="bg-emerald-100 hover:bg-emerald-200 px-3 text-emerald-800 transition-colors flex items-center justify-center shrink-0 border-l border-emerald-200 cursor-pointer" title="Open Link">
                     <ExternalLink size={13} />
                   </a>
                </div>

                {saveConfig ? (
                <div className="bg-blue-50/50 rounded-lg border border-blue-100 p-3 mt-4 text-xs">
                  <h4 className="text-[10px] uppercase font-bold text-blue-600 mb-2 flex items-center gap-1.5">
                    <FileText size={12} /> Monitor Responses
                  </h4>
                  <div className="flex bg-white border border-blue-200 rounded-md overflow-hidden shadow-xs">
                    <input type="text" readOnly value={publishedUrl.includes('/f/') ? publishedUrl.replace('/f/', '/r/') : publishedUrl.replace('?form=', '?responses=').replace('/form/', '/responses/')} className="w-full text-[11px] text-slate-700 pl-2 py-1.5 bg-transparent focus:outline-none" />
                    <button onClick={() => { navigator.clipboard.writeText(publishedUrl.includes('/f/') ? publishedUrl.replace('/f/', '/r/') : publishedUrl.replace('?form=', '?responses=').replace('/form/', '/responses/')); alert('Link copied!'); }} className="bg-blue-50 hover:bg-blue-100 px-2 text-blue-800 transition-colors flex items-center justify-center shrink-0 border-l border-blue-200 cursor-pointer" title="Copy">
                      <Copy size={12} />
                    </button>
                    <a href={publishedUrl.includes('/f/') ? publishedUrl.replace('/f/', '/r/') : publishedUrl.replace('?form=', '?responses=').replace('/form/', '/responses/')} target="_blank" rel="noopener noreferrer" className="bg-blue-50 hover:bg-blue-100 px-2 text-blue-800 transition-colors flex items-center justify-center shrink-0 border-l border-blue-200 cursor-pointer" title="Open Link">
                      <ExternalLink size={12} />
                    </a>
                  </div>
                </div>
                ) : (
                <div className="bg-slate-50 rounded-lg border border-slate-100 p-3 mt-4 text-xs">
                  <h4 className="text-[10px] uppercase font-bold text-slate-500 mb-2 flex items-center gap-1.5">
                    <FileText size={12} /> Responses Unlinked
                  </h4>
                  <p className="text-slate-500 text-[11px]">Connect this form to an Excel sheet to view responses.</p>
                </div>
                )}
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
                      <p>• Attachments Folder: {saveConfig.uploadFolderPath || (saveConfig.fileName ? saveConfig.fileName.replace(/\.xlsx?$/, '') : 'Not set')}</p>
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
