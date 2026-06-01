import React, { useState, useEffect } from 'react';
import { FormConfig, SavedForm, MSTokens } from '../types';
import FormSubmissionForm from './FormSubmissionForm';
import { getProfile } from '../utils/graphHelper';

export default function PublicForm({ formId }: { formId: string }) {
  const [formData, setFormData] = useState<SavedForm | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [respondentTokens, setRespondentTokens] = useState<MSTokens | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);
  const [validProfileChecked, setValidProfileChecked] = useState(false);

  useEffect(() => {
    fetch(`/api/forms/${formId}`)
      .then(res => {
        if (!res.ok) throw new Error('Form not found or unavailable.');
        return res.json();
      })
      .then(data => {
        setFormData(data);
        setLoading(false);

        // Check Time Constraints
        const ranges = data.config.settings?.dailyTimeRanges;
        if (ranges && ranges.length > 0) {
            const hasValidRanges = ranges.some((r: any) => r.start && r.end);
            if (hasValidRanges) {
                const now = new Date();
                const currentMinutes = now.getHours() * 60 + now.getMinutes();
                
                let isAllowed = false;
                for (const r of ranges) {
                    if (r.start && r.end) {
                        const [sH, sM] = r.start.split(':').map(Number);
                        const [eH, eM] = r.end.split(':').map(Number);
                        const startMinutes = sH * 60 + sM;
                        const endMinutes = eH * 60 + eM;
                        if (currentMinutes >= startMinutes && currentMinutes <= endMinutes) {
                            isAllowed = true;
                            break;
                        }
                    }
                }
                if (!isAllowed) {
                    setTimeError('This form is currently closed and only accepts submissions during specific times of the day.');
                }
            }
        }


        const storedAuth = localStorage.getItem('respondent_tokens');
        if (storedAuth) {
           try { setRespondentTokens(JSON.parse(storedAuth)); } catch(e){}
        }
      })
      .catch(err => {
        setError(err.message);
        setLoading(false);
      });
  }, [formId]);

  const [hasAlreadySubmitted, setHasAlreadySubmitted] = useState<boolean>(false);

  const [timeError, setTimeError] = useState<string | null>(null);

  useEffect(() => {
     if (formData?.config?.settings?.requireMicrosoftLogin || formData?.config?.settings?.allowMultipleSubmissions === false) {
         if (respondentTokens && !validProfileChecked) {
             getProfile(respondentTokens, setRespondentTokens).then(async (profile) => {
                 const email = profile.userPrincipalName || profile.mail || '';
                 (window as any).respondentEmail = email;
                 
                 // If disable multiple submissions, check if already submitted
                 if (formData.config.settings?.allowMultipleSubmissions === false) {
                     try {
                        const checkRes = await fetch(`/api/forms/${formId}/check-submission?email=${encodeURIComponent(email)}`);
                        const checkData = await checkRes.json();
                        if (checkData.hasSubmitted) {
                            setHasAlreadySubmitted(true);
                        }
                     } catch(e) {}
                 }

                 const allowedStr = formData.config.settings?.allowedDomains || '';
                 if (allowedStr.trim()) {
                     const allowed = allowedStr.split(',').map((s: string) => s.trim().toLowerCase());
                     const domain = email.split('@')[1]?.toLowerCase();
                     if (!allowed.includes(domain)) {
                         setAuthError(`Your email domain (${domain}) is not authorized to fill out this form.`);
                         return;
                     }
                 }
                 setValidProfileChecked(true);
             }).catch(err => {
                 console.error('Failed profile fetch, token valid?', err);
                 setRespondentTokens(null);
                 localStorage.removeItem('respondent_tokens');
             });
         }
     } else {
         setValidProfileChecked(true);
     }
  }, [formData, respondentTokens, validProfileChecked]);

  const handleRespondentLogin = async () => {
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
          setRespondentTokens(authTokens);
          localStorage.setItem('respondent_tokens', JSON.stringify(authTokens));
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

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="text-center font-medium text-slate-500 animate-pulse">Loading form...</div>
      </div>
    );
  }

  if (error || !formData) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="bg-red-50 text-red-600 p-6 rounded-xl border border-red-200">
          <h2 className="font-bold text-lg mb-2">Unavailable</h2>
          <p className="text-sm">{error || 'Could not load the requested form.'}</p>
        </div>
      </div>
    );
  }

  if ((formData.config.settings?.requireMicrosoftLogin || formData.config.settings?.allowMultipleSubmissions === false) && (!respondentTokens || !validProfileChecked)) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 bg-cover bg-center" style={{ backgroundImage: formData?.config.settings?.coverUrl ? `url(${formData.config.settings.coverUrl})` : 'none' }}>
         <div className="bg-white/95 backdrop-blur p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData.config.settings?.logoUrl && <img src={formData.config.settings.logoUrl} alt="Logo" className="h-12 mx-auto mb-4" />}
            <h1 className="text-xl font-bold text-slate-800 mb-2">Login Required</h1>
            <p className="text-sm text-slate-500 mb-6">This form requires you to sign in with your Microsoft 365 account to verify your identity.</p>
            {authError && <div className="mb-4 text-xs font-bold text-rose-600 bg-rose-50 p-2 rounded border border-rose-200">{authError}</div>}
            <button 
              onClick={handleRespondentLogin}
              disabled={isLoggingIn}
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-4 rounded-xl transition-colors disabled:opacity-50 shadow-sm"
              style={formData.config.settings?.themeColor ? { backgroundColor: formData.config.settings.themeColor } : {}}
            >
              {isLoggingIn ? 'Connecting...' : 'Sign in with Microsoft'}
            </button>
         </div>
       </div>
    );
  }

  if (timeError) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 bg-cover bg-center" style={{ backgroundImage: formData?.config.settings?.coverUrl ? `url(${formData.config.settings.coverUrl})` : 'none' }}>
         <div className="bg-white/95 backdrop-blur p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData?.config.settings?.logoUrl && <img src={formData.config.settings.logoUrl} alt="Logo" className="h-12 mx-auto mb-4" />}
            <h1 className="text-2xl font-bold text-slate-800 mb-2">Form Closed</h1>
            <p className="text-sm text-slate-600 font-medium mb-6">{timeError}</p>
         </div>
       </div>
    );
  }

  if (hasAlreadySubmitted) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
         <div className="bg-white p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData.config.settings?.logoUrl && <img src={formData.config.settings.logoUrl} alt="Logo" className="h-12 mx-auto mb-4" />}
            <h1 className="text-xl font-bold text-slate-800 mb-2">Already Responded</h1>
            <p className="text-sm text-slate-500 mb-6">You can only fill out this form once.</p>
         </div>
       </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 py-10 px-4 sm:px-6">
      <div className="max-w-3xl mx-auto">
        <FormSubmissionForm 
           formConfig={formData.config} 
           formId={formData.id} 
           saveConfig={formData.excelConfig} 
           publicMode={true}
        />
      </div>
    </div>
  );
}
