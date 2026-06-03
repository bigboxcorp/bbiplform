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
  const [respondentEmail, setRespondentEmail] = useState<string>('');

  useEffect(() => {
    fetch(`/api/forms/${formId}`)
      .then(res => {
        if (!res.ok) throw new Error('Form not found or unavailable.');
        return res.json();
      })
      .then(data => {
        // Check local storage for multiple submissions rule before continuing
        if (data.config?.settings?.allowMultipleSubmissions === false) {
            if (localStorage.getItem(`__form_submitted_${data.id}`)) {
                setHasAlreadySubmitted(true);
            }
        }
        
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
                    let nextSlot = null;
                    let nextSlotTomorrow = null;

                    // Sort ranges by start time
                    const sortedRanges = [...ranges].sort((a: any, b: any) => {
                         if (!a.start || !b.start) return 0;
                         const aMins = parseInt(a.start.split(':')[0]) * 60 + parseInt(a.start.split(':')[1]);
                         const bMins = parseInt(b.start.split(':')[0]) * 60 + parseInt(b.start.split(':')[1]);
                         return aMins - bMins;
                    });

                    for (const r of sortedRanges) {
                         if (r.start && r.end) {
                             const sH = parseInt(r.start.split(':')[0]);
                             const sM = parseInt(r.start.split(':')[1]);
                             const startMinutes = sH * 60 + sM;
                             if (startMinutes > currentMinutes) {
                                 nextSlot = r;
                                 break;
                             }
                             if (!nextSlotTomorrow) nextSlotTomorrow = r;
                         }
                    }

                    if (!nextSlot && nextSlotTomorrow) nextSlot = nextSlotTomorrow;

                    if (nextSlot) {
                       const formatTime = (timeStr: string) => {
                           let [h, m] = timeStr.split(':').map(Number);
                           const ampm = h >= 12 ? 'PM' : 'AM';
                           h = h % 12 || 12;
                           return `${h}:${m < 10 ? '0'+m : m} ${ampm}`;
                       };
                       setTimeError(`This form is currently closed. It will open again at ${formatTime(nextSlot.start)} to ${formatTime(nextSlot.end)}.`);
                    } else {
                       setTimeError('This form is currently closed and only accepts submissions during specific times of the day.');
                    }
                }
            }
        }


        let storedAuth = localStorage.getItem('respondent_tokens');
        if (!storedAuth) storedAuth = localStorage.getItem('microsoft_tokens');
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

  const needsLogin = formData?.config?.settings?.requireMicrosoftLogin || 
                     formData?.config?.settings?.allowMultipleSubmissions === false ||
                     formData?.config?.settings?.collectEmails;

  useEffect(() => {
     if (respondentTokens && !validProfileChecked) {
         getProfile(respondentTokens, setRespondentTokens).then(async (profile) => {
             const email = profile.userPrincipalName || profile.mail || '';
             (window as any).respondentEmail = email;
             setRespondentEmail(email);
             
             if (needsLogin) {
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
             }
             setValidProfileChecked(true);
         }).catch(err => {
             console.error('Failed profile fetch, token valid?', err);
             setRespondentTokens(null);
             localStorage.removeItem('respondent_tokens');
             if (!needsLogin) setValidProfileChecked(true);
         });
     } else if (!needsLogin) {
         setValidProfileChecked(true);
     }
  }, [formData, respondentTokens, validProfileChecked, formId, needsLogin]);

  const [failedAttempts, setFailedAttempts] = useState(0);
  const [lockoutTimer, setLockoutTimer] = useState(0);

  useEffect(() => {
    let t: NodeJS.Timeout;
    if (lockoutTimer > 0) {
      t = setTimeout(() => {
         setLockoutTimer(lockoutTimer - 1);
         if (lockoutTimer - 1 === 0) setFailedAttempts(0);
      }, 1000);
    }
    return () => clearTimeout(t);
  }, [lockoutTimer]);

  const handleRespondentLogin = async () => {
    if (lockoutTimer > 0) return;

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
         const newFails = failedAttempts + 1;
         setFailedAttempts(newFails);
         if (newFails >= 3) {
            setLockoutTimer(60);
            setAuthError('Too many failed attempts. Try again in 60 seconds.');
         } else {
            setAuthError('Login closed or failed. Please try again.');
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
          setRespondentTokens(authTokens);
          localStorage.setItem('respondent_tokens', JSON.stringify(authTokens));
          setAuthError(null);
          setFailedAttempts(0);
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

  const secureUrl = (url?: string) => url ? url.replace(/^http:\/\//i, 'https://') : '';

  const getBgStyle = () => {
    let style: any = {};
    if (formData?.config.settings?.backgroundColor) {
        style.backgroundColor = formData.config.settings.backgroundColor;
    }
    if (formData?.config.settings?.backgroundUrl) {
        style.backgroundImage = `url(${secureUrl(formData.config.settings.backgroundUrl)})`;
    } else if (formData?.config.settings?.coverUrl) {
        style.backgroundImage = `url(${secureUrl(formData.config.settings.coverUrl)})`;
    }
    return style;
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50" style={getBgStyle()}>
        <div className="text-center font-medium text-slate-500 animate-pulse">Loading form...</div>
      </div>
    );
  }

  if (error || !formData) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50" style={getBgStyle()}>
        <div className="bg-red-50 text-red-600 p-6 rounded-xl border border-red-200">
          <h2 className="font-bold text-lg mb-2">Unavailable</h2>
          <p className="text-sm">{error || 'Could not load the requested form.'}</p>
        </div>
      </div>
    );
  }

  if (needsLogin && (!respondentTokens || !validProfileChecked)) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 bg-cover bg-center" style={getBgStyle()}>
         <div className="bg-white/95 backdrop-blur p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData.config.settings?.logoUrl && <img src={secureUrl(formData.config.settings.logoUrl)} alt="Logo" style={{ height: formData.config.settings.logoSize ? `${Math.min(formData.config.settings.logoSize, 80)}px` : '48px' }} className="mx-auto mb-4 object-contain" />}
            <h1 className="text-xl font-bold text-slate-800 mb-2">Login Required</h1>
            <p className="text-sm text-slate-500 mb-6">This form requires you to sign in with your Microsoft 365 account to verify your identity.</p>
            {authError && <div className="mb-4 text-xs font-bold text-rose-600 bg-rose-50 p-2 rounded border border-rose-200">{authError}</div>}
            <button 
              onClick={handleRespondentLogin}
              disabled={isLoggingIn || lockoutTimer > 0}
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-4 rounded-xl transition-colors disabled:opacity-50 shadow-sm"
              style={formData.config.settings?.themeColor && lockoutTimer === 0 ? { backgroundColor: formData.config.settings.themeColor } : {}}
            >
              {lockoutTimer > 0 ? `Try again in ${lockoutTimer}s` : isLoggingIn ? 'Connecting...' : 'Sign in with Microsoft'}
            </button>
         </div>
       </div>
    );
  }

  if (timeError) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 bg-cover bg-center" style={getBgStyle()}>
         <div className="bg-white/95 backdrop-blur p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData?.config.settings?.logoUrl && <img src={secureUrl(formData.config.settings.logoUrl)} alt="Logo" style={{ height: formData.config.settings.logoSize ? `${Math.min(formData.config.settings.logoSize, 80)}px` : '48px' }} className="mx-auto mb-4 object-contain" />}
            <h1 className="text-2xl font-bold text-slate-800 mb-2">Form Closed</h1>
            <p className="text-sm text-slate-600 font-medium mb-6">{timeError}</p>
         </div>
       </div>
    );
  }

  if (hasAlreadySubmitted) {
    return (
       <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 bg-cover bg-center" style={getBgStyle()}>
         <div className="bg-white/95 p-8 rounded-2xl shadow-xl max-w-sm w-full text-center border border-slate-200">
            {formData.config.settings?.logoUrl && <img src={secureUrl(formData.config.settings.logoUrl)} alt="Logo" style={{ height: formData.config.settings.logoSize ? `${Math.min(formData.config.settings.logoSize, 80)}px` : '48px' }} className="mx-auto mb-4 object-contain" />}
            <h1 className="text-xl font-bold text-slate-800 mb-2">Already Responded</h1>
            <p className="text-sm text-slate-500 mb-6">You can only fill out this form once.</p>
         </div>
       </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 py-10 px-4 sm:px-6 bg-cover bg-center" style={getBgStyle()}>
      <div className="max-w-3xl mx-auto relative">
        {formData.config.settings?.requireMicrosoftLogin && validProfileChecked && respondentEmail && (
           <div className="absolute -top-7 right-0 text-[10px] text-slate-500 bg-white/50 px-2 py-0.5 rounded backdrop-blur border border-slate-200">
             Logged in as: <span className="font-semibold text-slate-700">{respondentEmail}</span>
           </div>
        )}
        <FormSubmissionForm 
           formConfig={formData.config} 
           formId={formData.id} 
           saveConfig={formData.excelConfig} 
           publicMode={true}
           userEmail={respondentEmail}
        />
      </div>
    </div>
  );
}
