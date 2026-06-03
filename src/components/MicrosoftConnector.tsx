import React, { useEffect, useState } from 'react';
import { FormConfig, MSTokens, MSGroup, MSChannel, MSDriveItem, MSWorksheet, MSTable, ExcelSaveConfig } from '../types';
import { 
  getProfile, 
  getJoinedTeams, 
  getTeamChannels, 
  getTeamDrive, 
  getExcelFilesInChannel, 
  getWorkbookWorksheets, 
  getWorkbookTables,
  createExcelFileWithTable,
  addColumnToTable,
  clearTableData,
  getTableColumns,
  createWorksheet,
  createTableInWorksheet,
  getFoldersInChannel,
  createFolderInChannel
} from '../utils/graphHelper';
import { 
  Lock, 
  RefreshCw, 
  HelpCircle, 
  CheckCircle2, 
  ChevronRight, 
  AlertCircle, 
  FolderLock, 
  FileSpreadsheet, 
  PlusCircle, 
  Grid3X3, 
  Sparkles,
  Layers,
  Settings,
  XCircle,
  LogOut,
  Copy,
  ExternalLink
} from 'lucide-react';

interface MicrosoftConnectorProps {
  formConfig: FormConfig;
  setFormConfig: (c: FormConfig) => void;
  tokens: MSTokens | null;
  setTokens: (t: MSTokens | null) => void;
  saveConfig: ExcelSaveConfig | null;
  setSaveConfig: (c: ExcelSaveConfig | null) => void;
  publishedUrl?: string | null;
  formId?: string | null;
}

export default function MicrosoftConnector({
  formConfig,
  setFormConfig,
  tokens,
  setTokens,
  saveConfig,
  setSaveConfig,
  publishedUrl,
  formId
}: MicrosoftConnectorProps) {
  const secureUrl = (url?: string) => url ? url.replace(/^http:\/\//i, 'https://') : '';

  // Config & Env status
  const [hasCreds, setHasCreds] = useState<boolean | null>(null);
  const [appUrl, setAppUrl] = useState('');
  const [msClientId, setMsClientId] = useState('');
  const [isConfigLoading, setIsConfigLoading] = useState(true);

  // Profile status
  const [profile, setProfile] = useState<any>(null);

  // Lists & dropdowns state
  const [teams, setTeams] = useState<MSGroup[]>([]);
  const [channels, setChannels] = useState<MSChannel[]>([]);
  const [excelFiles, setExcelFiles] = useState<MSDriveItem[]>([]);
  const [worksheets, setWorksheets] = useState<MSWorksheet[]>([]);
  const [tables, setTables] = useState<MSTable[]>([]);
  
  // Selection States
  const [selectedTeamId, setSelectedTeamId] = useState('');
  const [selectedChannelId, setSelectedChannelId] = useState('');
  const [selectedChannelName, setSelectedChannelName] = useState('');
  const [driveId, setDriveId] = useState('');
  const [selectedFileId, setSelectedFileId] = useState('');
  const [selectedFileName, setSelectedFileName] = useState('');
  const [selectedSheetName, setSelectedSheetName] = useState('');
  const [selectedTableName, setSelectedTableName] = useState('');
  const [uploadFolderPath, setUploadFolderPath] = useState('Submissions_Attachments');

  // UI state variables
  const [isLoading, setIsLoading] = useState(false);
  const [apiError, setApiError] = useState<string | null>(null);
  const [authStatus, setAuthStatus] = useState<'idle' | 'authorizing' | 'success' | 'error'>('idle');

  // Excel Creation details
  const [newExcelFileName, setNewExcelFileName] = useState('');
  const [isCreatingExcel, setIsCreatingExcel] = useState(false);
  const [isModifyingTable, setIsModifyingTable] = useState(false);
  
  const [attachmentFolders, setAttachmentFolders] = useState<any[]>([]);
  const [isCreatingFolder, setIsCreatingFolder] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  
  const [isCreatingWorksheet, setIsCreatingWorksheet] = useState(false);
  const [newWorksheetName, setNewWorksheetName] = useState('');
  
  const [isCreatingTable, setIsCreatingTable] = useState(false);
  const [newTableName, setNewTableName] = useState('');

  const handleClearTable = async () => {
    if (!selectedFileId || !selectedTableName) return;
    if (!confirm('Are you sure you want to clear all data in this table? This cannot be undone.')) return;
    setIsModifyingTable(true);
    try {
      await clearTableData(driveId, selectedFileId, selectedTableName, tokens!, setTokens);
      alert('Table data cleared successfully.');
    } catch(err: any) {
      alert('Failed to clear table: ' + err.message);
    } finally {
      setIsModifyingTable(false);
    }
  };

  const handleAddMissingColumns = async () => {
    if (!selectedFileId || !selectedTableName) return;
    setIsModifyingTable(true);
    try {
      const existingCols = await getTableColumns(driveId, selectedFileId, selectedTableName, tokens!, setTokens);
      const existingColNames = existingCols.map((c: any) => c.name.toLowerCase());
      
      const requiredCols = [
        'Submission ID',
        'Submitted At',
        ...formConfig.fields.map(f => f.label)
      ];

      let addedCols = 0;
      for (const col of requiredCols) {
        if (!existingColNames.includes(col.toLowerCase())) {
          await addColumnToTable(driveId, selectedFileId, selectedTableName, col, tokens!, setTokens);
          addedCols++;
          await new Promise(resolve => setTimeout(resolve, 800));
        }
      }
      alert(`Successfully added ${addedCols} missing columns.`);
    } catch(err: any) {
      alert('Failed to add columns: ' + err.message);
    } finally {
      setIsModifyingTable(false);
    }
  };

  // 1. Fetch system credentials status on startup
  useEffect(() => {
    fetchConfig();
  }, []);

  useEffect(() => {
    const restoreSetup = async () => {
       if (saveConfig && !selectedTeamId && tokens) {
         if (saveConfig.uploadFolderPath) setUploadFolderPath(saveConfig.uploadFolderPath);
         try {
            setIsLoading(true);
            setSelectedTeamId(saveConfig.groupId);
            setSelectedChannelId(saveConfig.channelId);
            setSelectedChannelName(saveConfig.channelName);
            setSelectedFileId((saveConfig as any).fileId || saveConfig.driveItemId);
            setSelectedFileName(saveConfig.fileName);
            setSelectedTableName(saveConfig.tableName);
            setSelectedSheetName(saveConfig.sheetName);
            
            const driveInfo = await getTeamDrive(saveConfig.groupId, tokens, setTokens);
            setDriveId(driveInfo.id);
            const userChannels = await getTeamChannels(saveConfig.groupId, driveInfo.id, tokens, setTokens);
            setChannels(userChannels);
            
            const files = await getExcelFilesInChannel(driveInfo.id, saveConfig.channelName, tokens, setTokens);
            setExcelFiles(files);
            
            const sheetsList = await getWorkbookWorksheets(driveInfo.id, (saveConfig as any).fileId || saveConfig.driveItemId, tokens, setTokens);
            setWorksheets(sheetsList);
            const tablesList = await getWorkbookTables(driveInfo.id, (saveConfig as any).fileId || saveConfig.driveItemId, tokens, setTokens);
            setTables(tablesList);
         } catch(e) { console.warn(e); } finally { setIsLoading(false); }
       }
    };
    restoreSetup();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [saveConfig, tokens]);

  const fetchConfig = async () => {
    try {
      setIsConfigLoading(true);
      const res = await fetch('/api/config');
      const data = await res.json();
      setHasCreds(data.hasCredentials);
      setAppUrl(data.appUrl);
      setMsClientId(data.microsoftClientId);
    } catch (err) {
      console.error('Failed to retrieve server config details:', err);
    } finally {
      setIsConfigLoading(false);
    }
  };

  // 2. Fetch Microsoft profile and available teams when token connects
  useEffect(() => {
    if (tokens && tokens.accessToken) {
      loadProfileAndTeams();
    } else {
      setProfile(null);
      setTeams([]);
    }
  }, [tokens]);

  const loadProfileAndTeams = async () => {
    if (!tokens) return;
    try {
      setIsLoading(true);
      setApiError(null);
      const userProf = await getProfile(tokens, setTokens);
      setProfile(userProf);
      
      const userTeams = await getJoinedTeams(tokens, setTokens);
      setTeams(userTeams);
    } catch (err: any) {
      setApiError(`Could not load M365 Profile. Please login again. Details: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // 3. User selects a Team (Group) -> Load channels
  const handleTeamChange = async (teamId: string) => {
    setSelectedTeamId(teamId);
    setSelectedChannelId('');
    setSelectedChannelName('');
    setDriveId('');
    setExcelFiles([]);
    setSelectedFileId('');
    
    if (!teamId || !tokens) return;

    try {
      setIsLoading(true);
      setApiError(null);
      // Get SharePoint Drive ID
      const driveInfo = await getTeamDrive(teamId, tokens, setTokens);
      setDriveId(driveInfo.id);

      // Get Folders in Drive instead of just Teams Channels
      const userChannels = await getTeamChannels(teamId, driveInfo.id, tokens, setTokens);
      setChannels(userChannels);
    } catch (err: any) {
      setApiError(`Failed to load Team folders: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // 4. User selects a Channel -> Load files
  const handleChannelChange = async (channelId: string) => {
    setSelectedChannelId(channelId);
    setSelectedFileId('');
    setSelectedFileName('');
    setExcelFiles([]);
    setAttachmentFolders([]);

    const selectedChanObj = channels.find(c => c.id === channelId);
    if (!selectedChanObj || !driveId || !tokens) return;

    setSelectedChannelName(selectedChanObj.displayName || '');

    try {
      setIsLoading(true);
      setApiError(null);
      // List Excel documents in channel folder
      const files = await getExcelFilesInChannel(
        driveId,
        selectedChanObj.displayName || '',
        tokens,
        setTokens
      );
      setExcelFiles(files);
      
      const folders = await getFoldersInChannel(
        driveId,
        selectedChanObj.displayName || '',
        tokens,
        setTokens
      );
      setAttachmentFolders(folders);
      
      if (!uploadFolderPath) setUploadFolderPath('Submissions_Attachments');
    } catch (err: any) {
      setApiError(`Failed to load Excel files in channel: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // 5. User selects an Excel File -> Load worksheets and tables
  const handleFileChange = async (fileId: string) => {
    setSelectedFileId(fileId);
    setSelectedSheetName('');
    setSelectedTableName('');
    setWorksheets([]);
    setTables([]);

    const fileObj = excelFiles.find(f => f.id === fileId);
    if (!fileObj || !driveId || !tokens) return;

    setSelectedFileName(fileObj.name || '');

    try {
      setIsLoading(true);
      setApiError(null);
      
      const sheetsList = await getWorkbookWorksheets(driveId, fileId, tokens, setTokens);
      setWorksheets(sheetsList);

      const tablesList = await getWorkbookTables(driveId, fileId, tokens, setTokens);
      setTables(tablesList);
    } catch (err: any) {
      setApiError(`Workbook structural errors: ${err.message}. Make sure files is not locked or corrupted.`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleAutoSetup = async () => {
    if (!driveId || !selectedChannelName || !tokens) {
      alert("Please select a Team and Channel first.");
      return;
    }
    const autoName = formConfig.title ? `${formConfig.title.replace(/[^a-zA-Z0-9 -]/g, '')}` : "Form Responses";
    setNewExcelFileName(autoName);

    try {
      setIsCreatingExcel(true);
      setApiError(null);
      
      const fieldsToSave = formConfig.fields.filter(f => f.type !== 'section_break');
      const headers = ['Submission ID', 'Submitted At'];
      
      if (formConfig.settings?.collectEmails) {
         headers.push('Submitter Email');
      }
      
      headers.push(...fieldsToSave.map(f => f.label));

      const result = await createExcelFileWithTable(
        driveId,
        selectedChannelName,
        autoName,
        'Sheet1',
        headers,
        tokens,
        setTokens
      );

      const createdFile: MSDriveItem = {
        id: result.file.id,
        name: result.file.name,
        webUrl: result.file.webUrl
      };

      setExcelFiles(prev => [createdFile, ...prev]);
      setSelectedFileId(createdFile.id);
      setSelectedFileName(createdFile.name);

      setWorksheets([{ id: 'Sheet1', name: 'Sheet1', position: 1 }]);
      setTables([{ id: result.table.id, name: result.table.name, displayName: result.table.displayName }]);
      
      setSelectedSheetName('Sheet1');
      setSelectedTableName(result.table.name);

      // Auto Map
      const teamObj = teams.find(t => t.id === selectedTeamId);
      const mapping: Record<string, string> = {};
      fieldsToSave.forEach(f => {
        mapping[f.id] = f.label;
      });
      if (formConfig.settings?.collectEmails) {
         mapping['respondent_email'] = 'Submitter Email';
      }

      setSaveConfig({
        groupName: teamObj?.displayName || '',
        channelId: selectedChannelId,
        channelName: selectedChannelName,
        fileName: createdFile.name,
        sheetName: 'Sheet1',
        tableName: result.table.name,
        driveId: driveId,
        groupId: selectedTeamId,
        driveItemId: createdFile.id,
        columnsMapping: mapping,
        uploadFolderPath: uploadFolderPath
      });
      setFormConfig({ ...formConfig, settings: { ...formConfig.settings, isMappingLocked: true } });

      alert(`✅ Auto Setup Complete! File "${autoName}" created and mapped automatically.`);
    } catch(err: any) {
       setApiError(`Could not auto-setup Excel File: ${err.message}`);
    } finally {
       setIsCreatingExcel(false);
    }
  };

  // 6. Action: Automated Creation of Excel file and formatted Table wrapper
  const handleCreateNewExcelFile = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newExcelFileName.trim() || !driveId || !selectedChannelName || !tokens) return;

    try {
      setIsCreatingExcel(true);
      setApiError(null);
      
      const baseName = newExcelFileName.trim();
      const finalFileName = baseName.endsWith('.xlsx') ? baseName : `${baseName}.xlsx`;

      // Derive headers from form fields (Required to set up table)
      const fieldsToSave = formConfig.fields.filter(f => f.type !== 'section_break');
      const headers = ['Submission ID', 'Submitted At'];
      
      if (formConfig.settings?.collectEmails) {
         headers.push('Submitter Email');
      }
      
      headers.push(...fieldsToSave.map(f => f.label));

      const result = await createExcelFileWithTable(
        driveId,
        selectedChannelName,
        finalFileName,
        'Sheet1',
        headers,
        tokens,
        setTokens
      );

      // File is created! Add it to the list
      const createdFile: MSDriveItem = {
        id: result.file.id,
        name: result.file.name,
        webUrl: result.file.webUrl
      };

      setExcelFiles(prev => [createdFile, ...prev]);
      setSelectedFileId(createdFile.id);
      setSelectedFileName(createdFile.name);

      // Sheet1 and Table1 is standard structure from our creation
      setWorksheets([{ id: 'Sheet1', name: 'Sheet1', position: 1 }]);
      setTables([{ id: result.table.id, name: result.table.name, displayName: result.table.displayName }]);
      
      setSelectedSheetName('Sheet1');
      setSelectedTableName(result.table.name);

      // Map everything automatically
      const mapping: Record<string, string> = {
        '__submission_id': 'Submission ID',
        '__submitted_at': 'Submitted At',
      };
      if (formConfig.settings?.collectEmails) {
         mapping['respondent_email'] = 'Submitter Email';
      }
      fieldsToSave.forEach(field => {
        mapping[field.id] = field.label;
      });

      const selectedTeamName = teams.find(t => t.id === selectedTeamId)?.displayName || 'Team';

      setSaveConfig({
        groupId: selectedTeamId,
        groupName: selectedTeamName,
        channelId: selectedChannelId,
        channelName: selectedChannelName,
        driveId: driveId,
        driveItemId: createdFile.id,
        fileName: createdFile.name,
        sheetName: 'Sheet1',
        tableName: result.table.name,
        columnsMapping: mapping
      });

      setNewExcelFileName('');
      alert(`Excel folder me file "${finalFileName}" successfully create ho gayi h! Tables aur mapping automatic match ho gayi hai.`);
    } catch (err: any) {
      setApiError(`Could not create new Excel File: ${err.message}`);
    } finally {
      setIsCreatingExcel(false);
    }
  };

  const handleCreateNewFolder = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFolderName.trim() || !driveId || !tokens || !selectedChannelName) return;

    try {
      setIsCreatingFolder(true);
      setApiError(null);
      await createFolderInChannel(
        driveId,
        selectedChannelName,
        newFolderName.trim(),
        tokens,
        setTokens
      );
      
      const folders = await getFoldersInChannel(driveId, selectedChannelName, tokens, setTokens);
      setAttachmentFolders(folders);
      setUploadFolderPath(newFolderName.trim());
      setNewFolderName('');
    } catch (err: any) {
      setApiError(`Failed to create Folder: ${err.message}`);
    } finally {
      setIsCreatingFolder(false);
    }
  };

  const handleCreateNewWorksheet = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newWorksheetName.trim() || !driveId || !selectedFileId || !tokens) return;
    try {
       setIsCreatingWorksheet(true);
       setApiError(null);
       const sheetName = newWorksheetName.trim();
       await createWorksheet(driveId, selectedFileId, sheetName, tokens, setTokens);
       setWorksheets(prev => [...prev, { id: sheetName, name: sheetName, position: prev.length + 1 }]);
       setSelectedSheetName(sheetName);
       setNewWorksheetName('');
    } catch(err: any) {
       setApiError(`Could not create Worksheet: ${err.message}`);
    } finally {
       setIsCreatingWorksheet(false);
    }
  };

  const handleCreateNewTable = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTableName.trim() || !driveId || !selectedFileId || !selectedSheetName || !tokens) return;
    try {
       setIsCreatingTable(true);
       setApiError(null);
       
       const headers = ['Submission ID', 'Submitted At', ...formConfig.fields.map(f => f.label)];
       const tName = newTableName.trim().replace(/[^a-zA-Z0-9]/g, '');

       let startAddress = 'A1';
       if (tables.length > 0) {
           // To leave 1 column gap, we assume each table is headers.length wide.
           // Offset = total previous tables * (headers.length + 1 gap column)
           const offsetCols = tables.length * (headers.length + 1); 
           let colNum = offsetCols + 1; // 1-based index for columns

           let newLetters = '';
           let temp = colNum;
           while (temp > 0) {
               let rem = (temp - 1) % 26;
               newLetters = String.fromCharCode(65 + rem) + newLetters;
               temp = Math.floor((temp - 1) / 26);
           }
           startAddress = `${newLetters}1`;
       }
       
       const result = await createTableInWorksheet(driveId, selectedFileId, selectedSheetName, tName, headers, startAddress, tokens, setTokens);
       
       setTables(prev => [...prev, { id: result.id, name: result.name, displayName: result.name }]);
       setSelectedTableName(result.name);
       setNewTableName('');

       alert(`✅ Table "${result.name}" created at ${startAddress} with ${headers.length} columns!`);
    } catch(err: any) {
       setApiError(`Could not create Table: ${err.message}`);
    } finally {
       setIsCreatingTable(false);
    }
  };

  // 7. Manual Save/Map Selection settings
  const handleApplyLayoutSelections = async () => {
    if (!selectedTeamId || !selectedChannelId || !selectedFileId || !selectedTableName) {
      alert('Please fill out all dropdown listings before saving config.');
      return;
    }

    try {
       setIsModifyingTable(true);
       const existingCols = await getTableColumns(driveId, selectedFileId, selectedTableName, tokens!, setTokens);
       const colNames = existingCols.map((c: any) => c.name.toLowerCase());
       
       let mismatchedCols = false;
       formConfig.fields.forEach(f => {
         if (!colNames.includes(f.label.trim().toLowerCase())) mismatchedCols = true;
       });

       if (mismatchedCols) {
         const proceed = window.confirm('Warning: Is table me column mapping properly form se match nahi kar rahi hai. Data save hone me error aa sakta hai.\n\n"Cancel" press karein and "Add Missing Form Columns" pe click karein existing table modify karne ke liye. Ya phir "Create Fresh Excel Document" try karein.\n\nFir bhi proceed karna chahte hain?');
         if (!proceed) {
             setIsModifyingTable(false);
             return;
         }
       }
    } catch(err) {
        console.error(err);
    } finally {
        setIsModifyingTable(false);
    }

    const teamObj = teams.find(t => t.id === selectedTeamId);
    
    const fieldsToSave = formConfig.fields.filter(f => f.type !== 'section_break');
    
    // Auto map form fields to Table Columns if possible
    const mapping: Record<string, string> = {
      '__submission_id': 'Submission ID',
      '__submitted_at': 'Submitted At'
    };
    if (formConfig.settings?.collectEmails) {
       mapping['respondent_email'] = 'Submitter Email';
    }
    
    // Default Map based on identical names
    fieldsToSave.forEach(f => {
      mapping[f.id] = f.label;
    });

    setSaveConfig({
      groupId: selectedTeamId,
      groupName: teamObj?.displayName || 'Group',
      channelId: selectedChannelId,
      channelName: selectedChannelName,
      driveId: driveId,
      driveItemId: selectedFileId,
      fileName: selectedFileName,
      sheetName: selectedSheetName || 'Sheet1',
      tableName: selectedTableName,
      columnsMapping: mapping,
      uploadFolderPath: uploadFolderPath
    });
    setFormConfig({ ...formConfig, settings: { ...formConfig.settings, isMappingLocked: true } });
  };

  // 8. Trigger Microsoft OAuth auth-code popup flow
  const handleConnectM365 = async () => {
    try {
      const timeout = localStorage.getItem('m365_admin_login_timeout');
      if (timeout) {
         const timeLeft = parseInt(timeout) - Date.now();
         if (timeLeft > 0) {
            setApiError(`Too many failed attempts. Please wait ${Math.ceil(timeLeft / 1000)} seconds.`);
            return;
         } else {
            localStorage.removeItem('m365_admin_login_timeout');
            localStorage.removeItem('m365_admin_login_fails');
         }
      }

      setAuthStatus('authorizing');
      setApiError(null);

      // Request auth url from express server
      const isLocalhost = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1');
      const origin = isLocalhost ? window.location.origin : window.location.origin.replace(/^http:\/\//i, 'https://');
      const clientRedirectUri = `${origin}/auth/callback`;
      const res = await fetch(`/api/auth/url?redirect_uri=${encodeURIComponent(clientRedirectUri)}`);
      if (!res.ok) {
        throw new Error(await res.text());
      }
      
      const { url } = await res.json();

      // Open Popup relative window
      const authWindow = window.open(
        url,
        'microsoft_oauth_popup',
        'width=600,height=700,status=no,resizable=yes'
      );

      if (!authWindow) {
        alert('Prompt: Popup blocker active. Please allow popups for M365 Login.');
        setAuthStatus('idle');
        return;
      }
      
      let failCount = Number(localStorage.getItem('m365_admin_login_fails') || '0');
      let checkInterval: NodeJS.Timeout;

      // Keep registering listener for popMessage responses
      const handleMsg = (e: MessageEvent) => {
        const origin = e.origin;
        if (origin !== window.location.origin) {
          return;
        }

        if (e.data?.type === 'OAUTH_AUTH_SUCCESS') {
          clearInterval(checkInterval);
          localStorage.removeItem('m365_admin_login_fails');
          const authTokens = e.data.tokens as MSTokens;
          setTokens(authTokens);
          localStorage.setItem('microsoft_tokens', JSON.stringify(authTokens));
          setAuthStatus('success');
          window.removeEventListener('message', handleMsg);
        } else if (e.data?.type === 'OAUTH_AUTH_ERROR') {
          clearInterval(checkInterval);
          setApiError(`Authentication failed: ${e.data.error}`);
          setAuthStatus('error');
          window.removeEventListener('message', handleMsg);
        }
      };

      window.addEventListener('message', handleMsg);
      
      checkInterval = setInterval(() => {
         try {
           if (authWindow?.closed) {
              clearInterval(checkInterval);
              window.removeEventListener('message', handleMsg);
              
              // Check if already succeeded (status updated)

             // We use a small timeout to let messages process
             setTimeout(() => {
                setAuthStatus(prev => {
                   if (prev === 'connecting') {
                      failCount++;
                      localStorage.setItem('m365_admin_login_fails', failCount.toString());
                      if (failCount >= 3) {
                         localStorage.setItem('m365_admin_login_timeout', (Date.now() + 60000).toString());
                         setApiError(`Authentication failed or canceled 3 times. Please wait 60 seconds before trying again.`);
                      } else {
                         setApiError(`Login window was closed before completing. (${failCount}/3)`);
                      }
                      return 'idle';
                   }
                   return prev;
                });
             }, 500);
          }
        } catch (e) {}
      }, 1000);
    } catch (err: any) {
      setApiError(`OAuth Initialization failed: ${err.message}`);
      setAuthStatus('idle');
    }
  };

  // 9. Signout Connection
  const handleSignout = () => {
    setTokens(null);
    setSaveConfig(null);
    localStorage.removeItem('microsoft_tokens');
    setSelectedTeamId('');
    setSelectedChannelId('');
    setSelectedFileId('');
    setExcelFiles([]);
    setProfile(null);
    setAuthStatus('idle');
  };

  if (isConfigLoading) {
    return (
      <div className="bg-white border rounded-2xl p-8 flex flex-col items-center justify-center min-h-[220px]">
        <RefreshCw size={24} className="text-blue-600 animate-spin" />
        <span className="text-xs text-gray-500 font-medium mt-3">Connecting to backend configs...</span>
      </div>
    );
  }

  // 10. Walkthrough If Credentials not configured in App Env Settings
  if (!hasCreds) {
    return (
      <div className="bg-white rounded-xl border border-red-200 shadow-sm p-6 space-y-5" id="credentials-banner">
        <div className="flex items-start gap-3">
          <div className="p-2.5 bg-red-50 text-red-600 rounded-xl shrink-0 border border-red-100">
            <Lock size={20} />
          </div>
          <div>
            <h2 className="font-bold text-slate-800 text-base">Microsoft Connection Pending Setup</h2>
            <p className="text-xs text-slate-500 mt-1">
              Configure your server environment secrets before activating real-time M365 document synchronization.
            </p>
          </div>
        </div>

        <div className="text-xs text-slate-650 space-y-3 leading-relaxed bg-slate-50 p-5 rounded-xl border border-slate-200">
          <p className="font-bold text-slate-800">👉 How to register your Teams app inside Azure Portal:</p>
          <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 font-semibold list-inside">
            <li>Go to the <span className="text-slate-800">Azure Active Directory</span> portal.</li>
            <li>Register as a <span className="text-slate-800">Multitenant Web App</span>.</li>
            <li>Configure Callback Redirect URL: <code className="bg-white px-2 py-0.5 rounded border border-slate-250 select-all font-mono font-bold text-blue-650">{`${appUrl}/auth/callback`}</code></li>
            <li>Add environment configuration variables <code className="bg-white px-1.5 py-0.5 rounded border border-slate-200 text-slate-700 font-mono">MICROSOFT_CLIENT_ID</code> and <code className="bg-white px-1.5 py-0.5 rounded border border-slate-200 text-slate-700 font-mono">MICROSOFT_CLIENT_SECRET</code> to your server settings.</li>
          </ol>
          <p className="text-slate-550 text-[11px] mt-2.5 italic">
            * Note: While settings are pending, the offline form engine remains active. You can build layouts and download compiled data manually as high-fidelity CSV spreadsheets below!
          </p>
        </div>

        <div className="flex flex-col sm:flex-row gap-3 pt-3 border-t border-slate-200">
          <a
            href="#workaround-documentation"
            className="flex-1 text-center py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-xs font-bold border border-slate-250 hover:shadow-xs transition-all active:scale-95"
          >
            Read Manual Setup Guide
          </a>
          <button 
            onClick={fetchConfig}
            className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow-sm hover:shadow active:scale-95 transition-all cursor-pointer"
          >
            <RefreshCw size={14} />
            Scan Environment Secrets
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6" id="microsoft-connector-component">
      {/* 11. Header Panel */}
      <div className="flex items-center justify-between border-b border-slate-200 pb-4">
        <div>
          <h2 className="font-bold text-slate-800 text-base">2. Connect Microsoft Teams</h2>
          <p className="text-xs text-slate-500 mt-0.5">Automated cloud storage inside shared Teams assets</p>
        </div>
      </div>

      {apiError && (
        <div className="p-3.5 bg-red-50 border border-red-100 text-red-800 rounded-xl flex items-start gap-3">
          <AlertCircle size={16} className="text-red-600 shrink-0 mt-0.5" />
          <p className="text-xs font-medium leading-relaxed">{apiError}</p>
        </div>
      )}

      {/* 12. OAuth State: Connect Button */}
      {!tokens ? (
        <div className="text-center py-10 bg-slate-50 border border-dashed border-slate-200 rounded-xl space-y-4">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-blue-50 text-blue-600 border border-blue-100">
            <Layers size={22} className="animate-pulse" />
          </div>
          <div className="max-w-xs mx-auto">
            <h4 className="font-bold text-slate-800 text-sm">Teams Integration Sync Offline</h4>
            <p className="text-xs text-slate-500 mt-1 leading-normal">Authenticate to map your dynamic custom form rows directly with Teams group worksheets.</p>
          </div>

          <button 
            onClick={handleConnectM365}
            disabled={authStatus === 'authorizing'}
            className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white rounded-xl text-xs font-bold shadow-sm hover:shadow transition-all flex items-center justify-center gap-2 mx-auto cursor-pointer active:scale-95"
            id="m365-sign-in-btn"
          >
            {authStatus === 'authorizing' ? (
              <>
                <RefreshCw size={14} className="animate-spin" />
                Validating connection...
              </>
            ) : (
              <>
                Sign In with Microsoft 365
              </>
            )}
          </button>
        </div>
      ) : (
        <div className="space-y-6">
          {/* Active login details */}
          {profile && (
            <div className="flex items-center gap-3 bg-emerald-50/50 p-3.5 rounded-xl border border-emerald-150">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse shrink-0"></span>
              <div className="flex-1 min-w-0">
                <span className="text-xs text-emerald-800 font-bold block">Connected Microsoft Identity:</span>
                <span className="text-[11px] text-emerald-750 font-semibold block truncate mt-0.5">
                  {profile.displayName} ({profile.mail || profile.userPrincipalName})
                </span>
              </div>
              <div className="text-[10px] bg-emerald-600 text-white px-2.5 py-0.5 rounded-full font-bold">ONLINE</div>
            </div>
          )}

          {/* 13. Dynamic selector grids */}
          {(!saveConfig || formConfig.settings?.isMappingLocked === false) && (
            <>
              {/* Common Team and Channel Selectors */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
                {/* Team select */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700 block">1. Choose M365 Workspace (Team / Group)</label>
                  <select
                    value={selectedTeamId}
                    onChange={(e) => handleTeamChange(e.target.value)}
                    disabled={isLoading}
                    className="w-full text-xs px-3.5 py-2.5 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 font-semibold text-slate-800 focus:border-blue-500 transition-all shadow-xs"
                    id="team-selector"
                  >
                    <option value="">Select Team Workspace</option>
                    {teams.map(team => (
                      <option key={team.id} value={team.id}>{team.displayName}</option>
                    ))}
                  </select>
                </div>

                {/* Channel select */}
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700 block">2. Choose Upload Folder (Channel)</label>
                  <select
                    value={selectedChannelId}
                    onChange={(e) => handleChannelChange(e.target.value)}
                    disabled={isLoading || !selectedTeamId}
                    className="w-full text-xs px-3.5 py-2.5 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 font-semibold text-slate-800 focus:border-blue-500 transition-all shadow-xs disabled:opacity-50"
                    id="channel-selector"
                  >
                    <option value="">Select Channel</option>
                    {channels.map(chan => (
                      <option key={chan.id} value={chan.id}>{chan.displayName}</option>
                    ))}
                  </select>
                </div>
              </div>

              {selectedChannelId && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start animate-fadeIn">
                  
                  {/* Left Column: Option A (Auto Setup) */}
                  <div className="space-y-4 bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                    <div className="flex items-center gap-2 mb-2">
                      <span className="p-1 px-2.5 bg-emerald-50 text-emerald-700 rounded border border-emerald-200 font-bold text-[10px]">AUTO SETUP</span>
                      <span className="text-sm font-bold text-slate-800 flex items-center gap-1.5">
                        <Sparkles size={14} className="text-yellow-500 fill-yellow-500" />
                        Generate New Context Automatically
                      </span>
                    </div>

                    {formConfig.fields.some(f => f.type === 'file') && (
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                        <label className="text-[11px] font-bold text-slate-800 block mb-1">Attachment Folder Path</label>
                        <p className="text-[10px] text-slate-500 mb-2">Select a directory or create a new one for file uploads in this channel.</p>
                        
                        <div className="space-y-3">
                          <select
                            value={uploadFolderPath}
                            onChange={(e) => setUploadFolderPath(e.target.value)}
                            disabled={isLoading}
                            className="w-full text-xs px-3 py-2 bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono shadow-xs"
                          >
                            <option value="">-- Choose Folder --</option>
                            <option value="Submissions_Attachments">Submissions_Attachments (Default)</option>
                            {attachmentFolders.map(folder => (
                              <option key={folder.id} value={folder.displayName || folder.name}>
                                📁 {folder.displayName || folder.name}
                              </option>
                            ))}
                          </select>

                          <div className="flex gap-2 items-center border-t border-slate-100 pt-3">
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Or Create New:</span>
                            <div className="flex-1 flex gap-2">
                              <input 
                                type="text" 
                                value={newFolderName}
                                onChange={(e) => setNewFolderName(e.target.value)}
                                placeholder="New Folder Name"
                                disabled={isCreatingFolder}
                                className="flex-1 text-xs px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:border-blue-500 shadow-xs"
                              />
                              <button 
                                type="button"
                                onClick={handleCreateNewFolder}
                                disabled={isCreatingFolder || !newFolderName.trim()}
                                className="bg-slate-200 hover:bg-slate-300 text-slate-800 disabled:opacity-50 rounded-lg text-[10px] font-bold px-3 py-1.5 flex items-center shadow-sm cursor-pointer transition-colors"
                              >
                                {isCreatingFolder ? '...' : 'Create Folder'}
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                    <div className="bg-white p-4 rounded-xl border border-slate-200">
                      <p className="text-[11px] text-slate-500 leading-relaxed font-medium mb-3">
                        Initialize a pristine live workspace spreadsheet inside the <code>{selectedChannelName}</code> workspace instantly. We will auto-create dynamic columns for you.
                      </p>

                      <div className="flex gap-2">
                        <button 
                          type="button"
                          onClick={handleAutoSetup}
                          disabled={isCreatingExcel}
                          className="w-full bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-lg text-xs font-bold px-4 py-2.5 flex items-center justify-center gap-1.5 shadow-sm hover:shadow active:scale-95 cursor-pointer transition-all"
                        >
                          {isCreatingExcel ? (
                            <><RefreshCw size={12} className="animate-spin" /> Auto Setting Up...</>
                          ) : (
                            <><Sparkles size={13} /> 1-Click Auto Setup (Recommended)</>
                          )}
                        </button>
                      </div>

                      <div className="relative my-4">
                        <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-slate-100"></div></div>
                        <div className="relative flex justify-center"><span className="bg-white px-2 text-[10px] uppercase text-slate-400 font-bold">OR CUSTOM NAME</span></div>
                      </div>

                      <form onSubmit={handleCreateNewExcelFile} className="flex gap-2">
                        <input 
                          type="text" 
                          value={newExcelFileName}
                          onChange={(e) => setNewExcelFileName(e.target.value)}
                          placeholder="Doc Name (e.g., Leads)"
                          disabled={isCreatingExcel}
                          className="flex-1 text-xs px-3 py-2 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-semibold text-slate-700 shadow-xs"
                        />
                        <button 
                          type="submit"
                          disabled={isCreatingExcel || !newExcelFileName.trim()}
                          className="bg-slate-800 hover:bg-slate-900 disabled:opacity-50 text-white rounded-lg text-xs font-bold px-4 py-2 flex items-center gap-1 shadow-sm hover:shadow active:scale-95 cursor-pointer shrink-0 transition-all"
                        >
                          <PlusCircle size={13} /> Create
                        </button>
                      </form>
                    </div>
                  </div>

                  {/* Right Column: Option B (Manual Setup) */}
                  <div className="space-y-4 bg-slate-50 border border-slate-200 rounded-xl p-5 shadow-sm">
                    <div className="flex items-center gap-2 mb-2">
                      <span className="p-1 px-2.5 bg-slate-200 text-slate-700 rounded border border-slate-300 font-bold text-[10px]">MANUAL SETUP</span>
                      <span className="text-sm font-bold text-slate-800">Map to Existing Document</span>
                    </div>

                    {formConfig.fields.some(f => f.type === 'file') && (
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                        <label className="text-[11px] font-bold text-slate-800 block mb-1">Attachment Folder Path</label>
                        <p className="text-[10px] text-slate-500 mb-2">Select a directory or create a new one for file uploads in this channel.</p>
                        
                        <div className="space-y-3">
                          <select
                            value={uploadFolderPath}
                            onChange={(e) => setUploadFolderPath(e.target.value)}
                            disabled={isLoading}
                            className="w-full text-xs px-3 py-2 bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono shadow-xs"
                          >
                            <option value="">-- Choose Folder --</option>
                            <option value="Submissions_Attachments">Submissions_Attachments (Default)</option>
                            {attachmentFolders.map(folder => (
                              <option key={folder.id} value={folder.displayName || folder.name}>
                                📁 {folder.displayName || folder.name}
                              </option>
                            ))}
                          </select>

                          <div className="flex gap-2 items-center border-t border-slate-100 pt-3">
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Or Create New:</span>
                            <div className="flex-1 flex gap-2">
                              <input 
                                type="text" 
                                value={newFolderName}
                                onChange={(e) => setNewFolderName(e.target.value)}
                                placeholder="New Folder Name"
                                disabled={isCreatingFolder}
                                className="flex-1 text-xs px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:border-blue-500 shadow-xs"
                              />
                              <button 
                                type="button"
                                onClick={handleCreateNewFolder}
                                disabled={isCreatingFolder || !newFolderName.trim()}
                                className="bg-slate-200 hover:bg-slate-300 text-slate-800 disabled:opacity-50 rounded-lg text-[10px] font-bold px-3 py-1.5 flex items-center shadow-sm cursor-pointer transition-colors"
                              >
                                {isCreatingFolder ? '...' : 'Create Folder'}
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                    <div className="space-y-3 bg-white p-4 rounded-xl border border-slate-200">
                      {/* Select file */}
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Excel Workbook</label>
                        <select
                          value={selectedFileId}
                          onChange={(e) => handleFileChange(e.target.value)}
                          disabled={isLoading}
                          className="w-full text-xs px-4 py-2 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-semibold text-slate-800 transition-all shadow-xs"
                        >
                          <option value="">Choose Worksheet File</option>
                          {excelFiles.map(file => (
                            <option key={file.id} value={file.id}>📄 {file.name}</option>
                          ))}
                        </select>
                        {excelFiles.length === 0 && !isLoading && (
                          <p className="text-[10px] text-amber-600 font-semibold mt-1">No Excel files found in the <code>{selectedChannelName}</code> folder yet.</p>
                        )}
                      </div>

                  {selectedFileId && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 animate-fadeIn">
                      {/* Select sheet */}
                      <div className="space-y-1.5 border border-slate-200 p-3 rounded-lg bg-slate-50">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Tab Worksheet</label>
                        <select
                          value={selectedSheetName}
                          onChange={(e) => setSelectedSheetName(e.target.value)}
                          className="w-full text-xs px-4 py-2 bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 mb-2"
                        >
                          <option value="">Select Sheet Tab</option>
                          {worksheets.map(sheet => (
                            <option key={sheet.id} value={sheet.name}>{sheet.name}</option>
                          ))}
                        </select>
                        <form onSubmit={handleCreateNewWorksheet} className="flex gap-2 border-t border-slate-200 pt-2">
                           <input 
                              type="text"
                              value={newWorksheetName}
                              onChange={e => setNewWorksheetName(e.target.value)}
                              placeholder="New tab name..."
                              className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500 focus:outline-none"
                              disabled={isCreatingWorksheet}
                           />
                           <button type="submit" disabled={isCreatingWorksheet || !newWorksheetName.trim()} className="bg-indigo-50 text-indigo-700 hover:bg-indigo-100 disabled:opacity-50 px-2 rounded text-[10px] font-bold border border-indigo-200 cursor-pointer text-nowrap">
                              {isCreatingWorksheet ? '...' : '+ New Tab'}
                           </button>
                        </form>
                      </div>

                      {/* Select table */}
                      <div className="space-y-1.5 border border-slate-200 p-3 rounded-lg bg-slate-50">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Workbook Table</label>
                        <select
                          value={selectedTableName}
                          onChange={(e) => setSelectedTableName(e.target.value)}
                          disabled={!selectedSheetName}
                          className="w-full text-xs px-4 py-2 bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 mb-2 disabled:opacity-50"
                        >
                          <option value="">Select Table Object</option>
                          {tables.map(table => (
                            <option key={table.id} value={table.name}>{table.name} ({table.displayName})</option>
                          ))}
                        </select>
                        {selectedSheetName && (
                          <form onSubmit={handleCreateNewTable} className="flex gap-2 border-t border-slate-200 pt-2">
                             <input 
                                type="text"
                                value={newTableName}
                                onChange={e => setNewTableName(e.target.value)}
                                placeholder="New table name..."
                                className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500 focus:outline-none"
                                disabled={isCreatingTable}
                             />
                             <button type="submit" disabled={isCreatingTable || !newTableName.trim()} className="bg-indigo-50 text-indigo-700 hover:bg-indigo-100 disabled:opacity-50 px-2 rounded text-[10px] font-bold border border-indigo-200 cursor-pointer text-nowrap">
                                {isCreatingTable ? '...' : '+ New Table'}
                             </button>
                          </form>
                        )}
                      </div>
                    </div>
                  )}

                  {selectedFileId && selectedTableName && (
                    <div className="flex flex-col gap-2 mt-4 pt-4 border-t border-slate-200">
                      <p className="text-[10px] uppercase font-bold text-slate-500 tracking-wider">Modify Existing Table Structure / Data</p>
                      <div className="flex flex-col sm:flex-row gap-2">
                        <button 
                           onClick={handleAddMissingColumns} 
                           disabled={isModifyingTable}
                           className="flex-1 text-[11px] border border-blue-200 bg-blue-50 hover:bg-blue-100 text-blue-700 py-2 rounded-lg font-bold disabled:opacity-50 transition-colors shadow-sm cursor-pointer"
                        >
                           {isModifyingTable ? 'Working...' : 'Add Missing Form Columns'}
                        </button>
                      </div>
                      <button
                        onClick={handleApplyLayoutSelections}
                        disabled={isModifyingTable}
                        className="w-full py-2.5 bg-slate-900 text-white rounded-lg text-xs font-bold hover:bg-slate-800 transition-all active:scale-95 cursor-pointer mt-2 shadow disabled:opacity-50"
                        id="save-mapping-btn"
                      >
                        📄 Lock Map & Link Table
                      </button>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
          </>
          )}

          {/* 14. Locked Activation Config Details summary */}
          {saveConfig && formConfig.settings?.isMappingLocked !== false && (
            <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-xl space-y-3 animate-fadeIn">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 text-emerald-800">
                  <CheckCircle2 size={16} className="text-emerald-600" />
                  <span className="text-xs font-bold uppercase tracking-wider">Sync Mapping Locked</span>
                </div>
                <button
                  onClick={() => {
                     if (window.confirm("Are you sure you want to unlock map syncing? This form will stop working and you will need to re-map it before it can receive responses again.")) {
                         setFormConfig({ ...formConfig, settings: { ...formConfig.settings, isMappingLocked: false } });
                     }
                  }}
                  className="px-2 py-1 bg-white text-rose-600 rounded text-[10px] font-bold border border-rose-200 hover:bg-rose-50 shadow-xs transition-colors cursor-pointer"
                >
                  Unlock Sync Mapping
                </button>
              </div>

              {publishedUrl && (
                 <>
                   <div className="bg-white border border-emerald-100 rounded-lg p-3 my-2 text-xs flex justify-between items-center">
                      <div>
                          <span className="text-[10px] uppercase font-bold text-slate-500 block">Live Form Link (Share with users)</span>
                          <a href={publishedUrl} target="_blank" rel="noopener noreferrer" className="text-blue-600 font-medium hover:underline flex items-center gap-1 mt-0.5">
                             {publishedUrl} <ExternalLink size={12} />
                          </a>
                      </div>
                      <button onClick={() => { navigator.clipboard.writeText(publishedUrl); alert('Link Copied!'); }} className="text-[10px] font-bold border border-slate-200 rounded px-2 py-1 bg-slate-50 hover:bg-slate-100 text-slate-600 flex items-center gap-1 cursor-pointer">
                          <Copy size={12} /> Copy
                      </button>
                   </div>
                   <div className="bg-white border border-blue-100 rounded-lg p-3 my-2 text-xs flex justify-between items-center">
                      <div>
                          <span className="text-[10px] uppercase font-bold text-blue-500 block">Monitor Responses (Admin Link)</span>
                          <a href={publishedUrl.replace('/form/', '/responses/')} target="_blank" rel="noopener noreferrer" className="text-blue-600 font-medium hover:underline flex items-center gap-1 mt-0.5">
                             {publishedUrl.replace('/form/', '/responses/')} <ExternalLink size={12} />
                          </a>
                      </div>
                      <button onClick={() => { navigator.clipboard.writeText(publishedUrl.replace('/form/', '/responses/')); alert('Link Copied!'); }} className="text-[10px] font-bold border border-slate-200 rounded px-2 py-1 bg-slate-50 hover:bg-slate-100 text-slate-600 flex items-center gap-1 cursor-pointer">
                          <Copy size={12} /> Copy
                      </button>
                   </div>
                 </>
              )}

              <div className="grid grid-cols-2 gap-x-4 gap-y-2.5 text-[11px] text-emerald-900 border-t border-emerald-150 pt-2.5">
                <div>
                  <span className="text-[9px] text-emerald-600 block leading-none font-bold uppercase">Team Workspace</span>
                  <span className="font-bold">{saveConfig.groupName}</span>
                </div>
                <div>
                  <span className="text-[9px] text-emerald-600 block leading-none font-bold uppercase">Teams Channel</span>
                  <span className="font-bold">{saveConfig.channelName}</span>
                </div>
                <div className="mt-1">
                  <span className="text-[9px] text-emerald-600 block leading-none font-bold uppercase">Excel File Target</span>
                  <span className="font-bold truncate block">{saveConfig.fileName}</span>
                </div>
                <div className="mt-1">
                  <span className="text-[9px] text-emerald-600 block leading-none font-bold uppercase">Excel Table Target</span>
                  <span className="font-bold">{saveConfig.tableName}</span>
                </div>
                {saveConfig.uploadFolderPath && (
                   <div className="mt-1 col-span-2">
                     <span className="text-[9px] text-emerald-600 block leading-none font-bold uppercase">Attachment Target Folder</span>
                     <span className="font-bold">{saveConfig.uploadFolderPath}</span>
                   </div>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* 15. Form Settings & Decoration (Global) */}
      <div className="pt-6 mt-6 border-t border-slate-200">
         <h3 className="text-[13px] font-bold text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-1.5">
            <Settings size={16} /> Advanced Form Settings
         </h3>
         
         <div className="bg-slate-50 border border-slate-200 rounded-xl p-5 space-y-4 shadow-sm">
            
            {/* Login Required */}
            <label className="flex items-start gap-3 cursor-pointer group mb-2">
              <input type="checkbox" checked={formConfig.settings?.requireMicrosoftLogin || false} onChange={(e) => {
                 let newVal = e.target.checked;
                 const newSettings = { ...formConfig.settings, requireMicrosoftLogin: newVal };
                 if (!newVal) {
                    newSettings.collectEmails = false;
                 }
                 setFormConfig({...formConfig, settings: newSettings});
              }} className="mt-0.5" />
              <div>
                 <span className="text-xs font-bold text-slate-800 block group-hover:text-blue-600">Require User Login Form Filling</span>
                 <span className="text-[10px] text-slate-500">Only people signed into a Microsoft 365 Account can submit the form.</span>
              </div>
            </label>

            {/* Collect Emails */}
            <label className="flex items-start gap-3 cursor-pointer group mb-2">
              <input type="checkbox" checked={!!formConfig.settings?.collectEmails} onChange={(e) => {
                 let newVal = e.target.checked;
                 const newSettings = { ...formConfig.settings, collectEmails: newVal };
                 if (newVal) {
                    newSettings.requireMicrosoftLogin = true;
                 }
                 setFormConfig({...formConfig, settings: newSettings});
              }} className="mt-0.5" />
              <div>
                 <span className="text-xs font-bold text-slate-800 block group-hover:text-blue-600">Collect Email</span>
                 <span className="text-[10px] text-slate-500">Automatically collects visitor email if the sheet has an 'Email' column. (Automatically enforces User Login)</span>
              </div>
            </label>

            {/* Allowed Domain filtering */}
            {formConfig.settings?.requireMicrosoftLogin && (
              <div className="pl-6 mb-2">
                 <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Restrict to Domain(s)</label>
                 <input type="text" placeholder="e.g. yourcompany.com" value={formConfig.settings?.allowedDomains || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, allowedDomains: e.target.value }})} className="w-full sm:w-1/2 p-2 text-[11px] rounded bg-white border border-slate-300 focus:border-blue-500 outline-none" />
                 <span className="text-[9px] text-slate-400 block mt-1">Leave blank to allow any Microsoft account. Seperate with commas for multiple domains.</span>
              </div>
            )}

            {/* Prevent Empty Submissions */}
            <label className="flex items-start gap-3 cursor-pointer group mb-2">
              <input type="checkbox" checked={!!formConfig.settings?.preventEmptySubmissions} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, preventEmptySubmissions: e.target.checked }})} className="mt-0.5" />
              <div>
                 <span className="text-xs font-bold text-slate-800 block group-hover:text-blue-600">Prevent Empty Submissions</span>
                 <span className="text-[10px] text-slate-500">If checked, a form cannot be submitted if all fields are left completely blank.</span>
              </div>
            </label>

            {/* Configure Generate Submission Sequence */}
            <div className="pt-3 mb-2 border-t border-slate-200">
               <div className="flex justify-between items-center mb-1">
                 <label className="text-[10px] font-bold text-slate-700 uppercase">Submission ID Format</label>
                 <button 
                   onClick={async () => {
                     if (!formId) return;
                     if (!confirm('Are you sure you want to reset the Submission ID counter? This will start the numbering back from the Start # you provide.')) return;
                     try {
                        const res = await fetch(`/api/forms/${formId}/reset-counter`, { method: 'POST' });
                        if (res.ok) {
                           alert('Submission counter reset successfully!');
                        }
                     } catch(e) {
                        alert('Could not reset sequence');
                     }
                   }}
                   className="text-[10px] bg-amber-50 text-amber-700 hover:bg-amber-100 px-2 py-1 rounded font-semibold border border-amber-200 transition-colors"
                 >
                   Reset Counter
                 </button>
               </div>
               <div className="flex gap-2 items-center">
                 <input type="text" placeholder="Prefix (e.g. S-)" value={formConfig.settings?.submissionPrefix || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, submissionPrefix: e.target.value }})} className="w-1/3 p-2 text-[11px] rounded bg-white border border-slate-300 focus:border-blue-500 outline-none" />
                 <input type="number" placeholder="Start # (e.g. 1)" value={formConfig.settings?.submissionStartNumber || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, submissionStartNumber: parseInt(e.target.value, 10) || 1 }})} className="w-1/3 p-2 text-[11px] rounded bg-white border border-slate-300 focus:border-blue-500 outline-none" />
               </div>
               <span className="text-[9px] text-slate-400 block mt-1">Sets the format for Sequential Submission IDs. Will look like format Prefix00Number</span>
            </div>

            {/* Allow Multiple Submissions */}
            <label className="flex items-start gap-3 cursor-pointer group">
              <input type="checkbox" checked={formConfig.settings?.allowMultipleSubmissions !== false} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, allowMultipleSubmissions: e.target.checked }})} className="mt-0.5" />
              <div>
                 <span className="text-xs font-bold text-slate-800 block group-hover:text-blue-600">Allow Multiple Responses (Per User)</span>
                 <span className="text-[10px] text-slate-500">If unchecked, requires login and will restrict the user to a single form submission.</span>
              </div>
            </label>

            <div className="pt-3 border-t border-slate-200">
               <label className="text-xs font-bold text-slate-800 block mb-2">Daily Time Constraints (Local Time)</label>
               <span className="text-[10px] text-slate-500 block mb-3">If configured, the form will only be available during these daily time slots. You can set up to 3 slots.</span>
               
               <div className="space-y-2">
                 {[0, 1, 2].map((idx) => {
                   const tr = formConfig.settings?.dailyTimeRanges?.[idx] || { start: '', end: '' };
                   const hasVal = tr.start || tr.end;
                   return (
                      <div key={idx} className={`flex items-center gap-2 p-2 rounded border ${hasVal ? 'bg-blue-50/50 border-blue-200' : 'bg-white border-slate-200'}`}>
                         <span className="text-[10px] font-bold text-slate-400 w-12">Slot {idx + 1}</span>
                         <input type="time" value={tr.start} onChange={(e) => {
                             const ranges = [...(formConfig.settings?.dailyTimeRanges || [])];
                             if (!ranges[idx]) ranges[idx] = { start: '', end: '' };
                             ranges[idx].start = e.target.value;
                             setFormConfig({...formConfig, settings: {...formConfig.settings, dailyTimeRanges: ranges}});
                         }} className="text-xs bg-white border border-slate-300 rounded px-2 py-1 outline-none" />
                         <span className="text-[10px] text-slate-500">to</span>
                         <input type="time" value={tr.end} onChange={(e) => {
                             const ranges = [...(formConfig.settings?.dailyTimeRanges || [])];
                             if (!ranges[idx]) ranges[idx] = { start: '', end: '' };
                             ranges[idx].end = e.target.value;
                             setFormConfig({...formConfig, settings: {...formConfig.settings, dailyTimeRanges: ranges}});
                         }} className="text-xs bg-white border border-slate-300 rounded px-2 py-1 outline-none" />
                         {hasVal && (
                             <button onClick={() => {
                                const ranges = [...(formConfig.settings?.dailyTimeRanges || [])];
                                ranges[idx] = { start: '', end: '' };
                                setFormConfig({...formConfig, settings: {...formConfig.settings, dailyTimeRanges: ranges}});
                             }} className="text-rose-500 hover:text-rose-700 p-1"><XCircle size={14} /></button>
                         )}
                      </div>
                   );
                 })}
               </div>
            </div>

         </div>
      </div>

      {/* 16. Visual Branding / Decorations */}
      <div className="pt-6 mt-6 border-t border-slate-200 mb-6">
         <h3 className="text-[13px] font-bold text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-1.5">
            <Settings size={16} /> Theme & Branding Layout
         </h3>
         
         <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4">
                <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Company Logo</label>
                <div className="flex flex-col gap-2 mb-2">
                  <input type="file" accept="image/*" onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (!f) return;
                      const r = new FileReader();
                      r.onload = () => setFormConfig({...formConfig, settings: { ...formConfig.settings, logoUrl: r.result as string }});
                      r.readAsDataURL(f);
                  }} className="w-full text-[11px] file:mr-2 file:py-1 file:px-2 file:rounded file:border-0 file:text-[10px] file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                  <input type="text" placeholder="Or Image Link URL" value={formConfig.settings?.logoUrl || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, logoUrl: e.target.value }})} className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500 font-mono" />
                </div>
                {formConfig.settings?.logoUrl && (
                   <div className="mt-2 space-y-2">
                       <div className="relative inline-block">
                           <img src={secureUrl(formConfig.settings.logoUrl)} alt="Logo Preview" className="h-10 border border-slate-200 rounded" />
                           <button onClick={() => setFormConfig({...formConfig, settings: { ...formConfig.settings, logoUrl: '' }})} className="absolute -top-2 -right-2 bg-white rounded-full shadow text-rose-500"><XCircle size={14}/></button>
                       </div>
                       <div className="flex items-center gap-2">
                           <span className="text-[9px] text-slate-500">Size (px):</span>
                           <input type="number" value={formConfig.settings.logoSize || 64} onChange={e => setFormConfig({...formConfig, settings: {...formConfig.settings, logoSize: parseInt(e.target.value)}})} className="w-16 text-[10px] px-1 py-0.5 border border-slate-300 rounded" />
                       </div>
                       <div className="flex items-center gap-2">
                           <span className="text-[9px] text-slate-500">Align:</span>
                           <select value={formConfig.settings.logoAlignment || 'center'} onChange={e => setFormConfig({...formConfig, settings: {...formConfig.settings, logoAlignment: e.target.value as any}})} className="text-[10px] px-1 py-0.5 border border-slate-300 rounded">
                              <option value="left">Left</option>
                              <option value="center">Center</option>
                              <option value="right">Right</option>
                           </select>
                       </div>
                   </div>
                )}
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4">
                <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Header Background (Cover / Color)</label>
                <div className="flex flex-col gap-2 mb-2">
                  <input type="file" accept="image/*" onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (!f) return;
                      const r = new FileReader();
                      r.onload = () => setFormConfig({...formConfig, settings: { ...formConfig.settings, coverUrl: r.result as string }});
                      r.readAsDataURL(f);
                  }} className="w-full text-[11px] file:mr-2 file:py-1 file:px-2 file:rounded file:border-0 file:text-[10px] file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                  <input type="text" placeholder="Or Image Link URL" value={formConfig.settings?.coverUrl || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, coverUrl: e.target.value }})} className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500 font-mono" />
                </div>
                {formConfig.settings?.coverUrl && (
                   <div className="relative inline-block w-full">
                       <img src={secureUrl(formConfig.settings.coverUrl)} alt="Cover Preview" className="w-full h-16 object-cover border border-slate-200 rounded" />
                       <button onClick={() => setFormConfig({...formConfig, settings: { ...formConfig.settings, coverUrl: '' }})} className="absolute top-1 right-1 bg-white rounded-full shadow text-rose-500"><XCircle size={14}/></button>
                   </div>
                )}
                
                <div className="mt-3">
                   <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Title & Description Alignment</label>
                   <select value={formConfig.settings?.headerAlignment || 'left'} onChange={e => setFormConfig({...formConfig, settings: {...formConfig.settings, headerAlignment: e.target.value as any}})} className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500">
                      <option value="left">Left Aligned</option>
                      <option value="center">Center Aligned</option>
                      <option value="right">Right Aligned</option>
                   </select>
                </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 sm:col-span-2">
                <div className="flex flex-col sm:flex-row gap-6">
                    <div>
                        <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Accent Theme Color</label>
                        <div className="flex items-center gap-2">
                           <input type="color" value={formConfig.settings?.themeColor || '#2563eb'} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, themeColor: e.target.value }})} className="w-8 h-8 rounded border-none cursor-pointer p-0 block bg-transparent" />
                           <span className="text-[11px] font-mono font-medium text-slate-500">{formConfig.settings?.themeColor || '#2563eb'}</span>
                        </div>
                    </div>
                    <div className="flex-1">
                        <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Page Background Color</label>
                        <div className="flex items-center gap-2">
                           <input type="color" value={formConfig.settings?.backgroundColor || '#f8fafc'} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, backgroundColor: e.target.value }})} className="w-8 h-8 rounded border-none cursor-pointer p-0 block bg-transparent" />
                           <span className="text-[11px] font-mono font-medium text-slate-500">{formConfig.settings?.backgroundColor || '#f8fafc'}</span>
                        </div>
                    </div>
                    <div className="flex-1">
                        <label className="text-[10px] font-bold text-slate-700 uppercase block mb-1">Page Background Image URL (Optional)</label>
                        <input type="text" placeholder="https://..." value={formConfig.settings?.backgroundUrl || ''} onChange={(e) => setFormConfig({...formConfig, settings: { ...formConfig.settings, backgroundUrl: e.target.value }})} className="w-full text-[10px] px-2 py-1.5 border border-slate-300 rounded focus:border-blue-500 font-mono" />
                    </div>
                </div>
            </div>
         </div>
      </div>
    </div>
  );
}
