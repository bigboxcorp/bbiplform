export interface FieldValidation {
  type: string;
  value1?: string | number;
  value2?: string | number;
  customErrorMessage?: string;
}

export interface LogicJump {
  value: string; // The selected option that triggers the jump
  action: 'goto_section' | 'submit';
  targetSectionId?: string; // If goto_section, which section? (ID of the section break)
}

export interface GridInputColumn {
  id: string;
  name: string;
  type: 'text' | 'dropdown';
  options?: string[]; // for dropdown
  allowOther?: boolean; // for dropdown other option
}

export interface FormField {
  id: string;
  label: string;
  type: 'short_text' | 'long_text' | 'number' | 'date' | 'time' | 'rating' | 'select' | 'radio' | 'checkbox' | 'file' | 'grid_radio' | 'grid_checkbox' | 'grid_input' | 'section_break';
  options?: string[]; // Only for select, radio, checkbox
  allowOther?: boolean; // For select, radio, checkbox
  gridRows?: string[]; // For grid types
  gridCols?: string[]; // For grid types (radio/checkbox)
  gridInputCols?: GridInputColumn[]; // For grid_input type
  fileOptions?: { maxAllowed: number; maxSizeMB: number; allowedTypes?: string[] }; // For file upload
  required: boolean;
  allowRemarks?: boolean;
  placeholder?: string;
  
  // Advanced validations
  minLength?: number;
  maxLength?: number;
  minValue?: number | string;
  maxValue?: number | string;
  pattern?: string; // Regex pattern
  validation?: FieldValidation;

  // Logic functionality
  logicJumps?: LogicJump[]; // e.g. [{ value: 'Yes', action: 'goto_section', targetSectionId: 'sec2' }]
  sectionEndAction?: 'next' | 'submit' | 'goto_section';
  sectionEndTarget?: string;
}

export interface TimeRange {
  start: string; // "HH:MM"
  end: string;
}

export interface FormSettings {
  requireMicrosoftLogin: boolean;
  collectEmails?: boolean;
  preventEmptySubmissions?: boolean;
  allowedDomains?: string; // Comma separated domains
  allowMultipleSubmissions?: boolean;
  themeColor?: string;
  backgroundColor?: string;
  logoUrl?: string;
  logoSize?: number;
  logoAlignment?: 'left' | 'center' | 'right';
  headerAlignment?: 'left' | 'center' | 'right';
  backgroundUrl?: string;
  coverUrl?: string;
  isMappingLocked?: boolean;
  dailyTimeRanges?: TimeRange[];
  submissionPrefix?: string;
  submissionStartNumber?: number;
  showProgressBar?: boolean;
  notificationEmails?: string;
  responsesViewLimit?: number;
  responsesViewOrder?: 'newest' | 'oldest';
}

export interface FormConfig {
  title: string;
  description: string;
  settings: FormSettings;
  fields: FormField[];
}

export interface SavedForm {
  id: string;
  config: FormConfig;
  excelConfig: ExcelSaveConfig;
  creatorTokens: MSTokens; // Server side only ideally, but we transport it
}

export interface MSTokens {
  accessToken: string;
  refreshToken?: string;
  expiresIn?: number;
  acquiredAt?: number;
}

export interface MSGroup {
  id: string;
  displayName: string;
  description?: string;
}

export interface MSChannel {
  id: string;
  displayName: string;
  description?: string;
}

export interface MSDrive {
  id: string;
  name: string;
  driveType: string;
}

export interface MSDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  folder?: any;
  file?: any;
}

export interface MSWorksheet {
  id: string;
  name: string;
  position: number;
}

export interface MSTable {
  id: string;
  name: string;
  displayName: string;
}

export interface ExcelSaveConfig {
  groupId: string;
  groupName: string;
  channelId: string;
  channelName: string;
  driveId: string;
  driveItemId: string;
  fileName: string;
  sheetName: string;
  tableName: string;
  columnsMapping: Record<string, string>; // Maps Field ID to Excel Column Name
  uploadFolderPath?: string; // Path inside the channel's root where files will be uploaded
}

export interface FormSubmission {
  id: string;
  submittedAt: string;
  data: Record<string, any>;
  status: 'pending' | 'success' | 'failed';
  error?: string;
}
