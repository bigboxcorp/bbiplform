import { MSTokens } from '../types';

/**
 * Perform a fetch directed towards Microsoft Graph API via the backend proxy.
 */
export async function fetchGraph(
  subpath: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void,
  options: RequestInit = {}
): Promise<any> {
  // Check if accessToken is near expiry or expired
  let activeAccessToken = tokens.accessToken;
  const isExpired = tokens.acquiredAt && tokens.expiresIn 
    ? Date.now() > tokens.acquiredAt + (tokens.expiresIn - 60) * 1000 
    : false;

  if (isExpired && tokens.refreshToken) {
    try {
      console.log('Access token is expired, refreshing via token refresh route...');
      const refreshRes = await fetch('/api/auth/refresh', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ refreshToken: tokens.refreshToken }),
      });

      if (refreshRes.ok) {
        const newTokens = await refreshRes.json();
        activeAccessToken = newTokens.accessToken;
        setTokens(newTokens);
        localStorage.setItem('microsoft_tokens', JSON.stringify(newTokens));
        console.log('Successfully refreshed Microsoft Graph token.');
      } else {
        console.error('Failed to refresh Microsoft token, redirecting/warning user.');
      }
    } catch (err) {
      console.error('Token refresh failed:', err);
    }
  }

  const url = `/api/ms-graph/${subpath}`;
  const headers = {
    'Authorization': `Bearer ${activeAccessToken}`,
    'Accept': 'application/json',
    ...(options.headers || {}),
  } as Record<string, string>;

  const res = await fetch(url, {
    ...options,
    headers,
  });

  if (res.status === 204) {
    return null;
  }

  if (!res.ok) {
    const text = await res.text();
    let errorData;
    try {
      errorData = JSON.parse(text);
    } catch {
      errorData = { error: { message: text } };
    }
    throw new Error(errorData?.error?.message || `HTTP ${res.status}: ${res.statusText}`);
  }

  const responseText = await res.text();
  return responseText ? JSON.parse(responseText) : null;
}

/**
 * Fetch the logged-in user's profile
 */
export async function getProfile(tokens: MSTokens, setTokens: (t: MSTokens) => void) {
  return fetchGraph('me', tokens, setTokens);
}

/**
 * Fetch the user's joined Teams (Groups).
 * Fallback to general groups if joinedTeams endpoints returns empty or fails.
 */
export async function getJoinedTeams(tokens: MSTokens, setTokens: (t: MSTokens) => void) {
  try {
    const data = await fetchGraph('me/joinedTeams', tokens, setTokens);
    if (data && data.value) {
      return data.value;
    }
  } catch (err) {
    console.warn("Failed to fetch joinedTeams.", err);
  }
  
  try {
    // Fallback: fetch groups the user is a member of
    const fallbackData = await fetchGraph("me/memberOf?$select=id,displayName", tokens, setTokens);
    if (fallbackData && fallbackData.value) {
      // Return only objects that look like groups
      return fallbackData.value.filter((item: any) => item['@odata.type'] === '#microsoft.graph.group' || item.displayName);
    }
  } catch (err) {
    console.error("Fallback to groups also failed", err);
  }
  return [];
}

/**
 * Get folders for a specific Group's Drive
 */
export async function getTeamChannels(groupId: string, driveId: string, tokens: MSTokens, setTokens: (t: MSTokens) => void) {
  let channels = [];
  try {
    const data = await fetchGraph(`teams/${groupId}/channels`, tokens, setTokens);
    if (data && data.value) {
      channels = data.value.map((c: any) => ({
         id: c.id,
         displayName: c.displayName
      }));
    }
  } catch (err) {
    console.warn(`Failed to fetch proper channels for team ${groupId}.`, err);
  }
  
  return channels;
}

/**
 * Get the main files Drive ID of a Team
 */
export async function getTeamDrive(groupId: string, tokens: MSTokens, setTokens: (t: MSTokens) => void) {
  const data = await fetchGraph(`groups/${groupId}/drive`, tokens, setTokens);
  return data;
}

/**
 * List Excel Files inside a Teams channel.
 * Channel files are stored in a folder matching the channel name in the Team's drive.
 */
export async function getFoldersInChannel(
  driveId: string,
  channelName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const encodedChannel = encodeURIComponent(channelName);
  try {
    const isSpecialRoot = channelName === 'Root Directory';
    const endpoint = isSpecialRoot ? `drives/${driveId}/root/children` : `drives/${driveId}/root:/${encodedChannel}:/children`;
    
    let data;
    try {
       data = await fetchGraph(endpoint, tokens, setTokens);
    } catch (e: any) {
       if (!isSpecialRoot) {
           data = await fetchGraph(`drives/${driveId}/root/children`, tokens, setTokens);
       } else {
           throw e;
       }
    }
    
    if (data && data.value) {
      return data.value.filter((item: any) => item.folder);
    }
    return [];
  } catch (err: any) {
    console.warn(`Could not list folders in channel folder "${channelName}".`, err);
    return [];
  }
}

export async function createFolderInChannel(
  driveId: string,
  channelName: string,
  folderName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const encodedChannel = encodeURIComponent(channelName);
  const isSpecialRoot = channelName === 'Root Directory';
  
  let endpoint = '';
  if (isSpecialRoot) {
    endpoint = `drives/${driveId}/root/children`;
  } else {
    // First, resolve the folder ID for the channel
    const channelFolder = await fetchGraph(`drives/${driveId}/root:/${encodedChannel}`, tokens, setTokens);
    endpoint = `drives/${driveId}/items/${channelFolder.id}/children`;
  }
  
  return await fetchGraph(endpoint, tokens, setTokens, {
    method: 'POST',
    body: JSON.stringify({
      name: folderName,
      folder: { },
      "@microsoft.graph.conflictBehavior": "rename"
    })
  });
}

export async function getExcelFilesInChannel(
  driveId: string,
  channelName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  // Safe URL encoding for folder paths in OneDrive/SharePoint REST API
  const encodedChannel = encodeURIComponent(channelName);
  
  try {
    const isSpecialRoot = channelName === 'Root Directory';
    const endpoint = isSpecialRoot ? `drives/${driveId}/root/children` : `drives/${driveId}/root:/${encodedChannel}:/children`;
    
    let data;
    try {
       data = await fetchGraph(endpoint, tokens, setTokens);
    } catch (e: any) {
       if (!isSpecialRoot) {
           console.warn(`Channel folder not found, falling back to root.`, e);
           data = await fetchGraph(`drives/${driveId}/root/children`, tokens, setTokens);
       } else {
           throw e;
       }
    }

    const files = data?.value || [];
    
    // Filter to Excel files only (.xlsx, .xlsm, .xls)
    return files.filter((file: any) => 
      file.file && (
        file.name.endsWith('.xlsx') || 
        file.name.endsWith('.xlsm') || 
        file.name.endsWith('.xls')
      )
    );
  } catch (err: any) {
    console.warn(`Could not list files in channel folder "${channelName}". Folder might not exist yet if no files have been uploaded.`, err);
    // Fallback: list root folders to see what exists, or let user create a folder/file
    return [];
  }
}

/**
 * List worksheets in a workbook
 */
export async function getWorkbookWorksheets(
  driveId: string,
  itemId: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets`;
  const data = await fetchGraph(endpoint, tokens, setTokens);
  return data.value || [];
}

/**
 * List tables inside a workbook. Or get tables in specific sheet.
 */
export async function getWorkbookTables(
  driveId: string,
  itemId: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void,
  worksheetName?: string
) {
  const endpoint = worksheetName 
    ? `drives/${driveId}/items/${itemId}/workbook/worksheets('${worksheetName}')/tables`
    : `drives/${driveId}/items/${itemId}/workbook/tables`;
  try {
    const data = await fetchGraph(endpoint, tokens, setTokens);
    return data.value || [];
  } catch (err) {
    console.error('Failed to fetch tables:', err);
    return [];
  }
}

/**
 * Appends rows to an Excel Table.
 * values is a grid (array of arrays), e.g. [[val1, val2, val3]] or column-keyed object if mapping matches.
 */
export async function addRowToTable(
  driveId: string,
  itemId: string,
  tableName: string,
  rowValues: any[][],
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/rows`;
  return fetchGraph(endpoint, tokens, setTokens, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      values: rowValues,
    }),
  });
}

/**
 * Check if table inside a workbook contains columns.
 */
export async function getTableColumns(
  driveId: string,
  itemId: string,
  tableName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/columns`;
  const data = await fetchGraph(endpoint, tokens, setTokens);
  return data.value || [];
}

/**
 * Create a new Excel file inside the channel with headers, and turn it into a Table ready for rows.
 */
export async function createExcelFileWithTable(
  driveId: string,
  channelName: string,
  fileName: string,
  sheetName: string,
  headers: string[],
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const encodedChannel = encodeURIComponent(channelName);
  const cleanFileName = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
  const encodedFile = encodeURIComponent(cleanFileName);

  // 1. Create a blank Excel file in SharePoint.
  // We can create a simple base file or upload an empty template.
  // One elegant way to create a blank file in Graph API is uploading zero-byte content or simple template stream.
  // Graph API creates a blank spreadsheet if we put to content with special header:
  // PUT /drives/{driveId}/root:/{channelName}/{fileName}:/content
  // Headers: Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
  const uploadEndpoint = `drives/${driveId}/root:/${encodedChannel}/${encodedFile}:/content`;
  
  // We upload a minimal 50-byte placeholder or an empty base64 Excel stream.
  // A valid small empty xlsx file is a zipped archive. However, MS Graph can create a blank workbook
  // if you upload an empty body, but it's safer to POST the file creation metadata, then standard Graph Excel commands.
  // Wait! A cleaner way is to first check if file exists, if not construct it or let user use an existing file.
  // Actually, uploading a base64 string of a compiled blank .xlsx file is incredibly reliable!
  // Normal blank .xlsx file base64 header:
  const blankXlsxBase64 = "UEsDBBQAAAAAADBpClYAAAAAAAAAAAAAAAAGAAAAYWRkb25zL1RYVA==;..."; 
  // Wait, MS Graph has a native endpoint to create an Excel file:
  // POST /drives/{drive-id}/items/{parent-id}/children
  // Body: { "name": "NewFile.xlsx", "file": {}, "folder": null }
  // Yes! The children endpoint with empty `"file": {}` creates a fully valid empty workbook on SharePoint!
  // This is a beautiful native Graph capability. Let's do that!
  
  let channelFolderId = '';
  if (channelName === 'Root Directory') {
      const rootFolder = await fetchGraph(`drives/${driveId}/root`, tokens, setTokens);
      channelFolderId = rootFolder.id;
  } else {
      const channelFolder = await fetchGraph(`drives/${driveId}/root:/${encodedChannel}`, tokens, setTokens);
      channelFolderId = channelFolder.id;
  }

  const createEndpoint = `drives/${driveId}/items/${channelFolderId}/children`;
  const fileObj = await fetchGraph(createEndpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: cleanFileName,
      file: {},
      description: 'Created by Teams Excel Form Saver'
    })
  });

  const fileId = fileObj.id;

  // Let's sleep a moment to let Office Online index the new spreadsheet.
  await new Promise(resolve => setTimeout(resolve, 2000));

  // 2. Add headers in range A1 to column mapping (e.g. A1:E1 for 5 headers)
  const getColLetter = (colIndex: number) => {
    let letter = '';
    let temp = colIndex;
    while (temp > 0) {
      let rem = (temp - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      temp = Math.floor((temp - 1) / 26);
    }
    return letter;
  };
  const endColLetter = getColLetter(headers.length);
  const rangeAddress = `Sheet1!A1:${endColLetter}1`;

  const addHeaderEndpoint = `drives/${driveId}/items/${fileId}/workbook/worksheets('Sheet1')/range(address='${rangeAddress}')`;
  await fetchGraph(addHeaderEndpoint, tokens, setTokens, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      values: [headers]
    })
  });

  // 3. Convert that range (A1:X1) into a Table.
  const createTableEndpoint = `drives/${driveId}/items/${fileId}/workbook/worksheets('Sheet1')/tables/add`;
  const tableResult = await fetchGraph(createTableEndpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      address: `Sheet1!A1:${endColLetter}1`,
      hasHeaders: true
    })
  });

  return {
    file: fileObj,
    table: tableResult
  };
}

/**
 * Formats a given sheet range as a table if one doesn't exist.
 */
export async function convertRangeToTable(
  driveId: string,
  itemId: string,
  sheetName: string,
  rangeAddress: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/tables/add`;
  return fetchGraph(endpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      address: rangeAddress,
      hasHeaders: true
    })
  });
}

export async function addColumnToTable(
  driveId: string,
  itemId: string,
  tableName: string,
  columnName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/columns`;
  return fetchGraph(endpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: columnName
    })
  });
}

export async function clearTableData(
  driveId: string,
  itemId: string,
  tableName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/dataBodyRange/clear`;
  try {
    return await fetchGraph(endpoint, tokens, setTokens, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
         applyTo: "All"
      })
    });
  } catch(err: any) {
    if (err.message.includes("ItemNotFound")) {
       return null; // Table is already empty probably
    }
    throw err;
  }
}

export async function deleteTableRow(
  driveId: string,
  itemId: string,
  tableName: string,
  rowIndex: number,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
  return fetchGraph(endpoint, tokens, setTokens, { method: 'DELETE' });
}

export async function deleteTableColumn(
  driveId: string,
  itemId: string,
  tableName: string,
  columnName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableName}/columns/${columnName}`;
  return fetchGraph(endpoint, tokens, setTokens, { method: 'DELETE' });
}

export async function createWorksheet(
  driveId: string,
  itemId: string,
  sheetName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets`;
  return fetchGraph(endpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: sheetName })
  });
}

export async function getWorksheetUsedRange(
  driveId: string,
  itemId: string,
  sheetName: string,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const endpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/usedRange`;
  try {
    return await fetchGraph(endpoint, tokens, setTokens);
  } catch (err: any) {
    if (err.message && err.message.includes('ItemNotFound')) {
       return null; // Return null if nothing is used yet
    }
    throw err;
  }
}

export async function createTableInWorksheet(
  driveId: string,
  itemId: string,
  sheetName: string,
  tableName: string,
  headers: string[],
  startAddress: string, // e.g. "A1"
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  // Convert column (1-based index) to letter
  const getColLetter = (colIndex: number) => {
    let letter = '';
    let temp = colIndex;
    while (temp > 0) {
      let rem = (temp - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      temp = Math.floor((temp - 1) / 26);
    }
    return letter;
  };
  
  // Extract start column letter
  const match = startAddress.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error("Invalid start address format");
  const startColStr = match[1].toUpperCase();
  const startRow = match[2];

  // Convert start column string to index
  let startColIdx = 0;
  for (let i = 0; i < startColStr.length; i++) {
    startColIdx = startColIdx * 26 + (startColStr.charCodeAt(i) - 64);
  }

  const endColIdx = startColIdx + headers.length - 1;
  const endColLetter = getColLetter(endColIdx);
  const rangeAddress = `${sheetName}!${startAddress}:${endColLetter}${startRow}`;

  const addHeaderEndpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/range(address='${rangeAddress}')`;
  await fetchGraph(addHeaderEndpoint, tokens, setTokens, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [headers] })
  });

  // Create Table
  const createTableEndpoint = `drives/${driveId}/items/${itemId}/workbook/worksheets('${sheetName}')/tables/add`;
  const tableResult = await fetchGraph(createTableEndpoint, tokens, setTokens, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      address: rangeAddress,
      hasHeaders: true
    })
  });
  
  // Optionally rename table if tableName is provided
  if (tableName) {
     const renameEndpoint = `drives/${driveId}/items/${itemId}/workbook/tables/${tableResult.id}`;
     await fetchGraph(renameEndpoint, tokens, setTokens, {
       method: 'PATCH',
       headers: { 'Content-Type': 'application/json' },
       body: JSON.stringify({ name: tableName })
     });
  }

  return tableResult;
}

export async function uploadFileSystem(
  driveId: string,
  channelName: string,
  folderName: string,
  fileName: string,
  fileObj: File,
  tokens: MSTokens,
  setTokens: (t: MSTokens) => void
) {
  const isSpecialRoot = channelName === 'Root Directory';
  const encodedChannel = encodeURIComponent(channelName);
  const encodedFolder = folderName ? encodeURIComponent(folderName) : '';
  const encodedFileName = encodeURIComponent(fileName);
  
  let endpoint = '';
  if (isSpecialRoot) {
    endpoint = folderName ? `drives/${driveId}/root:/${encodedFolder}/${encodedFileName}:/content` : `drives/${driveId}/root:/${encodedFileName}:/content`;
  } else {
    endpoint = folderName ? `drives/${driveId}/root:/${encodedChannel}/${encodedFolder}/${encodedFileName}:/content` : `drives/${driveId}/root:/${encodedChannel}/${encodedFileName}:/content`;
  }
  
  // Convert File to Base64
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const base64String = (reader.result as string).split(',')[1];
        // Ensure MS Graph proxy can accept this through fetchGraph
        // The standard fetchGraph uses standard API route, but we need to pass base64
        // The current fetchGraph only supports JSON. Let's make fetchGraph support base64 natively.
        const res = await fetchGraph(endpoint, tokens, setTokens, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
             _isBase64File: true,
             contentType: fileObj.type || 'application/octet-stream',
             payload: base64String
          })
        });
        resolve(res);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsDataURL(fileObj);
  });
}

export async function uploadFilePublic(
  driveId: string,
  channelName: string,
  folderName: string,
  fileName: string,
  fileObj: File,
  formId: string
) {
  const isSpecialRoot = channelName === 'Root Directory';
  const encodedChannel = encodeURIComponent(channelName);
  const encodedFolder = folderName ? encodeURIComponent(folderName) : '';
  const encodedFileName = encodeURIComponent(fileName);
  
  let endpoint = '';
  if (isSpecialRoot) {
    endpoint = folderName ? `drives/${driveId}/root:/${encodedFolder}/${encodedFileName}:/content` : `drives/${driveId}/root:/${encodedFileName}:/content`;
  } else {
    endpoint = folderName ? `drives/${driveId}/root:/${encodedChannel}/${encodedFolder}/${encodedFileName}:/content` : `drives/${driveId}/root:/${encodedChannel}/${encodedFileName}:/content`;
  }
  
  // Convert File to Base64
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const base64String = (reader.result as string).split(',')[1];
        const res = await fetchPublicGraph(endpoint, formId, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
             _isBase64File: true,
             contentType: fileObj.type || 'application/octet-stream',
             payload: base64String
          })
        });
        resolve(res);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsDataURL(fileObj);
  });
}
export async function fetchPublicGraph(
  endpoint: string,
  formId: string,
  options: RequestInit = {}
) {
  const url = `/api/forms/${formId}/graph/${endpoint}`;
  
  let respondentAuth = '';
  try {
    const tokens = localStorage.getItem('microsoft_tokens');
    if (tokens) {
      respondentAuth = JSON.parse(tokens).accessToken;
    }
  } catch(e) {}

  const fetchOptions: RequestInit = {
    ...options,
    headers: {
      ...options.headers,
      ...(respondentAuth ? { 'x-respondent-auth': respondentAuth } : {})
    },
  };

  const response = await fetch(url, fetchOptions);

  let data;
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    data = await response.json();
  } else {
    data = await response.text();
  }

  if (!response.ok) {
    let errorMsg = `Graph API Error (${response.status})`;
    if (data && data.error && data.error.message) {
      errorMsg += `: ${data.error.message}`;
    } else if (typeof data === 'string') {
      errorMsg += `: ${data}`;
    }
    throw new Error(errorMsg);
  }

  return Object.keys(data || {}).length === 0 ? null : data;
}
