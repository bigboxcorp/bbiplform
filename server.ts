import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import dotenv from 'dotenv';
import Database from 'better-sqlite3';
import { v4 as uuidv4 } from 'uuid';

// Load environment variables from .env file
dotenv.config();

// Initialize SQLite database
const db = new Database('data.db', { verbose: console.log });

// Safe migrations
try { db.exec(`ALTER TABLE forms ADD COLUMN submissionCounter INTEGER DEFAULT 0;`); } catch(e) {}

db.exec(`
  CREATE TABLE IF NOT EXISTS forms (
    id TEXT PRIMARY KEY,
    config TEXT,
    excelConfig TEXT,
    creatorTokens TEXT,
    creatorEmail TEXT,
    createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    submissionCounter INTEGER DEFAULT 0
  );
  CREATE TABLE IF NOT EXISTS edit_logs (
    id TEXT PRIMARY KEY,
    formId TEXT,
    action TEXT,
    details TEXT,
    createdAt DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS submissions_log (
    id TEXT PRIMARY KEY,
    formId TEXT,
    userEmail TEXT,
    submittedAt DATETIME DEFAULT CURRENT_TIMESTAMP
  );
`);

const DEFAULT_CLIENT_ID = 'c26c42ed-a704-4a2e-8bde-44a6727bb47b';

function getMicrosoftCredentials() {
  return {
    clientId: 'c26c42ed-a704-4a2e-8bde-44a6727bb47b',
    clientSecret: 'gFB8Q~vN5FBSFyKF39a2KkAzneoQ86cxGl6yPaQk',
    credentialsIssue: null
  };
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Middleware to parse JSON bodies
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ extended: true, limit: '50mb' }));

  // API Route: Configuration health check
  app.get('/api/config', (req, res) => {
    // Determine the host for callback redirection
    const host = req.get('host') || 'localhost:3000';
    const isLocalhost = host.includes('localhost') || host.includes('127.0.0.1');
    const protocol = isLocalhost ? 'http' : 'https';
    const calculatedUrl = `${protocol}://${host}`;

    const { clientId, clientSecret } = getMicrosoftCredentials();

    res.json({
      hasCredentials: !!(clientId && clientSecret),
      appUrl: process.env.APP_URL || calculatedUrl,
      microsoftClientId: clientId
    });
  });

  // API Route: Generate OAuth URL
  app.get('/api/auth/url', (req, res) => {
    const clientRedirectUri = req.query.redirect_uri as string;
    
    // Construct the redirect URL (callback)
    let redirectUri = clientRedirectUri;
    if (!redirectUri) {
      const host = req.get('host') || 'localhost:3000';
      const isLocalhost = host.includes('localhost') || host.includes('127.0.0.1');
      const protocol = isLocalhost ? 'http' : 'https';
      redirectUri = `${protocol}://${host}/auth/callback`;
    }

    const { clientId } = getMicrosoftCredentials();
    if (!clientId) {
      return res.status(400).json({ error: 'MICROSOFT_CLIENT_ID is not configured on the server.' });
    }

    // Scopes needed to access Teams channels, SharePoint directories, and Excel worksheets
    const scopes = [
      'offline_access',
      'openid',
      'profile',
      'User.Read',
      'Files.ReadWrite',
      'Files.ReadWrite.All',
      'Sites.ReadWrite.All',
      'Group.ReadWrite.All'
    ].join(' ');

    const params = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: scopes,
      state: redirectUri // Pass the redirect URI in state so callback knows who to negotiate token for
    });

    const authorizeUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
    res.json({ url: authorizeUrl });
  });

  // OAuth Callback Route
  app.get(['/auth/callback', '/auth/callback/'], async (req, res) => {
    const { code, state } = req.query;
    if (!code) {
      return res.status(400).send('No authorization code was returned from Microsoft.');
    }

    // Reconstruct redirectUri from state (preferred) or headers
    let redirectUri = '';
    if (state && typeof state === 'string' && state.startsWith('http')) {
      redirectUri = state;
    } else {
      const host = req.get('host') || 'localhost:3000';
      const isLocalhost = host.includes('localhost') || host.includes('127.0.0.1');
      const protocol = isLocalhost ? 'http' : 'https';
      redirectUri = `${protocol}://${host}/auth/callback`;
    }

    try {
      const { clientId, clientSecret } = getMicrosoftCredentials();

      if (!clientId || !clientSecret) {
        return res.send(`
          <html>
            <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; padding: 30px; text-align: center; background-color: #fef2f2;">
              <div style="max-width: 450px; margin: 40px auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #fca5a5;">
                <h3 style="color: #dc2626; margin-top: 0;">Server Configuration Error</h3>
                <p style="color: #7f1d1d; font-size: 14px; line-height: 1.5;">
                  Microsoft Client credentials are missing on the server. Please add 
                  <code>MICROSOFT_CLIENT_ID</code> and <code>MICROSOFT_CLIENT_SECRET</code> to your 
                  AI Studio environment secrets.
                </p>
                <script>
                  if (window.opener) {
                    window.opener.postMessage({ type: 'OAUTH_AUTH_ERROR', error: 'Missing client credentials on server' }, '*');
                  }
                </script>
              </div>
            </body>
          </html>
        `);
      }

      const tokenParams = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        code: code as string,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
      });

      const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenParams.toString()
      });

      const tokenData = await tokenRes.json() as any;

      if (!tokenRes.ok) {
        console.error('Token acquisition failed inside callback:', tokenData);
        return res.send(`
          <html>
            <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 30px; background-color: #fafafa;">
              <div style="max-width: 500px; margin: 40px auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #e5e7eb;">
                <h3 style="color: #dc2626; margin-top: 0;">Authentication Error</h3>
                <p style="color: #4b5563; font-size: 14px;">Failed to exchange authorization code for access tokens with Microsoft.</p>
                <div style="text-align: left; background: #f3f4f6; padding: 12px; border-radius: 6px; overflow-x: auto; font-family: monospace; font-size: 12px; margin-top: 15px;">
                  ${JSON.stringify(tokenData, null, 2)}
                </div>
                <script>
                  if (window.opener) {
                    window.opener.postMessage({ type: 'OAUTH_AUTH_ERROR', error: 'Token exchange failed' }, '*');
                  }
                </script>
              </div>
            </body>
          </html>
        `);
      }

      // Successful Auth Token Exchange!
      res.send(`
        <html>
          <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background-color: #f8fafc; margin: 0;">
            <div style="background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); text-align: center; max-width: 400px; border: 1px solid #e2e8f0;">
              <div style="display: inline-flex; align-items: center; justify-content: center; width: 60px; height: 60px; border-radius: 50%; background-color: #d1fae5; color: #059669; font-size: 32px; font-weight: bold; margin-bottom: 20px;">✓</div>
              <h3 style="color: #1e293b; margin: 0 0 10px 0; font-size: 20px;">Connected Successfully!</h3>
              <p style="color: #64748b; font-size: 14px; margin: 0 0 24px 0; line-height: 1.5;">Your Microsoft account has been connected with the Forms app. This window will close automatically.</p>
              <div style="font-size: 12px; color: #94a3b8;">Returning to app...</div>
            </div>
            <script>
              if (window.opener) {
                window.opener.postMessage({
                  type: 'OAUTH_AUTH_SUCCESS',
                  tokens: {
                    accessToken: '${tokenData.access_token}',
                    refreshToken: '${tokenData.refresh_token || ""}',
                    expiresIn: ${tokenData.expires_in || 3600},
                    acquiredAt: ${Date.now()}
                  }
                }, '*');
                setTimeout(() => window.close(), 1200);
              } else {
                window.location.href = '/';
              }
            </script>
          </body>
        </html>
      `);
    } catch (err: any) {
      console.error('Fatal callback error:', err);
      res.status(500).send(`Error processing OAuth callback: ${err.message}`);
    }
  });

  // API Route: Refresh tokens on behalf of client
  app.post('/api/auth/refresh', async (req, res) => {
    const { refreshToken } = req.body;
    if (!refreshToken) {
      return res.status(400).json({ error: 'Missing refreshToken in request body' });
    }

    try {
      const { clientId, clientSecret } = getMicrosoftCredentials();

      const tokenParams = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
      });

      const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenParams.toString()
      });

      const tokenData = await tokenRes.json() as any;

      if (!tokenRes.ok) {
        return res.status(tokenRes.status).json(tokenData);
      }

      res.json({
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || refreshToken,
        expiresIn: tokenData.expires_in || 3600,
        acquiredAt: Date.now()
      });
    } catch (err: any) {
      res.status(500).json({ error: `Refresh Token Error: ${err.message}` });
    }
  });

  // API Proxy: Generic pass-through endpoint to Microsoft Graph API
  // Resolves all CORS blocks by fetching client requests on the backend
  app.all('/api/ms-graph/*', async (req, res) => {
    const subpath = req.params[0] || '';
    const query = req.url.split('?')[1] || '';
    const graphUrl = `https://graph.microsoft.com/v1.0/${subpath}${query ? '?' + query : ''}`;

    const authHeader = req.headers['authorization'];
    if (!authHeader) {
      return res.status(401).json({ error: 'Missing authorization bearer token' });
    }

    const headers: Record<string, string> = {
      'Authorization': authHeader,
      'Accept': 'application/json',
    };

    if (req.headers['content-type']) {
      headers['Content-Type'] = req.headers['content-type'] as string;
    }

    try {
      const fetchOptions: RequestInit = {
        method: req.method,
        headers: headers,
      };

      if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(req.method) && req.body) {
        fetchOptions.body = JSON.stringify(req.body);
      }

      const graphRes = await fetch(graphUrl, fetchOptions);
      const graphStatus = graphRes.status;
      
      // Let special statuses pass through cleanly (e.g. 201 Created, 204 No Content, etc.)
      res.status(graphStatus);

      if (graphStatus === 204) {
        return res.end();
      }

      const contentType = graphRes.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        const bodyObj = await graphRes.json();
        res.json(bodyObj);
      } else {
        const txt = await graphRes.text();
        res.send(txt);
      }
    } catch (err: any) {
      console.error(`Microsoft Graph Proxy error:`, err);
      res.status(500).json({ error: `Graph Proxy Internal Error: ${err.message}` });
    }
  });


  // ======================================
  // Forms API (SQLite Powered)
  // ======================================
  app.get('/api/forms', (req, res) => {
    try {
      const email = req.query.email as string;
      let rows;
      if (email) {
         rows = db.prepare(`SELECT id, config, excelConfig, createdAt, updatedAt FROM forms WHERE creatorEmail = ? ORDER BY createdAt DESC`).all(email) as any[];
      } else {
         rows = db.prepare(`SELECT id, config, excelConfig, createdAt, updatedAt FROM forms ORDER BY createdAt DESC`).all() as any[];
      }
      const forms = rows.map(r => ({
        id: r.id,
        config: JSON.parse(r.config),
        excelConfig: JSON.parse(r.excelConfig),
        createdAt: r.createdAt,
        updatedAt: r.updatedAt
      }));
      res.json(forms);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/forms', (req, res) => {
    try {
      const formId = uuidv4();
      const { config, excelConfig, creatorTokens, creatorEmail } = req.body;
      
      // Handle the case where existing SQLite file lacks the new column (add it dynamically if needed, but db resets usually anyway)
      try {
         db.exec(`ALTER TABLE forms ADD COLUMN creatorEmail TEXT;`);
      } catch (e) { /* Column likely exists */ }

      const insert = db.prepare(`INSERT INTO forms (id, config, excelConfig, creatorTokens, creatorEmail) VALUES (?, ?, ?, ?, ?)`);
      insert.run(
        formId,
        JSON.stringify(config),
        JSON.stringify(excelConfig),
        JSON.stringify(creatorTokens),
        creatorEmail || null
      );

      res.status(201).json({ id: formId });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put('/api/forms/:id', (req, res) => {
    try {
      const formId = req.params.id;
      const { config, excelConfig, creatorTokens, creatorEmail } = req.body;

      try {
         db.exec(`ALTER TABLE forms ADD COLUMN creatorEmail TEXT;`);
      } catch (e) { /* Column likely exists */ }

      const update = db.prepare(`UPDATE forms SET config = ?, excelConfig = ?, creatorTokens = ?, creatorEmail = ?, updatedAt = CURRENT_TIMESTAMP WHERE id = ?`);
      update.run(
        JSON.stringify(config),
        JSON.stringify(excelConfig),
        JSON.stringify(creatorTokens),
        creatorEmail || null,
        formId
      );

      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get('/api/forms/:id', (req, res) => {
    try {
      const formId = req.params.id;
      const row = db.prepare(`SELECT * FROM forms WHERE id = ?`).get(formId) as any;
      
      if (!row) return res.status(404).json({ error: 'Form not found' });

      // Only send safe data to client
      res.json({
        id: row.id,
        config: JSON.parse(row.config),
        excelConfig: JSON.parse(row.excelConfig) // Need this for mapping, maybe? Wait no, maybe client needs some info, but omit creatorTokens!
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/forms/:id/submit', async (req, res) => {
    try {
      const formId = req.params.id;
      const row = db.prepare(`SELECT * FROM forms WHERE id = ?`).get(formId) as any;
      if (!row) return res.status(404).json({ error: 'Form not found' });

      const creatorTokens = JSON.parse(row.creatorTokens);
      const excelConfig = JSON.parse(row.excelConfig);
      const formConfig = JSON.parse(row.config);
      const submittedData = req.body;

      // Check if creatorTokens is expired (We should ideally refresh it if expired)
      // This is a minimal implementation, refresh logic could be called if the Graph API responds with 401.

      // Convert submittedData to an array based on excelConfig.columnsMapping
      // Build the row layout exactly as it will go to Excel (the columnsMapping maps generic __submission_id etc to Header names, but wait, we need to know the index)
      // Since we just need to proxy the graph API endpoint, wait...
      
      return res.status(200).json({ success: true, message: 'Will be proxy submitted by backend' }); // To be implemented, currently the UI is submitting it using graphProxy
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get('/api/forms/:id/check-submission', (req, res) => {
    try {
      const email = req.query.email as string;
      const formId = req.params.id;
      if (!email) return res.json({ hasSubmitted: false });
      
      const row = db.prepare(`SELECT id FROM submissions_log WHERE formId = ? AND userEmail = ?`).get(formId, email);
      res.json({ hasSubmitted: !!row });
    } catch(err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/forms/:id/generate-id', (req, res) => {
    try {
      const formId = req.params.id;
      
      db.prepare('BEGIN IMMEDIATE').run();
      const formRow = db.prepare(`SELECT config, submissionCounter FROM forms WHERE id = ?`).get(formId) as { config: string, submissionCounter: number | null };
      if (!formRow) {
          db.prepare('ROLLBACK').run();
          return res.status(404).json({ error: 'Form not found' });
      }
      
      const currentCounter = formRow.submissionCounter || 0;
      const nextCounter = currentCounter + 1;
      
      db.prepare(`UPDATE forms SET submissionCounter = ? WHERE id = ?`).run(nextCounter, formId);
      db.prepare('COMMIT').run();
      
      const config = JSON.parse(formRow.config);
      const prefix = config.settings.submissionPrefix || 'S-';
      const startNum = config.settings.submissionStartNumber !== undefined ? config.settings.submissionStartNumber : 1;
      
      const finalNum = startNum + nextCounter - 1;
      const finalStr = finalNum.toString().padStart(3, '0');
      const nextId = `${prefix}${finalStr}`;

      res.json({ nextId });
    } catch(err: any) {
      try { db.prepare('ROLLBACK').run(); } catch(e) {}
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/forms/:id/record-submission', (req, res) => {
    try {
      const formId = req.params.id;
      const { email } = req.body;
      db.prepare(`INSERT INTO submissions_log (id, formId, userEmail) VALUES (?, ?, ?)`).run(uuidv4(), formId, email || 'anonymous');
      res.json({ success: true });
    } catch(err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get('/api/logs', (req, res) => {
    try {
      const rows = db.prepare(`SELECT * FROM edit_logs ORDER BY createdAt DESC LIMIT 100`).all();
      res.json(rows);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post('/api/logs', (req, res) => {
    try {
      const { formId, action, details } = req.body;
      const id = uuidv4();
      db.prepare(`INSERT INTO edit_logs (id, formId, action, details) VALUES (?, ?, ?, ?)`).run(id, formId || 'global', action, details);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ======================================
  // Secure Form-Bound Graph Proxy
  // ======================================
  app.get('/api/forms/:id/responses', async (req, res) => {
    try {
      const formId = req.params.id;
      const row = db.prepare(`SELECT * FROM forms WHERE id = ?`).get(formId) as any;
      if (!row) return res.status(404).json({ error: 'Form not found' });

      const creatorTokens = JSON.parse(row.creatorTokens);
      const excelConfig = JSON.parse(row.excelConfig);
      const formConfig = JSON.parse(row.config);
      
      const driveId = excelConfig.driveId || (excelConfig.driveItemId.includes('!') ? excelConfig.driveItemId.split('!')[0] : null);
      if (!driveId) throw new Error("Could not determine driveId");

      const endpoint = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${excelConfig.driveItemId}/workbook/tables/${excelConfig.tableName}/rows`;
      const attemptGraphCall = async (token: string, url: string) => {
        return fetch(url, {
          headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json' }
        });
      };

      let response = await attemptGraphCall(creatorTokens.accessToken, endpoint);

      if (response.status === 401 && creatorTokens.refreshToken) {
        // Try refresh
        const { clientId, clientSecret } = getMicrosoftCredentials();
        const tokenParams = new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          refresh_token: creatorTokens.refreshToken,
          grant_type: 'refresh_token',
        });
        const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: tokenParams.toString()
        });
        if (tokenRes.ok) {
          const tokenData = await tokenRes.json() as any;
          creatorTokens.accessToken = tokenData.access_token;
          creatorTokens.refreshToken = tokenData.refresh_token || creatorTokens.refreshToken;
          creatorTokens.expiresIn = tokenData.expires_in || 3600;
          creatorTokens.acquiredAt = Date.now();
          db.prepare(`UPDATE forms SET creatorTokens = ? WHERE id = ?`).run(JSON.stringify(creatorTokens), formId);
          response = await attemptGraphCall(creatorTokens.accessToken, endpoint);
        }
      }

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to fetch from Graph (${response.status}): ${errText}`);
      }
      
      const payload = await response.json() as { value: { values: any[][] }[] };
      
      // we also need headers:
      const colsRes = await attemptGraphCall(creatorTokens.accessToken, `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${excelConfig.driveItemId}/workbook/tables/${excelConfig.tableName}/columns`);
      let colsPayload: any = { value: [] };
      if (colsRes.ok) {
         colsPayload = await colsRes.json();
      }

      res.json({
         title: formConfig.title,
         themeColor: formConfig.settings?.themeColor || '#2563eb',
         columns: colsPayload.value.map(c => c.name),
         data: payload.value.map(v => v.values[0])
      });
    } catch(err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.all('/api/forms/:id/graph/*', async (req, res) => {
    const formId = req.params.id;
    const subpath = req.params[0] || '';
    const query = req.url.split('?')[1] || '';
    const graphUrl = `https://graph.microsoft.com/v1.0/${subpath}${query ? '?' + query : ''}`;
    
    try {
      const row = db.prepare(`SELECT * FROM forms WHERE id = ?`).get(formId) as any;
      if (!row) return res.status(404).json({ error: 'Form not found' });
      
      const formConfig = JSON.parse(row.config);

      let creatorTokens = JSON.parse(row.creatorTokens);
      
      // Attempt Graph Call
      const attemptGraphCall = async (token: string) => {
        const fetchOptions: RequestInit = {
          method: req.method,
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json',
          }
        };

        if (req.headers['content-type']) {
          (fetchOptions.headers as Record<string, string>)['Content-Type'] = req.headers['content-type'] as string;
        }

        if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(req.method) && req.body) {
          if (req.body && req.body._isBase64File) {
            fetchOptions.body = Buffer.from(req.body.payload, 'base64');
            (fetchOptions.headers as Record<string, string>)['Content-Type'] = req.body.contentType || 'application/octet-stream';
          } else {
            fetchOptions.body = JSON.stringify(req.body);
          }
        }
        return fetch(graphUrl, fetchOptions);
      };

      let graphRes = await attemptGraphCall(creatorTokens.accessToken);
      
      // If 401, trying to refresh token
      if (graphRes.status === 401 && creatorTokens.refreshToken) {
        const { clientId, clientSecret } = getMicrosoftCredentials();
        const tokenParams = new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          refresh_token: creatorTokens.refreshToken,
          grant_type: 'refresh_token',
        });
        
        const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: tokenParams.toString()
        });
        
        if (tokenRes.ok) {
          const tokenData = await tokenRes.json() as any;
          creatorTokens = {
            accessToken: tokenData.access_token,
            refreshToken: tokenData.refresh_token || creatorTokens.refreshToken,
            expiresIn: tokenData.expires_in || 3600,
            acquiredAt: Date.now()
          };
          // Save back to DB
          db.prepare(`UPDATE forms SET creatorTokens = ? WHERE id = ?`).run(
            JSON.stringify(creatorTokens),
            formId
          );
          
          // Retry
          graphRes = await attemptGraphCall(creatorTokens.accessToken);
        }
      }

      const graphStatus = graphRes.status;
      res.status(graphStatus);
      if (graphStatus === 204) return res.end();

      const contentType = graphRes.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        const bodyObj = await graphRes.json();
        res.json(bodyObj);
      } else {
        const txt = await graphRes.text();
        res.send(txt);
      }
    } catch (err: any) {
      console.error(`Secure Graph Proxy Error:`, err);
      res.status(500).json({ error: `Secure Graph Proxy Error: ${err.message}` });
    }
  });

  // Serve static UI assets or run dev middleware
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    // Serve client production static files from dist/ folder
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Express application successfully configured and running on http://localhost:${PORT}`);
  });
}

startServer();
