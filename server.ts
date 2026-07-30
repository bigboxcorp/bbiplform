import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import Database from "better-sqlite3";
import { v4 as uuidv4 } from "uuid";
import nodemailer from "nodemailer";

// Load environment variables from .env file
dotenv.config();

// Setup SMTP Transporter
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST || "smtp.office365.com",
  port: parseInt(process.env.SMTP_PORT || "587", 10),
  secure: process.env.SMTP_PORT === "465", // true for 465, false for other ports
  requireTLS: true, // Force TLS for Microsoft Office 365
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

// Verify SMTP connection
if (process.env.SMTP_USER) {
  transporter.verify(function (error, success) {
    if (error) {
      console.error("SMTP Connection Error:", error);
    } else {
      console.log("SMTP Server is ready to take our messages");
    }
  });
}

// Initialize SQLite database
const db = new Database("data.db", { verbose: console.log });

// Safe migrations
try {
  db.exec(`ALTER TABLE forms ADD COLUMN submissionCounter INTEGER DEFAULT 0;`);
} catch (e) {}

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
  CREATE TABLE IF NOT EXISTS qrcodes (
    id TEXT PRIMARY KEY,
    title TEXT,
    type TEXT,
    targetData TEXT,
    createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
  );
`);

const DEFAULT_CLIENT_ID = "c26c42ed-a704-4a2e-8bde-44a6727bb47b";

function getMicrosoftCredentials() {
  return {
    clientId: "c26c42ed-a704-4a2e-8bde-44a6727bb47b",
    clientSecret: "gFB8Q~vN5FBSFyKF39a2KkAzneoQ86cxGl6yPaQk",
    credentialsIssue: null,
  };
}

async function getAppOnlyToken() {
  const { clientId, clientSecret } = getMicrosoftCredentials();
  const tenantId = process.env.MICROSOFT_TENANT_ID;
  if (!tenantId) return null;

  try {
    const params = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });
    const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });
    if (res.ok) {
      const data = await res.json();
      return data.access_token;
    }
  } catch (err) {
    console.error("Failed to get app-only token:", err);
  }
  return null;
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Middleware to parse JSON bodies
  app.use(express.json({ limit: "100mb" }));
  app.use(express.urlencoded({ extended: true, limit: "100mb" }));

  // API Route: Configuration health check
  app.get("/api/config", (req, res) => {
    // Determine the host for callback redirection
    const host = req.get("host") || "localhost:3000";
    const isLocalhost =
      host.includes("localhost") || host.includes("127.0.0.1");
    const protocol = isLocalhost ? "http" : "https";
    const calculatedUrl = `${protocol}://${host}`;

    const { clientId, clientSecret } = getMicrosoftCredentials();

    res.json({
      hasCredentials: !!(clientId && clientSecret),
      appUrl: process.env.APP_URL || calculatedUrl,
      microsoftClientId: clientId,
    });
  });

  // API Route: Generate OAuth URL
  app.get("/api/auth/url", (req, res) => {
    const clientRedirectUri = req.query.redirect_uri as string;

    // Construct the redirect URL (callback)
    let redirectUri = clientRedirectUri;
    if (!redirectUri) {
      const host = req.get("host") || "localhost:3000";
      const isLocalhost =
        host.includes("localhost") || host.includes("127.0.0.1");
      const protocol = isLocalhost ? "http" : "https";
      redirectUri = `${protocol}://${host}/auth/callback`;
    }

    const { clientId } = getMicrosoftCredentials();
    if (!clientId) {
      return res
        .status(400)
        .json({
          error: "MICROSOFT_CLIENT_ID is not configured on the server.",
        });
    }

    // Scopes needed to access Teams channels, SharePoint directories, and Excel worksheets
    const scopes = [
      "offline_access",
      "openid",
      "profile",
      "email",
      "User.Read.All",
      "Team.ReadBasic.All",
      "Files.ReadWrite",
      "Files.ReadWrite.All",
      "Sites.ReadWrite.All",
      "Group.ReadWrite.All",
      "Mail.Send",
      "Mail.Send.Shared",
    ].join(" ");

    const params = new URLSearchParams({
      client_id: clientId,
      response_type: "code",
      redirect_uri: redirectUri,
      response_mode: "query",
      scope: scopes,
      state: redirectUri, // Pass the redirect URI in state so callback knows who to negotiate token for
    });

    const authorizeUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
    res.json({ url: authorizeUrl });
  });

  // OAuth Callback Route
  app.get(["/auth/callback", "/auth/callback/"], async (req, res) => {
    const { code, state } = req.query;
    if (!code) {
      return res
        .status(400)
        .send("No authorization code was returned from Microsoft.");
    }

    // Reconstruct redirectUri from state (preferred) or headers
    let redirectUri = "";
    if (state && typeof state === "string" && state.startsWith("http")) {
      redirectUri = state;
    } else {
      const host = req.get("host") || "localhost:3000";
      const isLocalhost =
        host.includes("localhost") || host.includes("127.0.0.1");
      const protocol = isLocalhost ? "http" : "https";
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
        grant_type: "authorization_code",
      });

      const tokenRes = await fetch(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: tokenParams.toString(),
        },
      );

      const tokenData = (await tokenRes.json()) as any;

      if (!tokenRes.ok) {
        console.error("Token acquisition failed inside callback:", tokenData);
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
      console.error("Fatal callback error:", err);
      res.status(500).send(`Error processing OAuth callback: ${err.message}`);
    }
  });

  // API Route: Refresh tokens on behalf of client
  app.post("/api/auth/refresh", async (req, res) => {
    const { refreshToken } = req.body;
    if (!refreshToken) {
      return res
        .status(400)
        .json({ error: "Missing refreshToken in request body" });
    }

    try {
      const { clientId, clientSecret } = getMicrosoftCredentials();

      const tokenParams = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
      });

      const tokenRes = await fetch(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: tokenParams.toString(),
        },
      );

      const tokenData = (await tokenRes.json()) as any;

      if (!tokenRes.ok) {
        return res.status(tokenRes.status).json(tokenData);
      }

      res.json({
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token || refreshToken,
        expiresIn: tokenData.expires_in || 3600,
        acquiredAt: Date.now(),
      });
    } catch (err: any) {
      res.status(500).json({ error: `Refresh Token Error: ${err.message}` });
    }
  });

  // API Proxy: Generic pass-through endpoint to Microsoft Graph API
  // Resolves all CORS blocks by fetching client requests on the backend
  app.all("/api/ms-graph/*", async (req, res) => {
    const subpath = req.params[0] || "";
    const query = req.url.split("?")[1] || "";
    const graphUrl = `https://graph.microsoft.com/v1.0/${subpath}${query ? "?" + query : ""}`;

    const authHeader = req.headers["authorization"];
    if (!authHeader) {
      return res
        .status(401)
        .json({ error: "Missing authorization bearer token" });
    }

    const headers: Record<string, string> = {
      Authorization: authHeader,
      Accept: "application/json",
    };

    if (req.headers["content-type"]) {
      // Strip out charset for Microsoft Graph compatibility
      const cType = req.headers["content-type"] as string;
      if (cType.includes("application/json")) {
        headers["Content-Type"] = "application/json";
      } else {
        headers["Content-Type"] = cType;
      }
    }

    try {
      const fetchOptions: RequestInit = {
        method: req.method,
        headers: headers,
      };

      if (["POST", "PUT", "PATCH", "DELETE"].includes(req.method)) {
        if (
          req.body &&
          (Object.keys(req.body).length > 0 ||
            req.headers["content-type"]?.includes("json"))
        ) {
          fetchOptions.body = JSON.stringify(req.body);
          headers["Content-Type"] = "application/json";
        }
      }

      // Explicitly lowercase content-type override for native fetch just in case
      if (headers["Content-Type"]) {
        headers["content-type"] = headers["Content-Type"];
        delete headers["Content-Type"];
      }

      console.log(`[GRAPH CALL] ${req.method} ${graphUrl}`);
      console.log(`[GRAPH HEADERS]`, headers);
      if (fetchOptions.body) {
        const bodyStr = fetchOptions.body.toString();
        require('fs').appendFileSync('logs.txt', `\n[AUTH GRAPH CALL BODY LENGTH] ${bodyStr.length}\n`);
        if (bodyStr.length < 2000) {
            require('fs').appendFileSync('logs.txt', `[AUTH GRAPH CALL BODY] ${bodyStr}\n`);
        }
      }

      const graphRes = await fetch(graphUrl, fetchOptions);
      const graphStatus = graphRes.status;

      // Let special statuses pass through cleanly (e.g. 201 Created, 204 No Content, etc.)
      res.status(graphStatus);

      if (graphStatus === 204) {
        return res.end();
      }

      const contentType = graphRes.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        const bodyObj = await graphRes.json();
        res.json(bodyObj);
      } else {
        const txt = await graphRes.text();
        res.send(txt);
      }
    } catch (err: any) {
      console.error(`Microsoft Graph Proxy error:`, err);
      res
        .status(500)
        .json({ error: `Graph Proxy Internal Error: ${err.message}` });
    }
  });

  // ======================================
  // Forms API (SQLite Powered)
  // ======================================
  app.get("/api/mapped-tables", (req, res) => {
    try {
      const rows = db
        .prepare(`SELECT id, excelConfig FROM forms`)
        .all() as any[];
      const mappedTables: Record<string, string> = {}; // composite_key -> formId
      for (const row of rows) {
        try {
          const excelConfig = JSON.parse(row.excelConfig);
          if (excelConfig && excelConfig.driveItemId && excelConfig.tableName) {
            mappedTables[
              `${excelConfig.driveItemId}_${excelConfig.tableName}`
            ] = row.id;
          }
        } catch (e) {}
      }
      res.json(mappedTables);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/forms", (req, res) => {
    try {
      const email = req.query.email as string;
      let rows;
      if (email) {
        rows = db
          .prepare(
            `SELECT id, config, excelConfig, createdAt, updatedAt FROM forms WHERE creatorEmail = ? ORDER BY createdAt DESC`,
          )
          .all(email) as any[];
      } else {
        rows = db
          .prepare(
            `SELECT id, config, excelConfig, createdAt, updatedAt FROM forms ORDER BY createdAt DESC`,
          )
          .all() as any[];
      }
      const forms = rows.map((r) => ({
        id: r.id,
        config: JSON.parse(r.config),
        excelConfig:
          r.excelConfig && r.excelConfig !== "{}" && r.excelConfig !== "null"
            ? JSON.parse(r.excelConfig)
            : null,
        createdAt: r.createdAt,
        updatedAt: r.updatedAt,
      }));
      res.json(forms);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms", (req, res) => {
    try {
      const formId = uuidv4();
      const { config, excelConfig, creatorTokens, creatorEmail } = req.body;

      // Handle the case where existing SQLite file lacks the new column (add it dynamically if needed, but db resets usually anyway)
      try {
        db.exec(`ALTER TABLE forms ADD COLUMN creatorEmail TEXT;`);
      } catch (e) {
        /* Column likely exists */
      }

      const insert = db.prepare(
        `INSERT INTO forms (id, config, excelConfig, creatorTokens, creatorEmail) VALUES (?, ?, ?, ?, ?)`,
      );
      insert.run(
        formId,
        JSON.stringify(config),
        JSON.stringify(excelConfig),
        JSON.stringify(creatorTokens),
        creatorEmail || null,
      );

      res.status(201).json({ id: formId });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms/:id/duplicate", (req, res) => {
    try {
      const formId = req.params.id;
      const { creatorEmail } = req.body;
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;

      if (!row) return res.status(404).json({ error: "Form not found" });

      const newId = uuidv4();
      let config = JSON.parse(row.config);

      // Clear mapping from duplicated form
      config.settings = { ...(config.settings || {}), isMappingLocked: false };

      const insert = db.prepare(
        `INSERT INTO forms (id, config, excelConfig, creatorTokens, creatorEmail) VALUES (?, ?, ?, ?, ?)`,
      );
      insert.run(
        newId,
        JSON.stringify(config),
        JSON.stringify(null), // Clear excel config
        row.creatorTokens, // Keep tokens if we want? Or maybe clear it? It's fine to keep them if it's the same user.
        creatorEmail || row.creatorEmail || null,
      );

      res.json({ success: true, newId });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });
  app.put("/api/forms/:id", (req, res) => {
    try {
      const formId = req.params.id;
      const { config, excelConfig, creatorTokens, creatorEmail } = req.body;

      try {
        db.exec(`ALTER TABLE forms ADD COLUMN creatorEmail TEXT;`);
      } catch (e) {
        /* Column likely exists */
      }

      const update = db.prepare(
        `UPDATE forms SET config = ?, excelConfig = ?, creatorTokens = ?, creatorEmail = ?, updatedAt = CURRENT_TIMESTAMP WHERE id = ?`,
      );
      update.run(
        JSON.stringify(config),
        JSON.stringify(excelConfig),
        JSON.stringify(creatorTokens),
        creatorEmail || null,
        formId,
      );

      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete("/api/forms/:id", (req, res) => {
    try {
      const formId = req.params.id;
      const deleteStmt = db.prepare(`DELETE FROM forms WHERE id = ?`);
      const result = deleteStmt.run(formId);

      if (result.changes > 0) {
        res.json({ success: true });
      } else {
        res.status(404).json({ error: "Form not found" });
      }
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/forms/:id", (req, res) => {
    try {
      const formId = req.params.id;
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;

      if (!row) return res.status(404).json({ error: "Form not found" });

      // Only send safe data to client
      res.json({
        id: row.id,
        config: JSON.parse(row.config),
        excelConfig:
          row.excelConfig &&
          row.excelConfig !== "{}" &&
          row.excelConfig !== "null"
            ? JSON.parse(row.excelConfig)
            : null,
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms/:id/submit", async (req, res) => {
    try {
      const formId = req.params.id;
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;
      if (!row) return res.status(404).json({ error: "Form not found" });

      const creatorTokens = JSON.parse(row.creatorTokens);
      const excelConfig = JSON.parse(row.excelConfig);
      const formConfig = JSON.parse(row.config);
      const submittedData = req.body;

      // Check if creatorTokens is expired (We should ideally refresh it if expired)
      // This is a minimal implementation, refresh logic could be called if the Graph API responds with 401.

      // Convert submittedData to an array based on excelConfig.columnsMapping
      // Build the row layout exactly as it will go to Excel (the columnsMapping maps generic __submission_id etc to Header names, but wait, we need to know the index)
      // Since we just need to proxy the graph API endpoint, wait...

      return res
        .status(200)
        .json({ success: true, message: "Will be proxy submitted by backend" }); // To be implemented, currently the UI is submitting it using graphProxy
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/forms/:id/check-submission", (req, res) => {
    try {
      const email = req.query.email as string;
      const formId = req.params.id;
      if (!email) return res.json({ hasSubmitted: false });

      const row = db
        .prepare(
          `SELECT id FROM submissions_log WHERE formId = ? AND LOWER(userEmail) = LOWER(?)`,
        )
        .get(formId, email);
      res.json({ hasSubmitted: !!row });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms/:id/generate-id", (req, res) => {
    try {
      const formId = req.params.id;

      db.prepare("BEGIN IMMEDIATE").run();
      const formRow = db
        .prepare(`SELECT config, submissionCounter FROM forms WHERE id = ?`)
        .get(formId) as { config: string; submissionCounter: number | null };
      if (!formRow) {
        db.prepare("ROLLBACK").run();
        return res.status(404).json({ error: "Form not found" });
      }

      const currentCounter = formRow.submissionCounter || 0;
      const nextCounter = currentCounter + 1;

      db.prepare(`UPDATE forms SET submissionCounter = ? WHERE id = ?`).run(
        nextCounter,
        formId,
      );
      db.prepare("COMMIT").run();

      const config = JSON.parse(formRow.config);
      const prefix = config.settings.submissionPrefix || "S-";
      const startNum =
        config.settings.submissionStartNumber !== undefined
          ? config.settings.submissionStartNumber
          : 1;

      const finalNum = startNum + nextCounter - 1;
      const finalStr = finalNum.toString().padStart(3, "0");
      const nextId = `${prefix}${finalStr}`;

      res.json({ nextId });
    } catch (err: any) {
      try {
        db.prepare("ROLLBACK").run();
      } catch (e) {}
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms/:id/reset-counter", (req, res) => {
    try {
      const formId = req.params.id;
      db.prepare(`UPDATE forms SET submissionCounter = 0 WHERE id = ?`).run(
        formId,
      );
      db.prepare(`DELETE FROM submissions_log WHERE formId = ?`).run(formId);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/forms/:id/record-submission", (req, res) => {
    try {
      const formId = req.params.id;
      const { email } = req.body;
      db.prepare(
        `INSERT INTO submissions_log (id, formId, userEmail) VALUES (?, ?, ?)`,
      ).run(uuidv4(), formId, email || "anonymous");
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/test-email", async (req, res) => {
    try {
      if (!process.env.SMTP_USER || !process.env.SMTP_PASS) {
        return res
          .status(400)
          .json({
            error:
              "SMTP_USER and SMTP_PASS are not configured in the .env file.",
          });
      }

      // Verify connection configuration
      await new Promise((resolve, reject) => {
        transporter.verify(function (error, success) {
          if (error) {
            reject(error);
          } else {
            resolve(success);
          }
        });
      });

      // Send a test email
      const mailOptions = {
        from: process.env.SMTP_FROM || process.env.SMTP_USER,
        to: process.env.SMTP_USER, // send to self
        subject: `Test Email from Form Builder`,
        html: `<p>This is a test email to verify your SMTP configuration is working.</p>`,
      };

      const info = await transporter.sendMail(mailOptions);
      res.json({
        success: true,
        message: "SMTP connection successful and test email sent!",
        info,
      });
    } catch (err: any) {
      console.error("Test email failed:", err);
      res.status(500).json({
        error: "Failed to connect to SMTP server or send email",
        details: err.message,
        code: err.code,
        command: err.command,
      });
    }
  });

  app.post("/api/forms/:id/notify", async (req, res) => {
    try {
      const formId = req.params.id;
      const { emails, formTitle, submissionId } = req.body;

      if (!emails || !emails.length) {
        return res.json({ success: false, message: "No emails provided" });
      }

      // Fetch the form to get creatorTokens
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;
      if (!row) {
        return res.status(404).json({ error: "Form not found" });
      }

      if (!row.creatorTokens) {
        return res
          .status(400)
          .json({
            error: "Form creator has not connected their Microsoft account",
          });
      }

      let creatorTokens = JSON.parse(row.creatorTokens);

      // We might need to refresh the token if it's expired
      if (creatorTokens.expiresAt && Date.now() > creatorTokens.expiresAt) {
        try {
          const tokenResponse = await fetch(
            `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
            {
              method: "POST",
              headers: { "Content-Type": "application/x-www-form-urlencoded" },
              body: new URLSearchParams({
                client_id: process.env.MICROSOFT_CLIENT_ID!,
                client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
                refresh_token: creatorTokens.refreshToken,
                grant_type: "refresh_token",
              }),
            },
          );

          const tokens = await tokenResponse.json();
          if (tokens.access_token) {
            creatorTokens = {
              accessToken: tokens.access_token,
              refreshToken: tokens.refresh_token || creatorTokens.refreshToken,
              expiresAt: Date.now() + tokens.expires_in * 1000,
            };

            db.prepare(`UPDATE forms SET creatorTokens = ? WHERE id = ?`).run(
              JSON.stringify(creatorTokens),
              formId,
            );
          }
        } catch (err) {
          console.error("Failed to refresh token for notification:", err);
          // Proceed anyway, maybe it will work
        }
      }

      const emailMessage: any = {
        message: {
          subject: `New Submission: ${formTitle}`,
          body: {
            contentType: "HTML",
            content: `<p>A new submission has been received for your form <b>${formTitle}</b>.</p><p><b>Submission ID:</b> ${submissionId}</p><p>Please check your connected Microsoft Excel spreadsheet to view the details.</p>`,
          },
          toRecipients: emails.map((email: string) => ({
            emailAddress: { address: email },
          })),
        },
        saveToSentItems: "true",
      };

      const sendResponse = await fetch(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${creatorTokens.accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(emailMessage),
        },
      );

      if (!sendResponse.ok) {
        const errorData = await sendResponse.json();
        console.error("Graph API sendMail error:", errorData);
        return res
          .status(sendResponse.status)
          .json({
            error: "Failed to send email via Microsoft Graph",
            details: errorData,
          });
      }

      res.json({
        success: true,
        message: "Email sent successfully via Graph API!",
      });
    } catch (err: any) {
      console.error("Failed to send notification email:", err);
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/logs", (req, res) => {
    try {
      const rows = db
        .prepare(`SELECT * FROM edit_logs ORDER BY createdAt DESC LIMIT 100`)
        .all();
      res.json(rows);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/logs", (req, res) => {
    try {
      const { formId, action, details } = req.body;
      const id = uuidv4();
      db.prepare(
        `INSERT INTO edit_logs (id, formId, action, details) VALUES (?, ?, ?, ?)`,
      ).run(id, formId || "global", action, details);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // ======================================
  // Secure Form-Bound Graph Proxy
  // ======================================
  app.get("/api/forms/:id/responses", async (req, res) => {
    try {
      const formId = req.params.id;
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;
      if (!row) return res.status(404).json({ error: "Form not found" });

      const creatorTokens = JSON.parse(row.creatorTokens);
      const excelConfig = JSON.parse(row.excelConfig);
      const formConfig = JSON.parse(row.config);

      const driveId =
        excelConfig.driveId ||
        (excelConfig.driveItemId.includes("!")
          ? excelConfig.driveItemId.split("!")[0]
          : null);
      if (!driveId) throw new Error("Could not determine driveId");

      const endpoint = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${excelConfig.driveItemId}/workbook/tables/${excelConfig.tableName}/rows`;
      const attemptGraphCall = async (token: string, url: string) => {
        return fetch(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
          },
        });
      };

      let response = await attemptGraphCall(
        creatorTokens.accessToken,
        endpoint,
      );

      if (response.status === 401 && creatorTokens.refreshToken) {
        // Try refresh
        const { clientId, clientSecret } = getMicrosoftCredentials();
        const tokenParams = new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          refresh_token: creatorTokens.refreshToken,
          grant_type: "refresh_token",
        });
        const tokenRes = await fetch(
          "https://login.microsoftonline.com/common/oauth2/v2.0/token",
          {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: tokenParams.toString(),
          },
        );
        if (tokenRes.ok) {
          const tokenData = (await tokenRes.json()) as any;
          creatorTokens.accessToken = tokenData.access_token;
          creatorTokens.refreshToken =
            tokenData.refresh_token || creatorTokens.refreshToken;
          creatorTokens.expiresIn = tokenData.expires_in || 3600;
          creatorTokens.acquiredAt = Date.now();
          db.prepare(`UPDATE forms SET creatorTokens = ? WHERE id = ?`).run(
            JSON.stringify(creatorTokens),
            formId,
          );
          response = await attemptGraphCall(
            creatorTokens.accessToken,
            endpoint,
          );
        }
      }

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(
          `Failed to fetch from Graph (${response.status}): ${errText}`,
        );
      }

      const payload = (await response.json()) as {
        value: { values: any[][] }[];
      };

      // we also need headers:
      const colsRes = await attemptGraphCall(
        creatorTokens.accessToken,
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${excelConfig.driveItemId}/workbook/tables/${excelConfig.tableName}/columns`,
      );
      let colsPayload: any = { value: [] };
      if (colsRes.ok) {
        colsPayload = await colsRes.json();
      }

      let dataList = payload.value.map((v: any) => v.values[0]);

      const limit = formConfig.settings?.responsesViewLimit;
      const order = formConfig.settings?.responsesViewOrder || 'newest';

      if (order === 'newest') {
        dataList.reverse();
      }

      if (limit && limit > 0) {
        dataList = dataList.slice(0, limit);
      }

      res.json({
        title: formConfig.title,
        themeColor: formConfig.settings?.themeColor || "#2563eb",
        columns: colsPayload.value.map((c: any) => c.name),
        data: dataList,
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.all("/api/forms/:id/graph/*", async (req, res) => {
    const formId = req.params.id;
    const subpath = req.params[0] || "";
    const query = req.url.split("?")[1] || "";
    const graphUrl = `https://graph.microsoft.com/v1.0/${subpath}${query ? "?" + query : ""}`;

    try {
      const row = db
        .prepare(`SELECT * FROM forms WHERE id = ?`)
        .get(formId) as any;
      if (!row) return res.status(404).json({ error: "Form not found" });

      const formConfig = JSON.parse(row.config);

      let creatorTokens = JSON.parse(row.creatorTokens);

      // Attempt Graph Call
      const attemptGraphCall = async (token: string) => {
        const fetchOptions: RequestInit = {
          method: req.method,
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
          },
        };

        if (req.headers["content-type"]) {
          let cType = req.headers["content-type"] as string;
          if (cType.includes("application/json")) cType = "application/json";
          (fetchOptions.headers as Record<string, string>)["Content-Type"] =
            cType;
        }

        if (["POST", "PUT", "PATCH", "DELETE"].includes(req.method)) {
          if (req.body && req.body._isBase64File) {
            fetchOptions.body = Buffer.from(req.body.payload, "base64");
            (fetchOptions.headers as Record<string, string>)["content-type"] =
              req.body.contentType || "application/octet-stream";
          } else if (req.body) {
            fetchOptions.body = JSON.stringify(req.body);
            (fetchOptions.headers as Record<string, string>)["content-type"] =
              "application/json";
          }
        }

        // Remove uppercase Content-Type if it exists
        if ((fetchOptions.headers as Record<string, string>)["Content-Type"]) {
          delete (fetchOptions.headers as Record<string, string>)[
            "Content-Type"
          ];
        }

        if (fetchOptions.body) {
          const bodyStr = fetchOptions.body.toString();
          require('fs').appendFileSync('logs.txt', `\n[PUBLIC GRAPH CALL BODY LENGTH] ${bodyStr.length}\n`);
          if (bodyStr.length < 2000) {
              require('fs').appendFileSync('logs.txt', `[PUBLIC GRAPH CALL BODY] ${bodyStr}\n`);
          }
        }

        console.log(`[PUBLIC GRAPH CALL] ${req.method} ${graphUrl}`);
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
          grant_type: "refresh_token",
        });

        const tokenRes = await fetch(
          "https://login.microsoftonline.com/common/oauth2/v2.0/token",
          {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: tokenParams.toString(),
          },
        );

        if (tokenRes.ok) {
          const tokenData = (await tokenRes.json()) as any;
          creatorTokens = {
            accessToken: tokenData.access_token,
            refreshToken: tokenData.refresh_token || creatorTokens.refreshToken,
            expiresIn: tokenData.expires_in || 3600,
            acquiredAt: Date.now(),
          };
          // Save back to DB
          db.prepare(`UPDATE forms SET creatorTokens = ? WHERE id = ?`).run(
            JSON.stringify(creatorTokens),
            formId,
          );

          // Retry
          graphRes = await attemptGraphCall(creatorTokens.accessToken);
        }
      }

      const graphStatus = graphRes.status;
      res.status(graphStatus);
      if (graphStatus === 204) return res.end();

      const contentType = graphRes.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        const bodyObj = await graphRes.json();
        res.json(bodyObj);
      } else {
        const txt = await graphRes.text();
        res.send(txt);
      }
    } catch (err: any) {
      console.error(`Secure Graph Proxy Error:`, err);
      res
        .status(500)
        .json({ error: `Secure Graph Proxy Error: ${err.message}` });
    }
  });

  // -------------------------------------------------------------
  // Dynamic QR Code Routes
  // -------------------------------------------------------------
  
  app.get("/api/qrcodes", (req, res) => {
    try {
      const qrs = db.prepare("SELECT * FROM qrcodes ORDER BY createdAt DESC").all();
      res.json(qrs);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/qrcodes", (req, res) => {
    try {
      const { title, type, targetData } = req.body;
      const id = uuidv4().substring(0, 8); // Short ID for QR
      db.prepare("INSERT INTO qrcodes (id, title, type, targetData) VALUES (?, ?, ?, ?)").run(id, title, type, targetData);
      const qr = db.prepare("SELECT * FROM qrcodes WHERE id = ?").get(id);
      res.json(qr);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put("/api/qrcodes/:id", (req, res) => {
    try {
      const { id } = req.params;
      const { title, type, targetData } = req.body;
      db.prepare("UPDATE qrcodes SET title = ?, type = ?, targetData = ?, updatedAt = CURRENT_TIMESTAMP WHERE id = ?").run(title, type, targetData, id);
      const qr = db.prepare("SELECT * FROM qrcodes WHERE id = ?").get(id);
      if (!qr) return res.status(404).json({ error: "Not found" });
      res.json(qr);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.delete("/api/qrcodes/:id", (req, res) => {
    try {
      const { id } = req.params;
      db.prepare("DELETE FROM qrcodes WHERE id = ?").run(id);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  // Short link redirection for Dynamic QR
  app.get("/qr/:id", (req, res) => {
    try {
      const { id } = req.params;
      const qr = db.prepare("SELECT * FROM qrcodes WHERE id = ?").get(id) as any;
      if (!qr) {
        return res.status(404).send("QR code not found or deleted.");
      }
      // Redirect directly to the URL (works for links, images, and videos if they are URLs)
      if (qr.targetData && qr.targetData.startsWith("http")) {
        return res.redirect(302, qr.targetData);
      } else {
        return res.status(400).send("Invalid target data for this QR code.");
      }
    } catch (err: any) {
      res.status(500).send("Server Error");
    }
  });

  // Serve static UI assets or run dev middleware
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Serve client production static files from dist/ folder
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(
      `Express application successfully configured and running on http://localhost:${PORT}`,
    );
  });
}

startServer();
