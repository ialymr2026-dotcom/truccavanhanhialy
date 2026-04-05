import express from "express";
// import { createServer as createViteServer } from "vite"; // Move to dynamic import
import path from "path";
import { google } from "googleapis";
import cookieParser from "cookie-parser";
import dotenv from "dotenv";

import fs from "fs";
import admin from "firebase-admin";
dotenv.config();

let firebaseConfig: any = {};
try {
  const configPath = path.join(process.cwd(), "firebase-applet-config.json");
  if (fs.existsSync(configPath)) {
    firebaseConfig = JSON.parse(fs.readFileSync(configPath, "utf-8"));
  } else {
    console.warn("firebase-applet-config.json not found, using environment variables");
    firebaseConfig = {
      projectId: process.env.FIREBASE_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT || "default-project",
    };
  }
} catch (e) {
  console.error("Error loading firebase-applet-config.json:", e);
  firebaseConfig = { projectId: process.env.FIREBASE_PROJECT_ID || "default-project" };
}

// Initialize Firebase Admin
if (admin.apps.length === 0) {
  try {
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
        projectId: serviceAccount.project_id || firebaseConfig.projectId,
      });
      console.log("Firebase Admin initialized with Service Account from ENV");
    } else {
      admin.initializeApp({
        projectId: firebaseConfig.projectId,
      });
      console.log("Firebase Admin initialized with default credentials");
    }
  } catch (e) {
    console.error("Firebase Admin initialization error:", e);
  }
}
// const firestore = admin.firestore(); // Initialize lazily
const getFirestore = () => admin.firestore();

const saveTokens = async (tokens: any) => {
  try {
    const firestore = getFirestore();
    console.log("Saving tokens to Firestore (config/google_auth)...");
    await firestore.collection("config").doc("google_auth").set({
      tokens,
      updatedAt: new Date().toISOString()
    });
    console.log("Tokens saved to Firestore successfully.");
  } catch (e) {
    console.error("Error saving tokens to Firestore:", e);
  }
};

const loadTokens = async () => {
  try {
    const firestore = getFirestore();
    console.log("Attempting to load tokens from Firestore (config/google_auth)...");
    const doc = await firestore.collection("config").doc("google_auth").get();
    if (doc.exists) {
      const tokens = doc.data()?.tokens;
      console.log("Tokens found in Firestore:", !!tokens);
      return tokens || null;
    }
    console.log("No tokens document found in Firestore.");
  } catch (e) {
    console.error("Error loading tokens from Firestore:", e);
  }
  return null;
};

console.log("GOOGLE_CLIENT_ID exists:", !!process.env.GOOGLE_CLIENT_ID);
console.log("GOOGLE_CLIENT_SECRET exists:", !!process.env.GOOGLE_CLIENT_SECRET);
console.log("APP_URL:", process.env.APP_URL);

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(cookieParser());

const getOAuth2Client = () => {
  const appUrl = (process.env.APP_URL || 'http://localhost:3000').trim().replace(/\/$/, "");
  const clientId = process.env.GOOGLE_CLIENT_ID?.trim();
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET?.trim();
  
  if (!clientId || !clientSecret) {
    throw new Error("Missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET environment variables");
  }

  return new google.auth.OAuth2(
    clientId,
    clientSecret,
    `${appUrl}/api/auth/google/callback`
  );
};

app.get("/api/debug/env", (req, res) => {
  const id = process.env.GOOGLE_CLIENT_ID || "";
  const secret = process.env.GOOGLE_CLIENT_SECRET || "";
  const url = process.env.APP_URL || "";
  res.json({
    id_exists: !!id,
    id_length: id.length,
    id_start: id.substring(0, 5),
    id_end: id.substring(id.length - 5),
    secret_exists: !!secret,
    secret_length: secret.length,
    url_exists: !!url,
    url: url
  });
});

// Auth Routes
app.get("/api/auth/google", (req, res) => {
  try {
    const oauth2Client = getOAuth2Client();
    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/userinfo.email"],
      prompt: "consent"
    });
    res.redirect(url);
  } catch (error) {
    console.error("Error generating Auth URL:", error);
    res.status(500).json({ 
      error: "Failed to generate Google Auth URL", 
      details: error instanceof Error ? error.message : String(error) 
    });
  }
});

app.get("/api/auth/google/url", (req, res) => {
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/userinfo.email"],
    prompt: "consent"
  });
  res.json({ url });
});

app.get("/api/auth/google/callback", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  const { code } = req.query;
  try {
    const { tokens } = await oauth2Client.getToken(code as string);
    
    // Save tokens to server for "permanent" access for everyone
    await saveTokens(tokens);

    res.cookie("google_tokens", JSON.stringify(tokens), {
      httpOnly: true,
      secure: true,
      sameSite: "none",
    });
    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              setTimeout(() => window.close(), 1000);
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error) {
    console.error("Error getting tokens:", error);
    res.status(500).send("Authentication failed");
  }
});

app.get("/api/auth/status", async (req, res) => {
  const cookieTokens = req.cookies.google_tokens;
  const firestoreTokens = await loadTokens();
  
  console.log("Auth Status Check - Cookie:", !!cookieTokens, "Firestore:", !!firestoreTokens);
  
  res.json({ 
    authenticated: !!cookieTokens || !!firestoreTokens,
    source: cookieTokens ? "cookie" : (firestoreTokens ? "firestore" : "none")
  });
});

app.get("/api/debug/auth", async (req, res) => {
  const firestoreTokens = await loadTokens();
  res.json({
    has_service_account: !!process.env.FIREBASE_SERVICE_ACCOUNT,
    firebase_project_id: firebaseConfig.projectId,
    firestore_tokens_exist: !!firestoreTokens,
    cookie_tokens_exist: !!req.cookies.google_tokens,
    env_vars: {
      GOOGLE_CLIENT_ID: !!process.env.GOOGLE_CLIENT_ID,
      GOOGLE_CLIENT_SECRET: !!process.env.GOOGLE_CLIENT_SECRET,
      APP_URL: process.env.APP_URL
    }
  });
});

app.post("/api/sheets/update", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    tokens = JSON.parse(tokensStr);
  } else {
    tokens = await loadTokens();
  }

  if (!tokens) return res.status(401).json({ error: "Not authenticated" });

  oauth2Client.setCredentials(tokens);

  const { spreadsheetId, updates } = req.body;
  // updates: Array<{ date: string, person: string, shift: string }>

  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  try {
    // Get spreadsheet metadata to find sheet names
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];

    for (const update of updates) {
      const date = new Date(update.date);
      const month = date.getMonth() + 1;
      const monthStr = month < 10 ? `0${month}` : `${month}`;
      const year = date.getFullYear();
      const day = date.getDate();
      const dayStr = day < 10 ? `0${day}` : `${day}`;
      
      // Match LICH_MM_YYYY or other common formats
      const sheetName = sheetNames.find(n => 
        n === `LICH_${monthStr}_${year}` ||
        n === `LICH_${month}_${year}` ||
        n?.includes(`Tháng ${month}`) || 
        n?.includes(`Tháng ${monthStr}`) || 
        n?.includes(`T${month}`) ||
        n?.includes(`T${monthStr}`) ||
        n === `${month}` || 
        n === monthStr
      );

      if (!sheetName) {
        console.warn(`Sheet for month ${month} not found in:`, sheetNames);
        continue;
      }

      // Read the sheet to find the person and the date
      const range = `${sheetName}!A1:CZ500`; // Increased range to cover more columns (up to CZ) and rows
      const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = response.data.values;

      if (!rows || rows.length === 0) {
        console.warn(`No data found in sheet: ${sheetName}`);
        continue;
      }

      // Find person column (assuming names are in the first row)
      const header = rows[0];
      let personColIdx = -1;
      const searchName = update.name.trim().normalize('NFC').toLowerCase();
      
      for (let i = 1; i < header.length; i++) {
        const cell = header[i]?.toString().trim().normalize('NFC').toLowerCase();
        if (cell && cell.includes(searchName)) {
          personColIdx = i;
          break;
        }
      }

      if (personColIdx === -1) {
        console.warn(`Person "${update.name}" not found in sheet: ${sheetName}. Normalized search: "${searchName}". Header columns:`, header.map(h => h?.toString().normalize('NFC')));
        continue;
      }

      // Find date row (assuming dates are in the first column)
      let dateRowIdx = -1;
      const dateParts = update.date.split('-'); // YYYY-MM-DD
      const dayNum = parseInt(dateParts[2]);
      const currentDayStr = dayNum < 10 ? `0${dayNum}` : `${dayNum}`;
      const dd_mm_yyyy = `${currentDayStr}/${dateParts[1]}/${dateParts[0]}`;
      const d_m_yyyy = `${dayNum}/${parseInt(dateParts[1])}/${dateParts[0]}`;

      for (let i = 1; i < rows.length; i++) {
        const cell = rows[i][0]?.toString().trim();
        if (!cell) continue;
        
        // Match various date formats: full date, dd/mm/yyyy, or just the day number
        if (
          cell === update.date || 
          cell === dd_mm_yyyy || 
          cell === d_m_yyyy || 
          cell.includes(dd_mm_yyyy) ||
          cell === dayNum.toString() ||
          cell === currentDayStr
        ) {
          dateRowIdx = i;
          break;
        }
      }

      if (dateRowIdx === -1) {
        console.warn(`Date "${update.date}" not found in sheet: ${sheetName}. First column sample:`, rows.slice(0, 10).map(r => r[0]));
        continue;
      }

      // Update the cell in LICH sheet
      const colLetter = getColumnLetter(personColIdx);
      const cellRange = `${sheetName}!${colLetter}${dateRowIdx + 1}`;
      console.log(`Updating ${update.name} on ${update.date} to ${update.shift} at ${cellRange}`);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: cellRange,
        valueInputOption: "RAW",
        requestBody: { values: [[update.shift]] }
      });

      // Also try to update TONG_HOP sheet if it exists
      const summarySheetName = `TONG_HOP_${monthStr}_${year}`;
      if (sheetNames.includes(summarySheetName)) {
        const summaryRange = `${summarySheetName}!A1:CZ500`; // Increased range
        const summaryResponse = await sheets.spreadsheets.values.get({ spreadsheetId, range: summaryRange });
        const summaryRows = summaryResponse.data.values;
        
        if (summaryRows && summaryRows.length > 0) {
          // Find person row in TONG_HOP (Column A, starting from row 5)
          let summaryRowIdx = -1;
          for (let i = 4; i < summaryRows.length; i++) {
            const cell = summaryRows[i][0]?.toString().trim();
            if (cell && cell.toLowerCase() === update.name.toLowerCase().trim()) {
              summaryRowIdx = i;
              break;
            }
          }

          // Find day column in TONG_HOP (Row 1, starting from column B)
          const dayNum = parseInt(dateParts[2]);
          const summaryColIdx = dayNum; // Day 1 is in Column B (index 1)

          if (summaryRowIdx !== -1) {
            const summaryColLetter = getColumnLetter(summaryColIdx);
            const summaryCellRange = `${summarySheetName}!${summaryColLetter}${summaryRowIdx + 1}`;
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: summaryCellRange,
              valueInputOption: "RAW",
              requestBody: { values: [[update.shift]] }
            });
          }
        }
      }
    }

    // 4. Append notification to "Bảng báo cơm ca" sheet if it exists
    const mealSheetName = sheetNames.find(n => 
      n?.toLowerCase().includes("báo cơm") || 
      n?.toLowerCase().includes("cơm ca") ||
      n?.toLowerCase().includes("meal")
    );
    
    if (mealSheetName) {
      const now = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
      const mealUpdates = updates.map(u => [
        now, 
        u.name, 
        u.date, 
        u.shift, 
        "Đã thay đổi"
      ]);
      
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${mealSheetName}!A:E`,
          valueInputOption: "RAW",
          requestBody: { values: mealUpdates }
        });
        console.log(`Appended ${mealUpdates.length} rows to ${mealSheetName}`);
      } catch (appendError) {
        console.error("Error appending to meal sheet:", appendError);
      }
    }

    res.json({ success: true });
  } catch (error) {
    console.error("Error updating sheet:", error);
    res.status(500).json({ error: "Failed to update sheet" });
  }
});

// Catch-all for API routes that don't exist
app.all("/api/*", (req, res) => {
  res.status(404).json({ error: `API route ${req.path} not found` });
});

function getColumnLetter(columnIdx: number): string {
  let temp, letter = "";
  while (columnIdx >= 0) {
    temp = columnIdx % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIdx = (columnIdx - temp) / 26 - 1;
  }
  return letter;
}

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const viteModule = "vite";
    const { createServer: createViteServer } = await import(viteModule);
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);

    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }
}

startServer();

export default app;
