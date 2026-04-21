/**
 * ============================================================
 * TOASTMASTERS FLYER GENERATOR — Google Apps Script
 * Enthusiastic Embarkers Club #4039
 * ============================================================
 * HOW IT WORKS:
 * 1. You forward the VPE's meeting email to xdongwang52@gmail.com
 *    and type "make a meeting flyer" as the very first line
 * 2. This script detects the forwarded email (runs every 15 min)
 * 3. Claude API extracts date, theme, speaker FIRST NAMES from email
 * 4. Script looks up first names in Google Sheet → full names + photos
 * 5. Script emails you a ready-to-paste message for your Claude Project
 * 6. You paste the message into your Claude Project → Canva flyer is generated
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
// YOUR CONFIG — fill these in before running
// ─────────────────────────────────────────────────────────────
const CONFIG = {

  // Gmail that receives your forwarded emails
  TRIGGER_GMAIL: "xdongwang52@gmail.com",

  // Trigger config — script ONLY runs when:
  //   1. Email is FROM your Hotmail (you forwarding VPE email), AND
  //   2. The FIRST LINE of the email body contains "make a meeting flyer"
  //      (punctuation like "make a meeting flyer." is fine — it's ignored)
  HOTMAIL_ADDRESS: "redtrees2004@hotmail.com",
  TRIGGER_PHRASE: "make a meeting flyer",

  // Where to send the ready-to-paste notification
  NOTIFY_EMAIL: "redtrees2004@hotmail.com",

  // Claude API key — get from console.anthropic.com
  CLAUDE_API_KEY: "sk-ant-YOUR_KEY_HERE",

  // Fixed Canva template ID — never change
  CANVA_TEMPLATE_ID: "DAHHD3cPQl0",

  // Canva MCP server URL — do not change
  CANVA_MCP_URL: "https://mcp.canva.com/mcp",

  // Google Sheet ID
  SHEET_ID: "16cELM-a2GMDajc5Kqg55gtK1vreMfect1_pFEV8FMew",
  SHEET_TAB: "Members",

  // Google Drive folder containing member photos
  // Photos must be named First-Last: Jordan-Thomas.jpg, Fasi-Shariff.png, etc.
  // Use PHOTOS_FOLDER_ID (from Drive URL) for reliability — falls back to name lookup
  PHOTOS_FOLDER_NAME: "MemberPhotos",
  PHOTOS_FOLDER_ID:   "1vfo7rT_dUsIOfCYjeOwXkRDfCHbRno1q",

  // Default values — used when email doesn't mention them
  DEFAULTS: {
    time:        "7:00 - 8:30pm PT",
    location:    "2600 S. Main St Corona CA\nOr on Zoom",
    club_name:   "Enthusiastic Embarkers Club #4039 (Est. 1979)",
    social_link: "4039.toastmastersclubs.org",
  },

  // Generic fallback portrait used when a member has no photo in the sheet
  FALLBACK_PHOTO: "https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?w=400&h=400&fit=crop&crop=face",

  // Claude Project URL
  CLAUDE_PROJECT_URL: "https://bit.ly/3ORPJbF",
};


// ─────────────────────────────────────────────────────────────
// MAIN: Check Gmail for forwarded meeting email
// Runs automatically every 15 minutes via trigger
// ─────────────────────────────────────────────────────────────
function checkForForwardedEmail() {
  try {
    // Search for unread emails FROM your Hotmail to this Gmail
    const query = `from:${CONFIG.HOTMAIL_ADDRESS} to:${CONFIG.TRIGGER_GMAIL} is:unread`;
    const threads = GmailApp.search(query, 0, 10);

    if (threads.length === 0) {
      Logger.log("No unread emails from Hotmail found.");
      return;
    }

    for (const thread of threads) {
      const message = thread.getMessages()[thread.getMessages().length - 1];
      const body    = message.getPlainBody();
      const phrase  = CONFIG.TRIGGER_PHRASE.toLowerCase();

      // Check if FIRST LINE of body contains the trigger phrase
      // Strip punctuation and extra spaces so "make a meeting flyer." still matches
      const firstLine = body.trim().split("\n")[0].trim().toLowerCase().replace(/[^a-z0-9 ]/g, "").trim();
      if (!firstLine.includes(phrase)) {
        Logger.log(`Skipping — trigger phrase not found in first line: "${firstLine}"`);
        thread.markRead();
        continue;
      }

      Logger.log(`Trigger found in: "${message.getSubject()}"`);
      thread.markRead();

      // Strip the trigger phrase line — Claude only needs the meeting content
      const bodyWithoutTrigger = body.trim().split("\n").slice(1).join("\n").trim();

      runFlyerPipeline(message.getSubject(), bodyWithoutTrigger);
      return;
    }

    Logger.log("No emails with trigger phrase as first line found.");

  } catch (err) {
    Logger.log(`Error: ${err.message}`);
    notifyError(err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// PIPELINE: Full end-to-end flyer generation
// ─────────────────────────────────────────────────────────────
function runFlyerPipeline(subject, body) {
  Logger.log("=== Starting flyer pipeline ===");

  // Step 1: Extract meeting info using Claude API
  Logger.log("Step 1: Extracting meeting info...");
  const info = extractInfoWithClaude(subject, body);
  Logger.log(`Extracted: ${JSON.stringify(info)}`);

  // Step 2: Apply defaults for missing fields
  const meeting = {
    date:        info.date     || "Date TBC",
    time:        info.time     || CONFIG.DEFAULTS.time,
    theme:       info.theme    || "Theme TBC",
    location:    CONFIG.DEFAULTS.location,
    club_name:   CONFIG.DEFAULTS.club_name,
    social_link: CONFIG.DEFAULTS.social_link,
  };

  // Step 3: Look up speakers in Google Sheet by first name
  Logger.log("Step 2: Looking up speakers...");
  const firstNames = [info.speaker1, info.speaker2, info.speaker3];
  const speakers = firstNames.map((name, i) => {
    if (!name) return { fullName: `Speaker ${i+1} TBA`, photoUrl: "" };
    const found = lookupByFirstName(name);
    if (found) {
      Logger.log(`Found: ${name} → ${found.fullName}`);
      return {
        fullName: found.fullName,
        photoUrl: found.photoUrl || "",
      };
    }
    Logger.log(`Not found in sheet: ${name}`);
    return { fullName: name, photoUrl: "" };
  });

  // Step 4: Build ready-to-paste Claude Project message
  Logger.log("Step 3: Building flyer message...");
  const flyerMessage = buildFlyerMessage(meeting, speakers);

  // Step 5: Email you the message to paste into Claude Project
  Logger.log("Step 4: Sending message to your notify email...");
  sendFlyerMessage(meeting, flyerMessage);

  Logger.log("=== Pipeline complete! ===");
}


// ─────────────────────────────────────────────────────────────
// CLAUDE: Extract meeting info from email
// ─────────────────────────────────────────────────────────────
function extractInfoWithClaude(subject, body) {
  const prompt = `Extract meeting details from this Toastmasters email and return ONLY a JSON object.

Email subject: ${subject}
Email body: ${body}

Return ONLY this JSON (no explanation, no markdown, no backticks):
{
  "date": "day and date like Mon, Apr 28 — or null",
  "time": "time like 7:00 - 8:30pm PT — or null",
  "theme": "meeting theme — or null",
  "speaker1": "FIRST NAME ONLY of speaker 1 — or null",
  "speaker2": "FIRST NAME ONLY of speaker 2 — or null",
  "speaker3": "FIRST NAME ONLY of speaker 3 — or null"
}

Rules:
- Extract FIRST NAMES ONLY for speakers
- Return null for anything not clearly mentioned
- Return ONLY the raw JSON object`;

  const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": CONFIG.CLAUDE_API_KEY,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify({
      model: "claude-sonnet-4-5",
      max_tokens: 300,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  const responseData = JSON.parse(res.getContentText());
  Logger.log(`Claude extract response: ${JSON.stringify(responseData)}`);
  const raw = responseData.content[0].text.trim();

  try {
    return JSON.parse(raw);
  } catch (e) {
    const match = raw.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    throw new Error(`Could not parse Claude response: ${raw}`);
  }
}


// ─────────────────────────────────────────────────────────────
// BUILD: Format ready-to-paste message for your Claude Project
// ─────────────────────────────────────────────────────────────
function buildFlyerMessage(meeting, speakers) {
  const lines = [
    "New meeting flyer:",
    "Date: " + meeting.date,
    "Time: " + meeting.time,
    "Theme: " + meeting.theme,
    "Location: " + meeting.location.replace("\n", " / "),
    "Club Name: " + meeting.club_name,
    "Social Link: " + meeting.social_link,
  ];
  speakers.forEach(function(s, i) {
    lines.push("Speaker " + (i+1) + ": " + s.fullName + " — " + (s.photoUrl || "[no photo]"));
  });
  return lines.join("\n");
}


// ─────────────────────────────────────────────────────────────
// SHEET: Look up member by first name
// ─────────────────────────────────────────────────────────────
function lookupByFirstName(firstName) {
  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_TAB);

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const col = {
    first:  headers.indexOf("First Name"),
    full:   headers.indexOf("Full Name"),
    photo:  headers.indexOf("Photo URL"),
  };

  const search = firstName.toLowerCase().trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][col.first]).toLowerCase().trim() === search) {
      return {
        fullName: rows[i][col.full]  || firstName,
        photoUrl: rows[i][col.photo] || "",
      };
    }
  }
  return null;
}


// ─────────────────────────────────────────────────────────────
// NOTIFY: Email you the ready-to-paste Claude Project message
// ─────────────────────────────────────────────────────────────
function sendFlyerMessage(meeting, flyerMessage) {
  const subject = "🎤 Paste this into Claude to generate your flyer — " + meeting.date;

  const flyerHtml = flyerMessage.split("\n").join("<br>");

  const htmlBody = "<p>Your meeting details are ready. Follow these steps:</p>"
    + "<ol>"
    + "<li><a href='" + CONFIG.CLAUDE_PROJECT_URL + "'>Open your Claude Project</a></li>"
    + "<li>Paste the message below into the chat and press Enter:</li>"
    + "</ol>"
    + "<pre style='background:#f4f4f4;padding:12px;border-radius:4px;font-size:14px'>" + flyerMessage + "</pre>"
    + "<ol start='3'>"
    + "<li>Claude will generate your flyer automatically.</li>"
    + "<li>Approve the preview → get the Canva link → download PNG.</li>"
    + "</ol>"
    + "<p>— Toastmasters Flyer Bot</p>";

  GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, "", { htmlBody: htmlBody });
  Logger.log("Message sent to " + CONFIG.NOTIFY_EMAIL);
}


// ─────────────────────────────────────────────────────────────
// NOTIFY: Send error email
// ─────────────────────────────────────────────────────────────
function notifyError(message) {
  GmailApp.sendEmail(
    CONFIG.NOTIFY_EMAIL,
    "⚠️ Flyer Bot Error — Please generate manually",
    `The flyer bot ran into a problem:\n\n${message}\n\nPlease generate this week's flyer manually in your Claude Project.`
  );
}


// ─────────────────────────────────────────────────────────────
// MEMBER PHOTOS: Sync Drive folder → update Members sheet
// Run manually (or set a weekly trigger) when photos change.
// Photos must be named by first name: Jordan.jpg, Fasi.png, etc.
// ─────────────────────────────────────────────────────────────
function syncMemberPhotos() {
  const IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "webp", "heic"];

  // Step 1: Open the Google Drive folder
  // Prefers PHOTOS_FOLDER_ID (more reliable); falls back to name search
  let folder;
  try {
    if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID !== "YOUR_DRIVE_FOLDER_ID") {
      folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
    } else {
      const iter = DriveApp.getFoldersByName(CONFIG.PHOTOS_FOLDER_NAME);
      if (!iter.hasNext()) {
        throw new Error(`Drive folder "${CONFIG.PHOTOS_FOLDER_NAME}" not found. Check PHOTOS_FOLDER_NAME or set PHOTOS_FOLDER_ID.`);
      }
      folder = iter.next();
    }
  } catch (err) {
    throw new Error(`Could not open photos folder: ${err.message}`);
  }
  Logger.log(`Opened folder: ${folder.getName()}`);

  // Step 2: Open the Members sheet
  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col = {
    first: headers.indexOf("First Name"),
    full:  headers.indexOf("Full Name"),
    photo: headers.indexOf("Photo URL"),
  };

  // Step 3: Process each file in the folder
  const files = folder.getFiles();
  let updated = 0, added = 0, skipped = 0, errors = 0;

  while (files.hasNext()) {
    const file  = files.next();
    const fileName = file.getName();
    const ext   = fileName.split(".").pop().toLowerCase();

    // Skip non-image files (.pdf, .docx, etc.)
    if (!IMAGE_EXTENSIONS.includes(ext)) {
      Logger.log(`Skipping non-image: ${fileName}`);
      skipped++;
      continue;
    }

    // Parse "Jordan-Thomas.jpg" → firstName = "Jordan", fullName = "Jordan Thomas"
    const baseName  = fileName.substring(0, fileName.lastIndexOf(".")).trim();
    const nameParts = baseName.split("-");
    const firstName = nameParts[0].trim();
    const fullName  = nameParts.map(p => p.trim()).join(" ");

    // Make file publicly readable so Claude/Canva can fetch it
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      Logger.log(`WARNING: Could not set sharing for ${fileName}: ${err.message}`);
      errors++;
      continue;
    }

    const photoUrl = file.getUrl();

    // Find matching row in sheet (case-insensitive first name)
    const search = firstName.toLowerCase();
    let matchRow   = -1;
    let matchCount = 0;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][col.first]).toLowerCase().trim() === search) {
        matchCount++;
        if (matchRow === -1) matchRow = i;
      }
    }

    if (matchCount > 1) {
      // Duplicate names are unusual — warn and update only the first match
      Logger.log(`WARNING: Duplicate first name "${firstName}" (${matchCount} rows). Updating row ${matchRow + 1} only.`);
    }

    if (matchRow !== -1) {
      // Update existing member's Photo URL
      sheet.getRange(matchRow + 1, col.photo + 1).setValue(photoUrl);
      rows[matchRow][col.photo] = photoUrl; // keep local array in sync
      Logger.log(`Updated photo for: ${firstName}`);
      updated++;
    } else {
      // New member — add a placeholder row for the VPPR to complete
      const newRow = new Array(headers.length).fill("");
      newRow[col.first] = firstName;
      newRow[col.full]  = fullName;  // parsed from filename: "Jordan-Thomas" → "Jordan Thomas"
      newRow[col.photo] = photoUrl;
      sheet.appendRow(newRow);
      rows.push(newRow); // keep local array in sync
      Logger.log(`Added new row for: ${firstName} (${fullName}) — no further action needed`);
      added++;
    }
  }

  Logger.log(`=== syncMemberPhotos done: ${updated} updated, ${added} added, ${skipped} skipped, ${errors} errors ===`);

  // Step 4: Clear sheet URLs for photos that no longer exist in the folder
  clearDeletedPhotos();
}


// ─────────────────────────────────────────────────────────────
// TEST: Preview syncMemberPhotos() without writing to sheet
// Run this first to confirm the results look right
// ─────────────────────────────────────────────────────────────
function testSyncMemberPhotos() {
  const IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "webp", "heic"];
  Logger.log("=== DRY RUN — no changes will be made to the sheet ===");

  // Open folder
  let folder;
  if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID !== "YOUR_DRIVE_FOLDER_ID") {
    folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
  } else {
    const iter = DriveApp.getFoldersByName(CONFIG.PHOTOS_FOLDER_NAME);
    if (!iter.hasNext()) {
      Logger.log(`ERROR: Drive folder "${CONFIG.PHOTOS_FOLDER_NAME}" not found.`);
      return;
    }
    folder = iter.next();
  }
  Logger.log(`Folder found: ${folder.getName()}`);

  // Read sheet
  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col = { first: headers.indexOf("First Name") };

  // Report what syncMemberPhotos() would do for each file
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file     = files.next();
    const fileName = file.getName();
    const ext      = fileName.split(".").pop().toLowerCase();

    if (!IMAGE_EXTENSIONS.includes(ext)) {
      Logger.log(`[SKIP]        ${fileName} — not an image`);
      continue;
    }

    const baseName2  = fileName.substring(0, fileName.lastIndexOf(".")).trim();
    const firstName  = baseName2.split("-")[0].trim();
    const fullName2  = baseName2.split("-").map(p => p.trim()).join(" ");
    const search     = firstName.toLowerCase();
    let matchRow     = -1;
    let matchCount   = 0;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][col.first]).toLowerCase().trim() === search) {
        matchCount++;
        if (matchRow === -1) matchRow = i;
      }
    }

    if (matchRow !== -1) {
      const dupeNote = matchCount > 1 ? ` — WARNING: ${matchCount} duplicate rows, would update row ${matchRow + 1} only` : ` → row ${matchRow + 1}`;
      Logger.log(`[WOULD UPDATE] ${firstName}${dupeNote}`);
    } else {
      Logger.log(`[WOULD ADD]   ${firstName} (${fullName2}) — new row, fully populated from filename`);
    }
  }

  Logger.log("=== Dry run complete. Run syncMemberPhotos() to apply changes. ===");
}


// ─────────────────────────────────────────────────────────────
// MEMBER PHOTOS: Clear sheet URLs for deleted photos
// Called automatically at the end of syncMemberPhotos()
// Only clears URLs that point to files in the MemberPhotos folder
// ─────────────────────────────────────────────────────────────
function clearDeletedPhotos() {
  // Build a set of file IDs currently in the Drive folder
  let folder;
  if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID !== "YOUR_DRIVE_FOLDER_ID") {
    folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
  } else {
    const iter = DriveApp.getFoldersByName(CONFIG.PHOTOS_FOLDER_NAME);
    if (!iter.hasNext()) {
      Logger.log("clearDeletedPhotos: folder not found, skipping.");
      return;
    }
    folder = iter.next();
  }

  const existingIds = new Set();
  const files = folder.getFiles();
  while (files.hasNext()) existingIds.add(files.next().getId());

  // Scan the sheet for Drive URLs whose file no longer exists in the folder
  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col = {
    first: headers.indexOf("First Name"),
    photo: headers.indexOf("Photo URL"),
  };

  let cleared = 0;
  for (let i = 1; i < rows.length; i++) {
    const url = String(rows[i][col.photo] || "").trim();
    if (!url) continue;

    // Extract file ID from Drive URL: drive.google.com/file/d/FILE_ID/view
    const match = url.match(/\/d\/([^\/\?]+)/);
    if (!match) continue; // not a Drive URL — leave it alone

    const fileId = match[1];
    if (!existingIds.has(fileId)) {
      sheet.getRange(i + 1, col.photo + 1).setValue("");
      Logger.log(`Cleared stale photo URL for: ${rows[i][col.first]}`);
      cleared++;
    }
  }

  Logger.log(`=== clearDeletedPhotos done: ${cleared} URL(s) cleared ===`);
}


// ─────────────────────────────────────────────────────────────
// TEST: Preview clearDeletedPhotos() without writing to sheet
// ─────────────────────────────────────────────────────────────
function testClearDeletedPhotos() {
  Logger.log("=== DRY RUN — no changes will be made to the sheet ===");

  let folder;
  if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID !== "YOUR_DRIVE_FOLDER_ID") {
    folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
  } else {
    const iter = DriveApp.getFoldersByName(CONFIG.PHOTOS_FOLDER_NAME);
    if (!iter.hasNext()) {
      Logger.log("ERROR: folder not found.");
      return;
    }
    folder = iter.next();
  }

  const existingIds = new Set();
  const files = folder.getFiles();
  while (files.hasNext()) existingIds.add(files.next().getId());
  Logger.log(`Files currently in folder: ${existingIds.size}`);

  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col = {
    first: headers.indexOf("First Name"),
    photo: headers.indexOf("Photo URL"),
  };

  for (let i = 1; i < rows.length; i++) {
    const url = String(rows[i][col.photo] || "").trim();
    if (!url) continue;
    const match = url.match(/\/d\/([^\/\?]+)/);
    if (!match) continue;
    if (!existingIds.has(match[1])) {
      Logger.log(`[WOULD CLEAR] ${rows[i][col.first]} — file no longer in folder`);
    }
  }

  Logger.log("=== Dry run complete. Run clearDeletedPhotos() to apply. ===");
}


// ─────────────────────────────────────────────────────────────
// SETUP: Run once to authorize Gmail + Sheets
// ─────────────────────────────────────────────────────────────
function authorizeServices() {
  GmailApp.getInboxUnreadCount();
  SpreadsheetApp.openById(CONFIG.SHEET_ID);
  Logger.log("✅ Authorization complete! Now set the 15-minute trigger on checkForForwardedEmail.");
}


// ─────────────────────────────────────────────────────────────
// TEST: Run this to test with a sample email (no real email needed)
// ─────────────────────────────────────────────────────────────
function testWithSampleEmail() {
  const subject = "Toastmasters Meeting Monday Apr 28";
  const body = `Date: Mon Apr 28
Theme: Finding Your True North
Speaker 1: Jordan
Speaker 2: Fasi
Speaker 3: Kristina`;
  runFlyerPipeline(subject, body);
}
