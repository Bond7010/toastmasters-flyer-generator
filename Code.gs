/**
 * ============================================================
 * TOASTMASTERS FLYER GENERATOR — Google Apps Script
 * ============================================================
 * HOW IT WORKS:
 * 1. You forward the VPE's meeting email to your Gmail
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
  TRIGGER_GMAIL: "YOUR_GMAIL@gmail.com",

  // Trigger config — script ONLY runs when:
  //   1. Email is FROM your Hotmail (you forwarding VPE email), AND
  //   2. The FIRST LINE of the email body contains "make a meeting flyer"
  //      (punctuation like "make a meeting flyer." is fine — it's ignored)
  HOTMAIL_ADDRESS: "YOUR_HOTMAIL@hotmail.com",
  TRIGGER_PHRASE: "make a meeting flyer",

  // Where to send the ready-to-paste notification
  NOTIFY_EMAIL: "YOUR_NOTIFY_EMAIL@hotmail.com",

  // Claude API key — get from console.anthropic.com
  CLAUDE_API_KEY: "sk-ant-YOUR_KEY_HERE",

  // Fixed Canva template ID — copy from your template's URL
  // URL: canva.com/design/[THIS PART]/edit
  CANVA_TEMPLATE_ID: "YOUR_CANVA_TEMPLATE_ID",

  // Canva MCP server URL — do not change
  CANVA_MCP_URL: "https://mcp.canva.com/mcp",

  // Google Sheet ID — copy from your sheet's URL
  // URL: docs.google.com/spreadsheets/d/[THIS PART]/edit
  SHEET_ID: "YOUR_GOOGLE_SHEET_ID",
  SHEET_TAB: "Members",

  // Default values — used when email doesn't mention them
  DEFAULTS: {
    time:        "7:00 - 8:30pm PT",
    location:    "YOUR_MEETING_ADDRESS\nOr on Zoom",
    club_name:   "YOUR_CLUB_NAME #XXXX (Est. YYYY)",
    social_link: "YOURCLUB.toastmastersclubs.org",
  },

  // Generic fallback portraits when no photo available
  FALLBACKS: {
    male:    "https://images.unsplash.com/photo-1560250097-0b93528c311a?w=400&h=400&fit=crop&crop=face",
    female:  "https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=400&h=400&fit=crop&crop=face",
    neutral: "https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?w=400&h=400&fit=crop&crop=face",
  },

  // Claude Project URL — copy from your Claude Project's URL bar
  CLAUDE_PROJECT_URL: "https://claude.ai/project/YOUR_PROJECT_ID",
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
    if (!name) return { fullName: `Speaker ${i+1} TBA`, photoUrl: "", gender: "neutral" };
    const found = lookupByFirstName(name);
    if (found) {
      Logger.log(`Found: ${name} → ${found.fullName}`);
      return found;
    }
    Logger.log(`Not found in sheet: ${name}`);
    return { fullName: name, photoUrl: "", gender: "neutral" };
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
    gender: headers.indexOf("Gender"),
  };

  const search = firstName.toLowerCase().trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][col.first]).toLowerCase().trim() === search) {
      return {
        fullName: rows[i][col.full]   || firstName,
        photoUrl: rows[i][col.photo]  || "",
        gender:   rows[i][col.gender] || "neutral",
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
  const body = "Hi,\n\n"
    + "Your meeting details are ready. Follow these steps:\n\n"
    + "1. Open your Claude Project:\n"
    + "   " + CONFIG.CLAUDE_PROJECT_URL + "\n\n"
    + "2. Paste the message below into the chat and press Enter:\n\n"
    + "─────────────────────────────────\n"
    + flyerMessage + "\n"
    + "─────────────────────────────────\n\n"
    + "3. Claude will generate your flyer automatically.\n"
    + "4. Approve the preview → get the Canva link → download PNG.\n\n"
    + "— Toastmasters Flyer Bot";

  GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, body);
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
