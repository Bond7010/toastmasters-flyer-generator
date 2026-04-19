# Toastmasters Flyer Generator

A semi-automated weekly meeting flyer generator for Toastmasters clubs. Forward a meeting notice email → Google Apps Script extracts the details → Claude Project edits your Canva template → you download a ready-to-share PNG flyer.

**Cost:** ~$0.01 per flyer (Claude API usage only)

---

## How It Works

```
VPE sends meeting email
        ↓
You forward it to your Gmail
(type "make a meeting flyer" as the first line)
        ↓
Google Apps Script (runs every 15 min):
  • Detects the forwarded email
  • Calls Claude API to extract: date, theme, speaker first names
  • Looks up speakers in Google Sheet → full names + photo URLs
  • Emails you a ready-to-paste message
        ↓
You open your Claude Project, paste the message
        ↓
Claude edits your Canva template
(uploads photos, swaps text)
        ↓
You get the Canva edit link → download PNG → distribute
```

---

## Prerequisites

- [Google account](https://accounts.google.com) (for Gmail + Google Sheets + Apps Script)
- [Claude.ai account](https://claude.ai) with the Canva connector enabled
- [Canva account](https://canva.com) (free tier works)
- [Anthropic API key](https://console.anthropic.com) (free trial available)

---

## Setup Guide

### Step 1 — Create Your Canva Template

1. Sign in to Canva and create a new design (recommended: 1080×1080px for social media)
2. Design your flyer with placeholder text for:
   - Meeting date/time
   - Meeting theme
   - Club name and social link
   - Speaker 1, 2, 3 names
   - Speaker 1, 2, 3 photo slots
3. Save the design and copy the **template ID** from the URL:
   ```
   canva.com/design/[YOUR_TEMPLATE_ID]/edit
   ```

#### Finding Your Canva Element IDs

Each text box and image in your Canva template has a unique element ID. You need these so Claude knows which elements to update. To find them:

1. Open a Claude.ai conversation with the Canva connector active
2. Ask Claude: *"Start an editing session on my Canva template [YOUR_TEMPLATE_ID] and list all element IDs with their labels"*
3. Note the element IDs for: date/time, location, theme, club name, social link, speaker name slots (×3), speaker photo slots (×3)

You'll provide these element IDs to Claude in your Claude Project's skill file (see Step 4).

---

### Step 2 — Create the Google Sheet Member Database

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name the first tab **Members** (exact spelling matters)
3. Add these column headers in row 1 (exact spelling matters):

| First Name | Full Name | Photo URL | Gender |
|---|---|---|---|
| Jordan | Jordan Thomas | https://drive.google.com/file/d/FILE_ID/view | neutral |
| Alex | Alex Rivera | https://drive.google.com/file/d/FILE_ID/view | female |
| Sam | Sam Patel | | male |

**Notes:**
- **First Name** — must exactly match how the VPE writes names in meeting emails
- **Photo URL** — use a Google Drive sharing link (`Anyone with the link can view`). Leave blank to use a generic portrait fallback.
- **Gender** — `male`, `female`, or `neutral` — used to select the right fallback portrait when no photo is available
- The script auto-converts Drive share URLs to direct image URLs

4. Copy the **Sheet ID** from the URL:
   ```
   docs.google.com/spreadsheets/d/[YOUR_SHEET_ID]/edit
   ```

---

### Step 3 — Get a Claude API Key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign in or create an account
3. Navigate to **API Keys** → **Create Key**
4. Copy the key (starts with `sk-ant-`) — you'll only see it once

---

### Step 4 — Set Up Your Claude Project

1. Go to [claude.ai](https://claude.ai) and create a new **Project**
2. Add the Canva connector to the project (Settings → Integrations → Canva)
3. Upload a skill file to the project with instructions for generating flyers, including:
   - Your Canva template ID
   - Your element IDs (from Step 1)
   - Your club's default values (location, club name, social link)
4. Copy your **Project URL**:
   ```
   https://claude.ai/project/[YOUR_PROJECT_ID]
   ```

---

### Step 5 — Set Up Google Apps Script

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete the default `myFunction` code
3. Paste the entire contents of `Code.gs` from this repository
4. Fill in the `CONFIG` section at the top:

```javascript
const CONFIG = {
  TRIGGER_GMAIL:      "your-gmail@gmail.com",
  HOTMAIL_ADDRESS:    "your-email@hotmail.com",   // the address you forward FROM
  TRIGGER_PHRASE:     "make a meeting flyer",      // first line of forwarded email
  NOTIFY_EMAIL:       "your-email@hotmail.com",   // where to receive the formatted message
  CLAUDE_API_KEY:     "sk-ant-YOUR_KEY_HERE",
  CANVA_TEMPLATE_ID:  "YOUR_CANVA_TEMPLATE_ID",
  SHEET_ID:           "YOUR_GOOGLE_SHEET_ID",
  CLAUDE_PROJECT_URL: "https://claude.ai/project/YOUR_PROJECT_ID",

  DEFAULTS: {
    time:        "7:00 - 8:30pm PT",              // your club's usual meeting time
    location:    "123 Main St, City, State\nOr on Zoom",
    club_name:   "Your Club Name #XXXX (Est. YYYY)",
    social_link: "YOURCLUB.toastmastersclubs.org",
  },
  // ... (leave FALLBACKS and CANVA_MCP_URL unchanged)
};
```

5. Save the project (Ctrl+S)

#### Authorize Gmail + Sheets

1. In the script editor, select `authorizeServices` from the function dropdown
2. Click **Run**
3. Follow the Google authorization prompts (you'll need to approve Gmail and Sheets access)

#### Set Up the 15-Minute Trigger

1. In the script editor, click **Triggers** (clock icon in the left sidebar)
2. Click **+ Add Trigger**
3. Set:
   - Function: `checkForForwardedEmail`
   - Event source: Time-driven
   - Type: Minutes timer
   - Interval: Every 15 minutes
4. Click **Save**

---

### Step 6 — Test It

Run `testWithSampleEmail()` in the script editor to verify the full pipeline works without needing a real email. Check your notify email inbox for the formatted message.

---

## Weekly Workflow

Once set up, generating a flyer each week takes under 2 minutes:

1. **Forward** the VPE's meeting email to your Gmail
2. **Type** `make a meeting flyer` as the very first line (before the forwarded content)
3. **Wait** up to 15 minutes — the script emails you a formatted message
4. **Open** your Claude Project and **paste** the message
5. **Review** the flyer preview Claude generates
6. **Download** the PNG from the Canva link → share with your club

---

## Troubleshooting

**Script doesn't trigger**
- Check that the trigger phrase is the *very first line* of the email body
- Verify `HOTMAIL_ADDRESS` exactly matches the address you're forwarding from
- Check Apps Script execution logs: Extensions → Apps Script → Executions

**"Not found in sheet" for a speaker**
- The first name in the email must match the **First Name** column exactly (case-insensitive)
- Check for extra spaces or nicknames

**Claude can't parse the email**
- Check your `CLAUDE_API_KEY` is correct and has credits
- Run `testWithSampleEmail()` to isolate whether the issue is email detection or Claude parsing

**Canva elements not updating**
- Verify element IDs in your Claude Project skill file match your actual template
- Re-open the Claude Project, start a fresh conversation, and paste the message again

**Photo not appearing**
- Confirm the Google Drive photo is shared as "Anyone with the link can view"
- Try opening the photo URL in an incognito browser window to verify it's accessible

---

## File Reference

| File | Purpose |
|---|---|
| `Code.gs` | Google Apps Script — paste into script.google.com |
| `README.md` | This setup guide |
| `.gitignore` | Keeps credential files out of version control |

---

## Toastmasters Brand Colors

Use these in your Canva template to stay on-brand:

| Name | Hex |
|---|---|
| Loyal Blue | `#004165` |
| True Maroon | `#772432` |
| Cool Gray | `#A9B2B1` |
| Happy Yellow | `#F2DF74` |

**Fonts:** Montserrat (headlines), Source Sans 3 (body)

---

## Contributing

Found a bug or have an improvement? Open an issue or pull request. This tool is maintained by Toastmasters VPPRs for VPPRs.
