# 📧 Mail Merge Pro

A powerful, free, and open-source mail merge solution built entirely on **Google Apps Script**. Send personalized bulk emails directly from Google Sheets using Gmail — no third-party services, no API keys, no monthly fees.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| **Personalized Emails** | Use `{{column_name}}` variables in subject & body — auto-replaced per row |
| **Smart Greetings** | Auto-generates "Dear Sir/Ma'am" based on Name + Title (Mr/Mrs/Ms) columns |
| **Rich Text Editor** | Compose HTML emails with formatting, links, and styles |
| **File Attachments** | Attach files from Google Drive or use Quick Attach for resumes |
| **Sender Aliases** | Send from any verified Gmail alias |
| **CC & BCC Support** | Per-row CC/BCC from spreadsheet columns |
| **Scheduled Sending** | Schedule individual emails or entire campaigns for later |
| **Campaign Manager** | Create, reschedule, or cancel scheduled campaigns |
| **Activity Monitor** | Real-time progress tracking with send/fail counts |
| **Email Templates** | Save, load, rename, and delete reusable templates (stored in Drive) |
| **Default Templates** | One-click install of job application & networking templates |
| **Delay Options** | Random, fixed, or no delay between emails to avoid spam filters |
| **Row Range Selection** | Send to specific rows (e.g., rows 2–10) |
| **Send Cancellation** | Cancel mid-send — remaining rows marked as "Cancelled" |
| **Status Tracking** | Auto-writes send status ("Sent", "Failed", "Scheduled") to your sheet |
| **Daily Quota Check** | Warns you before exceeding Gmail's daily send limit |
| **Test Email** | Send a test email to yourself before the full campaign |

---

## 📁 Project Structure

```
mail-merge-pro/
├── Code.gs            # All server-side Apps Script logic
├── Sidebar.html       # Main sidebar UI (step-by-step wizard)
├── Options.html       # Templates, Scheduled, Activity, Campaigns dialogs
├── appsscript.json    # Manifest with OAuth scopes and config
├── test_data.csv      # Sample data for testing (anonymized)
├── LICENSE            # MIT License
└── README.md          # This file
```

---

## 🚀 Installation

### Step 1: Create a Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Add your data with headers in Row 1 (e.g., `Name`, `Email`, `Company`, `Position`)
3. You can import `test_data.csv` to get started quickly

### Step 2: Open Apps Script Editor

1. In your Google Sheet, go to **Extensions → Apps Script**
2. This opens the Apps Script editor

### Step 3: Add the Code

1. **Delete** the default code in `Code.gs`
2. Copy the contents of [`Code.gs`](Code.gs) from this repo and paste it in
3. Click **+ (Add a file)** → **HTML** → name it `Sidebar` (not `Sidebar.html`, the editor adds `.html` automatically)
4. Paste the contents of [`Sidebar.html`](Sidebar.html)
5. Repeat: **+ (Add a file)** → **HTML** → name it `Options`
6. Paste the contents of [`Options.html`](Options.html)

### Step 4: Configure the Manifest

1. In the Apps Script editor, click ⚙️ **Project Settings** (gear icon)
2. Check **"Show 'appsscript.json' manifest file in editor"**
3. Click on `appsscript.json` in the file list
4. Replace its contents with the [`appsscript.json`](appsscript.json) from this repo

### Step 5: Enable Gmail API

1. In the Apps Script editor, click **Services** (+ icon in the left sidebar)
2. Scroll down to **Gmail API** and click **Add**
3. Keep the default identifier as `Gmail`

### Step 6: Save & Authorize

1. Click **Save** (💾) or press `Ctrl+S`
2. Go back to your Google Sheet and **refresh the page**
3. You'll see a new menu: **📧 Mail Merge Pro**
4. Click **Open Mail Merge** — Google will ask you to authorize the script
5. Click through the authorization prompts (you may need to click "Advanced" → "Go to Mail Merge Pro")

---

## 📖 Usage

### Sending Emails

1. **Open** the sidebar: `📧 Mail Merge Pro → Open Mail Merge`
2. **Step 1** — Select your data sheet and map the columns (Email, Name, Title, CC, BCC)
3. **Step 2** — Compose your email using the rich text editor
   - Use `{{column_name}}` to insert personalized data (e.g., `{{Name}}`, `{{Company}}`)
   - Use `{{greeting}}` for smart auto-generated greetings
4. **Step 3** — Attach files from Google Drive (optional)
5. **Step 4** — Review, configure delay settings, and send!

### Template Variables

| Variable | Replaced With |
|----------|---------------|
| `{{greeting}}` | Smart greeting based on Name + Title column |
| `{{Name}}` | Value from the "Name" column for each row |
| `{{Company}}` | Value from the "Company" column for each row |
| `{{Position}}` | Value from the "Position" column for each row |
| `{{AnyColumn}}` | Value from any column header in your sheet |

### Scheduling Campaigns

1. In Step 4, select **"Schedule for later"** instead of sending immediately
2. Pick a date and time
3. The campaign will run automatically at the scheduled time via a time-based trigger
4. Manage scheduled campaigns via `📧 Mail Merge Pro → Scheduled Campaigns`

### Templates

- **Save**: After composing an email, save it as a template for reuse
- **Load**: Load any saved template when composing
- **Default Templates**: Click `📧 Mail Merge Pro → Templates` → "Install Default Templates" for pre-built job application and networking templates
- Templates are stored as JSON files in a `Mail Merge Pro Templates` folder in your Google Drive

---

## 🔒 OAuth Scopes

This script requests the following permissions (defined in `appsscript.json`):

| Scope | Why It's Needed |
|-------|-----------------|
| `gmail.send` | Send emails via Gmail |
| `gmail.compose` | Create drafts for scheduled emails |
| `gmail.settings.basic` | Read sender aliases (Send As addresses) |
| `spreadsheets` | Read spreadsheet data (recipients, columns) |
| `drive` | Read attachments, save/load templates |
| `script.container.ui` | Show sidebar and dialogs in Google Sheets |
| `userinfo.email` | Get the current user's email address |
| `script.scriptapp` | Create time-based triggers for scheduling |

> **Note:** This script runs entirely within your Google account. No data is sent to any external server.

---

## ⚙️ Configuration

### `appsscript.json`

- **`timeZone`**: Set to `"Asia/Kolkata"` by default. Change this to your timezone (e.g., `"America/New_York"`, `"Europe/London"`)
- **`runtimeVersion`**: Uses V8 runtime
- **`exceptionLogging`**: Logs to Stackdriver (Google Cloud Logging)

### Gmail Daily Limits

Google imposes daily email sending limits:
- **Free Gmail accounts**: ~500 emails/day
- **Google Workspace accounts**: ~2,000 emails/day

The script checks your remaining quota before sending and warns you if you'll exceed the limit.

---

## 🤝 Contributing

Contributions are welcome! Here's how:

1. **Fork** this repository
2. **Create a branch**: `git checkout -b feature/my-feature`
3. **Make your changes** and test in a Google Sheet
4. **Commit**: `git commit -m "Add my feature"`
5. **Push**: `git push origin feature/my-feature`
6. **Open a Pull Request**

### Development Tips

- Use [clasp](https://github.com/google/clasp) to develop locally and push to Apps Script
- Test with small datasets first before sending to large lists
- Use the **Send Test Email** feature to verify formatting

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 🤖 Disclaimer

This project was **vibe-coded** — built with the help of AI tools. While the code has been tested and works as intended, it may not follow every traditional software engineering convention. Contributions, code reviews, and improvements from the community are especially welcome!

---

## 🙏 Acknowledgments

Built with Google Apps Script, Gmail API, and Google Drive API.
