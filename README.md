# Report Automation System - Dadar Branch

Automated email reading from Gmail → AI parsing (Gemini) → Data writing to Google Sheets.

## 🏢 Branch Information

| Detail | Value |
|--------|-------|
| **Branch Name** | Dadar |
| **Gmail Account** | hardik.misedujam@gmail.com |
| **Google Sheet** | Dadar Sales Daily Working Report |
| **Department** | Sales Only (No HR) |
| **Schedule** | Every 30 min, Mon-Sat, 3:00 PM to 9:00 PM IST |

## 📋 System Overview

This automation system:
- Reads unread emails from specified sales employees
- Parses email content using Gemini AI (text + vision)
- Extracts: Total Dialed, Total Connected, Duration, Prospect
- Writes data to Google Sheets
- Marks late submissions (after 9:00 PM IST) as "Not Sent"
- Runs via GitHub Actions on a schedule

## 🛠️ Technologies Used

| Component | Purpose |
|-----------|---------|
| Gmail API (OAuth 2.0) | Read unread emails from allowed senders |
| Gemini API | Parse email body + screenshots (text + vision) |
| Google Sheets API (Service Account) | Write extracted data to spreadsheets |
| GitHub Actions | Schedule and run the automation |
| Tesseract OCR | Fallback for screenshot parsing |

## 📁 File Structure


Report_Automation_Dadar/
├── .github/
│ └── workflows/
│ └── scheduler.yml ← GitHub Actions schedule
├── Working_Report_Editor/
│ ├── config.py ← Branch configuration (employees, emails)
│ ├── config.yaml ← Model configuration
│ ├── error_handler.py ← Error handling & retry logic
│ ├── gemini_parser.py ← Email body parser (Gemini + regex)
│ ├── gmail_reader.py ← Gmail email fetcher (OAuth)
│ ├── main.py ← Main orchestrator
│ ├── requirements.txt ← Python dependencies
│ ├── runtime.txt ← Python version
│ ├── sheets_service.py ← Google Sheets writer
│ ├── tracker.py ← Duplicate detection & logging
│ ├── utils.py ← Helper functions
│ ├── validator.py ← Data validation
│ └── vision_parser.py ← Screenshot parser
├── .gitignore ← Git ignore rules
├── .python-version ← Python version
└── README.md ← Documentation




## 👥 Employee List

| Employee Name | Email Address |
|---------------|---------------|
| Sanskruti | sanskruti.edujam@gmail.com |
| Naved | naveds.edujam@gmail.com |
| Krishna | krishna.edujam@gmail.com |
| Abhinav | abhinav.edujam@gmail.com |
| Rohit | rohitg.edujam@gmail.com |
| Fahad | fahadk.edujam@gmail.com |
| Sarvesh | sarvesh.edujam@gmail.com |
| Shrushti | shrushtir.edujam@gmail.com |
| Imzamum | imzamium.edujam@gmail.com |

## 📊 Google Sheets Columns

| Column | Position |
|--------|----------|
| Date | 1 |
| Employee Name | 2 |
| Total Dialed | 3 |
| Total Connected | 4 |
| Duration | 5 |
| Prospect | 6 |
| Ref Added | 7 |
| Status Viewed | 8 |
| Document Collected | 9 |
| Report Status | 10 |

### Status Column Values

| Value | Meaning |
|-------|---------|
| (blank) | ✅ Data written successfully |
| Not Sent | ⚠️ Employee didn't submit OR submitted after 9:00 PM |

## 🕐 Duration Formats Supported

| Format | Example | Output |
|--------|---------|--------|
| Standard (HH:MM:SS) | 01:28:52 | 01:28:52 |
| Dots (2-digit hour) | 02.07.36 | 02:07:36 |
| Dots (1-digit hour) | 2.08.32 | 02:08:32 |
| Single digit seconds | 01:28:0 | 01:28:00 |
| Text (h/m/s) | 1h 42m 8s | 01:42:08 |
| sec as seconds | 1h 42m 8sec | 01:42:08 |
| min as minutes | 1hr 25min 46s | 01:25:46 |
| MINS SEC (uppercase) | 49 MINS 9 SEC | 00:49:09 |
| hr min sec (full words) | 1hr 9min 47sec | 01:09:47 |
| MM:SS format | 58:14 | 00:58:14 |
| Addition patterns | 1h 15m + 10 min | 01:25:00 |

## 🔐 GitHub Secrets Required

| Secret Name | Purpose |
|-------------|---------|
| `GOOGLE_CREDENTIALS` | Service Account JSON |
| `GEMINI_API_KEY_DADAR` | Gemini API key |
| `DADAR_SALES_SPREADSHEET_ID` | Branch-specific Sheet ID |
| `DADAR_CLIENT_ID` | Gmail OAuth client ID |
| `DADAR_CLIENT_SECRET` | Gmail OAuth client secret |
| `DADAR_REFRESH_TOKEN` | Gmail OAuth refresh token |

## ⏰ Schedule

The automation runs on the following schedule (IST):

| Time (IST) | UTC | Days |
|------------|-----|------|
| 3:00 PM - 8:30 PM | 09:30 - 15:00 | Monday - Saturday |
| 9:00 PM (Final) | 15:30 | Monday - Saturday |

**Cron Expression:** `0,30 9-14 * * 1-6` and `30 15 * * 1-6`

## 🚀 Deployment

### 1. Create Repository
```bash
git clone https://github.com/YOUR_USERNAME/Report_Automation_Dadar.git
cd Report_Automation_Dadar


2. Add GitHub Secrets
Go to Settings → Secrets and variables → Actions and add all required secrets.

3. Push Code

git add .
git commit -m "Initial commit - Dadar branch automation"
git push origin main

4. Test Workflow
Go to Actions tab

Select Report Automation Scheduler - Dadar

Click Run workflow → Run workflow

🔧 Setup Instructions
Getting OAuth Credentials
Go to Google Cloud Console

Create a new project or select existing

Enable Gmail API

Configure OAuth Consent Screen

Create OAuth Client ID (Desktop app)

Download credentials and run get_refresh_token.py

Copy CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN

Getting Service Account Credentials
Go to Google Cloud Console

IAM & Admin → Service Accounts

Create Service Account with Editor role

Generate JSON key

Copy entire JSON as GOOGLE_CREDENTIALS secret

Sharing Google Sheet
Open your Google Sheet

Click Share

Add the client_email from service account JSON

Set permission to Editor

📝 Quick Reference Commands

# Run locally (for testing)
cd Working_Report_Editor
python main.py

# Check logs
cat Working_Report_Editor/logs/processing_logs.csv

🐛 Common Errors & Solutions
Error	Solution
No OAuth credentials found	Check CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN secrets
PermissionError	Share sheet with Service Account client_email (Editor access)
No valid credentials found	Check GOOGLE_CREDENTIALS secret has full JSON
Sheet ID length: 0	Secret not being passed - check scheduler.yml env section
503 Service Unavailable	Retry logic will handle with exponential backoff
📚 Related Repositories
Report_Automation_Thane

Report_Automation_Thane_New

Report_Automation_Andheri

📄 License
This project is proprietary and confidential.

For support, contact: hardik.misedujam@gmail.com


---

## ✅ Quick Copy Instructions

1. **Select all** the text above (from `# Report Automation System - Dadar Branch` to the end)
2. **Copy** (Ctrl+C or Cmd+C)
3. Go to your `README.md` file in the Dadar branch repository
4. **Paste** (Ctrl+V or Cmd+V)
5. **Save** and **commit** the file

---

## 🔧 Customization Needed

| Placeholder | Replace With |
|-------------|--------------|
| `YOUR_USERNAME` | Your actual GitHub username |
| `hardik.misedujam@gmail.com` | Your actual support email |

---

**Now your Dadar branch is complete!** 🚀

