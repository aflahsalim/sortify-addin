# Sortify Outlook Add‑in

Sortify is an email‑safety classification system designed for Outlook (Web & Desktop).  
It helps users quickly understand whether an email is Safe, Suspicious, Spam, or Phishing — directly inside the Outlook task pane.

The system has two components:
1. Sortify Outlook Add‑in (this repository)
2. Sortify Backend (optional ML inference service)

This repository contains the add‑in component.

# What is Sortify?

Sortify is a privacy‑respecting email‑risk detection tool built for Outlook.  
It analyzes incoming emails using:
- Local heuristics (built into the add‑in)
- Optional machine‑learning inference (backend)

Sortify displays:
- A risk gauge (0–100)
- A color‑coded category (Safe, Support, Spam, Phishing)
- A short analysis panel (links, attachments, urgency, sender reputation)
- A confirmation popup for forwarding suspicious emails to support

Sortify does NOT store email content, log messages, or send full emails to external servers.

# Role of This Repository (Add‑in)

This add‑in:
- Runs inside Outlook (Web & Desktop)
- Reads email metadata using Office.js
- Computes a risk score using heuristics or backend
- Displays the UI (gauge + analysis panel)
- Provides a forwarding popup for escalation

# Features

- Gauge showing risk percentage and color zone
- Four categories: Safe, Support, Spam, Phishing
- Analysis panel with:
  - Sender reputation (Trusted / Unknown)
  - Link presence
  - Attachment presence
  - Urgency level
- Optional backend inference
- Confirmation popup for forwarding emails
- Graceful fallback values when data is missing

# Installation & Setup

1. Download the manifest file:
manifest.xml

2. Install in Outlook Web:
- Open Outlook in browser
- Settings → View all Outlook settings
- Mail → Customize actions → Add‑ins
- Add a custom add‑in → Add from file
- Upload manifest.xml

3. Install in Outlook Desktop:
- Home → Get Add‑ins
- My Add‑ins
- Custom Add‑ins → Add from file
- Upload manifest.xml

Sortify will appear in the Outlook ribbon and task pane.

# How to Use

1. Open any email
2. Open the Sortify task pane
3. Sortify automatically:
   - Reads email metadata
   - Computes a local heuristic score
   - Or calls the backend (if enabled)
4. Gauge + analysis panel appear instantly
5. Click “Forward to Support” to escalate
6. Confirm in the popup

Sortify does NOT store or log email content.

# Optional: Enable Backend Mode

1. Run the Sortify backend locally
2. Open the add‑in settings
3. Enter backend URL:
http://localhost:8000/api/infer
4. Save settings

If backend is unavailable, Sortify automatically falls back to local heuristics.

# Technical Details

Local Heuristics:
- Link detection
- Attachment detection
- Urgency keywords
- Sender domain trust
- Suspicious phrases
- Basic text features

UI Components:
- Gauge (0–100)
- Color zones:
  Green = Safe
  Blue = Support
  Orange = Spam
  Red = Phishing
- Analysis panel
- Confirmation popup

# Privacy & Security

- No email content stored
- No logs
- Only minimal metadata processed
- Forwarding requires explicit confirmation
- Add‑in cannot auto‑forward emails

# Limitations

- Sender reputation simplified to Trusted/Unknown
- No attachment scanning
- No persistent logs
- Backend optional
- No automated model retraining

# Purpose

To help users quickly understand the safety of an email by providing a clear, compact, and privacy‑respecting risk assessment directly within Outlook.
