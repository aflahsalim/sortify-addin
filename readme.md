# Sortify Outlook Add‑in

Sortify is an Outlook add‑in that evaluates the safety of incoming emails and presents the result directly inside the Outlook task pane.  
It uses local heuristics and, optionally, a Python backend to classify emails into four categories: Safe, Support, Spam, and Phishing.

## Features
- Semicircular gauge showing risk percentage and color‑coded category
- Analysis panel displaying:
  - Sender reputation (Trusted / Unknown)
  - Link presence
  - Attachment presence
  - Urgency level
- Optional backend inference for ML‑based scoring
- Confirmation popup for forwarding emails to Sortify Support
- Graceful fallback values when data is missing

## How It Works
1. User opens an email in Outlook  
2. Add‑in retrieves metadata via Office.js  
3. Add‑in computes a local heuristic score or calls the backend  
4. Gauge and analysis panel are rendered  
5. User may choose to forward the email to support via a confirmation popup  

## Tech Stack
- HTML, CSS, JavaScript  
- Office.js (Microsoft Office JavaScript APIs)  
- Optional backend integration via REST API

## Requirements & Constraints
- Works inside Outlook (web and desktop)  
- No email content is stored or logged  
- Backend inference is optional and may be disabled  
- Add‑in must remain responsive and lightweight  

## Purpose
To help users quickly understand the safety of an email by providing a clear, compact, and privacy‑respecting risk assessment directly within Outlook.
