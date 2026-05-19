# Security Review & Setup
## GloBE XML Export App — globe-xml-export.streamlit.app

**B2B Accounting AG**  
Prepared for client onboarding · May 2026

---

## Summary

The GloBE XML Export App is a browser-based tool hosted on Streamlit Community Cloud (Snowflake infrastructure). It converts the Swiss QDMTT Excel calculation template into a GloBE Information Return (GIR) XML file for submission to the ESTV.

**Overall risk: Low** — no data is stored, no database, no user accounts on the application side.

---

## Data Flow

```
Client browser → Streamlit Cloud (Snowflake) → XML returned to browser
```

1. User uploads the Excel file from their browser
2. The file is processed in memory on the Streamlit server
3. The generated XML is sent back to the browser as a download
4. **Nothing is written to disk or stored after the session ends**

---

## What Data Is Processed

| Data | Where it goes | Retained? |
|---|---|---|
| Excel file (financial figures) | Server memory during session | No |
| Generated XML file | Returned to browser | No |
| Company name, TIN | Server memory during session | No |
| IP address / browser info | Streamlit / Snowflake logs | Standard server logs only |

---

## Infrastructure Security

| Item | Detail |
|---|---|
| Hosting | Streamlit Community Cloud (Snowflake) |
| Data residency | US-based Snowflake infrastructure |
| Transport encryption | HTTPS / TLS 1.2+ enforced |
| Server-side storage | None — stateless, in-memory processing only |
| Source code | Private GitHub repository (B2B-Accounting-AG org) |
| Authentication | Email whitelist (see Access Control below) |

---

## Access Control

The app is set to **Private**. Access is restricted to email addresses explicitly approved by B2B Accounting AG.

### How to grant access to a new user

1. Go to **share.streamlit.io**
2. Open the app settings for `globe-xml-export`
3. Navigate to **Sharing**
4. Enter the user's email address and click **Invite**

The user will receive an email invitation and must sign in with their GitHub or Google account to access the app.

### How to revoke access

1. Go to **share.streamlit.io** → app settings → **Sharing**
2. Remove the user's email address

---

## Recommendations Before Going Live

| # | Action | Priority |
|---|---|---|
| 1 | Set app to Private and whitelist client email addresses | High |
| 2 | Validate final XML against official ESTV XSD when published | High |
| 3 | Confirm with client that uploading financial data to a US-hosted server is permitted under their data governance policy | Medium |
| 4 | Rotate the GitHub Personal Access Token that was exposed during setup | High |
| 5 | Advise client not to share the app URL publicly | Medium |

---

## Limitations

- **No audit trail** — the app does not log which files were processed or by whom
- **No version control on submissions** — clients should save the downloaded XML locally
- **US data residency** — Streamlit Community Cloud runs on US Snowflake infrastructure; if the client requires EU/CH data residency, a self-hosted deployment on a Swiss server would be required
- **Uptime** — Streamlit Community Cloud is a free-tier service with no SLA; for mission-critical use, a paid deployment is recommended

---

## Self-Hosted Alternative

If the client requires Swiss data residency or a guaranteed SLA, the app can be deployed on a dedicated server (e.g. a Swiss VPS) with:

```bash
pip install streamlit openpyxl
streamlit run globe_xml_app.py --server.port 443
```

B2B Accounting AG controls the source code and can redeploy at any time.

---

## Source Code & Change Management

All changes to the app go through the GitHub repository at:  
**github.com/B2B-Accounting-AG/globe-xml-export**

Only B2B Accounting AG has write access. The client receives the app URL only — they have no access to the source code.

---

*Document prepared by B2B Accounting AG · dan@danbuetler.com*
