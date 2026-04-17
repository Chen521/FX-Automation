# BoC FX Rates — Excel Add-in

Pull live and historical FX rates from the Bank of Canada Valet API directly into Excel.

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Tells Excel about your add-in (sideload this) |
| `taskpane.html` | The full add-in UI + logic (host this on a web server) |
| `commands.html` | Required stub for the ribbon button |

---

## Deployment steps

### 1. Host the HTML files

You need the two HTML files served over **HTTPS** from a web server. Options:

- **GitHub Pages** (free, easy): push files to a repo and enable Pages under Settings → Pages
- **Azure Static Web Apps** (free tier, integrates with Microsoft 365)
- **Any HTTPS web host** — Netlify, Vercel, AWS S3 + CloudFront, etc.

> The BoC API does not require authentication, so no secrets/keys are needed.

### 2. Update the manifest

Open `manifest.xml` and replace every instance of:
```
https://YOUR-HOSTED-URL
```
with your actual hosted URL (e.g. `https://youraccount.github.io/boc-fx-addin`).

### 3. Sideload the add-in into Excel

#### Excel on Windows / Mac (Desktop)
1. Open Excel → **Insert** tab → **Get Add-ins** → **My Add-ins**
2. Click **Upload My Add-in** (or "Manage My Add-ins" → Upload)
3. Browse to and select `manifest.xml`
4. The **FX Rates** button will appear in the **Home** tab ribbon

#### Excel on the Web
1. Go to [office.com](https://office.com) → open any Excel workbook
2. **Insert** → **Add-ins** → **Upload My Add-in**
3. Upload `manifest.xml`

#### Via Microsoft 365 Admin (for org-wide deployment)
1. Microsoft 365 Admin Center → **Settings** → **Integrated Apps**
2. Upload the manifest — it will be available to all users in your tenant

---

## Features

### Single pair tab
- Choose a currency pair (USD/CAD, EUR/CAD, GBP/CAD, etc.)
- Choose a date range (last 10/30/90 days, last year, or custom)
- Insert into the selected cell or a new sheet
- Optional column headers

### Bulk insert tab
- Select multiple pairs at once
- All pairs are inserted side-by-side with a shared date index
- Output goes to a dedicated "BoC_Bulk" sheet

### Live cell tab
- Insert the latest rate into any selected cell
- Option to insert rate + date as a pair (two adjacent cells)
- Refresh button to update the value

---

## Supported currency pairs

| Series code | Pair |
|-------------|------|
| FXUSDCAD | USD / CAD |
| FXEURCAD | EUR / CAD |
| FXGBPCAD | GBP / CAD |
| FXJPYCAD | JPY / CAD |
| FXCHFCAD | CHF / CAD |
| FXAUDCAD | AUD / CAD |
| FXHKDCAD | HKD / CAD |
| FXCNYCAD | CNY / CAD |
| FXMXNCAD | MXN / CAD |
| FXNZDCAD | NZD / CAD |

Add more by finding series codes at: https://www.bankofcanada.ca/valet/lists/series/json

---

## Notes

- The BoC Valet API only publishes **business day** rates — weekends and holidays will have no data.
- Data is typically published by end of day (Eastern time).
- No API key or authentication is required — it's a free public API.
