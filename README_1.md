# TRYBE-G Website

A single-page site for the TRYBE-G youth movement — built with HTML, Tailwind CSS, and vanilla JavaScript. No build step required.

## Files

| File | What it's for |
|---|---|
| `index.html` | The whole website — markup, styling, and behavior in one file. |
| `google-apps-script.gs` | Backend script for the **Profile Sync** feature, to paste into Google Apps Script. |
| `README.md` | This file. |

## Running it

Just open `index.html` in a browser, or upload it to any static host (Netlify, GitHub Pages, Vercel, etc.). It loads Tailwind and fonts from a CDN, so it needs an internet connection to look right — there's nothing to install or compile.

## What's a placeholder right now

Open `index.html` and find the `CONFIG` object near the top of the `<script>` block at the bottom of the file. Everything you need to plug in lives there:

```js
const CONFIG = {
  WHATSAPP_LINK: "https://chat.whatsapp.com/REPLACE_WITH_YOUR_INVITE_LINK",
  CONTACT_EMAIL: "hello@trybe-g.org",
  VENUE_TEXT: "Trybe Hub · Replace with your venue address",
  SOCIALS: {
    instagram: "https://instagram.com/REPLACE_WITH_HANDLE",
    tiktok: "https://www.tiktok.com/@REPLACE_WITH_HANDLE",
    youtube: "https://youtube.com/@REPLACE_WITH_HANDLE"
  },
  MEETING: { dayOfWeek: 5, hour: 18, minute: 0, label: "Every Friday · 6:00 PM · Trybe Hub" },
  USE_MOCK_DATA: true,
  SHEETS_API_URL: "https://script.google.com/macros/s/REPLACE_WITH_YOUR_DEPLOYMENT_ID/exec"
};
```

- **WHATSAPP_LINK** — your actual WhatsApp group invite link.
- **CONTACT_EMAIL / VENUE_TEXT** — shown in the footer.
- **SOCIALS** — your real Instagram, TikTok, and YouTube URLs (also update the three `href`s in the footer's social icon markup to match).
- **MEETING** — `dayOfWeek` is `0` for Sunday through `6` for Saturday. The countdown always counts down to the *next* occurrence of that day/time, so you never have to update it manually — it just keeps working week after week.

## Setting up the real Google Sheets sync

Right now `USE_MOCK_DATA` is `true`, so the Profile Sync panel works instantly against three fake demo records (try searching `ada@trybeg.org`) — no backend needed. That's what's running in the live preview.

To connect it to your real Google Sheet:

1. Open the Sheet you want to use as your member list.
2. Set up row 1 with these exact headers, in this order: `Timestamp | Name | Email | Phone | Day | Month | Year`
3. In the Sheet, go to **Extensions → Apps Script**.
4. Delete the starter code and paste in everything from `google-apps-script.gs`.
5. Click **Deploy → New deployment** → type **Web app** → Execute as **Me** → Who has access **Anyone**.
6. Copy the deployment URL (it ends in `/exec`).
7. Back in `index.html`, set:
   ```js
   USE_MOCK_DATA: false,
   SHEETS_API_URL: "https://script.google.com/macros/s/YOUR_ID_HERE/exec"
   ```

That's it — the search box now looks up real rows by email or phone, and "Save Changes" writes back to the same row.

If your real sheet has different column names or extra fields, `google-apps-script.gs` is short and commented — adjust `HEADERS` and `rowToMember_()` to match.

## Customizing the design

- **Colors, fonts, and the chamfered "cut corner" edges** are defined near the top of `index.html`, inside the `tailwind.config` script block and the `<style>` block right after it.
- The **TRYBE-G mark** used throughout the page (nav, hero watermark, identity diagram, footer) is embedded once as an SVG `<symbol id="mark">` near the top of `<body>`, then reused everywhere with `<use href="#mark">`. Change its color anywhere just by changing the surrounding element's text color class (it inherits via `fill="currentColor"`).
- Demo member records for the mock Sheet sync live in the `MOCK_MEMBERS` array in the JS — edit or remove them once you switch to the real backend.

## Browser support

Built with plain CSS/JS and Tailwind's utility classes — works in any current browser. Respects `prefers-reduced-motion` for the ambient rotation and scroll-reveal animations.
