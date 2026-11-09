# My Notes – Outlook add-in

This repository hosts **only** the My Notes add-in for Outlook.
CDF Updater (Excel) lives in its own repository: NPerryK/cdf-updater.

| File | Purpose |
|---|---|
| `my-notes.html`, `my-notes.js` | The add-in |
| `assets/mn-icon-*.png` | Icons for the new manifest |
| `taskpane.html`, `assets/icon-*.png` | Kept so the older manifest keeps working; `taskpane.html` forwards to `my-notes.html` |

Updates normally change `my-notes.html` and `my-notes.js` only.
