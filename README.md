# My Notes – Outlook add-in

This repository hosts **only** the My Notes add-in for Outlook.
CDF Updater (Excel) lives in its own repository: NPerryK/cdf-updater.

| File | Purpose |
|---|---|
| `my-notes.html`, `my-notes.js` | The add-in |
| `assets/mn-icon-*.png` | Icons for the new manifest |
| `taskpane.html`, `assets/icon-*.png` | Kept so the older manifest keeps working; `taskpane.html` forwards to `my-notes.html` |

`cdf-link.js` reads CDF status and Notes from column E of the CDF workbook and exports Outlook notes for the Excel lookup. On the Reports tab, open **CDF status** to paste an Excel snapshot or copy Outlook notes.
