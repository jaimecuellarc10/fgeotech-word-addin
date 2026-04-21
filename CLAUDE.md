# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Word Office JS task pane add-in for **Fortify Geotech** (ACT Geotechnical Engineers) that pulls project data from the **Total Synergy** project management platform and populates Word document fields via Content Controls.

- Hosted on GitHub Pages: `https://jaimecuellarc10.github.io/fgeotech-word-addin/`
- No build step — plain HTML/CSS/JS, no frameworks, no npm dependencies
- The add-in is sideloaded via `manifest.xml`

## Running Locally

Start the HTTPS server (required for local Word testing):
```bash
node server.js --https
```

First-time only — install dev certs:
```bash
npx office-addin-dev-certs install
```

To sideload into Word on Mac, copy the manifest to the Word add-in folder:
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

## Architecture

**`manifest.xml`** — Office Add-in manifest. Points icon and taskpane URLs to GitHub Pages. Contains the ribbon button definition under `TabHome`.

**`taskpane.html`** — The sidebar UI. Two sections: Settings (org slug + API key, persisted in `localStorage`) and Project Details (fields populated after a Load). Each field label shows its Content Control tag name for template setup reference.

**`taskpane.js`** — All client-side logic:
- `FIELD_MAP` — maps sidebar input IDs → Word Content Control tag names → Total Synergy API response paths (dot-notation for nested fields)
- `loadProject()` — calls Total Synergy API directly from the browser (`access-control-allow-origin: *` confirmed), unwraps `data.items[0]`
- `applyToDocument()` — uses `Word.run()` to find Content Controls by tag and replace their text
- Settings (API key, org slug) stored in `localStorage`

**`server.js`** — Only needed for local development. Serves static files over HTTPS and includes a `/proxy/projects` endpoint that was used before direct browser calls were confirmed to work. Not used in production.

**`start-fgeotech.bat`** — Windows double-click launcher for the local server (for colleagues who need to run locally).

## Total Synergy API

- Base URL: `https://api.totalsynergy.com/api/v2/`
- Auth header: `access-token: {apiKey}`
- Project search endpoint: `GET /api/v2/Organisation/{orgSlug}/Projects?criteria.projectNumber={number}`
- Returns `{ items: [...], totalItems: N }` — always unwrap `items[0]`
- Org slug for this company: `actgeotechnicalengineers`
- API keys are per-user, generated from the Total Synergy profile menu

## Word Content Controls

Templates must have **Plain Text Content Controls** tagged with names from `FIELD_MAP` (e.g. `synergy_project_number`, `synergy_client_name`). Set via Word Developer tab → Insert Content Control → Properties → Tag.

## Brand Colours

- Dark forest green: `#1D3B2A` (header, primary buttons, focus outlines)
- Sage green: `#8A9E78` (Apply button)
