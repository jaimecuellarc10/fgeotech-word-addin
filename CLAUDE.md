# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Word Office JS task pane add-in for **Fortify Geotech** (ACT Geotechnical Engineers) that pulls project data from the **Total Synergy** project management platform and populates Word document fields via Content Controls.

- Hosted on GitHub Pages: `https://jaimecuellarc10.github.io/fgeotech-word-addin/`
- No build step — plain HTML/CSS/JS, no frameworks, no npm dependencies
- The add-in is sideloaded via `manifest.xml` registered in the Windows registry

## Deployment (Windows)

**For colleagues:** run `install-fgeotech.bat` — it writes the registry key that makes Word pick up `manifest.xml` from the same folder.

Registry key written:
```
HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer  →  "FGeotech" = <path to manifest.xml>
```

**Cache clearing** (required after every code update to force Word to reload the add-in):
```bash
rm -rf "$LOCALAPPDATA/Microsoft/Office/16.0/Wef"
```
Word must be fully closed before clearing the cache.

## Running Locally

Start the HTTPS server (required for local Word testing):
```bash
node server.js --https
```

First-time only — install dev certs:
```bash
npx office-addin-dev-certs install
```

## Architecture

**`manifest.xml`** — Office Add-in manifest. Points icon and taskpane URLs to GitHub Pages. Contains the ribbon button definition under `TabHome`.

**`taskpane.html`** — The sidebar UI. Two sections: Settings (API key, persisted in `localStorage`) and Project Details (fields populated after a Load). Each field label shows its Content Control tag name for template setup reference. Script is versioned (`taskpane.js?v=N`) — bump N on every release to bust the WebView2 cache.

**`taskpane.js`** — All client-side logic:
- `FIELD_MAP` — maps sidebar input IDs → Word Content Control tag names → Total Synergy API response paths (dot-notation for nested fields). Fields with `apiPath: null` are either manual or computed.
- `loadProject()` — calls Total Synergy API directly from the browser, unwraps `data.items[0]`. Also fetches client email from the Contacts endpoint using `project.primaryContactId`.
- `populateFields()` — fills API-backed inputs and computes `synergy_project_full_address` by joining street + suburb + state + postcode.
- `applyToDocument()` — reads body OOXML, manipulates it directly, writes it back (see critical note below).
- `updateSdtByTag()` / `findClosingTag()` — OOXML helpers used by `applyToDocument`.
- Settings (API key) stored in `localStorage`.

**`server.js`** — Only needed for local development. Not used in production.

**`start-fgeotech.bat`** — Windows double-click launcher for the local server.

## Critical: Why OOXML Manipulation Instead of contentControls API

The `Word.run context.document.contentControls` API returns **0 items** for this project's templates. Root cause: the content controls in these documents are **table-cell-level SDTs** (`<w:sdt>` elements that are direct children of `<w:tr>`, wrapping `<w:tc>` elements). The Office JS `contentControls` collection only enumerates paragraph-level and inline SDTs — it silently skips row/cell-wrapping SDTs.

**Do not revert to `contentControls.load()`** — it will never work for these documents.

The working approach (`applyToDocument`):
1. `context.document.body.getOoxml()` — get the full body XML (~5.8 MB for the standard template)
2. Globally strip `<w:showingPlcHdr/>` and `<w:rStyle w:val="PlaceholderText"/>` so filled values display instead of placeholder text
3. For each field, call `updateSdtByTag()` which finds every `<w:sdt>` with matching `w:val="tagName"` and replaces text inside its `<w:sdtContent>`
4. `context.document.body.insertOoxml(xml, "Replace")` — write the modified XML back

`updateSdtByTag` loops through **all occurrences** of the same tag (a tag can appear multiple times in the document). It uses `findClosingTag()` for nesting-aware `</w:sdtContent>` matching. If an SDT's content has no `<w:t>` element (empty paragraph), it injects a `<w:r><w:t>` run before `</w:p>`.

## Total Synergy API

- Base URL: `https://api.totalsynergy.com/api/v2/`
- Auth header: `access-token: {apiKey}`
- Project search: `GET /api/v2/Organisation/actgeotechnicalengineers/Projects?criteria.projectNumber={number}`
- Returns `{ items: [...], totalItems: N }` — always unwrap `items[0]`
- Only **active** projects are returned. Completed/on-hold projects return empty results — the UI shows a clear error message for this case.
- Client email: `GET /api/v2/Organisation/actgeotechnicalengineers/Contacts/{primaryContactId}` — check `contact.email`, `contact.emailAddress`, or `contact.emails[0]`
- Org slug is hardcoded: `actgeotechnicalengineers`
- API keys are per-user, generated from the Total Synergy profile menu

## Word Content Controls — Tag Reference

Templates must use **Plain Text Content Controls** (Developer tab → Controls → Plain Text Content Control `Aa`). Set the Tag via Properties dialog. Tags are case-insensitive at apply time but should be lowercase for consistency.

| Tag | Source |
|-----|--------|
| `synergy_project_number` | API |
| `synergy_project_name` | API |
| `synergy_project_status` | API (read-only display) |
| `synergy_client_name` | API |
| `synergy_client_contact` | API |
| `synergy_client_email` | API (Contacts endpoint) |
| `synergy_project_manager` | API |
| `synergy_project_address` | API (street only) |
| `synergy_project_suburb` | API |
| `synergy_project_state` | API |
| `synergy_project_postcode` | API |
| `synergy_project_full_address` | Computed: street, suburb, state, postcode joined with `, ` |
| `synergy_project_office` | API |
| `synergy_report_writer` | Manual |
| `synergy_report_reviewer` | Manual |
| `synergy_investigation_type` | Manual |

## Brand Colours

- Dark forest green: `#1D3B2A` (header, primary buttons, focus outlines)
- Sage green: `#8A9E78` (Apply button)
