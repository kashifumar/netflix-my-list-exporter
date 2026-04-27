# Netflix My List Exporter — Firefox Extension

> Export every title in your Netflix **My List** to a richly formatted **Excel (.xlsx)** workbook — with embedded poster images, cast, genres, maturity ratings, season breakdowns, and a summary sheet — all from a single click.

---

## Table of Contents

1. [Overview](#overview)
2. [Key Features](#key-features)
3. [Tech Stack](#tech-stack)
4. [Architecture](#architecture)
5. [How It Works — Step by Step](#how-it-works--step-by-step)
6. [Engineering Challenges Solved](#engineering-challenges-solved)
7. [Excel Output Format](#excel-output-format)
8. [Project Structure](#project-structure)
9. [Installation](#installation)
10. [Usage](#usage)
11. [Permissions](#permissions)
12. [Known Limitations](#known-limitations)
13. [Author](#author)

---

## Overview

Netflix has no native export feature. For users who manage large watchlists — or anyone cataloguing content across a household — there's no way to get a structured, shareable view of their list outside the browser UI.

This extension solves that by reverse-engineering Netflix's internal data layer and extracting full metadata for every title in "My List", then generating a production-quality Excel workbook entirely client-side, with no backend and no data leaving the user's machine.

**Target browser:** Firefox (Manifest V2)

---

## Key Features

| Feature | Detail |
|---|---|
| **One-click export** | Toolbar button or right-click context menu on any Netflix page |
| **Dual extraction modes** | React fiber (fast, internal) + JawBone detail panel (rich, reliable) with automatic fallback |
| **Full metadata per title** | Title, Type, Language, Maturity Rating, Maturity Description, Rating Reason, Description, Cast, Genres, Season count, Episodes per season |
| **Embedded poster images** | Netflix CDN posters fetched and embedded directly into Excel cells |
| **Two-sheet workbook** | Sheet 1: full title catalogue with images; Sheet 2: summary statistics |
| **Popup-close resilience** | Extraction continues in the background if the popup is closed; state restored on reopen |
| **Profile session guard** | Detects Netflix profile-picker timeouts mid-extraction and reports gracefully |
| **Retry failed titles** | Collapsible failed-titles list with a one-click retry |
| **Zero external dependencies at runtime** | All processing is local; no data sent to any server |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Extension platform | Firefox WebExtension — Manifest V2 |
| Content script | Vanilla JavaScript (ES2020) — injected into `*.netflix.com` |
| Background script | MV2 non-persistent event page |
| Popup | HTML5 + CSS3 + Vanilla JS |
| Excel generation | [ExcelJS](https://github.com/exceljs/exceljs) v4.4.0 (browser bundle, MIT licence) |
| Inter-component communication | `browser.runtime` Ports (long-lived bidirectional channels) |
| Image fetching | `fetch()` in background scope to bypass popup-origin CORS |
| State persistence | `browser.storage.local` — survives popup close/reopen |
| Internationalisation | Firefox i18n API (`_locales/en/messages.json`) |

**Why ExcelJS over SheetJS?** SheetJS's image-embedding feature (`!images`) is locked behind a commercial Pro licence. ExcelJS provides equivalent functionality — including `wb.addImage()`, cell anchoring, row heights, frozen headers, and auto-filter — for free under the MIT licence and ships a self-contained browser bundle that loads as a plain `<script>` tag, which is exactly what a Firefox extension popup requires under MV2's strict Content Security Policy.

---

## Architecture

```
┌───────────────────────────────────────────────────────────────────────┐
│  Firefox Extension                                                    │
│                                                                       │
│  ┌──────────────────┐   Port 'popup-bg'    ┌───────────────────────┐ │
│  │    popup.js      │◄────────────────────►│    background.js      │ │
│  │                  │                      │                       │ │
│  │  • UI state      │  portRequest(action) │  • Extraction store   │ │
│  │  • ExcelJS build │◄──────────────────── │  • Image fetch proxy  │ │
│  │  • Download cmd  │ ──────────────────►  │  • Download API       │ │
│  │  • State restore │                      │  • Content relay      │ │
│  └──────────────────┘                      └──────────┬────────────┘ │
│                                                        │              │
│                                         Port 'content-bg'            │
│                                                        │              │
│                                            ┌───────────▼───────────┐ │
│                                            │     content.js        │ │
│                                            │                       │ │
│                                            │  • DOM scraping       │ │
│                                            │  • React fiber walk   │ │
│                                            │  • JawBone scrape     │ │
│                                            │  • falcorCache parse  │ │
│                                            │  • Progress stream    │ │
│                                            └───────────────────────┘ │
│                                            (runs inside Netflix tab)  │
└───────────────────────────────────────────────────────────────────────┘
```

### Three-component design

**`content.js`** — runs inside the Netflix tab under Firefox's sandboxed content-script scope. Responsible for all DOM interaction, data extraction, and progress streaming. Communicates upstream via the `content-bg` named port.

**`background.js`** — the persistent event page. Acts as the message relay between the popup and the content script, proxies image fetches (to avoid CORS), handles `browser.downloads`, and maintains a durable extraction state store in `browser.storage.local`.

**`popup.js`** — the user interface layer. Connects to the background via the `popup-bg` named port on open. Receives a `resumeState` message immediately so it can restore UI without latency. Builds the ExcelJS workbook and triggers the download entirely in-browser.

### Why the two-port relay?

In Firefox MV2, a popup page cannot receive messages directly from a content script because the popup's `runtime.onMessage` context has a different lifetime than the tab. The background script acts as a broker: the content script streams to `background.js` via `content-bg`, which persists state and forwards events to the popup via `popup-bg` if it is open. This also means the popup can close and reopen without losing extraction data.

---

## How It Works — Step by Step

### Phase 1 — Initialisation

1. User clicks the toolbar icon (or right-clicks and selects "Export Netflix My List to Excel")
2. Popup opens and immediately connects to `background.js` via the `popup-bg` port
3. Background sends the last-known `extractionStore` as a `resumeState` message — popup restores UI if extraction was already in progress or completed
4. Popup queries the active tab for profile name and current card count via the `sendToContent` relay

### Phase 2 — Card loading

5. Content script navigates to `netflix.com/browse/my-list` if not already there
6. `scrollToLoadAllCards()` scrolls to the bottom of the infinite-scroll grid repeatedly (up to 80 attempts) until no new cards appear across 4 consecutive scrolls
7. All title-card DOM elements are collected via a priority-ordered list of CSS selectors that covers every known Netflix UI variant

### Phase 3 — Metadata extraction (Mode: Internal API)

8. For each card, `extractFromFiber()` walks the React fiber tree upward from the card DOM element
   - The fiber key (`__reactFiber$XXXX`) is found by iterating `Object.keys()` on the unwrapped element — Firefox's XrayWrapper must be bypassed first with `el.wrappedJSObject`
   - Each fiber node is unwrapped individually before reading `memoizedProps`
   - When `props.videoModel` is found, it is converted to a `TitleRecord` via `videoModelToRecord()`
   - Contains: title, type (Movie/Series), maturity rating, season count
9. If `videoModel` had no description, `openAndScrapeJawBone()` is called:
   - Pushes `?jbv=<videoId>` to the URL via `history.pushState`
   - Dispatches a `popstate` event on the real page window (`window.wrappedJSObject`) to trigger Netflix's SPA router
   - Waits up to 6 s for `.ptrack-container` to appear
   - Scrapes `.previewModal--detailsMetadata-left` (description) and `.previewModal--detailsMetadata-right` (cast, genres)
   - Dismisses the panel via `Escape` keydown and restores the original URL
10. As a final fallback, `falcorCache.videos[videoId]` is read from `window.netflix.falcorCache` if the fiber walk returned nothing

### Phase 4 — Image fetching

11. Popup requests each poster URL from `background.js` via `portRequest('fetchImage', { url })`
12. Background fetches the image from Netflix's CDN using `fetch()` in the background scope (no CORS issue here), encodes it to a base64 data URI, and returns it to the popup

### Phase 5 — Excel generation

13. `buildExcelWorkbook()` creates an ExcelJS workbook in the popup
14. `buildMyListSheet()` writes all 13 columns, applies Netflix-red header fill, alternating row colours, frozen header row, auto-filter, and embeds each poster using `wb.addImage()` + `ws.addImage()` anchored with `tl`/`br` cell coordinates
15. `buildSummarySheet()` writes aggregate statistics
16. `wb.xlsx.writeBuffer()` serialises the workbook to an `ArrayBuffer`

### Phase 6 — Download

17. The popup sends the raw `ArrayBuffer` to the background via `portRequest('downloadBuffer', { buffer, filename })`
18. Background wraps the buffer in a `Blob`, creates a persistent `blob:` URL, and calls `browser.downloads.download({ saveAs: true })`
19. The OS "Save As" dialog appears; after the user saves, the background revokes the blob URL after 60 seconds

---

## Engineering Challenges Solved

### 1. Firefox XrayWrapper isolation

Firefox runs content scripts in a sandboxed scope where page-JS properties — including `window.netflix`, React fiber keys (`__reactFiber$XXXX`), and `window.wrappedJSObject` — are invisible. Every access to page-side data goes through two helpers:

```js
function getPageWindow() {
  try { return window.wrappedJSObject || window; } catch (_) { return window; }
}
function unwrapEl(el) {
  try { return (el && el.wrappedJSObject) || el; } catch (_) { return el; }
}
```

### 2. React fiber tree traversal

Netflix stores rich pre-loaded metadata in a `videoModel` prop on a parent React component above each title card. To find it, the extension walks the fiber tree upward using the `fiber.return` pointer. A subtle but critical constraint: each fiber node must be unwrapped from XrayWrapper before reading `memoizedProps`, but the tree must be advanced using the original (possibly wrapped) `fiber.return` — not `(fiber.wrappedJSObject || fiber).return`. Using the unwrapped node's `.return` silently jumps to a different branch of the tree and misses `videoModel` entirely.

### 3. Trusted-event restriction

Netflix's React event handlers reject synthetic `MouseEvent`s dispatched from content scripts because `event.isTrusted === false`. Direct hover simulation is therefore unreliable. The extension instead uses the `?jbv=<videoId>` URL approach: pushing a URL state triggers Netflix's own SPA router (which is trusted) to open the full detail panel without any mouse simulation.

### 4. Popup-close resilience

A long extraction can take several minutes. If the user closes the popup, all state would normally be lost. The extension solves this with a three-layer persistence strategy:

- Every progress event from `content.js` is written to `browser.storage.local` by `background.js`
- On popup reopen, the background sends the full store as the very first port message (`resumeState`)
- The popup reconstructs its UI — showing live progress, a completed download button, or an error — without missing a beat

### 5. Blob URL lifetime and the OS "Save As" race

Calling `browser.downloads.download({ saveAs: true })` from the popup causes the OS file dialog to steal focus, which makes Firefox close the popup. A `blob:` URL created in the popup context is immediately revoked when the popup page is destroyed — before the user has clicked "Save". The fix: send the raw `ArrayBuffer` to the background script (which is persistent), create the blob URL there, and trigger the download from that long-lived context. The URL survives the OS dialog interaction.

### 6. CORS for Netflix CDN images

The popup's `moz-extension://` origin cannot fetch Netflix CDN images directly — the CDN's CORS policy does not whitelist extension origins. Image fetches are proxied through `background.js`, which has a `*://*.netflix.com/*` host permission and runs in the browser's background context where the CORS restriction does not apply.

### 7. Infinite-scroll grid vs. slider

Netflix's My List page switched from a horizontal paginated slider to a vertical infinite-scroll grid. The slider required clicking a "next" arrow to swap virtual card batches into the DOM; the grid just needs repeated scrolling. The extension handles both layouts: `scrollToLoadAllCards()` scrolls the viewport, while `advanceSlider()` tries 13 known arrow-button selectors before falling back to scroll. The "best selector" strategy in `collectCards()` runs all known card selectors and returns whichever gives the most results.

---

## Excel Output Format

### Sheet 1 — "My List"

| Column | Width | Content |
|---|---|---|
| # | 5 | Row number |
| Title | 30 | Display name |
| Type | 10 | `Movie` or `Series` |
| Language | 14 | Original language |
| Maturity Rating | 14 | e.g. `TV-MA`, `PG-13` |
| Maturity Description | 30 | e.g. `Strong Language, Violence` |
| Rating Reason | 30 | Specific content descriptor |
| Description | 50 | Synopsis |
| Cast | 35 | Main cast, comma-separated |
| Genres | 25 | All genres, comma-separated |
| Seasons | 10 | Season count; `N/A` for movies |
| Episodes per Season | 28 | `S1: 8 eps, S2: 10 eps` (series only) |
| Poster | 18 | Embedded poster image (96 × 144 px) |

- Netflix-red frozen header row with white bold text
- Alternating white / light-grey row colours
- Row height 115 pt to accommodate poster images
- Images anchored `oneCell` — they move with rows but do not resize
- Auto-filter on all columns

### Sheet 2 — "Summary"

Counts of total titles, movies, series; number of failed extractions; active profile name; extraction mode; and export timestamp.

### Filename

`MyList-{ProfileName}-{YYYY-MM-DD}.xlsx`

---

## Project Structure

```
netflix-mylist-plugin-firefox/
│
├── manifest.json              MV2 manifest — permissions, content script, icons
├── background.js              Event page — state store, image proxy, download, relay
├── content.js                 Content script — extraction logic, DOM interaction
│
├── popup/
│   ├── popup.html             400 px popup UI
│   ├── popup.js               Orchestration — progress, ExcelJS, download
│   └── popup.css              Netflix dark-theme stylesheet
│
├── icons/
│   ├── icon-48.png            Toolbar icon
│   └── icon-96.png            High-DPI icon
│
├── _locales/
│   └── en/
│       └── messages.json      All English UI strings (Firefox i18n format)
│
├── libs/
│   └── exceljs.min.js         ExcelJS 4.4.0 browser bundle (local — MV2 CSP)
│
├── generate-icons.js          Node.js helper — renders PNG icons via canvas
└── download-libs.js           Node.js helper — downloads ExcelJS into libs/
```

---

## Installation

### Prerequisites

Node.js ≥ 14 is only needed for the two one-time setup scripts. The extension itself has no build step.

### Step 1 — Download ExcelJS

```bash
node download-libs.js
```

This saves `exceljs.min.js` (~1.1 MB) to `libs/`. Alternatively, download it manually from jsDelivr and place it at `libs/exceljs.min.js`.

### Step 2 — Generate icons

```bash
npm install canvas
node generate-icons.js
```

This writes `icons/icon-48.png` and `icons/icon-96.png`. Alternatively, place any 48×48 and 96×96 PNG files at those paths.

### Step 3 — Load in Firefox (temporary install, no signing required)

1. Open `about:debugging#/runtime/this-firefox`
2. Click **Load Temporary Add-on…**
3. Select `manifest.json` from the project directory
4. The extension appears in the toolbar

> Temporary add-ons are removed on Firefox restart. For a persistent install, sign the extension via addons.mozilla.org or use Firefox Developer Edition with `xpinstall.signatures.required = false`.

### Permanent install (Firefox Developer Edition / Nightly)

1. Zip the project contents (with `manifest.json` at the root)
2. Rename to `netflix-mylist-exporter.xpi`
3. Set `xpinstall.signatures.required = false` in `about:config`
4. Drag the `.xpi` onto a Firefox tab

---

## Usage

1. Open Netflix and log in to the profile you want to export
2. Click the extension icon in the toolbar (or right-click any Netflix page)
3. The popup shows your **profile name** and **title count**
4. Select an extraction mode:
   - **Internal API** (default) — fast; reads Netflix's React data layer directly
   - **Mouseover Simulation** — slower but works even if Netflix restructures its internal data
5. Click **Start Export**
6. Watch the progress bar — the currently processed title is shown beneath it
7. Once complete, click **Download Excel File**
8. The OS "Save As" dialog appears — choose your location and save

If any titles fail, a collapsible list appears with a **Retry Failed** button that re-attempts extraction using the fallback mode.

---

## Permissions

| Permission | Why |
|---|---|
| `activeTab` | Read the URL and communicate with the active Netflix tab |
| `downloads` | Trigger the Save-As dialog for the generated `.xlsx` file |
| `contextMenus` | Add the right-click "Export Netflix My List to Excel" entry |
| `storage` | Persist extraction state so the popup can close and reopen cleanly |
| `*://*.netflix.com/*` | Run the content script and fetch poster images from Netflix's CDN |

No data is sent to any external server. All DOM scraping, Excel generation, and image encoding happen locally in the browser.

---

## Known Limitations

- **Dynamic class names** — Netflix's DOM uses build-time generated class names. A major UI overhaul may require selector updates in `content.js`. The extension uses priority-ordered fallback selectors to be resilient to minor changes.
- **Trusted-event restriction** — Synthetic mouse events are not trusted by Netflix's React handlers. The JawBone panel is opened via URL state change instead, which is reliable but adds ~0.5–1 s per title that lacks fiber-level description data.
- **Profile session timeout** — Netflix may show the profile-picker during a long extraction. The extension detects this and reports it clearly, prompting the user to re-select their profile.
- **`episodesPerSeason` field** — Currently not populated. Netflix's fiber `vm.seasons` array is present but episode counts are not always pre-loaded in the grid view.
- **Language field** — Populated from LD+JSON structured data on the title page; not available via the fiber or JawBone paths.

---

## Author

**KASHIF UMAR**  
© 2025 All rights reserved. Unauthorized reproduction is not permitted.
