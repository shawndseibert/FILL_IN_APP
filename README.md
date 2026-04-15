# Site Material Request App

A lightweight mobile-first web app for framing material requests.

## What it does

- Lets manager enter subdivision and lot number.
- Lets manager add one or more items with quantity.
- Autocomplete item descriptions from the inventory CSV in real time.
- Requires item selection from inventory suggestions to ensure a valid POS#.
- On `Send`:
  - Generates email-ready text.
  - Downloads a `.psl` file with lines in this format:
    - `send "POS#<cr>QUANTITY<cr>"`

## Files

- `index.html` - App UI
- `styles.css` - Mobile-first styling
- `app.js` - App logic, CSV parsing, suggestion matching, output generation
- `resources/84 Inventory - 12.30.25 - W0889757.csv` - Inventory source

## Run locally

The app loads the CSV via `fetch`, so open it through a local web server (not directly as `file:///`).

### Option 1: VS Code Live Server

1. Install extension: **Live Server**
2. Right-click `index.html` and choose **Open with Live Server**

### Option 2: Any static server

Serve this folder with any static file server and open the URL in browser.

## Mobile usage

- Use the same server URL from your phone (on the same network), or host it publicly.
- The layout is optimized for touch with large tap targets and minimal clutter.
