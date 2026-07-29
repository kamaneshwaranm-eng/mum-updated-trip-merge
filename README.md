# Trip Merge Dashboard

A local dashboard to merge FNV / GRO / WhatsApp trip groups into a clean
`SoId, TrmId` output file, per city.

## Merge types

- **FNV Merge** — uses the 6620 Base file, grouped by trip/customer.
- **GRO Merge** — uses the FNV/GRO merge file + loading sheet.
- **FNV + GRO Merge (Bread/Milk)** — per-store mapping from the FNV/GRO merge file.
- **GRO Merge (Bread + Milk)** — uses the 6620 Base file.
- **WhatsApp Merge** — uses the **6620 Base file only**. You give it groups of
  trip IDs (or customer names) that should be merged together (one group per
  line, or one row per group in an uploaded Excel file). For each group, the
  item with the **most SoIds** in the base file becomes the anchor: its TrmId
  is used, and its own SoIds are skipped. Every other item's SoIds are mapped
  to that anchor TrmId. Available for **Bengaluru** and **Mumbai**.

## Run locally

Requires [Node.js](https://nodejs.org) 20+ (Node 22 recommended).

```bash
# 1. Install dependencies
npm install

# 2. Start the dev server
npm run dev
```

Then open the URL it prints (usually `http://localhost:3000`).

To build a production bundle and preview it:

```bash
npm run build
npm run preview
```

## Host on GitHub

```bash
# From inside this project folder
git init
git add .
git commit -m "Trip merge dashboard"

# Create a new repo on GitHub first (via github.com or `gh repo create`), then:
git remote add origin https://github.com/<your-username>/<your-repo>.git
git branch -M main
git push -u origin main
```

From there you can deploy it anywhere that supports Node (Vercel, Netlify,
Render, or your own server) by running the same `npm install && npm run build`
then serving the output, or run `npm run dev`/`npm run preview` on a server
you control for a fully local/internal setup.

## Everything runs in your browser

All file parsing and merging happens client-side (via the `xlsx` package) —
uploaded files are never sent to a server, so this is safe to run fully
offline/locally with your operational data.
