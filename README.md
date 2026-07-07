# Fleet Performance Scorecards

A rep- and territory-level KPI scorecard generator built for managed print / office technology sales operations — the kind of quota-and-pipeline reporting a Sales Ops or RevOps analyst would build to replace a manual, spreadsheet-driven KPI dashboard.

Live demo: _add your GitHub Pages URL here after first deploy_

## What it does

- **Three data sources** — built-in mock HubSpot-style data, upload your own CRM exports, or build a roster and deals entirely by hand in the app
- **Manual entry** — add/rename/remove reps (with territory), and add/edit/remove deals (rep, category, stage, amount, account, industry, imaging brand or software type, close date) directly in the UI, no file required. Renaming a rep automatically updates their deals and quota.
- **Multi-file CSV/Excel upload** — load a deals export and a separate activities export at once; rows are merged by rep name, with loose column-header matching so common HubSpot export headers work as-is
- **Three-category quotas** — Net New, Software Solutions, and Imaging are tracked as separate quota targets per rep, matching how a managed print/office tech company typically splits revenue goals
- **Editable quotas** — quota targets usually don't live in a CRM export, so there's a lightweight per-rep, per-category quota panel
- **Per-rep scorecards** — quota attainment for each of the three categories (rendered as toner-cartridge fill gauges), plus win rate, average deal cycle, open pipeline value, deal counts, activity volume, and:
  - **Weighted forecast** — projected period-end result using stage-based close probabilities, shown alongside actual attainment
  - **Quota coverage ratio** — open pipeline ÷ remaining quota gap, so you can see at a glance whether a rep has enough in flight to still hit their number
  - **Cross-sell rate** — % of a rep's closed-won deals that landed in an account which already had a closed-won deal in a *different* category (Net New/Software/Imaging) — expansion into a new category
  - **Upsell rate** — % of a rep's closed-won deals that landed in an account which already had a closed-won deal in the *same* category — repeat/expansion revenue within a category
  - **Avg deal size** — current average deal size across all won deals
- **Territory roll-up chart** — blended attainment *and* forecast by territory, side by side
- **Rep performance chart** — every rep's total closed-won vs. total quota in one horizontal bar chart, so reps can be compared against each other at a glance
- **Avg deal size trend** — a line chart of average deal size per category over time (by close month), so you can spot deal sizes shrinking or growing
- **Closed-won breakdowns** — separate views of closed-won revenue by Industry, Imaging Brand, Software Type, and Cross-Sell/Upsell (by account), filterable by territory
- **AI-generated performance reads** — optional, per-scorecard plain-language summaries via the Claude API, calling out which of the three categories is strongest and which needs attention
- **Local persistence** — uploaded data, quotas, and generated AI reads are saved to your browser's local storage, so they're still there next time you open the app on the same device/browser. Nothing is sent anywhere except the optional Claude API call for AI reads.
- **Print/export** — clean print stylesheet for sharing in a 1:1 or team meeting

## Tech stack

- React 18 + Vite
- Recharts (territory chart)
- PapaParse (CSV parsing) + SheetJS/xlsx (Excel parsing)
- lucide-react (icons)
- Claude API (`claude-sonnet-4-6`) for AI reads — called directly from the browser with a user-supplied API key

## Running locally

```bash
npm install
npm run dev
```

## Uploading your own data

The app expects loosely-HubSpot-shaped exports:

- **Deals file:** columns like `Deal Owner`, `Territory`, `Amount`, `Deal Stage`, `Create Date`, `Close Date`, `Category`, `Industry`, `Imaging Brand`, `Software Type`, `Account`/`Company`
- **Activities file:** columns like `Owner`, `Territory`, `Count`

Column names are matched case-insensitively and by partial match, so most CRM export headers work without renaming. `Category` values are mapped to one of the three quota categories (Net New / Software Solutions / Imaging) by keyword — deals without a recognizable category default to Net New. `Imaging Brand` and `Software Type` are only used when a deal's category is Imaging or Software Solutions, respectively. The `Account`/`Company` column is what powers cross-sell detection — without it, cross-sell rate reads 0% since there's no way to tell which deals share a customer. Quotas aren't typically part of an export, so set those — per category, per rep — in the Quotas panel after loading data.

Everything you load — files, quotas, AI reads — is saved to your browser's local storage automatically, so it persists across page refreshes on the same device. Click "clear" next to the uploaded file list to wipe saved data and start over.

## AI reads

Click "Set API key" in the header and paste an Anthropic API key to enable the "Generate AI read" button on each scorecard. A few important notes:

- The key is held in React state only, for the current browser session — it is never written to disk or any storage, and is not sent anywhere except directly to `api.anthropic.com`.
- **This is a public static site.** A key entered here is visible to anyone inspecting network requests in that browser session. This is fine for personal/demo use, but if you want to share this tool with a team, put the Claude API call behind a small server-side proxy (e.g., a Cloudflare Worker or a tiny Node/Express endpoint) that holds the key server-side, and point the fetch in `src/App.jsx` at that endpoint instead.

## Deploying to GitHub Pages

This repo includes a GitHub Actions workflow (`.github/workflows/deploy.yml`) that builds and deploys automatically on every push to `main`.

One-time setup after pushing this repo to GitHub:

1. Go to **Settings → Pages** in your repository.
2. Under **Build and deployment → Source**, select **GitHub Actions**.
3. Push to `main` (or run the workflow manually from the **Actions** tab).
4. Your site will be live at `https://<your-username>.github.io/<repo-name>/` within a minute or two.

The Vite config uses a relative base path (`base: "./"`), so it works regardless of the repo name — no need to hardcode it anywhere.

## Project structure

```
├── src/
│   ├── App.jsx        # main app: data parsing, KPI calc, scorecards
│   └── main.jsx        # React entry point
├── index.html
├── vite.config.js
├── package.json
└── .github/workflows/deploy.yml
```
