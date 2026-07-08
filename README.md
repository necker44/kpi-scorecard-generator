# Fleet Performance Scorecards

[![Launch App](https://img.shields.io/badge/Launch_App-2C4358?style=for-the-badge&logo=googlechrome&logoColor=white)](https://necker44.github.io/kpi-scorecard-generator/)

A rep- and territory-level KPI scorecard generator built for managed print / office technology sales operations — the kind of quota-and-pipeline reporting a Sales Ops or RevOps analyst would build to replace a manual, spreadsheet-driven KPI dashboard.

**Live demo:** https://necker44.github.io/kpi-scorecard-generator/

## What it does

- **Three data sources** — built-in mock HubSpot-style data, upload your own CRM exports, or build a roster and deals entirely by hand in the app
- **Manual entry** — add/rename/remove reps (with territory), and add/edit/remove deals (rep, category, stage, amount, account, industry, imaging brand or software type, close date) directly in the UI, no file required. Renaming a rep automatically updates their deals and quota.
- **Editable everywhere** — rep names and territories can be renamed inline in the Reps panel regardless of data source (mock, uploaded, or manual); renaming automatically updates that rep's deals, quotas, actuals, and any custom metric values
- **Customizable KPIs** — win rate, avg deal cycle, deal count, cross-sell rate, upsell rate, and avg deal size are all click-to-edit on every scorecard. Type a number to override the calculated value (marked with a small dot), clear the field to revert to the real calculation.
- **Multi-file CSV/Excel upload** — load a deals export and a separate activities export at once; rows are merged by rep name, with loose column-header matching so common HubSpot export headers work as-is
- **Quota Targets matrix** — every rep gets a full grid of quotas: Monthly / Quarterly / Annual **by** Net New / Software Solutions / Imaging (9 editable numbers per rep). Assigned once at the start of the year, completely independent of what they've actually sold. Annual drives the scorecard gauges/forecast; any period feeds the Rep Performance chart.
- **Actual Sales is the primary "sold" number** — a matching Monthly/Quarterly/Annual × category grid where you enter what a rep actually sold. Mock data comes pre-populated here (aggregated from the underlying mock deals) so it's ready to use immediately. Whenever a rep's Annual row is filled in, it becomes the real "sold" figure everywhere: scorecard gauges, forecast, coverage ratio, and the Rep Performance chart all use it instead of the calculated deal total — a small "actual sales" badge appears on the scorecard when this is active. Win rate, avg deal cycle, deal count, pipeline, cross-sell/upsell rate, and the breakdown charts still come from deal records regardless, since those need deal-level detail an aggregate total can't provide. Leave a rep's Annual row at 0 to fall back to the calculated deal total everywhere.
- **Per-rep scorecards** — quota attainment for each of the three categories (rendered as toner-cartridge fill gauges, driven by the Annual quota), plus win rate, average deal cycle, open pipeline value, deal counts, activity volume, and:
  - **Weighted forecast** — projected period-end result using stage-based close probabilities, shown alongside actual attainment
  - **Quota coverage ratio** — open pipeline ÷ remaining quota gap, so you can see at a glance whether a rep has enough in flight to still hit their number
  - **Cross-sell rate** — % of a rep's closed-won deals that landed in an account which already had a closed-won deal in a *different* category (Net New/Software/Imaging) — expansion into a new category
  - **Upsell rate** — % of a rep's closed-won deals that landed in an account which already had a closed-won deal in the *same* category — repeat/expansion revenue within a category
  - **Avg deal size** — current average deal size across all won deals
- **Territory roll-up chart** — blended attainment *and* forecast by territory, side by side
- **Rep performance chart** — every rep's actual sales vs. their quota target for the selected period (Monthly/Quarterly/Annual) in one horizontal bar chart, so reps can be compared against each other at a glance
- **Rep &amp; territory filters** — narrow the Avg Deal Size Trend and Closed-Won Breakdown views down to a single rep, a single territory, or both at once
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
