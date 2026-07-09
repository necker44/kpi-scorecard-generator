import React, { useState, useMemo, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
  LineChart, Line, Legend,
} from "recharts";
import { Upload, FileSpreadsheet, Sparkles, Printer, ChevronDown, ChevronUp, X, Loader2, Plus, Trash2 } from "lucide-react";

// ---------- Design tokens ----------
// Paper white bg, toner black ink, steel blue structure, amber accent (toner-low light),
// muted green (on-target) / rust (at-risk).
const COLORS = {
  paper: "#F7F4EE",
  paperDim: "#EEEAE1",
  ink: "#1D2226",
  inkSoft: "#4A5259",
  steel: "#3E5C76",
  steelDeep: "#2C4358",
  amber: "#D9992E",
  green: "#4C7A5D",
  rust: "#B5484D",
  line: "#D9D2C4",
};

// ---------- Quota categories ----------
// Each deal belongs to exactly one of these three revenue categories, mirroring how a
// managed print/office tech company typically splits quota: new logos, software/solutions
// attach, and core imaging hardware.
const CATEGORIES = [
  { key: "netNew", label: "Net New" },
  { key: "software", label: "Software Solutions" },
  { key: "imaging", label: "Imaging" },
];
const QUOTA_PERIODS = [
  { key: "monthly", label: "Monthly" },
  { key: "quarterly", label: "Quarterly" },
  { key: "annual", label: "Annual" },
];
const DEFAULT_CATEGORY_QUOTA = { netNew: 50000, software: 30000, imaging: 80000 };
// Full quota matrix: assigned once at the start of the year, per period, per category —
// independent of what a rep has actually sold. Annual drives the scorecard gauges/forecast;
// all three periods are editable and feed the Rep Performance chart via the period picker.
const DEFAULT_QUOTA_MATRIX = {
  monthly: { netNew: 4200, software: 2500, imaging: 6700 },
  quarterly: { netNew: 12600, software: 7500, imaging: 20000 },
  annual: { netNew: 50000, software: 30000, imaging: 80000 },
};
const EMPTY_ACTUAL_MATRIX = {
  monthly: { netNew: 0, software: 0, imaging: 0 },
  quarterly: { netNew: 0, software: 0, imaging: 0 },
  annual: { netNew: 0, software: 0, imaging: 0 },
};
function cloneMatrix(m) {
  return { monthly: { ...m.monthly }, quarterly: { ...m.quarterly }, annual: { ...m.annual } };
}

// Retention/growth metrics — these describe the existing customer base (renewals, expansion),
// so they're tracked per period but not split by Net New/Software/Imaging category the way
// dollar sales are. Lives right alongside Actual Sales and gets mock-populated the same way.
const RETENTION_METRICS = [
  { key: "nrr", label: "Net Renewal Retention", suffix: "%" },
  { key: "grr", label: "Gross Renewal Rate", suffix: "%" },
  { key: "yoyGrowth", label: "YoY Growth (Existing)", suffix: "%" },
];
const DEFAULT_RETENTION_MATRIX = {
  monthly: { nrr: 100, grr: 90, yoyGrowth: 8 },
  quarterly: { nrr: 100, grr: 90, yoyGrowth: 8 },
  annual: { nrr: 100, grr: 90, yoyGrowth: 8 },
};
function cloneRetentionMatrix(m) {
  return { monthly: { ...m.monthly }, quarterly: { ...m.quarterly }, annual: { ...m.annual } };
}
// Retention/growth quota targets are annual-only — there's no meaningful monthly/quarterly
// target for a rate metric like NRR the way there is for a dollar sales quota.
const DEFAULT_RETENTION_QUOTA = { nrr: 100, grr: 90, yoyGrowth: 10 };

// ---------- Mock HubSpot-style data ----------
const REPS = [
  { name: "Marcus Alvarado", territory: "OKC Metro" },
  { name: "Dana Whitfield", territory: "OKC Metro" },
  { name: "Priya Chandran", territory: "Tulsa" },
  { name: "Jordan Meeks", territory: "Tulsa" },
  { name: "Casey O'Brien", territory: "NW Oklahoma" },
  { name: "Levi Sandoval", territory: "South Texas" },
];

const INDUSTRIES = ["Healthcare", "Legal", "Education", "Manufacturing", "Government", "Financial Services"];
const IMAGING_BRANDS = ["Canon", "Konica Minolta", "Kyocera", "Sharp", "Ricoh", "Xerox"];
const SOFTWARE_TYPES = ["Document Management", "Print Management", "Workflow Automation", "ECM", "Managed IT / Cloud", "Cybersecurity"];
const ACCOUNT_POOL = [
  "Meridian Health Partners", "Crestview Legal Group", "Lakeside School District",
  "Ironforge Manufacturing", "Prairie County Government", "Sable Financial Group",
  "Northgate Medical Center", "Harrison & Cole Law", "Blue Ridge Academy",
  "Delta Fabrication Co", "Union County Courthouse", "Cornerstone Credit Union",
  "Riverside Dental Group", "Whitfield & Associates", "Oakmont Elementary",
  "Titan Steel Works", "City of Bellmont", "Summit Financial Advisors",
];

const STAGES_OPEN = ["Discovery", "Proposal Sent", "Negotiation"];

// Rough stage-to-close probabilities for weighted forecasting.
function stageProbability(stage) {
  const s = (stage || "").toLowerCase();
  if (s.includes("won")) return 1;
  if (s.includes("lost")) return 0;
  if (s.includes("discovery")) return 0.2;
  if (s.includes("proposal")) return 0.5;
  if (s.includes("negotiation")) return 0.75;
  return 0.3;
}

function sampleWithoutReplacement(pool, n, rnd) {
  const arr = [...pool];
  const result = [];
  for (let i = 0; i < n && arr.length > 0; i++) {
    const idx = Math.floor(rnd() * arr.length);
    result.push(arr[idx]);
    arr.splice(idx, 1);
  }
  return result;
}

function seededRandom(seed) {
  let s = seed;
  return () => {
    s = (s * 9301 + 49297) % 233280;
    return s / 233280;
  };
}

// Generates a full Monthly/Quarterly/Annual x Net New/Software/Imaging Actual Sales matrix
// directly from a rep's quota, using rough managed-print/MSP industry attainment benchmarks
// rather than bucketing sparse individual deals (which left plenty of real gaps at $0).
// Net New is hardest to attain consistently (new-logo acquisition); Imaging is the most
// mature, predictable line; Software attach is the newest motion and the most variable.
const ATTAINMENT_RANGES = {
  netNew: [0.65, 1.05],
  software: [0.55, 1.15],
  imaging: [0.80, 1.12],
};
function generateActualsFromQuota(quota, rnd) {
  const q = quota || DEFAULT_QUOTA_MATRIX;
  const annualAttainment = {};
  CATEGORIES.forEach(({ key }) => {
    const [lo, hi] = ATTAINMENT_RANGES[key];
    annualAttainment[key] = lo + rnd() * (hi - lo);
  });
  const buildPeriod = (periodKey, varianceSpread) => {
    const period = {};
    CATEGORIES.forEach(({ key }) => {
      const base = annualAttainment[key];
      const varied = Math.max(0.15, base + (rnd() - 0.5) * varianceSpread);
      const quotaVal = (q[periodKey] && q[periodKey][key]) || 0;
      period[key] = Math.round(quotaVal * varied);
    });
    return period;
  };
  return {
    annual: buildPeriod("annual", 0), // annual uses the base attainment rate exactly
    quarterly: buildPeriod("quarterly", 0.15),
    monthly: buildPeriod("monthly", 0.25),
  };
}

function generateMockData() {
  const rnd = seededRandom(42);
  const deals = [];
  const activities = [];
  const quotaMatrix = {};
  const actualMatrix = {};
  const retentionMatrix = {};
  const retentionQuota = {};
  let dealId = 1;

  REPS.forEach((rep) => {
    const annual = {
      netNew: 55000 + Math.floor(rnd() * 20000),
      software: 35000 + Math.floor(rnd() * 15000),
      imaging: 95000 + Math.floor(rnd() * 35000),
    };
    quotaMatrix[rep.name] = {
      annual,
      quarterly: { netNew: Math.round(annual.netNew / 4), software: Math.round(annual.software / 4), imaging: Math.round(annual.imaging / 4) },
      monthly: { netNew: Math.round(annual.netNew / 12), software: Math.round(annual.software / 12), imaging: Math.round(annual.imaging / 12) },
    };
    const repAccounts = sampleWithoutReplacement(ACCOUNT_POOL, 6 + Math.floor(rnd() * 4), rnd);
    const dealCount = 16 + Math.floor(rnd() * 12);
    for (let i = 0; i < dealCount; i++) {
      const stageRoll = rnd();
      const stage =
        stageRoll < 0.45 ? "Closed Won" :
        stageRoll < 0.62 ? "Closed Lost" :
        STAGES_OPEN[Math.floor(rnd() * STAGES_OPEN.length)];
      const created = new Date(2026, Math.floor(rnd() * 5), 1 + Math.floor(rnd() * 27));
      const cycle = 12 + Math.floor(rnd() * 55);
      const closeDate = new Date(created.getTime() + cycle * 86400000);

      const catRoll = rnd();
      const category = catRoll < 0.28 ? "netNew" : catRoll < 0.55 ? "software" : "imaging";
      const industry = INDUSTRIES[Math.floor(rnd() * INDUSTRIES.length)];
      const imagingBrand = category === "imaging" ? IMAGING_BRANDS[Math.floor(rnd() * IMAGING_BRANDS.length)] : null;
      const softwareType = category === "software" ? SOFTWARE_TYPES[Math.floor(rnd() * SOFTWARE_TYPES.length)] : null;
      const account = repAccounts[Math.floor(rnd() * repAccounts.length)];

      const amount =
        category === "imaging" ? Math.floor(4000 + rnd() * 38000) :
        category === "software" ? Math.floor(2000 + rnd() * 12000) :
        Math.floor(800 + rnd() * 6000);

      const deal = {
        id: dealId++,
        rep: rep.name,
        territory: rep.territory,
        amount,
        stage,
        category,
        industry,
        imagingBrand,
        softwareType,
        account,
        createdDate: created,
        closeDate: stage.startsWith("Closed") ? closeDate : null,
      };
      deals.push(deal);
    }
    const activityCount = 60 + Math.floor(rnd() * 140);
    activities.push({ rep: rep.name, territory: rep.territory, count: activityCount });

    // Actual Sales: generated directly from this rep's quota using industry attainment
    // benchmarks, guaranteeing every period/category cell is populated with a realistic
    // figure (rather than bucketing sparse individual deals, which left real gaps at $0).
    actualMatrix[rep.name] = generateActualsFromQuota(quotaMatrix[rep.name], rnd);

    // Realistic MSP/managed-print benchmarks: NRR ~95-112% (expansion can push over 100),
    // GRR ~85-96% (always <=100, can't gain revenue by definition), YoY growth ~2-16%.
    // Quarterly/monthly carry the same underlying rate with small natural variance rather
    // than being summed/divided like dollar quotas, since these are ratios, not totals.
    const round1 = (n) => Math.round(n * 10) / 10;
    const annualNRR = 95 + rnd() * 17;
    const annualGRR = 85 + rnd() * 11;
    const annualYoY = 2 + rnd() * 14;
    retentionMatrix[rep.name] = {
      annual: { nrr: round1(annualNRR), grr: round1(annualGRR), yoyGrowth: round1(annualYoY) },
      quarterly: { nrr: round1(annualNRR + (rnd() - 0.5) * 5), grr: round1(Math.min(annualGRR + (rnd() - 0.5) * 5, 100)), yoyGrowth: round1(annualYoY + (rnd() - 0.5) * 5) },
      monthly: { nrr: round1(annualNRR + (rnd() - 0.5) * 7), grr: round1(Math.min(annualGRR + (rnd() - 0.5) * 7, 100)), yoyGrowth: round1(annualYoY + (rnd() - 0.5) * 7) },
    };
    // Annual-only targets — company goals are usually a bit more uniform than actual results.
    retentionQuota[rep.name] = {
      nrr: round1(100 + (rnd() - 0.5) * 6),
      grr: round1(90 + (rnd() - 0.5) * 4),
      yoyGrowth: round1(10 + (rnd() - 0.5) * 4),
    };
  });

  return { deals, activities, quotaMatrix, actualMatrix, retentionMatrix, retentionQuota };
}

// ---------- File parsing helpers ----------
const HEADER_MAP = {
  rep: ["deal owner", "owner", "rep", "sales rep", "account manager", "assigned to"],
  territory: ["territory", "region", "pipeline", "area"],
  amount: ["amount", "deal amount", "value"],
  stage: ["deal stage", "stage", "status"],
  createdDate: ["create date", "created date", "created", "deal created date"],
  closeDate: ["close date", "closed date"],
  activityCount: ["count", "activities", "activity count", "number of activities"],
  category: ["category", "deal category", "product type", "line item type", "deal type", "revenue category"],
  industry: ["industry", "vertical", "industry vertical"],
  imagingBrand: ["imaging brand", "equipment brand", "brand", "manufacturer", "oem"],
  softwareType: ["software type", "solution type", "software category", "software"],
  account: ["account", "company", "account name", "customer", "client", "company name"],
};

function findKey(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const hit = keys.find((k) => k.trim().toLowerCase() === c);
    if (hit) return hit;
  }
  // fallback: partial match
  for (const c of candidates) {
    const hit = keys.find((k) => k.trim().toLowerCase().includes(c));
    if (hit) return hit;
  }
  return null;
}

function normalizeCategory(raw) {
  if (!raw) return "netNew";
  const v = String(raw).toLowerCase();
  if (v.includes("software") || v.includes("solution")) return "software";
  if (v.includes("imag") || v.includes("hardware") || v.includes("equipment") || v.includes("mfp") || v.includes("copier") || v.includes("device")) return "imaging";
  if (v.includes("net new") || v.includes("new business") || v.includes("new logo")) return "netNew";
  return "netNew";
}

function parseRowsAsDealsOrActivities(rows) {
  if (!rows.length) return { deals: [], activities: [] };
  const sample = rows[0];
  const repKey = findKey(sample, HEADER_MAP.rep);
  const amountKey = findKey(sample, HEADER_MAP.amount);
  const activityKey = findKey(sample, HEADER_MAP.activityCount);

  const deals = [];
  const activities = [];

  if (amountKey) {
    const territoryKey = findKey(sample, HEADER_MAP.territory);
    const stageKey = findKey(sample, HEADER_MAP.stage);
    const createdKey = findKey(sample, HEADER_MAP.createdDate);
    const closedKey = findKey(sample, HEADER_MAP.closeDate);
    const categoryKey = findKey(sample, HEADER_MAP.category);
    const industryKey = findKey(sample, HEADER_MAP.industry);
    const imagingBrandKey = findKey(sample, HEADER_MAP.imagingBrand);
    const softwareTypeKey = findKey(sample, HEADER_MAP.softwareType);
    const accountKey = findKey(sample, HEADER_MAP.account);

    rows.forEach((r, i) => {
      const rep = repKey ? String(r[repKey]).trim() : "Unassigned";
      if (!rep || rep === "undefined") return;
      const amt = amountKey ? parseFloat(String(r[amountKey]).replace(/[^0-9.-]/g, "")) : 0;
      const created = createdKey && r[createdKey] ? new Date(r[createdKey]) : null;
      const closed = closedKey && r[closedKey] ? new Date(r[closedKey]) : null;
      const category = normalizeCategory(categoryKey ? r[categoryKey] : null);
      deals.push({
        id: `u-${i}-${rep}`,
        rep,
        territory: territoryKey ? String(r[territoryKey]).trim() : "Unassigned",
        amount: isNaN(amt) ? 0 : amt,
        stage: stageKey ? String(r[stageKey]).trim() : "Unknown",
        createdDate: created && !isNaN(created) ? created : null,
        closeDate: closed && !isNaN(closed) ? closed : null,
        category,
        industry: industryKey && r[industryKey] ? String(r[industryKey]).trim() : "Unspecified",
        imagingBrand: category === "imaging" && imagingBrandKey && r[imagingBrandKey] ? String(r[imagingBrandKey]).trim() : null,
        softwareType: category === "software" && softwareTypeKey && r[softwareTypeKey] ? String(r[softwareTypeKey]).trim() : null,
        account: accountKey && r[accountKey] ? String(r[accountKey]).trim() : `${rep}-unknown-${i}`,
      });
    });
  } else if (activityKey || repKey) {
    const territoryKey = findKey(sample, HEADER_MAP.territory);
    rows.forEach((r) => {
      const rep = repKey ? String(r[repKey]).trim() : null;
      if (!rep) return;
      const count = activityKey ? parseFloat(r[activityKey]) : 1;
      activities.push({
        rep,
        territory: territoryKey ? String(r[territoryKey]).trim() : "Unassigned",
        count: isNaN(count) ? 1 : count,
      });
    });
  }
  return { deals, activities };
}

async function parseFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    const text = await file.text();
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    return parseRowsAsDealsOrActivities(parsed.data);
  } else {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    return parseRowsAsDealsOrActivities(rows);
  }
}

// Grouped by rep+account so books don't cross-pollinate. Within each account's win history:
//  - Cross-sell = a deal in a category the account hasn't bought before (expansion into a new category)
//  - Upsell = a deal in a category the account already bought (repeat/expansion within the same category)
// The very first won deal for an account is neither — it's the original sale.
function computeExpansionIds(deals) {
  const won = deals.filter((d) => d.stage.toLowerCase().includes("won"));
  const byKey = {};
  won.forEach((d) => {
    const acctKey = d.account && String(d.account).trim() ? d.account.trim() : `unassigned-${d.id}`;
    const key = `${d.rep}::${acctKey}`;
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push(d);
  });
  const crossSellIds = new Set();
  const upsellIds = new Set();
  Object.values(byKey).forEach((list) => {
    const sorted = [...list].sort((a, b) => (a.closeDate ? a.closeDate.getTime() : 0) - (b.closeDate ? b.closeDate.getTime() : 0));
    const seen = new Set();
    sorted.forEach((d) => {
      if (seen.size > 0) {
        if (seen.has(d.category)) upsellIds.add(d.id);
        else crossSellIds.add(d.id);
      }
      seen.add(d.category);
    });
  });
  return { crossSellIds, upsellIds };
}

// ---------- KPI computation ----------
function computeKPIs(deals, activities, quotas, roster = [], actualMatrix = {}) {
  const repNames = Array.from(new Set([...deals.map((d) => d.rep), ...roster.map((r) => r.name)]));
  const rosterTerritory = {};
  roster.forEach((r) => { rosterTerritory[r.name] = r.territory; });
  const byRep = {};
  const { crossSellIds, upsellIds } = computeExpansionIds(deals);

  repNames.forEach((rep) => {
    const repDeals = deals.filter((d) => d.rep === rep);
    const won = repDeals.filter((d) => d.stage.toLowerCase().includes("won"));
    const lost = repDeals.filter((d) => d.stage.toLowerCase().includes("lost"));
    const open = repDeals.filter(
      (d) => !d.stage.toLowerCase().includes("won") && !d.stage.toLowerCase().includes("lost")
    );
    const winRate = won.length + lost.length > 0 ? (won.length / (won.length + lost.length)) * 100 : null;

    const cycles = won
      .filter((d) => d.createdDate && d.closeDate)
      .map((d) => (d.closeDate - d.createdDate) / 86400000);
    const avgCycle = cycles.length ? cycles.reduce((a, b) => a + b, 0) / cycles.length : null;

    const territory = repDeals[0]?.territory || rosterTerritory[rep] || "Unassigned";
    const activityTotal = activities
      .filter((a) => a.rep === rep)
      .reduce((s, a) => s + (a.count || 0), 0);

    const crossSellWonCount = won.filter((d) => crossSellIds.has(d.id)).length;
    const crossSellRate = won.length > 0 ? (crossSellWonCount / won.length) * 100 : null;
    const upsellWonCount = won.filter((d) => upsellIds.has(d.id)).length;
    const upsellRate = won.length > 0 ? (upsellWonCount / won.length) * 100 : null;
    const avgDealSize = won.length > 0 ? won.reduce((s, d) => s + (d.amount || 0), 0) / won.length : null;

    // Actual Sales entry is the strict source of truth for "what's actually sold" — no
    // fallback to deal totals. If a rep's Annual row hasn't been filled in, that's 0 sold.
    const repActual = (actualMatrix[rep] && actualMatrix[rep].annual) || EMPTY_ACTUAL_MATRIX.annual;

    const repQuotas = quotas[rep] || DEFAULT_CATEGORY_QUOTA;
    const categories = {};
    let totalWon = 0, totalQuota = 0, totalPipeline = 0, totalForecast = 0;
    CATEGORIES.forEach(({ key, label }) => {
      const catOpenDeals = open.filter((d) => d.category === key);
      const catWon = repActual[key] || 0;
      const catPipeline = catOpenDeals.reduce((s, d) => s + (d.amount || 0), 0);
      const catForecast = catWon + catOpenDeals.reduce((s, d) => s + (d.amount || 0) * stageProbability(d.stage), 0);
      const catQuota = repQuotas[key] || 0;
      const catAttainment = catQuota ? (catWon / catQuota) * 100 : null;
      const catGap = Math.max(catQuota - catWon, 0);
      const catCoverage = catGap > 0 ? catPipeline / catGap : null; // null = quota already met
      categories[key] = {
        label, won: catWon, quota: catQuota, attainment: catAttainment, pipeline: catPipeline,
        forecast: catForecast, forecastAttainment: catQuota ? (catForecast / catQuota) * 100 : null,
        coverage: catCoverage, gap: catGap,
      };
      totalWon += catWon;
      totalQuota += catQuota;
      totalPipeline += catPipeline;
      totalForecast += catForecast;
    });
    const totalAttainment = totalQuota ? (totalWon / totalQuota) * 100 : null;
    const totalForecastAttainment = totalQuota ? (totalForecast / totalQuota) * 100 : null;
    const totalGap = Math.max(totalQuota - totalWon, 0);
    const totalCoverage = totalGap > 0 ? totalPipeline / totalGap : null;

    byRep[rep] = {
      rep,
      territory,
      categories,
      closedWonAmount: totalWon,
      quota: totalQuota,
      attainment: totalAttainment,
      pipelineAmount: totalPipeline,
      forecastAmount: totalForecast,
      forecastAttainment: totalForecastAttainment,
      coverage: totalCoverage,
      gap: totalGap,
      winRate,
      avgCycle,
      avgDealSize,
      crossSellRate,
      crossSellWonCount,
      upsellRate,
      upsellWonCount,
      activityTotal,
      dealCount: repDeals.length,
      wonCount: won.length,
    };
  });

  const territories = {};
  Object.values(byRep).forEach((r) => {
    if (!territories[r.territory]) {
      territories[r.territory] = {
        territory: r.territory,
        closedWonAmount: 0,
        quota: 0,
        pipelineAmount: 0,
        forecastAmount: 0,
        reps: 0,
        categories: {
          netNew: { won: 0, quota: 0 },
          software: { won: 0, quota: 0 },
          imaging: { won: 0, quota: 0 },
        },
      };
    }
    const t = territories[r.territory];
    t.closedWonAmount += r.closedWonAmount;
    t.quota += r.quota;
    t.pipelineAmount += r.pipelineAmount;
    t.forecastAmount += r.forecastAmount;
    t.reps += 1;
    CATEGORIES.forEach(({ key }) => {
      t.categories[key].won += r.categories[key].won;
      t.categories[key].quota += r.categories[key].quota;
    });
  });

  return { byRep, territories };
}

// Avg deal size per month per category, for the trend chart.
function computeAvgDealSizeTrend(deals) {
  const won = deals.filter((d) => d.stage.toLowerCase().includes("won") && d.closeDate);
  const byMonth = {};
  won.forEach((d) => {
    const monthKey = `${d.closeDate.getFullYear()}-${String(d.closeDate.getMonth() + 1).padStart(2, "0")}`;
    if (!byMonth[monthKey]) {
      byMonth[monthKey] = {};
      CATEGORIES.forEach(({ key }) => { byMonth[monthKey][key] = { sum: 0, count: 0 }; });
    }
    byMonth[monthKey][d.category].sum += d.amount || 0;
    byMonth[monthKey][d.category].count += 1;
  });
  return Object.keys(byMonth).sort().map((month) => {
    const entry = { month };
    CATEGORIES.forEach(({ key }) => {
      const c = byMonth[month][key];
      entry[key] = c.count > 0 ? Math.round(c.sum / c.count) : null;
    });
    return entry;
  });
}

// Group closed-won dollars by an arbitrary field (industry, imaging brand, software type)
function groupWonAmountBy(deals, field) {
  const won = deals.filter((d) => d.stage.toLowerCase().includes("won"));
  const map = {};
  won.forEach((d) => {
    const key = d[field] || "Unspecified";
    if (!map[key]) map[key] = { name: key, amount: 0, count: 0 };
    map[key].amount += d.amount || 0;
    map[key].count += 1;
  });
  return Object.values(map).sort((a, b) => b.amount - a.amount);
}

// Per-rep industry/imaging-brand/software-type mix, dollar totals scaled to match that rep's
// Actual Sales figures per category (same approach as the Closed-Won Breakdown chart), used to
// feed the "Ask About Your Data" Q&A with specifics an aggregate quota number can't answer.
function computeRepBreakdownMix(deals, actualMatrix, repName) {
  const repDeals = deals.filter((d) => d.rep === repName);
  const scale = {};
  CATEGORIES.forEach(({ key }) => {
    const dealTotal = repDeals.filter((d) => d.category === key && d.stage.toLowerCase().includes("won")).reduce((s, d) => s + (d.amount || 0), 0);
    const actualTotal = (actualMatrix[repName] && actualMatrix[repName].annual && actualMatrix[repName].annual[key]) || 0;
    scale[key] = dealTotal > 0 ? actualTotal / dealTotal : 0;
  });
  const wonScaled = repDeals
    .filter((d) => d.stage.toLowerCase().includes("won"))
    .map((d) => ({ ...d, amount: (d.amount || 0) * (scale[d.category] || 0) }));

  const groupBy = (list, field) => {
    const map = {};
    list.forEach((d) => {
      const k = d[field] || "Unspecified";
      map[k] = (map[k] || 0) + d.amount;
    });
    return map;
  };

  return {
    industry: groupBy(wonScaled, "industry"),
    imagingBrand: groupBy(wonScaled.filter((d) => d.category === "imaging"), "imagingBrand"),
    softwareType: groupBy(wonScaled.filter((d) => d.category === "software"), "softwareType"),
  };
}

function fmtMoney(n) {
  if (n == null) return "—";
  return "$" + Math.round(n).toLocaleString();
}
function fmtPct(n) {
  if (n == null) return "—";
  return Math.round(n) + "%";
}

// ---------- Toner gauge (signature element) ----------
function TonerGauge({ pct, label, compact = false }) {
  const clamped = Math.max(0, Math.min(pct ?? 0, 130));
  const fillHeight = Math.min(clamped, 100);
  const overfill = clamped > 100 ? clamped - 100 : 0;
  const color = pct == null ? COLORS.inkSoft : pct >= 100 ? COLORS.green : pct >= 75 ? COLORS.amber : COLORS.rust;
  const w = compact ? 26 : 34;
  const h = compact ? 52 : 68;

  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
      <div
        style={{
          position: "relative",
          width: w,
          height: h,
          border: `2px solid ${COLORS.ink}`,
          borderRadius: 4,
          background: COLORS.paperDim,
          overflow: "hidden",
          flexShrink: 0,
        }}
        aria-label={`${label ? label + " " : ""}quota attainment ${pct == null ? "unknown" : Math.round(pct) + "%"}`}
      >
        <div style={{ position: "absolute", top: -2, left: "50%", transform: "translateX(-50%)", width: w * 0.4, height: 4, background: COLORS.ink, borderRadius: "0 0 2px 2px" }} />
        <div
          style={{
            position: "absolute", bottom: 0, left: 0, right: 0,
            height: `${fillHeight}%`, background: color, transition: "height 400ms ease",
          }}
        />
        {overfill > 0 && (
          <div style={{ position: "absolute", top: 2, left: 2, right: 2, height: 2, background: COLORS.ink }} />
        )}
        {[25, 50, 75].map((t) => (
          <div key={t} style={{ position: "absolute", bottom: `${t}%`, left: 0, right: 0, height: 1, background: "rgba(29,34,38,0.25)" }} />
        ))}
      </div>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: compact ? 14 : 20, fontWeight: 600, color: COLORS.ink, lineHeight: 1 }}>
          {pct == null ? "—" : Math.round(pct) + "%"}
        </div>
        {label && (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, color: COLORS.inkSoft, letterSpacing: 0.3, marginTop: 2, textTransform: "uppercase" }}>
            {label}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Scorecard ----------
function Scorecard({ data, insight, onGenerateInsight, insightLoading, overrides, onSetOverride, onClearOverride, retention }) {
  const ret = retention || DEFAULT_RETENTION_MATRIX.annual;
  const [open, setOpen] = useState(true);
  const ov = overrides || {};
  const attainmentColor =
    data.attainment == null ? COLORS.inkSoft : data.attainment >= 100 ? COLORS.green : data.attainment >= 75 ? COLORS.amber : COLORS.rust;

  return (
    <div
      style={{
        background: "#FFFFFF",
        border: `1px solid ${COLORS.line}`,
        borderRadius: 6,
        boxShadow: "0 1px 2px rgba(29,34,38,0.06)",
        overflow: "hidden",
      }}
    >
      <div style={{ borderTop: `4px solid ${attainmentColor}` }} />
      <div style={{ padding: "18px 20px 14px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 24, fontWeight: 600, color: COLORS.ink, letterSpacing: 0.3, textTransform: "uppercase" }}>
              {data.rep}
            </div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.steel, marginTop: 2 }}>
              {data.territory}
            </div>
          </div>
          <button
            onClick={() => setOpen((o) => !o)}
            style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.inkSoft, padding: 4 }}
            aria-label={open ? "Collapse details" : "Expand details"}
          >
            {open ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
          </button>
        </div>

        {/* Three category gauges: Net New / Software / Imaging */}
        <div style={{ marginTop: 14, display: "flex", justifyContent: "space-between", gap: 8 }}>
          {CATEGORIES.map(({ key, label }) => (
            <TonerGauge key={key} pct={data.categories[key].attainment} label={label} compact />
          ))}
        </div>

        <div style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "baseline", borderTop: `1px dashed ${COLORS.line}`, paddingTop: 10 }}>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: COLORS.inkSoft, textTransform: "uppercase" }}>Total (Actual Sales)</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 600, color: COLORS.ink }}>
            {fmtMoney(data.closedWonAmount)} <span style={{ color: COLORS.inkSoft, fontWeight: 400 }}>of {fmtMoney(data.quota)}</span>
          </span>
        </div>

        <div style={{ marginTop: 6, display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase" }}>Forecast</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLORS.steel }}>
            {fmtMoney(data.forecastAmount)} <span style={{ color: COLORS.inkSoft }}>({fmtPct(data.forecastAttainment)} of quota)</span>
          </span>
        </div>

        {open && (
          <>
            <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, borderTop: `1px dashed ${COLORS.line}`, paddingTop: 14 }}>
              <EditableMetric label="Win rate" rawValue={ov.winRate != null ? ov.winRate : data.winRate} isOverride={ov.winRate != null} format={(v) => (v == null ? "—" : Math.round(v) + "%")} onSave={(v) => onSetOverride("winRate", v)} onClear={() => onClearOverride("winRate")} />
              <EditableMetric label="Avg cycle" rawValue={ov.avgCycle != null ? ov.avgCycle : data.avgCycle} isOverride={ov.avgCycle != null} format={(v) => (v == null ? "—" : Math.round(v) + "d")} onSave={(v) => onSetOverride("avgCycle", v)} onClear={() => onClearOverride("avgCycle")} />
              <Metric label="Pipeline" value={fmtMoney(data.pipelineAmount)} />
              <EditableMetric label="Deals" rawValue={ov.dealCount != null ? ov.dealCount : data.dealCount} isOverride={ov.dealCount != null} format={(v) => (v == null ? "—" : Math.round(v))} onSave={(v) => onSetOverride("dealCount", v)} onClear={() => onClearOverride("dealCount")} />
              <EditableMetric label="Won" rawValue={ov.wonCount != null ? ov.wonCount : data.wonCount} isOverride={ov.wonCount != null} format={(v) => (v == null ? "—" : Math.round(v))} onSave={(v) => onSetOverride("wonCount", v)} onClear={() => onClearOverride("wonCount")} />
              <Metric label="Activities" value={data.activityTotal || "—"} />
              <Metric label="Coverage" value={data.coverage == null ? "Met" : data.coverage.toFixed(1) + "×"} />
              <EditableMetric label="Cross-sell" rawValue={ov.crossSellRate != null ? ov.crossSellRate : data.crossSellRate} isOverride={ov.crossSellRate != null} format={(v) => (v == null ? "—" : Math.round(v) + "%")} onSave={(v) => onSetOverride("crossSellRate", v)} onClear={() => onClearOverride("crossSellRate")} />
              <EditableMetric label="Upsell" rawValue={ov.upsellRate != null ? ov.upsellRate : data.upsellRate} isOverride={ov.upsellRate != null} format={(v) => (v == null ? "—" : Math.round(v) + "%")} onSave={(v) => onSetOverride("upsellRate", v)} onClear={() => onClearOverride("upsellRate")} />
              <EditableMetric label="Avg deal size" rawValue={ov.avgDealSize != null ? ov.avgDealSize : data.avgDealSize} isOverride={ov.avgDealSize != null} format={(v) => fmtMoney(v)} onSave={(v) => onSetOverride("avgDealSize", v)} onClear={() => onClearOverride("avgDealSize")} />
              <Metric label="Net Renewal Retention" value={ret.nrr + "%"} />
              <Metric label="Gross Renewal Rate" value={ret.grr + "%"} />
              <Metric label="YoY Growth (Existing)" value={ret.yoyGrowth + "%"} />
            </div>
            <div style={{ fontSize: 10.5, color: COLORS.inkSoft, marginTop: 6 }}>Win rate, avg cycle, deals, won, cross-sell, upsell, and avg deal size are editable — click a value to type your own, clear the field to go back to the calculated number. Retention &amp; growth figures are set in the Retention &amp; Growth panel above.</div>

            <div style={{ marginTop: 14 }}>
              {insight ? (
                <div style={{ background: COLORS.paperDim, borderRadius: 4, padding: "10px 12px", fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 13.5, color: COLORS.ink, lineHeight: 1.5 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, fontWeight: 600, color: COLORS.steel, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 5 }}>
                    <Sparkles size={12} /> AI read
                  </div>
                  {insight}
                </div>
              ) : (
                <button
                  onClick={onGenerateInsight}
                  disabled={insightLoading}
                  style={{
                    display: "flex", alignItems: "center", gap: 6, background: "none",
                    border: `1px solid ${COLORS.steel}`, color: COLORS.steel, borderRadius: 4,
                    padding: "6px 12px", fontSize: 12.5, fontFamily: "'IBM Plex Sans', sans-serif",
                    cursor: insightLoading ? "default" : "pointer", opacity: insightLoading ? 0.6 : 1,
                  }}
                >
                  {insightLoading ? <Loader2 size={13} className="spin" /> : <Sparkles size={13} />}
                  {insightLoading ? "Reading the numbers…" : "Generate AI read"}
                </button>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function Metric({ label, value }) {
  return (
    <div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLORS.ink }}>{value}</div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: COLORS.inkSoft, letterSpacing: 0.4, textTransform: "uppercase", marginTop: 1 }}>{label}</div>
    </div>
  );
}

// A metric tile that's click-to-edit: click the value, type a custom number, blur/Enter to save.
// Clearing the field back to blank removes the override and reverts to the calculated value.
function EditableMetric({ label, rawValue, isOverride, format, onSave, onClear }) {
  const [editing, setEditing] = useState(false);

  const commit = (e) => {
    const v = e.target.value.trim();
    if (v === "") onClear();
    else {
      const parsed = parseFloat(v);
      if (!isNaN(parsed)) onSave(parsed);
    }
    setEditing(false);
  };

  if (editing) {
    return (
      <div>
        <input
          type="number"
          autoFocus
          defaultValue={rawValue ?? ""}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === "Enter") e.target.blur();
            if (e.key === "Escape") setEditing(false);
          }}
          style={{ width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, border: `1px solid ${COLORS.steel}`, borderRadius: 4, padding: "3px 5px" }}
        />
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: COLORS.inkSoft, letterSpacing: 0.4, textTransform: "uppercase", marginTop: 1 }}>{label}</div>
      </div>
    );
  }

  return (
    <div onClick={() => setEditing(true)} style={{ cursor: "pointer" }} title="Click to enter a custom value">
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: isOverride ? COLORS.steel : COLORS.ink, display: "flex", alignItems: "center", gap: 5 }}>
        {format(rawValue)}
        {isOverride && <span style={{ width: 5, height: 5, borderRadius: "50%", background: COLORS.amber, display: "inline-block" }} title="Custom value" />}
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: COLORS.inkSoft, letterSpacing: 0.4, textTransform: "uppercase", marginTop: 1 }}>{label}</div>
    </div>
  );
}

// ---------- Breakdown bar chart (industry / imaging brand / software type) ----------
function BreakdownChart({ data, emptyMessage }) {
  if (!data.length) {
    return <div style={{ color: COLORS.inkSoft, fontSize: 13, padding: "20px 4px" }}>{emptyMessage || "No closed-won deals in this group yet."}</div>;
  }
  return (
    <ResponsiveContainer width="100%" height={Math.max(140, data.length * 34)}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 24, left: 8, bottom: 4 }}>
        <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} horizontal={false} />
        <XAxis type="number" tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} tickFormatter={(v) => "$" + Math.round(v / 1000) + "k"} />
        <YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 12.5, fontFamily: "IBM Plex Sans" }} stroke={COLORS.inkSoft} />
        <Tooltip formatter={(v, n, p) => [fmtMoney(v), `${p.payload.count} deal${p.payload.count === 1 ? "" : "s"}`]} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12 }} />
        <Bar dataKey="amount" radius={[0, 3, 3, 0]} fill={COLORS.steel} />
      </BarChart>
    </ResponsiveContainer>
  );
}

// ---------- Local persistence (this browser/device only) ----------
const STORAGE_KEY = "kpi-scorecard-generator-v2";

function serializeDeals(deals) {
  return deals.map((d) => ({
    ...d,
    createdDate: d.createdDate ? d.createdDate.toISOString() : null,
    closeDate: d.closeDate ? d.closeDate.toISOString() : null,
  }));
}

function deserializeDeals(deals) {
  return (deals || []).map((d) => ({
    ...d,
    createdDate: d.createdDate ? new Date(d.createdDate) : null,
    closeDate: d.closeDate ? new Date(d.closeDate) : null,
  }));
}

function saveToLocalStorage(state) {
  try {
    const serializable = {
      source: state.source,
      quotaMatrix: state.quotaMatrix,
      actualMatrix: state.actualMatrix,
      retentionMatrix: state.retentionMatrix,
      retentionQuota: state.retentionQuota,
      fileNames: state.fileNames,
      insights: state.insights,
      deals: serializeDeals(state.deals),
      activities: state.activities,
      manualDeals: serializeDeals(state.manualDeals || []),
      manualRoster: state.manualRoster || [],
      mockRoster: state.mockRoster || [],
      uploadRoster: state.uploadRoster || [],
      metricOverrides: state.metricOverrides || {},
    };
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(serializable));
  } catch (e) {
    console.warn("Couldn't save scorecard data locally:", e);
  }
}

function loadFromLocalStorage() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    parsed.deals = deserializeDeals(parsed.deals);
    parsed.manualDeals = deserializeDeals(parsed.manualDeals);
    parsed.manualRoster = parsed.manualRoster || [];
    parsed.mockRoster = parsed.mockRoster || null;
    parsed.uploadRoster = parsed.uploadRoster || [];
    parsed.metricOverrides = parsed.metricOverrides || {};
    parsed.quotaMatrix = parsed.quotaMatrix || null;
    parsed.actualMatrix = parsed.actualMatrix || {};
    parsed.retentionMatrix = parsed.retentionMatrix || {};
    parsed.retentionQuota = parsed.retentionQuota || {};
    return parsed;
  } catch (e) {
    console.warn("Couldn't load saved scorecard data:", e);
    return null;
  }
}

function clearLocalStorage() {
  try {
    window.localStorage.removeItem(STORAGE_KEY);
  } catch (e) {
    // ignore
  }
}

// ---------- Manual entry editor (roster + deals, fully independent of any import) ----------
const STAGE_OPTIONS = ["Discovery", "Proposal Sent", "Negotiation", "Closed Won", "Closed Lost"];

function makeDealId() {
  return `m-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function ManualEntryEditor({ roster, setRoster, deals, setDeals, quotaMatrix, setQuotaMatrix, setActualMatrix, setRetentionMatrix }) {
  const repNames = roster.map((r) => r.name);

  const addRep = () => {
    let n = 1;
    let name = `New Rep ${roster.length + 1}`;
    while (repNames.includes(name)) { n += 1; name = `New Rep ${roster.length + n}`; }
    setRoster((r) => [...r, { name, territory: "" }]);
    setQuotaMatrix((q) => ({ ...q, [name]: cloneMatrix(DEFAULT_QUOTA_MATRIX) }));
    setRetentionMatrix((r) => ({ ...r, [name]: cloneRetentionMatrix(DEFAULT_RETENTION_MATRIX) }));
  };

  const renameRep = (index, newName) => {
    const oldName = roster[index].name;
    if (!newName || newName === oldName) {
      setRoster((r) => r.map((x, i) => (i === index ? { ...x, name: newName } : x)));
      return;
    }
    setRoster((r) => r.map((x, i) => (i === index ? { ...x, name: newName } : x)));
    setDeals((d) => d.map((deal) => (deal.rep === oldName ? { ...deal, rep: newName } : deal)));
    setQuotaMatrix((q) => {
      const next = { ...q };
      if (next[oldName]) {
        next[newName] = next[oldName];
        delete next[oldName];
      }
      return next;
    });
    setActualMatrix((a) => {
      const next = { ...a };
      if (next[oldName]) {
        next[newName] = next[oldName];
        delete next[oldName];
      }
      return next;
    });
    setRetentionMatrix((r) => {
      const next = { ...r };
      if (next[oldName]) {
        next[newName] = next[oldName];
        delete next[oldName];
      }
      return next;
    });
  };

  const updateTerritory = (index, territory) => {
    const name = roster[index].name;
    setRoster((r) => r.map((x, i) => (i === index ? { ...x, territory } : x)));
    setDeals((d) => d.map((deal) => (deal.rep === name ? { ...deal, territory } : deal)));
  };

  const removeRep = (index) => {
    const name = roster[index].name;
    setRoster((r) => r.filter((_, i) => i !== index));
    setDeals((d) => d.filter((deal) => deal.rep !== name));
    setQuotaMatrix((q) => {
      const next = { ...q };
      delete next[name];
      return next;
    });
    setActualMatrix((a) => {
      const next = { ...a };
      delete next[name];
      return next;
    });
    setRetentionMatrix((r) => {
      const next = { ...r };
      delete next[name];
      return next;
    });
  };

  const addDeal = () => {
    const firstRep = roster[0];
    setDeals((d) => [
      ...d,
      {
        id: makeDealId(),
        rep: firstRep ? firstRep.name : "",
        territory: firstRep ? firstRep.territory : "",
        amount: 0,
        stage: "Discovery",
        category: "netNew",
        industry: "Unspecified",
        imagingBrand: null,
        softwareType: null,
        account: "",
        createdDate: null,
        closeDate: null,
      },
    ]);
  };

  const updateDeal = (index, patch) => {
    setDeals((d) => d.map((deal, i) => (i === index ? { ...deal, ...patch } : deal)));
  };

  const removeDeal = (index) => {
    setDeals((d) => d.filter((_, i) => i !== index));
  };

  const inputStyle = { border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 7px", fontSize: 12.5, fontFamily: "'IBM Plex Sans', sans-serif", width: "100%" };
  const thStyle = { textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 6px 6px", whiteSpace: "nowrap" };
  const tdStyle = { padding: "4px 6px", verticalAlign: "top" };

  return (
    <div>
      <div style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 15, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3, color: COLORS.ink }}>
            Reps
          </div>
          <button onClick={addRep} style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: `1px solid ${COLORS.steel}`, color: COLORS.steel, borderRadius: 4, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>
            <Plus size={13} /> Add rep
          </button>
        </div>
        {roster.length === 0 ? (
          <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No reps yet — add one to get started.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse", width: "100%" }}>
              <thead>
                <tr>
                  <th style={thStyle}>Name</th>
                  <th style={thStyle}>Territory</th>
                  <th style={thStyle}></th>
                </tr>
              </thead>
              <tbody>
                {roster.map((r, i) => (
                  <tr key={i}>
                    <td style={tdStyle}><input style={{ ...inputStyle, minWidth: 140 }} value={r.name} onChange={(e) => renameRep(i, e.target.value)} /></td>
                    <td style={tdStyle}><input style={{ ...inputStyle, minWidth: 120 }} value={r.territory} onChange={(e) => updateTerritory(i, e.target.value)} placeholder="Territory" /></td>
                    <td style={tdStyle}>
                      <button onClick={() => removeRep(i)} style={{ background: "none", border: "none", color: COLORS.rust, cursor: "pointer", padding: 4 }} aria-label="Remove rep">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 15, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3, color: COLORS.ink }}>
            Deals
          </div>
          <button onClick={addDeal} disabled={roster.length === 0} style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: `1px solid ${COLORS.steel}`, color: COLORS.steel, borderRadius: 4, padding: "5px 10px", fontSize: 12, cursor: roster.length === 0 ? "not-allowed" : "pointer", opacity: roster.length === 0 ? 0.5 : 1 }}>
            <Plus size={13} /> Add deal
          </button>
        </div>
        {roster.length === 0 ? (
          <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>Add a rep first, then add their deals.</div>
        ) : deals.length === 0 ? (
          <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No deals yet — add one above.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse", width: "100%" }}>
              <thead>
                <tr>
                  <th style={thStyle}>Rep</th>
                  <th style={thStyle}>Category</th>
                  <th style={thStyle}>Stage</th>
                  <th style={thStyle}>Amount</th>
                  <th style={thStyle}>Account</th>
                  <th style={thStyle}>Industry</th>
                  <th style={thStyle}>Brand / SW type</th>
                  <th style={thStyle}>Close date</th>
                  <th style={thStyle}></th>
                </tr>
              </thead>
              <tbody>
                {deals.map((d, i) => (
                  <tr key={d.id}>
                    <td style={tdStyle}>
                      <select
                        style={{ ...inputStyle, minWidth: 130 }}
                        value={d.rep}
                        onChange={(e) => {
                          const rep = roster.find((r) => r.name === e.target.value);
                          updateDeal(i, { rep: e.target.value, territory: rep ? rep.territory : "" });
                        }}
                      >
                        {roster.map((r) => (
                          <option key={r.name} value={r.name}>{r.name}</option>
                        ))}
                      </select>
                    </td>
                    <td style={tdStyle}>
                      <select
                        style={{ ...inputStyle, minWidth: 110 }}
                        value={d.category}
                        onChange={(e) => {
                          const category = e.target.value;
                          updateDeal(i, {
                            category,
                            imagingBrand: category === "imaging" ? (d.imagingBrand || IMAGING_BRANDS[0]) : null,
                            softwareType: category === "software" ? (d.softwareType || SOFTWARE_TYPES[0]) : null,
                          });
                        }}
                      >
                        {CATEGORIES.map((c) => (
                          <option key={c.key} value={c.key}>{c.label}</option>
                        ))}
                      </select>
                    </td>
                    <td style={tdStyle}>
                      <select style={{ ...inputStyle, minWidth: 110 }} value={d.stage} onChange={(e) => updateDeal(i, { stage: e.target.value })}>
                        {STAGE_OPTIONS.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </td>
                    <td style={tdStyle}>
                      <input
                        type="number"
                        style={{ ...inputStyle, minWidth: 90, textAlign: "right" }}
                        value={d.amount}
                        onChange={(e) => updateDeal(i, { amount: parseFloat(e.target.value) || 0 })}
                      />
                    </td>
                    <td style={tdStyle}>
                      <input style={{ ...inputStyle, minWidth: 130 }} value={d.account} onChange={(e) => updateDeal(i, { account: e.target.value })} placeholder="Account / company" />
                    </td>
                    <td style={tdStyle}>
                      <select style={{ ...inputStyle, minWidth: 120 }} value={d.industry} onChange={(e) => updateDeal(i, { industry: e.target.value })}>
                        <option value="Unspecified">Unspecified</option>
                        {INDUSTRIES.map((ind) => (
                          <option key={ind} value={ind}>{ind}</option>
                        ))}
                      </select>
                    </td>
                    <td style={tdStyle}>
                      {d.category === "imaging" && (
                        <select style={{ ...inputStyle, minWidth: 120 }} value={d.imagingBrand || IMAGING_BRANDS[0]} onChange={(e) => updateDeal(i, { imagingBrand: e.target.value })}>
                          {IMAGING_BRANDS.map((b) => (<option key={b} value={b}>{b}</option>))}
                        </select>
                      )}
                      {d.category === "software" && (
                        <select style={{ ...inputStyle, minWidth: 150 }} value={d.softwareType || SOFTWARE_TYPES[0]} onChange={(e) => updateDeal(i, { softwareType: e.target.value })}>
                          {SOFTWARE_TYPES.map((s) => (<option key={s} value={s}>{s}</option>))}
                        </select>
                      )}
                      {d.category === "netNew" && <span style={{ color: COLORS.inkSoft, fontSize: 12 }}>—</span>}
                    </td>
                    <td style={tdStyle}>
                      <input
                        type="date"
                        style={{ ...inputStyle, minWidth: 130 }}
                        value={d.closeDate ? d.closeDate.toISOString().slice(0, 10) : ""}
                        onChange={(e) => updateDeal(i, { closeDate: e.target.value ? new Date(e.target.value) : null })}
                      />
                    </td>
                    <td style={tdStyle}>
                      <button onClick={() => removeDeal(i)} style={{ background: "none", border: "none", color: COLORS.rust, cursor: "pointer", padding: 4 }} aria-label="Remove deal">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Main App ----------
export default function App() {
  const saved = useMemo(() => loadFromLocalStorage(), []);
  const [source, setSource] = useState(saved?.source ?? "mock");
  const mockData = useMemo(() => generateMockData(), []);
  const initialMockRoster = useMemo(() => REPS.map((r) => ({ originalName: r.name, displayName: r.name, territory: r.territory })), []);
  const [mockRoster, setMockRoster] = useState(saved?.mockRoster ?? initialMockRoster);
  const [uploadedDeals, setUploadedDeals] = useState(saved?.deals ?? []);
  const [uploadRoster, setUploadRoster] = useState(saved?.uploadRoster ?? []);
  const [uploadedActivities, setUploadedActivities] = useState(saved?.activities ?? []);
  const [fileNames, setFileNames] = useState(saved?.fileNames ?? []);
  const [manualRoster, setManualRoster] = useState(saved?.manualRoster ?? []);
  const [manualDeals, setManualDeals] = useState(saved?.manualDeals ?? []);
  const [quotaMatrix, setQuotaMatrix] = useState(saved?.quotaMatrix ?? mockData.quotaMatrix);
  const [actualMatrix, setActualMatrix] = useState(
    saved?.actualMatrix && Object.keys(saved.actualMatrix).length > 0 ? saved.actualMatrix : mockData.actualMatrix
  );
  const [retentionMatrix, setRetentionMatrix] = useState(
    saved?.retentionMatrix && Object.keys(saved.retentionMatrix).length > 0 ? saved.retentionMatrix : mockData.retentionMatrix
  );
  const [retentionQuota, setRetentionQuota] = useState(
    saved?.retentionQuota && Object.keys(saved.retentionQuota).length > 0 ? saved.retentionQuota : mockData.retentionQuota
  );
  const [quotaPeriod, setQuotaPeriod] = useState("annual");
  const [metricOverrides, setMetricOverrides] = useState(saved?.metricOverrides ?? {});
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState(null);
  const [insights, setInsights] = useState(saved?.insights ?? {});
  const [insightLoading, setInsightLoading] = useState({});
  const [territoryFilter, setTerritoryFilter] = useState("All");
  const [repFilter, setRepFilter] = useState("All");
  const [breakdownTab, setBreakdownTab] = useState("industry");
  const [apiKey, setApiKey] = useState("");
  const [showKeyInput, setShowKeyInput] = useState(false);
  const [qaQuestion, setQaQuestion] = useState("");
  const [qaHistory, setQaHistory] = useState([]);
  const [qaLoading, setQaLoading] = useState(false);
  const [justSaved, setJustSaved] = useState(false);

  // Mock deals/activities carry their original generated rep names — remap them through
  // mockRoster so renaming a mock rep doesn't require regenerating the dataset.
  const mockDeals = useMemo(() => {
    const nameMap = {}, territoryMap = {};
    mockRoster.forEach((r) => { nameMap[r.originalName] = r.displayName; territoryMap[r.originalName] = r.territory; });
    return mockData.deals.map((d) => ({ ...d, rep: nameMap[d.rep] ?? d.rep, territory: territoryMap[d.rep] ?? d.territory }));
  }, [mockData, mockRoster]);
  const mockActivities = useMemo(() => {
    const nameMap = {};
    mockRoster.forEach((r) => { nameMap[r.originalName] = r.displayName; });
    return mockData.activities.map((a) => ({ ...a, rep: nameMap[a.rep] ?? a.rep }));
  }, [mockData, mockRoster]);

  const deals = source === "mock" ? mockDeals : source === "upload" ? uploadedDeals : manualDeals;
  const activities = source === "mock" ? mockActivities : source === "upload" ? uploadedActivities : [];
  const roster = source === "manual" ? manualRoster
    : source === "mock" ? mockRoster.map((r) => ({ name: r.displayName, territory: r.territory }))
    : uploadRoster;

  // Whenever the mock tab is active, make sure mock reps have their sample Actual Sales
  // filled in. A rep counts as "not yet populated" if every period/category is 0 — covers
  // both a first-ever visit and a browser that already has stale/zeroed data saved from
  // before this feature existed. Reps the user has actually entered numbers for are left alone.
  React.useEffect(() => {
    if (source !== "mock") return;
    setActualMatrix((prev) => {
      let changed = false;
      const next = { ...prev };
      mockRoster.forEach((r) => {
        const current = next[r.displayName];
        const isEmpty = !current || QUOTA_PERIODS.every((p) => {
          const c = (current && current[p.key]) || {};
          return (c.netNew || 0) + (c.software || 0) + (c.imaging || 0) === 0;
        });
        if (isEmpty) {
          const seeded = mockData.actualMatrix[r.originalName];
          if (seeded) {
            next[r.displayName] = cloneMatrix(seeded);
            changed = true;
          }
        }
      });
      return changed ? next : prev;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, mockRoster]);

  // Same idea for Retention & Growth — seed any mock rep missing an entry with their sample
  // NRR/GRR/YoY numbers. Unlike Actual Sales, 0% is a plausible real value here, so "missing
  // entirely" (rather than "all zero") is what triggers a re-seed.
  React.useEffect(() => {
    if (source !== "mock") return;
    setRetentionMatrix((prev) => {
      let changed = false;
      const next = { ...prev };
      mockRoster.forEach((r) => {
        if (!next[r.displayName]) {
          const seeded = mockData.retentionMatrix[r.originalName];
          if (seeded) {
            next[r.displayName] = cloneRetentionMatrix(seeded);
            changed = true;
          }
        }
      });
      return changed ? next : prev;
    });
    setRetentionQuota((prev) => {
      let changed = false;
      const next = { ...prev };
      mockRoster.forEach((r) => {
        if (!next[r.displayName]) {
          const seeded = mockData.retentionQuota[r.originalName];
          if (seeded) {
            next[r.displayName] = { ...seeded };
            changed = true;
          }
        }
      });
      return changed ? next : prev;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, mockRoster]);

  // Renaming/re-territorying works the same way regardless of which data source is active.
  const renameRepGlobal = (oldName, newName) => {
    if (!newName || newName === oldName) return;
    if (source === "mock") {
      setMockRoster((r) => r.map((x) => (x.displayName === oldName ? { ...x, displayName: newName } : x)));
    } else if (source === "upload") {
      setUploadedDeals((d) => d.map((deal) => (deal.rep === oldName ? { ...deal, rep: newName } : deal)));
      setUploadedActivities((a) => a.map((act) => (act.rep === oldName ? { ...act, rep: newName } : act)));
      setUploadRoster((r) => r.map((x) => (x.name === oldName ? { ...x, name: newName } : x)));
    } else if (source === "manual") {
      setManualRoster((r) => r.map((x) => (x.name === oldName ? { ...x, name: newName } : x)));
      setManualDeals((d) => d.map((deal) => (deal.rep === oldName ? { ...deal, rep: newName } : deal)));
    }
    setQuotaMatrix((q) => {
      const next = { ...q };
      if (next[oldName]) { next[newName] = next[oldName]; delete next[oldName]; }
      return next;
    });
    setActualMatrix((a) => {
      const next = { ...a };
      if (next[oldName]) { next[newName] = next[oldName]; delete next[oldName]; }
      return next;
    });
    setRetentionMatrix((r) => {
      const next = { ...r };
      if (next[oldName]) { next[newName] = next[oldName]; delete next[oldName]; }
      return next;
    });
    setRetentionQuota((r) => {
      const next = { ...r };
      if (next[oldName]) { next[newName] = next[oldName]; delete next[oldName]; }
      return next;
    });
    setMetricOverrides((o) => {
      if (!o[oldName]) return o;
      const next = { ...o };
      next[newName] = next[oldName];
      delete next[oldName];
      return next;
    });
  };

  const updateTerritoryGlobal = (repName, territory) => {
    if (source === "mock") {
      setMockRoster((r) => r.map((x) => (x.displayName === repName ? { ...x, territory } : x)));
    } else if (source === "upload") {
      setUploadedDeals((d) => d.map((deal) => (deal.rep === repName ? { ...deal, territory } : deal)));
      setUploadRoster((r) => r.map((x) => (x.name === repName ? { ...x, territory } : x)));
    } else if (source === "manual") {
      setManualRoster((r) => r.map((x) => (x.name === repName ? { ...x, territory } : x)));
      setManualDeals((d) => d.map((deal) => (deal.rep === repName ? { ...deal, territory } : deal)));
    }
  };

  // Adds a brand-new, zero-deal rep to whichever source is currently active. They show up
  // immediately in the Reps list, Quota Targets, Actual Sales, Retention & Growth, every chart,
  // and get their own scorecard — same as any other rep, just starting from a blank slate.
  const addRepGlobal = () => {
    const existingNames = new Set(Object.keys(byRep));
    let n = 1;
    let name = "New Rep";
    while (existingNames.has(name)) { n += 1; name = `New Rep ${n}`; }

    if (source === "mock") {
      setMockRoster((r) => [...r, { originalName: name, displayName: name, territory: "" }]);
    } else if (source === "upload") {
      setUploadRoster((r) => [...r, { name, territory: "" }]);
    } else {
      setManualRoster((r) => [...r, { name, territory: "" }]);
    }
    setQuotaMatrix((q) => ({ ...q, [name]: cloneMatrix(DEFAULT_QUOTA_MATRIX) }));
    setRetentionMatrix((r) => ({ ...r, [name]: cloneRetentionMatrix(DEFAULT_RETENTION_MATRIX) }));
    setRetentionQuota((r) => ({ ...r, [name]: { ...DEFAULT_RETENTION_QUOTA } }));
    if (source === "mock") {
      // Give a brand-new mock rep realistic starting actuals too, rather than a flat $0 that
      // would look broken next to a populated quota.
      setActualMatrix((a) => ({ ...a, [name]: generateActualsFromQuota(DEFAULT_QUOTA_MATRIX, Math.random) }));
    }
  };

  const setMetricOverride = (repName, field, value) => {
    setMetricOverrides((o) => ({ ...o, [repName]: { ...(o[repName] || {}), [field]: value } }));
  };
  const clearMetricOverride = (repName, field) => {
    setMetricOverrides((o) => {
      if (!o[repName]) return o;
      const next = { ...o, [repName]: { ...o[repName] } };
      delete next[repName][field];
      return next;
    });
  };

  React.useEffect(() => {
    if (source === "mock" && uploadedDeals.length === 0 && manualDeals.length === 0 && manualRoster.length === 0 && Object.keys(insights).length === 0 && Object.keys(metricOverrides).length === 0 && Object.keys(actualMatrix).length === 0) return;
    saveToLocalStorage({
      source, quotaMatrix, actualMatrix, retentionMatrix, retentionQuota, fileNames, insights,
      deals: uploadedDeals, activities: uploadedActivities,
      manualDeals, manualRoster, mockRoster, uploadRoster, metricOverrides,
    });
    setJustSaved(true);
    const t = setTimeout(() => setJustSaved(false), 1500);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, quotaMatrix, actualMatrix, retentionMatrix, retentionQuota, fileNames, insights, uploadedDeals, uploadedActivities, manualDeals, manualRoster, mockRoster, uploadRoster, metricOverrides]);

  const quotasForKPI = useMemo(() => {
    const out = {};
    Object.keys(quotaMatrix).forEach((rep) => { out[rep] = (quotaMatrix[rep] && quotaMatrix[rep].annual) || DEFAULT_CATEGORY_QUOTA; });
    return out;
  }, [quotaMatrix]);

  const { byRep, territories } = useMemo(() => computeKPIs(deals, activities, quotasForKPI, roster, actualMatrix), [deals, activities, quotasForKPI, roster, actualMatrix]);

  const territoryOptions = ["All", ...Object.keys(territories)];

  const filteredDeals = useMemo(
    () => deals
      .filter((d) => territoryFilter === "All" || d.territory === territoryFilter)
      .filter((d) => repFilter === "All" || d.rep === repFilter),
    [deals, territoryFilter, repFilter]
  );

  // The Industry/Imaging Brand/Software Type/Cross-Sell breakdowns need deal-level detail
  // (which Actual Sales doesn't track), so we scale each rep+category's won-deal amounts to
  // match that rep's Actual Sales figure for that category — the mix comes from deals, the
  // totals come from Actual Sales. A rep/category with no Actual Sales entered shows $0 here,
  // same as everywhere else.
  const dealsScaledToActual = useMemo(() => {
    const scale = {};
    filteredDeals.forEach((d) => {
      const repKey = d.rep;
      if (!scale[repKey]) scale[repKey] = {};
      if (scale[repKey][d.category] == null) {
        const dealTotal = filteredDeals
          .filter((x) => x.rep === repKey && x.category === d.category && x.stage.toLowerCase().includes("won"))
          .reduce((s, x) => s + (x.amount || 0), 0);
        const actualTotal = ((actualMatrix[repKey] && actualMatrix[repKey].annual && actualMatrix[repKey].annual[d.category]) || 0);
        scale[repKey][d.category] = dealTotal > 0 ? actualTotal / dealTotal : 0;
      }
    });
    return filteredDeals.map((d) => {
      if (!d.stage.toLowerCase().includes("won")) return d;
      const f = (scale[d.rep] && scale[d.rep][d.category] != null) ? scale[d.rep][d.category] : 0;
      return { ...d, amount: (d.amount || 0) * f };
    });
  }, [filteredDeals, actualMatrix]);

  const breakdowns = useMemo(() => {
    const { crossSellIds, upsellIds } = computeExpansionIds(dealsScaledToActual);
    return {
      industry: groupWonAmountBy(dealsScaledToActual, "industry"),
      imagingBrand: groupWonAmountBy(dealsScaledToActual.filter((d) => d.category === "imaging"), "imagingBrand"),
      softwareType: groupWonAmountBy(dealsScaledToActual.filter((d) => d.category === "software"), "softwareType"),
      crossSell: groupWonAmountBy(dealsScaledToActual.filter((d) => crossSellIds.has(d.id) || upsellIds.has(d.id)), "account"),
    };
  }, [dealsScaledToActual]);

  const repOptions = ["All", ...Object.keys(byRep).sort()];

  const repList = Object.values(byRep)
    .filter((r) => territoryFilter === "All" || r.territory === territoryFilter)
    .filter((r) => repFilter === "All" || r.rep === repFilter)
    .sort((a, b) => (b.attainment ?? -1) - (a.attainment ?? -1));

  const handleFiles = useCallback(async (fileList) => {
    setParsing(true);
    setError(null);
    try {
      let allDeals = [...uploadedDeals];
      let allActivities = [...uploadedActivities];
      const names = [...fileNames];
      for (const file of fileList) {
        const { deals: d, activities: a } = await parseFile(file);
        allDeals = allDeals.concat(d);
        allActivities = allActivities.concat(a);
        names.push(file.name);
      }
      setUploadedDeals(allDeals);
      setUploadedActivities(allActivities);
      setFileNames(names);
      const repNames = Array.from(new Set(allDeals.map((d) => d.rep)));
      setQuotaMatrix((q) => {
        const next = { ...q };
        repNames.forEach((r) => {
          if (!(r in next)) next[r] = cloneMatrix(DEFAULT_QUOTA_MATRIX);
        });
        return next;
      });
      setRetentionMatrix((r) => {
        const next = { ...r };
        repNames.forEach((rep) => {
          if (!(rep in next)) next[rep] = cloneRetentionMatrix(DEFAULT_RETENTION_MATRIX);
        });
        return next;
      });
      setRetentionQuota((r) => {
        const next = { ...r };
        repNames.forEach((rep) => {
          if (!(rep in next)) next[rep] = { ...DEFAULT_RETENTION_QUOTA };
        });
        return next;
      });
      setSource("upload");
    } catch (e) {
      setError("Couldn't read one of those files. Check that it's a valid CSV or Excel export.");
    } finally {
      setParsing(false);
    }
  }, [uploadedDeals, uploadedActivities, fileNames]);

  const resetUploads = () => {
    setUploadedDeals([]);
    setUploadedActivities([]);
    setUploadRoster([]);
    setFileNames([]);
    setInsights({});
    setQuotaMatrix(mockData.quotaMatrix);
    // Don't wipe local storage entirely — manual entry data should survive clearing an upload.
  };

  const generateInsight = async (repRaw) => {
    if (!apiKey) {
      setShowKeyInput(true);
      return;
    }
    const ov = metricOverrides[repRaw.rep] || {};
    const rep = {
      ...repRaw,
      winRate: ov.winRate != null ? ov.winRate : repRaw.winRate,
      avgCycle: ov.avgCycle != null ? ov.avgCycle : repRaw.avgCycle,
      dealCount: ov.dealCount != null ? ov.dealCount : repRaw.dealCount,
      wonCount: ov.wonCount != null ? ov.wonCount : repRaw.wonCount,
      crossSellRate: ov.crossSellRate != null ? ov.crossSellRate : repRaw.crossSellRate,
      upsellRate: ov.upsellRate != null ? ov.upsellRate : repRaw.upsellRate,
      avgDealSize: ov.avgDealSize != null ? ov.avgDealSize : repRaw.avgDealSize,
    };
    setInsightLoading((s) => ({ ...s, [repRaw.rep]: true }));
    try {
      const catLines = CATEGORIES.map(({ key, label }) => {
        const c = rep.categories[key];
        return `${label}: ${fmtMoney(c.won)} of ${fmtMoney(c.quota)} quota (${fmtPct(c.attainment)})`;
      }).join("\n");

      const prompt = `You are a sales operations analyst at a managed print/copier/IT services company. Write a 2-3 sentence, plain-language read on this rep's performance for a scorecard. Call out which of their three quota categories (Net New, Software Solutions, Imaging) is strongest and which needs attention, and mention whether their forecast suggests they'll close the gap. Be specific and direct, no fluff, no headers.

Rep: ${rep.rep}
Territory: ${rep.territory}
${catLines}
Overall: ${fmtMoney(rep.closedWonAmount)} of ${fmtMoney(rep.quota)} quota (${fmtPct(rep.attainment)})
Forecast: ${fmtMoney(rep.forecastAmount)} (${fmtPct(rep.forecastAttainment)} of quota)
Pipeline coverage: ${rep.coverage == null ? "quota already met" : rep.coverage.toFixed(1) + "x remaining gap"}
Cross-sell rate: ${fmtPct(rep.crossSellRate)} of won deals sold into an existing account in a new category
Upsell rate: ${fmtPct(rep.upsellRate)} of won deals that expanded an existing account within the same category
Avg deal size: ${fmtMoney(rep.avgDealSize)}
Win rate: ${fmtPct(rep.winRate)}
Won deals: ${rep.wonCount} of ${rep.dealCount} total deals
Avg deal cycle: ${rep.avgCycle == null ? "unknown" : Math.round(rep.avgCycle) + " days"}
Open pipeline: ${fmtMoney(rep.pipelineAmount)}
Activity count this period: ${rep.activityTotal || "unknown"}`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 300,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      if (!response.ok) throw new Error("API error");
      const data = await response.json();
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join(" ").trim();
      setInsights((s) => ({ ...s, [rep.rep]: text || "No read available." }));
    } catch (e) {
      setInsights((s) => ({ ...s, [rep.rep]: "Couldn't generate a read right now — check that your API key is valid." }));
    } finally {
      setInsightLoading((s) => ({ ...s, [rep.rep]: false }));
    }
  };

  // ---------- Ask About Your Data ----------
  const buildDataSummary = () => {
    const lines = Object.values(byRep).map((r) => {
      const mix = computeRepBreakdownMix(deals, actualMatrix, r.rep);
      const fmtMix = (obj) => Object.entries(obj).filter(([, v]) => v > 0).map(([k, v]) => `${k} ${fmtMoney(v)}`).join(", ") || "none";
      const catLine = CATEGORIES.map(({ key, label }) => `${label}: ${fmtMoney(r.categories[key].won)} of ${fmtMoney(r.categories[key].quota)} quota (${fmtPct(r.categories[key].attainment)})`).join("; ");
      const ret = (retentionMatrix[r.rep] && retentionMatrix[r.rep].annual) || DEFAULT_RETENTION_MATRIX.annual;
      return `${r.rep} (${r.territory}):
  Actual sales — ${catLine}; Total ${fmtMoney(r.closedWonAmount)} of ${fmtMoney(r.quota)} (${fmtPct(r.attainment)})
  Forecast: ${fmtMoney(r.forecastAmount)} (${fmtPct(r.forecastAttainment)}); Coverage: ${r.coverage == null ? "quota met" : r.coverage.toFixed(1) + "x"}
  Win rate: ${fmtPct(r.winRate)}; Avg cycle: ${r.avgCycle == null ? "n/a" : Math.round(r.avgCycle) + "d"}; Deals: ${r.dealCount} (${r.wonCount} won); Avg deal size: ${fmtMoney(r.avgDealSize)}
  Cross-sell: ${fmtPct(r.crossSellRate)}; Upsell: ${fmtPct(r.upsellRate)}; Activities: ${r.activityTotal || 0}; Open pipeline: ${fmtMoney(r.pipelineAmount)}
  Net Renewal Retention: ${ret.nrr}%; Gross Renewal Rate: ${ret.grr}%; YoY growth from existing customers: ${ret.yoyGrowth}% (annual)
  Software type mix (won $): ${fmtMix(mix.softwareType)}
  Imaging brand mix (won $): ${fmtMix(mix.imagingBrand)}
  Industry mix (won $): ${fmtMix(mix.industry)}`;
    });
    return lines.join("\n\n");
  };

  const askDataQuestion = async () => {
    if (!apiKey) {
      setShowKeyInput(true);
      return;
    }
    const q = qaQuestion.trim();
    if (!q) return;
    setQaLoading(true);
    try {
      const summary = buildDataSummary();
      const prompt = `You are a sales operations analyst assistant. Answer the question using ONLY the data below — don't invent numbers. Be concise (2-4 sentences), name specific reps and figures, and say plainly if the data doesn't cover what's being asked.

DATA (per rep, current period):
${summary}

QUESTION: ${q}`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 400,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      if (!response.ok) throw new Error("API error");
      const data = await response.json();
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join(" ").trim();
      setQaHistory((h) => [...h, { question: q, answer: text || "No answer available." }]);
      setQaQuestion("");
    } catch (e) {
      setQaHistory((h) => [...h, { question: q, answer: "Couldn't get an answer right now — check that your API key is valid." }]);
    } finally {
      setQaLoading(false);
    }
  };

  const territoryChartData = Object.values(territories).map((t) => ({
    name: t.territory,
    attainment: t.quota > 0 ? Math.round((t.closedWonAmount / t.quota) * 100) : 0,
    forecastPct: t.quota > 0 ? Math.round((t.forecastAmount / t.quota) * 100) : 0,
  }));

  const repChartData = repList.map((r) => {
    const repMatrix = quotaMatrix[r.rep] || DEFAULT_QUOTA_MATRIX;
    const periodQuota = repMatrix[quotaPeriod] || DEFAULT_QUOTA_MATRIX[quotaPeriod];
    const quotaSum = (periodQuota.netNew || 0) + (periodQuota.software || 0) + (periodQuota.imaging || 0);
    const repActual = (actualMatrix[r.rep] && actualMatrix[r.rep][quotaPeriod]) || EMPTY_ACTUAL_MATRIX[quotaPeriod];
    const actualSum = (repActual.netNew || 0) + (repActual.software || 0) + (repActual.imaging || 0);
    return { name: r.rep, won: actualSum, quota: quotaSum };
  });

  const retentionChartData = repList.map((r) => {
    const m = retentionMatrix[r.rep] || DEFAULT_RETENTION_MATRIX;
    const period = m[quotaPeriod] || DEFAULT_RETENTION_MATRIX[quotaPeriod];
    return { name: r.rep, nrr: period.nrr, grr: period.grr, yoyGrowth: period.yoyGrowth };
  });

  const avgDealSizeTrend = useMemo(() => computeAvgDealSizeTrend(filteredDeals), [filteredDeals]);

  const breakdownLabels = { industry: "Industry", imagingBrand: "Imaging Brand", softwareType: "Software Type", crossSell: "Cross-Sell / Upsell" };

  return (
    <div style={{ minHeight: "100vh", background: COLORS.paper, fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        .spin { animation: spin 1s linear infinite; }
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        input[type="number"] { font-family: 'IBM Plex Mono', monospace; }
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
        }
      `}</style>

      {/* Header */}
      <div style={{ background: COLORS.ink, color: COLORS.paper, padding: "22px 28px" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 30, fontWeight: 700, letterSpacing: 0.5, textTransform: "uppercase" }}>
              Fleet Performance Scorecards
            </div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: "#B8C4CC", marginTop: 2 }}>
              Rep &amp; territory KPI tracker — office tech &amp; managed print
            </div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: justSaved ? "#8FBF9F" : "#7A8790", marginTop: 6, transition: "color 300ms" }}>
              {justSaved ? "✓ saved to this device" : "data saved locally on this device"}
            </div>
          </div>
          <div className="no-print" style={{ display: "flex", gap: 10 }}>
            <button
              onClick={() => setShowKeyInput((s) => !s)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.paper}`, color: COLORS.paper, borderRadius: 4, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}
            >
              <Sparkles size={14} /> {apiKey ? "API key set" : "Set API key"}
            </button>
            <button
              onClick={() => window.print()}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.paper}`, color: COLORS.paper, borderRadius: 4, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}
            >
              <Printer size={14} /> Print / Export
            </button>
          </div>
        </div>
        {showKeyInput && (
          <div className="no-print" style={{ maxWidth: 1180, margin: "12px auto 0", background: "#26313A", borderRadius: 6, padding: "12px 16px" }}>
            <div style={{ fontSize: 12.5, color: "#B8C4CC", marginBottom: 8, lineHeight: 1.5 }}>
              Paste an Anthropic API key to enable AI reads on each scorecard. This stays in memory for this browser session only — it is never saved, and it's sent directly from your browser to Anthropic. Because this is a public static site, don't use a key you're not comfortable being visible in browser dev tools; for team use, put this behind a small server-side proxy instead.
            </div>
            <input
              type="password"
              placeholder="sk-ant-..."
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              style={{ width: "100%", maxWidth: 420, padding: "7px 10px", borderRadius: 4, border: "none", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}
            />
          </div>
        )}
      </div>

      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "24px 28px 60px" }}>
        {/* Data source controls */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
            <SourceTab active={source === "mock"} onClick={() => setSource("mock")} label="Mock HubSpot data" />
            <SourceTab active={source === "upload"} onClick={() => uploadedDeals.length ? setSource("upload") : document.getElementById("file-input").click()} label="Uploaded files" />
            <SourceTab active={source === "manual"} onClick={() => setSource("manual")} label="Manual entry" />
          </div>

          {source !== "manual" && (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                <label
                  htmlFor="file-input"
                  style={{
                    display: "flex", alignItems: "center", gap: 8, cursor: "pointer",
                    border: `1.5px dashed ${COLORS.steel}`, borderRadius: 5, padding: "10px 16px",
                    color: COLORS.steel, fontSize: 13.5, fontWeight: 500,
                  }}
                >
                  {parsing ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
                  {parsing ? "Reading files…" : "Upload CSV / Excel (multiple allowed)"}
                </label>
                <input
                  id="file-input"
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  multiple
                  style={{ display: "none" }}
                  onChange={(e) => e.target.files.length && handleFiles(Array.from(e.target.files))}
                />
                {fileNames.length > 0 && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                    {fileNames.map((n, i) => (
                      <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, background: COLORS.paperDim, borderRadius: 4, padding: "4px 8px", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>
                        <FileSpreadsheet size={12} /> {n}
                      </span>
                    ))}
                    <button onClick={resetUploads} style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: COLORS.rust, cursor: "pointer", fontSize: 12 }}>
                      <X size={12} /> clear
                    </button>
                  </div>
                )}
              </div>

              {error && <div style={{ marginTop: 10, color: COLORS.rust, fontSize: 13 }}>{error}</div>}

              <div style={{ marginTop: 12, fontSize: 12, color: COLORS.inkSoft, lineHeight: 1.5 }}>
                Expected columns — deals file: <em>Deal Owner, Territory, Amount, Deal Stage, Create Date, Close Date, Category, Industry, Imaging Brand, Software Type, Account/Company</em>. Activities file: <em>Owner, Territory, Count</em>. Column names are matched loosely (HubSpot export headers work as-is). Category values are mapped to Net New / Software Solutions / Imaging by keyword — deals without a recognizable category default to Net New. The Account/Company column powers cross-sell detection — without it, cross-sell rate will read 0%. Quotas aren't in most exports — set them below.
              </div>
            </>
          )}

          {source === "manual" && (
            <ManualEntryEditor
              roster={manualRoster}
              setRoster={setManualRoster}
              deals={manualDeals}
              setDeals={setManualDeals}
              quotaMatrix={quotaMatrix}
              setQuotaMatrix={setQuotaMatrix}
              setActualMatrix={setActualMatrix}
              setRetentionMatrix={setRetentionMatrix}
            />
          )}
        </div>


        {/* Ask About Your Data — Q&A over the current dataset via the Claude API */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 4, display: "flex", alignItems: "center", gap: 8 }}>
            <Sparkles size={16} /> Ask About Your Data
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 12, lineHeight: 1.5 }}>
            Ask a plain-language question about the current reps — e.g. "Who has the highest Cybersecurity sales?" or "Who has the best closing rate?" Claude answers using the actual sales, quotas, win rates, and industry/brand/software-type mix currently loaded, nothing outside it.
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <input
              type="text"
              value={qaQuestion}
              onChange={(e) => setQaQuestion(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !qaLoading) askDataQuestion(); }}
              placeholder="Who has the highest Cybersecurity sales?"
              style={{ flex: 1, minWidth: 220, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "8px 12px", fontSize: 13.5 }}
            />
            <button
              onClick={askDataQuestion}
              disabled={qaLoading || !qaQuestion.trim()}
              style={{
                display: "flex", alignItems: "center", gap: 6, background: COLORS.steel, color: "#FFF",
                border: "none", borderRadius: 4, padding: "8px 16px", fontSize: 13.5, fontWeight: 500,
                cursor: qaLoading || !qaQuestion.trim() ? "default" : "pointer", opacity: qaLoading || !qaQuestion.trim() ? 0.6 : 1,
              }}
            >
              {qaLoading ? <Loader2 size={14} className="spin" /> : <Sparkles size={14} />}
              {qaLoading ? "Thinking…" : "Ask"}
            </button>
          </div>
          {qaHistory.length > 0 && (
            <div style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 12 }}>
              {[...qaHistory].reverse().map((qa, i) => (
                <div key={i} style={{ background: COLORS.paperDim, borderRadius: 4, padding: "12px 14px" }}>
                  <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.ink, marginBottom: 4 }}>{qa.question}</div>
                  <div style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 1.5 }}>{qa.answer}</div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Reps: name + territory, editable regardless of source */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink }}>
              Reps
            </div>
            <button
              onClick={addRepGlobal}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${COLORS.steel}`, color: COLORS.steel, borderRadius: 4, padding: "6px 12px", fontSize: 12.5, cursor: "pointer" }}
            >
              <Plus size={14} /> Add Rep
            </button>
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 12 }}>Click a name or territory to edit it — works for mock, uploaded, and manual reps alike. "Add Rep" adds a new rep to whichever tab is currently active, with default quotas and targets ready to edit.</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Object.keys(byRep).length === 0 && <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No reps loaded yet.</div>}
            {Object.values(byRep).map((r) => (
              <div key={r.rep} style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <input
                  defaultValue={r.rep}
                  onBlur={(e) => { const v = e.target.value.trim(); if (v) renameRepGlobal(r.rep, v); }}
                  onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
                  style={{ fontSize: 13, fontWeight: 600, color: COLORS.ink, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "4px 7px", width: 150 }}
                  title="Click to rename this rep"
                />
                <input
                  defaultValue={r.territory}
                  onBlur={(e) => updateTerritoryGlobal(r.rep, e.target.value.trim())}
                  onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
                  placeholder="Territory"
                  style={{ fontSize: 12, color: COLORS.inkSoft, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "4px 7px", width: 120 }}
                />
              </div>
            ))}
          </div>
        </div>

        {/* Quota targets — full matrix: period (Monthly/Quarterly/Annual) x category (Net New/Software/Imaging).
            Assigned once at the start of the year, independent of what's actually been sold. Annual drives
            the scorecard gauges/forecast; any period feeds the Rep Performance chart via the period picker. */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 4 }}>
            Quota Targets
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 14, lineHeight: 1.5 }}>
            Each rep's assigned quota, broken out by period and by category — independent of what they've actually sold. Annual feeds the scorecard gauges below; pick a period in "Rep Performance: Total vs Quota" to compare against that specific target. Retention &amp; Growth targets (NRR, GRR, YoY growth) are annual-only, since a monthly or quarterly target for a rate metric isn't meaningful the way it is for a dollar quota.
          </div>
          {Object.keys(byRep).length === 0 && <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No reps loaded yet.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {Object.values(byRep).map((r) => {
              const m = quotaMatrix[r.rep] || DEFAULT_QUOTA_MATRIX;
              return (
                <div key={r.rep} style={{ borderBottom: `1px solid ${COLORS.paperDim}`, paddingBottom: 14 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 8 }}>{r.rep}</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse" }}>
                      <thead>
                        <tr>
                          <th style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px 0" }}></th>
                          {CATEGORIES.map((c) => (
                            <th key={c.key} style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px" }}>{c.label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {QUOTA_PERIODS.map((p) => (
                          <tr key={p.key}>
                            <td style={{ fontSize: 11.5, color: COLORS.inkSoft, padding: "4px 10px 4px 0", whiteSpace: "nowrap" }}>{p.label}</td>
                            {CATEGORIES.map((c) => (
                              <td key={c.key} style={{ padding: "4px 10px" }}>
                                <input
                                  type="number"
                                  value={(m[p.key] && m[p.key][c.key]) ?? ""}
                                  onChange={(e) =>
                                    setQuotaMatrix((q) => {
                                      const prev = q[r.rep] || cloneMatrix(DEFAULT_QUOTA_MATRIX);
                                      return { ...q, [r.rep]: { ...prev, [p.key]: { ...prev[p.key], [c.key]: parseFloat(e.target.value) || 0 } } };
                                    })
                                  }
                                  style={{ width: 95, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 8px", fontSize: 13, textAlign: "right" }}
                                />
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{ marginTop: 10 }}>
                    <div style={{ fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 4 }}>Retention &amp; Growth Targets (Annual only)</div>
                    <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
                      {RETENTION_METRICS.map((rm) => {
                        const rq = retentionQuota[r.rep] || DEFAULT_RETENTION_QUOTA;
                        return (
                          <label key={rm.key} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: COLORS.inkSoft }}>
                            {rm.label}
                            <input
                              type="number"
                              step="0.1"
                              value={rq[rm.key] ?? ""}
                              onChange={(e) =>
                                setRetentionQuota((q) => ({
                                  ...q,
                                  [r.rep]: { ...(q[r.rep] || DEFAULT_RETENTION_QUOTA), [rm.key]: parseFloat(e.target.value) || 0 },
                                }))
                              }
                              style={{ width: 70, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 8px", fontSize: 13, textAlign: "right" }}
                            />
                            <span style={{ fontSize: 12 }}>{rm.suffix}</span>
                          </label>
                        );
                      })}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Actual sales — manual entry, same period x category shape as quota. When filled in for a
            given period, this overrides the calculated closed-won total in the Rep Performance chart
            for that period. Doesn't touch the scorecards, forecast, or breakdowns below, which stay
            driven by actual deal records. */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink }}>
              Actual Sales (Manual Entry)
            </div>
            {source === "mock" && (
              <button
                onClick={() => {
                  setActualMatrix((prev) => {
                    const next = { ...prev };
                    mockRoster.forEach((r) => {
                      const seeded = mockData.actualMatrix[r.originalName];
                      if (seeded) next[r.displayName] = cloneMatrix(seeded);
                    });
                    return next;
                  });
                }}
                style={{ background: "none", border: `1px solid ${COLORS.steel}`, color: COLORS.steel, borderRadius: 4, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}
              >
                Reset to mock sample data
              </button>
            )}
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 14, lineHeight: 1.5 }}>
            Type in what a rep actually sold, by period and category — this is now the strict source of truth for "sold" everywhere in the app: the scorecard gauges, forecast, coverage ratio, Rep Performance chart, and Closed-Won Breakdown all use these numbers, not deal records. Mock data pre-populates this grid so it's usable right away; edit any cell freely. A rep/period left at 0 shows as 0 sold — nothing falls back to a calculated deal total. Win rate, avg deal cycle, deal count, pipeline, and cross-sell/upsell rate still come from deal records, since those need deal-level detail this grid doesn't capture.
          </div>
          {Object.keys(byRep).length === 0 && <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No reps loaded yet.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {Object.values(byRep).map((r) => {
              const m = actualMatrix[r.rep] || EMPTY_ACTUAL_MATRIX;
              const qm = quotaMatrix[r.rep] || DEFAULT_QUOTA_MATRIX;
              return (
                <div key={r.rep} style={{ borderBottom: `1px solid ${COLORS.paperDim}`, paddingBottom: 14 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 8 }}>{r.rep}</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse" }}>
                      <thead>
                        <tr>
                          <th style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px 0" }}></th>
                          {CATEGORIES.map((c) => (
                            <th key={c.key} style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px" }}>{c.label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {QUOTA_PERIODS.map((p) => (
                          <tr key={p.key}>
                            <td style={{ fontSize: 11.5, color: COLORS.inkSoft, padding: "4px 10px 4px 0", whiteSpace: "nowrap" }}>{p.label}</td>
                            {CATEGORIES.map((c) => {
                              const actualVal = (m[p.key] && m[p.key][c.key]) || 0;
                              const quotaVal = (qm[p.key] && qm[p.key][c.key]) || 0;
                              const pct = quotaVal > 0 ? Math.round((actualVal / quotaVal) * 100) : null;
                              return (
                                <td key={c.key} style={{ padding: "4px 10px" }}>
                                  <input
                                    type="number"
                                    value={(m[p.key] && m[p.key][c.key]) ?? ""}
                                    onChange={(e) =>
                                      setActualMatrix((a) => {
                                        const prev = a[r.rep] || cloneMatrix(EMPTY_ACTUAL_MATRIX);
                                        return { ...a, [r.rep]: { ...prev, [p.key]: { ...prev[p.key], [c.key]: parseFloat(e.target.value) || 0 } } };
                                      })
                                    }
                                    style={{ width: 95, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 8px", fontSize: 13, textAlign: "right" }}
                                  />
                                  {pct != null && (
                                    <div style={{ fontSize: 10, color: pct >= 100 ? COLORS.green : pct >= 75 ? COLORS.amber : COLORS.rust, marginTop: 2 }}>{pct}% of quota</div>
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Retention & growth — NRR, GRR, YoY growth from existing customers. Per period,
            not per category, since these describe the existing base rather than new sales. */}
        <div className="no-print" style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 4 }}>
            Retention &amp; Growth
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 14, lineHeight: 1.5 }}>
            Net Renewal Retention, Gross Renewal Rate, and year-over-year growth from existing customers — tracked per period, not split by category, since these describe the existing base rather than new sales. Mock data pre-populates realistic benchmarks (NRR ~95-112%, GRR ~85-96%, YoY growth ~2-16%); edit any cell.
          </div>
          {Object.keys(byRep).length === 0 && <div style={{ color: COLORS.inkSoft, fontSize: 13 }}>No reps loaded yet.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {Object.values(byRep).map((r) => {
              const m = retentionMatrix[r.rep] || DEFAULT_RETENTION_MATRIX;
              return (
                <div key={r.rep} style={{ borderBottom: `1px solid ${COLORS.paperDim}`, paddingBottom: 14 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 8 }}>{r.rep}</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse" }}>
                      <thead>
                        <tr>
                          <th style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px 0" }}></th>
                          {RETENTION_METRICS.map((rm) => (
                            <th key={rm.key} style={{ textAlign: "left", fontSize: 10.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 10px 6px" }}>{rm.label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {QUOTA_PERIODS.map((p) => (
                          <tr key={p.key}>
                            <td style={{ fontSize: 11.5, color: COLORS.inkSoft, padding: "4px 10px 4px 0", whiteSpace: "nowrap" }}>{p.label}</td>
                            {RETENTION_METRICS.map((rm) => (
                              <td key={rm.key} style={{ padding: "4px 10px" }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={(m[p.key] && m[p.key][rm.key]) ?? ""}
                                    onChange={(e) =>
                                      setRetentionMatrix((q) => {
                                        const prev = q[r.rep] || cloneRetentionMatrix(DEFAULT_RETENTION_MATRIX);
                                        return { ...q, [r.rep]: { ...prev, [p.key]: { ...prev[p.key], [rm.key]: parseFloat(e.target.value) || 0 } } };
                                      })
                                    }
                                    style={{ width: 80, border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 8px", fontSize: 13, textAlign: "right" }}
                                  />
                                  <span style={{ fontSize: 12, color: COLORS.inkSoft }}>{rm.suffix}</span>
                                </div>
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Territory rollup */}
        {territoryChartData.length > 0 && (
          <div style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: "18px 18px 8px", marginBottom: 24 }}>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 6 }}>
              Territory Attainment &amp; Forecast
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={territoryChartData} margin={{ top: 8, right: 12, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 12, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} />
                <YAxis tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} unit="%" />
                <Tooltip formatter={(v) => v + "%"} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} />
                <Bar dataKey="attainment" name="Actual" radius={[3, 3, 0, 0]}>
                  {territoryChartData.map((d, i) => (
                    <Cell key={i} fill={d.attainment >= 100 ? COLORS.green : d.attainment >= 75 ? COLORS.amber : COLORS.rust} />
                  ))}
                </Bar>
                <Bar dataKey="forecastPct" name="Forecast" fill={COLORS.steel} radius={[3, 3, 0, 0]} opacity={0.55} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Rep performance: total closed-won vs quota, side by side */}
        {repChartData.length > 0 && (
          <div style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: "18px 18px 8px", marginBottom: 24 }}>
            <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, flexWrap: "wrap", gap: 8 }}>
              <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink }}>
                Rep Performance: Total vs Quota
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 11.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.4 }}>Quota period</span>
                <select
                  value={quotaPeriod}
                  onChange={(e) => setQuotaPeriod(e.target.value)}
                  style={{ border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "5px 8px", fontSize: 12.5, background: "#FFF" }}
                >
                  {QUOTA_PERIODS.map((p) => (
                    <option key={p.key} value={p.key}>{p.label}</option>
                  ))}
                </select>
              </div>
            </div>
            <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 8 }}>
              Quota comes from the Quota Targets matrix above for the selected period. "Closed Won" is your Actual Sales entry for that rep/period — a rep with nothing entered shows $0, it does not fall back to a calculated deal total.
            </div>
            <ResponsiveContainer width="100%" height={Math.max(220, repChartData.length * 46)}>
              <BarChart data={repChartData} layout="vertical" margin={{ top: 8, right: 24, left: 8, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} tickFormatter={(v) => "$" + Math.round(v / 1000) + "k"} />
                <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} stroke={COLORS.inkSoft} />
                <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} />
                <Bar dataKey="won" name="Closed Won" fill={COLORS.steel} radius={[0, 3, 3, 0]} />
                <Bar dataKey="quota" name={`Quota (${QUOTA_PERIODS.find((p) => p.key === quotaPeriod).label})`} fill={COLORS.paperDim} stroke={COLORS.inkSoft} strokeWidth={1} radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Retention & growth by rep, for the same period picked above */}
        {retentionChartData.length > 0 && (
          <div style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: "18px 18px 8px", marginBottom: 24 }}>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 6 }}>
              Retention &amp; Growth by Rep ({QUOTA_PERIODS.find((p) => p.key === quotaPeriod).label})
            </div>
            <ResponsiveContainer width="100%" height={Math.max(220, retentionChartData.length * 50)}>
              <BarChart data={retentionChartData} layout="vertical" margin={{ top: 8, right: 24, left: 8, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} unit="%" />
                <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} stroke={COLORS.inkSoft} />
                <Tooltip formatter={(v) => v + "%"} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} />
                <Bar dataKey="nrr" name="Net Renewal Retention" fill={COLORS.steel} radius={[0, 3, 3, 0]} />
                <Bar dataKey="grr" name="Gross Renewal Rate" fill={COLORS.amber} radius={[0, 3, 3, 0]} />
                <Bar dataKey="yoyGrowth" name="YoY Growth (Existing)" fill={COLORS.green} radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
        <div className="no-print" style={{ display: "flex", alignItems: "center", gap: 20, marginBottom: 16, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 12.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.4 }}>Territory</span>
            <select
              value={territoryFilter}
              onChange={(e) => setTerritoryFilter(e.target.value)}
              style={{ border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "6px 10px", fontSize: 13, background: "#FFF" }}
            >
              {territoryOptions.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 12.5, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.4 }}>Rep</span>
            <select
              value={repFilter}
              onChange={(e) => setRepFilter(e.target.value)}
              style={{ border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "6px 10px", fontSize: 13, background: "#FFF" }}
            >
              {repOptions.map((r) => (
                <option key={r} value={r}>{r}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Avg deal size trend */}
        {avgDealSizeTrend.length > 1 && (
          <div style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: "18px 18px 8px", marginBottom: 24 }}>
            <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 6 }}>
              Avg Deal Size Trend
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={avgDealSizeTrend} margin={{ top: 8, right: 12, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} vertical={false} />
                <XAxis dataKey="month" tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} />
                <YAxis tick={{ fontSize: 11, fontFamily: "IBM Plex Mono" }} stroke={COLORS.inkSoft} tickFormatter={(v) => "$" + Math.round(v / 1000) + "k"} />
                <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontFamily: "IBM Plex Sans", fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "IBM Plex Sans" }} />
                <Line type="monotone" dataKey="netNew" name="Net New" stroke={COLORS.rust} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                <Line type="monotone" dataKey="software" name="Software Solutions" stroke={COLORS.amber} strokeWidth={2} dot={{ r: 3 }} connectNulls />
                <Line type="monotone" dataKey="imaging" name="Imaging" stroke={COLORS.steel} strokeWidth={2} dot={{ r: 3 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Revenue breakdowns: industry / imaging brand / software type */}
        <div style={{ background: "#FFFFFF", border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: 18, marginBottom: 24 }}>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 18, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, color: COLORS.ink, marginBottom: 4 }}>
            Closed-Won Breakdown
          </div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 12, lineHeight: 1.5 }}>
            Totals here reconcile to each rep's Actual Sales figures — the industry/brand/software-type mix comes from deal records (that detail isn't tracked in Actual Sales), but the dollar totals are scaled to match what's actually entered. A rep/category with nothing entered in Actual Sales shows $0 here too.
          </div>
          <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
            {Object.keys(breakdownLabels).map((k) => (
              <SourceTab key={k} active={breakdownTab === k} onClick={() => setBreakdownTab(k)} label={breakdownLabels[k]} />
            ))}
          </div>
          {breakdownTab === "crossSell" && (
            <div style={{ fontSize: 12, color: COLORS.inkSoft, marginBottom: 10, lineHeight: 1.5 }}>
              Accounts by expansion revenue — closed-won deals sold into an account that already had a closed-won deal, either in a different quota category (cross-sell) or the same category (upsell). Requires an Account/Company column on upload.
            </div>
          )}
          <BreakdownChart data={breakdowns[breakdownTab]} emptyMessage={breakdownTab === "crossSell" ? "No cross-sell deals detected yet — check that deals have an Account/Company column." : undefined} />
        </div>

        {/* Scorecards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: 16 }}>
          {repList.map((r) => (
            <Scorecard
              key={r.rep}
              data={r}
              insight={insights[r.rep]}
              insightLoading={insightLoading[r.rep]}
              onGenerateInsight={() => generateInsight(r)}
              overrides={metricOverrides[r.rep]}
              onSetOverride={(field, value) => setMetricOverride(r.rep, field, value)}
              onClearOverride={(field) => clearMetricOverride(r.rep, field)}
              retention={(retentionMatrix[r.rep] && retentionMatrix[r.rep].annual) || DEFAULT_RETENTION_MATRIX.annual}
            />
          ))}
          {repList.length === 0 && (
            <div style={{ color: COLORS.inkSoft, fontSize: 14, gridColumn: "1 / -1" }}>
              No rep data yet — switch to mock data or upload a file above.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function SourceTab({ active, onClick, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: "8px 16px", borderRadius: 4, fontSize: 13, fontWeight: 500, cursor: "pointer",
        border: `1px solid ${active ? COLORS.steel : COLORS.line}`,
        background: active ? COLORS.steel : "#FFF",
        color: active ? "#FFF" : COLORS.inkSoft,
      }}
    >
      {label}
    </button>
  );
}
