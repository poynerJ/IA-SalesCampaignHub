// Sales Campaign Generator — single-file backend.
//
// Storage: a GitHub repository, not a database. Each market/quarter proposal
// is one JSON file at data/{MARKET_CODE}__{quarter-slug}.json in the repo
// configured by GITHUB_REPO. Every save is a real git commit, so "save draft"
// is just "write the file"; "load" is "read the file"; "delete" is "remove
// the file" — no database service to provision, connect to, or have expire.
//
// Kept to the smallest possible file count (server.js + index.html +
// package.json + render.yaml + README) so the whole project can be
// uploaded/updated through GitHub's web "Upload files" screen without nested
// folders getting flattened or dropped. Reference data and the calculation
// engine — which would normally be separate files — are inlined below in
// clearly marked sections.

const express = require('express');
const path = require('path');
const crypto = require('crypto');
const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');

// ---------------------------------------------------------------------------
// SECTION 0 — Auth
//
// One shared password for everyone (APP_PASSWORD) — no per-user accounts.
// At sign-in the user also picks their region from a fixed list; picking
// "Admin" grants cross-region access (Proposal Tracker, approve/reject).
// Sessions live in memory (a Map), so they're lost on a redeploy/restart,
// which just means signing in again — acceptable for an internal tool.
// ---------------------------------------------------------------------------
const sessions = new Map(); // token -> { role, region, createdAt }
const SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours

function createSession(role, region) {
  const token = crypto.randomBytes(24).toString('hex');
  sessions.set(token, { role, region: region || null, createdAt: Date.now() });
  return token;
}

function getSession(token) {
  const session = sessions.get(token);
  if (!session) return null;
  if (Date.now() - session.createdAt > SESSION_TTL_MS) {
    sessions.delete(token);
    return null;
  }
  return session;
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  const session = token && getSession(token);
  if (!session) return res.status(401).json({ error: 'Not signed in.' });
  req.role = session.role;
  req.region = session.region;
  next();
}

function requireAdmin(req, res, next) {
  if (req.role !== 'admin') return res.status(403).json({ error: 'Admin sign-in required for this action.' });
  next();
}

// A region can edit Programme Input / Budget only while the proposal is
// editable (draft, or sent back for changes). Admins can always edit, e.g.
// to fix a typo without bouncing the whole proposal back.
const EDITABLE_STATUSES = ['draft', 'rejected', 'under_review'];
function assertProposalEditable(quarter, role) {
  if (role === 'admin') return;
  if (!EDITABLE_STATUSES.includes(quarter.status)) {
    const err = new Error(`This proposal is "${quarter.status}" and can't be edited right now.`);
    err.statusCode = 409;
    throw err;
  }
}

// Incentive packages / sales campaigns only unlock once the proposal is approved.
function assertProposalApproved(quarter, role) {
  if (role === 'admin') return;
  if (quarter.status !== 'approved') {
    const err = new Error('This unlocks once the proposal is approved. Submit Programme Input for approval first.');
    err.statusCode = 409;
    throw err;
  }
}
// ---------------------------------------------------------------------------
// SECTION 1 — GitHub-backed storage
//
// A thin client for GitHub's Contents API (github.com/repos/:owner/:repo/contents/:path).
// Requires three environment variables:
//   GITHUB_TOKEN  — a personal access token with write access to the repo
//   GITHUB_REPO   — "owner/repo-name"
//   GITHUB_BRANCH — defaults to "main"
// ---------------------------------------------------------------------------
const GITHUB_API = process.env.GITHUB_API_BASE || 'https://api.github.com';
const GITHUB_REPO = process.env.GITHUB_REPO || '';
const GITHUB_BRANCH = process.env.GITHUB_BRANCH || 'main';
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || '';
const DATA_DIR = 'data';

function assertGithubConfigured() {
  if (!GITHUB_TOKEN || !GITHUB_REPO) {
    const err = new Error('GitHub storage is not configured on the server: set GITHUB_TOKEN and GITHUB_REPO as environment variables.');
    err.statusCode = 500;
    throw err;
  }
}

async function ghFetch(method, apiPath, body) {
  assertGithubConfigured();
  const url = `${GITHUB_API}/repos/${GITHUB_REPO}/contents/${apiPath}`;
  let res;
  try {
    res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${GITHUB_TOKEN}`,
        Accept: 'application/vnd.github+json',
        'User-Agent': 'sales-campaign-generator',
        'Content-Type': 'application/json',
      },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
  } catch (networkErr) {
    const err = new Error('Could not reach GitHub. Check the server has network access and GITHUB_REPO is correct.');
    err.statusCode = 502;
    throw err;
  }
  return res;
}

async function ghErrorFor(res) {
  const body = await res.json().catch(() => ({}));
  const err = new Error(body.message || `GitHub API error (${res.status})`);
  err.statusCode = res.status === 404 ? 404 : 502;
  return err;
}

// "Q3 2026" -> "Q3-2026" — safe for a filename, and cheaply reversible for display.
function slugify(label) {
  return String(label).trim().replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

function fileName(region, quarterSlug) {
  return `${region}__${quarterSlug}.json`;
}

async function loadProposalRaw(region, quarterSlug) {
  const res = await ghFetch('GET', `${DATA_DIR}/${fileName(region, quarterSlug)}?ref=${encodeURIComponent(GITHUB_BRANCH)}`);
  if (res.status === 404) return null;
  if (!res.ok) throw await ghErrorFor(res);
  const json = await res.json();
  const content = Buffer.from(json.content, 'base64').toString('utf8');
  return { data: JSON.parse(content), sha: json.sha };
}

async function saveProposalRaw(region, quarterSlug, data, sha) {
  data.updated_at = new Date().toISOString();
  const content = Buffer.from(JSON.stringify(data, null, 2)).toString('base64');
  const res = await ghFetch('PUT', `${DATA_DIR}/${fileName(region, quarterSlug)}`, {
    message: `Save proposal: ${region} ${data.quarter_label} (${data.status})`,
    content,
    branch: GITHUB_BRANCH,
    ...(sha ? { sha } : {}),
  });
  if (!res.ok) {
    if (res.status === 409) {
      const err = new Error('This proposal was changed by someone else in the meantime. Reload the page and try again.');
      err.statusCode = 409;
      throw err;
    }
    if (res.status === 422) {
      const err = new Error('A proposal with this market and quarter already exists.');
      err.statusCode = 409;
      throw err;
    }
    throw await ghErrorFor(res);
  }
  const json = await res.json();
  return json.content.sha;
}

async function deleteProposalRaw(region, quarterSlug, sha) {
  const res = await ghFetch('DELETE', `${DATA_DIR}/${fileName(region, quarterSlug)}`, {
    message: `Delete proposal: ${region} ${quarterSlug}`,
    sha,
    branch: GITHUB_BRANCH,
  });
  if (!res.ok) throw await ghErrorFor(res);
}

// Lists every proposal file. GitHub doesn't track empty folders, so a fresh
// repo with no proposals yet returns a 404 here — treated as an empty list.
async function listProposalsRaw() {
  const res = await ghFetch('GET', `${DATA_DIR}?ref=${encodeURIComponent(GITHUB_BRANCH)}`);
  if (res.status === 404) return [];
  if (!res.ok) throw await ghErrorFor(res);
  const items = await res.json();
  if (!Array.isArray(items)) return [];
  return items
    .filter((it) => it.type === 'file' && it.name.endsWith('.json'))
    .map((it) => {
      const base = it.name.replace(/\.json$/, '');
      const sep = base.indexOf('__');
      const region = sep === -1 ? base : base.slice(0, sep);
      const quarterSlug = sep === -1 ? '' : base.slice(sep + 2);
      return { region, quarterSlug };
    });
}

async function getProposalOr404(region, quarterSlug) {
  const loaded = await loadProposalRaw(region, quarterSlug);
  if (!loaded) {
    const err = new Error('Proposal not found');
    err.statusCode = 404;
    throw err;
  }
  return loaded;
}

// SECTION 2 — Reference data (seeded from the original workbook: ISO CODE,
// Currency Conversion, Model Coding, and the Programme Input type catalogue)
// ---------------------------------------------------------------------------
// Reference data extracted directly from the original workbook
// ('ISO CODE', 'Currency Conversion', 'Model Coding' and the Programme Input
// type catalogue tabs). These are *not* things a user should have to type in
// by hand each quarter, so they are seeded once and looked up automatically.

const CURRENCY_RATES = {
  AUD: 1.75, BRL: 6.5, CAD: 1.55, CHF: 0.94, DKK: 7.45, GBP: 0.85, HKD: 9,
  IDR: 18900, INR: 100, JPY: 170, KRW: 1600, MXN: 22, MYR: 4.8, NOK: 11.8,
  NZD: 1.9, RMB: 8.3, SEK: 11, SGD: 1.45, THB: 38, TTD: 7.6, TWD: 34,
  USD: 1.15, ZAR: 20.75, EUR: 0,
};

const MODEL_GROUPS = [
  { code: 'G01', model_name: 'Station Wagon (G01)', sort_order: 1 },
  { code: 'G02', model_name: 'Utility Wagon (G02)', sort_order: 2 },
  { code: 'G03', model_name: 'Utility Wagon (G03)', sort_order: 3 },
  { code: 'G08', model_name: '2-Seat Quartermaster (G08)', sort_order: 4 },
  { code: 'G09', model_name: 'Quartermaster (G09)', sort_order: 5 },
  { code: 'G11', model_name: 'Chassis Cab (G11)', sort_order: 6 },
  { code: 'G13', model_name: 'Arcane Works (G13)', sort_order: 7 },
];

const MODEL_YEARS = ['MY23', 'MY24', 'MY25', 'MY26'];

// Not specified in the source workbook (there was no marketing-channel field) —
// a reasonable standard set for a regional automotive campaign. Edit freely.
const MARKETING_CHANNELS = [
  'Digital & Social', 'Email / CRM', 'Dealer POS / In-Store', 'Direct Mail',
  'Print / OOH', 'TV / Radio', 'Events & Experiential', 'Retailer Website',
];

// name -> [code_letter, sort_order, default_budget_alignment, default_channel_type]
const PROGRAMME_TYPES = [
  ['Affinity', 'A', 1, 'CoR (VIE)', 'Fleet'],
  ['Cash Support', 'C', 2, 'CoR (VIE)', 'Retail'],
  ['Courtesy Vehicle/SLV', 'V', 3, 'CoR (VIE)', 'Demo/Courtesy Vehicle/SLV'],
  ['Demonstrator', 'D', 4, 'CoR (VIE)', 'Demo/Courtesy Vehicle/SLV'],
  ['Employee', 'E', 5, 'CoR (VIE)', 'Employee/Alt. Internal'],
  ['Finance (Leasing)', 'L', 6, 'Interest/Subvention (EBITDA)', 'Leasing'],
  ['Finance (Purchase)', 'P', 7, 'Interest/Subvention (EBITDA)', 'Retail (Financial Services)'],
  ['Fleet', 'F', 8, 'CoR (VIE)', 'Fleet'],
  ['Wholesale Support', 'W', 9, 'CoR (VIE)', 'Wholesale'],
  ['Loyalty', 'Z', 10, 'CoR (VIE)', 'Retail'],
  ['Trade-In Support', 'X', 11, 'CoR (VIE)', 'Retail'],
  ['Accessory', 'Y', 12, 'CoR (VIE)', 'Retail'],
  ['Floorplan Support', 'O', 13, 'CoR (VIE)', 'Wholesale'],
  ['Network Bonus', 'N', 14, 'CoR (VIE)', 'Network Bonus'],
  ['Accrual programme', 'R', 15, 'CoR (VIE)', 'Accrual/Discretionary'],
  ['Regional Discretionary', 'Q', 16, 'CoR (VIE)', 'Accrual/Discretionary'],
].map(([name, code_letter, sort_order, default_budget_alignment, default_channel_type]) => ({
  name, code_letter, sort_order, default_budget_alignment, default_channel_type,
}));

// The six regions the business actually plans by. Each gets a short code
// used to build programme/campaign codes (e.g. "UI3P0"), matching the
// pattern from the original workbook (UKI's own code there was "UI").
const REGIONS = [
  { code: 'UI', region: 'UKI', display_name: 'UKI' },
  { code: 'EU', region: 'EUROPE', display_name: 'Europe' },
  { code: 'AP', region: 'APAC', display_name: 'APAC' },
  { code: 'AM', region: 'AMERICAS', display_name: 'Americas' },
  { code: 'ME', region: 'MEA', display_name: 'MEA' },
  { code: 'CN', region: 'CHINA', display_name: 'China' },
];

// ---------------------------------------------------------------------------
// SECTION 3 — Calculation engine (mirrors the workbook's formula logic:
// Programme Input, Incentive Package Input, Sales Campaign Builder,
// Programme Level Summary, Campaign Level Summary, QCP Sign-Off Summary)
// ---------------------------------------------------------------------------
// Calculation engine
//
// This module replicates the formula logic found in the original workbook
// (Programme Input, Incentive Package Input, Sales Campaign Builder,
// Programme Level Summary, Campaign Level Summary and QCP Sign-Off Summary)
// so the same figures come out the other end, from far fewer manual inputs.

/** EUR conversion, mirrors: =IF(currency="EUR", value, value / rate) */
function toEur(localValue, currencyCode, rates = CURRENCY_RATES) {
  const v = Number(localValue) || 0;
  if (currencyCode === 'EUR') return v;
  const rate = rates[currencyCode];
  if (!rate) return null; // mirrors IFERROR(...,"") in the sheet
  return v / rate;
}

/**
 * Programme ID / code, mirrors:
 * =IF(codeRequired="Yes", region & "3" & typeCodeLetter & sequenceNumber, "No Code Required")
 * sequenceNumber = count of programmes of the same type before this one (0-based), within the quarter.
 */
function programmeCode(region, typeCodeLetter, sequenceNumber, codeRequired) {
  if (!codeRequired) return 'No Code Required';
  return `${region}3${typeCodeLetter}${sequenceNumber}`;
}

/**
 * Model eligibility summary string, mirrors the BE:BK TEXTJOIN array formulas:
 * for each model group, join the model-years marked eligible, e.g. "G01:MY23:MY24".
 * modelEligibility shape: { G01: ['MY23','MY24'], G09: ['MY26'] }
 */
function modelSummary(modelEligibility, modelGroups) {
  const parts = [];
  for (const group of modelGroups) {
    const years = (modelEligibility && modelEligibility[group.code]) || [];
    if (years.length) {
      parts.push(`${group.code}: ${years.join(', ')}`);
    }
  }
  return parts.join('  |  ');
}

/**
 * Derive a programme's computed fields (channel type, budget alignment,
 * EUR cost, total programme cost) the way Programme Input's formula columns
 * (I, AR, AT, AY:BA) do.
 */
function deriveProgramme(programme, programmeTypesByName, rates = CURRENCY_RATES) {
  const typeInfo = programmeTypesByName[programme.programme_type] || {};
  const channelType = typeInfo.default_channel_type || '';
  const budgetAlignment = typeInfo.default_budget_alignment || 'CoR (VIE)';
  const eurCostPu = toEur(programme.local_cost_pu, programme.currency, rates);
  const corCostPu = budgetAlignment === 'CoR (VIE)' ? Number(programme.local_cost_pu) || 0 : 0;
  const interestCostPu = budgetAlignment === 'Interest/Subvention (EBITDA)' ? Number(programme.local_cost_pu) || 0 : 0;
  const totalCostLocal = (Number(programme.local_cost_pu) || 0) * (Number(programme.activation_volume) || 0);
  const totalCostEur = eurCostPu !== null ? eurCostPu * (Number(programme.activation_volume) || 0) : null;
  return {
    ...programme,
    channel_type: channelType,
    budget_alignment: budgetAlignment,
    eur_cost_pu: eurCostPu,
    cor_cost_pu: corCostPu,
    interest_cost_pu: interestCostPu,
    total_cost_local: totalCostLocal,
    total_cost_eur: totalCostEur,
  };
}

/**
 * Incentive package roll-up, mirrors Incentive Package Input columns J (name,
 * joined with "and"), K (CoR cost), L (Interest cost), M (total cost).
 */
function derivePackage(pkg, memberProgrammesDerived) {
  const names = memberProgrammesDerived.map((p) => p.title).filter(Boolean);
  let name;
  if (names.length <= 1) {
    name = names.join('');
  } else {
    name = names.slice(0, -1).join(', ') + ' and ' + names[names.length - 1];
  }
  const corCostPu = memberProgrammesDerived.reduce((sum, p) => sum + (p.cor_cost_pu || 0), 0);
  const interestCostPu = memberProgrammesDerived.reduce((sum, p) => sum + (p.interest_cost_pu || 0), 0);
  return {
    ...pkg,
    name,
    cor_cost_pu: corCostPu,
    interest_cost_pu: interestCostPu,
    total_cost_pu: corCostPu + interestCostPu,
  };
}

/**
 * Sales campaign roll-up: base programme + optional incentive package,
 * mirrors Sales Campaign Builder / Matrix Mapping combining a Primary
 * Campaign Category with its stacked Incentive Package.
 */
function deriveCampaign(campaign, baseProgrammeDerived, packageDerived, rates = CURRENCY_RATES) {
  const title = packageDerived
    ? `${baseProgrammeDerived.title}; ${packageDerived.name}`
    : baseProgrammeDerived.title;
  const code = packageDerived
    ? `${baseProgrammeDerived.programme_code}${packageDerived.secondary_code || ''}`
    : baseProgrammeDerived.programme_code;
  const costPuLocal = (Number(baseProgrammeDerived.local_cost_pu) || 0) + (packageDerived ? Number(packageDerived.total_cost_pu) || 0 : 0);
  // Incentive packages are assumed same-currency as their base programme for EUR roll-up purposes.
  const costPuEur = toEur(costPuLocal, baseProgrammeDerived.currency, rates);
  const volume = campaign.forecast_volume != null ? Number(campaign.forecast_volume) : Number(baseProgrammeDerived.activation_volume) || 0;
  return {
    ...campaign,
    campaign_code: code,
    campaign_title: title,
    forecast_volume: volume,
    cost_pu_local: costPuLocal,
    cost_pu_eur: costPuEur,
    total_cost_local: costPuLocal * volume,
    total_cost_eur: costPuEur !== null ? costPuEur * volume : null,
    currency: baseProgrammeDerived.currency,
    channel_type: baseProgrammeDerived.channel_type,
    budget_alignment: baseProgrammeDerived.budget_alignment,
  };
}

/** Programme Level Summary: roll-up by programme (was 'Programme Level Summary'). */
function programmeLevelSummary(derivedProgrammes) {
  const rows = derivedProgrammes.map((p) => ({
    programme_code: p.programme_code,
    title: p.title,
    programme_type: p.programme_type,
    channel_type: p.channel_type,
    budget_alignment: p.budget_alignment,
    currency: p.currency,
    activation_volume: p.activation_volume,
    local_cost_pu: p.local_cost_pu,
    eur_cost_pu: p.eur_cost_pu,
    total_cost_local: p.total_cost_local,
    total_cost_eur: p.total_cost_eur,
  }));
  const totals = rows.reduce(
    (acc, r) => {
      acc.activation_volume += Number(r.activation_volume) || 0;
      acc.total_cost_eur += Number(r.total_cost_eur) || 0;
      return acc;
    },
    { activation_volume: 0, total_cost_eur: 0 }
  );
  // Channel-type split, mirrors the "Campaign Type Split Summaries" block.
  const byChannel = {};
  for (const r of rows) {
    const key = r.channel_type || 'Unclassified';
    if (!byChannel[key]) byChannel[key] = { channel_type: key, activation_volume: 0, total_cost_eur: 0 };
    byChannel[key].activation_volume += Number(r.activation_volume) || 0;
    byChannel[key].total_cost_eur += Number(r.total_cost_eur) || 0;
  }
  // Programme-type split, for the proposal-approval view (Stage 4: "by channel type and programme type").
  const byProgrammeType = {};
  for (const r of rows) {
    const key = r.programme_type || 'Unclassified';
    if (!byProgrammeType[key]) byProgrammeType[key] = { programme_type: key, activation_volume: 0, total_cost_eur: 0 };
    byProgrammeType[key].activation_volume += Number(r.activation_volume) || 0;
    byProgrammeType[key].total_cost_eur += Number(r.total_cost_eur) || 0;
  }
  return { rows, totals, byChannel: Object.values(byChannel), byProgrammeType: Object.values(byProgrammeType) };
}

/** Campaign Level Summary: roll-up by generated sales campaign. */
function campaignLevelSummary(derivedCampaigns) {
  const rows = derivedCampaigns.map((c) => ({
    campaign_code: c.campaign_code,
    campaign_title: c.campaign_title,
    channel_type: c.channel_type,
    budget_alignment: c.budget_alignment,
    currency: c.currency,
    forecast_volume: c.forecast_volume,
    cost_pu_local: c.cost_pu_local,
    cost_pu_eur: c.cost_pu_eur,
    total_cost_local: c.total_cost_local,
    total_cost_eur: c.total_cost_eur,
  }));
  const totals = rows.reduce(
    (acc, r) => {
      acc.forecast_volume += Number(r.forecast_volume) || 0;
      acc.total_cost_eur += Number(r.total_cost_eur) || 0;
      return acc;
    },
    { forecast_volume: 0, total_cost_eur: 0 }
  );
  return { rows, totals };
}

/**
 * QCP Sign-Off Summary: budget vs proposal, mirrors the top of the
 * 'QCP Sign-Off Summary' tab (Budget Summary + Sign-Off Summary blocks).
 */
function qcpSignOffSummary(budgetMonths, derivedProgrammes, marketLabel) {
  const budget = budgetMonths.map((m) => ({
    month_label: m.month_label,
    cor_volume: Number(m.cor_volume) || 0,
    cor_cost_pu_eur: Number(m.cor_cost_pu_eur) || 0,
    cor_total_eur: (Number(m.cor_volume) || 0) * (Number(m.cor_cost_pu_eur) || 0),
    interest_budget_eur: Number(m.interest_budget_eur) || 0,
  }));
  const quarterBudget = budget.reduce(
    (acc, m) => {
      acc.cor_volume += m.cor_volume;
      acc.cor_total_eur += m.cor_total_eur;
      acc.interest_budget_eur += m.interest_budget_eur;
      return acc;
    },
    { cor_volume: 0, cor_total_eur: 0, interest_budget_eur: 0 }
  );

  const corProposalEur = derivedProgrammes
    .filter((p) => p.budget_alignment === 'CoR (VIE)')
    .reduce((sum, p) => sum + (Number(p.total_cost_eur) || 0), 0);
  const interestProposalEur = derivedProgrammes
    .filter((p) => p.budget_alignment === 'Interest/Subvention (EBITDA)')
    .reduce((sum, p) => sum + (Number(p.total_cost_eur) || 0), 0);

  const corVariance = corProposalEur - quarterBudget.cor_total_eur;
  const interestVariance = quarterBudget.interest_budget_eur
    ? interestProposalEur - quarterBudget.interest_budget_eur
    : null; // mirrors the sheet's "-" when no budget figure is set

  // CoR dissection by channel type, mirrors "Regional Programme Sign-Off Summary (CoR Dissection)".
  const corByChannel = {};
  for (const p of derivedProgrammes) {
    if (p.budget_alignment !== 'CoR (VIE)') continue;
    const key = p.channel_type || 'Unclassified';
    if (!corByChannel[key]) corByChannel[key] = { channel_type: key, total_cost_eur: 0, volume: 0 };
    corByChannel[key].total_cost_eur += Number(p.total_cost_eur) || 0;
    corByChannel[key].volume += Number(p.activation_volume) || 0;
  }
  const corDissection = Object.values(corByChannel).map((c) => ({
    ...c,
    weighted_cost_pu_eur: c.volume ? c.total_cost_eur / c.volume : 0,
  }));

  const approvalStatus = (variance) => (variance === null ? '-' : variance > 0 ? 'Requesting' : 'Within Budget');

  return {
    market: marketLabel,
    budget,
    quarterBudget,
    corProposalEur,
    interestProposalEur,
    corVariance,
    interestVariance,
    corDissection,
    corApprovalStatus: approvalStatus(corVariance),
    interestApprovalStatus: approvalStatus(interestVariance),
  };
}
// ---------------------------------------------------------------------------
// SECTION 4 — Proposal enrichment (bridges the raw GitHub JSON <-> calculation engine)
//
// Everything the calculation engine needs (programme codes, EUR conversion,
// summaries) is derived fresh from the raw stored data on every read — the
// stored JSON only ever holds the genuine decisions (title, cost, volume,
// flags), never the computed fields, so there's nothing to get out of sync.
// ---------------------------------------------------------------------------
function enrichProposal(data) {
  const typesByName = {};
  PROGRAMME_TYPES.forEach((t) => (typesByName[t.name] = t));

  const seqByType = {}; // mirrors COUNTIF($H$12:H12,H12)-1 — sequence within type, in array order
  const derivedProgrammes = (data.programmes || []).map((p) => {
    const key = p.programme_type;
    const seq = seqByType[key] || 0;
    seqByType[key] = seq + 1;
    const typeInfo = typesByName[p.programme_type] || {};
    const code = programmeCode(data.region_code, typeInfo.code_letter || '?', seq, p.code_required);
    const derived = deriveProgramme({ ...p, programme_code: code }, typesByName, CURRENCY_RATES);
    derived.model_eligibility_summary = modelSummary(p.model_eligibility, MODEL_GROUPS);
    return derived;
  });
  const programmesById = {};
  derivedProgrammes.forEach((p) => (programmesById[p.id] = p));

  const derivedPackages = (data.incentive_packages || []).map((pkg) => {
    const members = (pkg.programme_ids || []).map((id) => programmesById[id]).filter(Boolean);
    return derivePackage(pkg, members);
  });
  const packagesById = {};
  derivedPackages.forEach((p) => (packagesById[p.id] = p));

  const derivedCampaigns = (data.sales_campaigns || [])
    .map((c) => {
      const base = programmesById[c.base_programme_id];
      if (!base) return null;
      const pkg = c.incentive_package_id ? packagesById[c.incentive_package_id] : null;
      return deriveCampaign(c, base, pkg);
    })
    .filter(Boolean);

  return {
    ...data,
    programmes: derivedProgrammes,
    incentive_packages: derivedPackages,
    sales_campaigns: derivedCampaigns,
    programme_summary: programmeLevelSummary(derivedProgrammes),
    campaign_summary: campaignLevelSummary(derivedCampaigns),
    qcp_summary: (() => {
      const s = qcpSignOffSummary(data.budget_months || [], derivedProgrammes, data.region);
      s.quarterLabel = data.quarter_label;
      return s;
    })(),
  };
}

function nextId(items) {
  return 1 + Math.max(0, ...(items || []).map((i) => Number(i.id) || 0));
}


// ---------------------------------------------------------------------------
// SECTION 5 — Express app + routes
// ---------------------------------------------------------------------------
const app = express();
const PORT = process.env.PORT || 3000;
app.use(express.json());

// ---- Auth (public — no session required) ----
app.post('/api/auth/login', (req, res) => {
  const { region, password } = req.body || {};
  if (!password) return res.status(400).json({ error: 'Password is required.' });
  if (!region) return res.status(400).json({ error: 'Select a region.' });
  if (!process.env.APP_PASSWORD) {
    return res.status(500).json({ error: 'Server is not configured: set APP_PASSWORD.' });
  }
  if (password !== process.env.APP_PASSWORD) {
    return res.status(401).json({ error: 'Incorrect password.' });
  }
  if (region === 'Admin') {
    return res.json({ token: createSession('admin', null), role: 'admin', region: null });
  }
  const match = REGIONS.find((r) => r.region === region);
  if (!match) return res.status(400).json({ error: 'Unrecognised region.' });
  res.json({ token: createSession('region', match.region), role: 'region', region: match.region });
});
app.get('/health', (req, res) => res.json({ ok: true }));

const api = express.Router();
api.use(requireAuth); // everything below requires a signed-in session

api.get('/auth/session', (req, res) => res.json({ role: req.role, region: req.region }));

// ---- Reference data ----
api.get('/reference', (req, res) => {
  res.json({
    currencies: Object.keys(CURRENCY_RATES).sort().map((code) => ({ code, eur_rate: CURRENCY_RATES[code] })),
    regions: REGIONS,
    modelGroups: MODEL_GROUPS,
    modelYears: MODEL_YEARS,
    programmeTypes: PROGRAMME_TYPES,
    marketingChannels: MARKETING_CHANNELS,
    githubConfigured: Boolean(GITHUB_TOKEN && GITHUB_REPO),
  });
});

// ---- Proposal list (powers the market/quarter picker + "+ New proposal") ----
// A region user only ever sees/touches their own region's proposals; an
// admin (no home region) can reach any of them.
function assertRegionMatch(req) {
  if (req.role === 'admin') return;
  if (req.params.region !== req.region) {
    const err = new Error('You can only access proposals for your own region.');
    err.statusCode = 403;
    throw err;
  }
}
api.use('/proposals/:region', (req, res, next) => {
  try { assertRegionMatch(req); next(); } catch (err) { next(err); }
});

api.get('/proposals', async (req, res, next) => {
  try {
    const items = await listProposalsRaw();
    res.json(req.role === 'admin' ? items : items.filter((it) => it.region === req.region));
  } catch (err) { next(err); }
});

api.post('/proposals', async (req, res, next) => {
  try {
    // Region users always create in their own region; only an admin (who has
    // no home region) needs to specify one explicitly.
    const targetRegion = req.role === 'admin' ? req.body.region : req.region;
    const { quarter_label } = req.body;
    if (!targetRegion || !quarter_label) return res.status(400).json({ error: 'region and quarter_label are required' });
    const regionRef = REGIONS.find((r) => r.region === targetRegion);
    if (!regionRef) return res.status(400).json({ error: 'Unrecognised region' });
    const quarterSlug = slugify(quarter_label);
    if (!quarterSlug) return res.status(400).json({ error: 'quarter_label must contain at least one letter or number' });
    const existing = await loadProposalRaw(targetRegion, quarterSlug);
    if (existing) return res.status(409).json({ error: 'A proposal already exists for this region and quarter.' });
    const data = {
      region: regionRef.region, region_code: regionRef.code,
      quarter_label, status: 'draft', submitted_at: null, decided_at: null, decided_by: null, feedback: null,
      created_at: new Date().toISOString(),
      budget_months: [
        { month_label: 'Month 1', month_order: 1, cor_volume: 0, cor_cost_pu_eur: 0, interest_budget_eur: 0 },
        { month_label: 'Month 2', month_order: 2, cor_volume: 0, cor_cost_pu_eur: 0, interest_budget_eur: 0 },
        { month_label: 'Month 3', month_order: 3, cor_volume: 0, cor_cost_pu_eur: 0, interest_budget_eur: 0 },
      ],
      programmes: [], incentive_packages: [], sales_campaigns: [],
    };
    await saveProposalRaw(targetRegion, quarterSlug, data, null);
    res.status(201).json({ region: targetRegion, quarterSlug, ...enrichProposal(data) });
  } catch (err) { next(err); }
});

api.get('/proposals/:region/:quarterSlug', async (req, res, next) => {
  try {
    const loaded = await getProposalOr404(req.params.region, req.params.quarterSlug);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.delete('/proposals/:region/:quarterSlug', async (req, res, next) => {
  try {
    const loaded = await getProposalOr404(req.params.region, req.params.quarterSlug);
    if (req.role !== 'admin' && !EDITABLE_STATUSES.includes(loaded.data.status)) {
      return res.status(409).json({ error: 'Only a draft proposal can be deleted \u2014 ask an admin if an approved or submitted one needs removing.' });
    }
    await deleteProposalRaw(req.params.region, req.params.quarterSlug, loaded.sha);
    res.status(204).end();
  } catch (err) { next(err); }
});

// ---- Budget ----
api.put('/proposals/:region/:quarterSlug/budget/:monthOrder', async (req, res, next) => {
  try {
    const { region, quarterSlug, monthOrder } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalEditable(loaded.data, req.role);
    const month = (loaded.data.budget_months || []).find((m) => String(m.month_order) === String(monthOrder));
    if (!month) return res.status(404).json({ error: 'Budget month not found' });
    const { month_label, cor_volume, cor_cost_pu_eur, interest_budget_eur } = req.body;
    if (month_label !== undefined) month.month_label = month_label;
    if (cor_volume !== undefined) month.cor_volume = cor_volume;
    if (cor_cost_pu_eur !== undefined) month.cor_cost_pu_eur = cor_cost_pu_eur;
    if (interest_budget_eur !== undefined) month.interest_budget_eur = interest_budget_eur;
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

// ---- Programmes (Stage 1) ----
api.post('/proposals/:region/:quarterSlug/programmes', async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalEditable(loaded.data, req.role);
    const { title, programme_type, activation_volume, code_required, stackable, payout_method, currency, local_cost_pu, model_eligibility, marketed, marketing_channels } = req.body;
    if (!title || !programme_type || !currency) return res.status(400).json({ error: 'title, programme_type and currency are required' });
    loaded.data.programmes = loaded.data.programmes || [];
    loaded.data.programmes.push({
      id: nextId(loaded.data.programmes), title, programme_type, activation_volume: activation_volume || 0,
      code_required: code_required !== false, stackable: !!stackable,
      payout_method: payout_method || 'Discount On Invoice', currency, local_cost_pu: local_cost_pu || 0,
      model_eligibility: model_eligibility || {}, marketed: !!marketed, marketing_channels: marketing_channels || [],
    });
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(201).json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.put('/proposals/:region/:quarterSlug/programmes/:id', async (req, res, next) => {
  try {
    const { region, quarterSlug, id } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalEditable(loaded.data, req.role);
    const programme = (loaded.data.programmes || []).find((p) => String(p.id) === String(id));
    if (!programme) return res.status(404).json({ error: 'Programme not found' });
    const fields = ['title', 'programme_type', 'activation_volume', 'code_required', 'stackable', 'payout_method', 'currency', 'local_cost_pu', 'model_eligibility', 'marketed', 'marketing_channels'];
    fields.forEach((f) => { if (req.body[f] !== undefined) programme[f] = req.body[f]; });
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.delete('/proposals/:region/:quarterSlug/programmes/:id', async (req, res, next) => {
  try {
    const { region, quarterSlug, id } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalEditable(loaded.data, req.role);
    loaded.data.programmes = (loaded.data.programmes || []).filter((p) => String(p.id) !== String(id));
    loaded.data.incentive_packages = (loaded.data.incentive_packages || []).filter((pkg) => !(pkg.programme_ids || []).map(String).includes(String(id)));
    loaded.data.sales_campaigns = (loaded.data.sales_campaigns || []).filter((c) => String(c.base_programme_id) !== String(id));
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(204).end();
  } catch (err) { next(err); }
});

// ---- Incentive packages (Stage 2) ----
api.post('/proposals/:region/:quarterSlug/packages', async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalApproved(loaded.data, req.role);
    const { programme_ids, secondary_code } = req.body;
    if (!Array.isArray(programme_ids) || programme_ids.length < 2 || programme_ids.length > 5) {
      return res.status(400).json({ error: 'programme_ids must be an array of 2 to 5 stackable programme ids' });
    }
    const programmesById = {};
    (loaded.data.programmes || []).forEach((p) => (programmesById[p.id] = p));
    const missing = programme_ids.filter((id) => !programmesById[id]);
    if (missing.length) return res.status(400).json({ error: `Programme id(s) ${missing.join(', ')} not found in this proposal.` });
    const notStackable = programme_ids.filter((id) => !programmesById[id].stackable);
    if (notStackable.length) return res.status(400).json({ error: `Programme id(s) ${notStackable.join(', ')} are not marked as stackable.` });
    loaded.data.incentive_packages = loaded.data.incentive_packages || [];
    loaded.data.incentive_packages.push({ id: nextId(loaded.data.incentive_packages), secondary_code: secondary_code || null, programme_ids });
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(201).json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.delete('/proposals/:region/:quarterSlug/packages/:id', async (req, res, next) => {
  try {
    const { region, quarterSlug, id } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalApproved(loaded.data, req.role);
    loaded.data.incentive_packages = (loaded.data.incentive_packages || []).filter((p) => String(p.id) !== String(id));
    loaded.data.sales_campaigns = (loaded.data.sales_campaigns || []).filter((c) => String(c.incentive_package_id) !== String(id));
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(204).end();
  } catch (err) { next(err); }
});

// ---- Sales campaigns (Stage 3) ----
api.post('/proposals/:region/:quarterSlug/campaigns', async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalApproved(loaded.data, req.role);
    const { base_programme_id, incentive_package_id, forecast_volume } = req.body;
    if (!base_programme_id) return res.status(400).json({ error: 'base_programme_id is required' });
    const base = (loaded.data.programmes || []).find((p) => String(p.id) === String(base_programme_id));
    if (!base) return res.status(400).json({ error: 'base_programme_id does not belong to this proposal.' });
    if (incentive_package_id) {
      const pkg = (loaded.data.incentive_packages || []).find((p) => String(p.id) === String(incentive_package_id));
      if (!pkg) return res.status(400).json({ error: 'incentive_package_id does not belong to this proposal.' });
    }
    const volume = (forecast_volume === undefined || forecast_volume === null || forecast_volume === '') ? base.activation_volume : forecast_volume;
    loaded.data.sales_campaigns = loaded.data.sales_campaigns || [];
    loaded.data.sales_campaigns.push({ id: nextId(loaded.data.sales_campaigns), base_programme_id: Number(base_programme_id), incentive_package_id: incentive_package_id ? Number(incentive_package_id) : null, forecast_volume: volume });
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(201).json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.delete('/proposals/:region/:quarterSlug/campaigns/:id', async (req, res, next) => {
  try {
    const { region, quarterSlug, id } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    assertProposalApproved(loaded.data, req.role);
    loaded.data.sales_campaigns = (loaded.data.sales_campaigns || []).filter((c) => String(c.id) !== String(id));
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.status(204).end();
  } catch (err) { next(err); }
});

// ---- Approval workflow ----
api.post('/proposals/:region/:quarterSlug/submit', async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    if (!EDITABLE_STATUSES.includes(loaded.data.status)) {
      return res.status(409).json({ error: `Can't submit \u2014 this proposal is already "${loaded.data.status}".` });
    }
    if (!(loaded.data.programmes || []).length) return res.status(400).json({ error: 'Add at least one programme before submitting.' });
    loaded.data.status = 'submitted';
    loaded.data.submitted_at = new Date().toISOString();
    loaded.data.decided_at = null; loaded.data.decided_by = null; loaded.data.feedback = null;
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

api.post('/proposals/:region/:quarterSlug/decision', requireAdmin, async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const { decision, feedback, decided_by } = req.body;
    if (!['approved', 'rejected', 'under_review'].includes(decision)) {
      return res.status(400).json({ error: 'decision must be approved, rejected, or under_review' });
    }
    const loaded = await getProposalOr404(region, quarterSlug);
    if (loaded.data.status !== 'submitted') {
      return res.status(409).json({ error: `Can't decide on a proposal that is "${loaded.data.status}" (must be submitted).` });
    }
    loaded.data.status = decision;
    loaded.data.decided_at = new Date().toISOString();
    loaded.data.decided_by = decided_by || 'Central Admin';
    loaded.data.feedback = feedback || null;
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

// Region: revert an approved proposal back to draft (only if nothing built against it yet).
api.post('/proposals/:region/:quarterSlug/reopen', async (req, res, next) => {
  try {
    const { region, quarterSlug } = req.params;
    const loaded = await getProposalOr404(region, quarterSlug);
    if (loaded.data.status !== 'approved' && req.role !== 'admin') {
      return res.status(409).json({ error: 'Only an approved proposal can be reopened.' });
    }
    const hasBuilt = (loaded.data.incentive_packages || []).length || (loaded.data.sales_campaigns || []).length;
    if (req.role !== 'admin' && hasBuilt) {
      return res.status(409).json({ error: 'Incentive packages or sales campaigns already exist against this approval \u2014 ask an admin to reopen it.' });
    }
    loaded.data.status = 'draft';
    loaded.data.submitted_at = null; loaded.data.decided_at = null; loaded.data.decided_by = null; loaded.data.feedback = null;
    await saveProposalRaw(region, quarterSlug, loaded.data, loaded.sha);
    res.json(enrichProposal(loaded.data));
  } catch (err) { next(err); }
});

// ---- Admin: cross-market Proposal Tracker ----
api.get('/admin/proposals', requireAdmin, async (req, res, next) => {
  try {
    const items = await listProposalsRaw();
    const results = await Promise.all(items.map(async (item) => {
      const loaded = await loadProposalRaw(item.region, item.quarterSlug).catch(() => null);
      if (!loaded) return null;
      const enriched = enrichProposal(loaded.data);
      return {
        region: item.region, quarter_slug: item.quarterSlug,
        quarter_label: loaded.data.quarter_label,
        status: loaded.data.status, submitted_at: loaded.data.submitted_at, feedback: loaded.data.feedback,
        programme_count: (loaded.data.programmes || []).length,
        cor_proposal_eur: enriched.qcp_summary.corProposalEur,
        interest_proposal_eur: enriched.qcp_summary.interestProposalEur,
        cor_variance_eur: enriched.qcp_summary.corVariance,
      };
    }));
    const filtered = results.filter(Boolean);
    filtered.sort((a, b) => (a.status === 'submitted' ? 0 : 1) - (b.status === 'submitted' ? 0 : 1));
    res.json(filtered);
  } catch (err) { next(err); }
});

// ---- Proposal PDF (generated on demand, always reflects current data) ----
api.get('/proposals/:region/:quarterSlug/pdf', async (req, res, next) => {
  try {
    const loaded = await getProposalOr404(req.params.region, req.params.quarterSlug);
    const data = loaded.data;
    const enriched = enrichProposal(data);
    const summary = enriched.qcp_summary;
    const plSummary = enriched.programme_summary;
    const programmes = enriched.programmes;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${data.region.replace(/\s+/g, '_')}_${data.quarter_label.replace(/\s+/g, '_')}_Proposal.pdf"`);

    const doc = new PDFDocument({ margin: 50, size: 'A4' });
    doc.pipe(res);
    const RED = '#FF4638';
    const GREY = '#7F7D7A';
    const LEFT = 50;
    const RIGHT = 545;
    const fmtEur = (n) => {
      if (n === null || n === undefined) return '\u2014';
      const num = Number(n);
      return (num < 0 ? '-\u20ac' : '\u20ac') + Math.abs(num).toLocaleString(undefined, { maximumFractionDigits: 0 });
    };

    function row(cols, y, opts = {}) {
      const { bold = false, size = 9, color = '#000000' } = opts;
      doc.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(size).fillColor(color);
      let maxHeight = 0;
      cols.forEach(({ text, width, align }) => {
        const h = doc.heightOfString(String(text), { width, align: align || 'left' });
        if (h > maxHeight) maxHeight = h;
      });
      cols.forEach(({ text, x, width, align }) => {
        doc.text(String(text), x, y, { width, align: align || 'left' });
      });
      return y + maxHeight;
    }

    doc.font('Helvetica').fontSize(9).fillColor(GREY).text('INEOS GRENADIER  \u00b7  SALES CAMPAIGN GENERATOR', LEFT, 50, { characterSpacing: 1 });
    doc.font('Helvetica-Bold').fontSize(20).fillColor('#000000').text(`${data.region} \u2014 ${data.quarter_label}`, LEFT, 70);
    doc.font('Helvetica').fontSize(11).fillColor(RED).text('Quarterly Campaign Proposal', LEFT, 96);
    doc.fontSize(9).fillColor(GREY).text(
      `Status: ${data.status.toUpperCase()}` +
      (data.submitted_at ? `   \u00b7   Submitted: ${new Date(data.submitted_at).toLocaleDateString()}` : '') +
      (data.decided_at ? `   \u00b7   Decided: ${new Date(data.decided_at).toLocaleDateString()} by ${data.decided_by || ''}` : ''),
      LEFT, 114
    );

    let y = 150;
    doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Budget vs Proposal', LEFT, y);
    y += 24;
    const budgetRows = [
      ['CoR (VIE) \u2014 Budget', fmtEur(summary.quarterBudget.cor_total_eur)],
      ['CoR (VIE) \u2014 Proposal', fmtEur(summary.corProposalEur)],
      ['CoR (VIE) \u2014 Variance', fmtEur(summary.corVariance)],
      ['Interest/Subvention \u2014 Proposal', fmtEur(summary.interestProposalEur)],
    ];
    budgetRows.forEach(([label, value]) => {
      y = row([{ text: label, x: LEFT, width: 260 }, { text: value, x: LEFT + 260, width: 150 }], y) + 5;
    });

    y += 14;
    doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Spend by Channel Type', LEFT, y);
    y += 24;
    plSummary.byChannel.forEach((c) => {
      y = row([
        { text: c.channel_type, x: LEFT, width: 220 },
        { text: `Vol ${Math.round(c.activation_volume)}`, x: LEFT + 220, width: 80 },
        { text: fmtEur(c.total_cost_eur), x: LEFT + 300, width: 110 },
      ], y) + 5;
    });

    y += 14;
    doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Programme Detail', LEFT, y);
    y += 22;
    const cols = [
      { key: 'code', label: 'Code', x: LEFT, width: 60 },
      { key: 'title', label: 'Title', x: LEFT + 60, width: 190 },
      { key: 'type', label: 'Type', x: LEFT + 250, width: 95 },
      { key: 'vol', label: 'Vol', x: LEFT + 345, width: 35, align: 'right' },
      { key: 'cpu', label: 'Cost/U (EUR)', x: LEFT + 380, width: 65, align: 'right' },
      { key: 'total', label: 'Total (EUR)', x: LEFT + 445, width: 100, align: 'right' },
    ];
    y = row(cols.map((c) => ({ text: c.label, x: c.x, width: c.width, align: c.align })), y, { bold: true, size: 8, color: GREY }) + 4;
    doc.moveTo(LEFT, y).lineTo(RIGHT, y).strokeColor('#D9D6CF').stroke();
    y += 6;
    programmes.forEach((p) => {
      if (y > 760) { doc.addPage(); y = 50; }
      const values = {
        code: p.programme_code, title: p.title || '', type: p.programme_type || '',
        vol: Math.round(p.activation_volume || 0), cpu: fmtEur(p.eur_cost_pu), total: fmtEur(p.total_cost_eur),
      };
      y = row(cols.map((c) => ({ text: values[c.key], x: c.x, width: c.width, align: c.align })), y, { size: 8 }) + 6;
    });

    if (data.feedback) {
      y += 16;
      if (y > 740) { doc.addPage(); y = 50; }
      doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Admin Feedback', LEFT, y);
      y += 20;
      doc.font('Helvetica').fontSize(10).fillColor('#000000').text(data.feedback, LEFT, y, { width: RIGHT - LEFT });
    }

    doc.end();
  } catch (err) { next(err); }
});

// ---- Excel exports (for downstream ingest systems) ----
api.get('/proposals/:region/:quarterSlug/exports/programmes', async (req, res, next) => {
  try {
    const loaded = await getProposalOr404(req.params.region, req.params.quarterSlug);
    const enriched = enrichProposal(loaded.data);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Ingest_Programme_Export');
    ws.columns = [
      { header: 'Market', key: 'market', width: 10 }, { header: 'Quarter', key: 'quarter', width: 10 },
      { header: 'Programme Code', key: 'code', width: 16 }, { header: 'Title', key: 'title', width: 42 },
      { header: 'Programme Type', key: 'type', width: 20 }, { header: 'Channel Type', key: 'channel', width: 24 },
      { header: 'Budget Alignment', key: 'alignment', width: 26 }, { header: 'Currency', key: 'currency', width: 10 },
      { header: 'Local Cost P.U.', key: 'cost_pu', width: 16 }, { header: 'Cost P.U. (EUR)', key: 'cost_pu_eur', width: 16 },
      { header: 'Activation Volume', key: 'volume', width: 16 }, { header: 'Total Cost (EUR)', key: 'total_eur', width: 18 },
      { header: 'Code Required', key: 'code_required', width: 14 }, { header: 'Stackable', key: 'stackable', width: 12 },
      { header: 'Payout Method', key: 'payout', width: 20 }, { header: 'Marketed', key: 'marketed', width: 12 },
      { header: 'Marketing Channels', key: 'marketing_channels', width: 30 }, { header: 'Model Eligibility', key: 'models', width: 30 },
    ];
    ws.getRow(1).font = { bold: true };
    enriched.programmes.forEach((p) => {
      ws.addRow({
        market: loaded.data.region, quarter: loaded.data.quarter_label, code: p.programme_code, title: p.title,
        type: p.programme_type, channel: p.channel_type, alignment: p.budget_alignment, currency: p.currency,
        cost_pu: p.local_cost_pu, cost_pu_eur: p.eur_cost_pu, volume: p.activation_volume, total_eur: p.total_cost_eur,
        code_required: p.code_required ? 'Yes' : 'No', stackable: p.stackable ? 'Yes' : 'No', payout: p.payout_method,
        marketed: p.marketed ? 'Yes' : 'No', marketing_channels: (p.marketing_channels || []).join(', '),
        models: p.model_eligibility_summary || '',
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Ingest_Programme_Export.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) { next(err); }
});

api.get('/proposals/:region/:quarterSlug/exports/codes', async (req, res, next) => {
  try {
    const loaded = await getProposalOr404(req.params.region, req.params.quarterSlug);
    const enriched = enrichProposal(loaded.data);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Ingest_Code_Export');
    ws.columns = [
      { header: 'Market', key: 'market', width: 10 }, { header: 'Quarter', key: 'quarter', width: 10 },
      { header: 'Campaign Code', key: 'code', width: 18 }, { header: 'Campaign Title', key: 'title', width: 50 },
      { header: 'Channel Type', key: 'channel', width: 24 }, { header: 'Budget Alignment', key: 'alignment', width: 26 },
      { header: 'Currency', key: 'currency', width: 10 }, { header: 'Cost P.U. (Local)', key: 'cost_pu', width: 16 },
      { header: 'Cost P.U. (EUR)', key: 'cost_pu_eur', width: 16 }, { header: 'Forecast Volume', key: 'volume', width: 16 },
      { header: 'Total Cost (EUR)', key: 'total_eur', width: 18 },
    ];
    ws.getRow(1).font = { bold: true };
    enriched.sales_campaigns.forEach((c) => {
      ws.addRow({
        market: loaded.data.region, quarter: loaded.data.quarter_label, code: c.campaign_code, title: c.campaign_title,
        channel: c.channel_type, alignment: c.budget_alignment, currency: c.currency, cost_pu: c.cost_pu_local,
        cost_pu_eur: c.cost_pu_eur, volume: c.forecast_volume, total_eur: c.total_cost_eur,
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Ingest_Code_Export.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) { next(err); }
});

app.use('/api', api);

// Single-file frontend — see index.html (inline CSS/JS, no static asset paths to break).
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api')) return next();
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(err.statusCode || 500).json({ error: err.message || 'Internal server error' });
});

const server = app.listen(PORT, () => {
  console.log(`Sales Campaign Generator listening on port ${PORT}`);
  if (!GITHUB_TOKEN || !GITHUB_REPO) {
    console.warn('WARNING: GITHUB_TOKEN and/or GITHUB_REPO are not set — proposal storage will not work until they are configured.');
  }
});

// Node's default keep-alive timeout (5s) is shorter than most reverse proxies'
// (Render's included). When the proxy reuses a connection Node has already
// started tearing down, the client gets back a 200 with an empty body — the
// exact "empty response, even after retrying" symptom seen in testing. Both
// values must exceed the proxy's own keep-alive timeout; headersTimeout must
// exceed keepAliveTimeout or Node's own assertion on startup throws.
server.keepAliveTimeout = 65000;
server.headersTimeout = 66000;
