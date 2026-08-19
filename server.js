// Sales Campaign Generator — single-file backend.
//
// Deliberately kept to the smallest possible file count (server.js + index.html
// + package.json + render.yaml + README) so the whole project can be
// uploaded/updated through GitHub's web "Upload files" screen without nested
// folders getting flattened or dropped. Reference data, the calculation
// engine, and the DB schema — which would normally be separate files — are
// inlined below in clearly marked sections.

const express = require('express');
const path = require('path');
const crypto = require('crypto');
const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');
const { Pool, types } = require('pg');

// pg returns NUMERIC columns as strings by default (to avoid silent precision
// loss). Everything here is plain currency/volume well within float
// precision, so parse to JS numbers at the driver level — otherwise
// arithmetic like `cost + package.total` can silently string-concatenate.
types.setTypeParser(1700, (v) => (v === null ? null : parseFloat(v))); // numeric
types.setTypeParser(20, (v) => (v === null ? null : parseInt(v, 10))); // bigint

const connectionString = process.env.DATABASE_URL || 'postgres://localhost:5432/sales_campaign';
const isLocalDb = /localhost|127\.0\.0\.1/.test(connectionString);
const pool = new Pool({
  connectionString,
  ssl: isLocalDb ? false : { rejectUnauthorized: false },
});
const db = { query: (text, params) => pool.query(text, params) };

// ---------------------------------------------------------------------------
// SECTION 0 — Auth
//
// Simple shared-password gate matching the "regional or admin password"
// pattern: one password for regional users, one for the central admin — no
// per-user accounts. Sessions live in memory (a Map), so they're lost on a
// redeploy/restart, which just means signing in again — acceptable for an
// internal tool. Set REGION_PASSWORD and ADMIN_PASSWORD as environment
// variables before deploying; there is no built-in default password.
// ---------------------------------------------------------------------------
const sessions = new Map(); // token -> { role, createdAt }
const SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours

function createSession(role) {
  const token = crypto.randomBytes(24).toString('hex');
  sessions.set(token, { role, createdAt: Date.now() });
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

async function getQuarterOr404(quarterId) {
  const result = await db.query('SELECT * FROM quarters WHERE id = $1', [quarterId]);
  if (!result.rows[0]) {
    const err = new Error('Quarter not found');
    err.statusCode = 404;
    throw err;
  }
  return result.rows[0];
}

// ---------------------------------------------------------------------------
// SECTION 1 — Database schema
// ---------------------------------------------------------------------------
const SCHEMA_SQL = `
-- Sales Campaign Generator schema
-- Mirrors the logic of the "Q3 2026 UKI Sales Campaign Input File" workbook,
-- generalised to any number of markets and quarters.

CREATE TABLE IF NOT EXISTS reference_currency_rates (
  code            TEXT PRIMARY KEY,          -- ISO currency code, e.g. 'GBP'
  eur_rate        NUMERIC NOT NULL           -- units of this currency per 1 EUR (0 for EUR itself)
);

CREATE TABLE IF NOT EXISTS reference_markets (
  code            TEXT PRIMARY KEY,          -- ISO market code, e.g. 'UI', 'GB'
  market_name     TEXT NOT NULL,
  region          TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS reference_model_groups (
  code            TEXT PRIMARY KEY,          -- e.g. 'G01'
  model_name      TEXT NOT NULL,
  sort_order      INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS reference_programme_types (
  name                       TEXT PRIMARY KEY,   -- e.g. 'Cash Support'
  code_letter                TEXT NOT NULL,       -- single-letter code used in programme IDs
  sort_order                 INTEGER NOT NULL,
  default_budget_alignment   TEXT NOT NULL,       -- 'CoR (VIE)' | 'Interest/Subvention (EBITDA)'
  default_channel_type       TEXT NOT NULL
);

-- A "Market" the user is running a campaign quarter for (was a fixed template
-- of 10 markets in the workbook; here it's an open, addable list).
CREATE TABLE IF NOT EXISTS markets (
  id              SERIAL PRIMARY KEY,
  market_code     TEXT NOT NULL REFERENCES reference_markets(code),
  display_name    TEXT NOT NULL,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- One quarter's working set for a market (e.g. "Q3 2026") — this doubles as
-- "the proposal" for that market/quarter, carrying its approval workflow state.
-- draft -> submitted -> approved (unlocks packages/campaigns)
--                     -> rejected | under_review (region can edit + resubmit)
CREATE TABLE IF NOT EXISTS quarters (
  id              SERIAL PRIMARY KEY,
  market_id       INTEGER NOT NULL REFERENCES markets(id) ON DELETE CASCADE,
  label           TEXT NOT NULL,               -- 'Q3 2026'
  status          TEXT NOT NULL DEFAULT 'draft',
  submitted_at    TIMESTAMPTZ,
  decided_at      TIMESTAMPTZ,
  decided_by      TEXT,
  feedback        TEXT,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(market_id, label)
);

-- Upgrade path for databases created before the approval workflow existed.
ALTER TABLE quarters ADD COLUMN IF NOT EXISTS status TEXT NOT NULL DEFAULT 'draft';
ALTER TABLE quarters ADD COLUMN IF NOT EXISTS submitted_at TIMESTAMPTZ;
ALTER TABLE quarters ADD COLUMN IF NOT EXISTS decided_at TIMESTAMPTZ;
ALTER TABLE quarters ADD COLUMN IF NOT EXISTS decided_by TEXT;
ALTER TABLE quarters ADD COLUMN IF NOT EXISTS feedback TEXT;

-- Monthly Cost-of-Retail budget lines feeding the sign-off summary
-- (was the "Budget Summary" block at the top of QCP Sign-Off Summary).
CREATE TABLE IF NOT EXISTS budget_months (
  id              SERIAL PRIMARY KEY,
  quarter_id      INTEGER NOT NULL REFERENCES quarters(id) ON DELETE CASCADE,
  month_label     TEXT NOT NULL,               -- 'July'
  month_order     INTEGER NOT NULL,            -- 1,2,3
  cor_volume      NUMERIC NOT NULL DEFAULT 0,   -- Cost of Retail (New Vehicle) Volume
  cor_cost_pu_eur NUMERIC NOT NULL DEFAULT 0,   -- Cost of Retail (VIE) P.U. Budget (EUR)
  interest_budget_eur NUMERIC NOT NULL DEFAULT 0, -- Interest/Subvention budget for the month (EUR)
  UNIQUE(quarter_id, month_order)
);

-- Primary Programme Input rows (was the 139-column "Programme Input" grid,
-- reduced to the genuine decision fields; everything else is derived).
CREATE TABLE IF NOT EXISTS programmes (
  id                  SERIAL PRIMARY KEY,
  quarter_id          INTEGER NOT NULL REFERENCES quarters(id) ON DELETE CASCADE,
  row_letter          TEXT,                     -- 'A'..'P' sequence within the quarter, cosmetic
  title               TEXT NOT NULL,             -- 'MY24 5-Seat Utility Wagon £2,500 FDA/Cash Support'
  programme_type      TEXT NOT NULL REFERENCES reference_programme_types(name),
  activation_volume   NUMERIC NOT NULL DEFAULT 0,
  code_required       BOOLEAN NOT NULL DEFAULT true,
  stackable           BOOLEAN NOT NULL DEFAULT false,
  payout_method       TEXT NOT NULL DEFAULT 'Discount On Invoice',
  currency            TEXT NOT NULL REFERENCES reference_currency_rates(code),
  local_cost_pu       NUMERIC NOT NULL DEFAULT 0,
  model_eligibility   JSONB NOT NULL DEFAULT '{}'::jsonb, -- {"G01":["MY23","MY24"], ...}
  marketed            BOOLEAN NOT NULL DEFAULT false,      -- will this programme be consumer-marketed?
  marketing_channels  JSONB NOT NULL DEFAULT '[]'::jsonb,  -- ["Digital & Social", "Dealer POS", ...]
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Upgrade path for databases created before marketing fields existed.
ALTER TABLE programmes ADD COLUMN IF NOT EXISTS marketed BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE programmes ADD COLUMN IF NOT EXISTS marketing_channels JSONB NOT NULL DEFAULT '[]'::jsonb;

-- Incentive Packages: bundle 2-5 stackable programmes together
-- (was "Incentive Package Input" / "IP Builder").
CREATE TABLE IF NOT EXISTS incentive_packages (
  id                  SERIAL PRIMARY KEY,
  quarter_id          INTEGER NOT NULL REFERENCES quarters(id) ON DELETE CASCADE,
  secondary_code      TEXT,                      -- 'A'..'Z' sequence, cosmetic
  programme_ids       INTEGER[] NOT NULL,         -- ordered list of programme ids in this package
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Sales Campaigns: a base/primary programme plus an optional incentive package
-- (was "Sales Campaign Builder" + "Matrix Mapping").
CREATE TABLE IF NOT EXISTS sales_campaigns (
  id                     SERIAL PRIMARY KEY,
  quarter_id             INTEGER NOT NULL REFERENCES quarters(id) ON DELETE CASCADE,
  base_programme_id      INTEGER NOT NULL REFERENCES programmes(id) ON DELETE CASCADE,
  incentive_package_id   INTEGER REFERENCES incentive_packages(id) ON DELETE SET NULL,
  forecast_volume        NUMERIC NOT NULL DEFAULT 0,
  created_at             TIMESTAMPTZ NOT NULL DEFAULT now()
);
`;

// ---------------------------------------------------------------------------
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

const MARKETS = [
  { region: 'AFRICA', market: 'Botswana', code: 'BW' },
  { region: 'AFRICA', market: 'Ghana', code: 'GH' },
  { region: 'AFRICA', market: 'Kenya', code: 'KE' },
  { region: 'AFRICA', market: 'Morocco', code: 'MA' },
  { region: 'AFRICA', market: 'Namibia', code: 'NA' },
  { region: 'AFRICA', market: 'Nigeria', code: 'NG' },
  { region: 'AFRICA', market: 'Senegal', code: 'SN' },
  { region: 'AFRICA', market: 'South Africa', code: 'ZA' },
  { region: 'AFRICA', market: 'Tanzania', code: 'TZ' },
  { region: 'AMERICAS', market: 'Canada', code: 'CA' },
  { region: 'AMERICAS', market: 'Mexico', code: 'MX' },
  { region: 'AMERICAS', market: 'United States of America', code: 'US' },
  { region: 'APAC', market: 'APAC (Indirect)', code: 'AI' },
  { region: 'APAC', market: 'Australia', code: 'AU' },
  { region: 'APAC', market: 'Fiji', code: 'FJ' },
  { region: 'APAC', market: 'Mongolia', code: 'MN' },
  { region: 'APAC', market: 'New Zealand', code: 'NZ' },
  { region: 'APAC', market: 'South Korea', code: 'KR' },
  { region: 'APAC', market: 'Taiwan', code: 'TW' },
  { region: 'CHINA', market: 'China', code: 'CN' },
  { region: 'EUROPE', market: 'EUROPE (Dealer)', code: 'E1' },
  { region: 'EUROPE', market: 'EUROPE (Distributor)', code: 'E2' },
  { region: 'EUROPE', market: 'Austria', code: 'AT' },
  { region: 'EUROPE', market: 'Belgium', code: 'BE' },
  { region: 'EUROPE', market: 'Bosnia and Herzegovina', code: 'BA' },
  { region: 'EUROPE', market: 'Bulgaria', code: 'BG' },
  { region: 'EUROPE', market: 'Croatia', code: 'HR' },
  { region: 'EUROPE', market: 'Czech Republic', code: 'CZ' },
  { region: 'EUROPE', market: 'Finland', code: 'FI' },
  { region: 'EUROPE', market: 'France', code: 'FR' },
  { region: 'EUROPE', market: 'Germany', code: 'DE' },
  { region: 'EUROPE', market: 'Hungary', code: 'HU' },
  { region: 'EUROPE', market: 'Iceland', code: 'IS' },
  { region: 'EUROPE', market: 'Italy', code: 'IT' },
  { region: 'EUROPE', market: 'Luxembourg', code: 'LU' },
  { region: 'EUROPE', market: 'Macedonia', code: 'MK' },
  { region: 'EUROPE', market: 'Monaco', code: 'MC' },
  { region: 'EUROPE', market: 'Netherlands', code: 'NL' },
  { region: 'EUROPE', market: 'Norway', code: 'NO' },
  { region: 'EUROPE', market: 'Poland', code: 'PL' },
  { region: 'EUROPE', market: 'Romania', code: 'RO' },
  { region: 'EUROPE', market: 'Serbia', code: 'RS' },
  { region: 'EUROPE', market: 'Slovakia', code: 'SK' },
  { region: 'EUROPE', market: 'Spain', code: 'ES' },
  { region: 'EUROPE', market: 'Sweden', code: 'SE' },
  { region: 'EUROPE', market: 'Switzerland', code: 'CH' },
  { region: 'EUROPE', market: 'Ukraine', code: 'UA' },
  { region: 'ME', market: 'MIDDLE EAST', code: 'ME' },
  { region: 'ME', market: 'Bahrain', code: 'BH' },
  { region: 'ME', market: 'Kuwait', code: 'KW' },
  { region: 'ME', market: 'Oman', code: 'OM' },
  { region: 'ME', market: 'Qatar', code: 'QA' },
  { region: 'ME', market: 'Saudi Arabia', code: 'SA' },
  { region: 'ME', market: 'United Arab Emirates', code: 'AE' },
  { region: 'UKI', market: 'UKI', code: 'UI' },
  { region: 'UKI', market: 'Ireland', code: 'IE' },
  { region: 'UKI', market: 'United Kingdom', code: 'GB' },
].map(({ region, market, code }) => ({ code, market_name: market, region }));

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
 * =IF(codeRequired="Yes", marketCode & "3" & typeCodeLetter & sequenceNumber, "No Code Required")
 * sequenceNumber = count of programmes of the same type before this one (0-based), within the quarter.
 */
function programmeCode(marketCode, typeCodeLetter, sequenceNumber, codeRequired) {
  if (!codeRequired) return 'No Code Required';
  return `${marketCode}3${typeCodeLetter}${sequenceNumber}`;
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
// SECTION 4 — Migration + reference-data seed (idempotent — safe on every boot)
// ---------------------------------------------------------------------------
async function migrate() {
  await db.query(SCHEMA_SQL);

  for (const [code, eur_rate] of Object.entries(CURRENCY_RATES)) {
    await db.query(
      `INSERT INTO reference_currency_rates (code, eur_rate) VALUES ($1,$2)
       ON CONFLICT (code) DO UPDATE SET eur_rate = EXCLUDED.eur_rate`,
      [code, eur_rate]
    );
  }
  for (const m of MARKETS) {
    await db.query(
      `INSERT INTO reference_markets (code, market_name, region) VALUES ($1,$2,$3)
       ON CONFLICT (code) DO UPDATE SET market_name = EXCLUDED.market_name, region = EXCLUDED.region`,
      [m.code, m.market_name, m.region]
    );
  }
  for (const g of MODEL_GROUPS) {
    await db.query(
      `INSERT INTO reference_model_groups (code, model_name, sort_order) VALUES ($1,$2,$3)
       ON CONFLICT (code) DO UPDATE SET model_name = EXCLUDED.model_name, sort_order = EXCLUDED.sort_order`,
      [g.code, g.model_name, g.sort_order]
    );
  }
  for (const t of PROGRAMME_TYPES) {
    await db.query(
      `INSERT INTO reference_programme_types (name, code_letter, sort_order, default_budget_alignment, default_channel_type)
       VALUES ($1,$2,$3,$4,$5)
       ON CONFLICT (name) DO UPDATE SET code_letter = EXCLUDED.code_letter, sort_order = EXCLUDED.sort_order,
         default_budget_alignment = EXCLUDED.default_budget_alignment, default_channel_type = EXCLUDED.default_channel_type`,
      [t.name, t.code_letter, t.sort_order, t.default_budget_alignment, t.default_channel_type]
    );
  }
  console.log('Migration + reference seed complete.');
}

// ---------------------------------------------------------------------------
// SECTION 5 — Derivation helpers (bridge DB rows <-> calculation engine)
// ---------------------------------------------------------------------------
async function getReferenceMaps() {
  const [typesRes, ratesRes, groupsRes] = await Promise.all([
    db.query('SELECT * FROM reference_programme_types'),
    db.query('SELECT code, eur_rate FROM reference_currency_rates'),
    db.query('SELECT * FROM reference_model_groups ORDER BY sort_order'),
  ]);
  const typesByName = {};
  typesRes.rows.forEach((t) => (typesByName[t.name] = t));
  const rates = {};
  ratesRes.rows.forEach((r) => (rates[r.code] = Number(r.eur_rate)));
  return { typesByName, rates, modelGroups: groupsRes.rows };
}

async function marketCodeForQuarter(quarterId) {
  const result = await db.query(
    `SELECT rm.code FROM quarters q JOIN markets m ON m.id = q.market_id
     JOIN reference_markets rm ON rm.code = m.market_code WHERE q.id = $1`,
    [quarterId]
  );
  return result.rows[0] && result.rows[0].code;
}

async function withDerivedProgrammes(rows, quarterId) {
  const { typesByName, rates, modelGroups } = await getReferenceMaps();
  const marketCode = await marketCodeForQuarter(quarterId);
  const seqByType = {}; // mirrors COUNTIF($H$12:H12,H12)-1 — sequence within type, in creation order
  return rows.map((p) => {
    const key = p.programme_type;
    const seq = seqByType[key] || 0;
    seqByType[key] = seq + 1;
    const typeInfo = typesByName[p.programme_type] || {};
    const code = programmeCode(marketCode, typeInfo.code_letter || '?', seq, p.code_required);
    const derived = deriveProgramme({ ...p, programme_code: code }, typesByName, rates);
    derived.model_eligibility_summary = modelSummary(p.model_eligibility, modelGroups);
    return derived;
  });
}

async function deriveOneProgrammeWithinQuarter(programmeId, quarterId) {
  const result = await db.query('SELECT * FROM programmes WHERE quarter_id = $1 ORDER BY created_at', [quarterId]);
  const all = await withDerivedProgrammes(result.rows, quarterId);
  return all.find((p) => String(p.id) === String(programmeId));
}

async function allDerivedProgrammesByQuarter(quarterId) {
  const result = await db.query('SELECT * FROM programmes WHERE quarter_id = $1 ORDER BY created_at', [quarterId]);
  const derived = await withDerivedProgrammes(result.rows, quarterId);
  const byId = {};
  derived.forEach((p) => (byId[p.id] = p));
  return byId;
}

async function withDerivedPackages(rows, quarterId) {
  const programmesById = await allDerivedProgrammesByQuarter(quarterId);
  return rows.map((pkg) => {
    const members = (pkg.programme_ids || []).map((id) => programmesById[id]).filter(Boolean);
    return derivePackage(pkg, members);
  });
}

async function withDerivedCampaigns(rows, quarterId) {
  const programmesById = await allDerivedProgrammesByQuarter(quarterId);
  const packagesResult = await db.query('SELECT * FROM incentive_packages WHERE quarter_id = $1', [quarterId]);
  const derivedPackages = await withDerivedPackages(packagesResult.rows, quarterId);
  const packagesById = {};
  derivedPackages.forEach((p) => (packagesById[p.id] = p));
  return rows
    .map((c) => {
      const base = programmesById[c.base_programme_id];
      if (!base) return null;
      const pkg = c.incentive_package_id ? packagesById[c.incentive_package_id] : null;
      return deriveCampaign(c, base, pkg);
    })
    .filter(Boolean);
}

const app = express();
const PORT = process.env.PORT || 3000;
app.use(express.json());

// ---- Auth (public — no session required) ----
app.post('/api/auth/login', (req, res) => {
  const { password } = req.body || {};
  if (!password) return res.status(400).json({ error: 'Password is required.' });
  if (!process.env.REGION_PASSWORD && !process.env.ADMIN_PASSWORD) {
    return res.status(500).json({ error: 'Server is not configured: set REGION_PASSWORD and ADMIN_PASSWORD.' });
  }
  if (process.env.ADMIN_PASSWORD && password === process.env.ADMIN_PASSWORD) {
    return res.json({ token: createSession('admin'), role: 'admin' });
  }
  if (process.env.REGION_PASSWORD && password === process.env.REGION_PASSWORD) {
    return res.json({ token: createSession('region'), role: 'region' });
  }
  res.status(401).json({ error: 'Incorrect password.' });
});
app.get('/health', (req, res) => res.json({ ok: true }));

const api = express.Router();
api.use(requireAuth); // everything below requires a signed-in session

api.get('/auth/session', (req, res) => res.json({ role: req.role }));

// ---- Reference data ----
api.get('/reference', async (req, res, next) => {
  try {
    const [currencies, markets, modelGroups, programmeTypes] = await Promise.all([
      db.query('SELECT code, eur_rate FROM reference_currency_rates ORDER BY code'),
      db.query('SELECT code, market_name, region FROM reference_markets ORDER BY region, market_name'),
      db.query('SELECT code, model_name, sort_order FROM reference_model_groups ORDER BY sort_order'),
      db.query('SELECT name, code_letter, sort_order, default_budget_alignment, default_channel_type FROM reference_programme_types ORDER BY sort_order'),
    ]);
    res.json({
      currencies: currencies.rows, markets: markets.rows, modelGroups: modelGroups.rows,
      modelYears: MODEL_YEARS, programmeTypes: programmeTypes.rows, marketingChannels: MARKETING_CHANNELS,
    });
  } catch (err) { next(err); }
});

// ---- Markets ----
api.get('/markets', async (req, res, next) => {
  try {
    const result = await db.query(
      `SELECT m.id, m.market_code, m.display_name, m.created_at, rm.region
       FROM markets m JOIN reference_markets rm ON rm.code = m.market_code
       ORDER BY m.created_at DESC`
    );
    res.json(result.rows);
  } catch (err) { next(err); }
});

api.post('/markets', async (req, res, next) => {
  try {
    const { market_code, display_name } = req.body;
    if (!market_code) return res.status(400).json({ error: 'market_code is required' });
    const result = await db.query(
      `INSERT INTO markets (market_code, display_name) VALUES ($1,$2) RETURNING *`,
      [market_code, display_name || market_code]
    );
    res.status(201).json(result.rows[0]);
  } catch (err) { next(err); }
});

api.delete('/markets/:id', async (req, res, next) => {
  try { await db.query('DELETE FROM markets WHERE id = $1', [req.params.id]); res.status(204).end(); }
  catch (err) { next(err); }
});

// ---- Quarters ----
api.get('/markets/:marketId/quarters', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM quarters WHERE market_id = $1 ORDER BY created_at DESC', [req.params.marketId]);
    res.json(result.rows);
  } catch (err) { next(err); }
});

api.post('/markets/:marketId/quarters', async (req, res, next) => {
  try {
    const { label } = req.body;
    if (!label) return res.status(400).json({ error: 'label is required, e.g. "Q3 2026"' });
    const result = await db.query(`INSERT INTO quarters (market_id, label) VALUES ($1,$2) RETURNING *`, [req.params.marketId, label]);
    const quarterId = result.rows[0].id;
    const defaultLabels = ['Month 1', 'Month 2', 'Month 3'];
    for (let i = 0; i < 3; i++) {
      await db.query(`INSERT INTO budget_months (quarter_id, month_label, month_order) VALUES ($1,$2,$3)`, [quarterId, defaultLabels[i], i + 1]);
    }
    res.status(201).json(result.rows[0]);
  } catch (err) { next(err); }
});

api.delete('/quarters/:id', async (req, res, next) => {
  try { await db.query('DELETE FROM quarters WHERE id = $1', [req.params.id]); res.status(204).end(); }
  catch (err) { next(err); }
});

// ---- Approval workflow ----

// Region: send Programme Input + Budget for admin sign-off.
api.post('/quarters/:id/submit', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.id);
    if (!EDITABLE_STATUSES.includes(quarter.status)) {
      return res.status(409).json({ error: `Can't submit — this proposal is already "${quarter.status}".` });
    }
    const programmes = await db.query('SELECT id FROM programmes WHERE quarter_id = $1', [req.params.id]);
    if (programmes.rows.length === 0) {
      return res.status(400).json({ error: 'Add at least one programme before submitting.' });
    }
    const result = await db.query(
      `UPDATE quarters SET status = 'submitted', submitted_at = now(), decided_at = NULL, decided_by = NULL, feedback = NULL WHERE id = $1 RETURNING *`,
      [req.params.id]
    );
    res.json(result.rows[0]);
  } catch (err) { next(err); }
});

// Admin: approve, reject, or send back for changes with feedback.
api.post('/quarters/:id/decision', requireAdmin, async (req, res, next) => {
  try {
    const { decision, feedback, decided_by } = req.body;
    if (!['approved', 'rejected', 'under_review'].includes(decision)) {
      return res.status(400).json({ error: 'decision must be approved, rejected, or under_review' });
    }
    const quarter = await getQuarterOr404(req.params.id);
    if (quarter.status !== 'submitted') {
      return res.status(409).json({ error: `Can't decide on a proposal that is "${quarter.status}" (must be submitted).` });
    }
    const result = await db.query(
      `UPDATE quarters SET status = $1, decided_at = now(), decided_by = $2, feedback = $3 WHERE id = $4 RETURNING *`,
      [decision, decided_by || 'Central Admin', feedback || null, req.params.id]
    );
    res.json(result.rows[0]);
  } catch (err) { next(err); }
});

// Region: revert an approved proposal back to draft if changes are needed
// before any packages/campaigns have been built against it.
api.post('/quarters/:id/reopen', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.id);
    if (quarter.status !== 'approved' && req.role !== 'admin') {
      return res.status(409).json({ error: 'Only an approved proposal can be reopened.' });
    }
    const built = await db.query(
      'SELECT (SELECT count(*) FROM incentive_packages WHERE quarter_id=$1) AS pkgs, (SELECT count(*) FROM sales_campaigns WHERE quarter_id=$1) AS camps',
      [req.params.id]
    );
    if (req.role !== 'admin' && (Number(built.rows[0].pkgs) > 0 || Number(built.rows[0].camps) > 0)) {
      return res.status(409).json({ error: 'Incentive packages or sales campaigns already exist against this approval — ask an admin to reopen it.' });
    }
    const result = await db.query(
      `UPDATE quarters SET status = 'draft', submitted_at = NULL, decided_at = NULL, decided_by = NULL, feedback = NULL WHERE id = $1 RETURNING *`,
      [req.params.id]
    );
    res.json(result.rows[0]);
  } catch (err) { next(err); }
});

// Admin: cross-market view of every proposal, for the Proposal Tracker / dashboard.
api.get('/admin/quarters', requireAdmin, async (req, res, next) => {
  try {
    const result = await db.query(
      `SELECT q.*, m.display_name AS market_name, m.market_code
       FROM quarters q JOIN markets m ON m.id = q.market_id
       ORDER BY (q.status = 'submitted') DESC, q.submitted_at DESC NULLS LAST, q.created_at DESC`
    );
    const withTotals = await Promise.all(result.rows.map(async (q) => {
      const [budgetResult, programmesById] = await Promise.all([
        db.query('SELECT * FROM budget_months WHERE quarter_id = $1', [q.id]),
        allDerivedProgrammesByQuarter(q.id),
      ]);
      const summary = qcpSignOffSummary(budgetResult.rows, Object.values(programmesById), q.market_name);
      return {
        ...q,
        cor_proposal_eur: summary.corProposalEur,
        interest_proposal_eur: summary.interestProposalEur,
        cor_variance_eur: summary.corVariance,
        programme_count: Object.keys(programmesById).length,
      };
    }));
    res.json(withTotals);
  } catch (err) { next(err); }
});

// ---- Budget months ----
api.get('/quarters/:quarterId/budget', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM budget_months WHERE quarter_id = $1 ORDER BY month_order', [req.params.quarterId]);
    res.json(result.rows);
  } catch (err) { next(err); }
});

api.put('/quarters/:quarterId/budget/:monthId', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    assertProposalEditable(quarter, req.role);
    const { month_label, cor_volume, cor_cost_pu_eur, interest_budget_eur } = req.body;
    const result = await db.query(
      `UPDATE budget_months SET month_label = COALESCE($1, month_label),
        cor_volume = COALESCE($2, cor_volume), cor_cost_pu_eur = COALESCE($3, cor_cost_pu_eur),
        interest_budget_eur = COALESCE($4, interest_budget_eur)
       WHERE id = $5 AND quarter_id = $6 RETURNING *`,
      [month_label, cor_volume, cor_cost_pu_eur, interest_budget_eur, req.params.monthId, req.params.quarterId]
    );
    res.json(result.rows[0]);
  } catch (err) { next(err); }
});

// ---- Programmes ----
api.get('/quarters/:quarterId/programmes', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM programmes WHERE quarter_id = $1 ORDER BY created_at', [req.params.quarterId]);
    res.json(await withDerivedProgrammes(result.rows, req.params.quarterId));
  } catch (err) { next(err); }
});

api.post('/quarters/:quarterId/programmes', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    assertProposalEditable(quarter, req.role);
    const { title, programme_type, activation_volume, code_required, stackable, payout_method, currency, local_cost_pu, model_eligibility, row_letter, marketed, marketing_channels } = req.body;
    if (!title || !programme_type || !currency) return res.status(400).json({ error: 'title, programme_type and currency are required' });
    const result = await db.query(
      `INSERT INTO programmes (quarter_id, row_letter, title, programme_type, activation_volume, code_required, stackable, payout_method, currency, local_cost_pu, model_eligibility, marketed, marketing_channels)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13) RETURNING *`,
      [req.params.quarterId, row_letter || null, title, programme_type, activation_volume || 0, code_required !== false, !!stackable,
        payout_method || 'Discount On Invoice', currency, local_cost_pu || 0, JSON.stringify(model_eligibility || {}),
        !!marketed, JSON.stringify(marketing_channels || [])]
    );
    res.status(201).json(await deriveOneProgrammeWithinQuarter(result.rows[0].id, req.params.quarterId));
  } catch (err) { next(err); }
});

api.put('/programmes/:id', async (req, res, next) => {
  try {
    const existing = await db.query('SELECT quarter_id FROM programmes WHERE id = $1', [req.params.id]);
    if (!existing.rows[0]) return res.status(404).json({ error: 'not found' });
    const quarter = await getQuarterOr404(existing.rows[0].quarter_id);
    assertProposalEditable(quarter, req.role);
    const { title, programme_type, activation_volume, code_required, stackable, payout_method, currency, local_cost_pu, model_eligibility, row_letter, marketed, marketing_channels } = req.body;
    const result = await db.query(
      `UPDATE programmes SET title = COALESCE($1, title), programme_type = COALESCE($2, programme_type),
        activation_volume = COALESCE($3, activation_volume), code_required = COALESCE($4, code_required),
        stackable = COALESCE($5, stackable), payout_method = COALESCE($6, payout_method),
        currency = COALESCE($7, currency), local_cost_pu = COALESCE($8, local_cost_pu),
        model_eligibility = COALESCE($9, model_eligibility), row_letter = COALESCE($10, row_letter),
        marketed = COALESCE($11, marketed), marketing_channels = COALESCE($12, marketing_channels)
       WHERE id = $13 RETURNING *`,
      [title, programme_type, activation_volume, code_required, stackable, payout_method, currency, local_cost_pu,
        model_eligibility ? JSON.stringify(model_eligibility) : null, row_letter,
        marketed === undefined ? null : marketed, marketing_channels ? JSON.stringify(marketing_channels) : null, req.params.id]
    );
    if (!result.rows[0]) return res.status(404).json({ error: 'not found' });
    res.json(await deriveOneProgrammeWithinQuarter(result.rows[0].id, result.rows[0].quarter_id));
  } catch (err) { next(err); }
});

api.delete('/programmes/:id', async (req, res, next) => {
  try {
    const existing = await db.query('SELECT quarter_id FROM programmes WHERE id = $1', [req.params.id]);
    if (existing.rows[0]) {
      const quarter = await getQuarterOr404(existing.rows[0].quarter_id);
      assertProposalEditable(quarter, req.role);
    }
    await db.query('DELETE FROM programmes WHERE id = $1', [req.params.id]);
    res.status(204).end();
  } catch (err) { next(err); }
});

// ---- Incentive packages ----
api.get('/quarters/:quarterId/packages', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM incentive_packages WHERE quarter_id = $1 ORDER BY created_at', [req.params.quarterId]);
    res.json(await withDerivedPackages(result.rows, req.params.quarterId));
  } catch (err) { next(err); }
});

api.post('/quarters/:quarterId/packages', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    assertProposalApproved(quarter, req.role);
    const { programme_ids, secondary_code } = req.body;
    if (!Array.isArray(programme_ids) || programme_ids.length < 2 || programme_ids.length > 5) {
      return res.status(400).json({ error: 'programme_ids must be an array of 2 to 5 stackable programme ids' });
    }
    const memberCheck = await db.query(
      'SELECT id, stackable, quarter_id FROM programmes WHERE id = ANY($1::int[])',
      [programme_ids]
    );
    if (memberCheck.rows.length !== new Set(programme_ids).size) {
      return res.status(400).json({ error: 'One or more programme ids were not found.' });
    }
    const notStackable = memberCheck.rows.filter((r) => !r.stackable);
    if (notStackable.length) {
      return res.status(400).json({ error: `Programme id(s) ${notStackable.map((r) => r.id).join(', ')} are not marked as stackable.` });
    }
    const wrongQuarter = memberCheck.rows.filter((r) => String(r.quarter_id) !== String(req.params.quarterId));
    if (wrongQuarter.length) {
      return res.status(400).json({ error: 'All programmes in a package must belong to this quarter.' });
    }
    const result = await db.query(`INSERT INTO incentive_packages (quarter_id, secondary_code, programme_ids) VALUES ($1,$2,$3) RETURNING *`,
      [req.params.quarterId, secondary_code || null, programme_ids]);
    const [derived] = await withDerivedPackages([result.rows[0]], req.params.quarterId);
    res.status(201).json(derived);
  } catch (err) { next(err); }
});

api.delete('/packages/:id', async (req, res, next) => {
  try { await db.query('DELETE FROM incentive_packages WHERE id = $1', [req.params.id]); res.status(204).end(); }
  catch (err) { next(err); }
});

// ---- Sales campaigns ----
api.get('/quarters/:quarterId/campaigns', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM sales_campaigns WHERE quarter_id = $1 ORDER BY created_at', [req.params.quarterId]);
    res.json(await withDerivedCampaigns(result.rows, req.params.quarterId));
  } catch (err) { next(err); }
});

api.post('/quarters/:quarterId/campaigns', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    assertProposalApproved(quarter, req.role);
    const { base_programme_id, incentive_package_id, forecast_volume } = req.body;
    if (!base_programme_id) return res.status(400).json({ error: 'base_programme_id is required' });
    const baseRow = await db.query('SELECT activation_volume, quarter_id FROM programmes WHERE id = $1', [base_programme_id]);
    if (!baseRow.rows[0] || String(baseRow.rows[0].quarter_id) !== String(req.params.quarterId)) {
      return res.status(400).json({ error: 'base_programme_id does not belong to this quarter.' });
    }
    let volume = forecast_volume;
    if (volume === undefined || volume === null || volume === '') {
      volume = baseRow.rows[0].activation_volume;
    }
    const result = await db.query(
      `INSERT INTO sales_campaigns (quarter_id, base_programme_id, incentive_package_id, forecast_volume) VALUES ($1,$2,$3,$4) RETURNING *`,
      [req.params.quarterId, base_programme_id, incentive_package_id || null, volume]
    );
    const [derived] = await withDerivedCampaigns([result.rows[0]], req.params.quarterId);
    res.status(201).json(derived);
  } catch (err) { next(err); }
});

api.delete('/campaigns/:id', async (req, res, next) => {
  try { await db.query('DELETE FROM sales_campaigns WHERE id = $1', [req.params.id]); res.status(204).end(); }
  catch (err) { next(err); }
});

// ---- Summaries ----
api.get('/quarters/:quarterId/summary/programme-level', async (req, res, next) => {
  try {
    const programmesById = await allDerivedProgrammesByQuarter(req.params.quarterId);
    res.json(programmeLevelSummary(Object.values(programmesById)));
  } catch (err) { next(err); }
});

api.get('/quarters/:quarterId/summary/campaign-level', async (req, res, next) => {
  try {
    const result = await db.query('SELECT * FROM sales_campaigns WHERE quarter_id = $1', [req.params.quarterId]);
    const derived = await withDerivedCampaigns(result.rows, req.params.quarterId);
    res.json(campaignLevelSummary(derived));
  } catch (err) { next(err); }
});

api.get('/quarters/:quarterId/summary/qcp-sign-off', async (req, res, next) => {
  try {
    const [budgetResult, marketInfo] = await Promise.all([
      db.query('SELECT * FROM budget_months WHERE quarter_id = $1 ORDER BY month_order', [req.params.quarterId]),
      db.query(`SELECT m.display_name, q.label FROM quarters q JOIN markets m ON m.id = q.market_id WHERE q.id = $1`, [req.params.quarterId]),
    ]);
    const programmesById = await allDerivedProgrammesByQuarter(req.params.quarterId);
    const marketRow = marketInfo.rows[0];
    const summary = qcpSignOffSummary(budgetResult.rows, Object.values(programmesById), marketRow ? marketRow.display_name : '');
    summary.quarterLabel = marketRow ? marketRow.label : '';
    res.json(summary);
  } catch (err) { next(err); }
});

// ---- Proposal PDF (generated on demand, not stored — always reflects current data) ----
api.get('/quarters/:quarterId/proposal-pdf', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    const marketInfo = await db.query(`SELECT m.display_name FROM quarters q JOIN markets m ON m.id = q.market_id WHERE q.id = $1`, [req.params.quarterId]);
    const marketName = marketInfo.rows[0] ? marketInfo.rows[0].display_name : 'Unknown market';
    const budgetResult = await db.query('SELECT * FROM budget_months WHERE quarter_id = $1 ORDER BY month_order', [req.params.quarterId]);
    const programmesById = await allDerivedProgrammesByQuarter(req.params.quarterId);
    const programmes = Object.values(programmesById);
    const summary = qcpSignOffSummary(budgetResult.rows, programmes, marketName);
    const plSummary = programmeLevelSummary(programmes);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${marketName.replace(/\s+/g, '_')}_${quarter.label.replace(/\s+/g, '_')}_Proposal.pdf"`);

    const doc = new PDFDocument({ margin: 50, size: 'A4' });
    doc.pipe(res);
    const RED = '#FF4638';
    const GREY = '#7F7D7A';
    const LEFT = 50;
    const RIGHT = 545;
    const fmtEur = (n) => {
      if (n === null || n === undefined) return '—';
      const num = Number(n);
      return (num < 0 ? '-€' : '€') + Math.abs(num).toLocaleString(undefined, { maximumFractionDigits: 0 });
    };

    // Renders one row of fixed-position columns at the current y, then
    // advances y by the tallest cell. pdfkit's `continued` mode does NOT
    // create real columns — it just keeps appending text — so every column
    // here gets its own explicit x position instead.
    function row(cols, y, opts = {}) {
      const { bold = false, size = 9, color = '#000000' } = opts;
      doc.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(size).fillColor(color);
      let maxHeight = 0;
      cols.forEach(({ text, x, width, align }) => {
        const h = doc.heightOfString(String(text), { width, align: align || 'left' });
        if (h > maxHeight) maxHeight = h;
      });
      cols.forEach(({ text, x, width, align }) => {
        doc.text(String(text), x, y, { width, align: align || 'left' });
      });
      return y + maxHeight;
    }

    doc.font('Helvetica').fontSize(9).fillColor(GREY).text('INEOS GRENADIER  ·  SALES CAMPAIGN GENERATOR', LEFT, 50, { characterSpacing: 1 });
    doc.font('Helvetica-Bold').fontSize(20).fillColor('#000000').text(`${marketName} — ${quarter.label}`, LEFT, 70);
    doc.font('Helvetica').fontSize(11).fillColor(RED).text('Quarterly Campaign Proposal', LEFT, 96);
    doc.fontSize(9).fillColor(GREY).text(
      `Status: ${quarter.status.toUpperCase()}` +
      (quarter.submitted_at ? `   ·   Submitted: ${new Date(quarter.submitted_at).toLocaleDateString()}` : '') +
      (quarter.decided_at ? `   ·   Decided: ${new Date(quarter.decided_at).toLocaleDateString()} by ${quarter.decided_by || ''}` : ''),
      LEFT, 114
    );

    let y = 150;
    doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Budget vs Proposal', LEFT, y);
    y += 24;
    const budgetRows = [
      ['CoR (VIE) — Budget', fmtEur(summary.quarterBudget.cor_total_eur)],
      ['CoR (VIE) — Proposal', fmtEur(summary.corProposalEur)],
      ['CoR (VIE) — Variance', fmtEur(summary.corVariance)],
      ['Interest/Subvention — Proposal', fmtEur(summary.interestProposalEur)],
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

    if (quarter.feedback) {
      y += 16;
      if (y > 740) { doc.addPage(); y = 50; }
      doc.font('Helvetica-Bold').fontSize(13).fillColor('#000000').text('Admin Feedback', LEFT, y);
      y += 20;
      doc.font('Helvetica').fontSize(10).fillColor('#000000').text(quarter.feedback, LEFT, y, { width: RIGHT - LEFT });
    }

    doc.end();
  } catch (err) { next(err); }

});

// ---- Excel exports (for downstream ingest systems) ----
api.get('/quarters/:quarterId/exports/programmes', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    const marketInfo = await db.query(`SELECT m.display_name, m.market_code FROM quarters q JOIN markets m ON m.id = q.market_id WHERE q.id = $1`, [req.params.quarterId]);
    const market = marketInfo.rows[0];
    const programmesById = await allDerivedProgrammesByQuarter(req.params.quarterId);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Ingest_Programme_Export');
    ws.columns = [
      { header: 'Market', key: 'market', width: 10 },
      { header: 'Quarter', key: 'quarter', width: 10 },
      { header: 'Programme Code', key: 'code', width: 16 },
      { header: 'Title', key: 'title', width: 42 },
      { header: 'Programme Type', key: 'type', width: 20 },
      { header: 'Channel Type', key: 'channel', width: 24 },
      { header: 'Budget Alignment', key: 'alignment', width: 26 },
      { header: 'Currency', key: 'currency', width: 10 },
      { header: 'Local Cost P.U.', key: 'cost_pu', width: 16 },
      { header: 'Cost P.U. (EUR)', key: 'cost_pu_eur', width: 16 },
      { header: 'Activation Volume', key: 'volume', width: 16 },
      { header: 'Total Cost (EUR)', key: 'total_eur', width: 18 },
      { header: 'Code Required', key: 'code_required', width: 14 },
      { header: 'Stackable', key: 'stackable', width: 12 },
      { header: 'Payout Method', key: 'payout', width: 20 },
      { header: 'Marketed', key: 'marketed', width: 12 },
      { header: 'Marketing Channels', key: 'marketing_channels', width: 30 },
      { header: 'Model Eligibility', key: 'models', width: 30 },
    ];
    ws.getRow(1).font = { bold: true };
    Object.values(programmesById).forEach((p) => {
      ws.addRow({
        market: market ? market.market_code : '', quarter: quarter.label, code: p.programme_code, title: p.title,
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

api.get('/quarters/:quarterId/exports/codes', async (req, res, next) => {
  try {
    const quarter = await getQuarterOr404(req.params.quarterId);
    const marketInfo = await db.query(`SELECT m.market_code FROM quarters q JOIN markets m ON m.id = q.market_id WHERE q.id = $1`, [req.params.quarterId]);
    const market = marketInfo.rows[0];
    const result = await db.query('SELECT * FROM sales_campaigns WHERE quarter_id = $1', [req.params.quarterId]);
    const derived = await withDerivedCampaigns(result.rows, req.params.quarterId);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Ingest_Code_Export');
    ws.columns = [
      { header: 'Market', key: 'market', width: 10 },
      { header: 'Quarter', key: 'quarter', width: 10 },
      { header: 'Campaign Code', key: 'code', width: 18 },
      { header: 'Campaign Title', key: 'title', width: 50 },
      { header: 'Channel Type', key: 'channel', width: 24 },
      { header: 'Budget Alignment', key: 'alignment', width: 26 },
      { header: 'Currency', key: 'currency', width: 10 },
      { header: 'Cost P.U. (Local)', key: 'cost_pu', width: 16 },
      { header: 'Cost P.U. (EUR)', key: 'cost_pu_eur', width: 16 },
      { header: 'Forecast Volume', key: 'volume', width: 16 },
      { header: 'Total Cost (EUR)', key: 'total_eur', width: 18 },
    ];
    ws.getRow(1).font = { bold: true };
    derived.forEach((c) => {
      ws.addRow({
        market: market ? market.market_code : '', quarter: quarter.label, code: c.campaign_code, title: c.campaign_title,
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

async function start() {
  try {
    await migrate();
  } catch (err) {
    console.error('Migration failed — check DATABASE_URL. The app will still start, but API calls will fail until this is fixed:', err.message);
  }
  app.listen(PORT, () => console.log(`Sales Campaign Generator listening on port ${PORT}`));
}

start();
