# INEOS Grenadier — Sales Campaign Generator

A web-based sales campaign proposal-and-approval tool, branded to match the
INEOS Grenadier PowerPoint template, covering the full six-stage workflow:
Programme Input → Incentive Packages → Sales Campaign Builder → Proposal
Summary → Campaign Level Summary → QCP Sign-Off — plus a two-party approval
process between regional users and a central admin.

**Verified against the original workbook:** the real UKI Q3 2026 data
reproduces the workbook's cached figures exactly — €1,089,129.41 CoR (VIE)
proposal, €577,061.18 Interest/Subvention proposal, and the full channel-type
dissection, all to the cent — even after everything below was added.

## Still just 4 real files

```
server.js       everything backend: schema, reference data, calc engine, auth,
                 workflow, PDF + Excel export, all routes
index.html       everything frontend: HTML, CSS, JS, and the Grenadier logo, inlined
package.json      dependencies (express, pg, pdfkit, exceljs)
render.yaml       Render deployment blueprint (free tier)
```

Drag these into GitHub's "Add file → Upload files" screen (no folders to
worry about) — updating later is the same drag-and-drop, GitHub replaces by
filename.

## How the workflow works

**Stage 1 — Programme Input.** A regional user costs every programme for the
quarter. Each one can be marked "controlled via a network sales campaign
code" or left as a back-end programme (bonus/accrual payments that don't get
a code), marked stackable, and marked as consumer-marketed with a channel
list.

**Stages 2 & 3 — Incentive Packages / Sales Campaign Builder — locked** until
the proposal is approved. This is deliberate: a region can't build campaign
codes against numbers that haven't been signed off yet.

**Stage 4 — Proposal Summary.** All programmes by channel type and
programme type, for the approval conversation. Includes the
`Ingest_Programme_Export.xlsx` download.

**Stage 5 — Campaign Level Summary.** Row-by-row breakdown of every
generated campaign code (base programmes and stacked base+package
combinations), with true cost per unit. Includes the
`Ingest_Code_Export.xlsx` download.

**Stage 6 — QCP Sign-Off Summary.** Free-typed monthly budget (volume + cost
per unit), rolled up against the proposal to show over/under spend, CoR (VIE)
vs Interest/Subvention split, and a channel-type dissection. This is also
where a region submits the proposal, and where a **Download Proposal PDF**
button lives.

### The two-party approval loop

1. A region fills in **Budget + Programme Input** (Stage 1 + top of Stage 6)
   and clicks **Submit for Approval**. Programme Input and Budget lock —
   nothing can drift while a decision is pending.
2. A central admin (separate password, see below) reviews it on the
   **Proposal Tracker** — every region's proposals in one place, with live
   CoR/Interest figures — and **Approves**, **Rejects**, or sends it **Back
   for Changes** with a feedback note.
3. **Approved** unlocks Stages 2 and 3 for that region — packages and
   campaign codes can now be built — but Programme Input/Budget then lock
   too, so the approved numbers can't be edited out from under the
   approval. **Rejected** or **Under Review** re-opens Programme Input/Budget
   for editing and resubmission, with the admin's feedback shown inline.

An admin can bypass any lock (useful for fixing a typo without bouncing a
whole proposal), which a regional login cannot.

## Roles & passwords

There are no individual user accounts — a single shared password for
regional users, a separate one for the central admin, matching the pattern
from your existing INEOS sales tools ("enter your regional or admin
password"). **You must set both before the app will accept any sign-in:**

- `REGION_PASSWORD` — for regional users
- `ADMIN_PASSWORD` — for the central admin (Proposal Tracker, approve/reject)

Sessions are held in memory and last 12 hours; a redeploy/restart signs
everyone out (fine for an internal tool — just sign in again).

## Branding

Colours, fonts, and the Grenadier wordmark logo are pulled directly from
`INEOS_PowerPoint_Template_Brand_Font_V4.pptx`:

- **Red `#FF4638`** — primary actions, links
- **Orange `#FB6900`**, **Amber `#FFB805`**, **Green `#009B3B`** — these
  double as the approval status colours (submitted/under review/approved)
- Brand fonts are **PP Neue Montreal** (headings) and **Inter** (body) —
  both are licensed/paid fonts not available on Google Fonts, so headings
  use **Space Grotesk** as the nearest free equivalent, and body text uses
  **Inter** as-is (it's already free).
- The logo is the actual wordmark extracted from the template, inlined as
  base64 so there's no separate image file to lose.

I don't currently have a way to view the reference INEOS Sales Campaign
Performance Hub site visually in this environment — the styling here is
built from the brand template file directly, not a visual match to that
site. If you can share a screenshot of it, I can true up anything that's off.

## Running locally

```bash
npm install
DATABASE_URL="postgres://user:pass@localhost:5432/sales_campaign" \
REGION_PASSWORD="your-region-password" \
ADMIN_PASSWORD="your-admin-password" \
npm start
```

Open `http://localhost:3000`. Schema (and any new columns added since your
last deploy) are created/updated automatically on boot.

## Deploying (GitHub + Render)

1. Upload the 4 files to a GitHub repo.
2. [render.com](https://render.com) → **New** → **Blueprint** → connect the
   repo. Render provisions the web service and a free PostgreSQL database
   together.
3. **Set the two passwords:** Render dashboard → your web service →
   **Environment** → add `REGION_PASSWORD` and `ADMIN_PASSWORD` (the
   blueprint marks these `sync: false` deliberately, so they're never
   written into the repo).
4. First deploy runs the migration automatically. If it doesn't load, check
   the **Logs** tab — `Migration failed` means `DATABASE_URL` isn't wired up
   correctly.

Free-tier trade-offs, as before: cold start after 15 minutes idle (~30-60s),
and the free Postgres database expires after 30 days — upgrade it
(~$6/month) before then if this becomes a real production tool.

## What's simplified vs. the original ask

- **Marketing channels** aren't specified anywhere in the source workbook —
  I used a reasonable standard list (Digital & Social, Email/CRM, Dealer
  POS, Direct Mail, Print/OOH, TV/Radio, Events, Retailer Website). Easy to
  edit in `server.js` (`MARKETING_CHANNELS`).
- **The Proposal PDF** is generated fresh on every download (not stored),
  so it always reflects current data — including after approval, once
  packages/campaigns exist.
- **No per-user accounts or audit trail beyond the single
  decided_by/feedback field** on each proposal — if you need to know
  *which* admin made each decision as a hard audit requirement rather than
  a free-text name, that needs real accounts, which is a bigger change.
- Still not reproduced: the workbook's `Add. Sign-Off Summary`/`Business
  Case Tool` tabs and the full `Payout Matrix` margin fields (some already
  broken `#REF!` formulas in the live workbook).
