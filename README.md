# INEOS Grenadier — Sales Campaign Generator

A web-based sales campaign proposal-and-approval tool, branded to match the
INEOS Grenadier PowerPoint template, covering the full six-stage workflow:
Programme Input → Incentive Packages → Sales Campaign Builder → Proposal
Summary → Campaign Level Summary → QCP Sign-Off — with a two-party approval
process between regional users and a central admin.

**Storage: GitHub, not a database.** Every proposal is one JSON file in a
GitHub repository. Saving is a real commit, loading is a file read, deleting
is a file removal.

**Verified end-to-end** against a local mock of GitHub's Contents API (real
GitHub credentials aren't available in the environment this was built in):
login → create proposal → 14 real programmes → submit → approve → build a
stacked incentive package and campaign code → generate the PDF → delete —
with the actual JSON files landing on disk exactly as they would in a real
repo. Financial totals still match the original workbook exactly:
€1,089,129.41 CoR (VIE) proposal, €577,061.18 Interest/Subvention, to the cent.

## Signing in

One shared password for everyone. At sign-in, pick your region from:
**APAC, EUROPE, UKI, AMERICAS, MEA, CHINA** — or **Admin** for cross-region
access (Proposal Tracker, approve/reject). Set the password once as an
environment variable:

- `APP_PASSWORD` — the one password, for every region and Admin alike

There's no built-in default — the app won't accept any sign-in until this
is set.

**Region is a real access boundary, not just a filter.** A signed-in region
can only see and edit its own proposals; trying to reach another region's
proposal returns a clear "you can only access proposals for your own
region" error. Admin has no home region and can reach any of them.

## About that empty-response bug

The "server returned an empty response" error from earlier testing should
be resolved now — not just retried around, but actually fixed. The pattern
(worked sometimes, failed consistently other times, always a 200 status
with a genuinely empty body) is the signature of a well-known Node-behind-
a-reverse-proxy issue: Node's default keep-alive timeout (5 seconds) is
shorter than most proxies' (Render's included), so the proxy can reuse a
connection Node has already started closing. `server.js` now explicitly
sets `keepAliveTimeout` and `headersTimeout` above any reasonable proxy
timeout, which is the standard fix for this exact symptom. The automatic
one-time retry from before is still in place too, as a second line of
defence.

## Just 4 real files

```
server.js       everything backend: GitHub storage client, reference data,
                 calculation engine, auth, workflow, PDF + Excel export, all routes
index.html       everything frontend: HTML, CSS, JS, and the Grenadier logo, inlined
package.json      dependencies (express, pdfkit, exceljs — no database driver)
render.yaml       Render deployment blueprint (free tier, single service)
```

## How storage works

Each proposal lives at `data/{REGION}__{quarter-slug}.json` — e.g.
`data/UKI__Q3-2026.json`. The file holds only the genuine decisions
(programme titles, costs, volumes, flags, budget figures, status,
feedback); every computed field (campaign codes, EUR conversion,
summaries) is derived fresh on every read.

- **Save a draft:** every add/edit writes the file immediately as its own
  commit — no separate "save" step to forget, nothing silently lost.
- **Load it back:** picking a quarter reads the file straight from GitHub.
- **Delete it:** the "✕ Delete this proposal" link removes the file. A
  region can delete their own draft/rejected/under-review proposals; an
  admin can delete anything.
- **Conflicts:** if two people save the same proposal at nearly the same
  moment, GitHub's API rejects the second write and the app shows a clear
  "this was changed by someone else — reload and try again" message.

## Setting up GitHub storage

1. **Create (or choose) a repository** to hold the data.
2. **Create a Personal Access Token** with write access to it: GitHub →
   Settings → Developer settings → Personal access tokens → Fine-grained
   tokens → scoped to just that repo, **Contents: Read and write**.
3. **Set three environment variables:**
   - `GITHUB_TOKEN` — the token from step 2
   - `GITHUB_REPO` — `your-username/your-repo-name`
   - `GITHUB_BRANCH` — defaults to `main`

The `data/` folder is created automatically on the first save.

## How the workflow works

**Stage 1 — Programme Input.** Cost every programme, mark it as
controlled-via-campaign-code or a back-end programme (bonus/accrual),
stackable or not, and consumer-marketed with a channel list.

**Stages 2 & 3 — Incentive Packages / Sales Campaign Builder — locked**
until the proposal is approved.

**Stage 4 — Proposal Summary.** All programmes by channel type and
programme type. Includes `Ingest_Programme_Export.xlsx`.

**Stage 5 — Campaign Level Summary.** Every generated campaign code, base
and stacked. Includes `Ingest_Code_Export.xlsx`.

**Stage 6 — QCP Sign-Off Summary.** Free-typed monthly budget vs proposal,
CoR (VIE)/Interest split, channel dissection, submit button, PDF download.

### The two-party approval loop

1. Region fills in Budget + Programme Input, clicks **Submit for
   Approval** — locks while a decision is pending.
2. Admin reviews on the **Proposal Tracker** (every region's proposals,
   live CoR/Interest figures) and **Approves**, **Rejects**, or sends it
   **Back for Changes** with feedback.
3. **Approved** unlocks Stages 2–3 and locks Programme Input/Budget in
   exchange. **Rejected**/**Under Review** reopens Programme Input/Budget
   for editing and resubmission.

## Branding

Colours, fonts, and the Grenadier wordmark logo are pulled directly from
`INEOS_PowerPoint_Template_Brand_Font_V4.pptx`: red `#FF4638` for primary
actions, orange/amber/green doubling as approval status colours. Headings
use **Space Grotesk** (nearest free equivalent to the licensed PP Neue
Montreal); body text uses **Inter** (the brand font, free as-is).

## Running locally

```bash
npm install
GITHUB_TOKEN="ghp_..." \
GITHUB_REPO="your-username/your-repo" \
APP_PASSWORD="your-password" \
npm start
```

Open `http://localhost:3000`.

## Deploying (GitHub + Render)

1. Upload the 4 files to a GitHub repo for the code.
2. [render.com](https://render.com) → **New** → **Blueprint** → connect
   the code repo. Provisions a single free web service — no database.
3. Set the three secrets: Render dashboard → your service →
   **Environment** → add `APP_PASSWORD`, `GITHUB_TOKEN`, `GITHUB_REPO`.
4. If it doesn't load: check the **Logs** tab. The server logs a clear
   warning on boot if `GITHUB_TOKEN`/`GITHUB_REPO` aren't set, and any
   GitHub API error is surfaced as a real error message in the app.

Free-tier trade-off: a free web service cold-starts after 15 minutes idle
(~30–60s to wake).

## What's simplified / assumed

- **Marketing channels** aren't specified in the source workbook — a
  reasonable standard list is used. Edit `MARKETING_CHANNELS` in `server.js`.
- **"Admin" grants access to every region equally** — there's no concept
  of a regional admin who can only approve their own region's proposals.
  If that's needed, it's a straightforward addition to the region-matching
  logic in `server.js`.
- **The admin Proposal Tracker** fetches every proposal file to compute
  its figures — fine at the scale of a quarterly planning tool, would
  want an index file to stay fast at hundreds of proposals.
- Still not reproduced: the workbook's `Add. Sign-Off Summary`/`Business
  Case Tool` tabs and the full `Payout Matrix` margin fields.
