# INEOS Grenadier — Sales Campaign Generator

A web-based sales campaign proposal-and-approval tool, branded to match the
INEOS Grenadier PowerPoint template, covering the full six-stage workflow:
Programme Input → Incentive Packages → Sales Campaign Builder → Proposal
Summary → Campaign Level Summary → QCP Sign-Off — with a two-party approval
process between regional users and a central admin.

**Storage: GitHub, not a database.** Every proposal is one JSON file in a
GitHub repository you choose. Saving is a real commit, loading is a file
read, deleting is a file removal — there's no database service to
provision, connect to, or watch expire. This also means your proposal
history is a normal git log: `git log data/UI__Q3-2026.json` shows every
save, in order, with a timestamp.

**Verified end-to-end:** built a local mock of GitHub's Contents API to test
against (real GitHub credentials aren't available in the environment I
build in), then ran the complete lifecycle through it — create proposal,
add 14 real programmes, submit, approve, build a stacked incentive package
and campaign code, generate the PDF and both Excel exports, delete — with
the actual JSON files landing on disk exactly as they would in a real repo.
The financial totals still match the original workbook exactly:
€1,089,129.41 CoR (VIE) proposal, €577,061.18 Interest/Subvention, to the
cent, same as every version before this one.

## Just 4 real files

```
server.js       everything backend: GitHub storage client, reference data,
                 calculation engine, auth, workflow, PDF + Excel export, all routes
index.html       everything frontend: HTML, CSS, JS, and the Grenadier logo, inlined
package.json      dependencies (express, pdfkit, exceljs — no database driver)
render.yaml       Render deployment blueprint (free tier, single service)
```

## How storage works

Each proposal lives at `data/{MARKET_CODE}__{quarter-slug}.json` in the repo
you configure — e.g. `data/UI__Q3-2026.json`. The file holds only the
genuine decisions (programme titles, costs, volumes, flags, budget figures,
status, feedback); every computed field (campaign codes, EUR conversion,
summaries) is derived fresh on every read, so there's nothing to get out of
sync between what's saved and what's shown.

- **Save a draft:** every add/edit (a programme, a budget figure, a
  package) writes the file immediately — there's no separate "save" step
  to forget. Each write is its own commit, so nothing is ever silently lost.
- **Load it back:** picking a market and quarter reads the file straight
  from GitHub.
- **Delete it:** the "✕ Delete this proposal" link removes the file. A
  region can delete their own draft/rejected/under-review proposals; an
  admin can delete anything.
- **Conflicts:** if two people save the same proposal at nearly the same
  moment, GitHub's API rejects the second write and the app shows a clear
  "this was changed by someone else — reload and try again" message rather
  than silently overwriting anyone's work.

## Setting up GitHub storage

1. **Create (or choose) a repository** to hold the data — can be private,
   and can be the *same* repo the app's source code lives in, or a
   separate one just for data. Either works; a separate repo keeps the
   commit history focused on proposal activity rather than mixing it with
   code changes.
2. **Create a Personal Access Token** with write access to that repo:
   GitHub → Settings → Developer settings → Personal access tokens → Fine-grained
   tokens → generate one scoped to just that repository, with
   **Contents: Read and write** permission.
3. **Set three environment variables** on Render (or locally):
   - `GITHUB_TOKEN` — the token from step 2
   - `GITHUB_REPO` — `your-username/your-repo-name`
   - `GITHUB_BRANCH` — defaults to `main`

No other setup needed — the `data/` folder is created automatically on the
first save.

## Roles & passwords

Unrelated to the storage change — still a single shared password for
regional users and a separate one for the central admin:

- `REGION_PASSWORD` — for regional users
- `ADMIN_PASSWORD` — for the central admin (Proposal Tracker, approve/reject)

Both must be set before the app accepts any sign-in. Sessions are held in
memory and last 12 hours.

## How the workflow works

**Stage 1 — Programme Input.** Cost every programme, mark it as
controlled-via-campaign-code or a back-end programme (bonus/accrual),
stackable or not, and consumer-marketed with a channel list.

**Stages 2 & 3 — Incentive Packages / Sales Campaign Builder — locked**
until the proposal is approved, so campaign codes can never be built
against unapproved numbers.

**Stage 4 — Proposal Summary.** All programmes by channel type and
programme type. Includes `Ingest_Programme_Export.xlsx`.

**Stage 5 — Campaign Level Summary.** Every generated campaign code, base
and stacked, with true cost per unit. Includes `Ingest_Code_Export.xlsx`.

**Stage 6 — QCP Sign-Off Summary.** Free-typed monthly budget vs proposal,
CoR (VIE)/Interest split, channel dissection. Also where a region submits,
and where the **Download Proposal PDF** button lives.

### The two-party approval loop

1. Region fills in Budget + Programme Input, clicks **Submit for
   Approval** — that data locks while a decision is pending.
2. Admin reviews on the **Proposal Tracker** (every region's proposals,
   live CoR/Interest figures) and **Approves**, **Rejects**, or sends it
   **Back for Changes** with feedback.
3. **Approved** unlocks Stages 2–3 and locks Programme Input/Budget in
   exchange, so the approved numbers can't drift. **Rejected**/**Under
   Review** reopens Programme Input/Budget for editing and resubmission.

An admin can bypass any lock (fixing a typo without bouncing a whole
proposal); a regional login cannot.

## Branding

Colours, fonts, and the Grenadier wordmark logo are pulled directly from
`INEOS_PowerPoint_Template_Brand_Font_V4.pptx`: red `#FF4638` for primary
actions, with orange/amber/green doubling as the approval status colours.
Brand fonts are PP Neue Montreal + Inter (licensed, not on Google Fonts) —
headings use **Space Grotesk** as the nearest free equivalent, body text
uses **Inter** as-is. The logo is the actual extracted wordmark, inlined as
base64.

## Running locally

```bash
npm install
GITHUB_TOKEN="ghp_..." \
GITHUB_REPO="your-username/your-repo" \
REGION_PASSWORD="your-region-password" \
ADMIN_PASSWORD="your-admin-password" \
npm start
```

Open `http://localhost:3000`.

## Deploying (GitHub + Render)

1. Upload the 4 files (`server.js`, `index.html`, `package.json`,
   `render.yaml`) to a GitHub repo for the *code* — this can be the same
   repo as your *data* repo, or different.
2. [render.com](https://render.com) → **New** → **Blueprint** → connect
   the code repo. Render provisions a single free web service — no
   database this time.
3. Set the four secrets: Render dashboard → your service → **Environment**
   → add `REGION_PASSWORD`, `ADMIN_PASSWORD`, `GITHUB_TOKEN`, `GITHUB_REPO`.
4. If it doesn't load: check the **Logs** tab. The server logs a clear
   warning on boot if `GITHUB_TOKEN`/`GITHUB_REPO` aren't set, and any
   GitHub API error is surfaced as a real error message in the app rather
   than failing silently.

Free-tier trade-off that still applies: a free web service cold-starts
after 15 minutes idle (~30–60s to wake). No more 30-day expiry concern,
though — that was specific to Render's free Postgres, which this version
doesn't use at all.

## What's simplified vs. the original ask

- **Marketing channels** aren't specified in the source workbook — a
  reasonable standard list is used (Digital & Social, Email/CRM, Dealer
  POS, Direct Mail, Print/OOH, TV/Radio, Events, Retailer Website). Edit
  `MARKETING_CHANNELS` in `server.js`.
- **The admin Proposal Tracker** fetches every proposal file to compute
  its figures (no separate index file to keep in sync) — fine at the
  scale of a quarterly planning tool (tens of proposals), would need an
  index file to stay fast at hundreds+.
- **No per-user accounts or audit trail beyond the free-text
  decided_by/feedback field** — if a hard requirement for *which* admin
  made each call (not just a name they typed) matters, that needs real
  accounts, which is a bigger change than this one.
- Still not reproduced: the workbook's `Add. Sign-Off Summary`/`Business
  Case Tool` tabs and the full `Payout Matrix` margin fields.
