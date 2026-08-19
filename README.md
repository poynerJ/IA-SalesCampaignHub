# INEOS Grenadier — Sales Campaign Generator

A web-based sales campaign proposal-and-approval tool, branded to match the
INEOS Grenadier PowerPoint template, covering the full six-stage workflow:
Programme Input → Incentive Packages → Sales Campaign Builder → Proposal
Summary → Campaign Level Summary → QCP Sign-Off — with a two-party approval
process between regional users and a central admin.

**No login gate right now.** The password screen was the source of a
persistent, unresolved "empty response" failure in testing — removed for
the time being rather than leaving the app unusable while that's sorted
out separately. Open the app and you're straight in. A **"Working as"**
dropdown in the sidebar (UKI / EUROPE / APAC / AMERICAS / MEA / CHINA /
Admin) lets anyone switch context freely — there's no password and no
verification behind it. See **"Putting the login gate back"** below for
what that means in practice and how to reinstate it later.

**Storage: GitHub, not a database.** Every proposal is one JSON file in a
GitHub repository. Saving is a real commit, loading is a file read,
deleting is a file removal.

**Verified end-to-end** against a local mock of GitHub's Contents API:
create proposal → 14 real programmes → submit → approve → build a stacked
incentive package and campaign code → generate the PDF → delete, with no
headers or tokens involved anywhere. Financial totals still match the
original workbook exactly: €1,089,129.41 CoR (VIE) proposal, €577,061.18
Interest/Subvention, to the cent.

## Putting the login gate back later

The workflow logic (locked stages, submit/approve/reject, region
separation) all still exists and still works — it just now trusts
whichever region/role the client *says* it is, sent as plain `X-Role` /
`X-Region` headers, instead of a verified session token. That means right
now anyone can act as any region or as Admin simply by using the
dropdown — there's no real access control.

To reinstate it: put a real login/session check in front of the
`readContext` middleware in `server.js` (`SECTION 0 — Context`) so
`req.role`/`req.region` come from a verified session again instead of raw
headers. Everything downstream — the workflow locks, the region matching,
every route — is unchanged and doesn't need to be touched.

## Just 4 real files

```
server.js       everything backend: GitHub storage client, reference data,
                 calculation engine, workflow, PDF + Excel export, all routes
index.html       everything frontend: HTML, CSS, JS, and the Grenadier logo, inlined
package.json      dependencies (express, pdfkit, exceljs — no database driver)
render.yaml       Render deployment blueprint (free tier, single service)
```

## How storage works

Each proposal lives at `data/{REGION}__{quarter-slug}.json` — e.g.
`data/UKI__Q3-2026.json`. The file holds only the genuine decisions;
every computed field (campaign codes, EUR conversion, summaries) is
derived fresh on every read.

- **Save a draft:** every add/edit writes the file immediately as its own
  commit.
- **Load it back:** picking a quarter reads the file straight from GitHub.
- **Delete it:** the "✕ Delete this proposal" link removes the file.
- **Conflicts:** if two saves race, GitHub's API rejects the second write
  and the app shows a clear "reload and try again" message.

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

1. Switch "Working as" to a region, fill in Budget + Programme Input,
   click **Submit for Approval** — locks while a decision is pending.
2. Switch "Working as" to **Admin**, open the **Proposal Tracker**, and
   **Approve**, **Reject**, or send it **Back for Changes** with feedback.
3. **Approved** unlocks Stages 2–3 for that region and locks Programme
   Input/Budget in exchange. **Rejected**/**Under Review** reopens them
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
GITHUB_TOKEN="ghp_..." GITHUB_REPO="your-username/your-repo" npm start
```

Open `http://localhost:3000`.

## Deploying (GitHub + Render)

1. Upload the 4 files to a GitHub repo for the code.
2. [render.com](https://render.com) → **New** → **Blueprint** → connect
   the code repo. Provisions a single free web service — no database.
3. Set the two secrets: Render dashboard → your service →
   **Environment** → add `GITHUB_TOKEN`, `GITHUB_REPO`.
4. If it doesn't load: check the **Logs** tab. The server logs a clear
   warning on boot if `GITHUB_TOKEN`/`GITHUB_REPO` aren't set, and any
   GitHub API error is surfaced as a real error message in the app.

Free-tier trade-off: a free web service cold-starts after 15 minutes idle
(~30–60s to wake).

## What's simplified / assumed

- **No access control at all right now** — see "Putting the login gate
  back" above.
- **Marketing channels** aren't specified in the source workbook — a
  reasonable standard list is used. Edit `MARKETING_CHANNELS` in `server.js`.
- **The admin Proposal Tracker** fetches every proposal file to compute
  its figures — fine at the scale of a quarterly planning tool.
- Still not reproduced: the workbook's `Add. Sign-Off Summary`/`Business
  Case Tool` tabs and the full `Payout Matrix` margin fields.
