# CLOUD-SETUP.md — run the fleet without your PC (GitHub Actions, $0/month)

**How it works**

| Repo | Visibility | Contains | Why |
|---|---|---|---|
| `jtracker-fleet` | **public** | scripts + workflow only | Public repos get **unlimited free** Windows runner minutes — no card, no bill, ever |
| `jtracker-state` | **private** | JobTracker.xlsx, config.json, cookies, CVs, Tailored folders, email ledger | Your personal data — never public |

The workflow runs `Run-Pipeline.ps1` 4×/day (06:00 / 12:00 / 18:00 / 00:00 Dubai) on a
GitHub Windows machine: clone state → init → pipeline → push state back. Your PC can stay off.

**Cost: $0.** Public-repo Actions are free with no spending limit and no payment method on file.

---

## One-time setup

### Step 1 — tools (already installed on this PC)
MinGit 2.56 + GitHub CLI 2.103 in `%LOCALAPPDATA%\Programs\` (on your PATH for new terminals).

### Step 2 — sign in to GitHub (YOU do this, in your browser — I never see your password)
The web device flow (`gh auth login --web`) **fails on this network** (broken IPv6 resets
`github.com/login/oauth/*`; api.github.com and git-over-HTTPS are fine). Use the token path:
1. In the browser: GitHub → Settings → Developer settings → Personal access tokens →
   **Tokens (classic)** → Generate new token (classic): **No expiration**, scopes
   **`repo`** + **`workflow`** (both required — no scopes = "token has not been granted
   the required scopes" on repo create).
2. In a terminal: `gh auth login` → GitHub.com → HTTPS → **Paste an authentication token** → paste → Enter.

### Step 3 — create repos + push code *(I run this after Step 2)*
```
gh repo create jtracker-fleet --public  --source . --push
gh repo create jtracker-state --private
```

### Step 4 — seed the private state repo *(I run this)*
Copies into the state repo: `config.json`, `JobTracker.xlsx`, `CVS\` (32 MB),
`Tailored\` (9 MB), `edge-profile\{cookies,feedbot-history,feedbot-rotation}`,
`logs\email-ledger-*.jsonl`, then pushes.

### Step 5 — wire the secrets *(I run this; token is piped, never displayed)*
```
gh secret set STATE_PAT  --repo <you>/jtracker-fleet  < token-from-gh-auth-token
gh secret set STATE_REPO --repo <you>/jtracker-fleet --body "<you>/jtracker-state"
```

### Step 6 — first test run *(I run this)*
```
gh workflow run "JTracker Fleet" --repo <you>/jtracker-fleet
gh run watch --repo <you>/jtracker-fleet
```

### Step 7 — CUTOVER ✅ DONE (2026-11-14, after first fully-green cloud run 43556072909)
Both PC scheduled tasks disabled — cloud is now the only fleet:
```
Disable-ScheduledTask -TaskName SearchMachine2-Pipeline
Disable-ScheduledTask -TaskName HuntBot-min
```
(Re-enable anytime to roll back; the cloud workflow can be disabled in the Actions tab.)

---

## What may need attention later
- **Re-login handoffs:** runner IPs change between runs — LinkedIn/Gmail may occasionally
  demand a verification. When a run log shows that, re-login on your PC, then I re-seed
  cookies (`edge-profile\cookies.json`) and push them to the state repo.
- **Emails/day cap** stays at 8 (ledger lives in the state repo, persists across runs).
- **Viewing results:** `jtracker-state` → `JobTracker.xlsx` on github.com (any device),
  or download it. The OneDrive copy on your PC becomes a read-only snapshot after cutover.
- **60-day keepalive:** handled automatically (empty commit pushed when quiet >6 days).

## Troubleshooting (already hit during go-live)
- **`Init cloud environment` fails with `Cannot bind argument to parameter 'Path' ... empty string`
  and `project=` prints blank:** PS5.1 quirk — `$PSScriptRoot` evaluates **empty** in a
  param *default* when an earlier param is `[Parameter(Mandatory)]`. Fix (applied in both
  cloud scripts): default the param to `""` and resolve `$PSScriptRoot` in the body.
- **`gh` device flow dies with "error connecting to github.com":** see Step 2 (network
  blocks the OAuth endpoint; use token paste instead).
- **First push of personal files:** `.gitignore` must list them BEFORE `git add` — files
  already tracked are NOT un-ignored automatically; use `git rm --cached <file>`.

## Rollback
Re-enable the PC tasks (Step 7) and/or disable the workflow under Actions → JTracker Fleet → … → Disable workflow.
