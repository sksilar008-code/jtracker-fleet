# Search Machine - 2

Your personal job-hunting machine. Three tools + one scheduler, all PowerShell, no Python.

## What's here

| File | What it does |
|---|---|
| `SearchMachine2.ps1` | Searches your own files (PDF, Excel, Word, text, CSV, HTML) for any name/keyword |
| `HuntBot.ps1` | Searches job boards (RemoteOK, Remotive, Arbeitnow) for BIM/Revit jobs and adds new matches to `JobTracker.xlsx` |
| `ApplyBot.ps1` | Opens each job link from the tracker in Edge, fills the form with your profile, uploads your tailored CV + cover letter, and auto-submits |
| `Run-Pipeline.ps1` | Runs HuntBot then ApplyBot — this is what runs 4×/day |
| `Register-ScheduledTask.ps1` | Sets up the 4×/day schedule (run once) |
| `Tracker.psm1` / `EdgeCdp.psm1` | Helper modules (Excel read/write, Edge browser automation) |
| `config.json` | Your profile, search keywords, apply settings |
| `lib\`, `tools\` | PDF library files and Edge driver (don't touch) |
| `logs\` | Run logs + screenshots of every application attempt |
| `reports\` | Hunt reports and search reports |

## Daily use

**Search your files:**
```powershell
.\SearchMachine2.ps1 -Query "AECOM"
```

**Run the whole pipeline manually:**
```powershell
.\Run-Pipeline.ps1            # full run, applies for real
.\Run-Pipeline.ps1 -DryRun    # fills forms but never clicks submit
```

## The 4×/day schedule

Runs every 6 hours starting 06:00 via Windows Task Scheduler (`SearchMachine2-Pipeline`).
```powershell
Start-ScheduledTask -TaskName "SearchMachine2-Pipeline"      # run now
Get-ScheduledTask -TaskName "SearchMachine2-Pipeline"        # check
Unregister-ScheduledTask -TaskName "SearchMachine2-Pipeline" # remove
```
You must be logged in (screen locked is OK) for the browser to run.

## IMPORTANT: add your job-site accounts

Many sites (DrJobs, Indeed, FounditGulf...) require login before applying.
Jobs on those sites get status **"Needs account"**. Add your login once and
ApplyBot will use it automatically on the next run:

```powershell
Import-Module .\Tracker.psm1
Add-TrackerAccount -Platform "DrJobs" -Site "drjobs.ae" -Username "sksilar008@gmail.com" -PlainPassword "yourpassword"
```

Passwords are stored encrypted (Windows DPAPI — only YOUR Windows user can decrypt them).

## Tracker statuses explained

| Status | Meaning |
|---|---|
| New | Found by HuntBot, waiting to apply |
| Applied | Application confirmed submitted |
| Attempted - check | Submit clicked but no confirmation seen — check manually |
| Needs account | Site requires login — add it (see above) |
| Needs direct link | The URL is a search page, not a job post — HuntBot will try to find better links |
| Needs manual | Bot couldn't handle this one — apply by hand |
| Ready (dry run) | Form was filled in a dry run, nothing submitted |

## Honest limitations

- Auto-submit works best on simple sites. Sites with CAPTCHA (Indeed, LinkedIn)
  will often block the bot — those become "Needs manual".
- Every attempt is screenshotted in `logs\shots\` — you can always verify what happened.
- Applying too fast/too often on one site can get your account flagged. `maxJobsPerRun`
  in `config.json` (default 10) keeps this sane.
