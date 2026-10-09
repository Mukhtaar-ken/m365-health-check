# M365 Health Check

![Tests](https://github.com/Mukhtaar-ken/m365-health-check/actions/workflows/tests.yml/badge.svg)

A PowerShell module that checks Exchange Online mailboxes for problems before users notice them,
and writes the results as a CSV and HTML report, worst problems first.

## Why I built it

As an IT admin I've had to diagnose mailboxes by hand: a Recoverable Items folder that had quietly
filled up, or a mailbox close to its quota with no archive for old mail to go to. Both are easy to
miss until something breaks, and both are easy to spot if you look. So I wanted a repeatable check
I could run across every mailbox, with tests, instead of a one-off script.

## What it checks

| Check | Flags | Default thresholds |
|---|---|---|
| `Test-RecoverableItemsQuota` | Recoverable Items folder close to its quota | Warning 80%, Critical 95% |
| `Test-MailboxArchive` | No archive **and** the mailbox is filling up | Warning 80%, Critical 95% |

Each result is **OK**, **Warning**, **Critical** or **Unknown**. Unknown means the data was
missing, so the check refuses to guess.

## Install and run

Needs Windows PowerShell 5.1 or PowerShell 7.

**Try it with fake data** (no Microsoft 365 needed):

```powershell
git clone https://github.com/Mukhtaar-ken/m365-health-check.git
cd m365-health-check
Import-Module .\src\M365HealthCheck\M365HealthCheck.psd1
Get-Content .\tests\fixtures\mailboxes.json -Raw | ConvertFrom-Json | Invoke-M365HealthCheck
```

**Run it against a tenant** (read-only; needs an Exchange role that can read mailboxes,
such as View-Only Organization Management or Global Reader):

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
Import-Module .\src\M365HealthCheck\M365HealthCheck.psd1
Invoke-M365HealthCheck -OutputPath .\output
Disconnect-ExchangeOnline -Confirm:$false
```

The report goes to `output\`, which git ignores, so real mailbox data never gets committed.

## How it's built

```
Get-M365HealthData      -> the only part that talks to Exchange
        |
Test-* checks           -> only evaluate objects; never connect to anything
        |
Invoke-M365HealthCheck  -> runs every Test-* check, sorts worst first, writes the report
```

**Collecting and checking are kept apart on purpose.** The checks take plain objects, so they can be
tested with fake data in milliseconds, with no tenant and no credentials. `Get-M365HealthData` turns
Exchange's output (sizes like `"49.5 GB (53,150,220,288 bytes)"`) into those objects.

New checks are picked up automatically: `Invoke-M365HealthCheck` runs every exported function whose
name starts with `Test-`.

## How it's tested

- **Pester** tests for every function. The Exchange commands are replaced with mocks, so the tests
  never connect to Microsoft 365.
- **A manifest test** fails if a new function is added but not exported, a mistake that would
  otherwise go unnoticed.
- **GitHub Actions** runs Pester and **PSScriptAnalyzer** (code style) on every pull request and on
  every merge to `main`, on a clean Windows machine.
- I made each important test fail on purpose before trusting it. A test that has never failed
  might not be checking anything.

## What broke

**Every test passed, and the first real run was still wrong.** Against a real tenant, every mailbox
came back as 0% used and OK. Two bugs:

1. `Get-EXOMailbox` doesn't return `ExchangeGuid` unless you ask for it, so every size lookup used a
   blank ID. My mocks included `ExchangeGuid`, because that's what I *assumed* the real command
   returned.
2. The checks treated a missing size as 0%, so "no data" was reported as **OK**. That hid bug 1.

The fix: look mailboxes up by `ExternalDirectoryObjectId` (always returned), report missing data as
**Unknown**, warn if a mailbox can't be read, and change the mocks to match what the real tenant
actually returned. The new tests fail if the old bug comes back.

**Lesson: mocks only test what you think the API returns. Run against real data early.**

## What I'd change

- **Speed.** It makes three calls per mailbox, which is fine for a small tenant but slow for
  thousands. Next I'd fetch statistics in bulk or run mailboxes in parallel (PowerShell 7).
- **Shared logic.** Both checks repeat the same percentage and threshold code. With a third check,
  I'd move that into a private helper.
- **More checks.** For example, litigation hold left on a disabled account.
- **Branch protection** on `main`, so a pull request can't be merged until the tests pass.
