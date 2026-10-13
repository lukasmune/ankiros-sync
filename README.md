# ankiros-sync

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

Standalone ANKIROS exhibitor synchronizer. It uses the public API behind
`https://www.ankiros.com/tr/exhibitor-list`, stores normalized records in
SQLite, and exports Excel workbooks with summary, company master, and sales
lead sheets.

The API currently uses the 2026 exhibition identifier and exposes:

- Direct exhibitors from `getAllParticipantList`.
- Co-exhibitors from each company's `subcoCompany` detail field.

## Build from scratch (PowerShell)

```powershell
New-Item -ItemType Directory -Force ankiros-sync
Set-Location ankiros-sync
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe ankiros_sync.py sync --filter all
```

## PowerShell command aliases

The project includes [ankiros-sync.ps1](./ankiros-sync.ps1), matching the
command style used by `md-sync.ps1`:

```powershell
.\ankiros-sync.ps1 test
.\ankiros-sync.ps1 smoke
.\ankiros-sync.ps1 sync -Filter all
.\ankiros-sync.ps1 sync -Filter direct
.\ankiros-sync.ps1 sync -Filter co-exhibitors
.\ankiros-sync.ps1 export -Filter all
.\ankiros-sync.ps1 check
```

`test` runs the complete pytest suite. `smoke` intentionally calls the live
ANKIROS API. `check` compiles Python, runs pytest, and checks the Git diff.

Useful commands:

```powershell
.\.venv\Scripts\python.exe ankiros_sync.py sync --filter direct
.\.venv\Scripts\python.exe ankiros_sync.py sync --filter co-exhibitors
.\.venv\Scripts\python.exe ankiros_sync.py export --filter all
```

`--filter direct` exports only companies returned by the participant list.
`--filter co-exhibitors` exports only companies nested under a direct
exhibitor's `subcoCompany` detail. `--filter all` includes both and labels
each row with its relationship and parent company. The `Company Master` and
`Sales Leads` sheets include each exhibitor's country specifier. `Sales Leads`
also includes the API's address, stand, phone, email, and English product
classifications in addition to the relationship and slug.

Reports are written to `exports/ANKIROS_2026.xlsx`; the database is
`data/ankiros.db`.
