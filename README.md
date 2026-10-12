# ANKIROS Sync

Standalone ANKIROS exhibitor synchronizer. It uses the public API behind
`https://www.ankiros.com/tr/exhibitor-list`, stores normalized records in
SQLite, and exports Excel workbooks.

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

Useful commands:

```powershell
.\.venv\Scripts\python.exe ankiros_sync.py sync --filter direct
.\.venv\Scripts\python.exe ankiros_sync.py sync --filter co-exhibitors
.\.venv\Scripts\python.exe ankiros_sync.py export --filter all
```

`--filter direct` exports only companies returned by the participant list.
`--filter co-exhibitors` exports only companies nested under a direct
exhibitor's `subcoCompany` detail. `--filter all` includes both and labels
each row with its relationship and parent company.

Reports are written to `exports/ANKIROS_2026.xlsx`; the database is
`data/ankiros.db`.
