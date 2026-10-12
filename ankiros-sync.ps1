param(
    [ValidateSet("test", "smoke", "sync", "export", "check", "help")]
    [string]$Command = "help",
    [ValidateSet("all", "direct", "co-exhibitors")]
    [string]$Filter = "all"
)

$Python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
    $Python = "python"
}

switch ($Command) {
    "test" {
        & $Python -m pytest -q
        exit $LASTEXITCODE
    }
    "smoke" {
        & $Python -c "from ankiros_sync import fetch_exhibitors; records=fetch_exhibitors(); print('ANKIROS exhibitors:', len(records)); print('Direct exhibitors:', sum(r['relationship'] == 'direct' for r in records)); print('Co-exhibitors:', sum(r['relationship'] == 'co-exhibitor' for r in records))"
        exit $LASTEXITCODE
    }
    "sync" {
        & $Python ankiros_sync.py sync --filter $Filter
        exit $LASTEXITCODE
    }
    "export" {
        & $Python ankiros_sync.py export --filter $Filter
        exit $LASTEXITCODE
    }
    "check" {
        & $Python -m compileall -q .
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
        & $Python -m pytest -q
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
        git diff --check
        exit $LASTEXITCODE
    }
    "help" {
        Write-Host "Usage: .\ankiros-sync.ps1 <command> [-Filter <all|direct|co-exhibitors>]"
        Write-Host "  test    Run the complete pytest suite"
        Write-Host "  smoke   Query the live ANKIROS API and print record counts"
        Write-Host "  sync    Fetch, persist, and export ANKIROS exhibitors"
        Write-Host "  export  Export ANKIROS exhibitors without database synchronization"
        Write-Host "  check   Compile Python, run tests, and check the Git diff"
        Write-Host "  help    Show this help text"
    }
}
