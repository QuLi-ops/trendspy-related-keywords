$repo = "D:\github_respository\trendspy-related-keywords"
$stamp = Join-Path $repo ".last_collect_date.txt"
$today = (Get-Date).ToString("yyyy-MM-dd")

if (Test-Path $stamp) {
    $last = (Get-Content $stamp -Raw).Trim()
    if ($last -eq $today) {
        exit 0
    }
}

Set-Location $repo
& "$repo\venv\Scripts\python.exe" trends_monitor.py --test >> "$repo\trends_collect.log" 2>&1

if ($LASTEXITCODE -eq 0) {
    Set-Content -Path $stamp -Value $today -NoNewline
}