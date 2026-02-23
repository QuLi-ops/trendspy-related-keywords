1. 新建 D:\github_respository\trendspy-related-keywords\run_collect_once_per_day.ps1，内容：

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

2. 创建“开机触发”任务（每天最多跑一次）：

schtasks /Create /TN "TrendsCollectAtStartup" /SC ONSTART /DELAY 0001:00 /TR "powershell -NoProfile -ExecutionPolicy Bypass -File D:\github_respository\trendspy-related-keywords\run_collect_once_per_day.ps1" /F

3. 任务不用时可删除：

schtasks /Delete /TN "TrendsCollectAtStartup" /F

补充：

1. 临时停用（不删除）：

schtasks /Change /TN "TrendsCollectAtStartup" /DISABLE

2. 恢复启用：

schtasks /Change /TN "TrendsCollectAtStartup" /ENABLE