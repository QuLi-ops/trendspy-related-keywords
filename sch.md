## 设置计划
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

## 设置防火墙放行
1. 打开“Windows Defender 防火墙”。
2. 左侧点“高级设置”。
3. 点“出站规则” -> 右侧“新建规则”。
4. 规则类型选“程序”。
5. 程序路径填：D:\github_respository\trendspy-related-keywords\venv\Scripts\python.exe
6. 选“允许连接”。
7. 配置文件勾选“域/专用/公用”（先全勾，后面可收紧）。
8. 名称填：Allow Trendspy Python SMTP，完成。

可再加一个端口规则（可选）：

- 出站规则 -> 新建规则 -> 端口 -> TCP -> 特定本地端口填 465,587 -> 允许连接。

查看是否生效：

Get-NetFirewallRule -DisplayName "*Trendspy*","*SMTP 465*","*SMTP 587*" | Format-Table DisplayName,Enabled,Direction,Action

删规则（不用时）：

Remove-NetFirewallRule -DisplayName "Allow Trendspy Python Outbound","Allow SMTP 465 Outbound","Allow SMTP 587 Outbound"