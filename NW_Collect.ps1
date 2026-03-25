# ============================================================
#  ライフリズムナビ NW情報収集ツール
#  バッチファイルから呼び出されるPowerShellスクリプト
# ============================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Blue
Write-Host "  ║  ライフリズムナビ NW情報収集ツール       ║" -ForegroundColor Blue
Write-Host "  ║  ネットワーク情報を自動収集しています... ║" -ForegroundColor Blue
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Blue
Write-Host ""

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = Get-Location }
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputFile = Join-Path $scriptDir "NW_Report_$timestamp.html"

# =====================
#  1. Wi-Fi情報の取得
# =====================
Write-Host "  [1/5] Wi-Fi接続情報を取得中..." -ForegroundColor Cyan

$wifi = @{
    SSID       = "（未接続）"
    Auth       = "−"
    Cipher     = "−"
    Signal     = "−"
    RadioType  = "−"
    Channel    = "−"
    BSSID      = "−"
}

try {
    $wlanOutput = netsh wlan show interfaces 2>$null
    if ($wlanOutput) {
        foreach ($line in $wlanOutput) {
            if ($line -match '^\s+SSID\s+:\s+(.+)$') { $wifi.SSID = $Matches[1].Trim() }
            if ($line -match '^\s+BSSID\s+:\s+(.+)$') { $wifi.BSSID = $Matches[1].Trim() }
            if ($line -match '認証\s+:\s+(.+)$' -or $line -match 'Authentication\s+:\s+(.+)$') { $wifi.Auth = $Matches[1].Trim() }
            if ($line -match '暗号\s+:\s+(.+)$' -or $line -match 'Cipher\s+:\s+(.+)$') { $wifi.Cipher = $Matches[1].Trim() }
            if ($line -match 'シグナル\s+:\s+(.+)$' -or $line -match 'Signal\s+:\s+(.+)$') {
                $rawSignal = $Matches[1].Trim()
                # パーセンテージからRSSI(dBm)に変換: dBm = (% / 2) - 100
                if ($rawSignal -match '(\d+)%') {
                    $pct = [int]$Matches[1]
                    $rssi = [math]::Round($pct / 2 - 100)
                    # 品質ラベル（metageek基準 / LRN導入NW資料P27準拠）
                    # -30dBm: Amazing / -67dBm: Very Good / -70dBm: Okay(GW最低ライン) / -80dBm: Not Good / -90dBm: Unusable
                    if     ($rssi -ge -30) { $quality = "Amazing（素晴らしい）";   $qualityEn = "amazing" }
                    elseif ($rssi -ge -67) { $quality = "Very Good（良好）";       $qualityEn = "verygood" }
                    elseif ($rssi -ge -70) { $quality = "Okay（使用可能）";        $qualityEn = "okay" }
                    elseif ($rssi -ge -80) { $quality = "Not Good（よくない）";    $qualityEn = "notgood" }
                    else                   { $quality = "Unusable（使用不可）";    $qualityEn = "unusable" }
                    $wifi.Signal = "$rssi dBm"
                    $wifi.SignalRaw = $rssi
                    $wifi.SignalQuality = $quality
                    $wifi.SignalQualityEn = $qualityEn
                } else {
                    $wifi.Signal = $rawSignal
                }
            }
            if ($line -match '無線の種類\s+:\s+(.+)$' -or $line -match 'Radio type\s+:\s+(.+)$') { $wifi.RadioType = $Matches[1].Trim() }
            if ($line -match 'チャネル\s+:\s+(.+)$' -or $line -match 'Channel\s+:\s+(.+)$') {
                if ($line -notmatch '無線|Radio') { $wifi.Channel = $Matches[1].Trim() }
            }
        }
    }
} catch { }

# =====================
#  2. IP情報の取得
# =====================
Write-Host "  [2/5] IPアドレス情報を取得中..." -ForegroundColor Cyan

$ip = @{
    Address   = "−"
    Mask      = "−"
    Gateway   = "−"
    DNS1      = "−"
    DNS2      = "−"
    DHCP      = "−"
    MAC       = "−"
    Hostname  = $env:COMPUTERNAME
}

try {
    # アクティブなネットワークアダプタからIP情報を取得
    $adapters = Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object {
        $_.IPv4DefaultGateway -ne $null
    } | Select-Object -First 1

    if ($adapters) {
        $ip.Address = ($adapters.IPv4Address | Select-Object -First 1).IPAddress
        $ip.Gateway = ($adapters.IPv4DefaultGateway | Select-Object -First 1).NextHop

        $ifIndex = $adapters.InterfaceIndex
        $adapterConfig = Get-NetIPAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($adapterConfig) {
            $prefixLen = $adapterConfig.PrefixLength
            # プレフィックス長からサブネットマスクを計算
            $maskBits = ('1' * $prefixLen).PadRight(32, '0')
            $ip.Mask = "{0}.{1}.{2}.{3}" -f `
                [Convert]::ToInt32($maskBits.Substring(0,8), 2),
                [Convert]::ToInt32($maskBits.Substring(8,8), 2),
                [Convert]::ToInt32($maskBits.Substring(16,8), 2),
                [Convert]::ToInt32($maskBits.Substring(24,8), 2)
        }

        $dnsServers = ($adapters.DNSServer | Where-Object { $_.AddressFamily -eq 2 }).ServerAddresses
        if ($dnsServers -and $dnsServers.Count -ge 1) { $ip.DNS1 = $dnsServers[0] }
        if ($dnsServers -and $dnsServers.Count -ge 2) { $ip.DNS2 = $dnsServers[1] }

        $netAdapter = Get-NetAdapter -InterfaceIndex $ifIndex -ErrorAction SilentlyContinue
        if ($netAdapter) { $ip.MAC = $netAdapter.MacAddress }

        $dhcpEnabled = (Get-NetIPInterface -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).Dhcp
        $ip.DHCP = if ($dhcpEnabled -eq 'Enabled') { "有効" } else { "無効（固定IP）" }
    }
} catch {
    # フォールバック: ipconfig /all からパース
    try {
        $ipconfigOutput = ipconfig /all 2>$null
        foreach ($line in $ipconfigOutput) {
            if ($line -match 'IPv4.*:\s+([\d\.]+)') { if ($ip.Address -eq "−") { $ip.Address = $Matches[1] } }
            if ($line -match '(サブネット|Subnet).*:\s+([\d\.]+)') { if ($ip.Mask -eq "−") { $ip.Mask = $Matches[2] } }
            if ($line -match '(デフォルト|Default).*:\s+([\d\.]+)') { if ($ip.Gateway -eq "−") { $ip.Gateway = $Matches[2] } }
            if ($line -match 'DNS.*:\s+([\d\.]+)') {
                if ($ip.DNS1 -eq "−") { $ip.DNS1 = $Matches[1] }
                elseif ($ip.DNS2 -eq "−") { $ip.DNS2 = $Matches[1] }
            }
            if ($line -match '(物理|Physical).*:\s+([\w-]+)') { if ($ip.MAC -eq "−") { $ip.MAC = $Matches[2] } }
        }
    } catch { }
}

# =====================
#  3. インターネット接続確認
# =====================
Write-Host "  [3/5] インターネット接続を確認中..." -ForegroundColor Cyan

$netStatus = "NG"
try {
    $pingResult = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($pingResult) { $netStatus = "OK" }
} catch { }

# =====================
#  4. IPアドレススキャン
# =====================
Write-Host "  [4/5] ネットワーク内のIPアドレスをスキャン中..." -ForegroundColor Cyan

$scanResults = @()
$scanSummary = @{ Total = 0; Used = 0; Available = 0; NetworkAddr = "−"; BroadcastAddr = "−"; PrefixLength = 0; ScanTime = "−" }

try {
    if ($ip.Address -ne "−" -and $ip.Mask -ne "−") {
        # IPアドレスとサブネットマスクからネットワーク範囲を計算
        $ipBytes = [System.Net.IPAddress]::Parse($ip.Address).GetAddressBytes()
        $maskBytes = [System.Net.IPAddress]::Parse($ip.Mask).GetAddressBytes()

        # プレフィックス長を計算
        $prefixLength = 0
        foreach ($b in $maskBytes) {
            $bits = [Convert]::ToString($b, 2)
            $prefixLength += ($bits.ToCharArray() | Where-Object { $_ -eq '1' }).Count
        }
        $scanSummary.PrefixLength = $prefixLength

        # ネットワークアドレスとブロードキャストアドレスを計算
        $networkBytes = @(0,0,0,0)
        $broadcastBytes = @(0,0,0,0)
        for ($i = 0; $i -lt 4; $i++) {
            $networkBytes[$i] = $ipBytes[$i] -band $maskBytes[$i]
            $broadcastBytes[$i] = $ipBytes[$i] -bor (255 - $maskBytes[$i])
        }
        $networkAddr = "$($networkBytes[0]).$($networkBytes[1]).$($networkBytes[2]).$($networkBytes[3])"
        $broadcastAddr = "$($broadcastBytes[0]).$($broadcastBytes[1]).$($broadcastBytes[2]).$($broadcastBytes[3])"
        $scanSummary.NetworkAddr = $networkAddr
        $scanSummary.BroadcastAddr = $broadcastAddr

        # ホストIPの範囲を計算（ネットワークアドレス+1 ～ ブロードキャスト-1）
        $netInt = [uint32]($networkBytes[0]) * 16777216 + [uint32]($networkBytes[1]) * 65536 + [uint32]($networkBytes[2]) * 256 + [uint32]($networkBytes[3])
        $bcastInt = [uint32]($broadcastBytes[0]) * 16777216 + [uint32]($broadcastBytes[1]) * 65536 + [uint32]($broadcastBytes[2]) * 256 + [uint32]($broadcastBytes[3])
        $startHost = $netInt + 1
        $endHost = $bcastInt - 1
        $totalHosts = $endHost - $startHost + 1
        $scanSummary.Total = $totalHosts

        # /16以上の大きなサブネットの場合に警告
        if ($prefixLength -le 16) {
            Write-Host ""
            Write-Host "  ⚠ /$prefixLength のネットワークです（$totalHosts ホスト）" -ForegroundColor Yellow
            Write-Host "  スキャンに30分以上かかる場合があります。しばらくお待ちください..." -ForegroundColor Yellow
            Write-Host ""
        } elseif ($prefixLength -le 20) {
            Write-Host "  ⚠ /$prefixLength のネットワーク（$totalHosts ホスト）- 数分かかります..." -ForegroundColor Yellow
        }

        $scanStart = Get-Date

        # --- 並列Pingスキャン ---
        # PowerShell 5.1互換: Runspace Poolで並列実行
        $maxThreads = 100
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
        $runspacePool.Open()

        $scriptBlock = {
            param($targetIP)
            $result = @{ IP = $targetIP; Status = "Available" ; Hostname = "" }
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $reply = $ping.Send($targetIP, 1000)
                if ($reply.Status -eq 'Success') {
                    $result.Status = "Used"
                    try {
                        $dns = [System.Net.Dns]::GetHostEntry($targetIP)
                        $result.Hostname = $dns.HostName
                    } catch { }
                }
                $ping.Dispose()
            } catch { }
            return $result
        }

        $jobs = @()
        $completed = 0
        $progressInterval = [Math]::Max(1, [Math]::Floor($totalHosts / 50))

        Write-Host "  スキャン開始: $networkAddr/$prefixLength ($totalHosts ホスト)" -ForegroundColor Gray

        for ($hostInt = $startHost; $hostInt -le $endHost; $hostInt++) {
            $o1 = [Math]::Floor($hostInt / 16777216) % 256
            $o2 = [Math]::Floor($hostInt / 65536) % 256
            $o3 = [Math]::Floor($hostInt / 256) % 256
            $o4 = $hostInt % 256
            $targetIP = "$o1.$o2.$o3.$o4"

            $ps = [powershell]::Create().AddScript($scriptBlock).AddArgument($targetIP)
            $ps.RunspacePool = $runspacePool
            $jobs += @{
                Pipe = $ps
                Result = $ps.BeginInvoke()
                IP = $targetIP
            }

            # プログレス表示（一定間隔で）
            if ($jobs.Count % $progressInterval -eq 0 -or $hostInt -eq $endHost) {
                $pct = [Math]::Floor(($jobs.Count / $totalHosts) * 100)
                $bar = ('█' * [Math]::Floor($pct / 5)).PadRight(20, '░')
                Write-Host "`r  [$bar] $pct% ($($jobs.Count)/$totalHosts) 送信中..." -NoNewline -ForegroundColor Cyan
            }
        }

        Write-Host ""
        Write-Host "  応答を待機中..." -ForegroundColor Gray

        # 結果収集
        foreach ($job in $jobs) {
            try {
                $result = $job.Pipe.EndInvoke($job.Result)
                if ($result) {
                    $scanResults += $result
                }
            } catch { }
            $job.Pipe.Dispose()
            $completed++

            if ($completed % $progressInterval -eq 0 -or $completed -eq $jobs.Count) {
                $pct = [Math]::Floor(($completed / $jobs.Count) * 100)
                $bar = ('█' * [Math]::Floor($pct / 5)).PadRight(20, '░')
                Write-Host "`r  [$bar] $pct% ($completed/$($jobs.Count)) 収集中..." -NoNewline -ForegroundColor Cyan
            }
        }

        $runspacePool.Close()
        $runspacePool.Dispose()

        Write-Host ""

        # ARPテーブルも参照して補完（ping応答しないがARP応答するデバイス）
        # dynamic = DHCP割当の可能性大 / static = 固定IP設定済みの可能性大
        Write-Host "  ARPテーブルを確認中..." -ForegroundColor Gray
        try {
            $arpOutput = arp -a 2>$null
            foreach ($line in $arpOutput) {
                if ($line -match '^\s+([\d\.]+)\s+([\w-]+)\s+(\w+)') {
                    $arpIP = $Matches[1]
                    $arpMAC = $Matches[2]
                    $arpType = $Matches[3]
                    if ($arpMAC -eq 'ff-ff-ff-ff-ff-ff') { continue }
                    $existing = $scanResults | Where-Object { $_.IP -eq $arpIP }
                    if ($existing) {
                        # 使用中補完（ping非応答でもARP応答があれば使用中）
                        if ($existing.Status -eq 'Available' -and $arpType -ne 'static') {
                            $existing.Status = 'Used'
                        }
                        # ARPタイプを記録（後段のDHCP推定に利用）
                        $existing | Add-Member -NotePropertyName 'ArpType' -NotePropertyValue $arpType -Force
                    }
                }
            }
        } catch { }

        $scanEnd = Get-Date
        $elapsed = $scanEnd - $scanStart
        $scanSummary.ScanTime = "{0:mm}分{0:ss}秒" -f $elapsed

        # 集計
        $scanSummary.Used = ($scanResults | Where-Object { $_.Status -eq 'Used' }).Count
        $scanSummary.Available = ($scanResults | Where-Object { $_.Status -eq 'Available' }).Count

        Write-Host ""
        Write-Host "  スキャン完了！ （所要時間: $($scanSummary.ScanTime)）" -ForegroundColor Green
        Write-Host "  使用中: $($scanSummary.Used) / 空き: $($scanSummary.Available) / 合計: $totalHosts" -ForegroundColor White
    } else {
        Write-Host "  ⚠ IPアドレスまたはサブネットマスクが取得できなかったため、スキャンをスキップしました。" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  ⚠ IPスキャン中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
}

# =====================
#  DHCP推定分析
# =====================
# IP数値変換ヘルパー
function IntToIP([uint32]$n) {
    "$([Math]::Floor($n/16777216)%256).$([Math]::Floor($n/65536)%256).$([Math]::Floor($n/256)%256).$($n%256)"
}
function IPToInt([string]$ip) {
    $p = $ip -split '\.'
    [uint32]$p[0]*16777216 + [uint32]$p[1]*65536 + [uint32]$p[2]*256 + [uint32]$p[3]
}

$dhcpAnalysisHtml = ""
$ipMapHtml       = ""

if ($scanResults.Count -gt 0) {
    # ── ① IPアドレス分布マップ（16列グリッド）────────────────────
    $usedIPSet = @{}
    foreach ($r in ($scanResults | Where-Object { $_.Status -eq 'Used' })) {
        $lastOctet = [int]($r.IP -split '\.')[-1]
        $usedIPSet[$lastOctet] = $r
    }
    $prefix3 = ($scanSummary.NetworkAddr -replace '\.\d+$','')

    $ipMapHtml = "<div style='font-family:monospace;font-size:13px;line-height:1.8;'>`n"
    # 行ラベル付き16列グリッド
    $ipMapHtml += "<div style='display:inline-grid;grid-template-columns:repeat(17,auto);gap:2px;align-items:center;'>`n"
    for ($rowStart = 1; $rowStart -le 240; $rowStart += 16) {
        # 行ラベル
        $ipMapHtml += "<span style='color:#94a3b8;font-size:11px;padding-right:4px;'>.$rowStart</span>`n"
        for ($i = 0; $i -lt 16; $i++) {
            $n = $rowStart + $i
            if ($n -gt 254) { $ipMapHtml += "<span></span>`n"; continue }
            $tip = "$prefix3.$n"
            $hn = if ($usedIPSet.ContainsKey($n) -and $usedIPSet[$n].Hostname) { " ($($usedIPSet[$n].Hostname))" } else { "" }
            if ($usedIPSet.ContainsKey($n)) {
                $ipMapHtml += "<span title='$tip$hn' style='display:inline-block;width:18px;height:18px;background:#fca5a5;border:1px solid #f87171;border-radius:3px;cursor:default;'></span>`n"
            } else {
                $ipMapHtml += "<span title='$tip' style='display:inline-block;width:18px;height:18px;background:#d1fae5;border:1px solid #6ee7b7;border-radius:3px;cursor:default;'></span>`n"
            }
        }
    }
    $ipMapHtml += "</div>`n"
    $ipMapHtml += "<div style='margin-top:8px;font-size:11px;color:#64748b;'>
      <span style='display:inline-block;width:12px;height:12px;background:#fca5a5;border:1px solid #f87171;border-radius:2px;vertical-align:middle;'></span> 使用中 &nbsp;
      <span style='display:inline-block;width:12px;height:12px;background:#d1fae5;border:1px solid #6ee7b7;border-radius:2px;vertical-align:middle;'></span> 空き &nbsp;
      ※ セルにカーソルを合わせるとIPアドレスが表示されます
    </div></div>`n"

    # ── ② 周囲から離れたIP（孤立IP検出・要目視確認）─────────────
    $usedSorted = $usedIPSet.Keys | Sort-Object
    $isolatedHtml = ""
    $isolatedText = @()
    if ($usedSorted.Count -gt 0) {
        for ($i = 0; $i -lt $usedSorted.Count; $i++) {
            $n = $usedSorted[$i]
            $leftGap  = if ($i -gt 0) { $n - $usedSorted[$i-1] } else { 999 }
            $rightGap = if ($i -lt $usedSorted.Count - 1) { $usedSorted[$i+1] - $n } else { 999 }
            if ($leftGap -ge 5 -and $rightGap -ge 5) {
                $hn = if ($usedIPSet[$n].Hostname) { " ($($usedIPSet[$n].Hostname))" } else { "" }
                $isolatedHtml += "<span style='display:inline-block;background:#fff7ed;color:#c2410c;border:1px solid #fed7aa;border-radius:4px;padding:2px 10px;margin:2px;font-size:13px;font-family:monospace;'>$prefix3.$n$hn</span>`n"
                $isolatedText += "$prefix3.$n$hn"
            }
        }
    }
    if (-not $isolatedHtml) { $isolatedHtml = "<span style='color:#6b7280;font-size:13px;'>検出なし（全使用中IPが近接しています）</span>" }


    $dhcpAnalysisHtml = @"
<div style='margin-bottom:16px;'>
  <div style='font-size:12px;font-weight:700;color:#1e293b;margin-bottom:8px;'>📍 IPアドレス分布マップ（赤=使用中 / 緑=空き）</div>
  $ipMapHtml
</div>
<div style='margin-bottom:16px;padding:12px;background:#fff7ed;border:1px solid #fed7aa;border-radius:6px;'>
  <div style='font-size:12px;font-weight:700;color:#c2410c;margin-bottom:6px;'>🔍 周囲のIPと大きく離れているアドレス（前後5以上の空白）</div>
  <div style='line-height:2.2;'>$isolatedHtml</div>
  <div style='font-size:11px;color:#92400e;margin-top:6px;'>固定IPが設定されている可能性がありますが、断定はできません。ルーター管理画面での確認を推奨します。</div>
</div>

"@
}

# スキャン結果をHTML用に変数へ格納
$scanUsedHtml = ""
$scanAvailHtml = ""
$scanUsedText = @()
$scanAvailText = @()

foreach ($r in ($scanResults | Sort-Object { $parts = $_.IP -split '\.'; [uint32]$parts[0]*16777216 + [uint32]$parts[1]*65536 + [uint32]$parts[2]*256 + [uint32]$parts[3] })) {
    if ($r.Status -eq 'Used') {
        $hostname = if ($r.Hostname) { " ($($r.Hostname))" } else { "" }
        $scanUsedHtml += "<span style='display:inline-block;background:#fee2e2;color:#dc2626;border:1px solid #fca5a5;border-radius:4px;padding:2px 8px;margin:2px;font-size:13px;font-family:monospace;'>$($r.IP)$hostname</span>`n"
        $scanUsedText += "$($r.IP)$hostname"
    } else {
        $scanAvailHtml += "<span style='display:inline-block;background:#d1fae5;color:#059669;border:1px solid #6ee7b7;border-radius:4px;padding:2px 8px;margin:2px;font-size:13px;font-family:monospace;'>$($r.IP)</span>`n"
        $scanAvailText += $r.IP
    }
}

# =====================
#  5. HTMLレポート生成
# =====================
Write-Host "  [5/5] レポートを生成中..." -ForegroundColor Cyan

$badgeClass = if ($netStatus -eq "OK") { "badge-ok" } else { "badge-ng" }
$badgeText = if ($netStatus -eq "OK") { "接続OK" } else { "接続NG" }
$dateStr = Get-Date -Format "yyyy/MM/dd HH:mm"

$html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NW情報収集レポート</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI','Meiryo','Yu Gothic',sans-serif; background:#f0f4f8; color:#1a202c; padding:20px; }
.container { max-width:820px; margin:0 auto; }
.header { background:linear-gradient(135deg,#2563eb,#1d4ed8); color:#fff; padding:24px 32px; border-radius:12px 12px 0 0; }
.header h1 { font-size:22px; font-weight:700; }
.header p { font-size:13px; opacity:0.85; margin-top:4px; }
.badge { display:inline-block; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
.badge-ok { background:#10b981; color:#fff; }
.badge-ng { background:#ef4444; color:#fff; }
.card { background:#fff; border-radius:0 0 12px 12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; overflow:hidden; }
.card + .card, .card + .manual-section, .manual-section + .card { border-radius:12px; }
.section { padding:20px 32px; border-bottom:1px solid #e5e7eb; }
.section:last-child { border-bottom:none; }
.section-title { font-size:15px; font-weight:700; color:#2563eb; margin-bottom:14px; display:flex; align-items:center; gap:8px; }
.section-title .icon { font-size:18px; }
table { width:100%; border-collapse:collapse; }
td { padding:10px 12px; font-size:14px; border-bottom:1px solid #f1f5f9; }
td:first-child { color:#64748b; font-weight:500; width:220px; }
td:last-child { color:#1e293b; font-weight:600; }
.auto-tag { font-size:10px; color:#10b981; font-weight:600; margin-left:6px; }
.manual-section { background:#fffbeb; border:2px solid #fbbf24; border-radius:12px; margin-bottom:20px; }
.manual-section .section-title { color:#d97706; }
.manual-section .note { font-size:12px; color:#92400e; background:#fef3c7; padding:8px 12px; border-radius:6px; margin-bottom:14px; }
input[type="text"], select, textarea { width:100%; padding:10px 14px; border:2px solid #e2e8f0; border-radius:8px; font-size:14px; font-family:inherit; transition:border-color 0.2s; }
input:focus, select:focus, textarea:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 3px rgba(37,99,235,0.1); }
.form-row { display:flex; gap:16px; margin-bottom:12px; }
.form-group { flex:1; }
.form-group label { display:block; font-size:12px; font-weight:600; color:#64748b; margin-bottom:4px; }
.btn-group { display:flex; gap:12px; justify-content:center; padding:24px; flex-wrap:wrap; }
.btn { padding:12px 32px; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; border:none; transition:all 0.2s; }
.btn-primary { background:#2563eb; color:#fff; }
.btn-primary:hover { background:#1d4ed8; transform:translateY(-1px); }
.btn-secondary { background:#f1f5f9; color:#475569; }
.btn-secondary:hover { background:#e2e8f0; }
.btn-success { background:#10b981; color:#fff; }
.btn-success:hover { background:#059669; }
.footer { text-align:center; padding:16px; color:#94a3b8; font-size:12px; }
.tip { font-size:11px; color:#94a3b8; margin-top:4px; }
.scan-scroll::-webkit-scrollbar { width:6px; } .scan-scroll::-webkit-scrollbar-thumb { background:#cbd5e1; border-radius:3px; }
@media print { body{background:#fff;} .btn-group{display:none;} .manual-section{border-color:#ccc; page-break-before:always;} .facility-card{page-break-before:always;} input,select,textarea{border-color:#ccc;} .footer{page-break-before:avoid;} .scan-scroll{max-height:none !important;overflow:visible !important;} }
</style>
</head>
<body>
<div class="container">

<!-- ヘッダー -->
<div class="header">
<h1>&#128225; NW情報収集レポート</h1>
<p>ライフリズムナビ 導入前ネットワーク調査</p>
<p style="margin-top:8px">取得日時: $dateStr ／ PC名: $($ip.Hostname) ／ インターネット: <span class="badge $badgeClass">$badgeText</span></p>
</div>

<!-- 自動取得: Wi-Fi -->
<div class="card">
<div class="section">
<div class="section-title"><span class="icon">&#128246;</span> Wi-Fi接続情報 <span class="auto-tag">&#10003; 自動取得</span></div>
<table>
<tr><td>SSID（ネットワーク名）</td><td>$($wifi.SSID)</td></tr>
<tr><td>認証方式</td><td>$($wifi.Auth)</td></tr>
<tr><td>暗号化方式</td><td>$($wifi.Cipher)</td></tr>
<tr><td>電波強度（RSSI）</td><td>$(if ($wifi.SignalRaw) {
    # metageek基準 5段階の色分け
    $rssiColor = switch ($wifi.SignalQualityEn) {
        'amazing'  { '#10b981' }  # 緑
        'verygood' { '#3b82f6' }  # 青
        'okay'     { '#f59e0b' }  # 黄
        'notgood'  { '#ef4444' }  # 赤
        'unusable' { '#7f1d1d' }  # 暗赤
        default    { '#64748b' }
    }
    $barWidth = [math]::Max(5, [math]::Min(100, (($wifi.SignalRaw + 100) * 1.5)))
    # GW最低ライン(-70dBm)の注記
    $gwNote = if ($wifi.SignalRaw -lt -70) { "<div style='margin-top:6px;padding:6px 10px;background:#fef2f2;border:1px solid #fecaca;border-radius:6px;font-size:11px;color:#991b1b;'>⚠ SleepSensorゲートウェイの最低要件 -70dBm を下回っています。AP増設や設置位置の見直しを検討してください。</div>" } elseif ($wifi.SignalRaw -ge -70 -and $wifi.SignalRaw -lt -67) { "<div style='margin-top:6px;padding:6px 10px;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;font-size:11px;color:#92400e;'>△ GW最低ライン(-70dBm)付近です。余裕をもったAP配置を推奨します。</div>" } else { "" }
    "<span style='font-size:18px;font-weight:700;color:$rssiColor;'>$($wifi.Signal)</span> <span style='display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;color:#fff;background:$rssiColor;margin-left:8px;vertical-align:middle;'>$($wifi.SignalQuality)</span><div style='margin-top:8px;'><div style='background:#e5e7eb;border-radius:4px;height:10px;width:200px;display:inline-block;vertical-align:middle;position:relative;'><div style='background:$rssiColor;border-radius:4px;height:10px;width:$($barWidth)%;'></div><div style='position:absolute;left:45%;top:-1px;width:2px;height:12px;background:#f59e0b;' title='GW最低ライン: -70dBm'></div></div> <span style='font-size:10px;color:#94a3b8;margin-left:4px;'>▲-70dBm(GW最低ライン)</span></div>$gwNote"
} else { $wifi.Signal })</td></tr>
<tr><td>無線規格</td><td>$($wifi.RadioType)</td></tr>
<tr><td>チャネル</td><td>$($wifi.Channel)</td></tr>
<tr><td>BSSID（APのMACアドレス）</td><td>$($wifi.BSSID)</td></tr>
<tr><td>PCのMACアドレス</td><td>$($ip.MAC)</td></tr>
</table>
</div>

<!-- 自動取得: IP -->
<div class="section">
<div class="section-title"><span class="icon">&#127760;</span> IPアドレス情報 <span class="auto-tag">&#10003; 自動取得</span></div>
<table>
<tr><td>IPアドレス</td><td>$($ip.Address)</td></tr>
<tr><td>サブネットマスク</td><td>$($ip.Mask)</td></tr>
<tr><td>デフォルトゲートウェイ</td><td>$($ip.Gateway)</td></tr>
<tr><td>DNSサーバー（プライマリ）</td><td>$($ip.DNS1)</td></tr>
<tr><td>DNSサーバー（セカンダリ）</td><td>$($ip.DNS2)</td></tr>
<tr><td>DHCP</td><td>$($ip.DHCP)</td></tr>
</table>
</div>

</div>

<!-- 自動取得: IPスキャン結果 -->
<div class="card">
<div class="section">
<div class="section-title"><span class="icon">&#128269;</span> IPアドレススキャン結果 <span class="auto-tag">&#10003; 自動取得</span></div>
<table>
<tr><td>スキャン範囲</td><td>$($scanSummary.NetworkAddr)/$($scanSummary.PrefixLength)</td></tr>
<tr><td>ホスト数（合計）</td><td>$($scanSummary.Total)</td></tr>
<tr><td>使用中 IP</td><td style="color:#dc2626;font-weight:700;">$($scanSummary.Used)</td></tr>
<tr><td>空き IP</td><td style="color:#059669;font-weight:700;">$($scanSummary.Available)</td></tr>
<tr><td>スキャン所要時間</td><td>$($scanSummary.ScanTime)</td></tr>
</table>
</div>

<div class="section">
<div class="section-title" style="color:#dc2626;"><span class="icon">&#128308;</span> 使用中のIPアドレス ($($scanSummary.Used)件)</div>
<div style="line-height:2.2;">
$scanUsedHtml
</div>
</div>

<div class="section">
<div class="section-title" style="color:#059669;"><span class="icon">&#128994;</span> 空きIPアドレス ($($scanSummary.Available)件)</div>
<div style="line-height:2.2;max-height:300px;overflow-y:auto;">
$scanAvailHtml
</div>
<p style="font-size:11px;color:#94a3b8;margin-top:8px;">※ Pingに応答しないデバイスは「空き」として表示される場合があります。ファイアウォール等でICMP（Ping）がブロックされている端末が存在する可能性にご注意ください。</p>
</div>

<div class="section">
<div class="section-title"><span class="icon">&#128268;</span> IPアドレス分布分析 <span class="auto-tag">&#10003; 自動解析</span></div>
$dhcpAnalysisHtml
</div>

</div>

<!-- 手動入力セクション -->
<div class="manual-section">
<div class="section" style="padding-bottom:8px">
<div class="section-title"><span class="icon">&#9999;&#65039;</span> 手動入力項目</div>
<div class="note">以下はPCから自動取得できない項目です。分かる範囲でご記入ください。不明な場合は空欄のままで大丈夫です。</div>
</div>

<div class="section">
<div class="form-row">
<div class="form-group">
<label>Wi-Fiパスワード</label>
<input type="text" id="wifiPass" placeholder="お客様のIT担当者に確認">
</div>
<div class="form-group">
<label>周波数帯（推奨: 5GHz）</label>
<select id="freqBand">
<option value="">-- 選択 --</option>
<option value="2.4GHz">2.4GHz</option>
<option value="5GHz">5GHz（推奨）</option>
<option value="2.4GHz / 5GHz 両方">両方</option>
<option value="不明">不明</option>
</select>
<p class="tip">※ ライフリズムナビは5GHz推奨です</p>
</div>
</div>

<div class="form-row">
<div class="form-group">
<label>プライバシーセパレータ</label>
<select id="privSep">
<option value="">-- 選択 --</option>
<option value="OFF（正常）">OFF（正常）</option>
<option value="ON（要変更）">ON（要変更 ※カメラ映像閲覧不可）</option>
<option value="不明（要確認）">不明（要確認）</option>
</select>
<p class="tip">※ カメラ設置時はOFF必須</p>
</div>
<div class="form-group">
<label>MACアドレス認証</label>
<select id="macAuth">
<option value="">-- 選択 --</option>
<option value="なし">なし</option>
<option value="あり（要登録）">あり（要登録）</option>
<option value="不明（要確認）">不明（要確認）</option>
</select>
<p class="tip">※ ありの場合、全デバイスのMAC登録が必要</p>
</div>
</div>

<div class="form-row">
<div class="form-group">
<label>固定IPの払い出し可否</label>
<select id="staticIP">
<option value="">-- 選択 --</option>
<option value="払い出し可能">払い出し可能</option>
<option value="払い出し不可">払い出し不可</option>
<option value="要確認">要確認</option>
</select>
</div>
<div class="form-group">
<label>払い出し可能な固定IPアドレス</label>
<input type="text" id="availIP" placeholder="例: 192.168.1.100 ～ 110">
<p class="tip">※ カメラ・SIPサーバー・監視用GW分を含む</p>
</div>
</div>

<div class="form-row">
<div class="form-group">
<label>UTM / ファイアウォールの有無</label>
<select id="utmFw">
<option value="">-- 選択 --</option>
<option value="なし">なし</option>
<option value="あり">あり（ホワイトリスト設定が必要）</option>
<option value="不明">不明</option>
</select>
</div>
<div class="form-group">
<label>プロキシ環境</label>
<select id="proxy">
<option value="">-- 選択 --</option>
<option value="なし">なし</option>
<option value="あり">あり（GWにプロキシ設定で対応可）</option>
<option value="不明">不明</option>
</select>
</div>
</div>

<div class="form-row">
<div class="form-group">
<label>キャプティブポータル（ログイン画面）の有無</label>
<select id="captive">
<option value="">-- 選択 --</option>
<option value="なし">なし</option>
<option value="あり（要対策）">あり（※GW接続不可のため要対策）</option>
<option value="不明">不明</option>
</select>
</div>
<div class="form-group">
<label>VLAN分割の有無</label>
<select id="vlan">
<option value="">-- 選択 --</option>
<option value="なし">なし（フラットネットワーク）</option>
<option value="あり">あり（ルーティング設定要確認）</option>
<option value="不明">不明</option>
</select>
</div>
</div>

<div class="form-group" style="padding:0 32px 16px">
<label>備考・特記事項</label>
<textarea id="notes" rows="3" placeholder="NW構成図の有無、特殊なネットワーク構成、ルーター機種名など"></textarea>
</div>

</div>
</div>

<!-- 施設情報 -->
<div class="card facility-card">
<div class="section">
<div class="section-title"><span class="icon">&#127970;</span> 施設情報</div>
<div class="form-row">
<div class="form-group">
<label>施設名</label>
<input type="text" id="facName" placeholder="施設名を入力">
</div>
<div class="form-group">
<label>担当者名（記入者）</label>
<input type="text" id="staff" placeholder="記入者名">
</div>
</div>
<div class="form-group">
<label>調査場所（フロア・居室番号など）</label>
<input type="text" id="location" placeholder="例: 3F ナースステーション付近">
</div>
</div>
</div>

<!-- ボタン -->
<div class="btn-group">
<button class="btn btn-primary" onclick="copyAll()">&#128203; テキストコピー</button>
<button class="btn btn-success" onclick="saveLocal()">&#128190; 保存（HTML）</button>
<button class="btn btn-secondary" onclick="window.print()">&#128424; 印刷 / PDF保存</button>
</div>

<div class="footer">
ライフリズムナビ NW情報収集ツール v1.4（RSSI表示対応） &mdash; エコナビスタ株式会社
</div>

</div>

<script>
function gatherText() {
    var L = [];
    L.push('=== ライフリズムナビ NW情報収集レポート ===');
    L.push('取得日時: $dateStr');
    L.push('PC名: $($ip.Hostname)');
    L.push('インターネット: $netStatus');
    L.push('');
    L.push('【施設情報】');
    L.push('施設名: ' + v('facName'));
    L.push('担当者: ' + v('staff'));
    L.push('調査場所: ' + v('location'));
    L.push('');
    L.push('【Wi-Fi情報（自動取得）】');
    L.push('SSID: $($wifi.SSID)');
    L.push('認証方式: $($wifi.Auth)');
    L.push('暗号化: $($wifi.Cipher)');
    L.push('電波強度（RSSI）: $($wifi.Signal) [$($wifi.SignalQuality)]');
    L.push('無線規格: $($wifi.RadioType)');
    L.push('チャネル: $($wifi.Channel)');
    L.push('BSSID: $($wifi.BSSID)');
    L.push('PCのMAC: $($ip.MAC)');
    L.push('');
    L.push('【IP情報（自動取得）】');
    L.push('IPアドレス: $($ip.Address)');
    L.push('サブネットマスク: $($ip.Mask)');
    L.push('デフォルトGW: $($ip.Gateway)');
    L.push('DNS（プライマリ）: $($ip.DNS1)');
    L.push('DNS（セカンダリ）: $($ip.DNS2)');
    L.push('DHCP: $($ip.DHCP)');
    L.push('');
    L.push('【IPアドレススキャン結果】');
    L.push('スキャン範囲: $($scanSummary.NetworkAddr)/$($scanSummary.PrefixLength)');
    L.push('ホスト数: $($scanSummary.Total)');
    L.push('使用中: $($scanSummary.Used) / 空き: $($scanSummary.Available)');
    L.push('スキャン時間: $($scanSummary.ScanTime)');
    L.push('');
    L.push('使用中IP:');
    L.push('$(($scanUsedText | ForEach-Object { "  $_" }) -join "`n")');
    L.push('');
    L.push('空きIP:');
    L.push('$(($scanAvailText | ForEach-Object { "  $_" }) -join "`n")');
    L.push('');
    L.push('※Ping非応答のデバイスは空きとして表示される場合があります');
    L.push('');
    L.push('【IP分布分析】');
    L.push('');
    L.push('【手動入力項目】');
    L.push('Wi-Fiパスワード: ' + v('wifiPass'));
    L.push('周波数帯: ' + s('freqBand'));
    L.push('プライバシーセパレータ: ' + s('privSep'));
    L.push('MACアドレス認証: ' + s('macAuth'));
    L.push('固定IP払い出し: ' + s('staticIP'));
    L.push('利用可能固定IP: ' + v('availIP'));
    L.push('UTM/FW: ' + s('utmFw'));
    L.push('プロキシ: ' + s('proxy'));
    L.push('キャプティブポータル: ' + s('captive'));
    L.push('VLAN: ' + s('vlan'));
    L.push('備考: ' + v('notes'));
    return L.join('\n');
}

function v(id) { return document.getElementById(id).value || '未入力'; }
function s(id) { var el = document.getElementById(id); return el.options[el.selectedIndex].text || '未選択'; }

function copyAll() {
    var text = gatherText();
    navigator.clipboard.writeText(text).then(function() {
        alert('クリップボードにコピーしました！\nメールやチャットに貼り付けできます。');
    }).catch(function() {
        // フォールバック
        var ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        alert('コピーしました！');
    });
}

function saveLocal() {
    // 手動入力値をHTMLに反映して保存
    var blob = new Blob([document.documentElement.outerHTML], {type:'text/html;charset=utf-8'});
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    var fname = (document.getElementById('facName').value || 'NW情報') + '_' + new Date().toISOString().slice(0,10) + '.html';
    a.download = fname;
    a.click();
    URL.revokeObjectURL(a.href);
}
</script>

</body>
</html>
"@

# UTF-8 BOM付きで書き出し（ブラウザ互換性のため）
[System.IO.File]::WriteAllText($outputFile, $html, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "  ✅ 完了しました！" -ForegroundColor Green
Write-Host "  📄 保存先: $outputFile" -ForegroundColor White
Write-Host ""

# ブラウザで開く
Start-Process $outputFile

Write-Host "  ブラウザでレポートが開きます。このウィンドウは閉じて構いません。" -ForegroundColor Gray
Start-Sleep -Seconds 5
