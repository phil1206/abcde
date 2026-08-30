# ============================================================
# update.ps1 — 台股零股儀表板 每日資料更新程式
# 資料來源：台灣證券交易所 OpenAPI（免費公開資料）
# 執行方式：由工作排程器每日自動執行，或雙擊「更新資料.bat」
# ============================================================
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataDir = Join-Path $root 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory $dataDir | Out-Null }
$logFile = Join-Path $root 'update.log'

function Log($m) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts  $m" | Add-Content -Path $logFile -Encoding UTF8
    Write-Host $m
}

$UA = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
function Fetch($url) { Invoke-RestMethod -Uri $url -Headers $UA -TimeoutSec 60 }

# 民國日期轉西元 ISO（支援 "1150828" 與 "115/08/28"）
function RocToIso([string]$s) {
    $s = $s -replace '/', ''
    if ($s.Length -lt 7) { return $null }
    $y = [int]$s.Substring(0, $s.Length - 4) + 1911
    $m = $s.Substring($s.Length - 4, 2)
    $d = $s.Substring($s.Length - 2, 2)
    return "$y-$m-$d"
}
function ToNum($s) {
    if ($null -eq $s) { return $null }
    $s = "$s" -replace ',', ''
    if ($s -eq '' -or $s -eq '--' -or $s -eq '-') { return $null }
    try { return [double]$s } catch { return $null }
}

# ---------- 觀察清單（零股友善的知名標的，可自行增減代號） ----------
$WATCH = @('0050','0056','00878','006208','00919','2330','2317','2454','2412','2882','2891','2886','1216','2002','2603')

try {
    Log '===== 開始更新 ====='

    # ---------- 1. 抓取當日全市場資料 ----------
    $all    = Fetch 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL'
    $bwibbu = Fetch 'https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL'
    $mkt    = Fetch 'https://openapi.twse.com.tw/v1/exchangeReport/FMTQIK'
    $rev    = Fetch 'https://openapi.twse.com.tw/v1/opendata/t187ap05_L'
    $dataDate    = RocToIso $all[0].Date
    $dataDateNum = $dataDate -replace '-', ''
    Log "行情資料日期：$dataDate（共 $($all.Count) 檔）"

    # 三大法人買賣超（以行情日期查詢）
    $t86 = $null
    try {
        $t86raw = Fetch "https://www.twse.com.tw/rwd/zh/fund/T86?date=$dataDateNum&selectType=ALLBUT0999&response=json"
        if ($t86raw.stat -eq 'OK') { $t86 = $t86raw }
    } catch { Log "T86 抓取失敗：$($_.Exception.Message)" }

    # ---------- 2. 建立查表 ----------
    $priceMap = @{}
    foreach ($r in $all) {
        $c = ToNum $r.ClosingPrice
        if ($null -ne $c) {
            $priceMap[$r.Code] = @{
                name = $r.Name; close = $c; chg = (ToNum $r.Change)
                vol = (ToNum $r.TradeVolume); tx = (ToNum $r.Transaction)
            }
        }
    }
    $peMap = @{}
    foreach ($r in $bwibbu) {
        $peMap[$r.Code] = @{ pe = (ToNum $r.PEratio); yield = (ToNum $r.DividendYield); pb = (ToNum $r.PBratio) }
    }
    $revMap = @{}
    foreach ($r in $rev) {
        $revMap[$r.'公司代號'] = @{ yoy = (ToNum $r.'營業收入-去年同月增減(%)'); ym = $r.'資料年月'; ind = $r.'產業別' }
    }
    $instMap = @{}
    if ($t86) {
        $fld  = $t86.fields
        $iF   = [array]::IndexOf($fld, '外陸資買賣超股數(不含外資自營商)')
        $iT   = [array]::IndexOf($fld, '投信買賣超股數')
        $iAll = [array]::IndexOf($fld, '三大法人買賣超股數')
        foreach ($row in $t86.data) {
            $code = "$($row[0])".Trim()
            $instMap[$code] = @{ f = (ToNum $row[$iF]); t = (ToNum $row[$iT]); net = (ToNum $row[$iAll]) }
        }
        Log "三大法人資料：$($t86.data.Count) 檔"
    } else {
        Log '三大法人資料暫無（可能為假日或尚未公布）'
    }

    # ---------- 3. 歷史資料（累積收盤價與法人動向，供均線/連買計算） ----------
    $histFile = Join-Path $dataDir 'history.json'
    $hist = @{}
    if (Test-Path $histFile) {
        $hist = Get-Content $histFile -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
        if ($null -eq $hist) { $hist = @{} }
    }

    # 首次執行：回補觀察清單近 5 個月的收盤價
    foreach ($code in $WATCH) {
        if (-not $hist.ContainsKey($code)) { $hist[$code] = @() }
        if (@($hist[$code]).Count -lt 65) {
            Log "回補 $code 歷史股價..."
            $merged = @{}
            foreach ($e in @($hist[$code])) { $merged[$e.d] = $e }
            for ($i = 4; $i -ge 0; $i--) {
                $qm = (Get-Date $dataDate).AddMonths(-$i).ToString('yyyyMM') + '01'
                try {
                    $h = Fetch "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=$qm&stockNo=$code&response=json"
                    if ($h.stat -eq 'OK') {
                        foreach ($row in $h.data) {
                            $d = RocToIso "$($row[0])"; $c = ToNum $row[6]
                            if ($d -and $null -ne $c) { $merged[$d] = @{ d = $d; c = $c } }
                        }
                    }
                } catch { Log "  $code $qm 回補失敗：$($_.Exception.Message)" }
                Start-Sleep -Milliseconds 700
            }
            $hist[$code] = @($merged.Values | Sort-Object d)
        }
    }

    # 每日累積：合併本月每日收盤（多日未開機也能自動補齊缺漏），並寫入當日法人買賣超
    foreach ($code in ($WATCH + '_TAIEX')) {
        if (-not $hist.ContainsKey($code)) { $hist[$code] = @() }
    }
    foreach ($code in $WATCH) {
        $byDate = @{}
        foreach ($e in @($hist[$code])) { $byDate[$e.d] = $e }
        $months = @((Get-Date $dataDate).ToString('yyyyMM'))
        if ((Get-Date $dataDate).Day -le 10) { $months += (Get-Date $dataDate).AddMonths(-1).ToString('yyyyMM') }
        foreach ($m in $months) {
            try {
                $h = Fetch "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=${m}01&stockNo=$code&response=json"
                if ($h.stat -eq 'OK') {
                    foreach ($row in $h.data) {
                        $d = RocToIso "$($row[0])"; $c = ToNum $row[6]
                        if ($d -and $null -ne $c) {
                            if ($byDate.ContainsKey($d)) { $byDate[$d].c = $c }
                            else { $byDate[$d] = @{ d = $d; c = $c } }
                        }
                    }
                }
            } catch { Log "  $code 本月行情合併失敗：$($_.Exception.Message)" }
            Start-Sleep -Milliseconds 400
        }
        if ($instMap.ContainsKey($code) -and $byDate.ContainsKey($dataDate)) {
            $byDate[$dataDate].f = $instMap[$code].f; $byDate[$dataDate].t = $instMap[$code].t
        }
        $hist[$code] = @($byDate.Values | Sort-Object d)
    }
    # 大盤歷史（FMTQIK 為當月每日）
    $taiexArr = @($hist['_TAIEX'])
    foreach ($r in $mkt) {
        $d = RocToIso $r.Date
        if ($d -and -not ($taiexArr | Where-Object { $_.d -eq $d })) {
            $taiexArr += @{ d = $d; c = (ToNum $r.TAIEX); chg = (ToNum $r.Change) }
        }
    }
    $hist['_TAIEX'] = @($taiexArr | Sort-Object d | Select-Object -Last 120)

    # 補繳：電腦多日未開機時，補回缺漏日的法人動向與上榜紀錄（最多回補 8 個交易日）
    if (-not $hist.ContainsKey('_PICKS')) { $hist['_PICKS'] = @() }
    $pickDates = @($hist['_PICKS'] | ForEach-Object { $_.d })
    $missDays = @($hist['_TAIEX'] | ForEach-Object { $_.d } | Where-Object { $_ -lt $dataDate -and $_ -notin $pickDates } | Sort-Object | Select-Object -Last 8)
    foreach ($day in $missDays) {
        $dn = $day -replace '-', ''
        Log "補繳 $day 的法人與上榜紀錄..."
        try {
            $t86x = Fetch "https://www.twse.com.tw/rwd/zh/fund/T86?date=$dn&selectType=ALLBUT0999&response=json"
            Start-Sleep -Milliseconds 600
            $mix = Fetch "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date=$dn&type=ALLBUT0999&response=json"
            Start-Sleep -Milliseconds 600
            $bwx = Fetch "https://www.twse.com.tw/rwd/zh/afterTrading/BWIBBU_d?date=$dn&selectType=ALL&response=json"
            Start-Sleep -Milliseconds 600
            if ($t86x.stat -ne 'OK' -or $mix.stat -ne 'OK' -or $bwx.stat -ne 'OK') { Log '  該日資料不齊，跳過'; continue }
            $fldx = $t86x.fields
            $ic = [array]::IndexOf($fldx, '證券代號'); $if = [array]::IndexOf($fldx, '外陸資買賣超股數(不含外資自營商)')
            $it = [array]::IndexOf($fldx, '投信買賣超股數'); $in2 = [array]::IndexOf($fldx, '三大法人買賣超股數')
            $imap = @{}
            foreach ($row in $t86x.data) { $imap["$($row[$ic])".Trim()] = @{ f = (ToNum $row[$if]); t = (ToNum $row[$it]); net = (ToNum $row[$in2]) } }
            foreach ($code in $WATCH) {
                if (-not $imap.ContainsKey($code)) { continue }
                foreach ($e in @($hist[$code])) {
                    if ($e.d -eq $day) { $e.f = $imap[$code].f; $e.t = $imap[$code].t }
                }
            }
            $qt = $mix.tables | Where-Object { $_.fields -and ($_.fields -contains '證券代號') -and ($_.fields -contains '收盤價') } | Select-Object -First 1
            $ic1 = [array]::IndexOf($qt.fields, '證券代號'); $icl = [array]::IndexOf($qt.fields, '收盤價'); $iv = [array]::IndexOf($qt.fields, '成交股數')
            $pmx = @{}
            foreach ($row in $qt.data) { $c2 = "$($row[$ic1])".Trim(); $cl2 = ToNum $row[$icl]; if ($null -ne $cl2) { $pmx[$c2] = @{ close = $cl2; vol = (ToNum $row[$iv]) } } }
            $ic3 = [array]::IndexOf($bwx.fields, '證券代號'); $iy = [array]::IndexOf($bwx.fields, '殖利率(%)'); $ip = [array]::IndexOf($bwx.fields, '本益比')
            $bmx = @{}
            foreach ($row in $bwx.data) { $bmx["$($row[$ic3])".Trim()] = @{ y = (ToNum $row[$iy]); pe = (ToNum $row[$ip]) } }
            $codes = @()
            foreach ($c2 in $imap.Keys) {
                if (-not $pmx.ContainsKey($c2)) { continue }
                $p2 = $pmx[$c2]
                if ($null -eq $p2.vol -or $p2.vol -lt 500000 -or $p2.close -lt 5) { continue }
                $sc = 0
                if ($null -ne $imap[$c2].net -and ($imap[$c2].net * $p2.close / 1e8) -ge 1) { $sc++ }
                $b2 = $bmx[$c2]
                if ($b2 -and $null -ne $b2.y -and $b2.y -ge 3.5) { $sc++ }
                if ($b2 -and $null -ne $b2.pe -and $b2.pe -gt 0 -and $b2.pe -le 20) { $sc++ }
                if ($revMap.ContainsKey($c2) -and $null -ne $revMap[$c2].yoy -and $revMap[$c2].yoy -ge 20) { $sc++ }
                if ($sc -ge 3) { $codes += $c2 }
            }
            $hist['_PICKS'] = @(@($hist['_PICKS']) + @{ d = $day; codes = $codes } | Sort-Object { $_.d })
            Log "  $day 補回 $($codes.Count) 檔上榜"
        } catch { Log "  補繳 $day 失敗：$($_.Exception.Message)" }
    }
    # 每檔最多保留 200 筆
    foreach ($code in $WATCH) { $hist[$code] = @($hist[$code] | Sort-Object d | Select-Object -Last 200) }
    $hist | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $histFile -Encoding UTF8

    # ---------- 4. 計算觀察清單的均線、訊號與參考價位 ----------
    function MA($closes, $n) {
        if ($closes.Count -lt $n) { return $null }
        $s = ($closes | Select-Object -Last $n | Measure-Object -Sum).Sum
        return [math]::Round($s / $n, 2)
    }
    function Streak($arr, $key) {
        # 從最新往回數，連續買超（正值）的天數；一遇到非正值即停
        $n = 0
        for ($i = $arr.Count - 1; $i -ge 0; $i--) {
            $v = $arr[$i][$key]
            if ($null -ne $v -and $v -gt 0) { $n++ } elseif ($null -ne $v) { break } else { break }
        }
        return $n
    }

    $watchOut = @()
    foreach ($code in $WATCH) {
        if (-not $priceMap.ContainsKey($code)) { continue }
        $p = $priceMap[$code]
        $closes = @($hist[$code] | ForEach-Object { $_.c })
        $ma5  = MA $closes 5
        $ma20 = MA $closes 20
        $ma60 = MA $closes 60
        $prev = $p.close - $p.chg
        $chgPct = if ($prev -gt 0) { [math]::Round($p.chg / $prev * 100, 2) } else { $null }

        $signal = '資料累積中'; $advice = '歷史資料累積中，訊號稍後提供'
        if ($ma20 -and $ma60) {
            if ($p.close -gt $ma20 -and $ma20 -gt $ma60) {
                $signal = '偏多'
                $advice = "均線多頭排列，趨勢向上。參考做法：可分批小量布局；若跌破20日線（約 $ma20 元）先減碼觀察。"
            } elseif ($p.close -lt $ma60 -and $ma20 -lt $ma60) {
                $signal = '偏空'
                $advice = "均線空頭排列，趨勢偏弱。參考做法：暫緩進場；持股者若跌破前低可考慮停損，站回60日線（約 $ma60 元）前保守以對。"
            } else {
                $signal = '中性'
                $advice = "均線糾結、區間整理。參考做法：先觀望，等股價站穩20日線（約 $ma20 元）之上再考慮分批進場。"
            }
        }

        $inst = if ($instMap.ContainsKey($code)) { $instMap[$code] } else { $null }
        $instNetVal = if ($inst -and $null -ne $inst.net) { [math]::Round($inst.net * $p.close / 1e8, 2) } else { $null }
        $revI = if ($revMap.ContainsKey($code)) { $revMap[$code] } else { $null }
        $peI  = if ($peMap.ContainsKey($code)) { $peMap[$code] } else { $null }

        $watchOut += [ordered]@{
            code = $code; name = $p.name
            price = $p.close; chg = $p.chg; chgPct = $chgPct
            pe = $peI.pe; yield = $peI.yield; pb = $peI.pb
            ma5 = $ma5; ma20 = $ma20; ma60 = $ma60
            signal = $signal; advice = $advice
            foreignNet = $inst.f; trustNet = $inst.t; instNet = $inst.net; instNetVal = $instNetVal
            fStreak = (Streak @($hist[$code]) 'f'); tStreak = (Streak @($hist[$code]) 't')
            revYoY = $revI.yoy; revYm = $revI.ym; industry = $revI.ind
            histClose = @($hist[$code] | Select-Object -Last 60 | ForEach-Object { $_.c })
        }
    }

    # ---------- 5. 全市場篩選 ----------
    # (a) 三大法人買超金額排行
    $screenInst = @()
    if ($t86) {
        $screenInst = $instMap.Keys | Where-Object { $priceMap.ContainsKey($_) -and $null -ne $instMap[$_].net } |
            ForEach-Object {
                $c = $_; $p = $priceMap[$c]
                [ordered]@{
                    code = $c; name = $p.name; price = $p.close
                    instNet = $instMap[$c].net
                    instNetVal = [math]::Round($instMap[$c].net * $p.close / 1e8, 2)
                    pe = $peMap[$c].pe; yield = $peMap[$c].yield
                }
            } | Sort-Object { $_.instNetVal } -Descending | Select-Object -First 15
    }
    # (b) 高殖利率零股精選（殖利率>=4%、股價<=200、本益比 0~25、有一定成交量）
    $screenYield = $peMap.Keys | Where-Object {
            $priceMap.ContainsKey($_) -and
            $null -ne $peMap[$_].yield -and $peMap[$_].yield -ge 4 -and
            $null -ne $peMap[$_].pe -and $peMap[$_].pe -gt 0 -and $peMap[$_].pe -le 25 -and
            $priceMap[$_].close -le 200 -and $priceMap[$_].close -ge 8 -and
            $priceMap[$_].vol -ge 300000
        } | ForEach-Object {
            $c = $_; $p = $priceMap[$c]
            [ordered]@{
                code = $c; name = $p.name; price = $p.close
                pe = $peMap[$c].pe; yield = $peMap[$c].yield; pb = $peMap[$c].pb
                revYoY = $revMap[$c].yoy
            }
        } | Sort-Object { $_.yield } -Descending | Select-Object -First 15
    # (c) 營收高成長（去年同月增減 >= 30%，且買得起、有量）
    $screenRev = $revMap.Keys | Where-Object {
            $priceMap.ContainsKey($_) -and
            $null -ne $revMap[$_].yoy -and $revMap[$_].yoy -ge 30 -and
            $priceMap[$_].close -le 500 -and $priceMap[$_].vol -ge 300000
        } | ForEach-Object {
            $c = $_; $p = $priceMap[$c]
            [ordered]@{
                code = $c; name = $p.name; price = $p.close
                revYoY = [math]::Round($revMap[$c].yoy, 1); industry = $revMap[$c].ind
                pe = $peMap[$c].pe; yield = $peMap[$c].yield
            }
        } | Sort-Object { $_.revYoY } -Descending | Select-Object -First 15

    # (d) 每日自動精選：籌碼＋基本面綜合評分，附推薦理由
    $picks = @()
    foreach ($c in $priceMap.Keys) {
        $p = $priceMap[$c]
        if ($null -eq $p.vol -or $p.vol -lt 500000 -or $p.close -lt 5) { continue }
        $reasons = @(); $score = 0
        $inst = if ($instMap.ContainsKey($c)) { $instMap[$c] } else { $null }
        $instVal = if ($inst -and $null -ne $inst.net) { $inst.net * $p.close / 1e8 } else { $null }
        if ($null -ne $instVal -and $instVal -ge 1) { $score++; $reasons += "三大法人今日買超 $([math]::Round($instVal,1)) 億（籌碼轉強）" }
        $pi = if ($peMap.ContainsKey($c)) { $peMap[$c] } else { $null }
        if ($pi -and $null -ne $pi.yield -and $pi.yield -ge 3.5) { $score++; $reasons += "殖利率 $($pi.yield)%（存股領息有底）" }
        if ($pi -and $null -ne $pi.pe -and $pi.pe -gt 0 -and $pi.pe -le 20) { $score++; $reasons += "本益比 $($pi.pe) 倍（估值不貴）" }
        $ri = if ($revMap.ContainsKey($c)) { $revMap[$c] } else { $null }
        if ($ri -and $null -ne $ri.yoy -and $ri.yoy -ge 20) { $score++; $reasons += "最新月營收年增 $([math]::Round($ri.yoy,1))%（成長動能）" }
        if ($score -ge 3) {
            $picks += [ordered]@{
                code = $c; name = $p.name; price = $p.close; score = $score
                instNetVal = if ($null -ne $instVal) { [math]::Round($instVal,2) } else { $null }
                reasons = ($reasons -join '；')
            }
        }
    }
    $picksAll = @($picks)
    $picks = @($picksAll | Sort-Object { $_.score }, { $_.instNetVal } -Descending | Select-Object -First 5)

    # 累積每日上榜名單（存全部符合門檻者，不只前五），計算連續上榜天數
    if (-not $hist.ContainsKey('_PICKS')) { $hist['_PICKS'] = @() }
    $pickHist = @($hist['_PICKS'] | Where-Object { $_.d -ne $dataDate })
    $pickHist += @{ d = $dataDate; codes = @($picksAll | ForEach-Object { $_.code }) }
    $hist['_PICKS'] = @($pickHist | Sort-Object { $_.d } | Select-Object -Last 30)
    $ph = @($hist['_PICKS'])
    foreach ($p in $picks) {
        $n = 0
        for ($i = $ph.Count - 1; $i -ge 0; $i--) {
            if (@($ph[$i].codes) -contains $p.code) { $n++ } else { break }
        }
        $p.streak = $n
    }
    $hist | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $histFile -Encoding UTF8

    # ---------- 5.5 昨夜美股與市場天氣燈號 ----------
    $us = $null
    try {
        $usItems = @()
        $usDefs = @(
            @{ sym = '^IXIC'; name = '那斯達克' },
            @{ sym = '^SOX';  name = '費城半導體' },
            @{ sym = 'TSM';   name = '台積電ADR' }
        )
        foreach ($u in $usDefs) {
            $q = Fetch "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($u.sym))?range=10d&interval=1d"
            $res = $q.chart.result[0]; $meta = $res.meta
            $ts = @($res.timestamp); $cl = @($res.indicators.quote[0].close)
            $pairs = @()
            for ($i = 0; $i -lt $cl.Count; $i++) { if ($null -ne $cl[$i]) { $pairs += @{ t = $ts[$i]; c = [double]$cl[$i] } } }
            if ($pairs.Count -lt 2) { continue }
            $mDate = [DateTimeOffset]::FromUnixTimeSeconds([long]$meta.regularMarketTime).UtcDateTime.Date
            $bDate = [DateTimeOffset]::FromUnixTimeSeconds([long]$pairs[-1].t).UtcDateTime.Date
            if ($bDate -eq $mDate) { $latest = $pairs[-1].c; $prev = $pairs[-2].c }
            else { $latest = [double]$meta.regularMarketPrice; $prev = $pairs[-1].c }
            $usItems += [ordered]@{
                sym = $u.sym; name = $u.name
                price = [math]::Round($latest, 2)
                pct = [math]::Round(($latest - $prev) / $prev * 100, 2)
                date = $mDate.ToString('MM/dd')
            }
            Start-Sleep -Milliseconds 400
        }
        if ($usItems.Count -ge 2) {
            $g = @{}; foreach ($x in $usItems) { $g[$x.sym] = $x.pct }
            $sox  = if ($g.ContainsKey('^SOX'))  { $g['^SOX'] }  else { 0 }
            $ixic = if ($g.ContainsKey('^IXIC')) { $g['^IXIC'] } else { 0 }
            $tsm  = if ($g.ContainsKey('TSM'))   { $g['TSM'] }   else { 0 }
            if ($sox -le -2 -or $ixic -le -2 -or $tsm -le -3) {
                $mood = 'red'; $moodText = '🔴 美股重挫'
                $advice = '最近一晚美股大跌，台股很可能跳空開低。建議今日暫緩新買單（已掛的預約單考慮取消），等市場止穩一兩天再照計畫進行。'
            } elseif ($sox -le -1 -or $ixic -le -0.75 -or $tsm -le -1.5) {
                $mood = 'yellow'; $moodText = '🟡 美股偏弱'
                $advice = '最近一晚美股偏弱，台股容易開低震盪。買單保守一點：金額減半、限價往下掛，接到便宜貨是好事，接不到也不勉強。'
            } elseif ($sox -ge 2 -or $ixic -ge 1.5) {
                $mood = 'hot'; $moodText = '🟠 美股大漲'
                $advice = '最近一晚美股大漲，台股很可能跳空開高。限價單不要往上追價，買不到就等回檔——開高追買是新手最常見的虧錢方式。'
            } else {
                $mood = 'green'; $moodText = '🟢 市場平穩'
                $advice = '最近一晚美股平穩，沒有特殊風險訊號，照你原本的計畫執行即可。'
            }
            $us = [ordered]@{ items = $usItems; mood = $mood; moodText = $moodText; advice = $advice }
            Log "美股天氣：$moodText（費半 $sox%／那指 $ixic%／台積ADR $tsm%）"
        }
    } catch { Log "美股資料抓取失敗：$($_.Exception.Message)" }

    # ---------- 6. 輸出 data.js ----------
    $taiexLast = $hist['_TAIEX'] | Select-Object -Last 1
    $out = [ordered]@{
        updated  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        dataDate = $dataDate
        t86Date  = if ($t86) { "$($t86.date)".Insert(6,'-').Insert(4,'-') } else { $null }
        taiex    = [ordered]@{
            last = $taiexLast.c; chg = $taiexLast.chg
            history = @($hist['_TAIEX'] | Select-Object -Last 60)
        }
        us          = $us
        watch       = $watchOut
        picks       = @($picks)
        screenInst  = @($screenInst)
        screenYield = @($screenYield)
        screenRev   = @($screenRev)
    }
    $json = $out | ConvertTo-Json -Depth 8 -Compress
    "window.DASH_DATA = $json;" | Set-Content -Path (Join-Path $dataDir 'data.js') -Encoding UTF8
    Log "更新完成：觀察 $($watchOut.Count) 檔／法人榜 $(@($screenInst).Count)／殖利率榜 $(@($screenYield).Count)／營收榜 $(@($screenRev).Count)"

    # ---------- 7. 自動上傳到 GitHub Pages（手機版網站） ----------
    if (Test-Path (Join-Path $root '.git')) {
        try {
            git -C $root add -A 2>&1 | Out-Null
            $st = git -C $root status --porcelain
            if ($st) {
                git -C $root commit -m "每日資料更新 $(Get-Date -Format 'yyyy-MM-dd HH:mm')" 2>&1 | Out-Null
                git -C $root push 2>&1 | Out-Null
                Log '已上傳到 GitHub Pages'
            } else {
                Log '資料無變化，免上傳'
            }
        } catch { Log "GitHub 上傳失敗（本機網站不受影響）：$($_.Exception.Message)" }
    }
} catch {
    Log "更新失敗：$($_.Exception.Message)"
    exit 1
}
