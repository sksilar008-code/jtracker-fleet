# EdgeCdp.psm1 - Automate Microsoft Edge via Chrome DevTools Protocol (no external libraries)
# Uses built-in System.Net.WebSockets only. Works on Windows PowerShell 5.1.

$script:EdgeProc = $null
$script:WS = $null
$script:MsgId = 0
$script:Port = 9333
$script:InRepair = $false    # guards Repair-CdpConnection re-entry (its verify uses Send-Cdp)
$script:InRetry = $false     # one auto-retry per Send-Cdp call
$script:RepairFailAt = $null # cooldown timestamp after a failed repair (no repeated 5s stalls)

function Start-EdgeCdp {
    param([switch]$Headless, [int]$Port = 9333)
    $script:Port = $Port
    $edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if (-not (Test-Path $edge)) { $edge = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }
    # LIVE profile on LOCAL disk, OUTSIDE OneDrive. The old location (project folder,
    # inside OneDrive) made OneDrive sync the live profile while Edge wrote to it
    # (cookies/sqlite/Local State) - sync locks stalled the browser and killed CDP
    # sockets mid-run (mass "connection lost" failures in the 10/27 12:00 pipeline run).
    # Logins still survive between runs via the cookie JAR (edge-profile\cookies.json in
    # the project dir) which every bot restores via Network.setCookies at session start.
    $ud = Join-Path $env:LOCALAPPDATA "JTrackerEdge\profile"
    New-Item -ItemType Directory -Force -Path $ud | Out-Null
    $args = @("--remote-debugging-port=$Port", "--user-data-dir=`"$ud`"", "--no-first-run", "--no-default-browser-check", "--start-maximized", "--disable-blink-features=AutomationControlled", "about:blank")
    if ($Headless) { $args = @("--headless=new") + $args }
    # Cloud mode (GitHub Actions runner - no interactive desktop): force headless with a
    # fixed viewport so page layouts match what the bots expect from a maximized window.
    if ($env:JTRACKER_CLOUD -eq "1") {
        $args = @("--headless=new", "--disable-gpu", "--window-size=1600,1000") + $args
    }
    # Kill any stale bot-Edge processes first (prevents port fights + zombie pile-ups)
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "JTrackerEdge|edge-profile" } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {} }
    for ($w = 0; $w -lt 15; $w++) {
        $still = @(Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "JTrackerEdge|edge-profile" })
        if ($still.Count -eq 0) { break }
        Start-Sleep -Seconds 1
    }
    # Also wait for the DEBUG PORT to actually stop answering - a dying Edge can keep
    # the port open briefly. If we connect to that dying instance the socket dies on
    # first use ("connection lost") which is what broke pipeline runs (FeedBot after
    # 3 prior bots). Only start once the port is truly free.
    for ($w = 0; $w -lt 20; $w++) {
        $busy = $false
        try { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 1; $busy = $true } catch {}
        if (-not $busy) { break }
        # port still answering - nuke anything holding it, then re-check
        Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "JTrackerEdge|edge-profile|remote-debugging-port=$Port" } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
        Start-Sleep -Seconds 1
    }
    Start-Sleep -Seconds 2

    # Launch + connect, with a real round-trip verification. Retry the whole start
    # once if the socket comes up dead (leftover half-closed Edge from prior bot).
    $maxStartTries = 2
    for ($attempt = 1; $attempt -le $maxStartTries; $attempt++) {
        $script:EdgeProc = Start-Process $edge -ArgumentList ($args -join " ") -PassThru
        # Wait for the debugging endpoint
        $ready = $false
        for ($i = 0; $i -lt 60; $i++) {
            try {
                $v = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 2
                if ($v) { $ready = $true; break }
            } catch { Start-Sleep -Milliseconds 500 }
        }
        if (-not $ready) { throw "Edge DevTools endpoint did not start on port $Port" }
        # Connect to the page target
        $targets = Invoke-RestMethod "http://127.0.0.1:$Port/json" -TimeoutSec 5
        $page = $targets | Where-Object { $_.type -eq "page" } | Select-Object -First 1
        if (-not $page) { throw "No page target found" }
        $script:WS = New-Object System.Net.WebSockets.ClientWebSocket
        $script:WS.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
        $script:WS.ConnectAsync([uri]$page.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
        Send-Cdp "Page.enable" | Out-Null
        Send-Cdp "DOM.enable" | Out-Null
        Send-Cdp "Runtime.enable" | Out-Null
        # Verify the connection is genuinely alive with a real evaluate round-trip
        $alive = $false
        try { if ((Invoke-CdpJs "1+1" -TimeoutSec 10) -eq 2) { $alive = $true } } catch {}
        if ($alive) { return }
        # dead socket - tear down and retry once
        try { Stop-EdgeCdp } catch {}
        Start-Sleep -Seconds 3
    }
    throw "Edge CDP connection failed verification after $maxStartTries attempts"
}

function Enable-Stealth {
    # Makes the bot look like a normal human browser (call right after Start-EdgeCdp)
    Send-Cdp "Emulation.setUserAgentOverride" @{
        userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36 Edg/154.0.0.0"
        acceptLanguage = "en-US,en;q=0.9"
    } | Out-Null
    Send-Cdp "Page.addScriptToEvaluateOnNewDocument" @{
        source = "Object.defineProperty(navigator,'webdriver',{get:()=>undefined}); Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]}); Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']}); window.chrome={runtime:{}};"
    } | Out-Null
}

function Enable-AssetBlocking {
    # Block heavy subresources (images/fonts/media) - job pages are TEXT. On slow networks
    # unblocked media keeps the renderer main thread busy past Chromium's hang-watchdog
    # threshold, which KILLS the renderer and closes the CDP socket ("connection lost" at
    # a consistent ~30s). Every session death in the 10/27 13:14 run happened exactly
    # like this while the same URLs loaded fine in fast network windows.
    Send-Cdp "Network.enable" | Out-Null
    Send-Cdp "Network.setBlockedURLs" @{
        urls = @("*.png","*.jpg","*.jpeg","*.gif","*.webp","*.avif","*.svg","*.woff","*.woff2","*.ttf","*.otf","*.eot","*.mp4","*.webm","*.mp3","*.avi","*.mov")
    } | Out-Null
}

function Invoke-CdpRealClick {
    # Real mouse click (pressed+released) at the center of an element found by a JS expression.
    # Triggers React/Angular handlers that synthetic JS .click() misses.
    param([string]$JsElementExpr)
    $rect = Invoke-CdpJs "(function(){ var el = ($JsElementExpr); if(!el || !(el.offsetWidth||el.offsetHeight)) return null; el.scrollIntoView({block:'center'}); var r = el.getBoundingClientRect(); return JSON.stringify({x: r.left + r.width/2, y: r.top + r.height/2}); })()"
    if (-not $rect) { return $false }
    $p = $rect | ConvertFrom-Json
    Send-Cdp "Input.dispatchMouseEvent" @{ type = "mousePressed"; x = [int]$p.x; y = [int]$p.y; button = "left"; clickCount = 1 } | Out-Null
    Start-Sleep -Milliseconds 80
    Send-Cdp "Input.dispatchMouseEvent" @{ type = "mouseReleased"; x = [int]$p.x; y = [int]$p.y; button = "left"; clickCount = 1 } | Out-Null
    return $true
}

function Switch-EdgeTab {
    # If clicking opened a NEW TAB, hop to it. Returns $true if switched.
    param([string]$PreviousUrl)
    try {
        $targets = Invoke-RestMethod "http://127.0.0.1:$script:Port/json" -TimeoutSec 5
        $pages = @($targets | Where-Object { $_.type -eq "page" -and $_.url -ne $PreviousUrl -and $_.url -notmatch "^about:blank" })
        if ($pages.Count -eq 0) { return $false }
        $newPage = $pages[-1]
        try { if ($script:WS -and $script:WS.State -eq "Open") { $script:WS.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [Threading.CancellationToken]::None).Wait() } } catch {}
        $script:WS = New-Object System.Net.WebSockets.ClientWebSocket
        $script:WS.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
        $script:WS.ConnectAsync([uri]$newPage.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
        Send-Cdp "Page.enable" | Out-Null
        Send-Cdp "DOM.enable" | Out-Null
        Send-Cdp "Runtime.enable" | Out-Null
        Start-Sleep -Seconds 3
        return $true
    } catch { return $false }
}

function Stop-EdgeCdp {
    try { if ($script:WS -and $script:WS.State -eq "Open") { $script:WS.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [Threading.CancellationToken]::None).Wait() } } catch {}
    try { if ($script:EdgeProc -and -not $script:EdgeProc.HasExited) { $script:EdgeProc.Kill() } } catch {}
    $script:WS = $null; $script:EdgeProc = $null
}

function Connect-EdgeCdp {
    # Attach to an ALREADY RUNNING Edge that has the debug port open (e.g. the user's login window)
    param([int]$Port = 9333)
    $script:Port = $Port
    $targets = Invoke-RestMethod "http://127.0.0.1:$Port/json" -TimeoutSec 5
    $page = $targets | Where-Object { $_.type -eq "page" } | Select-Object -First 1
    if (-not $page) { throw "No page target found on port $Port" }
    $script:WS = New-Object System.Net.WebSockets.ClientWebSocket
    $script:WS.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
    $script:WS.ConnectAsync([uri]$page.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
}

function Repair-CdpConnection {
    # Re-attach a dead WebSocket to the STILL-RUNNING browser. 10/27 experiment finding:
    # on "connection lost" deaths the msedge process survives (/json/version keeps
    # answering) - only the page-target WS closes. Reconnecting takes ~1s vs ~10s+ for a
    # full browser restart and preserves all session/page state. Returns $true on success.
    if ($script:InRepair) { return $false }
    if ($script:RepairFailAt -and ((Get-Date) - $script:RepairFailAt).TotalSeconds -lt 10) { return $false }
    $script:InRepair = $true
    try {
        $port = if ($script:Port) { $script:Port } else { 9333 }
        $targets = $null
        try { $targets = Invoke-RestMethod "http://127.0.0.1:$port/json" -TimeoutSec 5 } catch { return $false }
        $page = @($targets | Where-Object { $_.type -eq "page" -and $_.webSocketDebuggerUrl }) | Select-Object -First 1
        if (-not $page) { return $false }
        if ($script:WS) { try { $script:WS.Dispose() } catch {}; $script:WS = $null }
        $ws = New-Object System.Net.WebSockets.ClientWebSocket
        $ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
        $ws.ConnectAsync([uri]$page.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait(8000) | Out-Null
        if ($ws.State -ne [System.Net.WebSockets.WebSocketState]::Open) { try { $ws.Dispose() } catch {}; return $false }
        $script:WS = $ws
        $null = Invoke-CdpJs "1+1" -TimeoutSec 10   # verify the re-attached socket really works
        $script:RepairFailAt = $null
        return $true
    } catch {
        try { if ($script:WS) { $script:WS.Dispose() }; $script:WS = $null } catch {}
        $script:RepairFailAt = Get-Date
        return $false
    } finally {
        $script:InRepair = $false
    }
}

function Send-Cdp {
    param([string]$Method, [hashtable]$Params = @{}, [int]$TimeoutSec = 60)
    # Auto-repair: a dead socket with a live browser (the dominant 10/27 failure mode)
    # gets re-attached transparently; only a genuinely gone browser still throws.
    if (-not $script:WS -or $script:WS.State -ne "Open") {
        if (-not (Repair-CdpConnection)) {
            throw "CDP not connected (call Start-EdgeCdp or Connect-EdgeCdp first)"
        }
    }
    try {
        return Send-CdpCore -Method $Method -Params $Params -TimeoutSec $TimeoutSec
    } catch {
        $msg = "" + $_.Exception.Message
        if ($msg -match "CDP timeout|CDP error in") { throw }   # valid outcomes - never retry
        if ($script:InRetry) { throw }
        $script:InRetry = $true
        try {
            if (Repair-CdpConnection) { return Send-Cdp -Method $Method -Params $Params -TimeoutSec $TimeoutSec }
        } finally { $script:InRetry = $false }
        throw
    }
}

function Send-CdpCore {
    param([string]$Method, [hashtable]$Params = @{}, [int]$TimeoutSec = 60)
    if (-not $script:WS -or $script:WS.State -ne "Open") { throw "CDP not connected (call Start-EdgeCdp or Connect-EdgeCdp first)" }
    $script:MsgId++
    $id = $script:MsgId
    $msg = @{ id = $id; method = $Method; params = $Params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($msg)
    $seg = New-Object "System.ArraySegment[byte]" (,$bytes)
    $ct = New-Object Threading.CancellationTokenSource ([TimeSpan]::FromSeconds($TimeoutSec))
    $script:WS.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct.Token).Wait()
    $ms = New-Object IO.MemoryStream
    $buf = New-Object byte[] 65536
    while ($true) {
        $ms.SetLength(0)
        do {
            $rseg = New-Object "System.ArraySegment[byte]" (,$buf)
            try {
                $res = $script:WS.ReceiveAsync($rseg, $ct.Token)
                $res.Wait()
            } catch {
                # Timeout (token fired) with an OPEN socket = browser alive, page/renderer
                # busy. Reporting that as "connection lost" made bots kill and restart
                # HEALTHY browsers on slow networks (10/27 12:31 run: 11/11 jobs skipped
                # via the restart spiral).
                if ($script:WS -and $script:WS.State -eq "Open" -and $ct.IsCancellationRequested) {
                    throw "CDP timeout during $Method after ${TimeoutSec}s (browser alive - page slow)"
                }
                throw "CDP receive failed during $Method - connection lost"
            }
            $ms.Write($buf, 0, $res.Result.Count)
        } while (-not $res.Result.EndOfMessage)
        $text = [Text.Encoding]::UTF8.GetString($ms.ToArray())
        $data = $null
        try { $data = $text | ConvertFrom-Json } catch { continue }  # skip unparseable events (e.g. duplicate header keys)
        if ($data -and $data.id -eq $id) {
            if ($data.error) { throw "CDP error in $Method : $($data.error.message)" }
            return $data.result
        }
        # else: async event, keep reading
    }
}

function Invoke-CdpJs {
    param([string]$Expression, [int]$TimeoutSec = 30)
    $r = Send-Cdp "Runtime.evaluate" @{ expression = $Expression; returnByValue = $true; awaitPromise = $true } -TimeoutSec $TimeoutSec
    if ($r.exceptionDetails) { return $null }
    return $r.result.value
}

function Navigate-Cdp {
    param([string]$Url, [int]$TimeoutSec = 45)
    try { Send-Cdp "Page.navigate" @{ url = $Url } -TimeoutSec 15 | Out-Null } catch {}
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        # Poll with SHORT timeouts: mid-load the renderer main thread is busy and evaluates
        # are slow - that is NORMAL, not a failure. A poll timing out means "still loading";
        # only a closed socket ("connection lost") aborts navigation.
        $state = $null
        try { $state = Invoke-CdpJs "document.readyState" -TimeoutSec 8 }
        catch {
            # Continue ONLY on a genuine busy-renderer timeout ("CDP timeout"). Every other
            # error means a dead socket ("CDP not connected", "connection lost", ...) -
            # abort navigation immediately so callers can restart/retry fast instead of
            # spinning the whole deadline on a dead connection (that silently burned ~45s
            # per job on the 10/27 runs).
            if ($_.Exception.Message -match "CDP timeout") { continue }
            throw
        }
        if ($state -eq "complete" -or $state -eq "interactive") {
            $body = Get-CdpBodyText
            if ($body -and $body.Length -gt 50) { break }
        }
    }
    Start-Sleep -Seconds 2
}

function Get-CdpBodyText {
    try { return (Invoke-CdpJs "document.body ? document.body.innerText : ''") } catch { return "" }
}

function Get-CdpUrl { return (Invoke-CdpJs "location.href") }

function Save-CdpScreenshot {
    param([string]$Path)
    try {
        $r = Send-Cdp "Page.captureScreenshot" @{ format = "png" }
        [IO.File]::WriteAllBytes($Path, [Convert]::FromBase64String($r.data))
    } catch {}
}

function Set-CdpFileInput {
    param([string]$FilePath)
    # Attaches a file to the first (or best-matching) <input type=file> on the page
    $doc = Send-Cdp "DOM.getDocument" @{ depth = 1 }
    $node = $null
    foreach ($sel in @("input[type='file'][name*='resume' i]","input[type='file'][name*='cv' i]","input[type='file'][id*='resume' i]","input[type='file'][id*='cv' i]","input[type='file']")) {
        try {
            $n = Send-Cdp "DOM.querySelector" @{ nodeId = $doc.root.nodeId; selector = $sel }
            if ($n.nodeId -gt 0) { $node = $n.nodeId; break }
        } catch {}
    }
    if (-not $node) { return $false }
    Send-Cdp "DOM.setFileInputFiles" @{ files = @($FilePath); nodeId = $node } | Out-Null
    return $true
}

Export-ModuleMember -Function Start-EdgeCdp, Connect-EdgeCdp, Switch-EdgeTab, Stop-EdgeCdp, Send-Cdp, Invoke-CdpJs, Navigate-Cdp, Get-CdpBodyText, Get-CdpUrl, Save-CdpScreenshot, Set-CdpFileInput, Enable-Stealth, Enable-AssetBlocking, Repair-CdpConnection, Invoke-CdpRealClick
