# Step 1 — Define helper functions
- **Connect-ToGraph** — connects to **Microsoft Graph** using the **Application.Read.All** scope.
- **Get-UrgencyLevel** — takes $DaysRemaining and returns **EXPIRED** (< **0** days), **CRITICAL** (≤ **30**), **WARNING** (≤ **60**), **NOTICE** (≤ **90**), or **OK**.
- **Get-UrgencyColor** — maps each urgency level to a console color (**Red** for **EXPIRED/CRITICAL**, **Yellow** for **WARNING**, **Cyan** for **NOTICE**, **Green** for **default/OK**).
- **Write-ConsoleSummary** — filters results down to anything not OK; if nothing's flagged, prints a clean "all good" message and returns. Otherwise -prints a report header (scan date + flagged count), then loops through EXPIRED → CRITICAL → WARNING → NOTICE, printing each group's items (sorted by days remaining) with app name, credential type, days label, expiry date, and name.

```powershell
# ============================================================
# FUNCTIONS
# ============================================================

function Connect-ToGraph {
    Write-Host "`n[*] Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    Write-Host "[+] Connected to Graph" -ForegroundColor Green
}

function Get-UrgencyLevel {
    param([int]$DaysRemaining)
    if ($DaysRemaining -lt 0)      { return "EXPIRED" }
    elseif ($DaysRemaining -le 30) { return "CRITICAL" }
    elseif ($DaysRemaining -le 60) { return "WARNING" }
    elseif ($DaysRemaining -le 90) { return "NOTICE" }
    else                           { return "OK" }
}

function Get-UrgencyColor {
    param([string]$Urgency)
    switch ($Urgency) {
        "EXPIRED"  { return "Red" }
        "CRITICAL" { return "Red" }
        "WARNING"  { return "Yellow" }
        "NOTICE"   { return "Cyan" }
        default    { return "Green" }
    }
}

function Write-ConsoleSummary {
    param([array]$Results)

    $flagged = $Results | Where-Object { $_.Urgency -ne "OK" }

    if ($flagged.Count -eq 0) {
        Write-Host "`n[OK] No secrets or certificates expiring within 90 days." -ForegroundColor Green
        return
    }

    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host "  APP SECRET & CERTIFICATE EXPIRY REPORT" -ForegroundColor White
    Write-Host "  Scanned: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  Flagged: $($flagged.Count) credential(s)" -ForegroundColor DarkGray
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

    foreach ($urgency in @("EXPIRED", "CRITICAL", "WARNING", "NOTICE")) {
        $group = $flagged | Where-Object { $_.Urgency -eq $urgency }
        if ($group.Count -eq 0) { continue }

        $color = Get-UrgencyColor -Urgency $urgency
        Write-Host "  [$urgency] — $($group.Count) item(s)" -ForegroundColor $color

        foreach ($item in $group | Sort-Object DaysRemaining) {
            $daysLabel = if ($item.DaysRemaining -lt 0) {
                "EXPIRED $([Math]::Abs($item.DaysRemaining))d ago"
            } else {
                "Expires in $($item.DaysRemaining)d"
            }
            Write-Host "    • $($item.AppDisplayName)" -ForegroundColor White
            Write-Host "      Type: $($item.CredentialType) | $daysLabel | Expiry: $($item.ExpiryDate) | Name: $($item.CredentialHint)" -ForegroundColor DarkGray
        }
        Write-Host ""
    }
}
```
# Step 2 — Connect to Graph
- Calls Connect-ToGraph.
