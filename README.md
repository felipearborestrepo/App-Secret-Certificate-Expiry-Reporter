# App Secret & Certificate Expiry Reporter

# Define helper Functions
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
# Step 1 — Connect to Graph
- Calls Connect-ToGraph.
```powershell
# ============================================================
# STEP 1 — Connect to Graph
# ============================================================

Connect-ToGraph
```
# Step 2 — Pull all App Registrations
- Calls **Get-MgApplication -All** with properties **Id, DisplayName, AppId, PasswordCredentials, KeyCredentials**, and reports how many apps were found.
```powershell
# ============================================================
# STEP 2 — Pull all App Registrations
# ============================================================

Write-Host "[*] Retrieving all App Registrations..." -ForegroundColor Cyan
$apps = Get-MgApplication -All -Property "Id,DisplayName,AppId,PasswordCredentials,KeyCredentials"
Write-Host "[+] Found $($apps.Count) app registration(s)" -ForegroundColor Green
```
# Step 3 — Loop through every app and check credentials
- For each app:
- Loops through **PasswordCredentials (secrets)**: calculates days remaining, gets urgency, skips if OK, otherwise adds a result (AppDisplayName, AppId, ObjectId, CredentialType = "Secret", CredentialHint, ExpiryDate, DaysRemaining, Urgency).
- Loops through **KeyCredentials (certificates)**: same logic, CredentialType = "Certificate", CredentialHint falling back to KeyId if unnamed.
```powershell
Write-Host "[*] Evaluating secrets and certificates..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$today   = Get-Date

foreach ($app in $apps) {

    # Check Client Secrets
    foreach ($secret in $app.PasswordCredentials) {
        $expiry   = $secret.EndDateTime
        $daysLeft = [int]($expiry - $today).TotalDays
        $urgency  = Get-UrgencyLevel -DaysRemaining $daysLeft

        if ($urgency -eq "OK") { continue }

        $results.Add([PSCustomObject]@{
            AppDisplayName  = $app.DisplayName
            AppId           = $app.AppId
            ObjectId        = $app.Id
            CredentialType  = "Secret"
            CredentialHint  = if ($secret.DisplayName) { $secret.DisplayName } else { "(unnamed)" }
            ExpiryDate      = $expiry.ToString("yyyy-MM-dd")
            DaysRemaining   = $daysLeft
            Urgency         = $urgency
        })
    }

    # Check Certificates
    foreach ($cert in $app.KeyCredentials) {
        $expiry   = $cert.EndDateTime
        $daysLeft = [int]($expiry - $today).TotalDays
        $urgency  = Get-UrgencyLevel -DaysRemaining $daysLeft

        if ($urgency -eq "OK") { continue }

        $results.Add([PSCustomObject]@{
            AppDisplayName  = $app.DisplayName
            AppId           = $app.AppId
            ObjectId        = $app.Id
            CredentialType  = "Certificate"
            CredentialHint  = if ($cert.DisplayName) { $cert.DisplayName } else { $cert.KeyId }
            ExpiryDate      = $expiry.ToString("yyyy-MM-dd")
            DaysRemaining   = $daysLeft
            Urgency         = $urgency
        })
    }
}
```
# Step 4 — Print results to console
- Calls **Write-ConsoleSummary -Results $results**.
```powershell
# ============================================================
# STEP 4 — Print results to console
# ============================================================

Write-ConsoleSummary -Results $results
```
# Step 5 — Export to CSV
- If results exist, sorts by DaysRemaining and exports to **$HOME/Desktop/AppSecretExpiry.csv**. Otherwise prints a **"nothing to export"** message.
```powershell
if ($results.Count -gt 0) {
    $results | Sort-Object DaysRemaining | Export-Csv -Path "$HOME/Desktop/AppSecretExpiry.csv" -NoTypeInformation
    Write-Host "[+] CSV exported to Desktop" -ForegroundColor Green
} else {
    Write-Host "[+] Nothing to export — no credentials expiring within 90 days." -ForegroundColor Green
}
```

# Step 6 — Summary
- Counts **flagged results** by urgency and prints a final tally: total apps scanned, credentials flagged, and the breakdown (expired, critical, warning, notice).
```powershell
# ============================================================
# STEP 6 — Summary
# ============================================================

$expired  = ($results | Where-Object Urgency -eq "EXPIRED").Count
$critical = ($results | Where-Object Urgency -eq "CRITICAL").Count
$warning  = ($results | Where-Object Urgency -eq "WARNING").Count
$notice   = ($results | Where-Object Urgency -eq "NOTICE").Count

Write-Host "`n── Summary ──────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Total apps scanned : $($apps.Count)"  -ForegroundColor White
Write-Host "  Credentials flagged: $($results.Count)" -ForegroundColor White
Write-Host "  Expired            : $expired"  -ForegroundColor Red
Write-Host "  Critical (<=30d)   : $critical" -ForegroundColor Red
Write-Host "  Warning  (<=60d)   : $warning"  -ForegroundColor Yellow
Write-Host "  Notice   (<=90d)   : $notice"   -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────`n" -ForegroundColor DarkGray
```
# Step 8 — Disconnect
- Calls **Disconnect-MgGraph** and confirms disconnection.
```powershell
Disconnect-MgGraph | Out-Null
Write-Host "[*] Disconnected from Graph.`n" -ForegroundColor DarkGray
```
