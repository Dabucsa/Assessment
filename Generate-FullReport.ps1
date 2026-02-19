#Requires -Version 5.1
<#
.SYNOPSIS
    Genera reporte HTML completo de M365 Security Assessment
.DESCRIPTION
    Combina report_data.json (REQUERIDO) + security_adoption.json + secure_score.json
.PARAMETER OutputPath
    Carpeta con los archivos generados (default: .\output)
#>
param([string]$OutputPath = ".\output")
$ErrorActionPreference = "Stop"

# ============================================================================
# CARGAR DATOS
# ============================================================================
$LicFile    = Get-ChildItem $OutputPath -Filter "*_report_data.json" -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$AdoptFile  = Get-ChildItem $OutputPath -Filter "*_security_adoption.json" -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$ScoreFile  = Get-ChildItem $OutputPath -Filter "*_secure_score.json" -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$UserCsv    = Get-ChildItem $OutputPath -Filter "*_02_Users.csv" -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $LicFile) { Write-Error "No se encontro report_data.json"; exit 1 }
if (-not $UserCsv) { Write-Error "No se encontro Users CSV"; exit 1 }

Write-Host "[*] Licensing:  $($LicFile.Name)" -ForegroundColor Yellow
if ($AdoptFile) { Write-Host "[*] Adoption:   $($AdoptFile.Name)" -ForegroundColor Yellow }
if ($ScoreFile) { Write-Host "[*] Score:      $($ScoreFile.Name)" -ForegroundColor Yellow }
Write-Host "[*] Users CSV:  $($UserCsv.Name)" -ForegroundColor Yellow

$LicData = Get-Content $LicFile.FullName -Raw | ConvertFrom-Json
$AdoptData = if ($AdoptFile) { Get-Content $AdoptFile.FullName -Raw | ConvertFrom-Json } else { $null }
$ScoreData = if ($ScoreFile) { Get-Content $ScoreFile.FullName -Raw | ConvertFrom-Json } else { $null }
$Users = Import-Csv $UserCsv.FullName

# ============================================================================
# HELPERS
# ============================================================================
function HtmlEncode { param([string]$Text); if (-not $Text) { return "" }; return [System.Net.WebUtility]::HtmlEncode($Text) }
function JEsc([string]$s) { if (-not $s) { return "" }; return $s -replace '\\','\\\\' -replace '"','\"' -replace "`r",'' -replace "`n",' ' -replace "`t",' ' }

$CatShortNames = @{
    "Entra_ID_P1"="Entra P1";"Entra_ID_P2"="Entra P2";"Entra_ID_Governance"="Entra Gov"
    "MDE_P1"="MDE P1";"MDE_P2"="MDE P2";"MDO_P1"="MDO P1";"MDO_P2"="MDO P2"
    "MDA"="MDA";"MDI"="MDI";"Intune_P1"="Intune P1";"Intune_P2"="Intune P2"
    "Purview_AIP_P1"="AIP P1";"Purview_AIP_P2"="AIP P2"
    "Purview_MIP_P1"="MIP P1";"Purview_MIP_P2"="MIP P2"
    "Purview_DLP"="DLP";"Purview_Audit"="Audit";"Purview_eDiscovery"="eDiscov"
    "Purview_InsiderRisk"="Insider";"Purview_CommCompliance"="CommComp";"Purview_DataLifecycle"="DataLife"
    "Copilot_M365"="Copilot"
}

Write-Host "[*] Preparando datos..." -ForegroundColor Yellow

# ============================================================================
# PRE-COMPUTAR: Donut chart + Capacity bars (PowerShell, como original)
# ============================================================================
$ChartColors = @("#5b9bf5","#34d399","#f87171","#fbbf24","#a78bfa","#fb923c","#38bdf8","#f472b6","#c084fc","#2dd4bf")
$PaidSkus = @($LicData.SKUs | Where-Object { $_.Total -gt 0 -and $_.FriendlyName -notmatch '(?i)(free|trial|exploratory)' -and $_.SKU_PartNumber -notmatch '(?i)(free|trial|exploratory)' })
if ($PaidSkus.Count -eq 0) { $PaidSkus = @($LicData.SKUs | Where-Object { $_.Total -gt 0 }) }
$ExcludedSkus = @($LicData.SKUs | Where-Object { $_.Total -gt 0 -and ($PaidSkus -notcontains $_) })
$TotalSeatsAll = ($PaidSkus | Measure-Object -Property Total -Sum).Sum
if (-not $TotalSeatsAll) { $TotalSeatsAll = 1 }
$TotalSeatsFormatted = $TotalSeatsAll.ToString("N0")

$DonutStops = ""
$DonutLegend = ""
$CumulativePct = 0
for ($si = 0; $si -lt $PaidSkus.Count; $si++) {
    $Sku = $PaidSkus[$si]
    $Color = $ChartColors[$si % $ChartColors.Count]
    $Pct = [math]::Round(($Sku.Total / $TotalSeatsAll) * 100, 2)
    $EndPct = $CumulativePct + $Pct
    $DonutStops += "$Color ${CumulativePct}% ${EndPct}%"
    if ($si -lt $PaidSkus.Count - 1) { $DonutStops += ", " }
    $DonutLegend += "<div style='display:flex;align-items:center;gap:8px;margin:4px 0'><span style='width:10px;height:10px;border-radius:50%;background:$Color;flex-shrink:0'></span><span style='font-size:12px'>$(HtmlEncode $Sku.FriendlyName)</span><span style='color:var(--muted);font-size:11px;margin-left:auto;white-space:nowrap'>$($Sku.Total.ToString("N0")) ($([math]::Round($Pct,1))%)</span></div>"
    $CumulativePct = $EndPct
}
if ($ExcludedSkus.Count -gt 0) {
    $ExcNames = ($ExcludedSkus | ForEach-Object { $_.FriendlyName }) -join ", "
    $DonutLegend += "<div style='font-size:10px;color:var(--muted);margin-top:8px;border-top:1px solid var(--border);padding-top:6px'>Excluidos del grafico: $ExcNames</div>"
}

# Capacity bars (horizontal)
$CapacityBars = ""
if ($LicData.Capacity) {
    foreach ($C in $LicData.Capacity) {
        if ($C.TotalSeats -eq 0) { continue }
        $ProdName = if ($CatShortNames.ContainsKey($C.Product)) { $CatShortNames[$C.Product] } else { $C.Product -replace "_"," " }
        $BarPct = [math]::Min($C.PctUsed, 100)
        $BarColor = switch ($C.Status) { "UNDERLICENSED" { "var(--red)" } "OVERLAP" { "var(--cyan,#00bcd4)" } "OVERLICENSED" { "var(--yellow)" } "LOW_USAGE" { "var(--yellow)" } default { "var(--green)" } }
        $CapacityBars += "<div style='display:flex;align-items:center;gap:8px;margin:5px 0'><span style='width:75px;font-size:11px;flex-shrink:0;text-align:right'>$ProdName</span><div style='flex:1;background:var(--bg4);border-radius:4px;height:16px;overflow:hidden'><div style='width:${BarPct}%;height:100%;background:$BarColor;border-radius:4px'></div></div><span style='width:95px;font-size:10px;color:var(--muted);flex-shrink:0'>$($C.UsersActive)/$($C.TotalSeats) ($($C.PctUsed)%)</span></div>"
    }
}

# ============================================================================
# SKU rows
# ============================================================================
$SkuRows = ""
foreach ($Sku in $LicData.SKUs) {
    if ($Sku.Total -eq 0) { continue }
    $RowClass = if ($Sku.Unassigned -lt 0) { "row-alert" } elseif ($Sku.PctUsed -gt 95) { "row-warn" } else { "" }
    $FriendlyCats = ""
    if ($Sku.IncludedCategories) {
        $FriendlyCats = ($Sku.IncludedCategories -split "\s*\|\s*" | ForEach-Object {
            $C = $_.Trim(); if ($CatShortNames.ContainsKey($C)) { $CatShortNames[$C] } else { $C -replace "_"," " }
        }) -join " | "
    }
    $SkuRows += "<tr class='$RowClass'><td><strong>$(HtmlEncode $Sku.FriendlyName)</strong><br><small class='muted'>$(HtmlEncode $Sku.SKU_PartNumber)</small></td><td class='tc'>$($Sku.Total.ToString("N0"))</td><td class='tc'>$($Sku.Assigned.ToString("N0"))</td><td class='tc'>$($Sku.Unassigned.ToString("N0"))</td><td><div class='pg' style='width:100px'><div class='pg-fill $(if($Sku.PctUsed -ge 75){"pg-green"}elseif($Sku.PctUsed -ge 40){"pg-yellow"}else{"pg-red"})' style='width:$([math]::Min($Sku.PctUsed,100))%'></div></div> <small>$($Sku.PctUsed)%</small></td><td class='small muted' style='max-width:200px'>$(HtmlEncode $FriendlyCats)</td></tr>"
}

# ============================================================================
# SKU Matrix (que incluye cada licencia)
# ============================================================================
$MatrixHeaders = ""
$MatrixCats = $LicData.SecurityCategories
foreach ($Cat in $MatrixCats) {
    $Label = if ($CatShortNames.ContainsKey($Cat)) { $CatShortNames[$Cat] } else { $Cat }
    $MatrixHeaders += "<th class='th-prod tc' title='$Cat'>$Label</th>"
}
$MatrixRows = ""
if ($LicData.SkuMatrix) {
    foreach ($Row in $LicData.SkuMatrix) {
        $Cells = ""
        foreach ($Cat in $MatrixCats) {
            $Val = $Row.$Cat
            if ($Val -eq "SI") { $Cells += "<td class='tc'><span class='dot dot-on'></span></td>" }
            else { $Cells += "<td class='tc'><span class='dot dot-na'></span></td>" }
        }
        $MatrixRows += "<tr><td><strong>$(HtmlEncode $Row.SKU)</strong><br><small class='muted'>$(HtmlEncode $Row.PartNumber)</small></td><td class='tc'>$($Row.Total)</td><td class='tc'>$($Row.Assigned)</td>$Cells</tr>"
    }
}

# ============================================================================
# Duplicates JSON (for JS pagination)
# ============================================================================
$DupJsonList = [System.Collections.Generic.List[string]]::new()
if ($LicData.Duplicates) {
    foreach ($Dup in $LicData.Duplicates) {
        $DupJsonList.Add("[`"$(JEsc $Dup.DisplayName)`",`"$(JEsc $Dup.UPN)`",`"$(JEsc $Dup.DuplicateProduct)`",`"$(JEsc $Dup.ProvidedBySKUs)`"]")
    }
}
$DupDataJson = "[" + ($DupJsonList -join ",") + "]"
$DupCount = if ($LicData.Duplicates) { ($LicData.Duplicates | Measure-Object).Count } else { 0 }

# ============================================================================
# Capacity grouped rows (with rowspan)
# ============================================================================
$ProductToGroup = @{}
if ($LicData.CategoryGroups) {
    foreach ($G in $LicData.CategoryGroups.PSObject.Properties) {
        foreach ($Cat in $G.Value) { $ProductToGroup[$Cat] = $G.Name }
    }
}
$CapacityRows = ""
$CapCurrentGroup = ""
if ($LicData.Capacity) {
    foreach ($C in $LicData.Capacity) {
        $ProdShort = if ($CatShortNames.ContainsKey($C.Product)) { $CatShortNames[$C.Product] } else { $C.Product -replace "_"," " }
        $ProdGroup = if ($ProductToGroup.ContainsKey($C.Product)) { $ProductToGroup[$C.Product] } else { "" }
        $GroupCell = ""
        if ($ProdGroup -ne $CapCurrentGroup) {
            $CapCurrentGroup = $ProdGroup
            $GroupCount = @($LicData.Capacity | Where-Object { $Pg = if ($ProductToGroup.ContainsKey($_.Product)) { $ProductToGroup[$_.Product] } else { "" }; $Pg -eq $CapCurrentGroup }).Count
            $GroupCell = "<td class='cat-cell' rowspan='$GroupCount'>$(HtmlEncode $CapCurrentGroup)</td>"
        }
        $OverlapVal = if ($C.PSObject.Properties['OverlapUsers']) { $C.OverlapUsers } else { 0 }
        $StatusBadge = switch ($C.Status) { "UNDERLICENSED" { "<span class='badge b-red'>RIESGO</span>" } "OVERLAP" { "<span class='badge' style='background:#00838f;color:#e0f7fa'>Overlap</span>" } "OVERLICENSED" { "<span class='badge b-yellow'>Sobreasignado</span>" } "LOW_USAGE" { "<span class='badge b-yellow'>Uso bajo</span>" } default { "<span class='badge b-green'>OK</span>" } }
        $RowClass = switch ($C.Status) { "UNDERLICENSED" { "row-alert" } "OVERLICENSED" { "row-warn" } "OVERLAP" { "" } default { "" } }
        $OverlapCell = if ($OverlapVal -gt 0) { "<td class='tc' style='color:#4dd0e1'>$OverlapVal</td>" } else { "<td class='tc muted'>-</td>" }
        $CapacityRows += "<tr class='$RowClass'>$GroupCell<td><strong>$(HtmlEncode $ProdShort)</strong></td><td class='tc'>$($C.TotalSeats.ToString("N0"))</td><td class='tc'><strong>$($C.UsersActive.ToString("N0"))</strong></td><td class='tc'>$($C.UsersDisabled.ToString("N0"))</td>$OverlapCell<td class='tc'>$($C.Unused.ToString("N0"))</td><td><div class='pg' style='width:100px'><div class='pg-fill $(if($C.PctUsed -ge 75){"pg-green"}elseif($C.PctUsed -ge 40){"pg-yellow"}else{"pg-red"})' style='width:$([math]::Min($C.PctUsed,100))%'></div></div> <small>$($C.PctUsed)%</small></td><td class='tc'>$StatusBadge</td><td class='small muted' style='max-width:280px'>$(HtmlEncode $C.ProvidedBy)</td></tr>"
    }
}

# ============================================================================
# Department rows (with ActiveProducts column)
# ============================================================================
$DeptRows = ""
if ($LicData.Departments) {
    foreach ($Dept in $LicData.Departments) {
        $SkuBadges = ""
        if ($Dept.TopSKUs) { foreach ($S in ($Dept.TopSKUs -split "\s*\|\s*")) { $S = $S.Trim(); if ($S) { $SkuBadges += " <span class='badge b-blue'>$(HtmlEncode $S)</span>" } } }
        $ProductBadges = ""
        if ($Dept.ActiveProducts) { foreach ($P in ($Dept.ActiveProducts -split "\s*\|\s*")) { $P = $P.Trim(); if ($P) { $ProductBadges += " <span class='badge b-default'>$(HtmlEncode $P)</span>" } } }
        $DeptRows += "<tr><td><strong>$(HtmlEncode $Dept.Department)</strong></td><td class='tc'>$($Dept.UserCount)</td><td class='small' style='max-width:350px;white-space:normal;line-height:1.8'>$SkuBadges</td><td class='small' style='max-width:500px;white-space:normal;line-height:1.8'>$ProductBadges</td></tr>"
    }
}

# ============================================================================
# Waste JSON (for JS pagination)
# ============================================================================
$WasteJsonList = [System.Collections.Generic.List[string]]::new()
$WasteDetails = $LicData.Waste.Details
if ($WasteDetails) {
    foreach ($W in $WasteDetails) {
        $Ae = if ($W.AccountEnabled -eq "True") { 1 } else { 0 }
        $WasteJsonList.Add("[`"$(JEsc $W.DisplayName)`",`"$(JEsc $W.UPN)`",$Ae,`"$(JEsc $W.LastSignIn)`",`"$(JEsc $W.AssignedSKUs)`",`"$(JEsc $W.WasteReasons)`"]")
    }
}
$WasteDataJson = "[" + ($WasteJsonList -join ",") + "]"
$InactiveDaysJs = if ($LicData.InactiveDays) { $LicData.InactiveDays } else { 90 }

# ============================================================================
# Users compact JSON
# ============================================================================
$SecurityCats = $LicData.SecurityCategories
$CatSepIndices = [System.Collections.Generic.List[int]]::new()
$PrevGroup = ""
for ($ci = 0; $ci -lt $SecurityCats.Count; $ci++) {
    $Cat = $SecurityCats[$ci]
    $CurGroup = ""
    foreach ($G in $LicData.CategoryGroups.PSObject.Properties) {
        if ($G.Value -contains $Cat) { $CurGroup = $G.Name; break }
    }
    if ($CurGroup -ne $PrevGroup -and $PrevGroup -ne "") { $CatSepIndices.Add($ci) }
    $PrevGroup = $CurGroup
}

$UserJsonList = [System.Collections.Generic.List[string]]::new()
foreach ($U in $Users) {
    $CatValues = [System.Collections.Generic.List[int]]::new()
    foreach ($Cat in $SecurityCats) {
        $Val = $U.$Cat
        if ($Val -eq "Enabled") { $CatValues.Add(1) } elseif ($Val -eq "Disabled") { $CatValues.Add(-1) } else { $CatValues.Add(0) }
    }
    $St = 0
    if ($U.AccountEnabled -eq "False") { $St = $St -bor 1 }
    if ($U.IsInactive -eq "True") { $St = $St -bor 2 }
    if ($U.WasteFlags) { $St = $St -bor 4 }
    if ($U.WasteFlags -match 'DisabledPlans') { $St = $St -bor 8 }
    if ($U.DisabledPlans) { $St = $St -bor 16 }
    $Mt = switch ($U.AssignmentMethod) { "Group" { 1 } "Group+Direct" { 2 } default { 0 } }
    $Ca = if ($U.HasConditionalAccess -eq "True") { 1 } else { 0 }
    $Days = if ($U.DaysSinceSignIn -match '^\d+$') { $U.DaysSinceSignIn } else { "-1" }
    $UserJsonList.Add("[`"$(JEsc $U.DisplayName)`",`"$(JEsc $U.UPN)`",`"$(JEsc $U.Department)`",`"$(JEsc $U.AssignedSKUs)`",`"$(JEsc $U.LastSignIn)`",$Days,$St,$Mt,$Ca,[$($CatValues -join ',')],`"$(JEsc $U.WasteFlags)`",`"$(JEsc $U.DisabledPlans)`"]")
}

$CatMetaList = [System.Collections.Generic.List[string]]::new()
$CatFullNames = @{
    "Entra_ID_P1"="Entra ID P1";"Entra_ID_P2"="Entra ID P2";"Entra_ID_Governance"="Entra ID Governance"
    "MDE_P1"="Defender for Endpoint P1";"MDE_P2"="Defender for Endpoint P2"
    "MDO_P1"="Defender for Office P1";"MDO_P2"="Defender for Office P2"
    "MDA"="Defender for Cloud Apps";"MDI"="Defender for Identity"
    "Intune_P1"="Intune Plan 1";"Intune_P2"="Intune Plan 2"
    "Purview_AIP_P1"="AIP P1";"Purview_AIP_P2"="AIP P2";"Purview_MIP_P1"="MIP P1";"Purview_MIP_P2"="MIP P2"
    "Purview_DLP"="Purview DLP";"Purview_Audit"="Purview Audit";"Purview_eDiscovery"="Purview eDiscovery"
    "Purview_InsiderRisk"="Insider Risk";"Purview_CommCompliance"="Comm Compliance";"Purview_DataLifecycle"="Data Lifecycle"
    "Copilot_M365"="M365 Copilot"
}
foreach ($Cat in $SecurityCats) {
    $Full = if ($CatFullNames.ContainsKey($Cat)) { $CatFullNames[$Cat] } else { $Cat }
    $CatMetaList.Add("`"$($Full -replace '"','\"')`"")
}

$UniqueDepts = @($Users | ForEach-Object { $_.Department } | Where-Object { $_ } | Sort-Object -Unique)
$UniqueSkus  = @($Users | ForEach-Object { $_.AssignedSKUs -split "\s*\|\s*" } | Where-Object { $_ } | Sort-Object -Unique)

# ============================================================================
# Consolidar master JSON
# ============================================================================
$MasterData = @{ lic = $LicData; adopt = $AdoptData; score = $ScoreData }
$MasterJson = $MasterData | ConvertTo-Json -Depth 10 -Compress

# ============================================================================
# Cards resumen
# ============================================================================
$TotalUsers = $LicData.TotalLicensedUsers
$TotalSkus  = ($LicData.SKUs | Measure-Object).Count
$E5Assigned  = ($LicData.SKUs | Where-Object { $_.FriendlyName -like "*E5*" } | Measure-Object -Property Assigned -Sum).Sum
$E3Assigned  = ($LicData.SKUs | Where-Object { $_.FriendlyName -like "*E3*" } | Measure-Object -Property Assigned -Sum).Sum
if (-not $E5Assigned) { $E5Assigned = 0 }
if (-not $E3Assigned) { $E3Assigned = 0 }
$WasteTotal     = $LicData.Waste.TotalWasteUsers
$WasteDisabled  = $LicData.Waste.DisabledAccounts
$WasteInactive  = $LicData.Waste.InactiveUsers
$WasteDuplicate = $LicData.Waste.DuplicateLicenses
$WasteDisPlans  = if ($LicData.Waste.DisabledPlans) { $LicData.Waste.DisabledPlans } else {
    # Count from Users CSV if available
    if ($UsersCsv) { ($UsersCsv | Where-Object { $_.DisabledPlans -ne '' -and $null -ne $_.DisabledPlans }).Count } else { 0 }
}
$MethodGroup  = $LicData.AssignmentMethods.Group
$MethodDirect = $LicData.AssignmentMethods.Direct
$MethodMixed  = $LicData.AssignmentMethods.Mixed

# ============================================================================
# Score data for header gauge
# ============================================================================
$ScorePct = 0; $ScoreCur = 0; $ScoreMax = 0
if ($ScoreData -and $ScoreData.Score) {
    $ScorePct = $ScoreData.Score.Pct
    $ScoreCur = $ScoreData.Score.Current
    $ScoreMax = $ScoreData.Score.Max
}
$GaugeColor = if ($ScorePct -ge 70) { "var(--green)" } elseif ($ScorePct -ge 40) { "var(--yellow)" } else { "var(--red)" }
$GaugeCircum = 314.16
$GaugeOffset = [math]::Round($GaugeCircum * (1 - $ScorePct / 100), 2)

# Pre-compute SVG ring gauge for dashboard hero
$RingRadius = 52
$RingCircum = [math]::Round(2 * [math]::PI * $RingRadius, 2)
$RingOffset = [math]::Round($RingCircum * (1 - $ScorePct / 100), 2)
$ScoreRingSvg = @"
<svg viewBox='0 0 130 130' width='150' height='150'>
  <circle cx='65' cy='65' r='$RingRadius' fill='none' stroke='var(--bg4)' stroke-width='9'/>
  <circle cx='65' cy='65' r='$RingRadius' fill='none' stroke='$GaugeColor' stroke-width='9' stroke-dasharray='$RingCircum' stroke-dashoffset='$RingOffset' stroke-linecap='round' transform='rotate(-90 65 65)' style='transition:stroke-dashoffset 1s ease'/>
  <text x='65' y='58' text-anchor='middle' fill='$GaugeColor' font-size='26' font-weight='700' font-family='Segoe UI,system-ui,sans-serif'>${ScorePct}%</text>
  <text x='65' y='76' text-anchor='middle' fill='var(--muted)' font-size='8' font-family='Segoe UI,system-ui,sans-serif' text-transform='uppercase' letter-spacing='.5'>SECURE SCORE</text>
  <text x='65' y='92' text-anchor='middle' fill='var(--muted)' font-size='9' font-family='Segoe UI,system-ui,sans-serif'>$ScoreCur / $ScoreMax</text>
</svg>
"@

# Pre-compute category bars for hero section  
$CatBarsHtml = ""
if ($ScoreData -and $ScoreData.Categories) {
    foreach ($Cat in $ScoreData.Categories) {
        $CatPct = if ($Cat.MaxScore -and $Cat.MaxScore -gt 0) { [math]::Min([math]::Round($Cat.Score / $Cat.MaxScore * 100), 100) } elseif ($Cat.PctScore) { [math]::Min($Cat.PctScore, 100) } else { 0 }
        $CatBarColor = if ($CatPct -ge 70) { "var(--green)" } elseif ($CatPct -ge 40) { "var(--yellow)" } else { "var(--red)" }
        $CatLabel = if ($Cat.MaxScore) { "$($Cat.Score)/$($Cat.MaxScore) ($CatPct%)" } else { "N/A" }
        $CatBarsHtml += "<div class='bar-row'><span class='bar-label'>$($Cat.Category)</span><div class='bar-track'><div class='bar-fill' style='width:${CatPct}%;background:$CatBarColor'></div></div><span class='bar-value'>$CatLabel</span></div>"
    }
}

# Pre-compute assignment bar percentages
$AssignTotal = [math]::Max(([int]$MethodGroup + [int]$MethodDirect + [int]$MethodMixed), 1)
$PctGroup = [math]::Round([int]$MethodGroup / $AssignTotal * 100, 1)
$PctDirect = [math]::Round([int]$MethodDirect / $AssignTotal * 100, 1)
$PctMixed = [math]::Round([int]$MethodMixed / $AssignTotal * 100, 1)

# MFA % from adoption data (directory-wide from aggregate API)
$MfaPct = "N/A"
if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.MFA) {
    $MfaPct = "$($AdoptData.Entra.MFA.PctRegistered)%"
}

$RiskyCount = 0
if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.RiskyUsers) {
    $RiskyCount = $AdoptData.Entra.RiskyUsers.TotalAtRisk
}

# ============================================================================
# GENERAR HTML
# ============================================================================
$Timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$HtmlPath = Join-Path $OutputPath "${Timestamp}_Full_Security_Report.html"
$TenantName = $LicData.TenantName
$TenantDomain = $LicData.TenantDomain
$GeneratedAt = $LicData.GeneratedAt

# Build tab navigation with badges and alert indicators
$PendingRecs = 0
if ($ScoreData -and $ScoreData.AllRecommendations) {
    $PendingRecs = @($ScoreData.AllRecommendations | Where-Object { $_.ImplementationStatus -ne 'Implemented' }).Count
}
$DeptCount = if ($LicData.Departments) { ($LicData.Departments | Measure-Object).Count } else { 0 }
$CapRiskCount = if ($LicData.Capacity) { @($LicData.Capacity | Where-Object { $_.Status -eq 'UNDERLICENSED' }).Count } else { 0 }

$NavItems = [System.Collections.Generic.List[string]]::new()
$NavItems.Add("<button class='main-tab active' data-tab='resumen'>&#128202; Resumen</button>")
if ($AdoptData) {
    $MfaAlertDot = if ($AdoptData.Entra -and $AdoptData.Entra.MFA -and $AdoptData.Entra.MFA.PctRegistered -lt 50) { "<span class='t-alert' style='background:var(--red)'></span>" } else { "" }
    $NavItems.Add("<button class='main-tab' data-tab='postura'>&#128737; Postura $MfaAlertDot</button>")
}
if ($ScoreData -and $ScoreData.AllRecommendations) {
    $RecAlertDot = if ($PendingRecs -gt 0) { "<span class='t-alert' style='background:var(--yellow)'></span>" } else { "" }
    $NavItems.Add("<button class='main-tab' data-tab='recomendaciones'>&#128161; Recomendaciones <span class='t-badge'>$PendingRecs</span>$RecAlertDot</button>")
}
$CapAlertDot = if ($CapRiskCount -gt 0) { "<span class='t-alert' style='background:var(--red)'></span>" } else { "" }
$NavItems.Add("<button class='main-tab' data-tab='licenciamiento'>&#128179; Licenciamiento <span class='t-badge'>$TotalSkus</span>$CapAlertDot</button>")
$NavItems.Add("<button class='main-tab' data-tab='usuarios'>&#128100; Usuarios <span class='t-badge'>$($Users.Count)</span></button>")
$WasteAlertDot = if ([int]$WasteTotal -gt 0) { "<span class='t-alert' style='background:var(--red)'></span>" } else { "" }
$NavItems.Add("<button class='main-tab' data-tab='optimizacion'>&#9888; Optimizacion <span class='t-badge'>$([int]$WasteTotal)</span>$WasteAlertDot</button>")

$Nav = $NavItems -join "`n  "

$Html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 Security Assessment - $TenantName</title>
<style>
:root{--bg:#0d1017;--bg2:#151921;--bg3:#1e2433;--bg4:#283044;--accent:#5b9bf5;--accent2:#8b6cc1;--green:#34d399;--yellow:#fbbf24;--red:#f87171;--blue:#60a5fa;--text:#e2e8f0;--muted:#6b7a90;--border:#2a3346}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;line-height:1.5}
.hdr{background:linear-gradient(135deg,var(--bg2),var(--bg3));border-bottom:1px solid var(--border);padding:28px 36px;display:flex;align-items:center;gap:32px;flex-wrap:wrap}
.hdr-info{flex:1;min-width:280px}
.hdr h1{font-size:22px;font-weight:700;color:var(--accent)}
.hdr .meta{color:var(--muted);font-size:12px;margin-top:6px}
.hdr .tenant{display:inline-block;background:var(--bg4);border:1px solid var(--border);border-radius:16px;padding:3px 14px;font-size:11px;margin-top:8px;color:var(--muted)}
.hdr-gauge{position:relative;width:120px;height:120px;border-radius:50%;flex-shrink:0}
.hdr-gauge-inner{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:82px;height:82px;border-radius:50%;background:var(--bg2);display:flex;flex-direction:column;align-items:center;justify-content:center}
.hdr-gauge-val{font-size:24px;font-weight:700}.hdr-gauge-lbl{font-size:8px;color:var(--muted);text-transform:uppercase;letter-spacing:.3px}
.wrap{max-width:1500px;margin:0 auto;padding:24px 36px}
.sec{font-size:16px;font-weight:600;margin:28px 0 14px;padding-bottom:6px;border-bottom:2px solid var(--border);display:flex;align-items:center;gap:8px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:12px;margin-bottom:24px}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:16px;text-align:center}
.card .v{font-size:28px;font-weight:700;color:var(--accent)}
.card .l{color:var(--muted);font-size:10px;margin-top:4px;text-transform:uppercase;letter-spacing:.5px}
.card.g .v{color:var(--green)}.card.y .v{color:var(--yellow)}.card.r .v{color:var(--red)}.card.p .v{color:var(--accent2)}.card.bl .v{color:var(--blue)}
.tw{overflow-x:auto;border-radius:8px;border:1px solid var(--border);margin-bottom:24px}
table{width:100%;border-collapse:collapse}
thead{background:var(--bg3)}
th{padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover{background:var(--bg3)}
.tc{text-align:center}.muted{color:var(--muted)}.small{font-size:11px}
.sku-cell{max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--muted)}
.row-alert td{background:rgba(248,113,113,.06)}.row-warn td{background:rgba(251,191,36,.04)}
.row-disabled td{opacity:.35}.row-inactive td{opacity:.55}
.badge{display:inline-block;padding:1px 7px;border-radius:8px;font-size:10px;font-weight:600;white-space:nowrap}
.b-green{background:rgba(52,211,153,.12);color:var(--green)}.b-yellow{background:rgba(251,191,36,.12);color:var(--yellow)}
.b-red{background:rgba(248,113,113,.12);color:var(--red)}.b-blue{background:rgba(96,165,250,.12);color:var(--blue)}
.b-default{background:var(--bg4);color:var(--muted)}
.pg{background:var(--bg4);border-radius:3px;height:7px;overflow:hidden;display:inline-block;vertical-align:middle;margin-right:4px}
.pg-fill{height:100%;border-radius:3px}
.pg-green{background:var(--green)}.pg-yellow{background:var(--yellow)}.pg-red{background:var(--red)}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%}
.dot-on{background:var(--green)}.dot-off{background:var(--red)}.dot-na{background:var(--bg4);border:1px solid var(--border)}
.td-sep{border-left:2px solid var(--border)}
.cat-cell{background:var(--bg3)!important;font-weight:600;font-size:11px;color:var(--accent);white-space:nowrap;vertical-align:middle;border-right:2px solid var(--border);padding:10px 14px}
.sticky-name{position:sticky;left:0;background:var(--bg2);z-index:1;min-width:160px}
tr:hover .sticky-name{background:var(--bg3)}
.search{width:100%;padding:9px 14px;background:var(--bg2);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:13px;margin-bottom:12px}
.search:focus{outline:none;border-color:var(--accent)}
.tabs{display:flex;gap:4px;margin-bottom:16px;flex-wrap:wrap}
.tab{padding:6px 14px;border-radius:6px;font-size:12px;cursor:pointer;background:var(--bg2);border:1px solid var(--border);color:var(--muted);transition:.2s}
.tab:hover,.tab.active{background:var(--bg3);color:var(--accent);border-color:var(--accent)}
.action{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:14px 18px;margin-bottom:10px}
.action .a-title{font-weight:600;font-size:13px;margin-bottom:4px}
.action .a-desc{color:var(--muted);font-size:12px}
.action.a-red{border-left:3px solid var(--red)}.action.a-yellow{border-left:3px solid var(--yellow)}.action.a-blue{border-left:3px solid var(--blue)}
.legend{display:flex;gap:16px;flex-wrap:wrap;margin:8px 0 16px;font-size:11px;color:var(--muted)}
.legend span{display:flex;align-items:center;gap:4px}
.footer{text-align:center;padding:24px;color:var(--muted);font-size:11px;border-top:1px solid var(--border);margin-top:36px}
nav{position:sticky;top:0;z-index:10;background:var(--bg);border-bottom:1px solid var(--border);padding:10px 36px;display:flex;gap:6px;flex-wrap:wrap}
.main-tab{position:relative;padding:8px 16px;font-size:12px;font-weight:500;cursor:pointer;border-radius:6px;background:var(--bg2);border:1px solid var(--border);color:var(--muted);transition:all .2s;white-space:nowrap;display:inline-flex;align-items:center;gap:6px}
.main-tab:hover{background:var(--bg3);color:var(--accent);border-color:var(--accent)}
.main-tab.active{background:var(--bg3);color:var(--accent);border-color:var(--accent);font-weight:600}
.main-tab .t-badge{font-size:9px;padding:1px 6px;border-radius:8px;background:var(--bg4);color:var(--muted);font-weight:600}
.main-tab.active .t-badge{background:rgba(91,155,245,.15);color:var(--accent)}
.main-tab .t-alert{width:7px;height:7px;border-radius:50%;position:absolute;top:4px;right:4px}
.tab-panel{display:none;animation:tabFadeIn .3s ease}.tab-panel.active{display:block}
@keyframes tabFadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.csv-btn{background:var(--bg2);border:1px solid var(--border);color:var(--muted);border-radius:4px;padding:4px 12px;font-size:11px;cursor:pointer;display:inline-flex;align-items:center;gap:4px;float:right}
.csv-btn:hover{background:var(--bg3);color:var(--accent);border-color:var(--accent)}
.gauge-svg{width:120px;height:120px;flex-shrink:0}
.dash-hero{display:flex;gap:32px;align-items:center;background:linear-gradient(135deg,var(--bg2),var(--bg3));border:1px solid var(--border);border-radius:14px;padding:32px;margin-bottom:24px;position:relative;overflow:hidden}
.dash-hero::before{content:'';position:absolute;top:-60px;right:-60px;width:200px;height:200px;border-radius:50%;background:rgba(91,155,245,.04);pointer-events:none}
.dash-ring{flex-shrink:0;text-align:center}
.dash-ring svg{filter:drop-shadow(0 0 16px rgba(91,155,245,.12))}
.dash-cats{flex:1;min-width:220px}
.dash-cats-title{font-size:11px;font-weight:600;color:var(--accent);margin-bottom:14px;text-transform:uppercase;letter-spacing:.6px}
.dash-divider{width:1px;background:var(--border);align-self:stretch;margin:0 8px;opacity:.5}
.dash-controls{min-width:180px}
.dash-controls-title{font-size:11px;font-weight:600;color:var(--accent);margin-bottom:14px;text-transform:uppercase;letter-spacing:.6px}
.dash-ctrl-row{display:flex;justify-content:space-between;align-items:center;padding:5px 0}
.dash-ctrl-row .dcr-v{font-size:20px;font-weight:700}
.dash-ctrl-row .dcr-l{font-size:10px;color:var(--muted);text-transform:uppercase}
.dash-cols{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:24px}
.dash-panel2{margin-bottom:24px}
@media(max-width:1000px){.dash-cols{grid-template-columns:1fr}.dash-hero{flex-direction:column;align-items:stretch;text-align:center}}
.dash-panel{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:20px 24px;border-top:3px solid var(--border);transition:transform .15s}
.dash-panel:hover{transform:translateY(-2px)}
.dash-panel.dp-blue{border-top-color:var(--accent)}
.dash-panel.dp-teal{border-top-color:var(--green)}
.dash-panel.dp-amber{border-top-color:var(--yellow)}
.dash-panel.dp-red{border-top-color:var(--red)}
.dp-title{font-size:13px;font-weight:600;margin-bottom:16px;display:flex;align-items:center;gap:8px;color:var(--text)}
.dp-row{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid rgba(42,51,70,.4)}
.dp-row:last-child{border-bottom:none}
.dp-label{font-size:12px;color:var(--muted)}
.dp-val{font-size:18px;font-weight:700}
.dp-val.v-green{color:var(--green)}.dp-val.v-red{color:var(--red)}.dp-val.v-yellow{color:var(--yellow)}.dp-val.v-blue{color:var(--blue)}.dp-val.v-accent{color:var(--accent)}.dp-val.v-muted{color:var(--muted);font-size:15px}
.dp-bar{height:8px;background:var(--bg4);border-radius:4px;margin-top:12px;overflow:hidden;display:flex}
.dp-bar span{height:100%;transition:width .5s}
.dp-bar-legend{display:flex;gap:14px;margin-top:8px;font-size:10px;color:var(--muted)}
.dp-bar-legend i{width:8px;height:8px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}
.dp-sep{border-top:1px solid var(--border);margin:12px 0;opacity:.4}
.dp-big{font-size:36px;font-weight:700;color:var(--accent);line-height:1}
.dp-biglabel{font-size:10px;color:var(--muted);text-transform:uppercase;margin-top:4px}
.priority-list{display:flex;flex-direction:column;gap:8px;margin-bottom:24px}
.priority-item{display:flex;align-items:flex-start;gap:14px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:14px 18px;transition:border-color .2s,transform .15s}
.priority-item:hover{border-color:var(--accent);transform:translateX(4px)}
.priority-num{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0;color:#fff}
.pn-red{background:var(--red)}.pn-yellow{background:var(--yellow);color:var(--bg)}.pn-blue{background:var(--blue)}
.priority-body{flex:1}.priority-body .pb-title{font-weight:600;font-size:13px;margin-bottom:3px}.priority-body .pb-desc{color:var(--muted);font-size:11px;line-height:1.5}
.narrative-box{background:linear-gradient(135deg,var(--bg2),var(--bg3));border:1px solid var(--border);border-radius:12px;padding:20px 24px;font-size:13px;line-height:1.8;color:var(--text);border-left:3px solid var(--accent);margin-bottom:24px}
.pg-btn{background:var(--bg2);border:1px solid var(--border);color:var(--muted);border-radius:4px;padding:3px 9px;font-size:11px;cursor:pointer;min-width:28px;text-align:center}
.pg-btn:hover{background:var(--bg3);color:var(--accent);border-color:var(--accent)}
.pg-active{background:var(--accent)!important;color:var(--bg)!important;border-color:var(--accent)!important;font-weight:700}
.pg-dis{opacity:.3;pointer-events:none}
.pg-controls{min-height:30px}
.sc{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:20px;margin-bottom:16px}
.sc-title{font-size:14px;font-weight:600;margin-bottom:12px;color:var(--accent);display:flex;align-items:center;gap:8px}
.sc-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:12px}
.sc-stat{text-align:center;padding:8px}.sc-stat .v{font-size:22px;font-weight:700}.sc-stat .l{font-size:10px;color:var(--muted);text-transform:uppercase}
.bar-row{display:flex;align-items:center;gap:8px;margin:5px 0}
.bar-label{width:75px;font-size:11px;flex-shrink:0;text-align:right}
.bar-track{flex:1;background:var(--bg4);border-radius:4px;height:16px;overflow:hidden}
.bar-fill{height:100%;border-radius:4px}
.bar-value{width:95px;font-size:10px;color:var(--muted);flex-shrink:0}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:24px}
@media(max-width:900px){.two-col{grid-template-columns:1fr}.hdr{flex-direction:column;align-items:flex-start}}
@media print{
  .search,nav,.pg-controls,.tabs,.no-print,.csv-btn{display:none!important}
  .tab-panel{display:block!important;animation:none!important}
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}
  body{font-size:10px}.wrap{max-width:100%;padding:10px}
  .card,.sc{break-inside:avoid}.action{break-inside:avoid}
  .sec{break-before:auto;break-after:avoid}
  table{font-size:9px}td{padding:4px 6px}.tw{overflow:visible}
  .sticky-name{position:static!important}
  @page{size:landscape;margin:8mm}
}
</style>
</head>
<body>

<!-- ====== HEADER ====== -->
<div class="hdr">
  <div class="hdr-info">
    <h1>Microsoft 365 Security Assessment</h1>
    <div class="meta">Generado: $GeneratedAt</div>
    <div class="tenant">$TenantName | $TenantDomain | $($LicData.TenantId)</div>
    <button class="no-print" onclick="window.print()" style="margin-top:10px;padding:8px 20px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">&#128196; Exportar PDF</button>
  </div>
</div>

<nav>
  $Nav
</nav>

<div class="wrap">

<div class="tab-panel active" id="tab-resumen">
<!-- ====== DASHBOARD EJECUTIVO ====== -->
<div class="sec" id="dashboard">Dashboard Ejecutivo</div>

<!-- Hero: Secure Score + Categories + Controls -->
$(if ($ScoreData -and $ScoreData.Categories) {
@"
<div class="dash-hero">
  <div class="dash-ring">
    $ScoreRingSvg
  </div>
  <div class="dash-divider"></div>
  <div class="dash-cats">
    <div class="dash-cats-title">Score por Categoria</div>
    $CatBarsHtml
  </div>
  <div class="dash-divider"></div>
  <div class="dash-controls">
    <div class="dash-controls-title">Secure Score &mdash; Controles</div>
    <div class="dash-ctrl-row"><div><span class="dcr-v" style="color:var(--green)">$($ScoreData.Summary.Implemented)</span></div><div class="dcr-l">Implementados</div></div>
    <div class="dash-ctrl-row"><div><span class="dcr-v" style="color:var(--yellow)">$($ScoreData.Summary.Partial)</span></div><div class="dcr-l">Parciales</div></div>
    <div class="dash-ctrl-row"><div><span class="dcr-v" style="color:var(--red)">$($ScoreData.Summary.NotImplemented)</span></div><div class="dcr-l">No Implementados</div></div>
    <div style="border-top:1px solid var(--border);margin-top:8px;padding-top:8px"><div class="dash-ctrl-row"><div><span class="dcr-v" style="color:var(--text)">$($ScoreData.Summary.TotalControls)</span></div><div class="dcr-l">Total Controles</div></div></div>
  </div>
</div>
"@
})

<!-- Three KPI Panels -->
<div class="dash-cols">
  <!-- Panel 1: Licenciamiento -->
  <div class="dash-panel dp-blue">
    <div class="dp-title">&#128179; Licenciamiento</div>
    <div class="dp-row"><span class="dp-label">Usuarios Licenciados</span><span class="dp-val v-accent">$([int]$TotalUsers |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-row"><span class="dp-label">SKUs Activos</span><span class="dp-val">$TotalSkus</span></div>
    <div class="dp-row"><span class="dp-label">Licencias E5</span><span class="dp-val v-green">$([int]$E5Assigned |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-row"><span class="dp-label">Licencias E3</span><span class="dp-val v-muted">$([int]$E3Assigned |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-sep"></div>
    <div style="font-size:11px;color:var(--muted);margin-bottom:6px">Metodo de Asignacion</div>
    <div class="dp-bar">
      <span style="width:${PctGroup}%;background:var(--green)" title="Grupo: $([int]$MethodGroup)"></span>
      <span style="width:${PctDirect}%;background:var(--yellow)" title="Directa: $([int]$MethodDirect)"></span>
      <span style="width:${PctMixed}%;background:var(--accent)" title="Mixta: $([int]$MethodMixed)"></span>
    </div>
    <div class="dp-bar-legend">
      <span><i style="background:var(--green)"></i>Grupo ($([int]$MethodGroup))</span>
      <span><i style="background:var(--yellow)"></i>Directa ($([int]$MethodDirect))</span>
      <span><i style="background:var(--accent)"></i>Mixta ($([int]$MethodMixed))</span>
    </div>
  </div>

  <!-- Panel 2: Acceso y Seguridad -->
  <div class="dash-panel $(if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.MFA -and $AdoptData.Entra.MFA.PctRegistered -lt 50) { 'dp-red' } else { 'dp-teal' })">
    <div class="dp-title">&#128737;&#65039; Acceso y Seguridad</div>
    $(if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.MFA) {
      $Mp = $AdoptData.Entra.MFA.PctRegistered
      $MfaVClass = if ($Mp -ge 70) { 'v-green' } else { 'v-red' }
      "<div class='dp-row'><span class='dp-label'>MFA Registrado</span><span class='dp-val $MfaVClass'>${Mp}%</span></div>"
    })
    $(if ($RiskyCount -gt 0) {
      "<div class='dp-row'><span class='dp-label'>Usuarios de Alto Riesgo</span><span class='dp-val v-red'>$RiskyCount</span></div>"
    } else {
      "<div class='dp-row'><span class='dp-label'>Usuarios de Alto Riesgo</span><span class='dp-val v-green'>0</span></div>"
    })
    $(if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.ConditionalAccess) {
      "<div class='dp-row'><span class='dp-label'>Politicas CA Activas</span><span class='dp-val'>$($AdoptData.Entra.ConditionalAccess.Enabled)</span></div>"
    })
    $(if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.PIM) {
      $pimRoles = $AdoptData.Entra.PIM.ActiveRoles
      $pimUsers = $AdoptData.Entra.PIM.ActiveUsers
      $pimPerm  = $AdoptData.Entra.PIM.PermanentUsers
      $pimElig  = $AdoptData.Entra.PIM.EligibleUsers
      "<div class='dp-sep'></div>"
      "<div class='dp-row'><span class='dp-label'>Roles Privilegiados</span><span class='dp-val'>$pimRoles</span></div>"
      "<div class='dp-row'><span class='dp-label'>Asignaciones Permanentes</span><span class='dp-val $(if($pimPerm -gt 5){'v-yellow'}else{'v-green'})'>$pimPerm</span></div>"
    })
  </div>

  <!-- Panel 3: Optimizacion -->
  <div class="dash-panel $(if ([int]$WasteTotal -gt 0) { 'dp-amber' } else { 'dp-teal' })">
    <div class="dp-title">&#9888;&#65039; Optimizacion de Licencias</div>
    <div style="text-align:center;padding:8px 0">
      <div class="dp-big" style="color:$(if ([int]$WasteTotal -gt 0) { 'var(--yellow)' } else { 'var(--green)' })">$([int]$WasteTotal |ForEach-Object{$_.ToString("N0")})</div>
      <div class="dp-biglabel">Licencias a Revisar</div>
    </div>
    <div class="dp-sep"></div>
    <div class="dp-row"><span class="dp-label">Cuentas Deshabilitadas</span><span class="dp-val $(if([int]$WasteDisabled -gt 0){'v-red'}else{'v-green'})">$([int]$WasteDisabled |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-row"><span class="dp-label">Inactivos $($LicData.InactiveDays)d+</span><span class="dp-val $(if([int]$WasteInactive -gt 0){'v-yellow'}else{'v-green'})">$([int]$WasteInactive |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-row"><span class="dp-label">Licencias Duplicadas</span><span class="dp-val $(if([int]$WasteDuplicate -gt 0){'v-blue'}else{'v-green'})">$([int]$WasteDuplicate |ForEach-Object{$_.ToString("N0")})</span></div>
    <div class="dp-row"><span class="dp-label">Planes Deshabilitados</span><span class="dp-val $(if([int]$WasteDisPlans -gt 0){'v-yellow'}else{'v-green'})">$([int]$WasteDisPlans |ForEach-Object{$_.ToString("N0")})</span></div>
  </div>
</div>

<!-- Defender Suite Panel -->
$(
$DefenderPanelHtml = ""
if ($AdoptData) {
    $DefProducts = [System.Collections.Generic.List[string]]::new()
    $DefCtrlOk = 0; $DefCtrlPend = 0; $DefCtrlTot = 0
    $DefIcon = @{ MDO="&#128231;"; MDA="&#9729;"; MDI="&#128737;"; MDE="&#128737;" }
    $DefName = @{ MDO="Office 365"; MDA="Cloud Apps"; MDI="Identity"; MDE="Endpoint" }
    foreach ($dk in @("MDO","MDA","MDI","MDE")) {
        $dp = $AdoptData.$dk
        if ($dp) {
            $sc = $dp.SecureScoreControls
            if ($sc) {
                $dpOk   = [int]$sc.FullyEnabled
                $dpTot  = [int]$sc.Total
                $dpPend = $dpTot - $dpOk
                $dpPct  = if ($dpTot -gt 0) { [math]::Round($dpOk / $dpTot * 100) } else { 0 }
                $dpCol  = if ($dpPct -ge 70) { 'var(--green)' } elseif ($dpPct -ge 40) { 'var(--yellow)' } else { 'var(--red)' }
                $DefProducts.Add("<div style='display:flex;align-items:center;gap:8px;margin:4px 0'><span style='width:90px;font-size:11px;flex-shrink:0'>$($DefIcon[$dk]) $($DefName[$dk])</span><div style='flex:1;background:var(--bg4);border-radius:3px;height:10px;overflow:hidden'><div style='width:${dpPct}%;height:100%;background:${dpCol};border-radius:3px'></div></div><span style='width:55px;font-size:10px;color:var(--muted);text-align:right;flex-shrink:0'>$dpOk/$dpTot</span></div>")
                $DefCtrlOk += $dpOk
                $DefCtrlPend += $dpPend
                $DefCtrlTot += $dpTot
            } else {
                $DefProducts.Add("<div style='display:flex;align-items:center;gap:8px;margin:4px 0'><span style='width:90px;font-size:11px;flex-shrink:0'>$($DefIcon[$dk]) $($DefName[$dk])</span><span style='font-size:10px;color:var(--muted)'>Licenciado</span></div>")
            }
        }
    }
    if ($DefProducts.Count -gt 0) {
        $DefPctAll = if ($DefCtrlTot -gt 0) { [math]::Round($DefCtrlOk / $DefCtrlTot * 100) } else { 0 }
        $DefBorderCol = if ($DefPctAll -ge 70) { 'dp-teal' } elseif ($DefPctAll -ge 40) { 'dp-amber' } else { 'dp-red' }
        $DefenderPanelHtml = @"
<div class="dash-panel2">
  <div class="dash-panel $DefBorderCol" style="margin:0">
    <div class="dp-title">&#128737;&#65039; Microsoft Defender</div>
    <div style="text-align:center;padding:6px 0">
      <div class="dp-big" style="color:$(if($DefPctAll -ge 70){'var(--green)'}elseif($DefPctAll -ge 40){'var(--yellow)'}else{'var(--red)'})">$DefPctAll%</div>
      <div class="dp-biglabel">Controles Defender ($DefCtrlOk de $DefCtrlTot)</div>
    </div>
    <div class="dp-sep"></div>
    <div class="dp-row"><span class="dp-label">Implementados</span><span class="dp-val v-green">$DefCtrlOk</span></div>
    <div class="dp-row"><span class="dp-label">Pendientes</span><span class="dp-val v-red">$DefCtrlPend</span></div>
    <div class="dp-sep"></div>
    $($DefProducts -join "`n    ")
  </div>
</div>
"@
    }
}
$DefenderPanelHtml
)

<!-- ====== ACCIONES PRIORITARIAS ====== -->
$(
$ActionList = [System.Collections.Generic.List[object]]::new()
if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.MFA -and $AdoptData.Entra.MFA.PctRegistered -lt 50) {
    $ActionList.Add(@{ Color='red'; Title="Implementar MFA para todos los usuarios"; Desc="Solo $($AdoptData.Entra.MFA.PctRegistered)% de los usuarios del directorio tienen MFA registrado. Esto es un riesgo critico de seguridad." })
}
if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.RiskyUsers -and $AdoptData.Entra.RiskyUsers.High -gt 0) {
    $ActionList.Add(@{ Color='red'; Title="Investigar $($AdoptData.Entra.RiskyUsers.High) usuarios de alto riesgo"; Desc="Identity Protection ha detectado usuarios de alto riesgo. Requieren investigacion inmediata y remediacion." })
}
if ($WasteDisabled -gt 0) {
    $ActionList.Add(@{ Color='red'; Title="Remover licencias de $WasteDisabled cuentas deshabilitadas"; Desc="Cuentas deshabilitadas consumiendo licencias. Reasignar para optimizar costos." })
}
if ($WasteInactive -gt 0) {
    $ActionList.Add(@{ Color='yellow'; Title="Revisar $WasteInactive usuarios sin sign-in en $($LicData.InactiveDays)+ dias"; Desc="Usuarios activos sin conexion reciente. Verificar si requieren licencia o se pueden reasignar." })
}
if ($MethodDirect -gt ($TotalUsers * 0.5)) {
    $ActionList.Add(@{ Color='yellow'; Title="Migrar asignacion directa a grupos"; Desc="$MethodDirect usuarios con licencias directas. Group-based licensing es mas escalable y reduce errores." })
}
if ($WasteDuplicate -gt 0) {
    $ActionList.Add(@{ Color='blue'; Title="Consolidar $WasteDuplicate licencias duplicadas"; Desc="Usuarios con multiples SKUs que proveen el mismo producto. El standalone se puede remover." })
}
if ([int]$WasteDisPlans -gt 0) {
    $ActionList.Add(@{ Color='yellow'; Title="Revisar $WasteDisPlans usuarios con planes de seguridad deshabilitados"; Desc="Usuarios con licencia de pago pero con la mayoria de planes de seguridad apagados. Verificar si fue intencional o un error de configuracion." })
}
if ($AdoptData -and $AdoptData.Entra -and $AdoptData.Entra.PIM -and $AdoptData.Entra.PIM.PermanentUsers -gt 5) {
    $ActionList.Add(@{ Color='yellow'; Title="Reducir $($AdoptData.Entra.PIM.PermanentUsers) asignaciones permanentes de roles privilegiados"; Desc="Migrar a asignaciones elegibles (just-in-time) con PIM para reducir superficie de ataque." })
}

if ($ActionList.Count -gt 0) {
    $ActionsHtml = "<div class='sec' id='actions'>Acciones Prioritarias</div><div class='priority-list'>"
    for ($ai = 0; $ai -lt $ActionList.Count; $ai++) {
        $Act = $ActionList[$ai]
        $ActionsHtml += "<div class='priority-item'><div class='priority-num pn-$($Act.Color)'>$($ai + 1)</div><div class='priority-body'><div class='pb-title'>$(HtmlEncode $Act.Title)</div><div class='pb-desc'>$(HtmlEncode $Act.Desc)</div></div></div>"
    }
    $ActionsHtml += "</div>"
    $ActionsHtml
} else {
    "<div class='sec' id='actions'>Acciones Prioritarias</div><div class='priority-item'><div class='priority-num' style='background:var(--green)'>&#10003;</div><div class='priority-body'><div class='pb-title'>Sin acciones criticas</div><div class='pb-desc'>El tenant esta bien optimizado. No se detectaron problemas urgentes.</div></div></div>"
}
)

<!-- ====== RESUMEN NARRATIVO ====== -->
$(
$PaidSkuCount = @($LicData.SKUs | Where-Object { $_.Total -gt 0 -and $_.FriendlyName -notmatch '(?i)(free|trial|exploratory)' }).Count
$TotalPaidSeats = ($LicData.SKUs | Where-Object { $_.Total -gt 0 -and $_.FriendlyName -notmatch '(?i)(free|trial|exploratory)' } | Measure-Object -Property Total -Sum).Sum
if (-not $TotalPaidSeats) { $TotalPaidSeats = 0 }
$TotalAssigned = ($LicData.SKUs | Where-Object { $_.Total -gt 0 -and $_.FriendlyName -notmatch '(?i)(free|trial|exploratory)' } | Measure-Object -Property Assigned -Sum).Sum
if (-not $TotalAssigned) { $TotalAssigned = 0 }
$PctAssigned = if ($TotalPaidSeats -gt 0) { [math]::Round(($TotalAssigned / $TotalPaidSeats) * 100, 0) } else { 0 }
$DominantMethod = if ($MethodGroup -ge $MethodDirect -and $MethodGroup -ge $MethodMixed) { 'grupo' } elseif ($MethodDirect -ge $MethodMixed) { 'directa' } else { 'mixta' }
$NarrHtml = "<div class='narrative-box'>"
$NarrHtml += "<strong>$(HtmlEncode $LicData.TenantName)</strong> cuenta con <strong>$($TotalUsers.ToString('N0'))</strong> usuarios licenciados en <strong>$PaidSkuCount</strong> SKU$(if($PaidSkuCount -ne 1){'s'}) de pago ($($TotalPaidSeats.ToString('N0')) seats), con una asignacion del <strong>${PctAssigned}%</strong>. "
$NarrHtml += "El metodo predominante es por <strong>$DominantMethod</strong>."
if ($ScoreData -and $ScoreData.Score) {
    $NarrHtml += " El Secure Score actual es <strong>$ScoreCur / $ScoreMax ($ScorePct%)</strong>."
}
if ($WasteTotal -gt 0) {
    $WastePct = [math]::Round(($WasteTotal / [math]::Max($TotalUsers,1)) * 100, 0)
    $NarrHtml += "<br>Se identificaron <strong>$($WasteTotal.ToString('N0'))</strong> licencias a revisar (<strong>${WastePct}%</strong> del total)"
    $WasteParts = @()
    if ($WasteDisabled -gt 0) { $WasteParts += "$($WasteDisabled.ToString('N0')) deshabilitadas" }
    if ($WasteInactive -gt 0) { $WasteParts += "$($WasteInactive.ToString('N0')) inactivos ($($LicData.InactiveDays)d+)" }
    if ($WasteDuplicate -gt 0) { $WasteParts += "$($WasteDuplicate.ToString('N0')) duplicadas" }
    if ([int]$WasteDisPlans -gt 0) { $WasteParts += "$($WasteDisPlans.ToString('N0')) con planes deshabilitados" }
    if ($WasteParts.Count -gt 0) { $NarrHtml += ": " + ($WasteParts -join ", ") }
    $NarrHtml += "."
}
$NarrHtml += "</div>"
$NarrHtml
)
</div><!-- /tab-resumen -->

<div class="tab-panel" id="tab-postura">
<!-- ====== POSTURA DE SEGURIDAD ====== -->
$(if ($AdoptData) {
@"
<div class="sec" id="posture">Postura de Seguridad</div>
<div id="postureContent"></div>
"@
})
</div><!-- /tab-postura -->

<div class="tab-panel" id="tab-recomendaciones">
<!-- ====== RECOMENDACIONES DE SEGURIDAD ====== -->
$(if ($ScoreData -and $ScoreData.AllRecommendations) {
@"
<div class="sec" id="recs">&#128161; Recomendaciones de Seguridad</div>
<p class="muted small" style="margin-bottom:12px">Todas las recomendaciones del Secure Score. Filtra por categoria, estado, o busca por texto.</p>
<button class="csv-btn no-print" onclick="exportRecsCSV()">&#128190; CSV</button>
<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
  <input type="text" class="search" id="recSearch" placeholder="Buscar recomendacion..." onkeyup="window._recFilter()" style="max-width:400px;margin:0;flex:1">
  <div class="tabs no-print" id="recSvcTabs"></div>
  <div class="tabs no-print" id="recSubTabs" style="display:none"></div>
  <div class="tabs no-print" id="recTabs"></div>
  <div class="tabs no-print" id="recStatusTabs"></div>
</div>
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
  <div id="recCount" class="small muted"></div>
  <div class="pg-controls" id="recPgTop"></div>
</div>
<div class="tw">
<table id="recTable">
  <thead><tr><th class="tc">#</th><th class="tc">Impacto</th><th>Recomendacion</th><th>Categoria</th><th>Servicio</th><th class="tc">Estado</th></tr></thead>
  <tbody id="recTbody"></tbody>
</table>
</div>
<div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px">
  <div id="recCountBot" class="small muted"></div>
  <div class="pg-controls" id="recPgBot"></div>
</div>
"@
})
</div><!-- /tab-recomendaciones -->

<div class="tab-panel" id="tab-licenciamiento">
<!-- ====== VISTA GENERAL (GRAFICOS) ====== -->
<div class="sec" id="charts">Vista General</div>
<div class="two-col">
  <div class="card" style="padding:20px;text-align:left">
    <div style="font-size:13px;font-weight:600;margin-bottom:12px;color:var(--accent)">Distribucion de Seats por SKU</div>
    <div style="display:flex;align-items:center;gap:24px">
      <div style="position:relative;width:140px;height:140px;flex-shrink:0">
        <div style="width:140px;height:140px;border-radius:50%;background:conic-gradient($DonutStops)"></div>
        <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:90px;height:90px;border-radius:50%;background:var(--bg2);display:flex;align-items:center;justify-content:center">
          <div style="text-align:center">
            <div style="font-size:20px;font-weight:700;color:var(--accent)">$TotalSeatsFormatted</div>
            <div style="font-size:9px;color:var(--muted)">SEATS</div>
          </div>
        </div>
      </div>
      <div style="flex:1">$DonutLegend</div>
    </div>
  </div>
  <div class="card" style="padding:20px;text-align:left">
    <div style="font-size:13px;font-weight:600;margin-bottom:12px;color:var(--accent)">Asignacion por Producto</div>
    <div style="font-size:10px;color:var(--muted);margin-bottom:8px">Usuarios asignados / Seats disponibles</div>
    $CapacityBars
  </div>
</div>

<!-- ====== SKUs ====== -->
<div class="sec" id="skus">Inventario de Licencias (SKUs)</div>
<button class="csv-btn no-print" onclick="exportTableToCSV('skuExport','SKUs.csv')">&#128190; CSV</button>
<div class="tw">
<table id="skuExport">
  <thead><tr><th>Producto</th><th class="tc">Total</th><th class="tc">Asignadas</th><th class="tc">Libres</th><th>Uso</th><th>Categorias Incluidas</th></tr></thead>
  <tbody>$SkuRows</tbody>
</table>
</div>

<!-- ====== MATRIZ SKU vs PRODUCTOS ====== -->
<div class="sec" id="skumatrix">Que incluye cada Licencia</div>
<p class="muted small" style="margin-bottom:12px">Matriz que muestra que productos de seguridad y compliance estan incluidos en cada SKU del tenant.</p>
$(if ($MatrixRows) { @"
<div class="tw">
<table>
  <thead><tr>
    <th>SKU</th><th class="tc">Total</th><th class="tc">Asignadas</th>
    $MatrixHeaders
  </tr></thead>
  <tbody>$MatrixRows</tbody>
</table>
</div>
"@ } else { "<p class='muted'>No hay datos de matriz disponibles.</p>" })

<!-- ====== CAPACIDAD vs ASIGNACION ====== -->
<div class="sec" id="capacity">Capacidad vs Asignacion</div>
<button class="csv-btn no-print" onclick="exportTableToCSV('capExport','Capacidad.csv')">&#128190; CSV</button>
<p class="muted small" style="margin-bottom:12px">
  Compara los seats disponibles por producto contra los usuarios asignados. <strong>Activos</strong> = plan habilitado. <strong>Deshab.</strong> = plan deshabilitado por admin.
  <strong style="color:var(--red)">RIESGO</strong> = mas usuarios que seats.
  <strong style="color:#4dd0e1">Overlap</strong> = multiples SKUs proveen la misma funcionalidad a los mismos usuarios.
  <strong style="color:var(--yellow)">Sobreasignado</strong> = menos del 30% de asignacion.
</p>
$(if ($CapacityRows) { @"
<div class="tw">
<table id="capExport">
  <thead><tr><th>Categoria</th><th>Producto</th><th class="tc">Seats</th><th class="tc">Activos</th><th class="tc">Deshabilitados</th><th class="tc" title="Usuarios con la misma funcionalidad en 2+ SKUs">Overlap</th><th class="tc">Libres</th><th>Asignacion</th><th class="tc">Estado</th><th>Provisto por</th></tr></thead>
  <tbody>$CapacityRows</tbody>
</table>
</div>
"@ } else { "<p class='muted'>No hay datos de capacidad disponibles.</p>" })

<!-- ====== DEPARTAMENTOS ====== -->
<div class="sec" id="depts">Resumen por Departamento</div>
<button class="csv-btn no-print" onclick="exportTableToCSV('deptExport','Departamentos.csv')">&#128190; CSV</button>
<div class="tw">
<table id="deptExport">
  <thead><tr><th>Departamento</th><th class="tc">Usuarios</th><th>SKUs</th><th>Productos Activos</th></tr></thead>
  <tbody>$DeptRows</tbody>
</table>
</div>
</div><!-- /tab-licenciamiento -->

<div class="tab-panel" id="tab-optimizacion">
<!-- ====== LICENCIAS A REVISAR ====== -->
<div class="sec" id="waste">Licencias a Revisar ($WasteTotal usuarios)</div>
<button class="csv-btn no-print" onclick="exportWasteCSV()">&#128190; CSV</button>
<p class="muted small" style="margin-bottom:12px">
  Usuarios con licencias que podrian optimizarse: cuentas deshabilitadas, inactivos $($LicData.InactiveDays)+ dias, licencias duplicadas o planes de seguridad deshabilitados.
</p>
$(if ($WasteJsonList.Count -gt 0) { @"
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
  <input type="text" class="search" id="wasteSearch" placeholder="Buscar en licencias a revisar..." onkeyup="window._wasteFilter()" style="max-width:400px;margin:0">
  <div class="pg-controls" id="wastePgTop"></div>
</div>
<div class="tw">
<table id="wasteTable">
  <thead><tr><th>Nombre</th><th>UPN</th><th class="tc">Activa</th><th class="tc">Ultimo Sign-In</th><th>SKUs</th><th>Motivo</th></tr></thead>
  <tbody id="wasteTbody"></tbody>
</table>
</div>
<div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px">
  <div id="wasteCount" class="small muted"></div>
  <div class="pg-controls" id="wastePgBot"></div>
</div>
"@ } else { "<p class='muted'>No se detectaron licencias a revisar.</p>" })

<!-- ====== DUPLICADOS ====== -->
<div class="sec" id="duplicates">Licencias Duplicadas ($DupCount entradas)</div>
<button class="csv-btn no-print" onclick="exportDupsCSV()">&#128190; CSV</button>
<p class="muted small" style="margin-bottom:12px">Usuarios con multiples SKUs que proveen el mismo producto. El SKU standalone podria removerse.</p>
$(if ($DupJsonList.Count -gt 0) { @"
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
  <input type="text" class="search" id="dupSearch" placeholder="Buscar en duplicados..." onkeyup="window._dupFilter()" style="max-width:400px;margin:0">
  <div class="pg-controls" id="dupPgTop"></div>
</div>
<div class="tw">
<table id="dupTable">
  <thead><tr><th>Nombre</th><th>UPN</th><th>Producto Duplicado</th><th>Provisto por SKUs</th></tr></thead>
  <tbody id="dupTbody"></tbody>
</table>
</div>
<div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px">
  <div id="dupCount" class="small muted"></div>
  <div class="pg-controls" id="dupPgBot"></div>
</div>
"@ } else { "<p class='muted'>No se detectaron licencias duplicadas.</p>" })
</div><!-- /tab-optimizacion -->

<div class="tab-panel" id="tab-usuarios">
<!-- ====== USUARIOS ====== -->
<div class="sec" id="users">Detalle por Usuario ($($Users.Count))</div>
<button class="csv-btn no-print" onclick="exportUsersCSV()">&#128190; CSV</button>
<input type="text" class="search" id="userSearch" placeholder="Buscar por nombre, UPN, departamento, SKU..." onkeyup="applyFilters()">
<div style="margin-bottom:10px" class="no-print"><div style="margin-bottom:6px;font-size:11px;color:var(--muted)">Estado:</div>
<div class="tabs" id="statusFilters"><button class="tab active" onclick="setStatusFilter('all',this)">Todos</button><button class="tab" onclick="setStatusFilter('active',this)">Activos</button><button class="tab" onclick="setStatusFilter('disabled',this)">Deshabilitados</button><button class="tab" onclick="setStatusFilter('inactive',this)">Inactivos</button><button class="tab" onclick="setStatusFilter('disPlans',this)">Planes Deshab.</button><button class="tab" onclick="setStatusFilter('waste',this)">A Revisar</button></div></div>
<div style="margin-bottom:10px" class="no-print"><div style="margin-bottom:6px;font-size:11px;color:var(--muted)">Departamento:</div>
<div class="tabs" id="deptFilters"><button class="tab active" onclick="filterByDept('all')">Todos</button>$(foreach ($D in $UniqueDepts) { " <button class='tab' onclick=`"filterByDept('$(HtmlEncode ($D -replace "'","\'"))')`">$(HtmlEncode $D)</button>" })</div></div>
<div style="margin-bottom:10px" class="no-print"><div style="margin-bottom:6px;font-size:11px;color:var(--muted)">SKU:</div>
<div id="skuChips" class="tabs"><button class="tab active" onclick="clearSkuFilter()">Todos</button>$(foreach ($S in $UniqueSkus) { " <button class='tab' onclick=`"toggleSkuFilter(this,'$(HtmlEncode ($S -replace "'","\'"))')`">$(HtmlEncode $S)</button>" })</div>
<div style="margin-top:6px"><label style="font-size:11px;color:var(--muted);cursor:pointer"><input type="checkbox" id="skuMatchAll" onchange="applySkuFilter()"> Mostrar solo usuarios con TODOS los SKUs seleccionados</label></div></div>

<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px">
  <div id="userCount" class="small muted"></div>
  <div class="pg-controls" id="pgTop"></div>
</div>
<div class="legend">
  <span><span class="dot dot-on"></span> Activo</span><span><span class="dot dot-off"></span> Deshabilitado</span><span><span class="dot dot-na"></span> No incluido</span>
  <span><span class="badge b-green" style="font-size:9px">CA</span> Conditional Access</span>
  <span><span class="badge b-blue" style="font-size:9px">G</span> Grupo</span>
  <span><span class="badge b-default" style="font-size:9px">D</span> Directa</span>
  <span><span class="badge b-yellow" style="font-size:9px">G+D</span> Mixta</span>
</div>
<div class="tw">
<table id="userTable">
  <thead><tr><th>Nombre</th><th>UPN</th><th>Depto</th><th>SKUs</th><th class="tc">Sign-In</th><th class="tc">Dias</th>
$(
    $catHeaders = ""
    for ($ci = 0; $ci -lt $SecurityCats.Count; $ci++) {
        $Cat = $SecurityCats[$ci]
        $Full = if ($CatFullNames.ContainsKey($Cat)) { $CatFullNames[$Cat] } else { $Cat }
        $Short = ($Cat -replace '_',' ' -replace 'Purview ','').Substring(0, [math]::Min(($Cat -replace '_',' ' -replace 'Purview ','').Length, 8))
        $Sep = if ($CatSepIndices -contains $ci) { " td-sep" } else { "" }
        $catHeaders += "    <th class='tc$Sep' title='$(HtmlEncode $Full)' style='font-size:9px;max-width:45px;overflow:hidden;writing-mode:vertical-lr;text-orientation:mixed;height:80px'>$Short</th>`n"
    }
    $catHeaders
)  </tr></thead>
  <tbody id="userTbody"></tbody>
</table>
</div>
<div style="display:flex;justify-content:space-between;align-items:center;margin-top:10px;flex-wrap:wrap;gap:8px">
  <div id="userCountBottom" class="small muted"></div>
  <div class="pg-controls" id="pgBottom"></div>
</div>
</div><!-- /tab-usuarios -->

</div><!-- /wrap -->

<div class="footer">Microsoft 365 Security Assessment | $TenantName | $GeneratedAt</div>

<!-- ====== EMBEDDED DATA ====== -->
<script>
const D=$MasterJson;
const _WD=$WasteDataJson;
const _DD=$DupDataJson;
const _ID=$InactiveDaysJs;
const _UD=$("[" + ($UserJsonList -join ",`n") + "]");
const _CM=$("[" + ($CatMetaList -join ",") + "]");
const _CS=$("[" + ($CatSepIndices -join ",") + "]");
const _DEPTS=$(($UniqueDepts | ForEach-Object { "`"$($_ -replace '"','\"')`"" }) -join "," | ForEach-Object { "[$_]" });
const _SKUS=$(($UniqueSkus | ForEach-Object { "`"$($_ -replace '"','\"')`"" }) -join "," | ForEach-Object { "[$_]" });
</script>

<script>
// ============================================================================
// JS RENDER ENGINE
// ============================================================================
const esc=s=>{if(!s)return'';const d=document.createElement('div');d.textContent=s;return d.innerHTML};
const fmt=n=>typeof n==='number'?n.toLocaleString('es'):n||'0';
const clr=p=>p>=70?'var(--green)':p>=40?'var(--yellow)':'var(--red)';
const badge=(t,c)=>'<span class="badge b-'+c+'">'+esc(t)+'</span>';
const pgBar=(p,w)=>{const c=p>=75?'pg-green':p>=40?'pg-yellow':'pg-red';return'<div class="pg" style="width:'+(w||'80px')+'"><div class="pg-fill '+c+'" style="width:'+Math.min(p,100)+'%"></div></div>'};

// Pagination helper
function pgHtml(page,tp,ns){
  if(tp<=1)return'';
  let h='<div style="display:flex;gap:4px;align-items:center">';
  h+='<button class="pg-btn'+(page===1?' pg-dis':'')+'" onclick="window._pg.'+ns+'(1)">&laquo;</button>';
  h+='<button class="pg-btn'+(page===1?' pg-dis':'')+'" onclick="window._pg.'+ns+'('+(page-1)+')">&lsaquo;</button>';
  const s=Math.max(1,page-2),e=Math.min(tp,page+2);
  for(let p=s;p<=e;p++)h+='<button class="pg-btn'+(p===page?' pg-active':'')+'" onclick="window._pg.'+ns+'('+p+')">'+p+'</button>';
  h+='<button class="pg-btn'+(page===tp?' pg-dis':'')+'" onclick="window._pg.'+ns+'('+(page+1)+')">&rsaquo;</button>';
  h+='<button class="pg-btn'+(page===tp?' pg-dis':'')+'" onclick="window._pg.'+ns+'('+tp+')">&raquo;</button>';
  h+='<span style="color:var(--muted);font-size:11px;margin-left:8px">Pag '+page+'/'+tp+'</span></div>';
  return h;
}
window._pg={};

const L=D.lic,A=D.adopt,S=D.score;

// ============================================================================
// SCORE POR CATEGORIA (fix: cap percentages, handle missing MaxScore)
// ============================================================================
if(S&&S.Categories){
  const el=document.getElementById('scoreBars');
  if(el){
    let h='';
    S.Categories.forEach(c=>{
      const p=c.MaxScore>0?Math.min(Math.round(c.Score/c.MaxScore*100),100):(c.PctScore?Math.min(c.PctScore,100):0);
      const col=clr(p);
      const label=c.MaxScore?c.Score+'/'+c.MaxScore+' ('+p+'%)':'N/A';
      h+='<div class="bar-row"><span class="bar-label">'+esc(c.Category)+'</span>';
      h+='<div class="bar-track"><div class="bar-fill" style="width:'+p+'%;background:'+col+'"></div></div>';
      h+='<span class="bar-value">'+label+'</span></div>';
    });
    el.innerHTML=h;
  }
  const el2=document.getElementById('scoreControls');
  if(el2&&S.Summary){
    let h='<div class="sc-stat"><div class="v" style="color:var(--green)">'+S.Summary.Implemented+'</div><div class="l">Implementados</div></div>';
    h+='<div class="sc-stat"><div class="v" style="color:var(--yellow)">'+S.Summary.Partial+'</div><div class="l">Parciales</div></div>';
    h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+S.Summary.NotImplemented+'</div><div class="l">No Implementados</div></div>';
    h+='<div class="sc-stat"><div class="v">'+S.Summary.TotalControls+'</div><div class="l">Total</div></div>';
    el2.innerHTML=h;
  }
}

// ============================================================================
// POSTURA DE SEGURIDAD — Scorecard redesign
// ============================================================================
if(A){
  const pc=document.getElementById('postureContent');
  if(pc){
    // Helper: semaphore dot + verdict
    function sem(pct){
      if(pct>=80) return {c:'var(--green)',t:'Saludable',i:'&#9679;'};
      if(pct>=50) return {c:'var(--yellow)',t:'Parcial',i:'&#9679;'};
      return {c:'var(--red)',t:'Critico',i:'&#9679;'};
    }
    function semDot(pct){const s=sem(pct);return '<span style="color:'+s.c+';font-size:14px;margin-right:6px">'+s.i+'</span><span style="color:'+s.c+';font-weight:600;font-size:13px">'+s.t+'</span>';}
    // Progress mini-bar
    function miniBar(pct,label,color){
      const c=color||clr(pct);
      return '<div style="margin:6px 0"><div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:3px"><span>'+label+'</span><span style="color:'+c+';font-weight:700">'+pct+'%</span></div>'
        +'<div style="height:6px;background:rgba(255,255,255,0.07);border-radius:3px"><div style="height:100%;width:'+Math.min(pct,100)+'%;background:'+c+';border-radius:3px;transition:width .5s"></div></div></div>';
    }

    let h='';
    const licCount=A.TotalLicensedUsers||0;

    // ═══════════════════════════════════════════════════════
    // Entra ID - Identidad
    // ═══════════════════════════════════════════════════════
    if(A.Entra){
      const E=A.Entra;
      // --- MFA & SSPR Card ---
      if(E.MFA){
        const m=E.MFA;
        const mainPct=m.PctRegistered;
        const mainReg=m.Registered;
        const mainTot=m.TotalUsers;
        h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128272; MFA & SSPR '+semDot(mainPct)+'</div>';
        h+='<div class="two-col" style="margin-bottom:0">';
        // MFA panel
        h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px">Multi-Factor Authentication</div>';
        const mfaLabel=(m.LicTotal&&m.LicTotal<=500)?'Usuarios Licenciados':'Directorio';
        h+=miniBar(m.PctRegistered,mfaLabel+' ('+fmt(m.Registered)+'/'+fmt(m.TotalUsers)+')');
        h+='</div>';
        // SSPR panel
        if(E.SSPR){
          const s=E.SSPR;
          const ssprLabel=(s.LicTotal&&s.LicTotal<=500)?'Usuarios Licenciados':'Directorio';
          h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px">Self-Service Password Reset</div>';
          h+=miniBar(s.PctRegistered,ssprLabel+' ('+fmt(s.Registered)+'/'+fmt(s.TotalUsers)+')');
          h+='</div>';
        }
        h+='</div>';

        // Auth Methods - Licensed vs Tenant
        const amLic=E.AuthMethodsLicensed;
        const amAll=E.AuthMethods;
        const amSrc=amLic||amAll;
        if(amSrc){
          const amLabel=(m.LicTotal&&m.LicTotal<=500)?'Usuarios Licenciados':'Directorio';
          h+='<div style="margin-top:12px"><div class="sc-title" style="font-size:12px">Metodos de Autenticacion ('+amLabel+')</div>';
          const methods=[['Authenticator',amSrc.Authenticator],['Phone',amSrc.PhoneAuth],['FIDO2',amSrc.FIDO2],['WHfB',amSrc.WHfB],['Passwordless',amSrc.Passwordless],['Email',amSrc.Email]];
          const base=(m.LicTotal&&m.LicTotal<=500)?m.TotalUsers:((E.MFA&&E.MFA.MfaCapable>0)?E.MFA.MfaCapable:methods.reduce((s,m)=>Math.max(s,m[1]||0),0)||1);
          methods.forEach(m=>{
            if(!m[1])return;
            const p=Math.round(m[1]/base*100);
            h+='<div class="bar-row"><span class="bar-label">'+m[0]+'</span>';
            h+='<div class="bar-track"><div class="bar-fill" style="width:'+Math.min(p,100)+'%;background:var(--accent)"></div></div>';
            h+='<span class="bar-value">'+m[1]+' ('+p+'%)</span></div>';
          });
          h+='</div>';
        }
        h+='</div>';
      }

      // --- Conditional Access ---
      if(E.ConditionalAccess){
        const ca=E.ConditionalAccess;
        const covPct=ca.LicCoveragePct||0;
        h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128272; Conditional Access '+semDot(covPct)+'</div>';
        // Coverage bar
        if(ca.LicTotal>0){
          h+=miniBar(covPct,'Cobertura Usuarios Licenciados ('+ca.LicCovered+'/'+ca.LicTotal+')');
          if(ca.HasAllUsersPolicy) h+='<div style="font-size:10px;color:var(--green);margin-bottom:8px">&#10003; Al menos una politica activa aplica a todos los usuarios</div>';
        }
        h+='<div class="sc-grid" style="margin-bottom:10px">';
        h+='<div class="sc-stat"><div class="v" style="color:var(--green)">'+ca.Enabled+'</div><div class="l">Activas</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--blue)">'+ca.ReportOnly+'</div><div class="l">Report-Only</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--muted)">'+ca.Disabled+'</div><div class="l">Deshabilitadas</div></div>';
        h+='<div class="sc-stat"><div class="v">'+ca.Total+'</div><div class="l">Total</div></div>';
        h+='</div>';
        h+='</div>';
      }

      // --- Identity Protection & PIM ---
      if(E.RiskyUsers||E.PIM){
        h+='<div class="sc"><div class="sc-title">&#128272; Entra ID - Identity Protection & PIM</div>';
        h+='<div class="two-col" style="margin-bottom:0">';
        if(E.RiskyUsers){
          const r=E.RiskyUsers;
          const riskOk=r.TotalAtRisk===0;
          h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px;display:flex;align-items:center;justify-content:space-between">Usuarios Riesgosos '+(riskOk?semDot(100):semDot(0))+'</div>';
          // Risk summary bar
          if(r.TotalAtRisk>0){
            const totalR=r.High+r.Medium+r.Low||1;
            const pH=Math.round(r.High/totalR*100);
            const pM=Math.round(r.Medium/totalR*100);
            const pL=100-pH-pM;
            h+='<div style="display:flex;height:8px;border-radius:4px;overflow:hidden;margin:8px 0">';
            if(r.High>0) h+='<div style="width:'+pH+'%;background:var(--red)" title="High: '+r.High+'"></div>';
            if(r.Medium>0) h+='<div style="width:'+pM+'%;background:var(--yellow)" title="Medium: '+r.Medium+'"></div>';
            if(r.Low>0) h+='<div style="width:'+pL+'%;background:var(--blue)" title="Low: '+r.Low+'"></div>';
            h+='</div>';
          }
          h+='<div class="sc-grid">';
          h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+r.High+'</div><div class="l">High</div></div>';
          h+='<div class="sc-stat"><div class="v" style="color:var(--yellow)">'+r.Medium+'</div><div class="l">Medium</div></div>';
          h+='<div class="sc-stat"><div class="v" style="color:var(--blue)">'+r.Low+'</div><div class="l">Low</div></div>';
          h+='<div class="sc-stat"><div class="v">'+r.TotalAtRisk+'</div><div class="l">Total</div></div>';
          h+='</div>';
          // Risk policies status
          h+='<div style="margin-top:10px;font-size:11px">';
          h+='<div style="font-weight:600;color:var(--muted);margin-bottom:4px">Politicas de Proteccion Basadas en Riesgo</div>';
          const sirOk=r.HasSignInRiskPolicy;
          const urOk=r.HasUserRiskPolicy;
          h+='<div style="padding:2px 0"><span style="color:'+(sirOk?'var(--green)':'var(--red)')+'">&#'+(sirOk?'10003':'10007')+'; </span>Sign-in Risk Policy '+(sirOk?'<span style="color:var(--green)">activa</span>':'<span style="color:var(--red)">no configurada</span>')+'</div>';
          h+='<div style="padding:2px 0"><span style="color:'+(urOk?'var(--green)':'var(--red)')+'">&#'+(urOk?'10003':'10007')+'; </span>User Risk Policy '+(urOk?'<span style="color:var(--green)">activa</span>':'<span style="color:var(--red)">no configurada</span>')+'</div>';
          h+='</div>';
          // Contextual insight
          if(r.TotalAtRisk>0&&r.High>0){
            h+='<div style="font-size:11px;color:var(--red);margin-top:8px;padding:6px 10px;background:rgba(244,67,54,0.08);border-radius:6px;border-left:3px solid var(--red)">&#9888; '+r.High+' usuario(s) en riesgo alto requieren investigacion inmediata. '+(sirOk&&urOk?'Las politicas de riesgo estan activas y deberian forzar remediacion automatica.':'Se recomienda configurar politicas de riesgo en Conditional Access para forzar cambio de contrasena o MFA.')+'</div>';
          } else if(r.TotalAtRisk===0){
            h+='<div style="font-size:11px;color:var(--green);margin-top:8px;padding:6px 10px;background:rgba(76,175,80,0.08);border-radius:6px;border-left:3px solid var(--green)">&#10003; No se detectaron usuarios en riesgo activo. Identity Protection esta monitoreando correctamente.</div>';
          }
          h+='</div>';
        }
        if(E.PIM){
          const p=E.PIM;
          const permU=p.PermanentUsers||0;
          const activeU=p.ActiveUsers||0;
          const activeSP=p.ActiveSPs||0;
          const eligU=p.EligibleUsers||0;
          const totalActive=p.ActiveTotal||0;
          // Ratio: elegibles vs permanentes humanos — mas elegibles = mejor
          const jitPct=activeU>0?Math.round(eligU/(eligU+permU)*100):100;
          const pimOk=permU<=5;
          h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px;display:flex;align-items:center;justify-content:space-between">PIM (Privileged Identity) '+(pimOk?semDot(80):semDot(30))+'</div>';
          // Main insight: JIT adoption bar
          h+=miniBar(jitPct,'Adopcion JIT ('+eligU+' elegibles vs '+permU+' permanentes)');
          // Stacked composition bar: eligibles | active humans | SPs
          const barTotal=eligU+activeU+activeSP||1;
          const bE=Math.round(eligU/barTotal*100);
          const bU=Math.round(activeU/barTotal*100);
          const bS=100-bE-bU;
          h+='<div style="display:flex;height:8px;border-radius:4px;overflow:hidden;margin:10px 0 4px">';
          h+='<div style="width:'+bE+'%;background:var(--green)" title="Elegibles: '+eligU+'"></div>';
          h+='<div style="width:'+bU+'%;background:var(--yellow)" title="Usuarios permanentes: '+activeU+'"></div>';
          if(activeSP>0) h+='<div style="width:'+bS+'%;background:var(--bg4)" title="Service Principals: '+activeSP+'"></div>';
          h+='</div>';
          h+='<div style="display:flex;gap:14px;font-size:10px;color:var(--muted);margin-bottom:10px">';
          h+='<span><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:var(--green);margin-right:4px;vertical-align:middle"></i>Elegibles JIT: '+eligU+'</span>';
          h+='<span><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:var(--yellow);margin-right:4px;vertical-align:middle"></i>Usuarios permanentes: '+permU+'</span>';
          if(activeSP>0) h+='<span><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:var(--bg4);border:1px solid var(--border);margin-right:4px;vertical-align:middle"></i>Service Principals: '+activeSP+'</span>';
          h+='</div>';
          // Top roles as compact bar chart
          if(p.TopPermanentRoles&&p.TopPermanentRoles.length){
            h+='<div style="margin-top:4px"><div style="font-size:10px;color:var(--muted);margin-bottom:6px">Roles con mas usuarios permanentes</div>';
            const maxR=p.TopPermanentRoles[0].Count||1;
            p.TopPermanentRoles.forEach(r=>{
              const w=Math.round(r.Count/maxR*100);
              h+='<div class="bar-row"><span class="bar-label" style="width:140px;font-size:10px">'+esc(r.Role)+'</span>';
              h+='<div class="bar-track"><div class="bar-fill" style="width:'+w+'%;background:var(--yellow)"></div></div>';
              h+='<span class="bar-value" style="font-size:10px">'+r.Count+'</span></div>';
            });
            h+='</div>';
          }
          // Contextual note
          if(permU>5)h+='<div style="font-size:11px;color:var(--yellow);margin-top:8px;padding:6px 10px;background:rgba(255,193,7,0.08);border-radius:6px;border-left:3px solid var(--yellow)">&#9888; '+permU+' usuarios con roles permanentes. Se recomienda migrar a asignaciones elegibles con activacion JIT.</div>';
          else h+='<div style="font-size:11px;color:var(--green);margin-top:8px;padding:6px 10px;background:rgba(76,175,80,0.08);border-radius:6px;border-left:3px solid var(--green)">&#10003; Buen uso de PIM — pocos usuarios con roles permanentes.</div>';
          if(activeSP>0)h+='<div style="font-size:10px;color:var(--muted);margin-top:6px;font-style:italic">&#8505; Los '+activeSP+' Service Principals con roles activos son aplicaciones del sistema y son esperados.</div>';
          h+='</div>';
        }
        h+='</div></div>';
      }
    }

    // ── ctrlBar: compact Secure Score bar for Defender cards ──
    function ctrlBar(sc,svcFilter){
      if(!sc)return '';
      const ok=sc.FullyEnabled||0,pa=sc.Partial||0,ni=sc.NotImplemented||0;
      const tot=ok+pa+ni||1;
      const wOk=Math.round(ok/tot*100),wPa=Math.round(pa/tot*100),wNi=100-wOk-wPa;
      let b='<div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--border)">';
      b+='<div style="font-size:10px;color:var(--muted);margin-bottom:4px">Secure Score &mdash; '+ok+'/'+sc.Total+' controles implementados</div>';
      b+='<div style="display:flex;height:8px;border-radius:4px;overflow:hidden">';
      b+='<div style="width:'+wOk+'%;background:var(--green)"></div>';
      b+='<div style="width:'+wPa+'%;background:var(--yellow)"></div>';
      b+='<div style="width:'+wNi+'%;background:var(--red)"></div></div>';
      b+='<div style="display:flex;gap:12px;font-size:10px;color:var(--muted);margin-top:4px">';
      b+='<span>&#10003; '+ok+' OK</span><span>&#9899; '+pa+' Parcial</span><span>&#10007; '+ni+' Pendiente</span></div>';
      b+='<div style="margin-top:6px;text-align:right"><a href="javascript:void(0)" onclick="_recSvc(\''+svcFilter+'\')" style="font-size:10px;color:var(--accent);text-decoration:none">Ver detalle en Recomendaciones &#8594;</a></div>';
      b+='</div>';
      return b;
    }

    // ═══════════════════════════════════════════════════════
    // MDE
    // ═══════════════════════════════════════════════════════
    if(A.MDE){
      const M=A.MDE;
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128737; Defender for Endpoint '+semDot(M.CoveragePct)+'</div>';
      h+=miniBar(M.CoveragePct,'Cobertura ('+M.UniqueUsersWithDevice+'/'+M.UsersWithLicense+' usuarios)');
      h+='<div class="sc-grid">';
      h+='<div class="sc-stat"><div class="v">'+M.DevicesOnboarded+'</div><div class="l">Dispositivos</div></div>';
      h+='<div class="sc-stat"><div class="v" style="color:var(--yellow)">'+M.DevicesStale7d+'</div><div class="l">Inactivos 7d+</div></div>';
      if(M.Alerts30d){
        h+='<div class="sc-stat"><div class="v" style="color:'+(M.Alerts30d.High>0?'var(--red)':'var(--green)')+'">'+M.Alerts30d.Total+'</div><div class="l">Alertas 30d</div></div>';
        if(M.Alerts30d.High>0)h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+M.Alerts30d.High+'</div><div class="l">Alta Severidad</div></div>';
      }
      h+='</div>';
      h+=ctrlBar(M.SecureScoreControls,'MDATP');
      h+='</div>';
    }

    // ═══════════════════════════════════════════════════════
    // MDO
    // ═══════════════════════════════════════════════════════
    if(A.MDO){
      const O=A.MDO;
      const mdoThreats=(O.PhishingDetected||0)+(O.MalwareDetected||0);
      const mdoTotal=O.EmailsProcessed30d||1;
      const mdoBlockPct=mdoTotal>0?Math.round((O.Blocked||0)/mdoTotal*100):0;
      const hasKQL=O.EmailsProcessed30d>0;
      const mdoClean=hasKQL?(mdoThreats===0?100:(mdoBlockPct>90?80:50)):(O.SecureScoreControls?70:50);
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128231; Defender for Office 365 '+semDot(mdoClean)+'</div>';
      if(hasKQL){
        h+='<div class="sc-grid">';
        h+='<div class="sc-stat"><div class="v">'+fmt(O.EmailsProcessed30d)+'</div><div class="l">Emails 30d</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+O.PhishingDetected+'</div><div class="l">Phishing</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+O.MalwareDetected+'</div><div class="l">Malware</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--green)">'+O.Blocked+'</div><div class="l">Bloqueados</div></div>';
        if(O.SafeLinks)h+='<div class="sc-stat"><div class="v">'+O.SafeLinks.Blocked+'</div><div class="l">Safe Links</div></div>';
        h+='</div>';
      } else {
        h+='<div style="font-size:11px;color:var(--muted);margin:6px 0;font-style:italic">&#9432; Sin datos de flujo de correo en los ultimos 30 dias. Configuracion evaluada via Secure Score.</div>';
      }
      h+=ctrlBar(O.SecureScoreControls,'MDO');
      if(O.Note)h+='<div style="font-size:11px;color:var(--muted);margin-top:6px;font-style:italic">'+esc(O.Note)+'</div>';
      h+='</div>';
    }

    // ═══════════════════════════════════════════════════════
    // MDA
    // ═══════════════════════════════════════════════════════
    if(A.MDA){
      const C=A.MDA;
      const mdaHealth=C.SecureScoreControls?(C.SecureScoreControls.FullyEnabled/Math.max(C.SecureScoreControls.Total,1)*100):70;
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#9729; Defender for Cloud Apps '+semDot(mdaHealth)+'</div>';
      if(C.UniqueApps>0||C.Events30d>0){
        h+='<div class="sc-grid">';
        h+='<div class="sc-stat"><div class="v">'+C.UniqueApps+'</div><div class="l">Apps Descubiertas</div></div>';
        h+='<div class="sc-stat"><div class="v">'+fmt(C.Events30d)+'</div><div class="l">Eventos 30d</div></div>';
        h+='<div class="sc-stat"><div class="v">'+C.UniqueUsers+'</div><div class="l">Usuarios</div></div>';
        if(C.Alerts30d)h+='<div class="sc-stat"><div class="v" style="color:'+(C.Alerts30d.Total>0?'var(--yellow)':'var(--green)')+'">'+C.Alerts30d.Total+'</div><div class="l">Alertas 30d</div></div>';
        h+='</div>';
      } else {
        h+='<div style="font-size:11px;color:var(--muted);margin:6px 0;font-style:italic">&#9432; Sin datos de actividad de Cloud Apps. Configuracion evaluada via Secure Score.</div>';
      }
      h+=ctrlBar(C.SecureScoreControls,'MCAS');
      h+='</div>';
    }

    // ═══════════════════════════════════════════════════════
    // MDI
    // ═══════════════════════════════════════════════════════
    if(A.MDI){
      const I2=A.MDI;
      const failRatio=I2.LogonEvents30d>0?Math.round(I2.FailedLogons/I2.LogonEvents30d*100):0;
      const mdiHealth=failRatio<20?90:(failRatio<50?60:30);
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128737; Defender for Identity '+semDot(mdiHealth)+'</div>';
      if(I2.LogonEvents30d>0){
        h+='<div class="sc-grid">';
        h+='<div class="sc-stat"><div class="v" style="color:var(--green)">'+I2.DCsMonitored+'</div><div class="l">DCs Monitoreados</div></div>';
        h+='<div class="sc-stat"><div class="v">'+fmt(I2.LogonEvents30d)+'</div><div class="l">Eventos Logon 30d</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--green)">'+fmt(I2.SuccessLogons)+'</div><div class="l">Exitosos</div></div>';
        h+='<div class="sc-stat"><div class="v" style="color:var(--yellow)">'+fmt(I2.FailedLogons)+'</div><div class="l">Fallidos ('+failRatio+'%)</div></div>';
        if(I2.Alerts30d)h+='<div class="sc-stat"><div class="v" style="color:'+(I2.Alerts30d.Total>0?'var(--red)':'var(--green)')+'">'+I2.Alerts30d.Total+'</div><div class="l">Alertas 30d</div></div>';
        h+='</div>';
      } else {
        h+='<div style="font-size:11px;color:var(--muted);margin:6px 0;font-style:italic">&#9432; Sin datos de logon en los ultimos 30 dias. Configuracion evaluada via Secure Score.</div>';
      }
      h+=ctrlBar(I2.SecureScoreControls,'Azure ATP');
      if(I2.Note)h+='<div style="font-size:11px;color:var(--muted);margin-top:8px;font-style:italic">&#9888; '+esc(I2.Note)+'</div>';
      h+='</div>';
    }

    // ═══════════════════════════════════════════════════════
    // Intune
    // ═══════════════════════════════════════════════════════
    if(A.Intune){
      const I=A.Intune;
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#128241; Microsoft Intune '+semDot(I.CompliancePct)+'</div>';
      h+='<div class="two-col" style="margin-bottom:0">';
      h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px">Dispositivos</div>';
      h+=miniBar(I.CompliancePct,'Cumplimiento ('+I.DevicesEnrolled+' enrolled)');
      h+='<div class="sc-grid" style="margin-top:8px">';
      h+='<div class="sc-stat"><div class="v" style="color:var(--red)">'+I.NonCompliant+'</div><div class="l">No Compliant</div></div>';
      h+='<div class="sc-stat"><div class="v" style="color:var(--yellow)">'+I.Stale30d+'</div><div class="l">Inactivos 30d+</div></div>';
      h+='</div></div>';
      if(I.Platforms){
        h+='<div class="sc" style="margin-bottom:0"><div class="sc-title" style="font-size:12px">Plataformas</div>';
        const plats=Object.entries(I.Platforms).sort((a,b)=>b[1]-a[1]);
        const ptotal=plats.reduce((s,p)=>s+p[1],0)||1;
        plats.forEach(p=>{
          const pp=Math.round(p[1]/ptotal*100);
          h+='<div class="bar-row"><span class="bar-label">'+p[0]+'</span>';
          h+='<div class="bar-track"><div class="bar-fill" style="width:'+pp+'%;background:var(--accent)"></div></div>';
          h+='<span class="bar-value">'+p[1]+' ('+pp+'%)</span></div>';
        });
        h+='</div>';
      }
      h+='</div></div>';
    }

    // ═══════════════════════════════════════════════════════
    // Copilot
    // ═══════════════════════════════════════════════════════
    if(A.Copilot){
      const Co=A.Copilot;
      h+='<div class="sc"><div class="sc-title" style="display:flex;align-items:center;justify-content:space-between">&#129302; Microsoft 365 Copilot '+semDot(Co.AdoptionPct)+'</div>';
      h+=miniBar(Co.AdoptionPct,'Adopcion ('+Co.ActiveUsers30d+'/'+Co.LicensedUsers+' usuarios activos 30d)');
      h+='</div>';
    }

    pc.innerHTML=h;
  }
}

// ============================================================================
// RECOMMENDATIONS (all, paginated, filterable by pillar + status)
// ============================================================================
if(S&&S.AllRecommendations&&S.AllRecommendations.length){
  const recs=S.AllRecommendations;
  const PS=10;
  let rPage=1,rCat='all',rSt='all',rSvc='all',rSvcSub='all',rQ='';
  const categories=[...new Set(recs.map(r=>r.Category))].sort();
  const statuses=[...new Set(recs.map(r=>r.ImplementationStatus))].sort();
  // Product groups — cleaner than 13+ individual service tabs
  const svcGroups=[
    {key:'defender',label:'\u{1F6E1} Defender',svcs:['MDATP','MDO','Azure ATP','MCAS']},
    {key:'identity',label:'\u{1F510} Entra ID',svcs:['AzureAD']},
    {key:'m365',label:'\u{1F4E7} Microsoft 365',svcs:['EXO','MS Teams','SPO','FORMS','SWAY']},
    {key:'governance',label:'\u{1F4CB} Gobernanza',svcs:['MIP','AppG','Admincenter']}
  ];
  const defSubLabels={'MDATP':'Endpoint','MDO':'Office 365','Azure ATP':'Identity','MCAS':'Cloud Apps'};
  const catLabels={'Apps':'Aplicaciones','Data':'Datos','Device':'Dispositivos','Identity':'Identidad'};
  function svcsForGroup(gk){const g=svcGroups.find(x=>x.key===gk);return g?g.svcs:[];}

  // Build product group tabs
  const svcEl=document.getElementById('recSvcTabs');
  if(svcEl){
    let t='<button class="tab active" onclick="window._recGrp(\'all\',this)">Todos</button>';
    svcGroups.forEach(g=>{
      const cnt=recs.filter(r=>g.svcs.includes(r.Service)).length;
      t+='<button class="tab" onclick="window._recGrp(\''+g.key+'\',this)">'+g.label+' <span style="font-size:9px;opacity:.6">('+cnt+')</span></button>';
    });
    svcEl.innerHTML=t;
  }
  // Show/hide defender sub-tabs
  function showDefSubs(show){
    const subEl=document.getElementById('recSubTabs');
    if(!subEl)return;
    if(!show){subEl.style.display='none';return;}
    const defSvcs=svcsForGroup('defender').filter(s=>recs.some(r=>r.Service===s));
    let t='<button class="tab active" onclick="window._recSub(\'all\',this)">Todos</button>';
    defSvcs.forEach(s=>{
      const cnt=recs.filter(r=>r.Service===s).length;
      t+='<button class="tab" onclick="window._recSub(\''+s+'\',this)">'+(defSubLabels[s]||s)+' <span style="font-size:9px;opacity:.6">('+cnt+')</span></button>';
    });
    subEl.innerHTML=t;
    subEl.style.display='';
  }
  // Build category tabs (renamed for clarity)
  const tabEl=document.getElementById('recTabs');
  if(tabEl){
    let t='<button class="tab active" onclick="window._recCat(\'all\',this)">Todos</button>';
    categories.forEach(c=>{t+='<button class="tab" onclick="window._recCat(\''+c+'\',this)">'+(catLabels[c]||esc(c))+'</button>';});
    tabEl.innerHTML=t;
  }
  // Build status tabs
  const stEl=document.getElementById('recStatusTabs');
  if(stEl){
    let t='<button class="tab active" onclick="window._recSt(\'all\',this)">Todos</button>';
    statuses.forEach(s=>{
      const label=s==='Implemented'?'Implementado':s==='Partial'?'Parcial':'Pendiente';
      t+='<button class="tab" onclick="window._recSt(\''+s+'\',this)">'+label+'</button>';
    });
    stEl.innerHTML=t;
  }

  function getRecFiltered(){
    return recs.filter(r=>{
      if(rSvc!=='all'){
        const grp=svcGroups.find(g=>g.key===rSvc);
        if(grp&&!grp.svcs.includes(r.Service))return false;
        if(!grp&&r.Service!==rSvc)return false;
      }
      if(rSvcSub!=='all'&&r.Service!==rSvcSub)return false;
      if(rCat!=='all'&&r.Category!==rCat)return false;
      if(rSt!=='all'&&r.ImplementationStatus!==rSt)return false;
      if(rQ&&!(r.Title+' '+r.Category+' '+r.Service).toLowerCase().includes(rQ))return false;
      return true;
    });
  }

  function renderRecs(){
    const f=getRecFiltered(),tp=Math.max(1,Math.ceil(f.length/PS));
    if(rPage>tp)rPage=tp;
    const s=(rPage-1)*PS,e=Math.min(s+PS,f.length);
    let h='';
    for(let i=s;i<e;i++){
      const r=f[i],imp=r.Improvement;
      const ic=imp>=5?'var(--red)':imp>=2?'var(--yellow)':'var(--muted)';
      const sb=r.ImplementationStatus==='Implemented'?badge('OK','green'):r.ImplementationStatus==='Partial'?badge('Parcial','yellow'):badge('Pendiente','red');
      h+='<tr><td class="tc">'+(i+1)+'</td><td class="tc" style="font-weight:700;color:'+ic+'">+'+imp+'</td>';
      h+='<td>'+esc(r.Title)+'</td><td class="small">'+esc(r.Category)+'</td><td class="small muted">'+esc(r.Service)+'</td>';
      h+='<td class="tc">'+sb+'</td></tr>';
    }
    document.getElementById('recTbody').innerHTML=h;
    const txt='Mostrando '+(e-s)+' de '+f.length+' recomendaciones ('+recs.length+' total)';
    ['recCount','recCountBot'].forEach(id=>{const el=document.getElementById(id);if(el)el.textContent=txt;});
    ['recPgTop','recPgBot'].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML=pgHtml(rPage,tp,'r');});
  }

  window._pg.r=p=>{rPage=p;renderRecs();};
  window._recFilter=()=>{rQ=document.getElementById('recSearch').value.toLowerCase();rPage=1;renderRecs();};
  window._recCat=(c,btn)=>{rCat=c;rPage=1;document.querySelectorAll('#recTabs .tab').forEach(t=>t.classList.remove('active'));if(btn)btn.classList.add('active');renderRecs();};
  window._recSt=(s,btn)=>{rSt=s;rPage=1;document.querySelectorAll('#recStatusTabs .tab').forEach(t=>t.classList.remove('active'));if(btn)btn.classList.add('active');renderRecs();};
  // Product group click
  window._recGrp=(g,btn)=>{
    rSvc=g;rSvcSub='all';rPage=1;
    document.querySelectorAll('#recSvcTabs .tab').forEach(t=>t.classList.remove('active'));
    if(btn)btn.classList.add('active');
    showDefSubs(g==='defender');
    renderRecs();
  };
  // Defender sub-product click
  window._recSub=(s,btn)=>{
    rSvcSub=s;rPage=1;
    document.querySelectorAll('#recSubTabs .tab').forEach(t=>t.classList.remove('active'));
    if(btn)btn.classList.add('active');
    renderRecs();
  };
  // Called from Postura cards — auto-navigate to recomendaciones + select correct group & sub
  window._recSvc=(rawSvc,btn)=>{
    // Find which group this service belongs to
    const grp=svcGroups.find(g=>g.svcs.includes(rawSvc));
    rSvc=grp?grp.key:'all';
    rSvcSub=rawSvc;
    rPage=1;
    // Highlight group tab
    document.querySelectorAll('#recSvcTabs .tab').forEach(t=>t.classList.remove('active'));
    if(grp){document.querySelectorAll('#recSvcTabs .tab').forEach(t=>{if(t.textContent.includes(grp.label.slice(2)))t.classList.add('active');});}
    // Show defender sub-tabs and select the right one
    if(grp&&grp.key==='defender'){
      showDefSubs(true);
      document.querySelectorAll('#recSubTabs .tab').forEach(t=>{t.classList.remove('active');if(t.textContent.includes(defSubLabels[rawSvc]||rawSvc))t.classList.add('active');});
    } else { showDefSubs(false); }
    // Switch to recomendaciones tab
    document.querySelectorAll('.main-tab').forEach(t=>{t.classList.remove('active');if(t.dataset.tab==='recomendaciones')t.classList.add('active');});
    document.querySelectorAll('.tab-panel').forEach(p=>p.style.display='none');
    const rp=document.getElementById('tab-recomendaciones');if(rp)rp.style.display='block';
    renderRecs();
  };
  renderRecs();
}

// ============================================================================
// WASTE TABLE (paginated)
// ============================================================================
(function(){
  if(!_WD||!_WD.length)return;
  const PS=10;
  let page=1,q='';
  const idx=_WD.map(w=>(w[0]+' '+w[1]+' '+w[4]+' '+w[5]).toLowerCase());

  function getF(){return _WD.filter((_,i)=>!q||idx[i].includes(q));}

  function renderW(){
    const f=getF(),tp=Math.max(1,Math.ceil(f.length/PS));
    if(page>tp)page=tp;
    const s=(page-1)*PS,e=Math.min(s+PS,f.length);
    let h='';
    for(let i=s;i<e;i++){
      const w=f[i];
      const dot=w[2]?'<span class="dot dot-on"></span>':'<span class="dot dot-off"></span>';
      const reasons=w[5].split('|').map(r=>{
        r=r.trim();
        if(r.includes('Disabled'))return badge(r,'red');
        if(r.includes('Inactive'))return badge(r,'yellow');
        if(r.includes('Duplicate'))return badge(r,'blue');
        return badge(r,'default');
      }).join(' ');
      h+='<tr><td>'+esc(w[0])+'</td><td class="small muted">'+esc(w[1])+'</td>';
      h+='<td class="tc">'+dot+'</td><td class="tc small">'+esc(w[3])+'</td>';
      h+='<td class="small sku-cell" title="'+esc(w[4]).replace(/\s*\|\s*/g,'\\n')+'">'+esc(w[4])+'</td>';
      h+='<td class="small">'+reasons+'</td></tr>';
    }
    document.getElementById('wasteTbody').innerHTML=h;
    const txt='Mostrando '+(e-s)+' de '+f.length;
    const el=document.getElementById('wasteCount');if(el)el.textContent=txt;
    ['wastePgTop','wastePgBot'].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML=pgHtml(page,tp,'w');});
  }

  window._pg.w=p=>{page=p;renderW();};
  window._wasteFilter=()=>{q=document.getElementById('wasteSearch').value.toLowerCase();page=1;renderW();};
  renderW();
})();

// ============================================================================
// DUPLICATES TABLE (paginated)
// ============================================================================
(function(){
  if(!_DD||!_DD.length)return;
  const PS=10;
  let page=1,q='';
  const idx=_DD.map(d=>(d[0]+' '+d[1]+' '+d[2]+' '+d[3]).toLowerCase());

  function getF(){return _DD.filter((_,i)=>!q||idx[i].includes(q));}

  function renderD(){
    const f=getF(),tp=Math.max(1,Math.ceil(f.length/PS));
    if(page>tp)page=tp;
    const s=(page-1)*PS,e=Math.min(s+PS,f.length);
    let h='';
    for(let i=s;i<e;i++){
      const d=f[i];
      const skus=d[3].split('|').map(s=>badge(s.trim(),'blue')).join(' ');
      h+='<tr><td>'+esc(d[0])+'</td><td class="small muted">'+esc(d[1])+'</td>';
      h+='<td>'+badge(d[2],'yellow')+'</td>';
      h+='<td class="small">'+skus+'</td></tr>';
    }
    const tbody=document.getElementById('dupTbody');if(tbody)tbody.innerHTML=h;
    const txt='Mostrando '+(e-s)+' de '+f.length;
    const el=document.getElementById('dupCount');if(el)el.textContent=txt;
    ['dupPgTop','dupPgBot'].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML=pgHtml(page,tp,'d');});
  }

  window._pg.d=p=>{page=p;renderD();};
  window._dupFilter=()=>{q=document.getElementById('dupSearch').value.toLowerCase();page=1;renderD();};
  renderD();
})();

// ============================================================================
// USERS TABLE (paginated, filterable)
// ============================================================================
(function(){
  if(!_UD||!_UD.length)return;
  const PS=_UD.length>5000?25:10,tbody=document.getElementById('userTbody');
  let page=1,statusF='all',deptF='all',searchQ='',skuFilterVals=[],skuMatchAll=false,initialized=false;
  const searchIdx=_UD.map(u=>(u[0]+' '+u[1]+' '+u[2]+' '+u[3]).toLowerCase());

  function getFiltered(){
    const res=[];
    for(let i=0;i<_UD.length;i++){
      const u=_UD[i],st=u[6];
      if(searchQ&&!searchIdx[i].includes(searchQ))continue;
      if(statusF==='active'&&(st&3))continue;
      if(statusF==='disabled'&&!(st&1))continue;
      if(statusF==='inactive'&&!(st&2))continue;
      if(statusF==='waste'&&!(st&4))continue;
      if(statusF==='disPlans'&&!(st&16))continue;
      if(deptF!=='all'&&u[2]!==deptF)continue;
      if(skuFilterVals.length>0){
        const us=u[3];
        if(skuMatchAll){if(!skuFilterVals.every(s=>us.indexOf(s)!==-1))continue;}
        else{if(!skuFilterVals.some(s=>us.indexOf(s)!==-1))continue;}
      }
      res.push(u);
    }
    return res;
  }

  function renderRow(u){
    const st=u[6],mt=u[7],ca=u[8],cats=u[9],wf=u[10],dp=u[11];
    let cls=(st&1)?'row-disabled':((st&2)?'row-inactive':'');
    let badges=ca?'<span class="badge b-green" style="font-size:9px">CA</span> ':'';
    badges+=mt===1?'<span class="badge b-blue" style="font-size:9px">G</span>':mt===2?'<span class="badge b-yellow" style="font-size:9px">G+D</span>':'<span class="badge b-default" style="font-size:9px">D</span>';
    if(st&8)badges+=' <span class="badge b-yellow" style="font-size:9px" title="Planes deshab.: '+esc(dp)+'">&#9888;</span>';
    if(st&4)badges+=' <span class="dot dot-off" title="'+esc(wf)+'"></span>';
    let c='<td class="sticky-name">'+esc(u[0])+' '+badges+'</td>';
    c+='<td class="small muted">'+esc(u[1])+'</td>';
    c+='<td class="small">'+esc(u[2])+'</td>';
    c+='<td class="small sku-cell" title="'+esc(u[3]).replace(/\s*\|\s*/g,'\\n')+'">'+esc(u[3])+'</td>';
    c+='<td class="tc small">'+esc(u[4])+'</td>';
    const days=u[5];
    c+='<td class="tc small">'+(days<0?'<span class="muted">N/A</span>':days)+'</td>';
    for(let i=0;i<cats.length;i++){
      const v=cats[i],fn=_CM[i]||'';
      const sep=_CS.includes(i)?' td-sep':'';
      const title=v===1?fn+' - Activo':v===-1?fn+' - Deshabilitado':fn+' - No disponible';
      const dot=v===1?'dot-on':v===-1?'dot-off':'dot-na';
      c+='<td class="tc'+sep+'"><span class="dot '+dot+'" title="'+esc(title)+'"></span></td>';
    }
    return'<tr class="'+cls+'">'+c+'</tr>';
  }

  function render(){
    const f=getFiltered(),tp=Math.max(1,Math.ceil(f.length/PS));
    if(page>tp)page=tp;
    const s=(page-1)*PS,e=Math.min(s+PS,f.length);
    let h='';for(let i=s;i<e;i++)h+=renderRow(f[i]);
    tbody.innerHTML=h;
    const shown=e-s,total=_UD.length,ft=f.length;
    const txt=ft===total?'Mostrando '+shown+' de '+total+' usuarios':'Mostrando '+shown+' de '+ft+' filtrados ('+total+' total)';
    ['userCount','userCountBottom'].forEach(id=>{const el=document.getElementById(id);if(el)el.textContent=txt;});
    ['pgTop','pgBottom'].forEach(id=>{const el=document.getElementById(id);if(el)el.innerHTML=pgHtml(page,tp,'u');});
  }

  window._pg.u=p=>{page=p;render();};
  window.applyFilters=()=>{searchQ=document.getElementById('userSearch').value.toLowerCase();page=1;render();};
  window.applySkuFilter=()=>{skuMatchAll=document.getElementById('skuMatchAll').checked;page=1;render();};
  window.toggleSkuFilter=(btn,sku)=>{
    const idx=skuFilterVals.indexOf(sku);
    if(idx===-1){skuFilterVals.push(sku);btn.classList.add('active');}else{skuFilterVals.splice(idx,1);btn.classList.remove('active');}
    document.querySelector('#skuChips .tab:first-child').classList.toggle('active',skuFilterVals.length===0);
    skuMatchAll=document.getElementById('skuMatchAll').checked;page=1;render();
  };
  window.clearSkuFilter=()=>{skuFilterVals=[];document.querySelectorAll('#skuChips .tab').forEach((t,i)=>{t.classList.toggle('active',i===0);});page=1;render();};
  window.setStatusFilter=(f,btn)=>{statusF=f;page=1;document.querySelectorAll('#statusFilters .tab').forEach(t=>t.classList.remove('active'));if(btn)btn.classList.add('active');render();};
  window.filterByDept=d=>{deptF=d;page=1;document.querySelectorAll('#deptFilters .tab').forEach(t=>{t.classList.toggle('active',(d==='all'&&t.textContent==='Todos')||t.textContent===d);});render();};
  // Deferred init: render users table only when Usuarios tab is first opened
  window._initUsersTab=()=>{if(!initialized){initialized=true;render();}};
  // If usuarios tab is the default, render immediately
  if(document.getElementById('tab-usuarios')&&document.getElementById('tab-usuarios').classList.contains('active')){render();initialized=true;}
})();

// ============================================================================
// TAB NAVIGATION SYSTEM
// ============================================================================
function switchTab(tabId, btn) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.main-tab').forEach(t => t.classList.remove('active'));
  const panel = document.getElementById('tab-' + tabId);
  if (panel) { panel.classList.add('active'); panel.scrollIntoView({behavior:'smooth',block:'start'}); }
  if (btn) btn.classList.add('active');
  history.replaceState(null, '', '#' + tabId);
  // Lazy init: render users table on first visit
  if (tabId === 'usuarios' && window._initUsersTab) window._initUsersTab();
}

// Bind tab click handlers
document.querySelectorAll('.main-tab').forEach(btn => {
  btn.addEventListener('click', function() { switchTab(this.dataset.tab, this); });
});

// Handle URL hash on load
(function() {
  const hash = window.location.hash.replace('#', '');
  if (hash) {
    const btn = document.querySelector('.main-tab[data-tab="' + hash + '"]');
    if (btn) { setTimeout(() => switchTab(hash, btn), 100); }
  }
})();

// ============================================================================
// CSV EXPORT FUNCTIONS
// ============================================================================
function downloadCSV(csv, filename) {
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}

function exportTableToCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  let csv = '\ufeff';
  table.querySelectorAll('tr').forEach(row => {
    const cols = row.querySelectorAll('th, td');
    const rowData = Array.from(cols).map(col => {
      let text = col.innerText.replace(/"/g, '""').replace(/[\r\n]+/g, ' ').trim();
      return '"' + text + '"';
    });
    csv += rowData.join(',') + '\r\n';
  });
  downloadCSV(csv, filename);
}

function exportWasteCSV() {
  let csv = '\ufeff"Nombre","UPN","Activa","Ultimo Sign-In","SKUs","Motivo"\r\n';
  _WD.forEach(w => {
    csv += '"' + [w[0],w[1],w[2]?'Si':'No',w[3],w[4],w[5]].map(v => String(v||'').replace(/"/g,'""')).join('","') + '"\r\n';
  });
  downloadCSV(csv, 'Licencias_Revisar.csv');
}

function exportRecsCSV() {
  if (!S || !S.AllRecommendations) return;
  let csv = '\ufeff"Impacto","Recomendacion","Categoria","Servicio","Estado"\r\n';
  S.AllRecommendations.forEach(r => {
    csv += '"' + ['+'+r.Improvement,r.Title,r.Category,r.Service,r.ImplementationStatus].map(v => String(v||'').replace(/"/g,'""')).join('","') + '"\r\n';
  });
  downloadCSV(csv, 'Recomendaciones.csv');
}

function exportDupsCSV() {
  let csv = '\ufeff"Nombre","UPN","Producto Duplicado","Provisto por SKUs"\r\n';
  _DD.forEach(d => {
    csv += '"' + [d[0],d[1],d[2],d[3]].map(v => String(v||'').replace(/"/g,'""')).join('","') + '"\r\n';
  });
  downloadCSV(csv, 'Duplicados.csv');
}

function exportUsersCSV() {
  const h = ['Nombre','UPN','Departamento','SKUs','Ultimo Sign-In','Dias','Estado','Metodo','CA'];
  _CM.forEach(c => h.push(c));
  let csv = '\ufeff' + h.map(x => '"'+x+'"').join(',') + '\r\n';
  _UD.forEach(u => {
    const st=u[6],mt=u[7],ca=u[8],cats=u[9];
    const estado=(st&1)?'Deshabilitado':(st&2)?'Inactivo':'Activo';
    const metodo=mt===1?'Grupo':mt===2?'Mixto':'Directa';
    const row=[u[0],u[1],u[2],u[3],u[4],u[5]<0?'N/A':u[5],estado,metodo,ca?'Si':'No'];
    cats.forEach(v => row.push(v===1?'Activo':v===-1?'Deshabilitado':'N/A'));
    csv += row.map(v => '"'+String(v||'').replace(/"/g,'""')+'"').join(',') + '\r\n';
  });
  downloadCSV(csv, 'Usuarios_Detalle.csv');
}

</script>
</body>
</html>
"@

$Html | Out-File -FilePath $HtmlPath -Encoding UTF8
Write-Host "[OK] Reporte: $HtmlPath" -ForegroundColor Green
