@echo off
set "BUILD_DIR=%~dp0"
set "BAT_FILE=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:BAT_FILE,[Text.Encoding]::Default); $i=$s.IndexOf('#'+'PS_START'); Invoke-Expression $s.Substring($i)"
set "RC=%errorlevel%"
if not "%RC%"=="0" pause
exit /b %RC%

#PS_START
$ErrorActionPreference = "Stop"

$ProjectFolder = $env:BUILD_DIR.TrimEnd('\')
$OutputName    = "Term02Tool.xlsm"
$OutputFile    = Join-Path $ProjectFolder $OutputName
$xlOpenXMLWorkbookMacroEnabled = 52

# Exact name of the panel module of this project (this copy builds Term02Tool.xlsm)
$PanelRoot = "modBuildPanel02"
# Any panel-module reference inside ThisWorkbook.cls is rewritten to the real name
$PanelRefPattern = 'modBuild(?:Bac)?Panel\w*'
$PanelProc = "RefreshPanelState"
$BuildPanelProc = "BuildControlPanel"

# Panel sheet detection: any sheet whose name contains one of these (Arabic built from char codes)
$ArabicPanel = -join ([char]0x0644,[char]0x0648,[char]0x062D,[char]0x0629)   # لوحة
$PanelSheetPatterns = @("*Panel*", "*$ArabicPanel*")

# Fallback ThisWorkbook code (used only if ThisWorkbook.cls is not found)
$DefaultThisWorkbookCode = @'
Option Explicit

Private Sub Workbook_Activate()
    On Error Resume Next
    modBuildPanel.RefreshPanelState
    On Error GoTo 0
End Sub
'@

$Excel = $null; $Workbook = $null; $VBProject = $null
$Success = $false
$HasRefresh = $true

# Read a document-module .cls and strip the export header (ANSI code page)
function Get-DocModuleCode([string]$Path) {
    $lines = [IO.File]::ReadAllLines($Path, [Text.Encoding]::Default)
    $i = 0
    while ($i -lt $lines.Count -and
           $lines[$i] -cmatch '^(VERSION |BEGIN|END\s*$|\s+MultiUse|Attribute VB_)') { $i++ }
    if ($i -ge $lines.Count) { return "" }
    return ($lines[$i..($lines.Count - 1)] -join "`r`n")
}

# ---------------------------------------------------------------------------
# Import diagnostics: runs only when a module import fails
# ---------------------------------------------------------------------------
$DiagLog = New-Object System.Collections.Generic.List[string]

function Write-Diag([string]$Text, [string]$Color = "Gray") {
    Write-Host $Text -ForegroundColor $Color
    $DiagLog.Add($Text)
}

function Show-ImportDiagnostics($File, $VBProject, $Err) {
    Write-Diag ""
    Write-Diag "===== IMPORT DIAGNOSTICS: $($File.Name) =====" "Yellow"

    # 1) Real error chain (the HRESULT is usually in the inner exception)
    $e = $Err.Exception
    while ($e) {
        Write-Diag ("Error : [{0}] 0x{1:X8} {2}" -f $e.GetType().Name, $e.HResult, $e.Message) "Red"
        $e = $e.InnerException
    }

    # 2) File format checks
    $b = [IO.File]::ReadAllBytes($File.FullName)
    $warn = New-Object System.Collections.Generic.List[string]
    if ($b.Length -eq 0) { $warn.Add("The file is EMPTY.") }
    if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { $warn.Add("UTF-8 BOM at the start of the file (VBA expects plain ANSI).") }
    if ($b.Length -ge 2 -and (($b[0] -eq 0xFF -and $b[1] -eq 0xFE) -or ($b[0] -eq 0xFE -and $b[1] -eq 0xFF))) { $warn.Add("UTF-16 file (VBA expects ANSI).") }
    if ($b -contains 0) { $warn.Add("NUL bytes found (UTF-16 or corrupted file).") }

    $highBytes = 0
    foreach ($x in $b) { if ($x -gt 127) { $highBytes++ } }
    $isUtf8 = $false
    if ($highBytes -gt 0) {
        try {
            $strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
            [void]$strict.GetString($b)
            $isUtf8 = $true
        } catch {}
        if ($isUtf8) { $warn.Add("File is UTF-8 with $highBytes non-ASCII bytes; VBA reads ANSI, so non-English text (Arabic) can be corrupted.") }
    }
    if ($isUtf8) { $enc = [System.Text.Encoding]::UTF8 } else { $enc = [System.Text.Encoding]::Default }
    $txt = $enc.GetString($b).TrimStart([char]0xFEFF)

    $crlf = [regex]::Matches($txt, "\r\n").Count
    $lf   = [regex]::Matches($txt, "(?<!\r)\n").Count
    $cr   = [regex]::Matches($txt, "\r(?!\n)").Count
    if ($lf -gt 0 -or $cr -gt 0) { $warn.Add("Wrong line endings: CRLF=$crlf, LF-only=$lf, CR-only=$cr (VBA expects CRLF only).") }

    $lines = $txt -split "\r\n|\n|\r"
    $first = ""
    foreach ($l in $lines) { if ($l.Trim() -ne "") { $first = $l; break } }
    if ($File.Extension -ieq ".bas" -and $lines[0] -notmatch '^Attribute\s+VB_Name\s*=') {
        $warn.Add("First line is not 'Attribute VB_Name = ...' (it is: '$($lines[0])').")
    }

    $fileBase = [IO.Path]::GetFileNameWithoutExtension($File.Name)
    $vbName = $fileBase
    $m = [regex]::Match($txt, '(?im)^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"')
    if (-not $m.Success) {
        $warn.Add("No 'Attribute VB_Name' line found.")
    } else {
        $vbName = $m.Groups[1].Value
        if ($vbName -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') { $warn.Add("VB_Name '$vbName' is not a valid VBA module name.") }
        foreach ($c in $VBProject.VBComponents) {
            if ($c.Name -ieq $vbName) { $warn.Add("A component named '$vbName' ALREADY EXISTS in the workbook (name conflict).") }
        }
        if ($VBProject.Name -ieq $vbName) { $warn.Add("VB_Name equals the VBA project name.") }
    }

    $longLines = 0
    foreach ($l in $lines) { if ($l.Length -gt 1023) { $longLines++ } }
    if ($longLines -gt 0) { $warn.Add("$longLines line(s) longer than 1023 characters (VBA limit).") }

    $midAttr = @()
    for ($k = 5; $k -lt $lines.Count; $k++) { if ($lines[$k] -match '^\s*Attribute\s+VB_') { $midAttr += ($k + 1) } }
    if ($midAttr.Count -gt 0) { $warn.Add("Attribute lines found in the middle of the file at line(s): $($midAttr -join ', ').") }

    $encName = "ANSI"
    if ($isUtf8) { $encName = "UTF-8" }
    Write-Diag ("File  : {0} bytes, {1} lines, looks like {2}" -f $b.Length, $lines.Count, $encName)
    Write-Diag ("Names : file = {0} | VB_Name = {1}" -f $fileBase, $vbName)
    Write-Diag ("First : {0}" -f $first)
    if ($warn.Count -eq 0) { Write-Diag "Format checks: nothing suspicious found." "Green" }
    foreach ($w in $warn) { Write-Diag "SUSPECT: $w" "Yellow" }

    # 3) Probe A: import a normalized copy (ANSI, CRLF, no BOM)
    $tmp = Join-Path $env:TEMP ("diag_" + $File.Name)
    $okA = $false
    try {
        $norm = [regex]::Replace($txt, "\r\n|\n|\r", "`r`n")
        [IO.File]::WriteAllText($tmp, $norm, [System.Text.Encoding]::Default)
        $comp = $VBProject.VBComponents.Import($tmp)
        $okA = $true
        try { $VBProject.VBComponents.Remove($comp) } catch {}
    } catch {
        Write-Diag ("Probe A: a normalized copy ALSO fails -> {0}" -f $_.Exception.Message) "Yellow"
    } finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }

    $contentProblem = $false
    if ($okA) {
        Write-Diag "Probe A: a normalized copy (ANSI + CRLF + no BOM) imports OK -> the cause is the FILE FORMAT, not the code." "Green"
        Write-Diag "         Fix: save $($File.Name) as ANSI with CRLF line endings." "Green"
    } elseif ($File.Extension -ieq ".bas") {
        # 4) Probe B: import an empty module with the same VB_Name (name problem vs content problem)
        $stub = Join-Path $env:TEMP "diag_stub.bas"
        try {
            [IO.File]::WriteAllText($stub, "Attribute VB_Name = `"$vbName`"`r`nOption Explicit`r`n", [System.Text.Encoding]::Default)
            $comp = $VBProject.VBComponents.Import($stub)
            try { $VBProject.VBComponents.Remove($comp) } catch {}
            Write-Diag "Probe B: an empty module named '$vbName' imports OK -> the NAME is fine; the problem is inside the file CONTENT." "Yellow"
            Write-Diag "         Check the Attribute lines at the top and any unusual lines/characters in the file." "Yellow"
            $contentProblem = $true
        } catch {
            Write-Diag "Probe B: even an empty module named '$vbName' fails -> the problem is the module NAME (conflict/invalid)." "Red"
        } finally {
            Remove-Item $stub -Force -ErrorAction SilentlyContinue
        }
    }

    # 5) Probe C: find the first line whose inclusion makes the import fail (binary search on file prefixes)
    if ($contentProblem) {
        $testPrefix = {
            param([int]$n, [int]$skip = 0)
            $p = Join-Path $env:TEMP "diag_prefix.bas"
            $ok = $false
            try {
                $part = New-Object System.Collections.Generic.List[string]
                for ($i = 0; $i -lt $n; $i++) { if (($i + 1) -ne $skip) { $part.Add($lines[$i]) } }
                [IO.File]::WriteAllText($p, (($part -join "`r`n") + "`r`n"), [System.Text.Encoding]::Default)
                $c = $VBProject.VBComponents.Import($p)
                $ok = $true
                try { $VBProject.VBComponents.Remove($c) } catch {}
            } catch {}
            Remove-Item $p -Force -ErrorAction SilentlyContinue
            return $ok
        }
        $total = $lines.Count
        $lo = [Math]::Min(2, $total)
        if (-not (& $testPrefix $lo)) {
            Write-Diag "Probe C: even the first $lo line(s) fail -> the problem is in the file header." "Red"
        } else {
            $hi = $total
            while (($hi - $lo) -gt 1) {
                $mid = [int][Math]::Floor(($lo + $hi) / 2)
                if (& $testPrefix $mid) { $lo = $mid } else { $hi = $mid }
            }
            $bytes = [System.Text.Encoding]::Default.GetByteCount(($lines[0..($hi - 1)] -join "`r`n"))
            Write-Diag "Probe C: the import first FAILS when line $hi is included (lines 1-$lo import fine)." "Red"
            Write-Diag ("         Size of the file up to that line: {0} bytes" -f $bytes) "Red"
            $from = [Math]::Max(1, $hi - 3)
            $to = [Math]::Min($total, $hi + 1)
            for ($k = $from; $k -le $to; $k++) {
                $mark = "  "
                if ($k -eq $hi) { $mark = ">>" }
                $t = $lines[$k - 1]
                if ($t.Length -gt 160) { $t = $t.Substring(0, 160) + "..." }
                Write-Diag ("   {0} {1,5}: {2}" -f $mark, $k, $t) "Red"
            }
            for ($k = $hi - 1; $k -ge 0; $k--) {
                if ($lines[$k] -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)') {
                    Write-Diag ("         Inside procedure '{0}' which starts at line {1}" -f $Matches[5], ($k + 1)) "Red"
                    break
                }
            }
            if (& $testPrefix $total $hi) {
                Write-Diag "Probe D: the WHOLE file imports OK without line $hi -> that line alone is the problem." "Green"
            } else {
                Write-Diag "Probe D: still fails without line $hi -> a size limit or several problem lines (compare the size above)." "Yellow"
            }
        }
    }

    $logPath = Join-Path $ProjectFolder "ImportDiagnostics.txt"
    try { [IO.File]::WriteAllLines($logPath, $DiagLog.ToArray(), [System.Text.Encoding]::UTF8); Write-Diag "Log saved: $logPath" } catch {}
}

# ---------------------------------------------------------------------------
# .frm without .frx: create an empty UserForm and paste the code into it
# (automates the manual "add a form + paste the frm content" step)
# ---------------------------------------------------------------------------
function Import-FormManually($File, $VBProject) {
    $lines = [IO.File]::ReadAllLines($File.FullName, [Text.Encoding]::Default)
    $name = [IO.Path]::GetFileNameWithoutExtension($File.Name)
    foreach ($l in $lines) {
        if ($l -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"') { $name = $Matches[1]; break }
    }
    $i = 0
    while ($i -lt $lines.Count -and ($lines[$i] -match '^\s*$' -or $lines[$i] -match '^\s*VERSION\s')) { $i++ }
    $caption = $null
    $cw = 0; $ch = 0; $sup = -1
    $controls = 0
    if ($i -lt $lines.Count -and $lines[$i] -match '^\s*Begin\s') {
        $depth = 0
        while ($i -lt $lines.Count) {
            $l = $lines[$i]
            if ($l -match '^\s*Begin\s') { $depth++; if ($depth -gt 1) { $controls++ } }
            elseif ($l -match '^\s*End\s*$') { $depth-- }
            elseif ($depth -eq 1 -and $l -match '^\s*Caption\s*=\s*"(.*)"') { $caption = $Matches[1] }
            elseif ($depth -eq 1 -and $l -match '^\s*ClientWidth\s*=\s*(\d+)') { $cw = [double]$Matches[1] }
            elseif ($depth -eq 1 -and $l -match '^\s*ClientHeight\s*=\s*(\d+)') { $ch = [double]$Matches[1] }
            elseif ($depth -eq 1 -and $l -match '^\s*StartUpPosition\s*=\s*(\d+)') { $sup = [int]$Matches[1] }
            $i++
            if ($depth -eq 0) { break }
        }
    }
    while ($i -lt $lines.Count -and $lines[$i] -match '^\s*Attribute\s+VB_') { $i++ }
    $code = ""
    if ($i -lt $lines.Count) { $code = ($lines[$i..($lines.Count - 1)] -join "`r`n") }

    # The design-time size lives only in the FRM header, so it is restored at run time
    # inside UserForm_Initialize (outer size adjusted by the real window chrome, so the
    # inside area equals ClientWidth x ClientHeight exactly on any Windows theme / DPI).
    if ($cw -gt 0 -and $ch -gt 0) {
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $wPt = [Math]::Round($cw / 20, 2).ToString($inv)
        $hPt = [Math]::Round($ch / 20, 2).ToString($inv)
        $sizeCode = "    ' --- injected by XLSMBuilder: restore design-time size from the FRM header ---`r`n" +
                    "    Me.Width = Me.Width + ($wPt - Me.InsideWidth)`r`n" +
                    "    Me.Height = Me.Height + ($hPt - Me.InsideHeight)`r`n"
        $initRx = [regex]'(?im)^[ \t]*(?:Private\s+|Public\s+)?Sub\s+UserForm_Initialize\s*\(\s*\)[ \t]*\r?\n'
        $m = $initRx.Match($code)
        if ($m.Success) { $code = $code.Insert($m.Index + $m.Length, $sizeCode) }
        else { $code += "`r`nPrivate Sub UserForm_Initialize()`r`n" + $sizeCode + "End Sub`r`n" }
        Write-Host "  form size restored in code: inside ${wPt} x ${hPt} pt" -ForegroundColor DarkYellow
    }
    $comp = $VBProject.VBComponents.Add(3)     # 3 = vbext_ct_MSForm
    $comp.Name = $name
    if ($caption) { try { $comp.Properties.Item("Caption").Value = $caption } catch {} }
    # size from the header (twips -> points); outer size = client size + window chrome
    if ($cw -gt 0) { try { $comp.Properties.Item("Width").Value  = ($cw / 20) + 6  } catch { Write-Host "  (design-time Width not applied; run-time sizing is used)" -ForegroundColor DarkYellow } }
    if ($ch -gt 0) { try { $comp.Properties.Item("Height").Value = ($ch / 20) + 26 } catch { Write-Host "  (design-time Height not applied; run-time sizing is used)" -ForegroundColor DarkYellow } }
    if ($sup -ge 0) { try { $comp.Properties.Item("StartUpPosition").Value = $sup } catch {} }
    if ($comp.CodeModule.CountOfLines -gt 0) { $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines) }
    $comp.CodeModule.AddFromString($code)
    Write-Host "  (no FRX) form '$name' created and code pasted." -ForegroundColor DarkYellow
    if ($controls -gt 0) { Write-Warning "$($File.Name) declares $controls control(s) in the header; without an FRX they are not recreated (controls must be created by the code)." }
}

try {
    Write-Host "Starting Excel build..." -ForegroundColor Cyan

    $AllFiles = @(
        Get-ChildItem $ProjectFolder -Filter "*.bas" -File | Sort-Object Name
        Get-ChildItem $ProjectFolder -Filter "*.cls" -File | Sort-Object Name
        Get-ChildItem $ProjectFolder -Filter "*.frm" -File | Sort-Object Name
    )

    $HasThisWorkbookFile = [bool]($AllFiles | Where-Object { $_.Name -ieq "ThisWorkbook.cls" })
    $OtherFiles = @($AllFiles | Where-Object { $_.Name -ine "ThisWorkbook.cls" })

    Write-Host "Found $($OtherFiles.Count) module files."

    foreach ($f in $OtherFiles | Where-Object Extension -eq ".frm") {
        $frx = [IO.Path]::ChangeExtension($f.FullName, ".frx")
        if (-not (Test-Path $frx)) { Write-Host "Note: no FRX for $($f.Name) - the form will be created manually from the FRM code." -ForegroundColor DarkYellow }
    }

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.EnableEvents = $false

    $Workbook = $Excel.Workbooks.Add()

    try { $VBProject = $Workbook.VBProject }
    catch { throw "Enable 'Trust access to the VBA project object model' in Excel Trust Center > Macro Settings." }

    # 1) Import modules
    foreach ($File in $OtherFiles) {
        Write-Host "Importing: $($File.Name)"
        $frxPath = [IO.Path]::ChangeExtension($File.FullName, ".frx")
        if ($File.Extension -ieq ".frm" -and -not (Test-Path $frxPath)) {
            try { Import-FormManually $File $VBProject; continue }
            catch { Write-Host "Manual form creation failed: $_" -ForegroundColor Red; throw "Failed creating form from $($File.Name): $_" }
        }
        try { [void]$VBProject.VBComponents.Import($File.FullName) }
        catch {
            $importErr = $_
            try { Show-ImportDiagnostics $File $VBProject $importErr }
            catch { Write-Host "(diagnostics failed: $_)" -ForegroundColor DarkYellow }
            throw "Failed importing $($File.Name): $importErr"
        }
    }

    # 2) Detect the panel module by its fixed root
    $PanelName = $null
    $PanelMatches = @()
    foreach ($comp in $VBProject.VBComponents) {
        if ($comp.Name -ieq $PanelRoot) { $PanelMatches += $comp.Name }
    }
    $PanelMatches = @($PanelMatches | Sort-Object)

    if ($PanelMatches.Count -eq 0) {
        throw "Panel module '$PanelRoot' was not found among the imported files."
    } else {
        $PanelName = $PanelMatches[0]
        Write-Host "Panel module: $PanelName" -ForegroundColor Yellow
        if ($PanelMatches.Count -gt 1) {
            Write-Warning "Multiple panel modules found ($($PanelMatches -join ', ')). Using: $PanelName"
        }
        $pm = $VBProject.VBComponents.Item($PanelName).CodeModule
        $pText = $pm.Lines(1, $pm.CountOfLines)
        # RefreshPanelState is optional (not every panel module has it)
        $HasRefresh = ($pText -match "(?i)\b(Sub|Function)\s+$PanelProc\b")
        if (-not $HasRefresh) { Write-Host "Note: '$PanelProc' not found in $PanelName - Workbook_Activate hook skipped." -ForegroundColor DarkYellow }
        # Build procedure: BuildControlPanel, otherwise auto-detect a Build*/Create* entry point
        if ($pText -notmatch "(?i)\bPublic\s+Sub\s+$BuildPanelProc\b") {
            $subNames = @([regex]::Matches($pText, "(?im)^\s*(?:Public\s+)?Sub\s+(\w+)") | ForEach-Object { $_.Groups[1].Value })
            $cand = @($subNames | Where-Object { $_ -match "^(Build|Create|Make|Setup|Init)" })
            if ($cand.Count -ge 1) {
                $BuildPanelProc = $cand[0]
                Write-Host "Build procedure auto-detected: $BuildPanelProc" -ForegroundColor Yellow
                if ($cand.Count -gt 1) { Write-Warning "Several build candidates ($($cand -join ', ')). Using: $BuildPanelProc" }
            } else {
                throw "Procedure '$BuildPanelProc' not found in $PanelName. Public subs found: $($subNames -join ', ')"
            }
        }
    }

    # 3) Inject ThisWorkbook code into the existing document module
    $cm = $VBProject.VBComponents.Item($Workbook.CodeName).CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }

    if ($HasThisWorkbookFile) {
        Write-Host "ThisWorkbook: from ThisWorkbook.cls"
        $code = Get-DocModuleCode (Join-Path $ProjectFolder "ThisWorkbook.cls")
    } else {
        Write-Host "ThisWorkbook: built-in default code"
        if ($HasRefresh) { $code = $DefaultThisWorkbookCode } else { $code = "" }
    }

    # Replace the root (or any suffixed name) with the real module name
    if ($PanelName) {
        $code = [regex]::Replace($code, "\b$PanelRefPattern(?=\s*\.)", $PanelName)
        $openPattern = "(?ims)^\s*Private\s+Sub\s+Workbook_Open\b.*?^\s*End\s+Sub"
        $openMatch = [regex]::Match($code, $openPattern)
        if ($openMatch.Success) {
            $callPattern = "(?im)^\s*" + [regex]::Escape($PanelName) + "\s*\.\s*RefreshPanelState\s*\r?\n"
            $openCode = [regex]::Replace($openMatch.Value, $callPattern, "", 1)
            $code = $code.Substring(0, $openMatch.Index) + $openCode + $code.Substring($openMatch.Index + $openMatch.Length)
        }
        $activatePattern = "(?ims)^\s*Private\s+Sub\s+Workbook_Activate\b.*?^\s*End\s+Sub"
        $activateMatch = [regex]::Match($code, $activatePattern)
        if (-not $HasRefresh) {
            # no refresh procedure in this panel: nothing to hook on activation
        } elseif (-not $activateMatch.Success) {
            $code += "`r`nPrivate Sub Workbook_Activate()`r`n    On Error Resume Next`r`n    $PanelName.RefreshPanelState`r`n    On Error GoTo 0`r`nEnd Sub`r`n"
        } else {
            $activateCode = $activateMatch.Value
            $calls = [regex]::Matches($activateCode, $callPattern)
            if ($calls.Count -gt 1) {
                $kept = $calls[0].Value
                $tailStart = $calls[0].Index + $calls[0].Length
                $tail = [regex]::Replace($activateCode.Substring($tailStart), $callPattern, "")
                $activateCode = $activateCode.Substring(0, $calls[0].Index) + $kept + $tail
                $code = $code.Substring(0, $activateMatch.Index) + $activateCode + $code.Substring($activateMatch.Index + $activateMatch.Length)
            }
        }
    }
    if ($code.Trim().Length -gt 0) { $cm.AddFromString($code) }   # leave ThisWorkbook empty when there is nothing to inject

    # 4) Save
    $Workbook.SaveAs($OutputFile, $xlOpenXMLWorkbookMacroEnabled)
    $Success = $true

    Write-Host ""
    Write-Host "BUILD SUCCESSFUL" -ForegroundColor Green
    Write-Host "Output: $OutputFile"
}
catch {
    Write-Host "BUILD FAILED: $_" -ForegroundColor Red
    $global:LASTEXITCODE = 1
}
finally {
    if ($Workbook) { try { $Workbook.Close($false) } catch {} }
    if ($Excel)    { try { $Excel.Quit() } catch {} }

    foreach ($o in @($VBProject, $Workbook, $Excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($Success) {
    # Launch: visible Excel, macros enabled for this session only, events ON
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.UserControl = $true           # keep Excel open after the script ends
    $previousAutomationSecurity = $xl.AutomationSecurity
    try {
        $xl.AutomationSecurity = 1
        $xl.EnableEvents = $true
        $wb = $xl.Workbooks.Open($OutputFile)   # Workbook_Open runs here
        $buildMacro = "'$OutputName'!$PanelName.$BuildPanelProc"
        Write-Host "Running: $buildMacro" -ForegroundColor Yellow
        [void]$xl.Run($buildMacro)
        $wb.Save()
    } finally {
        $xl.AutomationSecurity = $previousAutomationSecurity
    }

    # Go to the panel sheet
    $target = $null
    foreach ($ws in $wb.Worksheets) {
        foreach ($p in $PanelSheetPatterns) {
            if ($ws.Name -like $p) { $target = $ws; break }
        }
        if ($target) { break }
    }
    if ($target) {
        $target.Activate()
        [void]$target.Range("A1").Select()
        Write-Host "Panel sheet: $($target.Name)" -ForegroundColor Yellow
    } else {
        Write-Warning "No panel sheet found (patterns: $($PanelSheetPatterns -join ', ')). Staying on the active sheet."
    }
    exit 0
} else {
    exit 1
}