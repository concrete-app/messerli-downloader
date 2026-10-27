<#
Exportiert die Dokumente aller Projekte aus FakturaE, indem es sich selbst durch die
Oberflaeche klickt.

Pro Projekt (ab dem in der Projektliste markierten):
  Ribbon "Dokumentenmanagement" klicken -> Fenster "Projekt-Eigenschaften" abwarten ->
  erste Dokumentzeile anklicken -> pro Dokument Enter (oeffnet es) -> FakturaE legt die
  entschluesselte Datei unter \\AKU-VS-002\Messerli\Daten\dbs\Temp\<User>\{...}\ ab ->
  Datei nach <Ziel>\<Projektnr> <Bezeichnung>\ kopieren -> geoeffnetes Programm schliessen
  -> "Schliessen" -> in der Projektliste eine Zeile runter -> naechstes Projekt.

Bereits exportierte Projekte (Ordner mit _export.csv) werden uebersprungen, das Skript
kann also nach einem Abbruch einfach neu gestartet werden.

Aufruf:
  powershell -ExecutionPolicy Bypass -File .\export_projekt.ps1 -Projekte 1 -MaxDoks 3
  powershell -ExecutionPolicy Bypass -File .\export_projekt.ps1
#>
param(
    [string]$Ziel = "$env:USERPROFILE\Documents\Concrete Export",
    [string]$TempRoot = "\\AKU-VS-002\Messerli\Daten\dbs\Temp\$env:USERNAME",
    [int]$Projekte = 0,          # 0 = bis zum Ende der Projektliste
    [int]$MaxDoks = 0,           # 0 = alle Dokumente pro Projekt
    [int]$DokuKnopfX = 530,      # Ribbon-Knopf "Dokumentenmanagement", relativ zur Ribbon-Leiste
    [int]$DokuKnopfY = 80,
    [int]$TimeoutSek = 30,
    [int]$ViewerWartenSek = 10,
    [int]$StartWarten = 15,
    [string]$DebugOrdner = "",   # wenn gesetzt: Bild der Liste vor jedem Doppelklick
    [switch]$AbMarkiert          # beim markierten statt beim obersten Projekt beginnen
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class W {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, int x, int y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vk);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [StructLayout(LayoutKind.Sequential)] public struct PT { public int X, Y; }
    [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(PT p);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    public static IntPtr TopAt(int x, int y) { var p = new PT(); p.X = x; p.Y = y; return GetAncestor(WindowFromPoint(p), 2); }
    public static List<IntPtr> VisibleWindows() {
        var list = new List<IntPtr>();
        EnumWindows((h, l) => { if (IsWindowVisible(h)) list.Add(h); return true; }, IntPtr.Zero);
        return list;
    }
}
"@
[void][W]::SetProcessDPIAware()

$UIA = [System.Windows.Automation.AutomationElement]
$Raw = [System.Windows.Automation.TreeWalker]::RawViewWalker

function Fokus($hwnd) {
    # ALT antippen, damit Windows SetForegroundWindow erlaubt
    [W]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
    [W]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
    [void][W]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 400
}

# Esc (egal in welchem Fenster) bricht den Lauf ab; das Skript selbst sendet nie Esc
function PruefeEsc {
    if ([W]::GetAsyncKeyState(0x1B) -band 0x8001) { throw "ABBRUCH_ESC" }
}

function Klick([int]$x, [int]$y) {
    [void][W]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 100
    [W]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero)
    [W]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 300
}

function DoppelKlick([int]$x, [int]$y) {
    [void][W]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 100
    foreach ($n in 1, 2) {
        [W]::mouse_event(0x02, 0, 0, 0, [UIntPtr]::Zero)
        [W]::mouse_event(0x04, 0, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 60
    }
    Start-Sleep -Milliseconds 300
}

# y-Mitte der blau markierten Zeile in der Liste (Bildschirmkoordinate), sonst $null
function MarkierteZeile($zeilen) {
    $r = $zeilen.Current.BoundingRectangle
    if ($r.IsEmpty -or $r.Width -lt 100 -or $r.Height -lt 20) { return $null }
    $bmp = New-Object System.Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size)
    # eine Pixelzeile gilt als markiert, wenn mindestens 30 % der Messpunkte ueber die ganze
    # Breite Markierungsblau sind (Symbole und weisse Schrift unterbrechen das Blau);
    # gesucht ist ein zusammenhaengender Balken von mindestens 10 px (eine Zeile ist ~18 px)
    $xs = @(for ($x = 2; $x -lt $bmp.Width - 2; $x += 6) { $x })
    $ergebnis = $null; $von = -1
    for ($y = 0; $y -le $bmp.Height; $y++) {
        $treffer = 0
        if ($y -lt $bmp.Height) {
            foreach ($x in $xs) {
                $c = $bmp.GetPixel($x, $y)
                if ($c.R -lt 60 -and $c.G -gt 90 -and $c.G -lt 150 -and $c.B -gt 180) { $treffer++ }
            }
        }
        $blau = $treffer -ge [Math]::Max(3, [int]($xs.Count * 0.3))
        if ($blau -and $von -lt 0) { $von = $y }
        if (-not $blau -and $von -ge 0) {
            if ($y - $von -ge 10) { $ergebnis = [int]($r.Y + ($von + $y) / 2); break }
            $von = -1
        }
    }
    $g.Dispose(); $bmp.Dispose()
    return $ergebnis
}

# Programme, die nie beendet werden (nur deren neue Fenster werden geschlossen)
$Geschuetzt = @("Fakt", "OUTLOOK", "explorer", "WindowsTerminal", "powershell", "pwsh", "claude", "firefox", "SnippingTool", "Teams", "ms-teams")

function ProzessName([int]$id) { try { (Get-Process -Id $id -ErrorAction Stop).ProcessName } catch { "" } }

# schliesst Fenster anderer Programme, die ueber dem Punkt (x,y) liegen
# $ziel = Fenster, das bei (x,y) zuoberst sein soll
function FreiMachen([int]$x, [int]$y, $ziel) {
    Stop-Process -Name OpenWith -Force -ErrorAction SilentlyContinue
    for ($k = 0; $k -lt 5; $k++) {
        PruefeEsc
        $h = [W]::TopAt($x, $y)
        $p = 0; [void][W]::GetWindowThreadProcessId($h, [ref]$p)
        if ($p -eq $fakt.Id) { return }
        $name = ProzessName $p
        if ($Geschuetzt -contains $name) {
            # Claude, Outlook, Terminal usw. nie schliessen: erst nach hinten, notfalls minimieren
            if ($ziel) { Fokus $ziel }
            if ([W]::TopAt($x, $y) -eq $h) {
                Write-Host "   (minimiere verdeckendes Fenster von $name)" -ForegroundColor DarkGray
                [void][W]::ShowWindow($h, 6)  # SW_MINIMIZE
                Start-Sleep -Milliseconds 500
            }
            continue
        }
        Write-Host "   (schliesse verdeckendes Fenster von $name)" -ForegroundColor DarkGray
        [void][W]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Seconds 1
        if ([W]::TopAt($x, $y) -eq $h) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 500 }
    }
}

# Datei zu einem FileSystemWatcher-Ereignis; Systemdateien (Thumbs.db usw.) zaehlen nicht
function DateiAusEreignis($ev) {
    $pfad = $ev.SourceEventArgs.FullPath
    Remove-Event -EventIdentifier $ev.EventIdentifier
    $name = Split-Path $pfad -Leaf
    if ($name -in "Thumbs.db", "desktop.ini" -or $name -like "~*" -or $name -like "*.tmp") { return $null }
    try {
        $f = Get-Item -LiteralPath $pfad -Force -ErrorAction Stop
        if ($f.PSIsContainer -or ($f.Attributes -band [IO.FileAttributes]::Hidden)) { return $null }
        return $f
    } catch { return $null }
}

function NeueFenster($fensterVorher) {
    foreach ($h in [W]::VisibleWindows()) {
        if ($fensterVorher -contains $h) { continue }
        $p = 0; [void][W]::GetWindowThreadProcessId($h, [ref]$p)
        # auch neue FakturaE-Fenster (Fehlermeldungen); Hauptfenster und Dialog sind in $fensterVorher
        if ($p -eq $PID) { continue }
        $h
    }
}

# wartet, bis der Viewer des Dokuments aufgeht (die Fotos-App braucht mehrere Sekunden),
# und schliesst dann alle neuen Fenster bzw. neu gestarteten Viewer-Prozesse
function ViewerSchliessen($fensterVorher, $prozesseVorher) {
    $bis = (Get-Date).AddSeconds($ViewerWartenSek)
    while ((Get-Date) -lt $bis -and -not (NeueFenster $fensterVorher) -and -not (Get-Process OpenWith -ErrorAction SilentlyContinue)) {
        PruefeEsc
        Start-Sleep -Milliseconds 300
    }
    Start-Sleep -Milliseconds 700
    # "Wie soll diese Datei geoeffnet werden?" (keine App fuer den Dateityp) - eigener Prozess ohne Fenster-Handle
    Stop-Process -Name OpenWith -Force -ErrorAction SilentlyContinue
    foreach ($h in NeueFenster $fensterVorher) {
        [void][W]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
    }
    Start-Sleep -Milliseconds 1500
    $neu = Get-Process | Where-Object { $prozesseVorher -notcontains $_.Id -and $_.Id -ne $PID -and $Geschuetzt -notcontains $_.ProcessName -and $_.MainWindowHandle -ne 0 }
    foreach ($pr in $neu) { Stop-Process -Id $pr.Id -Force -ErrorAction SilentlyContinue }
}

function Taste($k) { [System.Windows.Forms.SendKeys]::SendWait($k); Start-Sleep -Milliseconds 300 }

function SichererName($s) {
    $s = ($s -replace '[\\/:*?"<>|]', '_').Trim().TrimEnd('.')
    if ($s.Length -gt 120) { $s = $s.Substring(0, 120) }
    return $s
}

# Unterfenster des Hauptfensters (Projekt-Eigenschaften ist ein "owned" Fenster)
function Kind($el, [scriptblock]$wo) {
    $ch = $Raw.GetFirstChild($el)
    while ($ch) { if (& $wo $ch) { return $ch }; $ch = $Raw.GetNextSibling($ch) }
    return $null
}

function NachId($el, $id) {
    $c = New-Object System.Windows.Automation.PropertyCondition($UIA::AutomationIdProperty, "$id")
    return $el.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}

function Dialog { Kind $main { param($e) $e.Current.Name -like "Projekt-Eigenschaften:*" } }

# Projektnummer aus dem Titel des Hauptfensters ("... [.NET 10] 230004 - Kom. ...")
function ProjektImTitel {
    if ($main.Current.Name -match '\]\s+(\S+)\s+-\s') { return $Matches[1] }
    return $main.Current.Name
}

function Mitte($el) {
    $r = $el.Current.BoundingRectangle
    return @([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
}

# --- Ein Projekt exportieren ---------------------------------------------------------
function ExportiereProjekt($dlg, $ordner) {
    $hwnd = [IntPtr]$dlg.Current.NativeWindowHandle

    $label = NachId $dlg 1112      # "43 Dokumente"
    $anzahl = 0
    if ($label -and $label.Current.Name -match '^(\d+)') { $anzahl = [int]$Matches[1] }
    $gesamt = $anzahl
    if ($MaxDoks -gt 0) { $anzahl = [Math]::Min($anzahl, $MaxDoks) }
    $log = @()
    if ($anzahl -eq 0) { return $log }

    # erste Zeile der Dokumentliste anklicken
    $liste = NachId $dlg 1101
    $zeilen = Kind $liste { param($e) $e.Current.AutomationId -eq "1" }
    $r = $zeilen.Current.BoundingRectangle
    FreiMachen ([int]$r.X + 40) ([int]$r.Y + 8) $hwnd
    Fokus $hwnd
    Klick ([int]$r.X + 40) ([int]$r.Y + 8)
    Taste "^{HOME}"

    New-Item -ItemType Directory -Force $ordner | Out-Null

    $letzteQuelle = $null
    $fehlFolge = 0
    for ($i = 1; $i -le $anzahl; $i++) {
        PruefeEsc
        FreiMachen ([int]$r.X + 40) ([int]$r.Y + 10) $hwnd
        Fokus $hwnd
        if ($i -gt 1) {
            # Klick auf die markierte Zeile holt den Tastaturfokus zurueck in die Liste
            $y = MarkierteZeile $zeilen
            if ($null -ne $y) { Klick ([int]$r.X + 40) $y }
            Taste "{DOWN}"
        }

        $zeileY = MarkierteZeile $zeilen
        if ($null -eq $zeileY) {
            # zweiter Versuch: Stoerenfriede wegraeumen, Fenster nach vorne
            FreiMachen ([int]$r.X + 40) ([int]$r.Y + 10) $hwnd
            Fokus $hwnd
            $zeileY = MarkierteZeile $zeilen
        }
        if ($null -eq $zeileY) {
            if ($DebugOrdner) {
                New-Item -ItemType Directory -Force $DebugOrdner | Out-Null
                $dr = $dlg.Current.BoundingRectangle
                $bmp = New-Object System.Drawing.Bitmap ([int]$dr.Width), ([int]$dr.Height)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.CopyFromScreen([int]$dr.X, [int]$dr.Y, 0, 0, $bmp.Size)
                $g.DrawRectangle([System.Drawing.Pens]::Red, [int]($r.X - $dr.X), [int]($r.Y - $dr.Y), [int]$r.Width, [int]$r.Height)
                $bmp.Save((Join-Path $DebugOrdner ("doc{0:D3}_fehlt.png" -f $i))); $g.Dispose(); $bmp.Dispose()
                Write-Host "   (debug: Listenbereich $r, Dialog $dr)" -ForegroundColor DarkGray
            }
            Write-Host ("   [{0}/{1}] markierte Zeile nicht gefunden" -f $i, $anzahl) -ForegroundColor Yellow
            $log += [pscustomobject]@{ Nr = $i; Status = "Zeile nicht gefunden"; Datei = ""; Quelle = "" }
            $fehlFolge++
            if ($fehlFolge -ge 3) {
                Write-Host "   Projekt abgebrochen (3x keine Zeile) - wird beim naechsten Lauf wiederholt" -ForegroundColor Yellow
                break
            }
            continue
        }
        $fehlFolge = 0
        if ($DebugOrdner) {
            New-Item -ItemType Directory -Force $DebugOrdner | Out-Null
            $dr = $dlg.Current.BoundingRectangle
            $bmp = New-Object System.Drawing.Bitmap ([int]$dr.Width), ([int]$dr.Height)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.CopyFromScreen([int]$dr.X, [int]$dr.Y, 0, 0, $bmp.Size)
            $g.DrawEllipse([System.Drawing.Pens]::Red, [int]($r.X + 40 - $dr.X - 5), [int]($zeileY - $dr.Y - 5), 10, 10)
            $bmp.Save((Join-Path $DebugOrdner ("doc{0:D3}.png" -f $i))); $g.Dispose(); $bmp.Dispose()
            Write-Host "   (debug: Zeile y=$zeileY)" -ForegroundColor DarkGray
        }
        $vorher = [W]::VisibleWindows()
        $prozesse = (Get-Process).Id
        Get-Event -SourceIdentifier "FaktTemp" -ErrorAction SilentlyContinue | Remove-Event
        DoppelKlick ([int]$r.X + 40) $zeileY

        # Datei abwarten, die FakturaE in den Temp-Ordner schreibt. Die Ordner {..} gelten pro
        # Projekt und Dateityp, und FakturaE setzt das Originaldatum - darum per FileSystemWatcher.
        $datei = $null
        $bis = (Get-Date).AddSeconds($TimeoutSek)
        while (-not $datei -and (Get-Date) -lt $bis) {
            PruefeEsc
            $ev = Wait-Event -SourceIdentifier "FaktTemp" -Timeout 1
            if ($ev) { $d = DateiAusEreignis $ev; if ($d) { $datei = $d } }
        }
        if ($datei) {
            # noch weitere Ereignisse zur selben Oeffnung abwarten und die zuletzt geschriebene Datei nehmen
            while ($ev = Wait-Event -SourceIdentifier "FaktTemp" -Timeout 1) {
                $d = DateiAusEreignis $ev; if ($d) { $datei = $d }
            }
        }

        if (-not $datei) {
            Write-Host ("   [{0}/{1}] KEINE Datei erschienen" -f $i, $anzahl) -ForegroundColor Yellow
            $log += [pscustomobject]@{ Nr = $i; Status = "keine Datei"; Datei = ""; Quelle = "" }
        } else {
            if ($datei.FullName -eq $letzteQuelle) {
                Write-Host "   (Achtung: gleiche Datei wie das vorige Dokument)" -ForegroundColor Yellow
            }
            $letzteQuelle = $datei.FullName
            # warten bis die Datei fertig geschrieben ist
            do { $g = $datei.Length; Start-Sleep -Milliseconds 700; $datei.Refresh() } while ($datei.Length -ne $g)

            $name = SichererName $datei.Name
            $zielDatei = Join-Path $ordner $name
            $n = 2
            while (Test-Path -LiteralPath $zielDatei) {
                $zielDatei = Join-Path $ordner ("{0} ({1}){2}" -f [IO.Path]::GetFileNameWithoutExtension($name), $n, [IO.Path]::GetExtension($name))
                $n++
            }
            Copy-Item -LiteralPath $datei.FullName -Destination $zielDatei
            Write-Host ("   [{0}/{1}] {2}" -f $i, $anzahl, (Split-Path $zielDatei -Leaf))
            $log += [pscustomobject]@{ Nr = $i; Status = "ok"; Datei = (Split-Path $zielDatei -Leaf); Quelle = $datei.FullName }
        }

        # alles schliessen, was seit dem Oeffnen neu aufgegangen ist (Viewer, Rueckfragen)
        ViewerSchliessen $vorher $prozesse
    }
    # mit -MaxDoks begrenzter Testlauf: Projekt nicht als fertig markieren
    if ($anzahl -lt $gesamt) {
        $log += [pscustomobject]@{ Nr = ""; Status = "begrenzt ($anzahl von $gesamt)"; Datei = ""; Quelle = "" }
    }
    return $log
}

# --- Start -----------------------------------------------------------------------------
$fakt = Get-Process Fakt -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $fakt) { throw "FakturaE (Fakt.exe) laeuft nicht." }
$pidCond = New-Object System.Windows.Automation.PropertyCondition($UIA::ProcessIdProperty, $fakt.Id)
$main = $UIA::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, $pidCond)
$mainHwnd = [IntPtr]$main.Current.NativeWindowHandle
$ribbon = Kind $main { param($e) $e.Current.ClassName -like "BCGPRibbonBar*" }
$rr = $ribbon.Current.BoundingRectangle

New-Item -ItemType Directory -Force $TempRoot | Out-Null
$watcher = New-Object System.IO.FileSystemWatcher $TempRoot
$watcher.IncludeSubdirectories = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]"FileName, LastWrite, Size"
foreach ($art in "Created", "Changed", "Renamed") {
    Register-ObjectEvent $watcher $art -SourceIdentifier "FaktTemp$art" -Action {
        New-Event -SourceIdentifier "FaktTemp" -Sender $Event.Sender -EventArguments $Event.SourceEventArgs | Out-Null
    } | Out-Null
}
$watcher.EnableRaisingEvents = $true

$offen = Dialog
if ($offen) {
    Write-Host "Schliesse offenes Fenster Projekt-Eigenschaften ..."
    $s = NachId $offen 201; $m = Mitte $s; Fokus ([IntPtr]$offen.Current.NativeWindowHandle); Klick $m[0] $m[1]
    Start-Sleep -Seconds 1
}

New-Item -ItemType Directory -Force $Ziel | Out-Null
Start-Transcript -Path (Join-Path $Ziel "_protokoll.txt") -Append | Out-Null
Write-Host "Ziel: $Ziel"
Write-Host ""
if ($AbMarkiert) {
    Write-Host "Klicke innerhalb von $StartWarten Sekunden in der Projektliste auf das Projekt, bei dem"
    Write-Host "der Export beginnen soll, und fasse danach Maus und Tastatur nicht mehr an."
    Start-Sleep -Seconds $StartWarten
} else {
    Write-Host "Start beim obersten Projekt der Liste. Maus und Tastatur bitte nicht mehr anfassen."
    Start-Sleep -Seconds 3
    # erste sichtbare Zeile der Projektliste anklicken, dann ganz nach oben
    $pl = $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($UIA::NameProperty, "Projektliste")))
    $plZeilen = Kind (NachId $pl 101) { param($e) $e.Current.AutomationId -eq "1" }
    $pr = $plZeilen.Current.BoundingRectangle
    Fokus $mainHwnd
    Klick ([int]$pr.X + 40) ([int]$pr.Y + 8)
    Taste "^{HOME}"
    Start-Sleep -Seconds 1
}
Write-Host "Abbruch: Esc-Taste (wirkt in jedem Fenster)."
[void][W]::GetAsyncKeyState(0x1B)   # alten Tastenzustand verwerfen

$pl = $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($UIA::NameProperty, "Projektliste")))
$pr = (Kind (NachId $pl 101) { param($e) $e.Current.AutomationId -eq "1" }).Current.BoundingRectangle

try {
$p = 0
while ($Projekte -eq 0 -or $p -lt $Projekte) {
    $p++
    PruefeEsc
    FreiMachen ([int]$pr.X + 40) ([int]$pr.Y + 10) $mainHwnd
    Fokus $mainHwnd
    if ($p -gt 1) {
        # eine Zeile runter; der Titel des Hauptfensters zeigt das markierte Projekt
        $vorNr = ProjektImTitel
        Taste "{DOWN}"
        Start-Sleep -Milliseconds 700
        if ((ProjektImTitel) -eq $vorNr) {
            Fokus $mainHwnd; Taste "{DOWN}"; Start-Sleep -Seconds 1
            if ((ProjektImTitel) -eq $vorNr) {
                Write-Host "Projektliste springt nicht weiter (Ende der Liste oder Fokus verloren) - Lauf beendet."
                Write-Host "Ein neuer Lauf ueberspringt alle fertigen Projekte."
                break
            }
        }
    }

    Klick ([int]$rr.X + $DokuKnopfX) ([int]$rr.Y + $DokuKnopfY)
    $dlg = $null
    $bis = (Get-Date).AddSeconds(20)
    while (-not $dlg -and (Get-Date) -lt $bis) { Start-Sleep -Milliseconds 500; $dlg = Dialog }
    if (-not $dlg) { throw "Fenster 'Projekt-Eigenschaften' ging nicht auf. Stimmt -DokuKnopfX/-DokuKnopfY?" }
    Start-Sleep -Seconds 1

    $titel = $dlg.Current.Name.Substring("Projekt-Eigenschaften:".Length).Trim()
    $nr = ($titel -split '\s+')[0]
    $ordner = Join-Path $Ziel (SichererName $titel)
    $logDatei = Join-Path $ordner "_export.csv"

    # fertig ist ein Projekt nur, wenn das Protokoll existiert und jede Zeile "ok" ist
    $fertig = (Test-Path -LiteralPath $logDatei) -and
        -not @(Import-Csv -LiteralPath $logDatei -Delimiter ";" | Where-Object Status -ne "ok").Count
    if ($fertig) {
        Write-Host "[$p] $titel  -> schon exportiert, uebersprungen"
    } else {
        Write-Host "[$p] $titel"
        # Reste eines abgebrochenen Laufs entfernen, sonst entstehen "(2)"-Dubletten
        if (Test-Path -LiteralPath $ordner) { Get-ChildItem -LiteralPath $ordner -Force | Remove-Item -Recurse -Force }
        $log = @(ExportiereProjekt $dlg $ordner)
        New-Item -ItemType Directory -Force $ordner | Out-Null
        if ($log.Count -gt 0) {
            $log | Export-Csv -LiteralPath $logDatei -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            $fehlt = @($log | Where-Object Status -ne "ok").Count
            Write-Host ("   -> {0} Dateien, {1} fehlgeschlagen" -f ($log.Count - $fehlt), $fehlt)
        } else {
            Set-Content -LiteralPath $logDatei -Value '"Nr";"Status";"Datei";"Quelle"' -Encoding UTF8
            Write-Host "   -> keine Dokumente"
        }
    }

    # Projekt-Eigenschaften schliessen
    $dlg = Dialog
    if ($dlg) {
        Fokus ([IntPtr]$dlg.Current.NativeWindowHandle)
        $s = NachId $dlg 201; $m = Mitte $s; Klick $m[0] $m[1]
        $bis = (Get-Date).AddSeconds(10)
        while ((Dialog) -and (Get-Date) -lt $bis) { Start-Sleep -Milliseconds 300 }
    }
}
} catch {
    if ("$_" -ne "ABBRUCH_ESC") { throw }
    Write-Host ""
    Write-Host "Mit Esc abgebrochen. Das angefangene Projekt wird beim naechsten Lauf wiederholt." -ForegroundColor Yellow
}

Fokus $mainHwnd
Write-Host ""
Write-Host "Fertig."
