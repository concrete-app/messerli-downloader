# messerli-downloader

Exportiert die Dokumente aller Projekte aus **Messerli FakturaE** in einen normalen Ordner, ein Unterordner pro Projekt.

FakturaE bietet keinen Massenexport für das Dokumentenmanagement. Das PowerShell-Skript [export_projekt.ps1](export_projekt.ps1) klickt sich darum selbst durch die Oberfläche: Es öffnet jedes Dokument einmal, fängt die Datei ab, die FakturaE dabei entschlüsselt im Temp-Ordner ablegt, und kopiert sie ins Ziel.

> **Hinweis:** Dies ist ein inoffizielles Werkzeug. Es steht in keiner Verbindung zur Messerli Informatik AG. Nutzung auf eigenes Risiko. Vorher einen Testlauf mit wenigen Projekten machen.

## Voraussetzungen

- Windows mit PowerShell 5.1 oder neuer
- FakturaE (`Fakt.exe`) läuft und zeigt die Projektliste an
- Lesezugriff auf den Temp-Ordner der Messerli-Datenbank, typischerweise `\\<Server>\Messerli\Daten\dbs\Temp\<Benutzer>`

## Verwendung

Testlauf mit einem Projekt und höchstens drei Dokumenten:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_projekt.ps1 `
    -TempRoot "\\SERVER\Messerli\Daten\dbs\Temp\$env:USERNAME" -Projekte 1 -MaxDoks 3
```

Kompletter Export:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_projekt.ps1 `
    -TempRoot "\\SERVER\Messerli\Daten\dbs\Temp\$env:USERNAME"
```

Während das Skript läuft, Maus und Tastatur nicht berühren. **Abbrechen mit `Esc`**, das wirkt in jedem Fenster.

### Parameter

| Parameter | Standard | Beschreibung |
|---|---|---|
| `-TempRoot` | *(Pflicht)* | Temp-Ordner, in den FakturaE geöffnete Dokumente schreibt |
| `-Ziel` | `%USERPROFILE%\Documents\FakturaE Export` | Zielordner für den Export |
| `-Projekte` | `0` | Anzahl Projekte; `0` = bis zum Ende der Projektliste |
| `-MaxDoks` | `0` | Höchstzahl Dokumente pro Projekt; `0` = alle |
| `-AbMarkiert` | aus | Beim markierten statt beim obersten Projekt beginnen |
| `-StartWarten` | `15` | Sekunden, um mit `-AbMarkiert` das Startprojekt anzuklicken |
| `-DokuKnopfX` / `-DokuKnopfY` | `530` / `80` | Position des Ribbon-Knopfs „Dokumentenmanagement“, relativ zur Ribbon-Leiste |
| `-TimeoutSek` | `30` | Wartezeit auf die Datei nach dem Öffnen eines Dokuments |
| `-ViewerWartenSek` | `10` | Wartezeit auf den Viewer, bevor er geschlossen wird |
| `-DebugOrdner` | – | Wenn gesetzt, wird vor jedem Doppelklick ein Bildschirmfoto der Dokumentliste abgelegt |

## Funktionsweise

Für jedes Projekt in der Projektliste:

1. Ribbon-Knopf „Dokumentenmanagement“ klicken und auf das Fenster „Projekt-Eigenschaften“ warten
2. Jedes Dokument per Doppelklick öffnen. Ein `FileSystemWatcher` auf `-TempRoot` erkennt die Datei, die FakturaE dabei schreibt.
3. Die Datei nach `<Ziel>\<Projektnr> <Bezeichnung>\` kopieren und den Viewer schliessen, der sich dabei geöffnet hat
4. Das Fenster schliessen und in der Projektliste eine Zeile weiter

Die Steuerung läuft über UI Automation und simulierte Maus- und Tastatureingaben. Die markierte Zeile wird über ihre Farbe erkannt. Fenster anderer Programme, die die Liste verdecken, werden geschlossen. Geschützte Programme wie Outlook, Teams oder das Terminal werden nur minimiert.

## Ausgabe

```
<Ziel>\
├── _protokoll.txt                 Konsolenprotokoll aller Läufe
└── 12345 Beispielprojekt\
    ├── _export.csv                Status pro Dokument (ok / keine Datei / ...)
    ├── Offerte.pdf
    └── ...
```

Ein Projekt gilt als fertig, wenn seine `_export.csv` existiert und jede Zeile den Status `ok` hat. Nach einem Abbruch das Skript einfach neu starten: Fertige Projekte werden übersprungen, unvollständige werden neu exportiert.

## Fehlersuche

- **„Fenster 'Projekt-Eigenschaften' ging nicht auf“**: Die Position des Ribbon-Knopfs passt nicht zur Bildschirmauflösung bzw. Skalierung. `-DokuKnopfX` und `-DokuKnopfY` anpassen.
- **„markierte Zeile nicht gefunden“**: Mit `-DebugOrdner .\debug` laufen lassen und die Bilder prüfen.

## Lizenz

[MIT](LICENSE)
