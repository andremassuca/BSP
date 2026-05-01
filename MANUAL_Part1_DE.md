# BSP - Biomechanical Stability Program
## Kurzreferenz-Handbuch v1.0

**Autoren:** Andre Massuca (Entwicklung) | Pedro Aleixo (Biomechanik) | Luis Massuca (Koordination)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Anforderungen und Installation

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

**Windows:** `BUILD_Windows.bat` -> `BSP_Setup.exe`
**macOS:** `./BUILD_macOS.sh` -> `BSP.dmg`

Kompatibilitat: Windows 10/11, macOS 12 Monterey+, Python 3.9+

---

## 2. Protokolle

| Protokoll | Beschreibung |
|---|---|
| FMS Bipodal | N Versuche pro Fuss, 95%-Ellipse, Asymmetrie |
| Einbeinstand | N Versuche pro Fuss, laterale Schwankungsmetriken |
| Schiessen (ISCPSI) | Multi-Distanz, Multi-Intervall, Prazisionskorrelation |
| Bogenschiessen | Bis zu 30 beidbeinige Versuche, einzelnes Analysefenster (Bestatigung 1 bis 2), keine Distanz |

Unbegrenzte Anzahl von Probanden und Versuchen. PDF-Tabellen passen sich automatisch an.

---

## 3. Dateistruktur

```
hauptordner/
    001_Name/
        dir_1.xls, dir_2.xls, ...
        esq_1.xls, esq_2.xls, ...
    002_AndererName/
        ...
```

---

## 4. Neues in v23

- Stiller Absturz behoben (Analyse-Thread)
- RMS ML, AP, Radius im PDF-Bericht hinzugefugt
- Validierung der Abtastrate zwischen Versuchen
- PNG-Export mit konfigurierbarem DPI (72/150/180/300)
- "Was ist neu"-Seite im PDF automatisch generiert
- Unbegrenzte Eingaben - PDF-Tabellen passen sich an
- Aktualisierte Autorschaft: Andre Massuca & Luis Massuca
- Dynamische Jahreszahl in Zitaten (Systemjahr)

---

## 5. Tastaturkurzel

| Windows | macOS | Aktion |
|---|---|---|
| Ctrl+Enter | Cmd+Enter | Analyse starten |
| Ctrl+S | Cmd+S | Konfiguration speichern |
| F1 | F1 | Alle Kurzel |
| F2 | F2 | Schnellanalyse |

---

## 6. Problemlosung

| Problem | Losung |
|---|---|
| `ModuleNotFoundError` | Abhangigkeiten mit pip installieren |
| macOS Sicherheitswarnung | Rechtsklick -> Offnen |
| Stiller Absturz (v22) | Auf v23 aktualisieren |
| > 10 Versuche | Kein Limit - automatische Anpassung |
| Abtastrate-Warnung | Export-Einstellungen der Kraftmessplatte prufen |

---

## Akademisches Zitat

```
Massuca, A., & Massuca, L. (2026). BSP - Biomechanical Stability Program (v23).
https://github.com/andremassuca/BSP
```

*BSP v23 - Andre Massuca & Luis Massuca*
