# BSP - Biomechanical Stability Program
## Vollstandiges Benutzerhandbuch v1.0

**Autoren:** Andre Massuca (Entwicklung) | Pedro Aleixo (Biomechanik) | Luis Massuca (Koordination)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Systemanforderungen

| System | Mindestversion | Getestet |
|---|---|---|
| Windows | 10 / 11 | 10 22H2, 11 23H2 |
| macOS | 12 Monterey | Apple Silicon M1/M2/M3 und Intel |
| Python | 3.9 | 3.9 / 3.10 / 3.11 / 3.12 |

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

HiDPI-Unterstutzung: Windows DPI-Awareness automatisch aktiviert; macOS Retina nativ.

---

## 2. Protokolle

- **FMS Bipodal:** Rechts + Links, N Versuche, 95%-Ellipse, Asymmetrieindex
- **Einbeinstand:** Rechts + Links, RMS ML/AP, Amplituden
- **Schiessen:** Gesamtplatte + Fussauswahl, Multi-Distanz, Korrelation mit Treffern
- **Bogenschiessen:** Wie Schiessen, fur Bogenschiessen angepasst

Unbegrenzte Probanden und Versuche. PDF-Tabellen skalieren automatisch.

---

## 3. Neuerungen in v23

| Funktion | Beschreibung |
|---|---|
| Stiller Absturz behoben | Analyse-Thread-Fehler zeigen Diagnosefenster mit E-Mail-Option |
| RMS im PDF | RMS ML, AP, Radius (Quijoux et al., 2021) in Einzeltabelle |
| HurdleStep-Tabellen korrigiert | Spaltenuberschriften uberlappen nicht mehr |
| Unbegrenzte Eingaben | PDF passt Spalten und Schrift automatisch an |
| Abtastrate-Validierung | Erkennt doppelte Zeitstempel, Jitter >20%, Variation >10% |
| PNG-Optionen | Konfigurierbares DPI (72/150/180/300), Typauswahl |
| "Was ist neu"-PDF-Seite | Automatisch generiert |
| Dynamische Zitate | APA und BibTeX nutzen aktuelles Systemjahr |
| Aktualisierte Autorschaft | Andre Massuca & Luis Massuca |

---

## 4. Ausgabeoptionen

| Option | Beschreibung |
|---|---|
| Zusammenfassung Excel | DATEN + GRUPPE + SPSS + ESTATS |
| Einzeldateien | Eine xlsx pro Proband |
| PDF-Bericht | Deckblatt + Neuerungen + Legende + Probanden + Zitat |
| Word-Bericht | Statistiktabellen (erfordert python-docx) |
| HTML-Bericht | Interaktive Chart.js-Grafiken, eigenstandig |
| CSV-Export | Bereit fur SPSS / R / Excel |
| PNG-Export | Konfigurierbare DPI |

---

## 5. Abtastrate-Validierung (v23)

| Prufung | Schwellenwert | Aktion |
|---|---|---|
| Doppelte/invertierte Zeitstempel | Jedes Auftreten | Warnung im Log |
| Jitter innerhalb Versuch | >20% relative Variation | Warnung im Log |
| Variation zwischen Versuchen | >10% zwischen Versuchen | Warnung im Log |

Warnungen erscheinen gelb im Log ohne die Analyse zu unterbrechen.

---

## 6. Sprachen

PT / EN / ES / DE mit 336 Ubersetzungsschlussel. Alle Ausgaben (PDF, Excel, HTML)
passen sich der gewahlten Sprache an.

Metriknamen: DE "95%-Ellipsenflache (mm2)", "Mittl. CoP-Geschwindigkeit (mm/s)"
Seiten: "Rechter Fuss" / "Linker Fuss"

---

## 7. Fur die Verteilung kompilieren

**Windows:** `BUILD_Windows.bat` -> installiert in `%LOCALAPPDATA%\Programs\BSP\`
**macOS:** `./BUILD_macOS.sh` -> `BSP.dmg` (macOS 12+, Intel und Apple Silicon M1/M2/M3)

---

## 8. Fehlerbehebung

| Problem | Losung |
|---|---|
| Abhangigkeiten nicht gefunden | pip install... |
| macOS Sicherheitswarnung | Rechtsklick -> Offnen |
| PDF nicht generiert | Pfad und Option prufen |
| Stiller Absturz (v22) | Auf v23 aktualisieren |
| > 10 Versuche | Kein Limit in v23 |
| Abtastrate-Warnung | Export der Kraftmessplatte prufen |
| HTML ladt nicht | Internetverbindung prüfen (Chart.js CDN) |

---

## Akademisches Zitat

```
Massuca, A., & Massuca, L. (2026). BSP - Biomechanical Stability Program (v23).
https://github.com/andremassuca/BSP
```

```bibtex
@software{BSP_v23,
  author  = {Massuca, Andre and Massuca, Luis},
  title   = {BSP - Biomechanical Stability Program},
  year    = {2026},
  version = {23},
  url     = {https://github.com/andremassuca/BSP}
}
```

*BSP v23 - Andre Massuca & Luis Massuca*
