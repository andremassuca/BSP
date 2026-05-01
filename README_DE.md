<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

**Was es ist:** eine Anwendung zur Analyse der Haltungsstabilität von Athleten aus Kraftmessplatten-Daten. Ich habe das gebaut, weil zu viel davon in Excel von Hand gemacht wurde - das hier verwandelt rohen CoP in saubere Messgrößen, ohne Code schreiben zu müssen.

**Für wen:** Forscher, Trainer und Kliniker, die mit Kraftmessplatten arbeiten und von rohem CoP zu nutzbaren Messgrößen (ea95, RMS, Durchschnittsgeschwindigkeit, Haltungssteifheit) übergehen wollen.

## Was es macht

Vier Protokolle:

- **FMS Bipodal** - einfacher beidfüßiger Stand mit optionalen Links-/Rechts-Panels.
- **Einbeiniger Stand** - dominante / nicht-dominante Asymmetrie.
- **Pistolenschießen** - zwei Stabilisierungsfenster vor dem Abzug.
- **Bogenschießen** - bis zu 30 Durchgänge, einzelnes Fenster zwischen Bestätigung 1 und 2.

Für jedes Protokoll werden berechnet:

- Fläche der 95%-Konfidenzellipse (`ea95`), RMS und Amplituden in X/Y und radial.
- Durchschnittsgeschwindigkeit und Haltungssteifheitsindex (`stiff_x`, `stiff_y`).
- Dominante / nicht-dominante Asymmetrie wenn anwendbar.
- CoP-FFT zur spektralen Inspektion.

Ausgaben: Excel mit Zusammenfassung, pro Durchgang und Demografie (wenn Referenzdatei vorhanden), PDF pro Athlet und Gruppen-PDF, optionales lokales Web-Dashboard mit interaktiven Listen, Boxplots, Scatter mit Regression und Athleten-Detailansicht.

## 1-Minuten-Quickstart

```bash
git clone https://github.com/andremassuca/BSP.git
cd BSP
pip install -r requirements.txt
python estabilidade_gui.py
```

Der erste Start öffnet das Lizenzfenster → Passwort → Protokollauswahl.

Um die synthetische Testsuite auszuführen, bevor echte Daten verarbeitet werden:

```bash
python estabilidade_gui.py --testes     # GUI/PDF/Excel (~161 Tests)
python -m bsp_core --testes             # nur Mathe-Kern (~41 Tests)
```

Beide müssen **100% bestanden** melden.

## Arbeiten mit Daten

**Ordnerkonvention:** ein Ordner pro Messsitzung. Jeder Athlet trägt seine ID im Dateinamen.

**Bogenschießen**

- Ordner mit ZIPs oder Unterordnern; Dateinamen `{id}_{trial} - {TT-MM-JJJJ} - Stability export.xls` (tab-separiert, 50 Hz).
- `Inicio_fim_vfinal.xlsx` mit drei Blättern (`tempo do toque`, `confirmação_1`, `confirmação_2`). Analysefenster: `confirmação_1 → confirmação_2`. Der Loader toleriert cp1252-verstümmelte Namen.
- `Todos os registos dos 142 atletas em JUl_2024 _.xlsx` als demografische Referenz.

**Andere Protokolle**

- Von der Plattform-Software exportierte Dateien (`.xlsx` oder tab-separiertes `.txt`).
- `inicio_fim.xlsx` für Pistolenschießen, mit den zwei Zeitfenstern pro Durchgang.


## Bekannte Probleme

- **macOS, erster Start:** Gatekeeper blockiert nicht notarisierte Apps. **Rechtsklick → Öffnen**, bestätigen. Nur einmal nötig.
- **macOS, App startet nicht:** `~/Library/Logs/BSP/BSP_crash.log` prüfen. Wenn leer, `./BUILD_macOS.sh` im Terminal ausführen und `build_logs/` lesen.
- **Windows, fehlende DLL:** [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe) installieren.
- **Dashboard öffnet nicht:** uvicorn braucht 1–3 s zum Starten. Wenn nach 10 s nichts erscheint, sicherstellen, dass `fastapi`/`uvicorn` installiert sind.

## Zitieren

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Version 1.0) [Software]. https://github.com/andremassuca/BSP

```bibtex
@software{massuca_bsp_2025,
  author  = {Massu\c{c}a, Andr\'{e} Oliveira and Aleixo, Pedro and Massu\c{c}a, Lu\'{i}s M.},
  title   = {BSP: Biomechanical Stability Program},
  version = {1.0},
  year    = {2026},
  url     = {https://github.com/andremassuca/BSP},
}
```

## Lizenz

Akademischer Code, frei für Forschung und Lehre. Für kommerzielle Nutzung oder Weitergabe, den Autor kontaktieren.

**André O. Massuça** - [github.com/andremassuca](https://github.com/andremassuca)
**Pedro Aleixo**
**Luís M. Massuça**
