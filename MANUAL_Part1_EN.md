# BSP - Biomechanical Stability Program
## Quick Reference Manual v1.0

**Authors:** Andre Massuca (development) | Pedro Aleixo (biomechanics) | Luis Massuca (coordination)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Requirements and Installation

| OS | Min version | Python |
|---|---|---|
| Windows | 10 / 11 | 3.9+ |
| macOS | 12 Monterey | 3.9+ |

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

**Windows:** Double-click `BUILD_Windows.bat` -> generates `BSP_Setup.exe`
**macOS:** `chmod +x BUILD_macOS.sh && ./BUILD_macOS.sh` -> generates `BSP.dmg`

First launch on macOS: right-click -> Open (bypass Gatekeeper once).

---

## 2. First Run

1. **Automatic theme** - detects OS dark/light mode
2. **End User Licence** - accept once; saved in `~/.aom_estabilidade.json`
3. **Password** - available at https://github.com/andremassuca
4. **Protocol selection** - FMS / Unipodal / Functional Task (Shooting / Archery)

---

## 3. Input File Structure

```
main_folder/
    001_Subject_Name/
        dir_1.xls   (right foot, trial 1)
        dir_2.xls
        esq_1.xls   (left foot, trial 1)
        ...
    002_Another_Name/
        ...
```

Any number of subjects and trials is supported.
PDF tables auto-adapt columns for > 10 trials.

---

## 4. Available Protocols

| Protocol | Description |
|---|---|
| FMS Bipodal | N trials per foot, 95% ellipse, Dir/Esq asymmetry |
| Unipodal | N trials per foot, lateral oscillation metrics |
| Shooting (ISCPSI) | Multi-distance, multi-interval, stability vs precision score |
| Archery | Up to 30 bipodal trials, single analysis window (Confirmation 1 to 2), no distance |

---

## 5. Main Interface

**Left panel** - configuration (folder, timing file, output, PDF path, N trials)
**Right panel** - execution log (blue=info, green=ok, yellow=warning, red=error)

**Frame rate validation (v23):** BSP automatically checks sampling consistency
between trials. Warnings appear in yellow if jitter > 20% or > 10% rate variation.

---

## 6. Output Options

| Output | Description |
|---|---|
| Summary Excel | DADOS + GRUPO + SPSS + ESTATS sheets |
| Individual files | One xlsx per subject with ellipse and stabilogram |
| PDF Report | Cover + What's New + Legend + individuals + citation |
| Word Report | Statistical tables (requires python-docx) |
| HTML Report | Interactive Chart.js, standalone, dark/light mode |
| CSV Export | Ready for SPSS / R / Excel |
| PNG Export | Stabilograms and ellipses at configurable DPI |

**PNG export (v23):** DPI options: 72 (web), 150 (slides), 180 (default), 300 (print).
Select types: stabilogram and/or ellipse.

---

## 7. Statistical Analyses

Enable in **Options -> Automatic statistical tests (ESTATS sheet)**.

- Shapiro-Wilk normality test
- Paired t-test or Wilcoxon + Cohen's d (Dir vs Esq)
- Intra-subject CV% with traffic light (green < 15%, yellow 15-30%, red > 30%)
- Friedman test + Bonferroni post-hoc (Shooting)
- Pearson/Spearman correlation ea95 vs precision score (Shooting)

Requires n >= 3 subjects.

---

## 8. Keyboard Shortcuts

| Windows | macOS | Action |
|---|---|---|
| Ctrl+Enter | Cmd+Enter | Run analysis |
| Ctrl+S | Cmd+S | Save config |
| Ctrl+H | Cmd+H | History |
| F1 | F1 | All shortcuts |
| F2 | F2 | Quick analysis - single file |
| F5 | F5 | Clear log |

---

## 9. Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError` | `pip install numpy scipy openpyxl matplotlib reportlab` |
| macOS security warning | Right-click -> Open (once) |
| PDF not generated | Ensure PDF field not empty and option checked |
| Silent crash (v22) | Update to v23 - fixes analysis thread crash |
| > 10 trials | No limit - PDF auto-adapts columns and font size |
| Frame rate warning | Check force platform export settings |

---

## Academic Citation

```
Massuca, A., & Massuca, L. (2026). BSP - Biomechanical Stability Program (v23).
https://github.com/andremassuca/BSP
```

*BSP v23 - Andre Massuca & Luis Massuca*
