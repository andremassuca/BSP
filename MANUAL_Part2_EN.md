# BSP - Biomechanical Stability Program
## Full User Manual - v1.0

**Authors:** Andre Massuca (development) | Pedro Aleixo (biomechanics) | Luis Massuca (coordination)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Requirements and Installation

### Compatibility

| System | Minimum | Tested |
|---|---|---|
| Windows | 10 / 11 | 10 22H2, 11 23H2 |
| macOS | 12 Monterey | Apple Silicon M1/M2/M3 and Intel |
| Python | 3.9 | 3.9 / 3.10 / 3.11 / 3.12 |

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

**Windows:** `BUILD_Windows.bat` -> `BSP_Setup.exe` (standalone installer)
**macOS:** `./BUILD_macOS.sh` -> `BSP.dmg`

High-DPI support: Windows DPI awareness enabled automatically; macOS Retina display supported natively.

---

## 2. Protocols

| Protocol | Sides | Trials | Features |
|---|---|---|---|
| FMS Bipodal | Dir + Esq | N (default 5) | 95% Ellipse, Asymmetry index |
| Unipodal | Dir + Esq | N (default 5) | RMS ML/AP, Amplitude |
| Shooting | Entire plate + Selection | N per distance | Multi-interval, score correlation |
| Archery | Same as Shooting | N | Adapted for archery |

All protocols support unlimited subjects and trials. PDF tables auto-scale.

---

## 3. Input File Structure

### FMS / Unipodal
```
main_folder/
    001_Name/
        dir_1.xls, dir_2.xls, ...  (right foot)
        esq_1.xls, esq_2.xls, ...  (left foot)
```

### Shooting
```
main_folder/
    001_Shooter/
        trial5_1_-_date_-_Stability_export.xls
        trial10_1_-_date_-_Stability_export.xls
        hs_dir_1.xls  (Hurdle Step, optional)
```

---

## 4. New in v23

| Feature | Description |
|---|---|
| Silent crash fix | Analysis thread errors now show diagnostic dialog with email report |
| RMS in PDF | RMS ML, AP, Radius (Quijoux et al., 2021) added to individual table |
| HurdleStep table fix | Column headers no longer overlap with multiple trials |
| Unlimited inputs | PDF tables auto-adapt: columns shrink, font reduces for > 10 trials |
| Frame rate validation | Detects duplicate timestamps, jitter > 20%, rate variation > 10% |
| PNG options | Configurable DPI (72/150/180/300) and type selection |
| What's New PDF page | Automatically generated after the legend page |
| Dynamic citations | APA and BibTeX year uses current system year |
| Updated authorship | Andre Massuca & Luis Massuca |

---

## 5. Output Options

### PDF Report Structure
1. Cover (title, protocol, date, subjects count)
2. Cover overflow (large datasets)
3. Legend (all metrics with formulas)
4. **What's New (v23)** - auto-generated version notes
5. Table of contents (Shooting protocol only)
6. Group summary (Shooting)
7. Per-subject pages (table + stabilogram + ellipse)
8. Statistical tests (if enabled)
9. Academic citation (APA + BibTeX with dynamic year)

### PNG Export (v23)
- DPI: 72 (web), 150 (slides), 180 (default), 300 (publication)
- Types: stabilogram and/or 95% ellipse
- Output: `subject_folder/png/`

### HTML Report
Standalone `.html` with Chart.js charts. No installation required on recipient side.
Supports automatic dark/light mode via CSS `prefers-color-scheme`.

---

## 6. Frame Rate Validation (v23)

Automatic checks on each subject's trial data:

| Check | Threshold | Action |
|---|---|---|
| Duplicate/inverted timestamps | Any occurrence | Warning in log |
| Jitter within trial | > 20% relative variation | Warning in log |
| Rate variation across trials | > 10% between trials | Warning in log |

Warnings are yellow in the log and do not interrupt analysis.
The calculations use real timestamps so mild jitter does not affect results.

---

## 7. Multilingual Support

All 4 languages (PT / EN / ES / DE) share 336 translation keys.
Language change updates: UI labels, PDF metric labels, side labels,
Excel headers, HTML labels, EULA text.

Metric label examples:
- PT: "Area elipse 95% (mm2)"
- EN: "95% Ellipse area (mm2)"
- DE: "95%-Ellipsenflache (mm2)"

Side labels: EN "Right Foot" / "Left Foot", DE "Rechter Fuss" / "Linker Fuss"

---

## 8. Version History

### v23 (2026)
Silent crash fix, RMS in PDF, HurdleStep overlap fix, unlimited inputs,
frame rate validation, PNG DPI options, What's New PDF page, dynamic citations,
updated authorship (Andre Massuca & Luis Massuca), no institution references.

### v22 (2025)
Automatic theme, ETA progress bar, configuration profiles, quick analysis (F2),
interactive HTML report, citation page in PDF, SPSS headers by language.

### v21 (2024)
Shooting protocol, Selection CoP, Hurdle Step, Friedman tests,
score correlation, full multilingual support.

---

## 9. Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError` | Install all dependencies via pip |
| macOS Gatekeeper | Right-click -> Open (once) |
| PDF not generated | Ensure PDF path set and checkbox active |
| Silent crash (v22) | Update to v23 |
| > 10 trials | No limit - auto-adapts |
| Frame rate warning | Check force platform export settings |
| HTML chart not loading | Check internet connection (Chart.js via CDN) |
| Statistical tests greyed out | Enable option and ensure n >= 3 |

---

## Academic Citation

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

---

## References

- Schubert, P., & Kirchner, M. (2013). Ellipse area calculations. *Gait & Posture*, 39(1), 518-522.
- Prieto, T.E. et al. (1996). Measures of postural steadiness. *IEEE Trans. Biomed. Eng.*, 43(9), 956-966.
- Quijoux, F. et al. (2021). A review of COP variables. *Physiological Reports*, 9(22), e15067.
- Winter, D.A. (1995). Human balance and posture control. *Gait & Posture*, 3(4), 193-214.
- Cohen, J. (1988). *Statistical Power Analysis* (2nd ed.). Erlbaum.

*BSP v23 - Andre Massuca & Luis Massuca*
