<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

Force-plate analysis for postural stability. I built this because too much of this work was being done by hand in Excel - you point it at the files your platform exports and get clean metrics, PDFs, and an optional web dashboard without writing any code.

[PT](README_PT.md) · [ES](README_ES.md) · [DE](README_DE.md)

## What it does

Four protocols: FMS Bipodal, Single-leg Stance, Pistol Shooting, and Archery.

In each one it computes the 95% confidence ellipse area (`ea95`), RMS and amplitudes in X/Y and radial, mean velocity, postural stiffness index (`stiff_x`, `stiff_y`), dominant/non-dominant asymmetry where the protocol allows it, and CoP FFT for spectral inspection.

Archery also includes full demographic analysis for a reference population: Mann-Whitney and Kruskal-Wallis comparisons, Pearson/Spearman correlations, per-subgroup percentiles, and per-trial score linkage.

Output is an Excel file (summary, per-trial, demographics), per-athlete and group PDFs with ellipses and stabilograms, and an optional local web dashboard.

## Quickstart

You need Python 3.9 or later. If `python --version` gives an error or opens the Windows Store, download it from [python.org/downloads](https://www.python.org/downloads/) and tick "Add Python to PATH" during install.

```bash
git clone https://github.com/andremassuca/BSP.git
cd BSP
pip install -r requirements.txt
python estabilidade_gui.py
```

First run opens the licence window, then asks for a password, then shows the protocol picker. The password is in the access document sent separately.

Run the tests before touching real data:

```bash
python estabilidade_gui.py --testes   # GUI, PDF, Excel
python -m bsp_core --testes           # math core only
```

Both must report 100% pass.

## Data layout

One folder per recording session. Each athlete's files have the athlete ID in the filename.

**FMS Bipodal / Single-leg Stance** - `dir_*.xls` and `esq_*.xls` for right and left foot, with `inicio_fim.xlsx` holding the time windows per trial.

**Pistol Shooting** - `trial{distance}_{trial} - {date} - Stability export.xls`, with `inicio_fim.xlsx` providing two windows per trial (aiming period and trigger period).

**Archery** - `{id}_{trial} - {DD-MM-YYYY} - Stability export.xls` at 50 Hz, `Inicio_fim_vfinal.xlsx` with three sheets (`tempo do toque`, `confirmação_1`, `confirmação_2`), and optionally the demographic reference file. The analysis window is `confirmação_1 → confirmação_2`. cp1252-mangled sheet names are handled automatically.


## Building a standalone installer

**Windows**

```cmd
BUILD_Windows.bat
```

Produces `BSP_Setup.exe` (~150 MB, self-contained). First build takes 3–6 minutes; subsequent ones are faster because pip caches the wheels.

**macOS**

```bash
chmod +x BUILD_macOS.sh
./BUILD_macOS.sh
```

Produces `BSP.dmg`, compatible with macOS 12+ on Apple Silicon and Intel. First launch: right-click the `.app` inside the DMG, choose Open, confirm. After that you can drag it to Applications as normal.

Build logs go to `build_logs/`. If anything fails, start there.

## Known issues

| Symptom | Fix |
|---|---|
| macOS: "BSP cannot be opened" | Right-click the `.app` inside the DMG → Open → confirm. One-time only. |
| macOS: app closes immediately | Check `~/Library/Logs/BSP/BSP_crash.log`. If empty, run `./BUILD_macOS.sh` from terminal and read `build_logs/`. |
| Windows: missing DLL | Install [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe). |
| Archery: time-windows file won't load | Open the xlsx and rename the sheets to `confirmação_1` and `confirmação_2`, then save. |
| Demographic analysis empty | Check that the athlete IDs in the reference xlsx match the folder IDs. |

## Citing

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Version 1.0) [Computer software]. https://github.com/andremassuca/BSP

```bibtex
@software{massuca_bsp_2025,
  author  = {Massu\c{c}a, Andr\'{e} Oliveira and Aleixo, Pedro and Massu\c{c}a, Lu\'{i}s M.},
  title   = {BSP: Biomechanical Stability Program},
  version = {1.0},
  year    = {2026},
  url     = {https://github.com/andremassuca/BSP},
}
```

Methods: Schubert & Kirchner (2013), Prieto et al. (1996), Winter (1995), Quijoux et al. (2021), Carpenter et al. (2001), Maurer & Peterka (2005). Full references in [README_EN.md](README_EN.md).

## Licence

Academic code. Free for research and teaching. For commercial use or redistribution, contact the author.

**André O. Massuça** · [github.com/andremassuca](https://github.com/andremassuca) · See [AUTHORS.md](AUTHORS.md) for everyone involved.
