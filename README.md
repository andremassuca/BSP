<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

Force-plate analysis for postural stability. Built to replace manual CoP work in Excel with clean metrics, PDFs, and an optional web dashboard.

[PT](README_PT.md)

> **Source code is not yet publicly available.** The full release will follow publication of the associated methods paper.

## What it does

Four protocols: FMS Bipodal, Single-leg Stance, Pistol Shooting, and Archery.

In each one it computes the 95% confidence ellipse area (`ea95`), RMS and amplitudes in X/Y and radial, mean velocity, postural stiffness index (`stiff_x`, `stiff_y`), dominant/non-dominant asymmetry where the protocol allows it, and CoP FFT for spectral inspection.

Archery also includes full demographic analysis for a reference population: Mann-Whitney and Kruskal-Wallis comparisons, Pearson/Spearman correlations, per-subgroup percentiles, and per-trial score linkage.

Output is an Excel file (summary, per-trial, demographics), per-athlete and group PDFs with ellipses and stabilograms, and an optional local web dashboard.

## Data layout

One folder per recording session. Each athlete's files have the athlete ID in the filename.

**FMS Bipodal / Single-leg Stance** - `dir_*.xls` and `esq_*.xls` for right and left foot, with `inicio_fim.xlsx` holding the time windows per trial.

**Pistol Shooting** - `trial{distance}_{trial} - {date} - Stability export.xls`, with `inicio_fim.xlsx` providing two windows per trial (aiming period and trigger period).

**Archery** - `{id}_{trial} - {DD-MM-YYYY} - Stability export.xls` at 50 Hz, `Inicio_fim_vfinal.xlsx` with three sheets (`tempo do toque`, `confirmação_1`, `confirmação_2`), and optionally the demographic reference file. The analysis window is `confirmação_1` to `confirmação_2`. cp1252-mangled sheet names are handled automatically.

## Citing

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Version 1.0) [Computer software]. https://github.com/andremassuca/BSP

```bibtex
@software{massuca_bsp_2026,
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
