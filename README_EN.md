<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

**What it is:** an app for analyzing athletes' postural stability from force-plate data. Built because too much of this was being done by hand in Excel - this turns raw CoP into clean metrics without writing code.

**Who it's for:** researchers, coaches and clinicians using force plates who want to move past raw CoP into usable metrics (ea95, RMS, mean velocity, postural stiffness).

## What it does

Four protocols:

- **FMS Bipodal** - simple two-foot stance with optional left/right panels.
- **Single-leg Stance** - dominant / non-dominant asymmetry.
- **Pistol Shooting** - two stabilization windows before trigger pull.
- **Archery** - up to 30 trials, single window between confirmation 1 and 2.

In any protocol it computes:

- 95% confidence ellipse area (`ea95`), RMS and amplitudes in X/Y and radial.
- Mean velocity and postural stiffness index (`stiff_x`, `stiff_y`).
- Dominant / non-dominant asymmetry when applicable.
- CoP FFT for spectral inspection.

Outputs: Excel with summary, per-trial, demographics (when a reference file is provided), per-athlete and group PDFs, and an optional local web dashboard with interactive lists, box plots, scatter with regression, per-athlete detail.

## 1-minute quickstart

```bash
git clone https://github.com/andremassuca/BSP.git
cd BSP
pip install -r requirements.txt
python estabilidade_gui.py
```

First run opens the license window → password → protocol picker. The password is in the access document sent separately.

Run the synthetic test suite before processing real data:

```bash
python estabilidade_gui.py --testes     # GUI/PDF/Excel (~161 tests)
python -m bsp_core --testes             # math core only (~41 tests)
```

Both must report **100% pass**.

## Working with data

**Folder convention:** one folder per session. Each athlete has its ID in the filename.

**Archery**

- Folder with ZIPs or subfolders; filenames `{id}_{trial} - {DD-MM-YYYY} - Stability export.xls` (tab-separated, 50 Hz).
- `Inicio_fim_vfinal.xlsx` with three sheets (`tempo do toque`, `confirmação_1`, `confirmação_2`). The analysis window is `confirmação_1 → confirmação_2`. The loader handles cp1252-mangled names (`confirma��o_1`) automatically.
- `Todos os registos dos 142 atletas em JUl_2024 _.xlsx` as demographic reference. Columns PESO, ALTURA, IDADE, ESTILO, CATEGORIA, GÉNERO, P1..P30 and P_TOTAL. Columns 68–308 are ignored (manual data).

**Other protocols**

- Files exported by the platform software, `.xlsx` or tab-separated `.txt`.
- `inicio_fim.xlsx` for Pistol Shooting, with the two temporal windows per trial.


## Known issues

- **macOS, first run:** Gatekeeper blocks unnotarized apps. **Right-click → Open**, confirm. One-time only.
- **macOS, app won't start:** check `~/Library/Logs/BSP/BSP_crash.log`. If empty, run `./BUILD_macOS.sh` from terminal and read `build_logs/`.
- **Windows, missing DLL:** install [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe).
- **Confirmação_1/2 won't load:** check that sheet names aren't corrupted. Open, rename to `confirmação_1` / `confirmação_2`, save.
- **Dashboard doesn't open:** uvicorn takes 1–3 s to start. If nothing appears after 10 s, verify `fastapi`/`uvicorn` are installed.

## Cite

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

Relevant methods:

- Schubert, P., & Kirchner, M. (2013). *Ellipse area calculations and their applicability in posturography.* Gait & Posture, 39(1), 518–522.
- Prieto, T. E. et al. (1996). *Measures of postural steadiness.* IEEE Trans. Biomed. Eng., 43(9), 956–966.
- Winter, D. A. (1995). *Human balance and posture control during standing and walking.* Gait & Posture, 3(4), 193–214.
- Quijoux, F. et al. (2021). *A review of center of pressure (COP) variables to quantify standing balance in elderly people.* Physiological Reports, 9(22), e15067.
- Maurer, C., & Peterka, R. J. (2005). *A new interpretation of spontaneous sway measures based on a simple model of human postural control.* J. Neurophysiol., 93(1), 189–200.
- Carpenter, M. G. et al. (2001). *Sampling duration effects on centre of pressure summary measures.* Exp. Brain Res., 138(2), 210–218.

## License

Academic code, free for research and teaching. For commercial use or redistribution, contact the author.

**André O. Massuça** - [github.com/andremassuca](https://github.com/andremassuca)
**Pedro Aleixo**
**Luís M. Massuça**
