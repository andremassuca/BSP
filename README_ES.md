<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

**Qué es:** una aplicación para analizar la estabilidad postural de atletas a partir de datos de plataforma de fuerzas. La hice porque había demasiado trabajo manual en Excel que podía automatizarse y devolver métricas limpias.

**Para quién:** investigadores, entrenadores y clínicos que usan plataforma de fuerzas y quieren pasar del CoP bruto a métricas útiles (ea95, RMS, velocidad media, rigidez postural) sin escribir código.

## Qué hace

Cuatro protocolos:

- **FMS Bipodal** - apoyo bipodal simple con paneles izquierdo/derecho opcionales.
- **Apoyo Unipodal** - asimetría dominante / no-dominante.
- **Tiro** - dos ventanas temporales de estabilización antes del disparo.
- **Tiro con Arco** - hasta 30 ensayos, ventana única entre confirmación 1 y 2.

En cualquier protocolo calcula:

- Área de la elipse de 95% (`ea95`), RMS y amplitudes en X/Y y radial.
- Velocidad media e índice de rigidez postural (`stiff_x`, `stiff_y`).
- Asimetría dominante / no-dominante cuando corresponda.
- FFT del CoP para inspección espectral.

Genera: Excel con resumen, por ensayo y demografía (cuando hay referencia demográfica), PDF individual por atleta y PDF de grupo, y dashboard web local opcional con listas interactivas, box plots, scatter con regresión y detalle por atleta.

## Inicio en 1 minuto

```bash
git clone https://github.com/andremassuca/BSP.git
cd BSP
pip install -r requirements.txt
python estabilidade_gui.py
```

La primera ejecución abre la ventana de licencia → contraseña → selector de protocolo. La contraseña está en el documento de acceso.

Para ejecutar los tests sintéticos antes de procesar datos reales:

```bash
python estabilidade_gui.py --testes     # GUI/PDF/Excel (~161 tests)
python -m bsp_core --testes             # sólo núcleo matemático (~41 tests)
```

Ambos deben reportar **100% pasado**.

## Trabajar con los datos

**Convención de carpetas:** una carpeta por sesión. Cada atleta tiene su ID en el nombre del archivo.

**Tiro con Arco**

- Carpeta con ZIPs o subcarpetas; archivos `{id}_{ensayo} - {DD-MM-AAAA} - Stability export.xls` (tab-separated, 50 Hz).
- `Inicio_fim_vfinal.xlsx` con tres hojas (`tempo do toque`, `confirmação_1`, `confirmação_2`). La ventana de análisis es `confirmação_1 → confirmação_2`. El loader tolera nombres con cp1252 corrupto.
- `Todos os registos dos 142 atletas em JUl_2024 _.xlsx` como referencia demográfica.

**Otros protocolos**

- Archivos exportados por el software de la plataforma (`.xlsx` o `.txt` tab-separated).
- `inicio_fim.xlsx` para Tiro, con las dos ventanas temporales por ensayo.


## Problemas conocidos

- **macOS, primera ejecución:** Gatekeeper bloquea apps no notarizadas. **Clic derecho → Abrir**, confirmar. Sólo una vez.
- **macOS, la app no arranca:** revisar `~/Library/Logs/BSP/BSP_crash.log`. Si está vacío, correr `./BUILD_macOS.sh` desde terminal y leer `build_logs/`.
- **Windows, DLL ausente:** instalar [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe).
- **Dashboard no abre:** uvicorn tarda 1–3 s en arrancar. Si tras 10 s no aparece, verificar que `fastapi`/`uvicorn` están instalados.

## Citar

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Versión 1.0) [Software]. https://github.com/andremassuca/BSP

```bibtex
@software{massuca_bsp_2025,
  author  = {Massu\c{c}a, Andr\'{e} Oliveira and Aleixo, Pedro and Massu\c{c}a, Lu\'{i}s M.},
  title   = {BSP: Biomechanical Stability Program},
  version = {1.0},
  year    = {2026},
  url     = {https://github.com/andremassuca/BSP},
}
```

## Licencia

Código académico, libre para investigación y docencia. Para uso comercial o redistribución, contactar con el autor.

**André O. Massuça** - [github.com/andremassuca](https://github.com/andremassuca)
**Pedro Aleixo**
**Luís M. Massuça**
