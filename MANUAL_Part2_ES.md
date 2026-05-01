# BSP - Biomechanical Stability Program
## Manual Completo v1.0

**Autores:** Andre Massuca (desarrollo) | Pedro Aleixo (biomecanica) | Luis Massuca (coordinacion)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Requisitos e Instalacion

| Sistema | Version minima |
|---|---|
| Windows | 10 / 11 |
| macOS | 12 Monterey |
| Python | 3.9+ |

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

---

## 2. Novedades en v23

| Funcionalidad | Descripcion |
|---|---|
| Crash silencioso resuelto | Errores en hilo de analisis muestran ventana de diagnostico |
| RMS en PDF | RMS ML, AP, Radio (Quijoux et al., 2021) en tabla individual |
| Correccion tablas HurdleStep | Cabeceras ya no se superponen |
| Entradas ilimitadas | Tablas PDF auto-adaptan columnas y fuente |
| Validacion frecuencia muestreo | Detecta timestamps duplicados, jitter >20%, variacion >10% |
| PNG con opciones | DPI configurable (72/150/180/300) y seleccion de tipos |
| Pagina novedades en PDF | Generada automaticamente |
| Citaciones dinamicas | Ano APA y BibTeX usa el ano actual del sistema |
| Autoria actualizada | Andre Massuca & Luis Massuca |

---

## 3. Estructura de Ficheros

### FMS / Unipodal
```
carpeta_principal/
    001_Nombre/
        dir_1.xls, dir_2.xls   (pie derecho)
        esq_1.xls, esq_2.xls   (pie izquierdo)
```

### Tiro
```
carpeta_principal/
    001_Tirador/
        trial5_1_-_fecha_-_Stability_export.xls
        hs_dir_1.xls   (Hurdle Step, opcional)
```

---

## 4. Protocolos Disponibles

- **FMS Bipodal:** 2 lados, N ensayos, elipse 95%, índice de asimetría Dir/Izq
- **Apoyo Unipodal:** 2 lados, N ensayos, RMS ML/AP
- **Tiro:** plato completo + seleccion de pie, multi-distancia, Friedman
- **Tiro con Arco:** equivalente al Tiro

---

## 5. Opciones de Salida

- Resumen Excel (DATOS + GRUPO + SPSS + ESTATS)
- Ficheros individuales (uno por sujeto)
- Informe PDF (portada + novedades + leyenda + sujetos + citacion)
- Informe Word (.docx, requiere python-docx)
- Informe HTML interactivo (Chart.js standalone)
- CSV para SPSS / R
- PNG (DPI configurable)

---

## 6. Validacion de Frecuencia de Muestreo (v23)

| Verificacion | Umbral | Accion |
|---|---|---|
| Timestamps duplicados/invertidos | Cualquier ocurrencia | Aviso en log |
| Jitter dentro del ensayo | >20% variacion relativa | Aviso en log |
| Variacion entre ensayos | >10% entre ensayos | Aviso en log |

---

## 7. Idiomas

PT / EN / ES / DE con 336 claves de traduccion. Todos los outputs (PDF, Excel, HTML)
se adaptan al idioma seleccionado: metricas, lados, EULA.

---

## 8. Compilar para Distribucion

**Windows:** `BUILD_Windows.bat` -> instala en `%LOCALAPPDATA%\Programs\BSP\`
**macOS:** `./BUILD_macOS.sh` -> `BSP.dmg` compatible con macOS 12+ (Intel y Apple Silicon)

---

## 9. Solucion de Problemas

| Problema | Solucion |
|---|---|
| Dependencias no encontradas | `pip install numpy scipy openpyxl matplotlib reportlab` |
| Advertencia seguridad macOS | Boton derecho -> Abrir |
| PDF no generado | Verificar ruta y opcion activa |
| Crash silencioso (v22) | Actualizar a v23 |
| > 10 ensayos | Sin limite en v23 |
| Aviso frecuencia muestreo | Verificar exportacion de la plataforma de fuerzas |

---

## Citacion Academica

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
