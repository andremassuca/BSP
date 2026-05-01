# BSP - Biomechanical Stability Program
## Manual de Referencia Rapida v1.0

**Autores:** Andre Massuca (desarrollo) | Pedro Aleixo (biomecanica) | Luis Massuca (coordinacion)
**GitHub:** https://github.com/andremassuca/BSP

---

## 1. Requisitos e Instalacion

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

**Windows:** `BUILD_Windows.bat` -> `BSP_Setup.exe`
**macOS:** `./BUILD_macOS.sh` -> `BSP.dmg`

Compatibilidad: Windows 10/11, macOS 12 Monterey+, Python 3.9+

---

## 2. Protocolos

| Protocolo | Descripcion |
|---|---|
| FMS Bipodal | N ensayos por pie, Elipse 95%, asimetria |
| Apoyo Unipodal | N ensayos por pie, oscilacion lateral |
| Tiro (ISCPSI) | Multi-distancia, multi-intervalo, correlacion precision |
| Tiro con Arco | Hasta 30 ensayos bipodal, ventana unica (Confirmacion 1 a 2), sin distancia |

Numero ilimitado de sujetos y ensayos. Las tablas PDF se adaptan automaticamente.

---

## 3. Estructura de Ficheros

```
carpeta_principal/
    001_Nombre/
        dir_1.xls, dir_2.xls, ...
        esq_1.xls, esq_2.xls, ...
    002_OtroNombre/
        ...
```

---

## 4. Novedades en v23

- Crash silencioso resuelto (hilo de analisis)
- RMS ML, AP, Radio anadidos al informe PDF
- Validacion de frecuencia de muestreo entre ensayos
- Exportacion PNG con DPI configurable (72/150/180/300)
- Pagina "Novedades" en el PDF generada automaticamente
- Entradas ilimitadas - tablas PDF se adaptan automaticamente
- Autoria actualizada: Andre Massuca, Pedro Aleixo & Luis Massuca
- Citaciones con ano dinamico del sistema

---

## 5. Atajos de Teclado

| Windows | macOS | Accion |
|---|---|---|
| Ctrl+Enter | Cmd+Enter | Ejecutar analisis |
| Ctrl+S | Cmd+S | Guardar configuracion |
| F1 | F1 | Ver atajos |
| F2 | F2 | Analisis rapido |

---

## 6. Solucion de Problemas

| Problema | Solucion |
|---|---|
| `ModuleNotFoundError` | Instalar dependencias con pip |
| Aviso seguridad macOS | Boton derecho -> Abrir |
| Crash silencioso (v22) | Actualizar a v23 |
| > 10 ensayos | Sin limite - adaptacion automatica |
| Aviso frecuencia muestreo | Verificar exportacion de plataforma de fuerzas |

---

## Citacion Academica

```
Massuca, A., & Massuca, L. (2026). BSP - Biomechanical Stability Program (v23).
https://github.com/andremassuca/BSP
```

*BSP v23 - Andre Massuca & Luis Massuca*
