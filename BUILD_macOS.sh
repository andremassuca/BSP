#!/bin/bash
# ============================================================
#  BSP - Biomechanical Stability Program  v1.0
#  Gera BSP.dmg para macOS - abre e arrasta para Applications
#  Uso: chmod +x BUILD_macOS.sh && ./BUILD_macOS.sh
# ============================================================
#
# Se algo falhar, o log detalhado de cada passo fica em:
#   ./build_logs/01_tests.log
#   ./build_logs/02_pip.log
#   ./build_logs/03_icon.log
#   ./build_logs/04_pyinstaller.log
#   ./build_logs/05_smoke.log
#   ./build_logs/06_codesign.log
#   ./build_logs/07_dmg.log
# ============================================================

set -euo pipefail
cd "$(dirname "$0")"

mkdir -p build_logs

V='\033[0;32m'; Y='\033[0;33m'; R='\033[0;31m'; B='\033[0;34m'
BOLD='\033[1m'; RST='\033[0m'
ok()   { echo -e "  ${V}✓${RST}  $1"; }
fail() { echo -e "  ${R}✗${RST}  $1"; exit 1; }
info() { echo -e "  ${B}·${RST}  $1"; }
warn() { echo -e "  ${Y}⚠${RST}  $1"; }
hdr()  { echo -e "\n${BOLD}${B}$1${RST}"; }

echo ""
echo -e "${BOLD}${B}============================================================${RST}"
echo -e "${BOLD}${B}  BSP - Biomechanical Stability Program  v1.0${RST}"
echo -e "${BOLD}${B}  A gerar BSP.dmg para macOS${RST}"
echo -e "${BOLD}${B}  Logs em: $(pwd)/build_logs/${RST}"
echo -e "${BOLD}${B}============================================================${RST}"

# ── Python (inclui 3.14) ────────────────────────────────────
PYTHON=""
for cmd in python3.14 python3.13 python3.12 python3.11 python3.10 python3.9 python3 python; do
    if command -v "$cmd" &>/dev/null; then
        VER=$("$cmd" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null || echo "0.0")
        MAJ=$(echo "$VER"|cut -d. -f1); MIN=$(echo "$VER"|cut -d. -f2)
        if [ "$MAJ" -ge 3 ] && [ "$MIN" -ge 9 ]; then
            PYTHON="$cmd"; ok "Python $VER ($PYTHON)"; break
        fi
    fi
done
[ -z "$PYTHON" ] && fail "Python 3.9+ não encontrado - instalar em https://www.python.org/downloads/"

# Arquitectura
ARCH="$(uname -m)"
info "Arquitectura: $ARCH"

# ── Garantir nome correcto do ícone ──────────────────────────
[ -f "AOM.ico" ] && [ ! -f "BSP.ico" ] && cp AOM.ico BSP.ico
[ -f "AOM_icon_1024.png" ] && [ ! -f "BSP_icon_1024.png" ] && cp AOM_icon_1024.png BSP_icon_1024.png

# ── Testes ───────────────────────────────────────────────────
hdr "[1/7]  A correr testes sintéticos..."
if $PYTHON -X utf8 estabilidade_gui.py --testes >build_logs/01_tests.log 2>&1; then
    ok "Testes passaram"
else
    warn "Alguns testes falharam - continuar (ver build_logs/01_tests.log)"
fi

# ── Dependências ────────────────────────────────────────────
hdr "[2/7]  A instalar dependências..."
{
    $PYTHON -m pip install pyinstaller --upgrade
    if [ -f "requirements.txt" ]; then
        $PYTHON -m pip install -r requirements.txt --upgrade
    else
        $PYTHON -m pip install numpy scipy openpyxl matplotlib \
            reportlab python-docx Pillow --upgrade
    fi
} >build_logs/02_pip.log 2>&1 || fail "pip falhou - ver build_logs/02_pip.log"
ok "Dependências instaladas"

# ── Opcional: PyArmor (anti-reverse-engineering) ────────────
# Activado se a env-var BSP_OBFUSCATE=1 ou se pyarmor estiver instalado.
# Cria build_obf/ com source ofuscado; PyInstaller usa esse em vez do .py limpo.
SOURCE_DIR="."
if [ "${BSP_OBFUSCATE:-0}" = "1" ] || $PYTHON -m pyarmor --version &>/dev/null; then
    hdr "[2.5/7]  PyArmor: a ofuscar source antes de PyInstaller..."
    rm -rf build_obf
    if $PYTHON -m pyarmor gen --output build_obf \
            estabilidade_gui.py bsp_core.py bsp_i18n.py >build_logs/02b_pyarmor.log 2>&1; then
        SOURCE_DIR="build_obf"
        ok "Source ofuscado em build_obf/"
    else
        warn "PyArmor falhou - continuar com source original (ver build_logs/02b_pyarmor.log)"
    fi
else
    info "PyArmor nao disponivel; build sem ofuscacao adicional"
    info "(pip install pyarmor para activar)"
fi

# ── Ícone .icns ─────────────────────────────────────────────
hdr "[3/7]  A preparar ícone..."
ICON_ARG=""
if [ -f "BSP.icns" ]; then
    # .icns ja foi gerado pelo branding/_apply.py - usar directamente
    ok "BSP.icns ja existe (gerado pelo pipeline de branding)"
    ICON_ARG="--icon BSP.icns"
elif [ -f "BSP_icon_1024.png" ] && command -v sips &>/dev/null && command -v iconutil &>/dev/null; then
    # Fallback: gerar a partir do PNG via sips/iconutil
    {
        rm -rf BSP.iconset BSP.icns
        mkdir BSP.iconset
        for size in 16 32 128 256 512; do
            sips -z $size $size BSP_icon_1024.png \
                --out "BSP.iconset/icon_${size}x${size}.png"
            sips -z $((size*2)) $((size*2)) BSP_icon_1024.png \
                --out "BSP.iconset/icon_${size}x${size}@2x.png"
        done
        iconutil -c icns BSP.iconset -o BSP.icns
        rm -rf BSP.iconset
    } >build_logs/03_icon.log 2>&1 || fail "sips/iconutil falhou - ver build_logs/03_icon.log"
    ok "BSP.icns gerado a partir de BSP_icon_1024.png"
    ICON_ARG="--icon BSP.icns"
else
    warn "BSP.icns e BSP_icon_1024.png ausentes - usar ícone por defeito"
fi

# ── PyInstaller ─────────────────────────────────────────────
# --onedir em vez de --onefile - evita extracção lenta em /tmp
# a cada execução no macOS (especialmente Apple Silicon).
hdr "[4/7]  A compilar BSP.app com PyInstaller..."
if ! $PYTHON -m PyInstaller \
    --onedir \
    --windowed \
    --name BSP \
    --osx-bundle-identifier com.aomassuca.bsp \
    --log-level INFO \
    $ICON_ARG \
    --hidden-import numpy \
    --hidden-import scipy \
    --hidden-import scipy.stats \
    --hidden-import scipy.signal \
    --hidden-import openpyxl \
    --hidden-import matplotlib \
    --hidden-import matplotlib.backends.backend_tkagg \
    --hidden-import reportlab \
    --hidden-import reportlab.pdfgen \
    --hidden-import reportlab.lib \
    --hidden-import reportlab.platypus \
    --hidden-import docx \
    --hidden-import PIL \
    --hidden-import tkinter \
    --hidden-import tkinter.ttk \
    --hidden-import tkinter.filedialog \
    --collect-all scipy \
    --collect-all matplotlib \
    --collect-all reportlab \
    estabilidade_gui.py >build_logs/04_pyinstaller.log 2>&1; then
    warn "Ver warn-BSP.txt:"
    [ -f "build/BSP/warn-BSP.txt" ] && tail -40 build/BSP/warn-BSP.txt
    fail "PyInstaller falhou - ver build_logs/04_pyinstaller.log"
fi

[ ! -d "dist/BSP.app" ] && fail "BSP.app não foi criado - ver build_logs/04_pyinstaller.log"
ok "BSP.app compilado"

# ── Embeber ícone no bundle ─────────────────────────────────
if [ -f "BSP.icns" ]; then
    cp BSP.icns "dist/BSP.app/Contents/Resources/BSP.icns" 2>/dev/null || true
    touch "dist/BSP.app" 2>/dev/null || true
fi

# ── Smoke test ──────────────────────────────────────────────
hdr "[5/7]  Smoke test (--testes) no bundle..."
BSP_BIN="dist/BSP.app/Contents/MacOS/BSP"
if [ -x "$BSP_BIN" ]; then
    if "$BSP_BIN" --testes >build_logs/05_smoke.log 2>&1; then
        ok "Bundle arranca e passa testes"
    else
        warn "Bundle arranca mas testes falharam - ver build_logs/05_smoke.log"
        # Tentar capturar traceback do log da app
        if [ -d "$HOME/Library/Logs/BSP" ]; then
            warn "Últimas 20 linhas do log da app:"
            tail -20 "$HOME/Library/Logs/BSP/"*.log 2>/dev/null || true
        fi
    fi
else
    warn "Binário $BSP_BIN não executável"
fi

# ── Code-sign ad-hoc ────────────────────────────────────────
# Evita "bundle broken" no Apple Silicon. Não é notarização
# (essa requer Apple Developer Account pago).
hdr "[6/7]  Code-sign ad-hoc..."
if command -v codesign &>/dev/null; then
    if codesign --force --deep --sign - "dist/BSP.app" >build_logs/06_codesign.log 2>&1; then
        ok "Assinatura ad-hoc aplicada"
    else
        warn "codesign falhou - ver build_logs/06_codesign.log (bundle pode abrir na mesma)"
    fi
else
    warn "codesign não disponível"
fi

# ── DMG ─────────────────────────────────────────────────────
hdr "[7/7]  A criar BSP.dmg..."
rm -rf dmg_stage BSP.dmg BSP_rw.dmg 2>/dev/null || true
mkdir -p dmg_stage
cp -r dist/BSP.app dmg_stage/
ln -s /Applications dmg_stage/Applications

{
    hdiutil create -volname "BSP v1.0" -srcfolder dmg_stage \
        -ov -format UDRW -o BSP_rw.dmg
    hdiutil convert BSP_rw.dmg -format UDZO -imagekey zlib-level=9 -o BSP.dmg
} >build_logs/07_dmg.log 2>&1 || fail "hdiutil falhou - ver build_logs/07_dmg.log"

rm -f BSP_rw.dmg
rm -rf dmg_stage
ok "BSP.dmg criado"

echo ""
echo -e "${BOLD}${V}============================================================${RST}"
echo -e "${BOLD}${V}  BUILD COMPLETO${RST}"
echo ""
ok "BSP.dmg - $(du -sh BSP.dmg 2>/dev/null | cut -f1)"
echo ""
info "Para instalar:"
info "  1. Duplo-clique BSP.dmg"
info "  2. Arrastar BSP.app para Applications"
info "  3. Ejectar o disco - abrir via Launchpad ou Spotlight"
echo ""
warn "Primeira execução: clique-direito em BSP.app > Abrir"
warn "Ou: Definições > Privacidade > 'Abrir mesmo assim'"
echo ""
info "Se a app não arranca: ver ~/Library/Logs/BSP/BSP_crash.log"
echo -e "${BOLD}${V}============================================================${RST}"
echo ""
