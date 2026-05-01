"""
BSP branding - gerar bsp_icon.ico e bsp_icon.icns
a partir de bsp_icon_1024.png (gerado pelo Brand Guide HTML).

Dependências:
    pip install pillow icnsutil

Uso:
    cd branding/
    python make_ico_icns.py

Coloca os ficheiros gerados em branding/ e depois corre _apply.py.
"""

from pathlib import Path
import sys

try:
    from PIL import Image
except ImportError:
    sys.exit("Instala Pillow primeiro:  pip install pillow")

SRC  = Path(__file__).parent / "bsp_icon_1024.png"
ICO  = Path(__file__).parent / "bsp_icon.ico"
ICNS = Path(__file__).parent / "bsp_icon.icns"
ICONSET = Path(__file__).parent / "bsp_icon.iconset"

ICO_SIZES  = [16, 32, 48, 64, 128, 256]
ICNS_SIZES = [16, 32, 64, 128, 256, 512, 1024]  # + @2x via iconutil

# ── ICO ───────────────────────────────────────────────────────────────────────

def make_ico(src: Path, dst: Path):
    img = Image.open(src).convert("RGBA")
    frames = [img.resize((s, s), Image.LANCZOS) for s in ICO_SIZES]
    frames[0].save(
        dst,
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
        append_images=frames[1:],
    )
    print(f"[OK] {dst.name}  ({', '.join(str(s) for s in ICO_SIZES)} px)")


# ── ICNS (via iconutil - macOS only) ─────────────────────────────────────────

def make_iconset(src: Path, iconset_dir: Path):
    """Cria pasta .iconset com os PNGs necessários para iconutil."""
    iconset_dir.mkdir(exist_ok=True)
    img = Image.open(src).convert("RGBA")

    mapping = {
        "icon_16x16.png":       16,
        "icon_16x16@2x.png":    32,
        "icon_32x32.png":       32,
        "icon_32x32@2x.png":    64,
        "icon_64x64.png":       64,
        "icon_64x64@2x.png":    128,
        "icon_128x128.png":     128,
        "icon_128x128@2x.png":  256,
        "icon_256x256.png":     256,
        "icon_256x256@2x.png":  512,
        "icon_512x512.png":     512,
        "icon_512x512@2x.png":  1024,
    }
    for name, size in mapping.items():
        resized = img.resize((size, size), Image.LANCZOS)
        resized.save(iconset_dir / name, format="PNG")
    print(f"[OK] {iconset_dir.name}/  ({len(mapping)} ficheiros)")


def make_icns_via_iconutil(iconset_dir: Path, dst: Path):
    import subprocess
    result = subprocess.run(
        ["iconutil", "-c", "icns", str(iconset_dir), "-o", str(dst)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"[ERRO] iconutil: {result.stderr.strip()}")
        print("       Isto so funciona em macOS com Xcode Command Line Tools.")
    else:
        print(f"[OK] {dst.name}")


def make_icns_pillow_fallback(src: Path, dst: Path):
    """Fallback simples para Windows/Linux usando icnsutil (pip)."""
    try:
        import icnsutil
    except ImportError:
        print("[INFO] icnsutil nao instalado. Instala com:  pip install icnsutil")
        print("       Em macOS, usa make_iconset() + iconutil (sem dependencias extra).")
        return

    img = Image.open(src).convert("RGBA")
    ic = icnsutil.IcnsFile()
    size_map = {
        "ic04": 16,   # ic04 = 16x16
        "ic05": 32,   # ic05 = 32x32  (16@2x)
        "ic07": 128,
        "ic08": 256,
        "ic09": 512,
        "ic10": 1024,
        "ic11": 32,
        "ic12": 64,
        "ic13": 256,
        "ic14": 512,
    }
    added = set()
    for key, size in size_map.items():
        if size in added:
            continue
        resized = img.resize((size, size), Image.LANCZOS)
        from io import BytesIO
        buf = BytesIO()
        resized.save(buf, "PNG")
        try:
            ic.add_media(key, data=buf.getvalue())
        except Exception:
            pass
        added.add(size)

    # icnsutil >= 1.x espera path; versoes antigas aceitavam handle
    try:
        ic.write(str(dst))
    except (TypeError, AttributeError):
        with open(dst, "wb") as f:
            ic.write(f)
    print(f"[OK] {dst.name}  (via icnsutil)")


# ── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not SRC.exists():
        sys.exit(
            f"Ficheiro nao encontrado: {SRC}\n"
            "Exporta primeiro o bsp_icon_1024.png no Brand Guide HTML."
        )

    print(f"\nSource: {SRC}  ({Image.open(SRC).size[0]}x{Image.open(SRC).size[1]})\n")

    # ICO (funciona em qualquer OS)
    make_ico(SRC, ICO)

    # ICNS
    import platform
    if platform.system() == "Darwin":
        make_iconset(SRC, ICONSET)
        make_icns_via_iconutil(ICONSET, ICNS)
        # Limpa iconset temporario
        import shutil
        shutil.rmtree(ICONSET)
        print(f"[OK] iconset temporario removido")
    else:
        print("\n[INFO] Nao estas em macOS - a tentar icnsutil (pip install icnsutil)...")
        make_icns_pillow_fallback(SRC, ICNS)

    print(f"\nFicheiros gerados em: {Path(__file__).parent.resolve()}")
    print("Corre _apply.py para copiar para os destinos finais.\n")
