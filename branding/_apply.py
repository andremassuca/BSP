#!/usr/bin/env python3
"""Aplicar os assets de branding aos sítios finais.

Quando se mete um asset novo em branding/ corre-se isto e ele:
- copia para a raiz (BSP.ico, BSP_banner.png, etc.)
- re-encoda o banner em base64 dentro de estabilidade_gui.py
- diz o que falta

Idempotente. Se um ficheiro fonte não existe, deixa o destino como está.
"""
from __future__ import annotations
import base64
import hashlib
import os
import re
import shutil
import sys
from pathlib import Path

# Estrutura esperada:
#   <root>/                         <-- raiz do projecto
#       branding/                   <-- esta pasta
#           bsp_icon.ico
#           bsp_banner_dark.png
#           bsp_banner_light.png
#           bsp_icon_1024.png
#           bsp_pdf_cover.png
#       BSP.ico
#       BSP_banner.png
#       BSP_icon_1024.png
#       estabilidade_gui.py

ROOT = Path(__file__).resolve().parent.parent
BRAND = ROOT / 'branding'

# Source -> destination (relativo a ROOT)
COPIES = [
    ('bsp_icon.ico',         'BSP.ico'),
    ('bsp_icon.icns',        'BSP.icns'),
    ('bsp_banner_dark.png',  'BSP_banner.png'),
    ('bsp_icon_1024.png',    'BSP_icon_1024.png'),
    ('bsp_pdf_cover.png',    'BSP_pdf_cover.png'),
]


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def step(n, msg):
    print(f'\n[{n}] {msg}')


def copy_assets() -> list[str]:
    """Copia os ficheiros de branding/ para a raiz se mudaram."""
    changed = []
    for src_name, dst_name in COPIES:
        src = BRAND / src_name
        dst = ROOT / dst_name
        if not src.exists():
            print(f'  - skip {src_name} (nao existe em branding/)')
            continue
        if dst.exists() and md5(src) == md5(dst):
            print(f'  - {dst_name}: igual, sem alteracao')
            continue
        shutil.copy2(src, dst)
        size_kb = dst.stat().st_size / 1024
        print(f'  + {dst_name}: {size_kb:.1f} KB')
        changed.append(dst_name)
    return changed


def reencode_banner_in_gui() -> bool:
    """Re-encoda o banner em base64 dentro de estabilidade_gui.py.

    Procura a constante _BANNER_AOM = b\"\"\"...\"\"\" (ou similar) e
    substitui pelo conteudo do branding/bsp_banner_dark.png.
    Se a constante nao existe ou o banner-dark nao foi gerado, sai sem
    alterar nada.
    """
    src = BRAND / 'bsp_banner_dark.png'
    if not src.exists():
        print('  - skip (bsp_banner_dark.png nao existe)')
        return False
    gui = ROOT / 'estabilidade_gui.py'
    if not gui.exists():
        print('  - skip (estabilidade_gui.py nao encontrado)')
        return False
    txt = gui.read_text(encoding='utf-8')
    b64 = base64.b64encode(src.read_bytes()).decode('ascii')
    # Quebrar em linhas de 76 chars como costume Python
    lines = '\n'.join(b64[i:i+76] for i in range(0, len(b64), 76))
    # Procurar a constante actual:
    # 1. b''' / b""" multi-line (preferred forward)
    # 2. ''' / """ multi-line
    # 3. "..." ou '...' single-line (formato actual em v1.0)
    pat = re.compile(
        r"(_BANNER_(?:AOM|BSP)\s*=\s*)(b?'''.*?'''|b?\"\"\".*?\"\"\")",
        re.DOTALL,
    )
    m = pat.search(txt)
    if not m:
        # Single-line string com aspas simples ou duplas (suporta backslash escapes)
        pat2 = re.compile(
            r"(_BANNER_(?:AOM|BSP)\s*=\s*)([\"'])((?:\\.|(?!\2).)*)\2",
            re.DOTALL,
        )
        m = pat2.search(txt)
    if not m:
        # Forma agrupada com parens
        pat3 = re.compile(
            r"(_BANNER_(?:AOM|BSP)\s*=\s*)(\([^)]*\))",
            re.DOTALL,
        )
        m = pat3.search(txt)
    if not m:
        print('  - aviso: nao consegui localizar _BANNER_AOM/BSP em estabilidade_gui.py')
        print('         (banner sera lido apenas via BSP_banner.png em runtime)')
        return False
    # Re-escrever como string normal numa linha (compativel com formato actual)
    novo = f'{m.group(1)}"{b64}"'
    txt2 = txt[:m.start()] + novo + txt[m.end():]
    if txt2 != txt:
        gui.write_text(txt2, encoding='utf-8')
        print(f'  + banner re-encodado dentro de estabilidade_gui.py ({len(b64):,} chars b64)')
        return True
    print('  - banner inalterado (mesmo conteudo)')
    return False


def reencode_logo_in_gui() -> bool:
    """Re-encoda o logo 1024 em base64 dentro de estabilidade_gui.py (_LOGO_B64)."""
    src = BRAND / 'bsp_logo_1024.png'
    if not src.exists():
        print('  - skip (bsp_logo_1024.png nao existe)')
        return False
    gui = ROOT / 'estabilidade_gui.py'
    if not gui.exists():
        print('  - skip (estabilidade_gui.py nao encontrado)')
        return False
    import re as _re
    txt = gui.read_text(encoding='utf-8')
    b64 = base64.b64encode(src.read_bytes()).decode('ascii')
    pat = _re.compile(
        r'(# Logo BSP \(.*?\)\n_LOGO_B64 = \()[\s\S]*?(\)\n)',
        _re.DOTALL,
    )
    m = pat.search(txt)
    if not m:
        print('  - aviso: nao consegui localizar _LOGO_B64 em estabilidade_gui.py')
        return False
    novo = f'{m.group(1)}\n    "{b64}"\n{m.group(2)}'
    txt2 = txt[:m.start()] + novo + txt[m.end():]
    if txt2 != txt:
        gui.write_text(txt2, encoding='utf-8')
        print(f'  + logo re-encodado dentro de estabilidade_gui.py ({len(b64):,} chars b64)')
        return True
    print('  - logo inalterado (mesmo conteudo)')
    return False


def report_missing():
    """Lista os assets esperados que ainda nao foram fornecidos."""
    expected = [
        'bsp_logo.svg',
        'bsp_logo_1024.png',
        'bsp_icon.ico',
        'bsp_icon.icns',
        'bsp_icon_1024.png',
        'bsp_banner_dark.png',
        'bsp_banner_light.png',
        'bsp_banner_dark@2x.png',
        'bsp_banner_light@2x.png',
        'bsp_pdf_cover.png',
        'bsp_installer_banner.png',
    ]
    missing = [name for name in expected if not (BRAND / name).exists()]
    if missing:
        print('\nAssets em falta (meter em branding/ com este nome):')
        for n in missing:
            print(f'  - {n}')
    else:
        print('\nTodos os assets esperados estao presentes.')


def main() -> int:
    if not BRAND.exists():
        print(f'[ERRO] pasta branding nao existe: {BRAND}')
        return 2

    step(1, 'Copiar assets para a raiz')
    changed = copy_assets()

    step(2, 'Re-encodar banner dentro de estabilidade_gui.py')
    if 'BSP_banner.png' in changed:
        reencode_banner_in_gui()
    else:
        print('  - skip (banner nao mudou)')

    step(3, 'Re-encodar logo dentro de estabilidade_gui.py')
    if 'BSP_icon_1024.png' in changed or (BRAND / 'bsp_logo_1024.png').exists():
        reencode_logo_in_gui()
    else:
        print('  - skip (logo nao mudou)')

    step(4, 'Verificar assets em falta')
    report_missing()

    step(5, 'Sugerir proximos passos')
    if changed:
        print('  -> correr os testes:')
        print('     PYTHONUTF8=1 python estabilidade_gui.py --testes')
        print('  -> regerar o ZIP de distribuicao:')
        print('     python ../_make_zip.py   (se existir)')
    else:
        print('  - nada mudou; nada a fazer')

    return 0


if __name__ == '__main__':
    sys.exit(main())
