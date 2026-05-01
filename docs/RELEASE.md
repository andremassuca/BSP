# Procedimento de release - BSP

Este documento descreve como publicar uma nova versão do BSP. O formato dos assets está acordado com o auto-update do cliente, por isso não inventar nomes de ficheiros.

## 0. Pré-checks

Tudo tem de estar verde antes de começar:

```bash
python estabilidade_gui.py --testes   # 161/161
python -m bsp_core --testes           # 41/41
```

Em macOS ARM e Windows, abrir a versão anterior instalada e confirmar que nada regrediu em uso real (abrir projecto, gerar PDF, etc.).

## 1. Bump de versão

Actualizar:

- `VERSION` → novo número (ex: `1.0.1`).
- `estabilidade_gui.py` → `VERSAO = "1.0.1"` e qualquer header `v1.0 → v1.0.1`.
- `bsp_core.py` → `VERSAO = "1.0.1"`.
- `BUILD_Windows.bat`, `BUILD_macOS.sh` → headers com a versão.
- `bsp_installer.py`, `bsp_uninstaller.py` → headers.
- `docs/CHANGELOG.md` → nova secção "v1.0.1 - {descrição}".

Commit: `chore: bump to v1.0.1`.

## 2. Build

**Windows** (VM ou máquina Windows):

```bat
BUILD_Windows.bat
```

Produz `dist\BSP_Setup.exe` (ou `BSP-1.0.1-win-x64.exe`, consoante o script).

**macOS** (ARM e Intel separadamente - o DMG é arch-specific):

```bash
./BUILD_macOS.sh
```

Produz `dist/BSP-1.0.1-macos-arm64.dmg` (ou `-x64.dmg`).

**Verificar em conta limpa:**

- Windows: instalar num user sem admin elevado; abrir; gerar um PDF.
- macOS: montar DMG em user diferente; right-click → Abrir; gerar um PDF.

## 3. SHA256SUMS

Na pasta onde estão os builds (idealmente reuni-os numa pasta `releases/v1.0.1/`):

```bash
# macOS / Linux
shasum -a 256 BSP_Setup.exe BSP-1.0.1-macos-arm64.dmg BSP-1.0.1-macos-x64.dmg \
  > SHA256SUMS.txt

# Windows PowerShell
Get-FileHash -Algorithm SHA256 BSP_Setup.exe, BSP-1.0.1-macos-arm64.dmg |
  ForEach-Object { "$($_.Hash.ToLower())  $($_.Path | Split-Path -Leaf)" } |
  Out-File -Encoding ASCII SHA256SUMS.txt
```

Confirmar que `SHA256SUMS.txt` tem formato `<hash>  <filename>` (dois espaços, um filename por linha, hash em minúsculas). O cliente auto-update usa este formato literalmente.

## 4. Tag e release

```bash
git tag -a v1.0.1 -m "BSP v1.0.1"
git push origin v1.0.1

gh release create v1.0.1 \
    --title "BSP v1.0.1" \
    --notes-file docs/CHANGELOG.md \
    BSP_Setup.exe \
    BSP-1.0.1-macos-arm64.dmg \
    BSP-1.0.1-macos-x64.dmg \
    SHA256SUMS.txt
```

O campo `--notes-file` usa a primeira secção do CHANGELOG (é o que aparece no banner "Ver notas"). Se preferir notas curtas à parte, usar `--notes "texto aqui"`.

## 5. Smoke post-release

Numa máquina com versão anterior do BSP instalada:

1. Abrir a app. O banner "Versão v1.0.1 disponível" deve aparecer ~5 s após o arranque.
2. Clicar **Ver notas** → confirmar que o conteúdo do CHANGELOG aparece.
3. Clicar **Actualizar agora** → barra de progresso → download → SHA256 verifica → instalador abre.
4. Confirmar que a nova versão arranca.

Cenário negativo (testar uma vez, em fork):

- Publicar um release com SHA256SUMS.txt incorrecto.
- Cliente vai tentar o download e abortar com "Checksum nao corresponde". Banner fica como está, sem lançar instalador.

## 6. Se algo corre mal

- Se o release foi publicado com binário errado: `gh release delete v1.0.1 --yes` + `git push --delete origin v1.0.1`. Nunca re-utilizar a mesma tag sem apagar primeiro - o cliente pode ter cached o release antigo em memória.
- Se só o SHA256SUMS.txt estava errado: basta editar o release no GitHub, apagar o asset errado, upload do correcto.

## Notas sobre convenção de nomes

O cliente auto-update identifica o asset por regex (ver `_asset_nome_esperado` em [estabilidade_gui.py](../estabilidade_gui.py)):

- Windows: `BSP_Setup*.exe` ou `BSP-<ver>-win-x64.exe`.
- macOS ARM: `BSP-<ver>-macos-arm64.dmg`.
- macOS Intel: `BSP-<ver>-macos-x64.dmg`.

Qualquer outro nome não é detectado e o cliente fica com o banner mas sem "Actualizar agora" clicável (cai para "Ver notas" que leva ao GitHub).
