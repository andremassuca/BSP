# Changelog

## v1.0 - initial release

Primeira versão pública. Squash do histórico de desenvolvimento (v23 → v1.0).

### Novo

- **`bsp_core.py`** - módulo standalone com parsing, cálculo e demografia. Sem Tkinter, matplotlib ou reportlab. Importável por backends web e scripts. `python -m bsp_core --testes` corre 41 testes.
- **Tiro com Arco** - protocolo completo reescrito. Até 30 ensaios, janela única `confirmação_1 → confirmação_2`, sem Hurdle Step, sem distâncias. Parser tolera ficheiros tab-separated `{id}_{trial} - {DD-MM-YYYY} - Stability export.xls` a 50 Hz.
- **Análise demográfica** - loader dos 142 atletas (referência). Comparações Mann-Whitney e Kruskal-Wallis, correlações Pearson e Spearman, percentis por subgrupo, ligação scores ↔ CoP. Aba dedicada na UI e páginas extra no PDF.
- **Dashboard web** - FastAPI + frontend local. Botão "Dashboard" no desktop arranca uvicorn em porto livre. Endpoints REST para projectos, atletas, comparações, correlações. Storage em `~/.bsp/projects/` com cache hash-invalidated.
- **Auto-update one-click** - banner com "Actualizar agora", verificação via GitHub Releases API, validação SHA256, lançamento de instalador em Windows/macOS.
- **Palavra-passe remota rotativa** - o hash SHA256 da palavra-passe de acesso está num ficheiro público no repo (`.bsp_pass.sha256`). Qualquer versão instalada consulta esse ficheiro no arranque, faz cache local e aceita apenas o hash mais recente. Ao mudar a palavra-passe (simples edit + push), todas as instalações existentes ficam automaticamente sincronizadas no próximo arranque - sem precisar de actualização de binário. Existe fallback offline para a última password validada.
- **Tokens de UI** - constantes `_PAD_*`, `_FONT_*`, dicionário `ICO`, estilos ttk (`Primary.TButton`, `Secondary.TButton`, `Danger.TButton`, `Ghost.TButton`, `Treeview`, `Accent.Horizontal.TProgressbar`).
- **Citações junto ao código** - blocos com Prieto et al. (1996), Carpenter et al. (2001), Winter (1995), Maurer & Peterka (2005), Quijoux et al. (2021) inline onde os cálculos acontecem.

### Alterado

- **ISCPSI só no Tiro** - protocolo PROTO_TIRO mantém badge dourada ISCPSI; arco passa a verde alvo. READMEs e manuais purificados.
- **Logger macOS-safe** - crash logs em `~/Library/Logs/BSP/` (macOS) e `%LOCALAPPDATA%\BSP\` (Windows). Sem tentativas de escrita no Desktop.
- **Fonte cross-platform** - Segoe UI (Win) / SF Pro Text (macOS) / DejaVu Sans (Linux). Mono: Cascadia Mono / Menlo / DejaVu Sans Mono.
- **`BUILD_macOS.sh`** - `set -euo pipefail`, logs de erro em `build_logs/`, smoke test `--testes` após PyInstaller, code-signing ad-hoc, `--osx-bundle-identifier com.aomassuca.bsp`.
- **README** - PT-source-of-truth, tom humano, secção "Problemas conhecidos" honesta. EN/ES/DE traduzidos da versão PT.

### Removido

- **Hurdle Step** do protocolo Tiro com Arco (mantém-se em FMS).
- **`bsp_web.py`** - substituído pelo dashboard FastAPI.
- **Código multi-distância / duas-fases / trial5_1** do arco - legado eliminado.
- **Palavras banidas** - "seamlessly", "comprehensive", "robust", "leverages", etc. Zero ocorrências nos READMEs.

### Correcções

- Nomes de folhas cp1252-corrompidas (`confirma��o_1`) são normalizados no loader.
- Colunas 68–308 do ficheiro dos 142 atletas são ignoradas (dados manuais).
- Chi² 95 % com df=2 = 5.9915 (Schubert & Kirchner 2014) confirmado vs. tabela.
- Crash dialog no macOS já não é invisível (remove `overrideredirect` quando `_SYS == 'Darwin'`).

### Testes

- `estabilidade_gui.py --testes` → 168/168 (inclui [15] auto-update e [16] palavra-passe rotativa)
- `bsp_core.py --testes` → 41/41
- `import bsp_core` não puxa tkinter/matplotlib/reportlab/PIL (verificação explícita).
