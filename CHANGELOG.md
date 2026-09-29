# Changelog

All notable changes to BSP are recorded in this file. [Versão em português abaixo](#histórico-de-alterações).

## [1.0.2] - 2026-09-29

### Fixed
- Archery demographic reference: the reader now accepts only rows with an integer athlete ID and at least one anthropometric or score value. The code legend at the end of the reference sheet (rows "ESTILO", "1" and "2") was being loaded as three extra athletes, which inflated the reference count and could attach legend codes to subject folders numbered 1 or 2. New test in `estabilidade_gui.py --testes`.

### Changed
- The interface label "Demographic reference (142 athletes)" is now "Demographic reference" in all four languages. The reference sheet has 42 athletes (IDs 101 to 142); 142 was the highest ID, not the number of athletes. Code comments and docstrings corrected accordingly.

## [1.0.1] - 2026-09-29

### Added
- `README.md` in English as the main page, with a full Portuguese version in `README_PT.md`.
- Step-by-step installation guide for Windows aimed at users with no terminal experience, a quick start, a short macOS and Linux section, a guided first analysis with the example data and a troubleshooting table.
- Screenshots of the main window and of the results (95% ellipse and stabilogram) in `docs/images/`.
- Bilingual bug report template (`.github/ISSUE_TEMPLATE/bug_report.md`).

### Changed
- `CONTRIBUTING.md` and `CHANGELOG.md` now in English first, followed by Portuguese.
- CI: `actions/checkout` and `actions/setup-python` updated to v7 (Node 24); Ubuntu runner pinned to `ubuntu-24.04`.

## [1.0.0] - 2026-09-29

First public release of the source code.

### Included
- Four protocols: FMS Bipodal, Single-leg Stance, Pistol Shooting and Archery.
- Postural stability metrics: 95% ellipse area (`ea95`), RMS and amplitude in X, Y and radial, mean velocity, postural stiffness (`stiff_x`, `stiff_y`), right/left asymmetry and CoP FFT.
- Export to Excel (summary and per-subject files) and PDF report.
- Graphical interface (tkinter) in Portuguese, English, Spanish and German.
- Automated test suites: `python bsp_core.py --testes` and `python estabilidade_gui.py --testes`.
- Synthetic example data for the four protocols (`examples/`).
- `requirements.txt` with pinned versions, tested with Python 3.12 and 3.14.
- Continuous integration (GitHub Actions) on Ubuntu and Windows.
- MIT licence.

### Removed
- Password screen and the associated remote check.
- Authorship check that stopped the program from starting after code changes.
- Automatic update (download and execution of installers from GitHub).
- Launcher for the local web interface, which is not part of this distribution.
- One-time record sent to the authors' server when the licence was accepted (date and time, BSP version, interface language, operating system and a SHA-256 hash of the MAC address, machine name and user name). It was described in the terms of use as anonymous, which was not accurate: the hash can be recalculated by anyone who knows the machine and user names.

---

<a id="histórico-de-alterações"></a>
# Histórico de alterações

Todas as alterações relevantes ao BSP ficam registadas neste ficheiro.

## [1.0.2] - 2026-09-29

### Corrigido
- Referência demográfica do Tiro com Arco: o leitor passa a aceitar só linhas com ID de atleta inteiro e com pelo menos um valor antropométrico ou de pontuação. A legenda de códigos no fim da folha de referência (linhas "ESTILO", "1" e "2") estava a ser carregada como três atletas a mais, o que inflacionava a contagem da referência e podia associar códigos da legenda a pastas de indivíduos numeradas 1 ou 2. Novo teste em `estabilidade_gui.py --testes`.

### Alterado
- A etiqueta da interface "Referência demográfica (142 atletas)" passa a "Referência demográfica" nas quatro línguas. A folha de referência tem 42 atletas (IDs 101 a 142); 142 era o ID mais alto, não o número de atletas. Comentários e docstrings do código corrigidos em conformidade.

## [1.0.1] - 2026-09-29

### Adicionado
- `README.md` em inglês como página principal, com versão portuguesa completa em `README_PT.md`.
- Guia de instalação passo a passo para Windows pensado para quem não tem experiência de terminal, início rápido, secção curta para macOS e Linux, primeira análise guiada com os dados de exemplo e tabela de resolução de problemas.
- Capturas de ecrã da janela principal e dos resultados (elipse a 95% e estabilograma) em `docs/images/`.
- Modelo bilingue para reportar problemas (`.github/ISSUE_TEMPLATE/bug_report.md`).

### Alterado
- `CONTRIBUTING.md` e `CHANGELOG.md` passam a ter o inglês primeiro, seguido do português.
- CI: `actions/checkout` e `actions/setup-python` actualizadas para a v7 (Node 24); runner Ubuntu fixado em `ubuntu-24.04`.

## [1.0.0] - 2026-09-29

Primeira versão pública do código-fonte.

### Incluído
- Quatro protocolos: FMS Bipodal, Apoio Unipodal, Tiro e Tiro com Arco.
- Métricas de estabilidade postural: área da elipse a 95% (`ea95`), RMS e amplitude em X, Y e radial, velocidade média, rigidez postural (`stiff_x`, `stiff_y`), assimetria direita/esquerda e FFT do CoP.
- Exportação para Excel (resumo e ficheiros por indivíduo) e relatório PDF.
- Interface gráfica (tkinter) em português, inglês, espanhol e alemão.
- Suites de testes automáticos: `python bsp_core.py --testes` e `python estabilidade_gui.py --testes`.
- Dados de exemplo sintéticos para os quatro protocolos (`examples/`).
- `requirements.txt` com versões fixadas, testado com Python 3.12 e 3.14.
- Integração contínua (GitHub Actions) em Ubuntu e Windows.
- Licença MIT.

### Removido
- Ecrã de palavra-passe e a verificação remota associada.
- Verificação de autoria que impedia o arranque do programa após alterações ao código.
- Actualização automática (descarga e execução de instaladores a partir do GitHub).
- Lançador da interface web local, que não faz parte desta distribuição.
- Registo enviado uma única vez para o servidor dos autores ao aceitar a licença (data e hora, versão do BSP, língua da interface, sistema operativo e um hash SHA-256 do endereço MAC, do nome da máquina e do nome de utilizador). Os termos de utilização descreviam-no como anónimo, o que não era exacto: o hash pode ser recalculado por quem conheça o nome da máquina e o nome de utilizador.
