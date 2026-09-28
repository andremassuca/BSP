# Changelog

Todas as alterações relevantes ao BSP ficam registadas neste ficheiro.
All notable changes to BSP are recorded in this file.

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
- Recolha de estatísticas de utilização e identificador da máquina.
