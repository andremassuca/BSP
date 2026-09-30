# Changelog

All notable changes to BSP are recorded in this file. [Versão em português abaixo](#histórico-de-alterações).

## [1.0.0] - 2026-09-30

First public release.

- Four protocols: FMS Bipodal, Single-leg Stance, Pistol Shooting and Archery.
- Postural stability metrics from the force-plate centre of pressure: 95% ellipse area (`ea95`), RMS and amplitude in X, Y and radial, mean velocity, postural stiffness (`stiff_x`, `stiff_y`), right/left asymmetry and CoP FFT.
- Results in Excel (group summary, one file per subject and a sheet ready for SPSS) and an optional PDF report with ellipses and stabilograms.
- Processing warnings in the log, in an AVISOS sheet of the Excel results and in the PDF report, for trials whose file is missing and for time windows that extend beyond the force-plate recording.
- Graphical interface (tkinter) in Portuguese, English, Spanish and German.
- Automated test suites: `python bsp_core.py --testes` and `python estabilidade_gui.py --testes`.
- Synthetic example data for the four protocols (`examples/`).
- `requirements.txt` with pinned versions, for Python 3.12 or newer.
- Continuous integration (GitHub Actions) on Ubuntu and Windows.
- MIT licence.

---

<a id="histórico-de-alterações"></a>
# Histórico de alterações

Todas as alterações relevantes ao BSP ficam registadas neste ficheiro.

## [1.0.0] - 2026-09-30

Primeira versão pública.

- Quatro protocolos: FMS Bipodal, Apoio Unipodal, Tiro e Tiro com Arco.
- Métricas de estabilidade postural a partir do centro de pressão da plataforma de forças: área da elipse a 95% (`ea95`), RMS e amplitude em X, Y e radial, velocidade média, rigidez postural (`stiff_x`, `stiff_y`), assimetria direita/esquerda e FFT do CoP.
- Resultados em Excel (resumo do grupo, um ficheiro por indivíduo e uma folha pronta para o SPSS) e relatório PDF opcional com elipses e estabilogramas.
- Avisos de processamento no registo, num separador AVISOS do Excel de resultados e no relatório PDF, para ensaios sem ficheiro e para janelas temporais que ultrapassam o registo da plataforma.
- Interface gráfica (tkinter) em português, inglês, espanhol e alemão.
- Suites de testes automáticos: `python bsp_core.py --testes` e `python estabilidade_gui.py --testes`.
- Dados de exemplo sintéticos para os quatro protocolos (`examples/`).
- `requirements.txt` com versões fixadas, para Python 3.12 ou mais recente.
- Integração contínua (GitHub Actions) em Ubuntu e Windows.
- Licença MIT.
