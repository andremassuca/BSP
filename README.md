<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

[![Licença: MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-blue.svg)](LICENSE)
[![Testes](https://github.com/andremassuca/BSP/actions/workflows/testes.yml/badge.svg)](https://github.com/andremassuca/BSP/actions/workflows/testes.yml)

[English version below](#english)

## O que é o BSP

O BSP analisa a estabilidade postural a partir de dados de plataforma de forças. Lê as trajectórias do centro de pressão (CoP) exportadas pela plataforma, aplica as janelas temporais de cada ensaio e calcula as métricas habituais em posturografia:

- área da elipse de confiança a 95% (`ea95`);
- RMS e amplitude em X, Y e radial;
- velocidade média do CoP;
- índice de rigidez postural (`stiff_x`, `stiff_y`);
- assimetria direita/esquerda, quando o protocolo a prevê;
- FFT do CoP para inspecção espectral.

Os resultados saem em Excel (resumo e ficheiros por indivíduo) e, opcionalmente, em relatório PDF com elipses e estabilogramas.

Protocolos disponíveis:

| Protocolo | Ensaios | Janela de análise |
|---|---|---|
| FMS Bipodal | 5 direita + 5 esquerda | início/fim por ensaio |
| Apoio Unipodal | 5 direita + 5 esquerda | início/fim por ensaio |
| Tiro | 5 por distância | toque, pontaria, disparo e fim |
| Tiro com Arco | até 30 | entre confirmação 1 e confirmação 2 |

## Requisitos

- Windows 10 ou 11 (também corre em Linux e macOS).
- **Python 3.12 ou mais recente** (testado com 3.12 e 3.14), instalado a partir de [python.org](https://www.python.org/downloads/). O instalador oficial já inclui o `tkinter`, necessário para a interface gráfica.
- Cerca de 500 MB livres para o ambiente virtual.

## Instalação passo a passo (Windows)

1. **Instalar o Python.** Descarregar o instalador em python.org e, no primeiro ecrã, marcar a opção **"Add python.exe to PATH"**.

2. **Obter o BSP.** Na página do repositório, carregar em **Code > Download ZIP** e extrair a pasta, ou, com o Git instalado:

   ```
   git clone https://github.com/andremassuca/BSP.git
   ```

3. **Abrir um terminal na pasta do BSP.** No Explorador de Ficheiros, abrir a pasta `BSP`, escrever `cmd` na barra de endereço e carregar em Enter.

4. **Criar e activar um ambiente virtual:**

   ```
   py -m venv .venv
   .venv\Scripts\activate
   ```

   A linha de comandos passa a começar por `(.venv)`. Se usar o PowerShell e aparecer um erro de permissões ao activar, correr uma vez `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` ou usar a Linha de Comandos (`cmd`).

5. **Instalar as dependências:**

   ```
   pip install -r requirements.txt
   ```

6. **Confirmar que está tudo bem** (opcional, demora cerca de um minuto):

   ```
   cd src
   python bsp_core.py --testes
   python estabilidade_gui.py --testes
   ```

   As duas suites devem terminar com todos os testes a passar.

Nas utilizações seguintes basta abrir o terminal na pasta do BSP e activar o ambiente (`.venv\Scripts\activate`) antes de abrir o programa.

## Abrir a interface gráfica

A partir da pasta `src`, no Windows nativo (não em WSL):

```
cd src
python estabilidade_gui.py
```

Na primeira execução aparecem os termos de utilização. Depois escolhe-se o protocolo: **FMS Bipodal**, **Apoio Unipodal** ou **Tarefa Funcional** (submenu com **Tiro** e **Tiro com Arco**). O botão **Protocolo** na janela principal permite mudar de protocolo mais tarde.

## Como correr cada protocolo

Em todos os protocolos, a **Pasta de indivíduos** tem uma subpasta por indivíduo, com o nome no formato `ID_Nome` (por exemplo `01_Exemplo_A`). O BSP usa o número inicial para associar cada pasta à linha certa do ficheiro de tempos. No fim, carregar em **Executar Análise** e indicar onde guardar o **Resumo Excel** e, se quiser, o relatório PDF.

A pasta `examples/` tem dados sintéticos para os quatro protocolos. Servem para experimentar o programa sem dados reais.

### FMS Bipodal e Apoio Unipodal

- **Pasta de indivíduos:** `examples/fms_bipodal` (ou `examples/unipodal`).
- **Ficheiro Inicio_fim:** `examples/fms_bipodal/inicio_fim.xlsx` (ou `examples/unipodal/inicio_fim.xlsx`).
- Cada subpasta tem `dir_1` a `dir_5` e `esq_1` a `esq_5`.

### Tiro

- **Pasta de indivíduos:** `examples/tiro`.
- **Ficheiro de tempos (tiro):** `examples/tiro/tempos_tiro.xlsx`.
- Escolher em **Intervalos a calcular** as janelas pretendidas (toque a pontaria, pontaria a disparo, etc.).
- Os exemplos não trazem ficheiros de Hurdle Step. O aviso correspondente no registo é normal; pode desmarcar **Incluir análise bipodal (Hurdle Step)**.

### Tiro com Arco

- **Pasta de indivíduos:** `examples/tiro_arco`.
- **Ficheiro de tempos:** `examples/tiro_arco/Inicio_fim_vfinal.xlsx`.
- A referência demográfica é opcional e fica em branco nos exemplos.

## Formato dos dados de entrada

### Ficheiros de ensaio (exportação da plataforma)

Texto separado por tabulações, com extensão `.xls` (é o que o software da plataforma exporta; não é um ficheiro Excel binário). Algumas linhas de cabeçalho e depois uma tabela:

```
Stability export for measurement: dir_1
Patient name: ...
Measurement done on 01-01-2026

Frame	Time (ms)	COF X (mm)	COF Y (mm)	Force (N)	FrameDist (mm)
0	0	0.512	-1.203	700.0	0.000
1	10	0.530	-1.187	700.0	0.024
```

As quatro primeiras colunas (frame, tempo em ms, CoP X e CoP Y em mm) são obrigatórias. No Tiro com Arco as colunas chamam-se `Entire plate COF X`, `Entire plate COF Y` e `Entire plate COF Force` (tipicamente 50 Hz). Também são aceites ficheiros `.xlsx` com colunas `frame`, `t_ms`, `x`, `y`.

Nomes esperados dentro de cada pasta de indivíduo:

| Protocolo | Nome dos ficheiros |
|---|---|
| FMS / Unipodal | `dir_1 ... dir_5`, `esq_1 ... esq_5` (sufixos como ` - 01-01-2026 - Stability export.xls` são aceites) |
| Tiro | `trial{distância}_{ensaio} - ... .xls`, por exemplo `trial5_1 - 01_01_2026 - Stability_export.xls` |
| Tiro com Arco | `{id}_{ensaio} - {DD-MM-AAAA} - Stability export.xls`, por exemplo `901_1 - 01-01-2026 - Stability export.xls` |

### Ficheiros de tempos (Excel)

- **FMS / Unipodal, `inicio_fim.xlsx`:** uma linha por indivíduo. Coluna A com o nome (igual ao da pasta, sem o ID), depois pares início/fim em ms: os ensaios 1 a 5 são o pé direito e os ensaios 6 a 10 o pé esquerdo.
- **Tiro:** três folhas, `tempo (toque)`, `tempo (pontaria)` e `tempo (disparo)`. A linha 2 tem cabeçalhos no formato `5m (1)`, `5m (2)`, ...; a coluna A tem o ID do indivíduo. Na folha do toque, o primeiro bloco de colunas tem os tempos da câmara e o segundo bloco os tempos da placa; o BSP converte os restantes eventos para o relógio da placa.
- **Tiro com Arco, `Inicio_fim_vfinal.xlsx`:** três folhas, `tempo do toque`, `confirmação_1` e `confirmação_2`. Linha 1 com `atleta_id, ensaio_1, ..., ensaio_30`; uma linha por atleta. A janela de análise vai de `confirmação_1` a `confirmação_2`.

Os ficheiros em `examples/` servem de modelo. Para os regenerar: `python examples/gerar_exemplos.py`.

## Documentação adicional

- [MANUAL_Part1_PT.md](MANUAL_Part1_PT.md) e [MANUAL_Part2_PT.md](MANUAL_Part2_PT.md): manual de utilização.
- [CHANGELOG.md](CHANGELOG.md): histórico de versões.
- [CONTRIBUTING.md](CONTRIBUTING.md): como reportar problemas e propor alterações.

## Citar

Se usar o BSP num trabalho, cite o software (ver [CITATION.cff](CITATION.cff); o GitHub mostra o botão **Cite this repository**):

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Versão 1.0) [Software]. https://github.com/andremassuca/BSP

Estudo que usou o BSP: Campião, A., Aleixo, P., Massuça, A. O., Abrantes, J. M. S. C., & Massuça, L. M. (2026). Associations of BIA-Estimated Body Composition, Handgrip Strength, and Event-Defined Postural Control with Short-Range Police Precision-Shooting Accuracy: A Cross-Sectional Study. *Journal of Functional Morphology and Kinesiology*, 11(3), 272. https://doi.org/10.3390/jfmk11030272

Referências metodológicas:

- Schubert, P., & Kirchner, M. (2013). Ellipse area calculations and their applicability in posturography. *Gait & Posture*, 39(1), 518-522.
- Prieto, T. E., et al. (1996). Measures of postural steadiness. *IEEE Transactions on Biomedical Engineering*, 43(9), 956-966.
- Winter, D. A. (1995). Human balance and posture control during standing and walking. *Gait & Posture*, 3(4), 193-214.
- Quijoux, F., et al. (2021). A review of center of pressure (COP) variables to quantify standing balance in elderly people. *Physiological Reports*, 9(22), e15067.
- Maurer, C., & Peterka, R. J. (2005). A new interpretation of spontaneous sway measures based on a simple model of human postural control. *Journal of Neurophysiology*, 93(1), 189-200.

## Licença

O BSP é software livre, distribuído sob a [licença MIT](LICENSE).

## Autores

André O. Massuça, Pedro Aleixo e Luís M. Massuça. Ver [AUTHORS.md](AUTHORS.md).

---

<a id="english"></a>
## English

BSP computes postural stability metrics (95% confidence ellipse area, RMS, amplitude, mean velocity, postural stiffness, left/right asymmetry and CoP spectrum) from force-plate centre-of-pressure exports. It supports four protocols (FMS Bipodal, Single-leg Stance, Pistol Shooting and Archery) and writes Excel summaries and optional PDF reports.

**Requirements:** Python 3.12 or newer (tested with 3.12 and 3.14) from python.org, which includes `tkinter`.

**Install and run (Windows):**

```
git clone https://github.com/andremassuca/BSP.git
cd BSP
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
cd src
python estabilidade_gui.py
```

Run the test suites with `python bsp_core.py --testes` and `python estabilidade_gui.py --testes` from `src/`.

**Input data:** one subfolder per subject named `ID_Name`, each holding the tab-separated `.xls` exports from the force plate, plus an Excel file with the analysis window of each trial. Synthetic examples for every protocol are in `examples/`; the Portuguese section above describes each file format. More documentation in [MANUAL_Part1_EN.md](MANUAL_Part1_EN.md) and [MANUAL_Part2_EN.md](MANUAL_Part2_EN.md).

**Citation:** see [CITATION.cff](CITATION.cff). **Licence:** [MIT](LICENSE). **Contributing:** see [CONTRIBUTING.md](CONTRIBUTING.md).
