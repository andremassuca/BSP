<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

[![Licença: MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-blue.svg)](LICENSE)
[![Testes](https://github.com/andremassuca/BSP/actions/workflows/testes.yml/badge.svg)](https://github.com/andremassuca/BSP/actions/workflows/testes.yml)

**English version: [README.md](README.md)**

![Janela principal do BSP depois de analisar o exemplo sintético de FMS Bipodal](docs/images/bsp_janela_principal.png)

## Índice

- [O que é o BSP](#o-que-é-o-bsp)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
  - [Início rápido](#início-rápido)
  - [Guia passo a passo para Windows](#guia-passo-a-passo-para-windows)
  - [macOS e Linux](#macos-e-linux)
- [Primeira análise com os dados de exemplo](#primeira-análise-com-os-dados-de-exemplo)
- [Como correr cada protocolo](#como-correr-cada-protocolo)
- [Formato dos dados de entrada](#formato-dos-dados-de-entrada)
- [Resolução de problemas](#resolução-de-problemas)
- [Documentação adicional](#documentação-adicional)
- [Citar o BSP](#citar-o-bsp)
- [Licença](#licença)
- [Autores](#autores)

## O que é o BSP

O BSP analisa a estabilidade postural a partir de dados de plataforma de forças. Lê as trajectórias do centro de pressão (CoP) exportadas pela plataforma, aplica a janela temporal de cada ensaio e calcula as métricas habituais em posturografia:

- área da elipse de confiança a 95% (`ea95`);
- RMS e amplitude em X, Y e radial;
- velocidade média do CoP;
- índice de rigidez postural (`stiff_x`, `stiff_y`);
- assimetria direita/esquerda, quando o protocolo a prevê;
- FFT do CoP para inspecção espectral.

Os resultados saem em Excel (resumo do grupo e um ficheiro por indivíduo) e, opcionalmente, num relatório PDF com elipses e estabilogramas.

![Elipses de confiança a 95% de cinco ensaios e estabilograma de um ensaio (exemplo sintético)](docs/images/bsp_resultados.png)

Protocolos disponíveis:

| Protocolo | Ensaios | Janela de análise |
|---|---|---|
| FMS Bipodal | 5 direita + 5 esquerda | início/fim por ensaio |
| Apoio Unipodal | 5 direita + 5 esquerda | início/fim por ensaio |
| Tiro | 5 por distância | toque, pontaria, disparo e fim |
| Tiro com Arco | até 30 | entre confirmação 1 e confirmação 2 |

## Requisitos

- Windows 10 ou 11 (o BSP também corre em macOS e Linux).
- **Python 3.12 ou mais recente** (testado com 3.12 e 3.14). Versões anteriores não funcionam porque as bibliotecas científicas fixadas exigem a 3.12.
- Cerca de 500 MB livres em disco para o ambiente virtual.
- Ligação à internet durante a instalação.

O BSP corre inteiramente no teu computador e não envia dados para nenhum servidor. As únicas excepções são o relatório HTML opcional, que carrega a biblioteca Chart.js a partir de cdn.jsdelivr.net quando o abres no navegador, e os relatórios de erro, que só são enviados se o decidires, através do teu programa de email.

## Instalação

### Início rápido

Se já usa Python e o terminal, bastam estes comandos (Windows):

```
git clone https://github.com/andremassuca/BSP.git
cd BSP
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
cd src
python estabilidade_gui.py
```

Se alguma coisa aqui não lhe for familiar, siga o guia abaixo.

### Guia passo a passo para Windows

#### 1. Instalar o Python

1. Abra [python.org/downloads](https://www.python.org/downloads/) e descarregue o Python mais recente (3.12 ou superior) para Windows.
2. Execute o instalador. **No primeiro ecrã, marque "Add python.exe to PATH"** antes de carregar em *Install Now*. Isto permite que o Windows encontre o Python quando escreve comandos no terminal, a partir de qualquer pasta.
3. Mantenha as opções por omissão: incluem o `tcl/tk and IDLE`, de que o BSP precisa para mostrar a janela.

#### 2. Obter o BSP

Escolha uma das duas opções:

- **Opção A, sem Git (a mais simples):** na [página do repositório](https://github.com/andremassuca/BSP), carregue no botão verde **Code** e depois em **Download ZIP**. Abra o ficheiro descarregado, carregue em **Extrair tudo** e escolha uma pasta, por exemplo `Documentos`. Fica com uma pasta chamada `BSP-main`.
- **Opção B, com Git:** se tiver o [Git](https://git-scm.com/) instalado, abra um terminal na pasta onde quer o BSP e escreva:

  ```
  git clone https://github.com/andremassuca/BSP.git
  ```

  Fica com uma pasta chamada `BSP`.

Daqui em diante, "a pasta do BSP" é a `BSP-main` (opção A) ou a `BSP` (opção B).

#### 3. Abrir um terminal dentro da pasta do BSP

Abra a pasta do BSP no Explorador de Ficheiros (deve ver o `README.md`, o `requirements.txt` e as pastas `src` e `examples`) e use uma destas formas:

- carregue na barra de endereço no topo da janela, escreva `cmd` e carregue em Enter; ou
- carregue com o botão direito numa zona vazia da pasta e escolha **Abrir no Terminal**.

Abre-se uma janela preta ou azul. O texto antes do cursor (a *linha de comandos*) mostra a pasta onde está, e tem de terminar na pasta do BSP, por exemplo `C:\Users\aluno\Documents\BSP-main>`.

#### 4. Criar e activar o ambiente virtual

Um ambiente virtual (*venv*) é uma cópia do Python só para o BSP, para que as bibliotecas que ele usa não interfiram com o resto do computador. Escreva:

```
py -m venv .venv
```

Demora alguns segundos e cria uma pasta `.venv`. Depois active-o:

```
.venv\Scripts\activate
```

**Como saber que ficou activo:** a linha de comandos passa a começar por `(.venv)`, por exemplo `(.venv) C:\Users\aluno\Documents\BSP-main>`. Se aparecer um erro a dizer que a execução de scripts está desactivada, veja a [Resolução de problemas](#resolução-de-problemas).

#### 5. Instalar as bibliotecas que o BSP usa

Com o ambiente activo:

```
pip install -r requirements.txt
```

Descarrega o NumPy, o SciPy, o Matplotlib e as restantes bibliotecas (alguns minutos da primeira vez). Deve terminar com uma linha que começa por `Successfully installed`.

#### 6. Confirmar a instalação

Ainda com o ambiente activo, corra as duas suites de testes automáticos:

```
cd src
python bsp_core.py --testes
python estabilidade_gui.py --testes
```

Cada uma demora até um minuto e termina com uma linha do tipo `RESULTADO ... X/X testes passaram`. A instalação está correcta quando, nas duas suites, os dois números dessa linha são iguais, por exemplo `20/20`. Se o primeiro número for menor, houve testes que falharam: veja a [Resolução de problemas](#resolução-de-problemas).

#### 7. Abrir o BSP

Já está na pasta `src`, por isso:

```
python estabilidade_gui.py
```

Abre-se a janela do BSP. Da primeira vez mostra a licença e os termos de utilização: escolha a língua (PT, EN, ES ou DE) no topo e carregue em **Li e Aceito os Termos**. Depois escolha o protocolo. A língua pode ser mudada mais tarde no botão ⚙ da janela principal.

#### 8. Voltar a abrir o BSP noutro dia

Não é preciso instalar nada outra vez. Abra um terminal na pasta do BSP (passo 3) e escreva:

```
.venv\Scripts\activate
cd src
python estabilidade_gui.py
```

### macOS e Linux

Os passos são os mesmos, com três diferenças: o comando do Python é `python3`, o ambiente activa-se com `source` e, em Linux, o `tkinter` pode ter de ser instalado à parte.

```
git clone https://github.com/andremassuca/BSP.git
cd BSP
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cd src
python estabilidade_gui.py
```

- **Linux:** se aparecer `No module named 'tkinter'`, instale-o com o gestor de pacotes do sistema, por exemplo `sudo apt install python3-tk` em Ubuntu/Debian ou `sudo dnf install python3-tkinter` em Fedora. Em Ubuntu/Debian, se `python3 -m venv .venv` disser `ensurepip is not available`, instale primeiro o `python3-venv` (`sudo apt install python3-venv`).
- **macOS:** use o instalador de python.org, que já inclui o `tkinter`.

## Primeira análise com os dados de exemplo

A pasta `examples/` tem dados **sintéticos** (gerados por um programa, não são de pessoas reais) para os quatro protocolos. Este exemplo usa o FMS Bipodal.

1. Abra o BSP (passo 7 acima) e, no ecrã de protocolos, carregue em **FMS Bipodal**.
2. Em **FICHEIROS DE ENTRADA**, no campo **Indivíduos**, carregue no botão **...** e escolha a pasta `examples\fms_bipodal` dentro da pasta do BSP. Tem uma subpasta por indivíduo (`01_Exemplo_A`, `02_Exemplo_B`, `03_Exemplo_C`).
3. No campo **Inicio_fim**, carregue em **...** e escolha `examples\fms_bipodal\inicio_fim.xlsx`. Este ficheiro tem a janela de análise de cada ensaio.
4. Em **FICHEIROS DE SAÍDA**, no campo **Resumo Excel**, carregue em **...**, escolha a pasta onde quer os resultados (por exemplo a própria pasta do BSP) e escreva um nome como `resultados_fms.xlsx`.
5. Opcional: em **Individuais** escolha uma pasta para os ficheiros de cada indivíduo. Deixe o campo **PDF** vazio se não quiser o relatório PDF.
6. Carregue em **Executar Análise**. O registo à direita mostra cada ensaio à medida que é processado e, no fim, uma janela confirma "Análise concluída! 3 indivíduo(s)." Carregue em **OK**.
7. Carregue em **Abrir Pasta de Resultados** e abra o `resultados_fms.xlsx` no Excel. A folha **DADOS** tem uma linha por indivíduo com todas as métricas, a **GRUPO** tem as estatísticas do grupo e a **SPSS** tem os dados prontos para o SPSS.

A imagem no topo desta página mostra a janela no fim desta análise.

## Como correr cada protocolo

Em todos os protocolos, a pasta de **Indivíduos** tem uma subpasta por indivíduo, com o nome no formato `ID_Nome` (por exemplo `01_Exemplo_A`). O BSP usa o número inicial para associar cada pasta à linha certa do ficheiro de tempos.

### FMS Bipodal e Apoio Unipodal

- **Indivíduos:** `examples/fms_bipodal` (ou `examples/unipodal`).
- **Inicio_fim:** `examples/fms_bipodal/inicio_fim.xlsx` (ou `examples/unipodal/inicio_fim.xlsx`).
- Cada subpasta tem `dir_1` a `dir_5` (pé direito) e `esq_1` a `esq_5` (pé esquerdo).

### Tiro

- No ecrã de protocolos escolha **Tarefa Funcional** e depois **Tiro**.
- **Indivíduos:** `examples/tiro`.
- **Tempos (tiro):** `examples/tiro/tempos_tiro.xlsx`.
- Em **Intervalos a calcular**, escolha as janelas pretendidas (toque a pontaria, pontaria a disparo, etc.).
- Os exemplos não trazem ficheiros de Hurdle Step. O aviso correspondente no registo é normal; pode desmarcar **Incluir análise bipodal (Hurdle Step)**.

### Tiro com Arco

- No ecrã de protocolos escolha **Tarefa Funcional** e depois **Tiro com Arco**.
- **Indivíduos:** `examples/tiro_arco`.
- **Inicio_fim:** `examples/tiro_arco/Inicio_fim_vfinal.xlsx`.
- A **Ref. demográfica** é opcional e fica em branco nos exemplos.

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

## Resolução de problemas

| Problema | Solução |
|---|---|
| `'python' não é reconhecido como um comando interno ou externo` (ou `'py' não é reconhecido`) | O Python não está instalado ou foi instalado sem "Add python.exe to PATH". Volte a correr o instalador de python.org, escolha **Modify** (ou reinstale) e marque a opção do PATH. Feche e volte a abrir o terminal. |
| Ao activar o ambiente no PowerShell: `a execução de scripts está desativada neste sistema` | Corra uma vez `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, responda `S` (ou `Y`) e active outra vez. Em alternativa use a Linha de Comandos: escreva `cmd` na barra de endereço da pasta (passo 3). |
| `ModuleNotFoundError: No module named 'tkinter'` | Em Windows, volte a correr o instalador de python.org, escolha **Modify** e marque `tcl/tk and IDLE`. Em Linux, instale o `python3-tk` (ver [macOS e Linux](#macos-e-linux)). |
| `ModuleNotFoundError: No module named 'numpy'` (ou outra biblioteca) | O ambiente não está activo ou a instalação não terminou. Confirme que a linha de comandos começa por `(.venv)` e corra outra vez `pip install -r requirements.txt` a partir da pasta do BSP. |
| `Could not find a version that satisfies the requirement numpy==...` | O Python é demasiado antigo. Confirme com `py --version`: tem de ser 3.12 ou superior. Instale um Python recente, apague a pasta `.venv` e repita a partir do passo 4. |
| `pip install` falha por outro motivo (erros de rede, tempo esgotado) | Confirme a ligação à internet (algumas redes institucionais bloqueiam o pip; experimente outra rede). Actualize o pip com `python -m pip install --upgrade pip` e tente de novo. |
| `No such file or directory: 'requirements.txt'` | O terminal não está na pasta do BSP. Veja a linha de comandos e repita o passo 3. |
| Um teste falha | Abra uma issue (ver [CONTRIBUTING.md](CONTRIBUTING.md)) com o output completo do comando de testes. |

## Documentação adicional

- [MANUAL_Part1_PT.md](MANUAL_Part1_PT.md) e [MANUAL_Part2_PT.md](MANUAL_Part2_PT.md): manual de utilização (português).
- [MANUAL_Part1_EN.md](MANUAL_Part1_EN.md) e [MANUAL_Part2_EN.md](MANUAL_Part2_EN.md): manual de utilização (inglês).
- [CHANGELOG.md](CHANGELOG.md): histórico de versões.
- [CONTRIBUTING.md](CONTRIBUTING.md): como reportar problemas e propor alterações.

## Citar o BSP

Se usar o BSP num trabalho, cite o software (ver [CITATION.cff](CITATION.cff); o GitHub mostra o botão **Cite this repository**):

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Versão 1.0.0) [Software]. https://github.com/andremassuca/BSP

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
