# BSP - Biomechanical Stability Program
## Manual de Utilizacao Rapida - v1.0

**Autores:** Andre Massuca (desenvolvimento) | Pedro Aleixo (biomecanica) | Luis Massuca (coordenacao)
**GitHub:** https://github.com/andremassuca/BSP

---

## Indice

1. [Requisitos e Instalacao](#1-requisitos-e-instalacao)
2. [Primeira Execucao](#2-primeira-execucao)
3. [Estrutura de Ficheiros de Entrada](#3-estrutura-de-ficheiros-de-entrada)
4. [Protocolos Disponiveis](#4-protocolos-disponiveis)
5. [Interface Principal](#5-interface-principal)
6. [Opcoes de Saida](#6-opcoes-de-saida)
7. [Analises Estatisticas](#7-analises-estatisticas)
8. [Atalhos de Teclado](#8-atalhos-de-teclado)
9. [Modo CLI (linha de comandos)](#9-modo-cli-linha-de-comandos)
10. [Modo Escuro e Claro](#10-modo-escuro-e-claro)
11. [Termos de Utilizacao](#11-termos-de-utilizacao)
12. [Solucao de Problemas](#12-solucao-de-problemas)
13. [Referencias](#13-referencias)

---

## 1. Requisitos e Instalacao

### Dependencias Python

| Biblioteca | Obrigatoria | Funcao |
|---|---|---|
| Python 3.12+ | Sim | Motor principal (inclui tkinter) |
| numpy | Sim | Calculo numerico |
| scipy | Sim | Estatistica, interpolacao |
| openpyxl | Sim | Leitura/escrita Excel |
| matplotlib | Sim | Graficos |
| reportlab | Sim | Geracao de PDF |
| Pillow | Recomendado | Logos e banners |
| python-docx | Opcional | Relatorio Word |

### Instalar dependencias

A partir da pasta do BSP, com o ambiente virtual activo:

```bash
pip install -r requirements.txt
```

O passo a passo completo para Windows (ambiente virtual incluido) esta no [README_PT](README_PT.md#guia-passo-a-passo-para-windows).

---

## 2. Primeira Execucao

### Tema automatico
Na primeira execucao, o BSP detecta automaticamente o tema do sistema operativo
(macOS dark mode ou Windows 10/11 apps theme) e aplica-o automaticamente.

### Ecra de Termos de Utilizacao
Clica em **"Li e Aceito os Termos"** para continuar.

### Seleccao de Protocolo
- **FMS Bipodal** - analise padrao de estabilidade, 5 ensaios por pe
- **Apoio Unipodal** - apoio unipodal direito e esquerdo
- **Tarefa Funcional** -> submenu -> **Tiro** ou **Tiro com Arco**

---

## 3. Estrutura de Ficheiros de Entrada

```
pasta_principal/
    individuo_01_Nome/
        dir_1.xls
        dir_2.xls
        esq_1.xls
        ...
    individuo_02_Nome/
        ...
```

Cada sub-pasta = um individuo.
Os ficheiros `.xls` sao exportados directamente da plataforma de forca.
O numero de individuos e de ensaios e ilimitado.

### Ficheiro Inicio_fim (FMS / Unipodal)
Excel com colunas de inicio e fim (em segundos) para cada ensaio.

### Ficheiro de Tempos (Protocolo Tiro)
Excel com tempos de inicio e fim por distancia e ensaio.

---

## 4. Protocolos Disponiveis

### FMS Bipodal
- N ensaios por pe (direito + esquerdo), bipodal
- Elipse 95%, assimetria Dir/Esq, todas as metricas de CoP
- N ensaios configuravel (padrao: 5, sem limite maximo)

### Apoio Unipodal
- N ensaios por pe, unipodal
- Elipse 95%, metricas de oscilacao lateral

### Tarefa Funcional - Tiro (ISCPSI)
- Analise por distancia de tiro (detectada automaticamente)
- Multiplos intervalos temporais (toque a pontaria, pontaria a disparo, disparo ao fim, etc.)
- Correlacao metricas de estabilidade vs score de precisao
- Analise bipodal (Hurdle Step) - opcional

### Tarefa Funcional - Tiro com Arco
- Ate 30 ensaios bipodal por atleta
- Janela unica de analise (Confirmacao 1 a Confirmacao 2, definidas por Excel)
- Sem distancia (o ficheiro novo nao codifica distancia)
- Sem Hurdle Step
- Normalizacao pela massa corporal e altura (compara sujeitos de morfologia distinta)

---

## 5. Interface Principal

### Painel esquerdo - Parametros

| Campo | Descricao |
|---|---|
| Pasta de individuos | Pasta raiz com sub-pastas de cada individuo |
| Ficheiro Inicio_fim | Excel com intervalos de tempo |
| Resumo Excel | Ficheiro de saida principal |
| Pasta ficheiros individuais | Pasta para os ficheiros por individuo |
| Relatorio PDF | Caminho do PDF (vazio = nao gerar) |
| N. ensaios | Substituir o numero padrao do protocolo |

### Painel direito - Registo de Execucao
- **Azul** - informacao
- **Verde** - operacao concluida
- **Amarelo** - aviso (dados em falta, outliers, frame rate inconsistente)
- **Vermelho** - erro

### Validacao de frame rate (v23)
O BSP verifica automaticamente a consistencia da taxa de amostragem entre ensaios
do mesmo individuo. Se detectar jitter >20% ou variacao >10% entre ensaios,
emite aviso no log (cor amarela) sem interromper a analise.

### Barra de progresso com ETA
Mostra percentagem de conclusao e tempo estimado restante.

---

## 6. Opcoes de Saida

| Opcao | Descricao |
|---|---|
| Resumo Excel | Ficheiro principal com abas DADOS, GRUPO, SPSS |
| Ficheiros individuais | Um Excel por individuo |
| Relatorio PDF | Capa + novidades + legenda + paginas individuais + citacao |
| Relatorio Word | Tabelas estatisticas (requer python-docx) |
| Relatorio HTML | Graficos interactivos Chart.js, standalone |
| Exportar CSV | Dados em .csv para SPSS/R/Excel |
| Exportar PNG | Estabilogramas e elipses em imagem |

### Exportacao PNG (v23)
A exportacao de PNG aceita configuracao de:
- **DPI:** 72 (web), 150 (apresentacoes), 180 (padrao), 300 (publicacao)
- **Tipos:** estabilograma e/ou elipse 95%

### Relatorio PDF - Pagina de novidades (v23)
O PDF inclui agora uma pagina automatica "O que ha de novo" com as
funcionalidades de cada versao, a seguir a pagina de legenda.

---

## 7. Analises Estatisticas

Activar **"Testes estatisticos automaticos (aba ESTATS)"**.

- **Shapiro-Wilk** + IC 95% por metrica
- **t-pareado ou Wilcoxon** + **Cohen's d** (Dir vs. Esq)
- **CV de variabilidade** intra-individuo
- **Friedman** entre intervalos + **post-hoc Bonferroni** (Tiro)
- **Correlacao de Pearson/Spearman** EA95 vs. score (Tiro)

> Requer n >= 3 individuos.

---

## 8. Atalhos de Teclado

| Windows | macOS | Accao |
|---|---|---|
| Ctrl+Enter | Cmd+Enter | Executar analise |
| Ctrl+S | Cmd+S | Guardar configuracao |
| Ctrl+H | Cmd+H | Ver historico |
| Ctrl+P | Cmd+P | Calculadora de poder amostral |
| F1 | F1 | Ver todos os atalhos |
| F2 | F2 | Analise rapida - ficheiro unico |
| F5 | F5 | Limpar registo de execucao |

---

## 9. Modo CLI (linha de comandos)

```bash
python estabilidade_gui.py --cli PASTA_INDIVIDUOS \
    --inicio_fim caminho/inicio_fim.xlsx \
    --output resultados.xlsx \
    --individuais pasta_individuais \
    --pdf relatorio.pdf \
    --protocolo fms
```

| Argumento | Descricao | Default |
|---|---|---|
| `PASTA` | Pasta raiz de individuos (posicional, obrigatorio) | - |
| `--inicio_fim` | Ficheiro Excel de intervalos de tempo | None |
| `--output` | Ficheiro Excel de saida | `resultados_estabilidade.xlsx` |
| `--individuais` | Pasta para ficheiros individuais | `individuais` |
| `--pdf` | Caminho do PDF | None |
| `--protocolo` | `fms` / `unipodal` / `tiro` | `fms` |

---

## 10. Modo Escuro e Claro

Deteccao automatica na primeira execucao. Alterar manualmente em **botao configuracoes** no canto superior direito.

---

## 11. Termos de Utilizacao

Termos aceites na primeira execucao - guardados em `~/.aom_estabilidade.json`.

Para repor:
```bash
# Windows:
del "%USERPROFILE%\.aom_estabilidade.json"
# macOS / Linux:
rm ~/.aom_estabilidade.json
```

---

## 12. Solucao de Problemas

| Problema | Solucao |
|---|---|
| `ModuleNotFoundError` ao arrancar | Activar o ambiente virtual e correr `pip install -r requirements.txt` |
| Validacao falha antes de iniciar | Verificar pasta de individuos com sub-pastas |
| PDF nao gerado | Confirmar campo PDF nao vazio e opcao activa |
| Testes estatisticos em cinzento | Activar opcao e garantir n >= 3 individuos |
| Relatorio HTML nao abre | Verificar ligacao a internet (Chart.js CDN) |
| Aviso "frame rate inconsistente" | Verificar exportacao da plataforma de forca - todos os ensaios devem ter o mesmo rate |
| Erros com > 10 ensaios | Sem limite - colunas adaptam-se automaticamente |
| Programa fecha sem erro | Actualizar para v23 - corrige crash silencioso da thread de analise |

---

## 13. Referencias

- Schubert, P., & Kirchner, M. (2013). Ellipse area calculations and their applicability in posturography. *Gait & Posture*, 39(1), 518-522.
- Winter, D.A. (1995). Human balance and posture control during standing and walking. *Gait & Posture*, 3(4), 193-214.
- Prieto, T.E. et al. (1996). Measures of postural steadiness: differences between healthy young and elderly adults. *IEEE Trans. Biomed. Eng.*, 43(9), 956-966.
- Quijoux, F. et al. (2021). A review of centre of pressure (COP) variables to quantify standing balance in elderly people: Algorithms and open-access code. *Physiological Reports*, 9(22), e15067.
- Paillard, T., & Noe, F. (2015). Techniques and criteria for assessing balance and postural control in sport. *Frontiers in Physiology*, 6, 399.
- Cohen, J. (1988). *Statistical Power Analysis for the Behavioral Sciences* (2nd ed.). Erlbaum.

---

## Citacao Academica

```
Massuca, A. O., Aleixo, P., & Massuca, L. M. (2026). BSP: Biomechanical Stability Program
(Versao 1.0) [Software]. https://github.com/andremassuca/BSP
```

---

*BSP v1.0 - Andre O. Massuca, Pedro Aleixo & Luis M. Massuca*
