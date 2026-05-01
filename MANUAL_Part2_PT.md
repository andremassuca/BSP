# BSP - Biomechanical Stability Program
## Manual Completo - v1.0

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
9. [Modo CLI](#9-modo-cli)
10. [Interface Web](#10-interface-web)
11. [Tema e Acessibilidade](#11-tema-e-acessibilidade)
12. [Auto-update](#12-auto-update)
13. [Perfis de Configuracao](#13-perfis-de-configuracao)
14. [Analise Rapida - Ficheiro Unico](#14-analise-rapida---ficheiro-unico)
15. [Relatorio HTML Interactivo](#15-relatorio-html-interactivo)
16. [Validacao de Dados](#16-validacao-de-dados)
17. [Exportacao PNG](#17-exportacao-png)
18. [Relatorio PDF - Estrutura Completa](#18-relatorio-pdf---estrutura-completa)
19. [Termos de Utilizacao e Palavra-passe](#19-termos-de-utilizacao-e-palavra-passe)
20. [Compilar para Distribuicao](#20-compilar-para-distribuicao)
21. [Solucao de Problemas](#21-solucao-de-problemas)
22. [Historico de Versoes](#22-historico-de-versoes)
23. [Referencias](#23-referencias)

---

## 1. Requisitos e Instalacao

### Compatibilidade

| Sistema | Versao minima | Testado |
|---|---|---|
| Windows | 10 / 11 | 10 22H2, 11 23H2 |
| macOS | 12 Monterey | Apple Silicon M1/M2/M3 e Intel |
| Python | 3.9 | 3.9 / 3.10 / 3.11 / 3.12 |

### Dependencias Python

| Biblioteca | Obrigatoria | Funcao |
|---|---|---|
| numpy | Sim | Calculo numerico |
| scipy | Sim | Estatistica, filtros |
| openpyxl | Sim | Leitura/escrita Excel |
| matplotlib | Sim | Graficos |
| reportlab | Sim | Geracao de PDF |
| Pillow | Recomendado | Imagens e logos |
| python-docx | Opcional | Relatorio Word |
| streamlit | Opcional | Interface web |

```bash
pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow
```

### Windows - Instalador standalone

1. Instala Python 3.9+ (https://www.python.org/downloads/) - activar "Add Python to PATH"
2. Duplo-clique em `BUILD_Windows.bat` -> gera `BSP_Setup.exe`
3. Distribui o `BSP_Setup.exe` - instala em qualquer maquina Windows sem Python

O instalador regista o programa em Adicionar/Remover Programas e cria atalho no Desktop.

### macOS - DMG

```bash
chmod +x BUILD_macOS.sh && ./BUILD_macOS.sh
```

Gera `BSP.dmg` com `BSP.app`. Primeira abertura: botao direito -> Abrir.
Desinstalar: Launchpad -> clique longo -> arrastar para o Lixo.

### DPI e resolucao de ecra

Em Windows com ecrãs de alta resolucao (4K, HiDPI), o BSP activa automaticamente
DPI awareness via `SetProcessDpiAwareness`. Em macOS, o Tkinter adapta-se ao
Retina display sem configuracao adicional.

---

## 2. Primeira Execucao

1. **Tema automatico** - detecta dark/light mode do SO, aplica e guarda preferencia
2. **Termos de Utilizacao** - aceitar uma vez; guardado em `~/.aom_estabilidade.json`
3. **Palavra-passe** - disponivel em https://github.com/andremassuca
4. **Seleccao de protocolo** - FMS / Unipodal / Tarefa Funcional (Tiro / Tiro com Arco)

A janela redimensiona automaticamente para o ecra disponivel.
Em monitores pequenos (< 1280px), as zonas comprimem proporcionalmente.

---

## 3. Estrutura de Ficheiros de Entrada

### FMS Bipodal e Apoio Unipodal

```
pasta_principal/
    001_Nome_Apelido/
        dir_1.xls   (pe direito, ensaio 1)
        dir_2.xls
        esq_1.xls   (pe esquerdo, ensaio 1)
        esq_2.xls
    002_Outro_Nome/
        ...
```

O numero de individuos e de ensaios e ilimitado.
As tabelas PDF adaptam-se automaticamente: para > 10 ensaios,
as colunas comprimem e a fonte reduz para manter tudo legivel numa pagina A4.

### Protocolo de Tiro

```
pasta_principal/
    001_Atirador/
        trial5_1_-_13_01_2026_-_Stability_export.xls
        trial5_2_-_...
        trial10_1_-_...
        hs_dir_1.xls   (Hurdle Step, opcional)
    ...
```

O prefixo `trial[dist]_[ensaio]` e detectado automaticamente.
Suporta multiple distancias na mesma pasta.

### Ficheiro Inicio_fim

Excel com colunas:
- `nome` - nome do individuo (deve corresponder ao nome da pasta)
- `inicio_X` / `fim_X` - tempos em segundos para cada ensaio X

### Formato XLS da plataforma de forca

O BSP le ficheiros exportados directamente da plataforma Kistler/AMTI no formato
tabular padrao (Frame, Time, COF X, COF Y, ...). Suporta colunas adicionais
de seleccao (Selection CoP direito/esquerdo) a partir da coluna 6.

---

## 4. Protocolos Disponiveis

### FMS Bipodal
- Dois lados: Pe Direito e Pe Esquerdo
- N ensaios configuravel (padrao 5, sem limite)
- Elipse 95%, assimetria Dir/Esq (ratio de simetria standard)
- Testes estatisticos: Shapiro-Wilk, t-pareado/Wilcoxon, Cohen's d

### Apoio Unipodal
- Dois lados: Pe Direito e Pe Esquerdo
- N ensaios configuravel
- Metricas de oscilacao lateral (RMS ML, RMS AP, amplitudes)
- Sem indice de assimetria (cada pe analisado independentemente)

### Tarefa Funcional - Tiro
- Dados bipodais (Entire Plate CoP) + seleccao de pe direito/esquerdo
- Detecta distancias automaticamente a partir dos nomes dos ficheiros
- Multiplos intervalos: toque de pontaria, toque de disparo,
  pontaria-disparo, pos-disparo, total
- Hurdle Step bipodal opcional (ficheiros `hs_dir_*.xls`, `hs_esq_*.xls`)
- Relatorio de grupo com graficos de barras (ea95 por distancia e intervalo)
- Correlacao Pearson/Spearman ea95 vs score de precisao

### Tarefa Funcional - Tiro com Arco
- Identico ao protocolo de Tiro, adaptado para arco

---

## 5. Interface Principal

### Painel esquerdo - Configuracao

| Campo | Descricao |
|---|---|
| Pasta de individuos | Pasta raiz com sub-pastas; aceita qualquer numero de individuos |
| Ficheiro Inicio_fim | Excel de intervalos; opcional se os ficheiros tiverem tempo embutido |
| Resumo Excel | Ficheiro de saida (.xlsx) |
| Pasta individuais | Pasta para ficheiros individuais e PNG |
| Relatorio PDF | Caminho do relatorio PDF; deixar vazio para nao gerar |
| Relatorio Word | Caminho opcional do ficheiro .docx |
| N. ensaios | Override do padrao do protocolo; deixar vazio para usar o padrao |

### Validacao previa
Ao clicar Executar, o BSP verifica:
- Pasta de individuos existe e tem sub-pastas
- Ficheiro de saida Excel definido
- Ficheiro de tempos existe (se indicado)
Problemas sao listados numa janela de aviso antes de iniciar.

### Registo de execucao
Mostra progresso em tempo real. Cores:
- Azul - informacao
- Verde - ok
- Amarelo - aviso (dados em falta, outliers, jitter de frame rate)
- Vermelho - erro

### Barra de progresso
Mostra percentagem e ETA (tempo estimado restante) a partir dos 2%.

### Botoes utilitarios

| Botao | Funcao |
|---|---|
| Protocolo | Mudar protocolo sem reiniciar |
| Historico | Ver sessoes anteriores com pesquisa |
| Poder | Calculadora de tamanho amostral |
| Guardar | Guardar configuracao atual |
| Perfis | Gerir perfis nomeados |
| Rapido [F2] | Analise rapida de ficheiro unico |
| Atalhos [F1] | Ver todos os atalhos |

---

## 6. Opcoes de Saida

### Resumo Excel (obrigatorio)
Ficheiro `.xlsx` com:
- **Aba DADOS** - uma linha por individuo, todas as metricas por lado e ensaio
- **Aba GRUPO** - estatisticas do grupo (Media, DP, CV%)
- **Aba SPSS** - formato de exportacao para SPSS (uma coluna por metrica x ensaio)
- **Aba ESTATS** - testes estatisticos (se activado)

### Ficheiros Individuais
Um `.xlsx` por individuo com:
- **Aba RESUMO** - tabela de metricas por ensaio com outliers assinalados
- **Aba SPSS** - dados no formato SPSS
- **Abas T1..TN** - dados frame-a-frame com grafico de elipse embutido
- **Aba ELIPSE** - grafico de elipse 95% com todos os ensaios sobrepostos
- **Aba ESTABILOGRAMA** - CoP X e Y ao longo do tempo

### Relatorio PDF
Ver seccao 18 para estrutura completa.

### Relatorio Word (.docx)
Requer `python-docx`. Contem:
- Cabecalho com nome, data e protocolo
- Tabela de metricas por individuo e lado
- Citacao academica no rodape

### Relatorio HTML
Ver seccao 15.

### Exportar CSV
Ficheiros `.csv` prontos para R / SPSS / Excel:
- `*_grupo.csv` - resumo do grupo
- `*_individual.csv` - dados por individuo e ensaio

### Exportar PNG
Ver seccao 17.

---

## 7. Analises Estatisticas

Activar em **Opcoes -> Testes estatisticos automaticos (aba ESTATS)**.

### Shapiro-Wilk
Teste de normalidade para cada metrica e cada individuo.
Apresentado com p-value e IC 95%.

### Comparacao Dir vs. Esq
- t-pareado (se normalidade) ou Wilcoxon (se nao-normal)
- Cohen's d como medida de efeito
- Disponivel para FMS e Unipodal

### Variabilidade intra-individuo
- CV% (coeficiente de variacao) com semaforo:
  - Verde: CV < 15%
  - Amarelo: 15-30%
  - Vermelho: > 30%

### Protocolo de Tiro (adicional)
- **Indice de perturbacao** por ensaio
- **Friedman** entre intervalos + **post-hoc Bonferroni**
- **Correlacao Pearson/Spearman** ea95 vs. score de precisao

> Requer n >= 3 individuos. Para n < 10, aviso de baixo poder estatistico.

---

## 8. Atalhos de Teclado

| Windows | macOS | Accao |
|---|---|---|
| Ctrl+Enter | Cmd+Enter | Executar analise |
| Ctrl+S | Cmd+S | Guardar configuracao |
| Ctrl+H | Cmd+H | Ver historico |
| Ctrl+P | Cmd+P | Calculadora de poder |
| F1 | F1 | Ver todos os atalhos |
| F2 | F2 | Analise rapida - ficheiro unico |
| F5 | F5 | Limpar registo |
| Escape | Escape | Cancelar analise em curso |

---

## 9. Modo CLI

Util para automacao, scripting e integracao em pipelines.

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
| `PASTA` | Pasta raiz de individuos (obrigatorio) | - |
| `--inicio_fim` | Ficheiro de intervalos de tempo | None |
| `--output` | Ficheiro Excel de saida | `resultados_estabilidade.xlsx` |
| `--individuais` | Pasta para ficheiros individuais | `individuais` |
| `--pdf` | Caminho do PDF | None |
| `--protocolo` | `fms` / `unipodal` / `tiro` | `fms` |
| `--scores` | Ficheiro de scores (tiro) | None |
| `--sem_ind` | Nao gerar ficheiros individuais | False |
| `--sem_elipse` | Nao gerar grafico de elipse | False |
| `--sem_estab` | Nao gerar estabilograma | False |

---

## 10. Interface Web

```bash
pip install streamlit
streamlit run bsp_web.py
```

Abre automaticamente em **http://localhost:8501**

1. Selecciona protocolo no menu lateral
2. Upload da pasta de individuos em ficheiro .zip
3. Upload do Excel de inicio/fim (opcional)
4. Clica **Executar Analise**
5. Descarrega o ZIP de resultados quando concluido

Util para acesso remoto, demonstracoes e ambientes sem GUI.

---

## 11. Tema e Acessibilidade

**Deteccao automatica:** Windows (registo `AppsUseLightTheme`), macOS (`AppleInterfaceStyle`).

**Alterar manualmente:** botao de configuracoes no canto superior direito.

**Idiomas:** PT / EN / ES / DE. Todos os outputs (PDF, Excel, HTML) adaptam
os labels das metricas ao idioma seleccionado.
- Metricas PDF: traducao dinamica via `mets_pdf_localizadas()`
- Lados: "Pe Direito" / "Right Foot" / "Pie Derecho" / "Rechter Fuss"
- EULA: texto completo em 4 linguas

**Ecras de alta resolucao:**
- Windows: DPI awareness automatico (4K, HiDPI)
- macOS: Retina display suportado sem configuracao
- A interface redimensiona para qualquer tamanho de ecra

---

## 12. Auto-update

O BSP verifica `https://api.github.com/repos/andremassuca/BSP/releases/latest`
ao iniciar. Se houver nova versao, aparece banner com "Actualizar agora"
(download + verificacao SHA256 + lanca instalador), "Ver notas" e "Dispensar".

Para publicar actualizacao (para maintainers):
1. Actualizar `VERSAO = "1.0.1"` em `estabilidade_gui.py`
2. Actualizar ficheiro `VERSION` na raiz
3. Seguir `docs/RELEASE.md` (build em cada OS + `SHA256SUMS.txt`)
4. `gh release create v1.0.1 ...`

---

## 13. Perfis de Configuracao

Botao **Perfis** -> guardar e carregar configuracoes completas com nome proprio.

| Operacao | Descricao |
|---|---|
| Guardar actual | Guarda pasta, ficheiro de tempos, saida, PDF, protocolo |
| Carregar | Preenche todos os campos com os valores do perfil |
| Apagar | Remove o perfil seleccionado |

Guardados em `~/.aom_estabilidade_profiles.json`.

Exemplos tipicos:
- `FMS_Turma2026` - configuracao para a turma de 2026, protocolo FMS
- `Tiro_ISCPSI_Mar` - protocolo de tiro, dados de Marco
- `Unipodal_Desporto` - analise unipodal para grupo de atletas

---

## 14. Analise Rapida - Ficheiro Unico

**Acesso:** botao **Rapido** ou tecla **F2**

Analisa um unico ficheiro `.xls` sem configurar pasta, ficheiro de tempos ou saida.

1. Prima F2
2. Selecciona o ficheiro .xls
3. O BSP calcula todas as metricas e apresenta os resultados
4. Clica **Exportar Excel** para guardar

Metricas calculadas: amplitudes ML/AP, velocidades medias e de pico, deslocamento,
area da elipse 95%, semi-eixos, RMS ML/AP/Radius, stiffness, variancia, covariancia.

Util para uso clinico rapido, verificacao de dados e demonstracoes.

---

## 15. Relatorio HTML Interactivo

Gera ficheiro `.html` standalone (sem servidor, sem instalacao).

Conteudo:
- Graficos de barras interactivos por metrica (hover, zoom)
- Tabela de resumo do grupo (Media, DP, CV%)
- Bloco de citacao academica
- Suporte dark/light mode automatico

Partilha por email ou servidor sem qualquer instalacao do lado do receptor.

---

## 16. Validacao de Dados (v23)

O BSP valida automaticamente a qualidade dos dados de cada individuo.

### Validacao de frame rate
- Detecta timestamps duplicados ou invertidos em qualquer ensaio
- Detecta jitter de taxa de amostragem > 20% dentro de um ensaio
- Detecta variacao > 10% na frequencia entre ensaios do mesmo lado
- Avisos aparecem no log (cor amarela) sem interromper a analise

### O que fazer com avisos de frame rate
- Verificar exportacao da plataforma de forca (rate identico em todos os ensaios)
- Confirmar que os ficheiros nao estao corrompidos ou truncados
- Se o jitter for consistente em todos os individuos, pode ser caracteristica
  da plataforma e nao afecta os calculos (que usam timestamps reais)

### Outliers
Ensaios identificados como outliers (> 2 DP da media do grupo para ea95)
sao assinalados em vermelho nas tabelas Excel e no relatorio PDF.

---

## 17. Exportacao PNG (v23)

Gera imagens dos graficos individuais (estabilograma e/ou elipse 95%).

### Opcoes disponiveis

| Parametro | Opcoes | Default |
|---|---|---|
| DPI | 72, 96, 150, 180, 300 | 180 |
| Tipos | estabilograma, elipse | ambos |

**DPI recomendado por destino:**
- 72 dpi - ecra / web
- 150 dpi - apresentacoes PowerPoint
- 180 dpi - relatorios digitais
- 300 dpi - publicacao cientifica / impressao

Os PNG sao guardados em `pasta_individuais/png/`.

---

## 18. Relatorio PDF - Estrutura Completa

O relatorio PDF e gerado com a biblioteca ReportLab e tem estrutura fixa:

1. **Capa** - titulo, protocolo, data, numero de individuos, autores
2. **Capa 2** (se > muitos individuos) - continuacao da lista
3. **Legenda** - descricao de todas as metricas com formulas
4. **O que ha de novo (v23)** - novidades da versao actual e da anterior
5. **Indice** (apenas protocolo Tiro) - ligacao de paginas por individuo
6. **Pagina de grupo** (Tiro) - graficos de barras e tabela comparativa
7. **Seccao por individuo:**
   - Divisora com nome
   - Tabela de metricas (todos os ensaios, ilimitado, auto-ajuste de colunas)
   - Estabilograma
   - Paginas de intervalo (Tiro)
   - Resumo por distancia (Tiro)
   - Hurdle Step (Tiro, se existir)
8. **Estatisticas do grupo** (se activado e n >= 2)
9. **Citacao academica** - bloco APA e BibTeX com ano dinamico

---

## 19. Termos de Utilizacao e Palavra-passe

**Termos:** aceitos na primeira execucao; guardados em `~/.aom_estabilidade.json`.

Para repor:
```bash
# Windows:
del "%USERPROFILE%\.aom_estabilidade.json"
# macOS / Linux:
rm ~/.aom_estabilidade.json
```

**Palavra-passe:** disponivel em https://github.com/andremassuca

**Alteracao de idioma:** disponivel no ecra de password e na janela principal.
A mudanca de idioma actualiza todos os labels da interface e dos outputs em tempo real.

---

## 20. Compilar para Distribuicao

### Windows -> BSP_Setup.exe

```
BUILD_Windows.bat
```

O script faz automaticamente:
1. Instala todas as dependencias via pip
2. Compila `BSP.exe` com PyInstaller
3. Compila `BSP_Uninstall.exe`
4. Empacota tudo em `BSP_Setup.exe`

O instalador:
- Instala em `%LOCALAPPDATA%\Programs\BSP\`
- Cria atalho no Desktop
- Regista em Adicionar/Remover Programas
- Inclui desinstalador grafico

### macOS -> BSP.dmg

```bash
chmod +x BUILD_macOS.sh && ./BUILD_macOS.sh
```

Compativel com macOS 12 Monterey+ (Intel e Apple Silicon M1/M2/M3).

---

## 21. Solucao de Problemas

| Problema | Solucao |
|---|---|
| `ModuleNotFoundError` | `pip install numpy scipy openpyxl matplotlib reportlab python-docx Pillow` |
| Validacao falha | Verificar pasta de individuos com sub-pastas validas |
| Aviso de seguranca macOS | Botao direito -> Abrir (uma vez) |
| PDF nao gerado | Campo PDF nao vazio e opcao activa |
| Testes estatisticos em cinzento | Activar opcao e n >= 3 |
| HTML nao abre | Verificar ligacao a internet (Chart.js CDN) |
| Aviso frame rate | Verificar exportacao da plataforma; pode ser caracteristica do equipamento |
| Programa fecha sem erro (v22) | Actualizar para v23 - corrige crash silencioso |
| Texto borrado em Windows 4K | Actualizar para v23 - DPI awareness activo |
| Numero de ensaios > 10 | Sem limite - PDF adapta colunas e fonte automaticamente |

---

## 22. Historico de Versoes

### v23 (2026)
- Crash silencioso ao iniciar protocolo resolvido (thread de analise via after())
- RMS ML, AP e Radius adicionados ao relatorio PDF
- Sobreposicao nas tabelas HurdleStep corrigida
- Inputs ilimitados: tabelas PDF auto-adaptam-se a qualquer numero de ensaios
- Validacao de frame rate entre ensaios (avisos no log)
- Exportacao PNG com DPI configuravel e seleccao de tipos
- Pagina "O que ha de novo" no PDF
- Autoria actualizada: Andre Massuca, Pedro Aleixo & Luis Massuca
- Todos os outputs sem referencias a instituicoes
- Citacoes com ano dinamico (ano do sistema)
- Em-dashes (-) substituidos em todo o codigo e outputs

### v22 (2025)
- Tema automatico (deteccao dark/light mode do SO)
- ETA na barra de progresso
- Perfis de configuracao nomeados
- Analise rapida - ficheiro unico (F2)
- Relatorio HTML interactivo Chart.js
- Citacao academica no PDF (APA + BibTeX)
- Cabeçalhos SPSS traduzidos por idioma
- Banner de actualizacao automatico

### v21 (2024)
- Protocolo de Tiro com analise de Selection CoP
- Hurdle Step bipodal
- Testes de Friedman e post-hoc
- Correlacao ea95 vs score
- Multilingua PT/EN/ES/DE completo

---

## 23. Referencias

- Schubert, P., & Kirchner, M. (2013). Ellipse area calculations and their applicability in posturography. *Gait & Posture*, 39(1), 518-522.
- Winter, D.A. (1995). Human balance and posture control during standing and walking. *Gait & Posture*, 3(4), 193-214.
- Prieto, T.E. et al. (1996). Measures of postural steadiness. *IEEE Trans. Biomed. Eng.*, 43(9), 956-966.
- Quijoux, F. et al. (2021). A review of centre of pressure (COP) variables to quantify standing balance in elderly people. *Physiological Reports*, 9(22), e15067.
- Paillard, T., & Noe, F. (2015). Techniques and criteria for assessing balance. *Frontiers in Physiology*, 6, 399.
- Maurer, C., & Peterka, R.J. (2005). A new interpretation of spontaneous sway measures based on a simple model of human postural control. *Journal of Neurophysiology*, 93(1), 189-200.
- Cohen, J. (1988). *Statistical Power Analysis for the Behavioral Sciences* (2nd ed.). Erlbaum.
- Iglewicz, B., & Hoaglin, D. (1993). *How to Detect and Handle Outliers*. ASQC Quality Press.

---

## Citacao Academica

**Formato APA:**
```
Massuca, A., & Massuca, L. (2026). BSP - Biomechanical Stability Program (v23).
https://github.com/andremassuca/BSP
```

**Formato BibTeX:**
```bibtex
@software{BSP_v23,
  author  = {Massuca, Andre and Massuca, Luis},
  title   = {BSP - Biomechanical Stability Program},
  year    = {2026},
  version = {23},
  url     = {https://github.com/andremassuca/BSP}
}
```

---

*BSP v23 - Andre Massuca & Luis Massuca*
