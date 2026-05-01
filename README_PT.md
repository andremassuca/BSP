<picture>
  <source media="(prefers-color-scheme: dark)" srcset="branding/bsp_banner_dark.png">
  <source media="(prefers-color-scheme: light)" srcset="branding/bsp_banner_light.png">
  <img alt="BSP" src="branding/bsp_banner_light.png">
</picture>

# BSP - Biomechanical Stability Program

**O que é:** uma aplicação para analisar a estabilidade postural de atletas a partir de dados recolhidos em plataforma de forças. Faço isto porque havia muita coisa feita à mão em Excel que dava para automatizar e devolver em métricas limpas.

**Para quem:** investigadores, treinadores e clínicos que usam plataforma de forças e querem sair do CoP bruto para métricas úteis (ea95, RMS, velocidade média, rigidez postural) sem escrever código.

## O que faz

Quatro protocolos:

- **FMS Bipodal** - teste bipodal simples com painel esquerdo/direito opcional.
- **Apoio Unipodal** - assimetria dominante/não-dominante.
- **Tiro** - duas janelas temporais de estabilização antes do disparo.
- **Tiro com Arco** - até 30 ensaios, janela única entre confirmação 1 e 2.

Em qualquer protocolo calcula:

- Área da elipse de 95% (`ea95`), RMS e amplitudes em X/Y e radial.
- Velocidade média e rigidez postural (`stiff_x`, `stiff_y`).
- Assimetria dominante/não-dominante quando aplicável.
- FFT do CoP para inspecção espectral.

Gera: Excel com resumo, per-ensaio e demografia (quando há referência demográfica), PDF individual por atleta e PDF de grupo, e um dashboard web local opcional com listas interactivas, box plots, scatter com regressão e detalhe por atleta.

## Quickstart em 1 minuto

```bash
git clone https://github.com/andremassuca/BSP.git
cd BSP
pip install -r requirements.txt
python estabilidade_gui.py
```

Primeira corrida abre a janela de licença → palavra-passe → escolha de protocolo. A palavra-passe está no documento de acesso que enviei separadamente.

Para correr os testes sintéticos antes de processar dados reais:

```bash
python estabilidade_gui.py --testes     # GUI/PDF/Excel (~161 testes)
python -m bsp_core --testes             # só core matemático (~41 testes)
```

Ambos devem reportar **100% passou**.

## Como trabalhar com os dados

**Convenção de pastas:** uma pasta por recolha. Cada atleta tem o ID no nome do ficheiro.

**Tiro com Arco**

- Pasta com ZIPs ou sub-pastas, cada ficheiro `{id}_{ensaio} - {DD-MM-AAAA} - Stability export.xls` (tab-separated, 50 Hz).
- Ficheiro `Inicio_fim_vfinal.xlsx` com três folhas (`tempo do toque`, `confirmação_1`, `confirmação_2`). A janela de análise é `confirmação_1 → confirmação_2`. O loader lida com nomes corrompidos em cp1252 automaticamente.
- Ficheiro `Todos os registos dos 142 atletas em JUl_2024 _.xlsx` como referência demográfica. Colunas PESO, ALTURA, IDADE, ESTILO, CATEGORIA, GÉNERO, P1..P30 e P_TOTAL. Colunas 68–308 são ignoradas.

**Outros protocolos**

- Ficheiros exportados pelo software da plataforma, em `.xlsx` ou `.txt` tab-separated.
- `inicio_fim.xlsx` para Tiro, com as duas janelas temporais por ensaio.


## Problemas conhecidos

- **macOS, primeira execução:** Gatekeeper bloqueia apps não notarizadas. **Clique direito → Abrir**, confirmar. Só é preciso uma vez.
- **macOS, app não arranca:** verificar `~/Library/Logs/BSP/BSP_crash.log`. Se nada lá está, correr `./BUILD_macOS.sh` a partir do terminal e ler `build_logs/`.
- **Windows, DLL em falta:** instalar [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe).
- **Confirmação_1/2 não lê o ficheiro:** verificar se os nomes das folhas não estão corrompidos. Abrir, renomear para `confirmação_1` e `confirmação_2`, gravar.
- **Dashboard não abre no browser:** o uvicorn pode levar 1–3 s a arrancar. Se após 10 s nada aparecer, verificar se `fastapi`/`uvicorn` estão instalados.

## Citar

Se usar o BSP num artigo, por favor cite:

Massuça, A. O., Aleixo, P., & Massuça, L. M. (2026). *BSP: Biomechanical Stability Program* (Versão 1.0) [Software]. https://github.com/andremassuca/BSP

```bibtex
@software{massuca_bsp_2025,
  author  = {Massu\c{c}a, Andr\'{e} Oliveira and Aleixo, Pedro and Massu\c{c}a, Lu\'{i}s M.},
  title   = {BSP: Biomechanical Stability Program},
  version = {1.0},
  year    = {2026},
  url     = {https://github.com/andremassuca/BSP},
}
```

Métodos relevantes:

- Schubert, P., & Kirchner, M. (2013). *Ellipse area calculations and their applicability in posturography.* Gait & Posture, 39(1), 518–522.
- Prieto, T. E. et al. (1996). *Measures of postural steadiness.* IEEE Trans. Biomed. Eng., 43(9), 956–966.
- Winter, D. A. (1995). *Human balance and posture control during standing and walking.* Gait & Posture, 3(4), 193–214.
- Quijoux, F. et al. (2021). *A review of center of pressure (COP) variables to quantify standing balance in elderly people.* Physiological Reports, 9(22), e15067.
- Maurer, C., & Peterka, R. J. (2005). *A new interpretation of spontaneous sway measures based on a simple model of human postural control.* J. Neurophysiol., 93(1), 189–200.

## Licença

Código académico, livre para uso em investigação e docência. Para uso comercial ou redistribuição, contactar o autor.

**André O. Massuça** - [github.com/andremassuca](https://github.com/andremassuca)
**Pedro Aleixo**
**Luís M. Massuça**
