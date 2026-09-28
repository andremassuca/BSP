# Contribuir para o BSP

[English below](#english)

## Reportar um problema

1. Confirme que está a usar a versão mais recente do repositório.
2. Abra uma issue em [github.com/andremassuca/BSP/issues](https://github.com/andremassuca/BSP/issues) e indique:
   - o que fez, o que esperava e o que aconteceu;
   - o protocolo usado (FMS Bipodal, Apoio Unipodal, Tiro ou Tiro com Arco);
   - a versão do Windows (ou outro sistema) e do Python (`python --version`);
   - a mensagem de erro completa ou o registo de execução, se existir.
3. **Não anexe dados reais de participantes.** Se precisar de mostrar um ficheiro, use os dados de `examples/` ou apague os nomes e identificadores antes de o enviar.

## Propor alterações (pull requests)

1. Faça um fork do repositório e crie um ramo para a alteração (`git checkout -b corrige-leitura-arco`).
2. Faça a alteração e corra as duas suites de testes a partir de `src/`:
   ```
   python bsp_core.py --testes
   python estabilidade_gui.py --testes
   ```
3. Abra um pull request a explicar o que muda e porquê. Alterações que mudem o cálculo de uma métrica devem indicar a referência bibliográfica ou o teste que as justifica.

As issues e pull requests podem ser escritos em português ou em inglês.

---

<a id="english"></a>
## Contributing to BSP

**Reporting a problem:** open an issue at [github.com/andremassuca/BSP/issues](https://github.com/andremassuca/BSP/issues) describing what you did, what you expected and what happened, together with the protocol, your operating system, your Python version and the full error message. Please do not attach real participant data; use the files in `examples/` or remove names and identifiers first.

**Proposing changes:** fork the repository, create a branch, run both test suites from `src/` (`python bsp_core.py --testes` and `python estabilidade_gui.py --testes`) and open a pull request explaining what changes and why. Changes to how a metric is computed should cite the reference or test that supports them.

Issues and pull requests may be written in Portuguese or English.
