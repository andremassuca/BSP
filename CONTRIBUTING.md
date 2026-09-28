# Contributing to BSP

[Versão em português abaixo](#contribuir-para-o-bsp)

## Reporting a problem

1. Check that you are using the latest version of the repository.
2. Open an issue at [github.com/andremassuca/BSP/issues](https://github.com/andremassuca/BSP/issues) using the **Bug report** template, and include:
   - what you did, what you expected and what happened;
   - the protocol you used (FMS Bipodal, Single-leg Stance, Pistol Shooting or Archery);
   - your Windows version (or other operating system) and Python version (`python --version`);
   - the full error message or execution log, if there is one.
3. **Do not attach real participant data.** If you need to show a file, use the data in `examples/` or remove names and identifiers first.

## Proposing changes (pull requests)

1. Fork the repository and create a branch for your change (`git checkout -b fix-archery-reader`).
2. Make the change and run both test suites from `src/`:
   ```
   python bsp_core.py --testes
   python estabilidade_gui.py --testes
   ```
3. Open a pull request explaining what changes and why. Changes to how a metric is computed should cite the reference or test that supports them.

Issues and pull requests may be written in English or Portuguese.

---

<a id="contribuir-para-o-bsp"></a>
# Contribuir para o BSP

## Reportar um problema

1. Confirme que está a usar a versão mais recente do repositório.
2. Abra uma issue em [github.com/andremassuca/BSP/issues](https://github.com/andremassuca/BSP/issues) com o modelo **Bug report** e indique:
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

As issues e pull requests podem ser escritos em inglês ou em português.
