# Palavra-passe de acesso rotativa

A palavra-passe que aparece no ecrã de acesso do BSP **não está colada no binário**. O hash SHA256 da password actual está num ficheiro público no repositório (`.bsp_pass.sha256`) e cada versão instalada da app consulta esse ficheiro no arranque.

Isto significa que, para mudar a password em todas as instalações existentes, **basta um commit + push**. Não é preciso publicar nova versão, não é preciso que os utilizadores instalem nada. O próximo arranque da app (com ligação à Internet) apanha o novo hash.

## Como funciona

1. Ao arrancar, o ecrã de password dispara um fetch em background a:
   ```
   https://raw.githubusercontent.com/andremassuca/BSP/main/.bsp_pass.sha256
   ```
2. O conteúdo é parseado linha a linha. Cada linha que seja 64 caracteres hex lowercase é considerada um hash SHA256 válido. `#` começa um comentário.
3. Se o fetch responde com pelo menos um hash válido:
   - A lista é guardada em cache local (`~/.bsp_config.json` ou equivalente) sob a chave `_pass_hashes_cache`, com timestamp.
   - A app passa a aceitar **apenas** esses hashes. Qualquer password antiga que já não esteja na lista **deixa de funcionar** - esta é a rotação efectiva.
4. Se o fetch falha (sem rede, GitHub em baixo, etc.):
   - Usa a cache local se existir.
   - Se não há cache (primeira execução offline), usa o hash embedded no binário (bootstrap).

## Rotação - passo a passo

### 1. Gerar o hash da nova password

```bash
python -c "import hashlib; print(hashlib.sha256(b'MINHA_NOVA_PASSWORD').hexdigest())"
```

Saída exemplo:
```
a7b4...  (64 chars hex)
```

### 2. Editar `.bsp_pass.sha256` na raiz do repo

Substituir a linha actual pelo novo hash:

```
# ... comentários de topo ...

a7b4c8...novohash...aabbccdd
```

### 3. Commit + push

```bash
git add .bsp_pass.sha256
git commit -m "rotate access password"
git push origin main
```

Nos segundos seguintes, qualquer app que arranque e tenha rede passa a exigir a nova password. Caches locais são actualizadas automaticamente no primeiro fetch com sucesso.

## Grace period (rotação suave)

Se quiseres dar tempo aos utilizadores para trocarem a password que têm anotada, podes deixar **os dois hashes** no ficheiro durante algum tempo:

```
# Password nova (canónica)
a7b4c8...novohash...aabbccdd

# Password antiga (aceite até <data>)
76c10573cf1d0db18f6b87f4c6f7bcb393d62fc47eac43cea1aca9c5d169a0d9
```

Ambas funcionam. Quando chegar a data, remove a linha antiga e faz push - aí a password antiga deixa de ser aceite.

## Revogar uma password comprometida

Se descobrires que a password foi partilhada indevidamente:

1. Escolhe uma password nova.
2. Substitui o hash em `.bsp_pass.sha256` pelo novo **e remove o antigo imediatamente** (não uses grace period).
3. Push.

No próximo arranque com rede, todos os utilizadores têm de saber a password nova. Quem estiver offline continua a poder entrar com a antiga **até que a máquina volte a ter rede** - aí a cache é actualizada.

## Offline na primeira execução

A app tem o último hash conhecido **embedded no binário** (`_PASS_HASH`). Isto é apenas para a primeira execução quando ainda não há cache local e não há rede. Em qualquer execução subsequente com rede, o embedded é substituído pelo remoto.

Quando compilares uma nova versão (`BSP_Setup.exe`/`BSP.dmg`), o embedded fica com o valor corrente - convém alinhar com o ficheiro remoto no momento do build para reduzir surpresas offline.

## Segurança

- **O repo é público.** O ficheiro `.bsp_pass.sha256` é visível para toda a gente. Isto é aceitável porque:
  - O hash SHA256 não é reversível: não permite recuperar a password plaintext.
  - O controlo de acesso real é saber a password; o hash apenas valida.
- **Não comitar a password em plaintext.** Nunca. Só o hash.
- **Evitar passwords curtas ou de dicionário.** Se a password for fraca, um atacante pode fazer brute-force ao hash em tempo viável. Usa pelo menos 12 caracteres aleatórios (ou uma frase longa).
- **Rotação após partilha acidental.** Se uma pessoa a quem não confias souber a password, roda - é barato.

## Testar

Para verificar que o mecanismo funciona:

```bash
python estabilidade_gui.py --testes
# Procurar seccao [16] Palavra-passe rotativa - todos os testes devem passar.
```

Para testar manualmente a rotação:

1. Anotar a password actual.
2. Gerar um hash novo para uma password diferente e fazer push.
3. Abrir a app → o ecrã de acesso rejeita a password antiga.
4. Verificar em `~/.bsp_config.json` (ou equivalente): o campo `_pass_hashes_cache` mostra o novo hash.

## Ficheiros envolvidos

- [`.bsp_pass.sha256`](../.bsp_pass.sha256) - ficheiro público com os hashes aceites.
- [`estabilidade_gui.py`](../estabilidade_gui.py) - funções `_fetch_pass_hashes_remoto` e `_obter_pass_hashes_aceites`; uso no `_ecra_password`.
