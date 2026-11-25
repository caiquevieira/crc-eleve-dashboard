# Dashboard de Produtividade CRC · Eleve

https://caiquevieira.github.io/crc-eleve-dashboard/

Painel publicado no GitHub Pages com os indicadores da aba **Produtividade** da
planilha "CRC - ELEVE Leads Agendamentos" (Google Drive).

## Como funciona

1. A cada 30 minutos (07h–20h, seg a sáb) o GitHub Actions baixa a planilha do Drive.
2. `scripts/build_data.py` lê a aba Produtividade e gera `data.json` só com campos
   operacionais: datas, hora, campanha, CRC, clínica, confirmação, meio, status,
   fechou/motivo e valores.
3. A página (`site/index.html`) lê esse JSON e calcula os indicadores no navegador.

**Nome, telefone, data de nascimento e observação nunca são publicados.** A planilha
e o `data.json` não entram no repositório (veja `.gitignore`).

## Indicadores

| Indicador | Cálculo |
|---|---|
| Agendamentos | linhas com data do agendamento ou nome preenchido |
| Confirmação feita | FEZ CONFIRMAÇÃO = SIM ÷ agendamentos |
| Comparecimento | Compareceu ÷ (Compareceu + Faltou); quem tem FECHOU = Sim conta como compareceu |
| Fechamentos / conversão | FECHOU = Sim; conversão = fechamentos ÷ compareceram |
| Valor vendido / ticket médio | soma e média de VALOR DE VENDA |

Filtros: período (pela data do agendamento ou da consulta), CRC, campanha, clínica e meio.
Os filtros ficam na URL, então dá para compartilhar um link já filtrado.

## Configuração (uma vez)

1. A planilha precisa estar compartilhada como **"Qualquer pessoa com o link: Leitor"**.
2. No repositório: *Settings → Secrets and variables → Actions → New repository secret*
   - Nome: `SHEET_FILE_ID`
   - Valor: o ID do arquivo (o trecho entre `/d/` e `/edit` no link da planilha)
3. *Settings → Pages → Source*: **GitHub Actions**.
4. *Actions → Atualizar dashboard → Run workflow* para a primeira publicação.

O ID fica em segredo de propósito: o repositório é público e quem tiver o ID baixa a
planilha completa.

## Manutenção

- **Atualizar agora:** Actions → Atualizar dashboard → Run workflow.
- **Rodar localmente:**
  ```
  pip install openpyxl
  python scripts/build_data.py --xlsx "CRC - ELEVE Leads Agendamentos.xlsx" --out site/data.json
  python -m http.server -d site 8000
  ```
- O GitHub pausa agendamentos de repositórios sem commits há 60 dias. Se o painel parar
  de atualizar, reative em Actions.
- Ao adicionar colunas novas na aba, ajuste os índices `COL_*` no topo de `scripts/build_data.py`.

## Botão "Atualizar agora"

No topo do painel. Pede login e senha de administrador (com o responsável; não ficam no
repositório) e dispara a atualização na hora, recarregando o painel quando termina.

Como funciona: um token do GitHub com permissão **apenas de Actions** neste repositório fica
criptografado em `site/auth.json`; o login e a senha são a chave. Para configurar ou trocar a senha:

1. Crie um token *fine-grained* em GitHub → Settings → Developer settings → Personal access tokens:
   repositório `crc-eleve-dashboard`, permissão **Actions: Read and write**.
2. Cole o token em `token.txt` na raiz do projeto (o arquivo é ignorado pelo git).
3. `node scripts/encrypt_token.mjs <login> <senha>` → gera `site/auth.json` e apaga `token.txt`.
4. Faça commit e push de `site/auth.json`.
