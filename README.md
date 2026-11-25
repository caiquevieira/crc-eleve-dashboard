# CRM e Dashboard de Produtividade CRC · Eleve

- **CRM** (cadastro e acompanhamento dos agendamentos): https://caiquevieira.github.io/crc-eleve-dashboard/crm/
- **Dashboard**: https://caiquevieira.github.io/crc-eleve-dashboard/

A planilha do Google é o banco de dados: o CRM lê e grava direto na aba **Produtividade**,
e o dashboard é gerado a partir dela.

## CRM (Google Apps Script)

Código em `apps-script/`. Ele roda dentro da própria planilha, com login pela conta Google:
só quem tem **permissão de edição** na planilha consegue abrir o CRM.

Recursos: visões Hoje / Amanhã / Próximos 7 dias / A confirmar / Status pendente / Todos,
busca por nome ou telefone, link direto para o WhatsApp, mudança rápida de confirmação e
status na lista, e formulário completo para novo agendamento ou edição.

### Instalação (uma vez, na conta dona da planilha)

1. Abra o .xlsx no Drive → **Arquivo → Salvar como Planilhas Google**. Use essa nova planilha daqui em diante.
2. Compartilhe a nova planilha com as CRCs como **Editor**. Para o dashboard atualizar sozinho,
   deixe também **"Qualquer pessoa com o link: Leitor"**.
3. Na planilha: **Extensões → Apps Script**.
   - Substitua o conteúdo de `Código.gs` pelo de `apps-script/Code.gs`.
   - **+ → HTML**, nome `Index`, e cole `apps-script/Index.html`.
   - Salve.
4. **Implantar → Nova implantação → App da Web**
   - Executar como: **Usuário que acessa o app da Web**
   - Quem pode acessar: **Qualquer pessoa com uma Conta do Google**
   - Autorize o acesso. Como o app não é verificado pelo Google, clique em
     *Avançado → Acessar CRM (não seguro)*; isso é normal para scripts internos.
5. Copie a URL do app (termina em `/exec`) e, no GitHub:
   - Secret `SHEET_FILE_ID` = ID da **nova** planilha
   - Variable `CRM_URL` = URL do app (*Settings → Secrets and variables → Actions → Variables*)
   - Rode *Actions → Atualizar dashboard → Run workflow*.

Ao alterar o código do CRM: **Implantar → Gerenciar implantações → editar → Nova versão**
(assim a URL continua a mesma).

Cada CRC escolhe o próprio nome em "Você é" no primeiro acesso; os novos agendamentos já vêm com ele.
## Dashboard

Painel publicado no GitHub Pages com os indicadores da aba **Produtividade** da
planilha "CRC - ELEVE Leads Agendamentos" (Google Drive).

## Dashboard: como funciona

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
