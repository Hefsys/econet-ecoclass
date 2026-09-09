# Automação EcoClass (Econet) — NCM / NBS

Consulta automática da ferramenta **EcoClass** do Econet para todos os NCM e NBS
distintos da planilha de simulação IBS/CBS, gerando um arquivo com o retorno
completo (CST, cClassTrib, cCredPres, % de redução, alíquotas, tratamento
tributário, base legal e vigência).

## Dois scripts (mesmo resultado, caminhos diferentes)

| Script | Como | Tempo | Quando usar |
|--------|------|-------|-------------|
| **`econet_ecoclass.py`** | Chama a API JSON do EcoClass | ~6 min | Padrão — lote do dia a dia |
| **`econet_ecoclass_rpa.py`** | Robô abre o Chrome e opera a tela | ~50–60 min | Fallback se a API mudar; conferência visual de poucos códigos |

Os dois leem a mesma planilha e o `_rpa` reaproveita a leitura do outro.

## Configuração inicial

1. `cp .env.example .env` e preencha `ECONET_LOGIN` / `ECONET_SENHA`.
2. Coloque a planilha de simulação (`Simulacao_*.xlsx`) nesta pasta, **ou**
   aponte `ECOCLASS_WORKBOOK` no `.env` para o caminho dela.

> `.env` e a planilha do cliente **não** vão para o git (ver `.gitignore`).

## Como rodar (API — recomendado)

```bash
pip3 install requests openpyxl        # só na primeira vez
python3 econet_ecoclass.py
```

O script leva ~6–8 minutos (≈ 220 códigos, 2 chamadas de API por código).

## Como rodar (RPA — fallback)

```bash
pip3 install playwright openpyxl
python3 -m playwright install chromium         # baixa o navegador (uma vez)

python3 econet_ecoclass_rpa.py                 # todos os códigos, sem janela
python3 econet_ecoclass_rpa.py 19059090 22021000   # só esses, com janela visível
```

Gera `Econet_EcoClass_RPA.xlsx` (uma linha por resultado, com o texto de cada aba
CST / cClassTrib / cCredPres / Alíquota + CST e cClassTrib já extraídos).

### Passo a passo do robô (`econet_ecoclass_rpa.py`)

1. Abre `app/eco-class`, cai no `login.asp`, preenche `Log`/`Sen`, habilita o botão
   de login via JS (o reCAPTCHA não é validado no servidor) e entra.
2. Para cada código: escolhe a categoria no seletor (`Produtos e Bens Materiais`
   ou `Serviços`), digita o código em `#input-v-6`, clica na lupa `i.mdi-magnify`.
3. Lê as linhas de `table tbody tr` (código + descrição + selo do benefício).
4. Clica na linha → tela de detalhe; percorre as abas `.tab-header-item`
   (CST → cClassTrib → cCredPres → Alíquota), expande os painéis
   (`.mdi-chevron-down`) e copia o `.tab-content`.
5. Volta com `.mdi-chevron-left`; repete para cada linha de resultado.
6. Recarrega a página para o próximo código e grava tudo no xlsx.

## O que ele faz

1. Lê os códigos das abas da planilha `Simulacao_IBS_CBS_*.xlsx`:
   | Aba | Coluna | Tipo |
   |-----|--------|------|
   | NFCe | M | NCM |
   | NFe_Saidas | M | NCM |
   | NFe_Entradas | N | NCM |
   | NFSe_Entradas | P | NBS |
   Ignora valores em branco, `Não informado` e códigos inválidos (ex.: `999999999`).

2. Autentica no portal Econet (login lido de `ECONET_LOGIN`) e obtém o token da API
   `ecoclass-api.econeteditora.com.br`. Não abre navegador — chama a API direto.

3. Para cada código chama `POST /api/ecoclass/pesquisa` e, quando há detalhamento,
   `POST /api/ecoclass/obter-dados`.

4. Grava:
   - **`Econet_EcoClass_Resultado.xlsx`** — 1+ linha por código, com ~30 colunas.
     Aba `EcoClass` (dados) + aba `Resumo`. Cores:
     verde = benefício, amarelo = tributação integral, vermelho = não localizado / erro.
   - **`Econet_EcoClass_bruto.jsonl`** — resposta crua da API por código (auditoria).

## Configuração

- **Login/senha**: `.env` (`ECONET_LOGIN`, `ECONET_SENHA`) — ver `.env.example`.
- **Planilha**: `ECOCLASS_WORKBOOK` no `.env`, ou o primeiro `Simulacao_*.xlsx` da pasta.
- **Ajuste fino** (colunas das abas, pausa entre chamadas): topo de `econet_ecoclass.py`
  (`SHEETS`, `PAUSA_ENTRE_CHAMADAS`).

## Pasta `lovable/`

Versão da consulta para rodar dentro de um app **Lovable** (Supabase Edge Function
que faz proxy + cache da API). Ver `lovable/LOVABLE.md`.

## Observação importante (segmento restaurantes)

O EcoClass devolve a classificação **geral** da LC 214/2025 por NCM/NBS.
A empresa atua em **restaurantes / fornecimento de alimentação**, que tem regra
específica (regime do art. 274 e seguintes / redução setorial). Portanto os
`CST` / `cClassTrib` retornados **não são aplicáveis diretamente** — servem como
base para análise e comparação com o tratamento que a planilha já adota.

## Como a resposta é interpretada

| Situação na API | Status na planilha | CST / cClassTrib |
|---|---|---|
| `tipo_beneficio` preenchido | `Beneficio: <tipo>` | conforme retorno (ex.: 200 / 200003) |
| `tributacao_integral = true` | `Tributacao integral` | 000 / 000001, % CBS/IBS = 100 |
| lista vazia (não consta na LC 214) | `Nao localizado -> CST 000` | 000 / 000001 (regra da própria ferramenta) |
| falha de rede após 4 tentativas | `ERRO: ...` | — (rodar de novo) |
