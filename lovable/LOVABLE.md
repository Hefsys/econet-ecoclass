# Levar a consulta EcoClass para o Lovable

## Qual dos dois cenários dá para readaptar

| Cenário | Vai para o Lovable? | Por quê |
|---|---|---|
| **API** (`econet_ecoclass.py`) | ✅ **Sim** | São só chamadas HTTPS. Rodam dentro de uma **Supabase Edge Function** (Deno), que o Lovable já usa como back-end. |
| **RPA** (`econet_ecoclass_rpa.py`) | ❌ Não | Precisa de um navegador (Chromium + Playwright) e de um processo que fica de pé. O back-end do Lovable é serverless (Deno, sem browser, timeout ~150 s). Só rodaria num worker separado (VPS, Railway, Browserless) — aí não é mais "no Lovable". |

**Resposta curta: use o cenário da API.** O RPA fica como utilitário local de conferência.

## Arquitetura no Lovable

```
React (Lovable)  ──►  Edge Function "ecoclass"  ──►  API EcoClass (Econet)
      ▲                      │
      │                      ▼
      └──────────  Tabela  ecoclass_resultado  (cache no Supabase)
```

- O **front-end nunca fala direto com a Econet** (CORS + o login não pode ir para o browser).
- A Edge Function faz login (JWT em cache ~24 h), gera o token de captcha, chama
  `pesquisa` + `obter-dados`, **grava no cache** e devolve o resultado normalizado.
- Login e senha ficam em **Supabase Secrets**, nunca no código do front.

## Passos

1. **Criar a tabela de cache** (SQL Editor do Supabase, dentro do Lovable):
   o `create table ecoclass_resultado (...)` está comentado no topo de
   `supabase/functions/ecoclass/index.ts`.

2. **Publicar a função**: copie `supabase/functions/ecoclass/index.ts` para o
   projeto (o Lovable cria as Edge Functions em `supabase/functions/<nome>/`).

3. **Configurar os secrets**:
   ```
   ECONET_LOGIN = <login do Econet>
   ECONET_SENHA = <senha do Econet>
   ```
   (`SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY` já existem no ambiente da função.)

4. **Chamar do front (React)**:
   ```ts
   // consulta única
   const { data } = await supabase.functions.invoke("ecoclass", {
     body: { tipo: "NCM", codigo: "19059090" },
   });

   // lote (máx. 40 por chamada – itere em blocos para listas maiores)
   const { data } = await supabase.functions.invoke("ecoclass", {
     body: { itens: [
       { tipo: "NCM", codigo: "19059090" },
       { tipo: "NBS", codigo: "115090000" },
     ], refresh: false },
   });
   ```
   `refresh: true` ignora o cache e reconsulta a Econet.

5. **Importar a planilha no app** (opcional): parsear o XLSX no browser com
   `xlsx` (SheetJS), extrair as colunas de NCM/NBS (mesma regra do
   `econet_ecoclass.py`: NFCe!M, NFe_Saidas!M, NFe_Entradas!N, NFSe_Entradas!P),
   deduplicar e mandar em lotes de 40 para a função.

## Retorno da função

```json
{
  "resultados": [
    {
      "tipo": "NCM", "codigo": "19059090",
      "status": "Beneficio: Redução de Alíquota",
      "ncm_nbs_retornado": "1905.90.90",
      "descricao": "Pão comumente denominado pão francês...",
      "tipo_beneficio": "Redução de Alíquota",
      "cst": "200", "cst_descricao": "Alíquota reduzida",
      "cclasstrib": "200003", "cclasstrib_descricao": "Vendas de produtos...",
      "reducao_aliquota": true, "perc_cbs": 100, "perc_ibs": 100,
      "base_legal": "Artigo 125 da Lei Complementar n° 214/2025 | Anexo I...",
      "inicio_vigencia": "2026-01-01",
      "bruto": { "pesquisa": {...}, "obter_dados": {...} }
    }
  ]
}
```

Um mesmo código pode devolver **várias linhas** (regimes concorrentes: cesta
básica, insumo agropecuário, diferimento etc.) — igual à versão por API.

## Observações

- **Login compartilhado**: é uma conta só. Se muita gente usar o app ao mesmo
  tempo a Econet pode limitar. O cache resolve 95% disso (os ~220 códigos mudam
  pouco). Considere um botão "Atualizar" manual em vez de reconsultar sempre.
- **Segmento restaurantes**: continua valendo o aviso — o EcoClass devolve a
  classificação geral da LC 214; a regra do restaurante é análise por cima.
- **Termos de uso da Econet**: uso interno da consultoria, com a conta do
  cliente. Não exponha o app publicamente sem autenticação.
- Se a Econet trocar `x-api-key`, host da API (`ecoclass-api...`) ou o formato do
  captcha, é só ajustar as constantes no topo do `index.ts` — mesmos pontos do
  `econet_ecoclass.py`.
