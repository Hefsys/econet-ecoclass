// Supabase Edge Function (Deno)  ─  proxy + cache da API do EcoClass (Econet)
//
// Deploy:  supabase functions deploy ecoclass  (ou pelo painel do Lovable)
// Secrets: supabase secrets set ECONET_LOGIN=<login> ECONET_SENHA=<senha>
//
// Chamada a partir do front-end (Lovable / React):
//   const { data } = await supabase.functions.invoke("ecoclass", {
//     body: { tipo: "NCM", codigo: "19059090" }          // 1 código
//   });
//   // ou em lote (máx. 40 por chamada):
//   body: { itens: [{ tipo:"NCM", codigo:"19059090" }, { tipo:"NBS", codigo:"115090000" }], refresh:false }
//
// Tabela de cache (crie no Lovable / SQL):
//   create table ecoclass_resultado (
//     tipo text not null,
//     codigo text not null,
//     status text,
//     ncm_nbs_retornado text,
//     descricao text,
//     tipo_beneficio text,
//     cst text,
//     cst_descricao text,
//     cclasstrib text,
//     cclasstrib_descricao text,
//     reducao_aliquota boolean,
//     perc_cbs numeric,
//     perc_ibs numeric,
//     base_legal text,
//     inicio_vigencia date,
//     bruto jsonb,
//     atualizado_em timestamptz default now(),
//     primary key (tipo, codigo, ncm_nbs_retornado)
//   );

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36";
const APP = "https://app.econeteditora.com.br";
const PAG = `${APP}/app/eco-class`;
const CAPTCHA =
  "https://1u9nevrqx5.execute-api.sa-east-1.amazonaws.com/default/go-gerar-captcha";
const CAPTCHA_XKEY = "4XQdAIUfrJ1Zpk2Jh0uRyaD5cNywcxQA9E5pvjJI";
const ECO = "https://ecoclass-api.econeteditora.com.br/api/ecoclass";
const TIPO_API: Record<string, string> = { NCM: "produto_mercadoria", NBS: "servico" };

const LOGIN = Deno.env.get("ECONET_LOGIN")!;
const SENHA = Deno.env.get("ECONET_SENHA")!;

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

// ─── autenticação (JWT em cache no runtime, ~24h) ──────────────────────────
let jwtCache: { token: string; exp: number } | null = null;

async function loginCookie(): Promise<string> {
  await fetch(`https://www.econeteditora.com.br/user/login.asp?Pag=${encodeURIComponent(PAG)}`, {
    headers: { "User-Agent": UA },
  });
  const qs = new URLSearchParams({
    Pag: PAG, Origem: "Comex", Log: LOGIN, Sen: SENHA, "g-recaptcha-response": "",
  });
  const r = await fetch(`https://www.econeteditora.com.br/user/ver_log.asp?${qs}`, {
    headers: { "User-Agent": UA },
    redirect: "manual",
  });
  const c = r.headers.getSetCookie().find((x) => x.startsWith("bG9naW4="));
  if (!c) throw new Error("Login Econet falhou (verifique ECONET_LOGIN / ECONET_SENHA).");
  return decodeURIComponent(c.split(";")[0].slice("bG9naW4=".length));
}

async function getJwt(force = false): Promise<string> {
  if (!force && jwtCache && jwtCache.exp > Date.now() / 1000 + 60) return jwtCache.token;
  const cookie = await loginCookie();
  const r = await fetch("https://go-login.econeteditora.com.br/api/login", {
    method: "POST",
    headers: { "Content-Type": "application/json", "User-Agent": UA, Referer: `${APP}/` },
    body: JSON.stringify({ cookie }),
  });
  if (!r.ok && r.status !== 201) throw new Error(`go-login ${r.status}: ${await r.text()}`);
  const token = (await r.text()).trim().replace(/^"|"$/g, "");
  let exp = Date.now() / 1000 + 3600;
  try { exp = JSON.parse(atob(token.split(".")[1])).exp ?? exp; } catch { /* noop */ }
  jwtCache = { token, exp };
  return token;
}

// ─── chamadas ao EcoClass ─────────────────────────────────────────────────
async function captchaToken(target: string): Promise<string> {
  const r = await fetch(`${CAPTCHA}?path=${encodeURIComponent(target)}`, {
    headers: { "x-api-key": CAPTCHA_XKEY, Referer: `${APP}/` },
  });
  return (await r.json()).token;
}

async function ecoPost(path: string, body: unknown, retry = true): Promise<any> {
  const url = `${ECO}/${path}`;
  const jwt = await getJwt();
  const r = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: jwt,
      "g-recaptcha-response": await captchaToken(url),
      "Content-Type": "application/json",
      "User-Agent": UA,
      Origin: APP,
      Referer: `${APP}/`,
    },
    body: JSON.stringify(body),
  });
  if (r.status === 401 && retry) {
    await getJwt(true);
    return ecoPost(path, body, false);
  }
  if (!r.ok) throw new Error(`${path} ${r.status}: ${await r.text()}`);
  return r.json();
}

// ─── normalização (1 código -> 1..n linhas) ───────────────────────────────
const sn = (v: unknown) => v === 1 || v === true;
const bl = (arr: any[]) => (arr ?? []).map((b) => b.base_legal).filter(Boolean).join(" | ");

async function consulta(tipo: string, codigo: string) {
  const dig = codigo.replace(/\D/g, "");
  const resp = await ecoPost("pesquisa", { tipo: TIPO_API[tipo], pesquisa: dig });
  const data = resp?.data ?? {};
  const chave = tipo === "NCM" ? "ncm" : "nbs";
  const itens: any[] = data[chave] ?? [];

  if (itens.length === 0) {
    return [{
      tipo, codigo,
      status: "Nao localizado na LC 214/2025 -> CST 000 (Tributacao Integral)",
      ncm_nbs_retornado: "", descricao: "", tipo_beneficio: null,
      cst: "000", cst_descricao: "Tributacao integral",
      cclasstrib: "000001", cclasstrib_descricao: "Situacoes tributadas integralmente pelo IBS e CBS.",
      reducao_aliquota: false, perc_cbs: 100, perc_ibs: 100,
      base_legal: "", inicio_vigencia: null, bruto: resp,
    }];
  }

  const linhas = [];
  for (const item of itens) {
    const integral = !!item.tributacao_integral;
    const cc = item.cclasstrib_cbs_ibs ?? {};
    let det: any = null;
    if (item.id != null && (item.id_cclasstrib_cbs_ibs != null || integral)) {
      try {
        det = await ecoPost("obter-dados", {
          id: item.id,
          id_cclasstrib_cbs_ibs: item.id_cclasstrib_cbs_ibs ?? null,
          id_ccredpres: item.id_ccredpres ?? null,
          tipo: TIPO_API[tipo],
          tributacao_integral: integral,
          nbs_codigo: tipo === "NBS" ? (item.nbs ?? "").replace(/\D/g, "") : null,
        });
      } catch { /* segue sem detalhe */ }
    }
    const c0 = det?.data?.cClassTrib?.[0] ?? {};
    const cstd = c0.cst_cbs_ibs ?? {};
    const ali = det?.data?.aliquota ?? {};

    linhas.push({
      tipo, codigo,
      status: integral && !item.tipo_beneficio
        ? "Tributacao integral"
        : `Beneficio: ${item.tipo_beneficio ?? "ver cClassTrib"}`,
      ncm_nbs_retornado: item.ncm ?? item.nbs ?? "",
      descricao: item.descricao ?? "",
      tipo_beneficio: item.tipo_beneficio ?? null,
      cst: cstd.codigo ?? (integral ? "000" : ""),
      cst_descricao: cstd.descricao ?? (integral ? "Tributacao integral" : ""),
      cclasstrib: c0.cclasstrib ?? cc.cclasstrib ?? (integral ? "000001" : ""),
      cclasstrib_descricao: c0.descricao ?? cc.descricao ?? "",
      reducao_aliquota: sn(c0.reducao_aliquota ?? cc.reducao_aliquota),
      perc_cbs: item.valor_cbs != null ? Number(item.valor_cbs) : (integral ? 100 : null),
      perc_ibs: item.valor_ibs != null ? Number(item.valor_ibs) : (integral ? 100 : null),
      base_legal: bl(item.base_legal) || bl(c0.base_legal),
      inicio_vigencia: item.inicio_vigencia ?? c0.inicio_vigencia ?? null,
      aliquota_cbs_ref: ali?.cbs?.aliquota ?? null,
      aliquota_ibs_ref: ali?.ibs?.aliquota ?? null,
      bruto: { pesquisa: { data: { [chave]: [item] } }, obter_dados: det },
    });
  }
  return linhas;
}

// ─── handler ──────────────────────────────────────────────────────────────
Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const body = await req.json();
    const itens: { tipo: string; codigo: string }[] =
      body.itens ?? (body.codigo ? [{ tipo: body.tipo ?? "NCM", codigo: body.codigo }] : []);
    if (itens.length === 0) throw new Error("informe { tipo, codigo } ou { itens: [...] }");
    if (itens.length > 40) throw new Error("máx. 40 itens por chamada (faça em lotes)");
    const refresh = !!body.refresh;

    const sb = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const out: any[] = [];
    for (const { tipo, codigo } of itens) {
      if (!refresh) {
        const { data: cached } = await sb
          .from("ecoclass_resultado")
          .select("*")
          .eq("tipo", tipo).eq("codigo", codigo);
        if (cached && cached.length) { out.push(...cached); continue; }
      }
      let linhas;
      try {
        linhas = await consulta(tipo, codigo);
      } catch (e) {
        out.push({ tipo, codigo, status: `ERRO: ${e.message}` });
        continue;
      }
      await sb.from("ecoclass_resultado")
        .delete().eq("tipo", tipo).eq("codigo", codigo);
      await sb.from("ecoclass_resultado").insert(linhas);
      out.push(...linhas);
    }

    return new Response(JSON.stringify({ resultados: out }), {
      headers: { ...CORS, "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), {
      status: 400,
      headers: { ...CORS, "Content-Type": "application/json" },
    });
  }
});
