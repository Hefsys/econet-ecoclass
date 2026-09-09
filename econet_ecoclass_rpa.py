#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Versao RPA (robo que opera a tela) da consulta ao EcoClass.

Faz exatamente o que uma pessoa faria no site:
  abre o navegador -> faz login -> escolhe "Produtos e Bens Materiais" / "Servicos"
  -> digita o NCM/NBS -> clica na lupa -> abre o resultado -> percorre as abas
  CST / cClassTrib / cCredPres / Aliquota -> copia o conteudo -> volta -> proximo.

Quando usar esta versao (mais lenta, ~15 s por codigo):
  - se a API do EcoClass mudar / sair do ar;
  - para conferir visualmente um punhado de codigos (MODO_CONFERENCIA);
  - se a area juridica preferir "clicar como humano" em vez de chamar a API.

Para o processamento em lote do dia a dia use "econet_ecoclass.py" (via API, ~6 min).

Requisitos:
  pip3 install playwright openpyxl
  python3 -m playwright install chromium      # baixa o navegador (uma vez)

Uso:
  python3 econet_ecoclass_rpa.py                 # todos os codigos, headless
  python3 econet_ecoclass_rpa.py 19059090 22021000   # so esses codigos, com janela visivel
"""

import re
import sys
import time
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# reaproveita a leitura da planilha e as constantes da versao por API
from econet_ecoclass import (
    WORKBOOK, BASE_DIR, ECONET_LOGIN, ECONET_SENHA,
    PAG_DESTINO, extrair_codigos,
)

OUT_XLSX = BASE_DIR / "Econet_EcoClass_RPA.xlsx"
HEADLESS_PADRAO = True            # False abre a janela do Chrome (bom p/ depurar)
TIMEOUT = 30_000                  # ms
PAUSA = 0.8                       # s entre passos


# --------------------------------------------------------------------------- #
# Operacao da tela
# --------------------------------------------------------------------------- #
class EcoClassRPA:
    def __init__(self, page):
        self.p = page

    # -- login -----------------------------------------------------------------
    def login(self):
        self.p.goto(PAG_DESTINO, wait_until="networkidle", timeout=60_000)
        self.p.fill('input[name="Log"]', ECONET_LOGIN)
        self.p.fill('input[name="Sen"]', ECONET_SENHA)
        # o botao so habilita apos o callback do reCAPTCHA; o servidor nao valida,
        # entao habilitamos via JS (mesmo caminho que o site usa quando o captcha
        # esta "invisivel").
        self.p.evaluate("() => { const b=document.getElementById('login_ver'); if (b) b.disabled=false; }")
        self.p.click("#login_ver")
        self.p.wait_for_selector("text=FERRAMENTA", timeout=TIMEOUT)
        self.p.wait_for_load_state("networkidle")
        time.sleep(1)

    # -- utilitarios ---------------------------------------------------------
    def _home(self):
        """Volta para a tela inicial da ferramenta."""
        # tentar clicar no titulo (rapido); se nao voltar, recarregar a pagina
        try:
            self.p.click("h1:has-text('ecoclass')", timeout=3_000)
            self.p.wait_for_selector("#input-v-6", timeout=4_000)
            time.sleep(0.3)
            return
        except PWTimeout:
            pass
        self.p.goto(PAG_DESTINO, wait_until="networkidle")
        self.p.wait_for_selector("#input-v-6", timeout=TIMEOUT)
        time.sleep(0.4)

    def _set_categoria(self, tipo):
        """tipo = 'NCM' -> Produtos e Bens Materiais ; 'NBS' -> Servicos"""
        alvo = "Produtos e Bens Materiais" if tipo == "NCM" else "Serviços"
        combo = self.p.locator(".v-select").first
        if (combo.inner_text() or "").strip().startswith(alvo.split()[0]):
            return
        combo.click()
        time.sleep(0.3)
        self.p.click(f".v-list-item:has-text('{alvo}')", timeout=5_000)
        time.sleep(0.3)

    def _tab_text(self, nome):
        """Clica na aba (CST / cClassTrib / cCredPres / Alíquota) e devolve o texto."""
        try:
            self.p.click(f".tab-header-item:has-text('{nome}')", timeout=6_000)
            time.sleep(0.6)
            # expande qualquer painel recolhido (chevron)
            for ch in self.p.locator(".tab-content .mdi-chevron-down").all():
                try:
                    ch.click(timeout=1_000)
                    time.sleep(0.2)
                except Exception:
                    pass
            return _norm(self.p.locator(".tab-content").inner_text())
        except PWTimeout:
            return ""

    # -- consulta de um codigo --------------------------------------------------
    def consulta(self, tipo, codigo):
        self._home()
        self._set_categoria(tipo)
        box = self.p.locator("#input-v-6")
        box.click()
        box.fill("")
        box.type(codigo, delay=40)
        time.sleep(0.4)
        self.p.click("i.mdi-magnify")
        # espera as abas de resultado aparecerem
        try:
            self.p.wait_for_selector("text=Resultado da Pesquisa", timeout=TIMEOUT)
        except PWTimeout:
            return [{"indicador": "SEM RESPOSTA", "tela": _norm(self.p.locator("body").inner_text())}]
        time.sleep(1.2)

        chave = "NCM" if tipo == "NCM" else "NBS"
        linhas_res = self.p.locator("table tbody tr")
        n = linhas_res.count()
        if n == 0:
            corpo = _norm(self.p.locator("body").inner_text())
            return [{"indicador": "Nao localizado (CST 000 - Tributacao Integral)", "tela": corpo}]

        resultados = []
        for i in range(n):
            # a lista de resultados e recarregada a cada volta -> re-seleciona
            linhas_res = self.p.locator("table tbody tr")
            row = linhas_res.nth(i)
            row_txt = row.inner_text()
            cols = [c.strip() for c in re.split(r"[\t\n]+", row_txt) if c.strip()]
            cod_ret = cols[0] if cols else ""
            desc_ret = cols[1] if len(cols) > 1 else ""
            indicador = ""
            for tag in ("Redução de Alíquota", "Tributação Integral", "Suspensão",
                        "Diferimento com Redução de Alíquota", "Diferimento",
                        "Monofasia", "Imunidade", "Isenção", "Alíquota Zero",
                        "Crédito Presumido"):
                if tag.lower() in row_txt.lower():
                    indicador = tag
                    break
            if not indicador and len(cols) > 2:
                indicador = cols[-1]
            row.click()
            self.p.wait_for_selector(".tab-header-item", timeout=TIMEOUT)
            time.sleep(1.0)

            cab = _norm(self.p.locator(".tab-content, main").first.inner_text())
            m = re.search(r"Base legal\s*(.+?)(?:Atenção|CST\s*\d|$)", cab, re.S)
            base_legal = _norm(m.group(1)) if m else ""

            reg = {
                "codigo_retornado": cod_ret,
                "descricao_resultado": desc_ret,
                "indicador": indicador,
                "base_legal": base_legal,
                "cst_txt": self._tab_text("CST"),
                "cclasstrib_txt": self._tab_text("cClassTrib"),
                "ccredpres_txt": self._tab_text("cCredPres"),
                "aliquota_txt": self._tab_text("Alíquota"),
            }
            reg["descricao_lc"] = _first_desc(cab) or desc_ret
            resultados.append(reg)

            # volta para a lista de resultados
            try:
                self.p.click(".mdi-chevron-left", timeout=5_000)
                self.p.wait_for_selector("text=Resultado da Pesquisa", timeout=TIMEOUT)
                time.sleep(0.8)
            except PWTimeout:
                self.consulta_reset = True
                self._home()
                self._set_categoria(tipo)
                self.p.locator("#input-v-6").fill(codigo)
                self.p.click("i.mdi-magnify")
                self.p.wait_for_selector("text=Resultado da Pesquisa", timeout=TIMEOUT)
                time.sleep(1.0)
        return resultados


def _norm(t):
    return re.sub(r"[ \t]*\n[ \t]*", "\n", re.sub(r"[ \t]{2,}", " ", (t or "").strip()))


_IGNORAR_DESC = (
    "esta ferramenta apresenta", "na hipótese de inaplicabilidade",
    "atenção:", "a alíquota exibida", "resultado da pesquisa",
)


def _first_desc(txt):
    for ln in (txt or "").splitlines():
        ln = ln.strip()
        low = ln.lower()
        if len(ln) > 40 and not low.startswith(("base legal", "atenção", "resultado")) \
                and not any(s in low for s in _IGNORAR_DESC):
            return ln
    return ""


def _extrai(codigo_txt, pat):
    m = re.search(pat, codigo_txt or "")
    return m.group(1).strip() if m else ""


# --------------------------------------------------------------------------- #
# Gravacao
# --------------------------------------------------------------------------- #
COLS = [
    "Tipo", "Codigo consultado", "Abas de origem", "Qtd. itens na planilha",
    "Descricao (planilha)", "Codigo retornado", "Indicador", "Descricao (LC 214/2025)",
    "CST (codigo)", "cClassTrib (codigo)", "Base legal",
    "Aba CST (texto)", "Aba cClassTrib (texto)", "Aba cCredPres (texto)", "Aba Aliquota (texto)",
]


def grava(linhas, caminho):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EcoClass_RPA"
    ws.append(COLS)
    for c in range(1, len(COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for l in linhas:
        ws.append([l.get(c, "") for c in COLS])
    larg = {"Descricao (planilha)": 40, "Descricao (LC 214/2025)": 55, "Base legal": 40,
            "Aba CST (texto)": 40, "Aba cClassTrib (texto)": 60,
            "Aba cCredPres (texto)": 40, "Aba Aliquota (texto)": 40, "Abas de origem": 26}
    for i, c in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = larg.get(c, 16)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{ws.max_row}"
    wb.save(caminho)
    print(f"\nArquivo gravado: {caminho}  ({len(linhas)} linhas)")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    if not ECONET_LOGIN or not ECONET_SENHA:
        sys.exit("Defina ECONET_LOGIN e ECONET_SENHA (variavel de ambiente ou .env). Veja .env.example.")

    filtro = [a.strip() for a in sys.argv[1:]]
    headless = HEADLESS_PADRAO and not filtro   # se pediu codigos especificos, mostra a janela

    codigos = extrair_codigos(WORKBOOK)
    if filtro:
        codigos = [c for c in codigos if c["codigo"] in filtro]
        if not codigos:
            sys.exit(f"Nenhum dos codigos {filtro} foi encontrado na planilha.")

    linhas = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless, channel="chrome")
        ctx = browser.new_context(viewport={"width": 1500, "height": 1000},
                                  locale="pt-BR", timezone_id="America/Recife")
        rpa = EcoClassRPA(ctx.new_page())
        print("Fazendo login ...")
        rpa.login()
        print("Logado. Iniciando consultas.\n")

        for n, cod in enumerate(codigos, 1):
            print(f"[{n:>3}/{len(codigos)}] {cod['tipo']} {cod['codigo']} ...", end=" ", flush=True)
            base = {
                "Tipo": cod["tipo"], "Codigo consultado": cod["codigo"],
                "Abas de origem": " | ".join(f"{a}({q})" for a, q in cod["fontes"].items()),
                "Qtd. itens na planilha": cod["total_itens"],
                "Descricao (planilha)": " /// ".join(cod["amostras"]),
            }
            try:
                res = rpa.consulta(cod["tipo"], cod["codigo"])
            except Exception as exc:  # noqa: BLE001
                res = [{"indicador": f"ERRO: {exc}", "tela": ""}]
            for r in res:
                linha = dict(base)
                linha["Codigo retornado"] = r.get("codigo_retornado", "")
                linha["Indicador"] = r.get("indicador", "")
                linha["Descricao (LC 214/2025)"] = r.get("descricao_lc", "")
                linha["Base legal"] = r.get("base_legal", "")
                linha["CST (codigo)"] = _extrai(r.get("cst_txt", ""), r"\b(\d{3})\b")
                linha["cClassTrib (codigo)"] = _extrai(r.get("cclasstrib_txt", ""), r"\b(\d{6})\b")
                if not linha["Indicador"] and linha["CST (codigo)"] == "000":
                    linha["Indicador"] = "Tributacao Integral"
                linha["Aba CST (texto)"] = r.get("cst_txt", "")
                linha["Aba cClassTrib (texto)"] = r.get("cclasstrib_txt", "")
                linha["Aba cCredPres (texto)"] = r.get("ccredpres_txt", "")
                linha["Aba Aliquota (texto)"] = r.get("aliquota_txt", "")
                linhas.append(linha)
            print(res[0].get("indicador", "")[:50])
            time.sleep(PAUSA)
        browser.close()

    grava(linhas, OUT_XLSX)


if __name__ == "__main__":
    main()
