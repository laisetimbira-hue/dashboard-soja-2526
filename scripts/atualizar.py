#!/usr/bin/env python3
"""
SEED BASE — atualização automática do dashboard.

Baixa as duas planilhas do Google Drive, processa e injeta os dados
no index.html. Roda dentro do GitHub Actions.

Configuração via variáveis de ambiente (secrets do repositório):
  ID_ANALISES  → ID do arquivo ANALISES_DE_SEMENTES_DE_VERAO_25-26.xlsx
  ID_BENEF     → ID do arquivo LOTES_BENEFICIADOS_VERAO_25-26.xlsx
"""

import os
import re
import io
import json
import sys
from datetime import datetime, timezone, timedelta
from collections import defaultdict

import requests
import pandas as pd
import openpyxl

# ---------------------------------------------------------------- config

HTML_PATH = "index.html"

ID_ANALISES = os.environ.get("ID_ANALISES", "").strip()
ID_BENEF    = os.environ.get("ID_BENEF", "").strip()

# Fuso de Brasília — o GitHub Actions roda em UTC
HOJE = datetime.now(timezone(timedelta(hours=-3))).strftime("%d/%m/%Y")

APROVADOS = {"A", "B", "BV"}

# Nomes usados por terceiros que não batem com o código oficial do CV_META
BENEF_ALIASES = {
    "VALENTE": "6968RSF RR",
    "C 2605 E": "C2605 E",
}


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------- download

def extrair_id(valor):
    """Aceita o ID puro ou uma URL completa colada no secret."""
    v = (valor or "").strip().strip('"').strip("'")
    if not v:
        return ""
    # .../spreadsheets/d/ID/edit  ou  .../file/d/ID/view
    m = re.search(r"/d/([a-zA-Z0-9_-]{20,})", v)
    if m:
        return m.group(1)
    # ...?id=ID  ou  ...&id=ID
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{20,})", v)
    if m:
        return m.group(1)
    return v


def baixar(file_id, nome, obrigatorio=True):
    """Baixa um arquivo do Drive. Funciona para .xlsx e para Sheets nativo.

    Se obrigatorio=False e o ID não estiver configurado, devolve None em vez
    de abortar — o chamador decide o que fazer.
    """
    file_id = extrair_id(file_id)

    if not file_id:
        if not obrigatorio:
            log(f"  {nome}: ID não configurado — etapa será pulada")
            return None
        raise SystemExit(f"ERRO: ID do arquivo '{nome}' não configurado nos secrets.")

    log(f"  {nome}: id={file_id[:8]}...{file_id[-4:]} ({len(file_id)} caracteres)")

    tentativas = [
        ("Sheets nativo",
         f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"),
        ("Drive usercontent",
         f"https://drive.usercontent.google.com/download?id={file_id}&export=download"),
        ("Drive uc (legado)",
         f"https://drive.google.com/uc?export=download&id={file_id}"),
    ]

    sessao = requests.Session()
    sessao.headers.update({"User-Agent": "Mozilla/5.0 (compatible; SeedBaseBot/1.0)"})

    for rotulo, url in tentativas:
        try:
            r = sessao.get(url, timeout=120, allow_redirects=True)
        except Exception as e:
            log(f"    [{rotulo}] falhou na conexão: {type(e).__name__}: {e}")
            continue

        ctype = r.headers.get("content-type", "?").split(";")[0]
        log(f"    [{rotulo}] HTTP {r.status_code} · {ctype} · {len(r.content):,} bytes")

        if r.status_code != 200:
            continue

        if r.content[:2] == b"PK":          # assinatura de xlsx/zip
            log(f"    -> OK via {rotulo}")
            return io.BytesIO(r.content)

        # Arquivo grande: Google devolve uma página pedindo confirmação
        if b"confirm" in r.content[:60000].lower():
            token = re.search(rb'name="confirm"\s+value="([^"]+)"', r.content)
            uuid  = re.search(rb'name="uuid"\s+value="([^"]+)"', r.content)
            if token:
                params = {"id": file_id, "export": "download",
                          "confirm": token.group(1).decode()}
                if uuid:
                    params["uuid"] = uuid.group(1).decode()
                r2 = sessao.get("https://drive.usercontent.google.com/download",
                                params=params, timeout=180)
                if r2.status_code == 200 and r2.content[:2] == b"PK":
                    log(f"    -> OK via {rotulo} (após confirmação)")
                    return io.BytesIO(r2.content)

        # Diagnóstico: mostrar o começo do que veio
        amostra = r.content[:300].decode("utf-8", errors="replace").replace("\n", " ")
        log(f"    conteúdo recebido (início): {amostra[:200]}")

    raise SystemExit(
        f"\nERRO: não consegui baixar '{nome}' (id={file_id}).\n"
        f"Verifique, nesta ordem:\n"
        f"  1. O arquivo está compartilhado como 'Qualquer pessoa com o link' "
        f"com permissão de Leitor?\n"
        f"  2. O ID no secret corresponde a ESTE arquivo?\n"
        f"  3. O arquivo está na sua conta pessoal ou em um Drive corporativo "
        f"com restrição de domínio? Contas corporativas costumam bloquear "
        f"acesso externo mesmo com o link ativo.\n"
        f"Os códigos HTTP acima indicam o motivo: 404 = ID errado, "
        f"403 = sem permissão pública, 200 com HTML = página de login."
    )


# ---------------------------------------------------------------- helpers

def nz(v):
    if isinstance(v, str):
        return None if v.strip() == "" else v
    return None if pd.isna(v) else v


def r1(v):
    if v is None:
        return None
    try:
        return round(float(str(v).strip()), 1)
    except (TypeError, ValueError):
        return None


def to_float(v):
    if v is None:
        return None
    if isinstance(v, str) and v.startswith("="):
        return None  # fórmula não avaliada
    try:
        return float(str(v).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- classificação

def classificar(germ, vigor, germ_of, vigor_of, tz_rank):
    def gb(g):
        if g is None:
            return None
        return "A" if g >= 95 else "B" if g >= 85 else "C" if g >= 80 else "D"

    def vb(v):
        if v is None:
            return None
        return "A" if v >= 95 else "B" if v >= 80 else "BV" if v >= 70 else "D"

    rg, rv = gb(germ), vb(vigor)

    if germ_of is not None and vigor_of is not None:
        if vigor_of < 80 or germ_of < 80:
            final = "D"
        elif germ_of >= 95:
            final = "A"
        elif germ_of >= 90:
            final = "B"
        else:
            final = "C"
    elif vigor_of is not None:
        final = "A" if vigor_of >= 95 else "B" if vigor_of >= 80 else "D"
    else:
        if rg is None and rv is None:
            final = tz_rank if tz_rank else "X"
        elif rg == "D":
            final = "D"
        elif rg == "C":
            final = "D" if rv in ("BV", "D") else "C"
        elif rg in ("B", "A"):
            if rv == "D":
                final = "D"
            elif rv == "BV":
                final = "D" if tz_rank == "C" else "BV"
            elif rv == "B":
                final = "C" if tz_rank == "C" else "B"
            else:
                final = {"A": "B", "B": "C"}[rg] if tz_rank == "C" else rg
        else:
            final = "X"

    rotulos = {
        "A": "Ótimo", "B": "Bom", "C": "Condicional",
        "BV": "Condicional Vigor", "D": "Ruim", "X": "Pendente",
    }
    return rg, rv, final, rotulos[final]


# ---------------------------------------------------------------- ANALISES

def parse_analises(buf):
    df = pd.read_excel(buf, sheet_name="SOJA 25-26")
    df = df[df["LOTE"].notna()].reset_index(drop=True)

    lots = []
    for _, row in df.iterrows():
        germ     = r1(nz(row["GER. ROLO DE PAPEL %"]))
        vigor    = r1(nz(row["ROLO DE PAPEL E.A. %"]))
        germ_of  = r1(nz(row["GERM OFICIAL %"]))
        vigor_of = r1(nz(row["VIGOR OFICIAL %"]))
        tz_rank  = nz(row["RAKING TZ GERM"])
        rg, rv, ranking, status = classificar(germ, vigor, germ_of, vigor_of, tz_rank)

        kg_raw = nz(row["QUANT.    KG"])
        umid   = nz(row["UMIDADE %"])

        lots.append({
            "cultivar": str(nz(row["CULTIVAR"])).strip() if nz(row["CULTIVAR"]) else None,
            "lote": nz(row["LOTE"]),
            "peneira": nz(row["PENEIRA "]),
            "categoria": nz(row["CATEGORIA "]),
            "ubs": nz(row["UBS"]),
            "umidade": round(float(umid), 2) if umid is not None else None,
            "tz_viab": r1(nz(row["TZ VIABILDIADE (1-5) "])),
            "tz_vigor": r1(nz(row["TZ VIGOR (1-3)"])),
            "rank_tz": tz_rank,
            "germ": germ, "rank_germ": rg,
            "vigor": vigor, "rank_vigor": rv,
            "germ_oficial": germ_of, "vigor_oficial": vigor_of,
            "modo_classificacao": (
                "ambos" if (germ_of and vigor_of) else ("vigor" if vigor_of else "nenhum")
            ),
            "ranking": ranking, "status": status,
            "bags": int(row["QUANT BB"]) if pd.notna(row["QUANT BB"]) else None,
            "kg": int(kg_raw) if kg_raw and str(kg_raw).strip() not in ("", "None") else None,
            "pms_etq": r1(nz(row["PMS (g) ETIQUETA"])),
            "peso_bag": r1(nz(row["PESO BAG"])),
            "sementes_m": r1(nz(row["N° SEMENTES (MILHÕES)"])),
        })

    return lots


def montar_summary(lots):
    df = pd.DataFrame(lots)
    df["germ_eff"]  = df["germ_oficial"].combine_first(df["germ"])
    df["vigor_eff"] = df["vigor_oficial"].combine_first(df["vigor"])

    summary = []
    for cv, grp in df.groupby("cultivar"):
        ap = grp[grp["ranking"].isin(APROVADOS)]
        c = grp["ranking"].value_counts().to_dict()
        summary.append({
            "cultivar": cv,
            "total_kg": int(grp["kg"].sum()),
            "lotes": int(len(grp)),
            "germ_media": r1(ap["germ_eff"].dropna().mean()),
            "vigor_media": r1(ap["vigor_eff"].dropna().mean()),
            "tz_viab_media": r1(grp["tz_viab"].dropna().mean()),
            "tz_vigor_media": r1(grp["tz_vigor"].dropna().mean()),
            "rank_A": int(c.get("A", 0)), "rank_B": int(c.get("B", 0)),
            "rank_BV": int(c.get("BV", 0)), "rank_C": int(c.get("C", 0)),
            "rank_D": int(c.get("D", 0)), "rank_X": int(c.get("X", 0)),
        })
    summary.sort(key=lambda s: -s["total_kg"])

    pms_cv = []
    for cv, grp in df.groupby("cultivar"):
        e = {"cultivar": cv}
        total_bags = 0
        for pen, suf in [("P1", "p1"), ("P2", "p2")]:
            sub = grp[grp["peneira"] == pen]
            bags = int(sub["bags"].sum()) if len(sub) else 0
            total_bags += bags
            ap = sub[sub["ranking"].isin(APROVADOS)]
            c = sub["ranking"].value_counts().to_dict()
            e[f"bags_{suf}"]   = bags
            e[f"pms_{suf}"]    = r1(ap["pms_etq"].dropna().mean())
            e[f"aprov_{suf}"]  = int(c.get("A", 0) + c.get("B", 0) + c.get("BV", 0))
            e[f"reprov_{suf}"] = int(c.get("C", 0) + c.get("D", 0))
            e[f"pend_{suf}"]   = int(c.get("X", 0))
        e["bags_total"] = total_bags
        pms_cv.append(e)
    pms_cv.sort(key=lambda p: -p["bags_total"])

    return summary, pms_cv


def parse_comprados(buf):
    wb = openpyxl.load_workbook(buf)
    ws = wb["LOTES COMPRADOS"]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(header) if h}

    def cell(r, nome):
        return ws.cell(r, col[nome]).value if nome in col else None

    def txt(r, nome):
        v = cell(r, nome)
        return str(v).strip() if v else None

    out = []
    for r in range(2, ws.max_row + 1):
        cultivar = ws.cell(r, 1).value
        if not cultivar:
            continue

        bags     = to_float(cell(r, "QUANT BB"))
        peso_bag = to_float(cell(r, "PESO BAG"))
        kg_raw   = to_float(cell(r, "QUANT.    KG"))
        kg = int(bags * peso_bag) if bags and peso_bag else (int(kg_raw) if kg_raw else None)

        # PMS pode vir como fórmula (=PESO BAG/5) — nesse caso calcula na mão
        pms_raw = cell(r, "PMS (g) ETIQUETA")
        if pms_raw is not None and not (isinstance(pms_raw, str) and pms_raw.startswith("=")):
            pms_etq = round(float(pms_raw), 1)
        elif peso_bag:
            pms_etq = round(peso_bag / 5, 1)
        else:
            pms_etq = None

        out.append({
            "cultivar": str(cultivar).strip(),
            "lote": txt(r, "LOTE SCV"),
            "lote_origem": txt(r, "LOTE DE ORIGEM"),
            "empresa": txt(r, "EMPRESA"),
            "germ_oficial": to_float(cell(r, "GERM OFICIAL %")),
            "vigor_oficial": to_float(cell(r, "VIGOR OFICIAL %")),
            "bags": int(bags) if bags else None,
            "peso_bag": peso_bag,
            "kg": kg,
            "categoria_origem": str(ws.cell(r, 2).value).strip() if ws.cell(r, 2).value else None,
            "categoria_reembalado": str(ws.cell(r, 3).value).strip() if ws.cell(r, 3).value else None,
            "peneira": txt(r, "PENEIRA "),
            "pms_etq": pms_etq,
            "status": txt(r, "STATUS"),
        })
    return out


# ---------------------------------------------------------------- BENEFICIADOS

# Cada aba tem layout próprio. (linha_inicial, col_cultivar, col_categoria,
# col_peneira, col_pms, col_bags, pms_vem_do_peso)
LAYOUT_BENEF = {
    "SOJA 25-26": ("SCV",      3, 2, 4, 6, 7, 9, False),
    "CAMBAÍ":     ("CAMBAÍ",   3, 1, 3, 5, 9, 8, True),
    "SIMÃO":      ("SIMÃO",    4, 1, 3, 5, 6, 8, False),
    "GIOVELLI":   ("GIOVELLI", 3, 1, 3, 5, 6, 8, False),
    "CERENNA":    ("CERENNA",  4, 1, 3, 5, 6, 8, False),
}

CV_INVALIDOS = {
    "CULTIVAR", "DATA", "SAFRA", "CAT.", "N° BB", "QUANTIDADE",
    "LOTE", "PEN.", "PMS", "EMBALAGEM", "PESO", "LOCALIZAÇÃO",
    "STATUS", "OBSERVAÇÕES",
}


def parse_benef(buf):
    wb = openpyxl.load_workbook(buf)
    agrupado = defaultdict(lambda: {"bags": 0, "kg": 0.0, "pms_sum": 0.0, "pms_cnt": 0})

    for aba, cfg in LAYOUT_BENEF.items():
        if aba not in wb.sheetnames:
            log(f"  aviso: aba '{aba}' não encontrada — ignorando")
            continue

        ubs, linha0, c_cv, c_cat, c_pen, c_pms, c_bags, pms_do_peso = cfg
        ws = wb[aba]

        for r in range(linha0, ws.max_row + 1):
            cv = ws.cell(r, c_cv).value
            if not cv:
                continue
            cv = str(cv).strip()
            if cv in CV_INVALIDOS or "PARA SEMENTES" in cv or "SEMENTES COM VIGOR" in cv:
                continue

            bags = to_float(ws.cell(r, c_bags).value)
            if not bags or bags <= 0:
                continue
            bags = int(bags)

            bruto = to_float(ws.cell(r, c_pms).value)
            if not bruto:
                continue
            pms = bruto / 5 if pms_do_peso else bruto  # coluna traz peso do bag

            pen_raw = ws.cell(r, c_pen).value
            pen = "P1" if pen_raw and "P1" in str(pen_raw).upper() else "P2"
            cat_raw = ws.cell(r, c_cat).value
            cat = str(cat_raw).strip() if cat_raw else "C2"

            k = agrupado[(cv, ubs, cat, pen)]
            k["bags"]    += bags
            k["kg"]      += bags * pms * 5
            k["pms_sum"] += pms * bags
            k["pms_cnt"] += bags

    # consolidar por (cultivar, ubs)
    por_cv_ubs = defaultdict(lambda: defaultdict(
        lambda: {"bags": 0, "kg": 0.0, "pms_sum": 0.0, "pms_cnt": 0}
    ))
    for (cv, ubs, cat, pen), v in agrupado.items():
        d = por_cv_ubs[(cv, ubs)][(cat, pen)]
        for campo in ("bags", "kg", "pms_sum", "pms_cnt"):
            d[campo] += v[campo]

    benef = []
    for (cv, ubs), detalhes in por_cv_ubs.items():
        detail = []
        p1_s = p1_n = p2_s = p2_n = 0
        for (cat, pen), v in sorted(detalhes.items()):
            if v["bags"] == 0:
                continue
            detail.append({
                "categoria": cat, "peneira": pen,
                "bags": v["bags"], "kg": round(v["kg"]),
            })
            if pen == "P1":
                p1_s += v["pms_sum"]; p1_n += v["pms_cnt"]
            else:
                p2_s += v["pms_sum"]; p2_n += v["pms_cnt"]

        total = sum(d["bags"] for d in detail)
        if total == 0:
            continue

        benef.append({
            "cultivar": cv, "ubs": ubs,
            "bags_total": total,
            "kg_total": sum(d["kg"] for d in detail),
            "bags_p1": sum(d["bags"] for d in detail if d["peneira"] == "P1"),
            "bags_p2": sum(d["bags"] for d in detail if d["peneira"] == "P2"),
            "pms_todos_p1": round(p1_s / p1_n, 1) if p1_n else None,
            "pms_todos_p2": round(p2_s / p2_n, 1) if p2_n else None,
            "detail": detail,
        })

    benef.sort(key=lambda x: -x["kg_total"])
    return benef


# ---------------------------------------------------------------- injeção

def js(dados):
    return json.dumps(dados, ensure_ascii=False, separators=(",", ":"))


def substituir_const(html, nome, dados):
    novo, n = re.subn(
        rf"const {nome} = \[.*?\];",
        f"const {nome} = {js(dados)};",
        html, count=1, flags=re.DOTALL,
    )
    if n == 0:
        raise SystemExit(f"ERRO: não encontrei 'const {nome}' no {HTML_PATH}.")
    log(f"  {nome}: {len(dados)} itens")
    return novo


def atualizar_html(lots, summary, pms_cv, comprados, benef):
    with open(HTML_PATH, encoding="utf-8") as f:
        html = f.read()

    # guarda o SUMMARY atual como SUMMARY_PREV (alimenta o botão "Δ anterior")
    atual = re.search(r"const SUMMARY = (\[.*?\]);", html, re.DOTALL)
    if atual and "const SUMMARY_PREV" in html:
        html, _ = re.subn(
            r"const SUMMARY_PREV = \[.*?\];",
            f"const SUMMARY_PREV = {atual.group(1)};",
            html, count=1, flags=re.DOTALL,
        )

    html = substituir_const(html, "LOTS", lots)
    html = substituir_const(html, "SUMMARY", summary)
    html = substituir_const(html, "PMS_CV", pms_cv)
    html = substituir_const(html, "LOTES_COMPRADOS", comprados)
    if benef is None:
        log("  BENEF: preservado (planilha não informada)")
    else:
        html = substituir_const(html, "BENEF", benef)

    # data no header e no rodapé
    html, _ = re.subn(
        r'(<div class="val" id="h-data"[^>]*>)[^<]*(</div>)',
        rf"\g<1>{HOJE}\2", html,
    )
    html, _ = re.subn(r"Atualizado em [0-9/]+", f"Atualizado em {HOJE}", html)

    # histórico de aprovação
    total = len(lots)
    aprov = sum(1 for l in lots if l["ranking"] in APROVADOS)
    cont = defaultdict(int)
    for l in lots:
        cont[l["ranking"]] += 1

    ponto = {
        "data": HOJE,
        "rank_A": cont["A"], "rank_B": cont["B"], "rank_BV": cont["BV"],
        "rank_C": cont["C"], "rank_D": cont["D"], "rank_X": cont["X"],
        "total": total,
        "pctAprov": round(aprov / total * 100, 1) if total else 0,
    }

    m = re.search(r"HISTORICO_APROVACAO = (\[.*?\]);", html, re.DOTALL)
    if m:
        hist = json.loads(m.group(1))
        chaves = ["rank_A", "rank_B", "rank_BV", "rank_C", "rank_D", "rank_X", "total"]

        def igual(p):
            return all(p.get(k) == ponto[k] for k in chaves)

        if any(igual(p) for p in hist):
            for p in hist:
                if igual(p):
                    p["data"] = HOJE
            log("  histórico: composição igual, data atualizada")
        else:
            hist.append(ponto)
            log(f"  histórico: novo ponto (#{len(hist)})")

        html, _ = re.subn(
            r"HISTORICO_APROVACAO = \[.*?\];",
            f"HISTORICO_APROVACAO = {js(hist)};",
            html, count=1, flags=re.DOTALL,
        )

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    return total, aprov, dict(cont)


# ---------------------------------------------------------------- main

def main():
    log("Baixando planilhas do Drive...")
    buf_analises = baixar(ID_ANALISES, "ANALISES")
    buf_benef    = baixar(ID_BENEF, "BENEFICIADOS", obrigatorio=False)

    log("\nProcessando ANALISES...")
    lots = parse_analises(buf_analises)
    summary, pms_cv = montar_summary(lots)

    buf_analises.seek(0)
    comprados = parse_comprados(buf_analises)
    log(f"  {len(lots)} lotes · {len(comprados)} comprados")

    if buf_benef is None:
        benef = None
        log("\nBENEFICIADOS: pulado — o BENEF atual do index.html será preservado")
    else:
        log("\nProcessando BENEFICIADOS...")
        benef = parse_benef(buf_benef)
        log(f"  {len(benef)} entradas · {sum(b['bags_total'] for b in benef):,} bags")

    log(f"\nInjetando no {HTML_PATH}...")
    total, aprov, cont = atualizar_html(lots, summary, pms_cv, comprados, benef)

    pct = round(aprov / total * 100, 1) if total else 0
    log(f"\n{'='*46}")
    log(f"  {HOJE}")
    log(f"  {total} lotes · {aprov} aprovados ({pct}%)")
    log(f"  {cont}")
    log(f"{'='*46}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        log(f"\nFALHOU: {type(e).__name__}: {e}")
        sys.exit(1)
