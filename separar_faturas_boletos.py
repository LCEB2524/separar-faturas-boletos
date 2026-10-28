"""
Junta FATURA + BOLETOS + NOTAS FISCAIS de cada cliente em um único PDF.

Entrada (na pasta "conth", um nível acima deste script):
  1. Faturamento agrupado .... Emiss*Faturamento*.pdf   (uma fatura por cliente, campo CLIENTE)
  2. Boletos Sicredi ......... boleto*.pdf              (campo Pagador e "Ref. Notas Fiscais")
  3. Notas fiscais (DANFE) ... DANFE*.pdf               (pode haver mais de um arquivo)

Etapas:
  1. Lê os 3 arquivos.
  2. Cruza tudo e confere:
       Fatura x Boleto : cliente, valor da fatura = soma dos boletos, nº duplicata = nº documento,
                         NFs da fatura x "Ref. Notas Fiscais" do boleto
       Fatura x NF     : cada NF da fatura existe no arquivo de NFs, destinatário = cliente,
                         valor da NF = valor na fatura, soma das NFs = VALOR PEDIDOS
       Sobras          : boletos sem fatura e NFs que não estão em nenhuma fatura
  3. Mostra a conferência, salva em Excel (saida\\Conferencia) e só gera os PDFs depois do
     seu "s" (ou direto, com o parâmetro --sem-perguntar).

Saída: saida\\<CLIENTE>.pdf  = fatura + boletos + NFs (na ordem da fatura)
       Cliente sem boleto: "<CLIENTE> - SEM BOLETO.pdf"

Uso:
  python separar_faturas_boletos.py
  python separar_faturas_boletos.py --sem-perguntar
"""

import fnmatch
import logging
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pypdf

logging.disable(logging.WARNING)  # PDFs do banco geram muitos avisos inofensivos

# ------------------------------------------------------------------ configuração
PASTA = Path(__file__).resolve().parent.parent      # pasta "conth"
PADRAO_FATURAMENTO = "emiss*faturamento*.pdf"
PADRAO_BOLETOS = "boleto*.pdf"
PADRAO_NFS = "danfe*.pdf"
PASTA_SAIDA = "saida"
TOLERANCIA = 0.009                                    # diferença de centavos aceita


# ------------------------------------------------------------------ utilidades
def normalizar(txt):
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt.upper())
    return re.sub(r"\s+", " ", txt).strip()


def valor_br(v):
    return float(v.replace(".", "").replace(",", ".")) if v else 0.0


def fmt(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def mesmo_nome(a, b):
    """Iguais, ou um é o começo do outro (nomes cortados em boleto/NF)."""
    a, b = normalizar(a), normalizar(b)
    if not a or not b:
        return False
    curto, longo = sorted((a, b), key=len)
    return a == b or (len(curto) >= 15 and longo.startswith(curto))


def achar_arquivos(padrao, unico=True):
    achados = sorted(p for p in PASTA.iterdir()
                     if p.is_file() and fnmatch.fnmatch(p.name.lower(), padrao))
    if not achados:
        sys.exit(f"ERRO: não encontrei arquivo '{padrao}' em {PASTA}")
    if unico and len(achados) > 1:
        sys.exit(f"ERRO: mais de um arquivo '{padrao}' em {PASTA}: {[p.name for p in achados]}")
    return achados


def nome_arquivo(nome):
    return re.sub(r'[\\/:*?"<>|]', "", nome).strip().rstrip(".")


RE_NOTA_FATURA = re.compile(r"\d{8}\s+(\d{5,10})\s+([\d.]+,\d{2})N")


# ------------------------------------------------------------------ leitura
def ler_faturas(leitor):
    faturas = []
    for i, pag in enumerate(leitor.pages):
        t = pag.extract_text() or ""
        m = re.search(r"CLIENTE\s*:\s*\d+-(.+?)\s+(?:TELEFONE|FONE)", t)
        notas = {str(int(n)): valor_br(v) for n, v in RE_NOTA_FATURA.findall(t)}
        novo = m and not (faturas and faturas[-1]["nome"] == m.group(1).strip()
                          and not faturas[-1]["fechada"])
        if novo:
            faturas.append({"nome": m.group(1).strip(), "paginas": [], "notas": {},
                            "dups": [], "valor": 0.0, "valor_pedidos": 0.0,
                            "credito": 0.0, "acrescimo": 0.0,
                            "qtde": None, "fechada": False})
        if not faturas:
            continue
        f = faturas[-1]
        f["paginas"].append(i)
        f["notas"].update(notas)
        f["dups"] += re.findall(r"(\d{9}[A-Z])\s", t)
        vf = re.search(r"VALOR FATURA:\s*([\d.,]+)", t)
        vp = re.search(r"VALOR PEDIDOS:\s*([\d.,]+)", t)
        q = re.search(r"QTDE\. NOTAS:(\d+)", t)
        cr = re.search(r"CR[EÉ]DITO\s*\[-\]:\s*([\d.,]+)", t)
        ac = re.search(r"FRETE \+ OUTROS\s*\[\+\]:\s*([\d.,]+)", t)
        if cr:
            f["credito"] = valor_br(cr.group(1))
        if ac:
            f["acrescimo"] = valor_br(ac.group(1))
        if vf:
            f["valor"], f["fechada"] = valor_br(vf.group(1)), True
        if vp:
            f["valor_pedidos"] = valor_br(vp.group(1))
        if q:
            f["qtde"] = int(q.group(1))
    return faturas


def ler_boletos(leitor):
    boletos = []
    for i, pag in enumerate(leitor.pages):
        t = pag.extract_text() or ""
        m = re.search(r"Pagador:\s*(.+)", t)
        r = re.search(r"Ref\. Notas Fiscais:\s*([\d,\s]+?)\s*Cobrar", t)
        nfs = [str(int(x)) for x in re.sub(r"\s", "", r.group(1)).split(",") if x] if r else []
        doc = re.search(r"(\d{9}[A-Z]) R\$", t)
        v = re.search(r"Valor do Documento\s*\n([\d.,]+)", t)
        boletos.append({"pagina": i, "nome": m.group(1).strip() if m else "", "nfs": nfs,
                        "doc": doc.group(1) if doc else "",
                        "valor": valor_br(v.group(1)) if v else 0.0})
    return boletos


def ler_nfs(arquivos):
    """Devolve {numero: {...}} juntando todos os PDFs de DANFE (NF de várias folhas fica junta)."""
    nfs, leitores = {}, []
    for arq in arquivos:
        leitor = pypdf.PdfReader(str(arq))
        leitores.append(leitor)
        anterior = None
        for i, pag in enumerate(leitor.pages):
            t = pag.extract_text() or ""
            n = re.search(r"Nº\s*([\d.]+)", t)
            if not n:
                if anterior:                      # página sem número: continuação
                    nfs[anterior]["paginas"].append((leitor, i))
                continue
            num = str(int(n.group(1).replace(".", "")))
            if num in nfs and anterior == num:    # Folha 2, 3... da mesma NF
                nfs[num]["paginas"].append((leitor, i))
                continue
            dest = re.search(r"NOME / RAZÃO SOCIAL\s*\n\s*(?:\d+\.)?(.+)", t)
            if not dest:
                dest = re.search(r"DESTINATÁRIO:\s*(?:\d+\.)?(.+?)\s+-\s", t)
            v = re.search(r"VALOR TOTAL DA NOTA\s*\n\s*([\d.,]+)", t) or \
                re.search(r"VALOR TOTAL: R\$\s*([\d.,]+)", t)
            emi = re.search(r"EMISSÃO:\s*(\d{2}/\d{2}/\d{4})", t)
            if num in nfs:                         # mesma NF repetida em outro arquivo
                anterior = num
                continue
            nfs[num] = {"numero": num, "dest": dest.group(1).strip() if dest else "",
                        "valor": valor_br(v.group(1)) if v else 0.0,
                        "emissao": emi.group(1) if emi else "",
                        "arquivo": arq.name, "paginas": [(leitor, i)]}
            anterior = num
    return nfs, leitores


# ------------------------------------------------------------------ cruzamento
def cruzar(faturas, boletos, nfs):
    res = []
    for f in faturas:
        bs, como = [b for b in boletos if normalizar(b["nome"]) == normalizar(f["nome"])], "nome"
        if not bs:
            bs, como = [b for b in boletos if mesmo_nome(b["nome"], f["nome"])], "nome cortado"
        if not bs:
            bs, como = [b for b in boletos if set(b["nfs"]) & set(f["notas"])], "NF"
        if not bs:
            como = "-"

        alertas = []
        # ---- fatura x boleto
        soma_bol = round(sum(b["valor"] for b in bs), 2)
        if bs:
            if abs(soma_bol - f["valor"]) > TOLERANCIA:
                alertas.append(f"BOLETO: fatura {fmt(f['valor'])} x boletos {fmt(soma_bol)}")
            if sorted(f["dups"]) != sorted(b["doc"] for b in bs):
                alertas.append(f"BOLETO: duplicatas {f['dups']} x documentos {[b['doc'] for b in bs]}")
            ref = set().union(*[set(b["nfs"]) for b in bs])
            fora = sorted(set(f["notas"]) - ref)
            extra = sorted(ref - set(f["notas"]))
            if fora:
                alertas.append(f"BOLETO: NFs fora da 'Ref. Notas Fiscais': {', '.join(fora)}")
            if extra:
                alertas.append(f"BOLETO: NFs no boleto que não estão na fatura: {', '.join(extra)}")
        else:
            alertas.append("BOLETO: nenhum boleto encontrado para este cliente")

        # ---- fatura x NF
        achadas, faltando = [], []
        for num, v_fat in f["notas"].items():
            nf = nfs.get(num)
            if not nf:
                faltando.append(num)
                continue
            achadas.append(num)
            nf["usada"] = True
            if abs(nf["valor"] - v_fat) > TOLERANCIA:
                alertas.append(f"NF {num}: valor na NF {fmt(nf['valor'])} x fatura {fmt(v_fat)}")
            if not mesmo_nome(nf["dest"], f["nome"]):
                alertas.append(f"NF {num}: destinatário '{nf['dest']}' diferente do cliente")
        if faltando:
            alertas.append(f"NF: não encontradas no arquivo de NFs: {', '.join(faltando)}")
        # Valor das NFs (já inclui frete/outros) - crédito = valor da fatura
        soma_nf_fat = round(sum(f["notas"].values()), 2)
        esperado = round(soma_nf_fat - f["credito"], 2)
        if abs(esperado - f["valor"]) > TOLERANCIA:
            alertas.append(f"FATURA: soma das NFs {fmt(soma_nf_fat)} - crédito {fmt(f['credito'])} "
                           f"= {fmt(esperado)} x VALOR FATURA {fmt(f['valor'])}")
        if f["qtde"] is not None and f["qtde"] != len(f["notas"]):
            alertas.append(f"FATURA: QTDE. NOTAS {f['qtde']} x {len(f['notas'])} NFs lidas")

        res.append({"fatura": f, "boletos": bs, "como": como, "nfs": achadas,
                    "faltando": faltando, "soma_bol": soma_bol, "alertas": alertas})
    return res


# ------------------------------------------------------------------ conferência
def mostrar_conferencia(res, boletos, nfs, arqs):
    print("=" * 78)
    print(" CONFERÊNCIA")
    for k, v in arqs.items():
        print(f"   {k:12}: {v}")
    print("=" * 78)
    for r in res:
        f = r["fatura"]
        ok = "OK " if not r["alertas"] else "!! "
        print(f" {ok}{f['nome'][:42]:42} fat {fmt(f['valor']):>10} | "
              f"bol {len(r['boletos'])} | NF {len(r['nfs'])}/{len(f['notas'])}")
        for a in r["alertas"]:
            print(f"      - {a}")
    usados = {b["pagina"] for r in res for b in r["boletos"]}
    sobra_bol = [b for b in boletos if b["pagina"] not in usados]
    sobra_nf = [n for n in nfs.values() if not n.get("usada")]
    print("-" * 78)
    if sobra_bol:
        print(f" BOLETOS SEM FATURA ({len(sobra_bol)}):")
        for b in sobra_bol:
            print(f"      - {b['nome']} | {b['doc']} | {fmt(b['valor'])}")
    print(f" NFs do arquivo que não estão em nenhuma fatura: {len(sobra_nf)}")
    tot_ok = sum(1 for r in res if not r["alertas"])
    print("-" * 78)
    print(f" Clientes: {len(res)} | sem divergência: {tot_ok} | com divergência: {len(res) - tot_ok}")
    sem_bol = sum(r['fatura']['valor'] for r in res if not r['boletos'])
    print(f" Total faturas: {fmt(sum(r['fatura']['valor'] for r in res))} "
          f"(sem boleto: {fmt(sem_bol)}) | total boletos usados: "
          f"{fmt(sum(r['soma_bol'] for r in res))} | "
          f"NFs anexadas: {sum(len(r['nfs']) for r in res)} de "
          f"{sum(len(r['fatura']['notas']) for r in res)}")
    print("=" * 78)
    return sobra_bol, sobra_nf


def salvar_excel(res, sobra_bol, sobra_nf, pasta):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
    except ImportError:
        print(" (openpyxl não instalado: conferência em Excel não foi salva)")
        return None
    pasta.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clientes"
    ws.append(["Cliente", "Valor fatura", "Valor boletos", "Qtd boletos", "Cruzamento boleto",
               "NFs na fatura", "NFs anexadas", "NFs não encontradas", "Status", "Divergências"])
    for r in res:
        f = r["fatura"]
        ws.append([f["nome"], f["valor"], r["soma_bol"], len(r["boletos"]), r["como"],
                   len(f["notas"]), len(r["nfs"]), ", ".join(r["faltando"]),
                   "OK" if not r["alertas"] else "DIVERGÊNCIA", " | ".join(r["alertas"])])
        cor = "C6EFCE" if not r["alertas"] else "FFC7CE"
        for c in ws[ws.max_row]:
            c.fill = PatternFill("solid", fgColor=cor)
    ws2 = wb.create_sheet("Sobras")
    ws2.append(["Tipo", "Nome/Destinatário", "Documento/NF", "Valor", "Emissão"])
    for b in sobra_bol:
        ws2.append(["Boleto sem fatura", b["nome"], b["doc"], b["valor"], ""])
    for n in sobra_nf:
        ws2.append(["NF fora das faturas", n["dest"], n["numero"], n["valor"], n["emissao"]])
    for w in (ws, ws2):
        for c in w[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
        w.freeze_panes = "A2"
    for col, larg in zip("ABCDEFGHIJ", (45, 13, 13, 11, 16, 12, 12, 30, 13, 90)):
        ws.column_dimensions[col].width = larg
    for col, larg in zip("ABCDE", (22, 45, 14, 12, 12)):
        ws2.column_dimensions[col].width = larg
    arq = pasta / f"conferencia_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    wb.save(arq)
    return arq


# ------------------------------------------------------------------ principal
def main():
    sem_perguntar = "--sem-perguntar" in sys.argv
    arq_fat = achar_arquivos(PADRAO_FATURAMENTO)[0]
    arq_bol = achar_arquivos(PADRAO_BOLETOS)[0]
    arqs_nf = achar_arquivos(PADRAO_NFS, unico=False)

    lf, lb = pypdf.PdfReader(str(arq_fat)), pypdf.PdfReader(str(arq_bol))
    faturas, boletos = ler_faturas(lf), ler_boletos(lb)
    nfs, _ = ler_nfs(arqs_nf)
    res = cruzar(faturas, boletos, nfs)

    sobra_bol, sobra_nf = mostrar_conferencia(res, boletos, nfs, {
        "Faturamento": f"{arq_fat.name} ({len(lf.pages)} pág, {len(faturas)} faturas)",
        "Boletos": f"{arq_bol.name} ({len(lb.pages)} pág, {len(boletos)} boletos)",
        "NFs": f"{', '.join(a.name for a in arqs_nf)} ({len(nfs)} NFs)"})
    saida = arq_fat.parent / PASTA_SAIDA
    xls = salvar_excel(res, sobra_bol, sobra_nf, saida / "Conferencia")
    if xls:
        print(f" Conferência salva em: {xls}")

    if not sem_perguntar:
        resp = input("\n Gerar os arquivos por cliente agora? (s/n): ").strip().lower()
        if resp not in ("s", "sim"):
            print(" Nada foi gerado.")
            return

    saida.mkdir(exist_ok=True)
    paginas = 0
    for r in res:
        w = pypdf.PdfWriter()
        for p in r["fatura"]["paginas"]:
            w.add_page(lf.pages[p])
        for b in r["boletos"]:
            w.add_page(lb.pages[b["pagina"]])
        for num in r["nfs"]:
            for leitor, p in nfs[num]["paginas"]:
                w.add_page(leitor.pages[p])
        nome = nome_arquivo(r["fatura"]["nome"]) + ("" if r["boletos"] else " - SEM BOLETO")
        with open(saida / f"{nome}.pdf", "wb") as fh:
            w.write(fh)
        paginas += len(w.pages)
    esperado = (sum(len(r["fatura"]["paginas"]) for r in res)
                + sum(len(r["boletos"]) for r in res)
                + sum(len(nfs[n]["paginas"]) for r in res for n in r["nfs"]))
    print(f"\n {len(res)} arquivos gerados em {saida}")
    print(f" Páginas gravadas: {paginas} | esperadas: {esperado} "
          f"{'(OK)' if paginas == esperado else '(DIFERENTE - verificar!)'}")


if __name__ == "__main__":
    main()
