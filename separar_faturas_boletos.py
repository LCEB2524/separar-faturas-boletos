"""
Separa o faturamento agrupado e os boletos em um PDF por cliente.

Entrada (na mesma pasta configurada abaixo ou passada por argumento):
  - PDF de faturamento agrupado (uma fatura por página, campo "CLIENTE")
  - PDF de boletos Sicredi (um boleto por página, campo "Pagador" e "Ref. Notas Fiscais")

Cruzamento fatura x boleto, nesta ordem:
  1. Nome do cliente = Pagador (exato)
  2. Nome truncado (o boleto corta o nome do pagador em 40 caracteres)
  3. Números de NF da fatura x "Ref. Notas Fiscais" do boleto

Saída: pasta "saida" com um PDF por cliente (fatura + boletos).
Cliente sem boleto gera só a fatura, com " - SEM BOLETO" no nome.
Ao final mostra a conferência de valores, duplicatas e NFs.

Uso:
  python separar_faturas_boletos.py
  python separar_faturas_boletos.py "C:\\pasta\\faturamento.pdf" "C:\\pasta\\boletos.pdf"
"""

import logging
import re
import sys
import unicodedata
from pathlib import Path

import pypdf

logging.disable(logging.WARNING)  # PDFs do banco geram muitos avisos inofensivos

# ------------------------------------------------------------------ configuração
PASTA = Path(__file__).resolve().parent.parent      # pasta "conth" (um nível acima do script)
PADRAO_FATURAMENTO = "Emiss*Faturamento*.pdf"      # nome do PDF de faturamento
PADRAO_BOLETOS = "boleto*.pdf"                      # nome do PDF de boletos
PASTA_SAIDA = "saida"


# ------------------------------------------------------------------ utilidades
def normalizar(txt):
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", txt.upper()).strip()


def valor_br(v):
    return float(v.replace(".", "").replace(",", ".")) if v else 0.0


def achar_arquivo(padrao):
    import fnmatch
    achados = [p for p in PASTA.iterdir()
               if p.is_file() and fnmatch.fnmatch(p.name.lower(), padrao.lower())]
    if not achados:
        sys.exit(f"Não encontrei arquivo '{padrao}' em {PASTA}")
    if len(achados) > 1:
        sys.exit(f"Mais de um arquivo '{padrao}' em {PASTA}: {[p.name for p in achados]}")
    return achados[0]


def nome_arquivo(nome):
    return re.sub(r'[\\/:*?"<>|]', "", nome).strip().rstrip(".")


# ------------------------------------------------------------------ leitura
def ler_faturas(leitor):
    faturas = []
    for i, pag in enumerate(leitor.pages):
        t = pag.extract_text() or ""
        m = re.search(r"CLIENTE\s*:\s*\d+-(.+?)\s+(?:TELEFONE|FONE)", t)
        if not m:
            # página de continuação: anexa à fatura anterior
            if faturas:
                faturas[-1]["paginas"].append(i)
                faturas[-1]["nfs"] += [str(int(x)) for x in
                                       re.findall(r"\d{8}\s+(\d{8,10})\s+[\d.]+,\d{2}N", t)]
            continue
        nome = m.group(1).strip()
        nfs = [str(int(x)) for x in re.findall(r"\d{8}\s+(\d{8,10})\s+[\d.]+,\d{2}N", t)]
        if faturas and faturas[-1]["nome"] == nome and not faturas[-1]["fechada"]:
            faturas[-1]["paginas"].append(i)
            faturas[-1]["nfs"] += nfs
        else:
            faturas.append({"nome": nome, "paginas": [i], "nfs": nfs, "dups": [],
                            "valor": None, "qtde": None, "fechada": False})
        f = faturas[-1]
        f["dups"] += re.findall(r"(\d{9}[A-Z])\s", t)
        vf = re.search(r"VALOR FATURA:\s*([\d.,]+)", t)
        q = re.search(r"QTDE\. NOTAS:(\d+)", t)
        if vf:
            f["valor"] = vf.group(1)
            f["fechada"] = True
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
        boletos.append({"pagina": i, "nome": m.group(1).strip() if m else "",
                        "nfs": nfs, "doc": doc.group(1) if doc else "",
                        "valor": v.group(1) if v else ""})
    return boletos


# ------------------------------------------------------------------ cruzamento
def cruzar(faturas, boletos):
    resultado = []
    for f in faturas:
        nf = normalizar(f["nome"])
        bs, como = [b for b in boletos if normalizar(b["nome"]) == nf], "nome"
        if not bs:
            bs = [b for b in boletos
                  if len(normalizar(b["nome"])) >= 15 and nf.startswith(normalizar(b["nome"]))]
            como = "nome truncado"
        if not bs:
            bs = [b for b in boletos if set(b["nfs"]) & set(f["nfs"])]
            como = "NF"
        if not bs:
            como = "sem boleto"
        nfs_bol = set().union(*[set(b["nfs"]) for b in bs]) if bs else set()
        resultado.append({
            "fatura": f, "boletos": bs, "como": como,
            "valor_fat": valor_br(f["valor"]),
            "valor_bol": round(sum(valor_br(b["valor"]) for b in bs), 2),
            "dup_ok": sorted(f["dups"]) == sorted(b["doc"] for b in bs),
            "nf_so_fatura": sorted(set(f["nfs"]) - nfs_bol),
            "nf_so_boleto": sorted(nfs_bol - set(f["nfs"])),
        })
    return resultado


# ------------------------------------------------------------------ principal
def main():
    if len(sys.argv) == 3:
        arq_fat, arq_bol = Path(sys.argv[1]), Path(sys.argv[2])
    else:
        arq_fat, arq_bol = achar_arquivo(PADRAO_FATURAMENTO), achar_arquivo(PADRAO_BOLETOS)
    saida = arq_fat.parent / PASTA_SAIDA
    saida.mkdir(exist_ok=True)

    lf, lb = pypdf.PdfReader(str(arq_fat)), pypdf.PdfReader(str(arq_bol))
    faturas, boletos = ler_faturas(lf), ler_boletos(lb)
    res = cruzar(faturas, boletos)

    usados = {b["pagina"] for r in res for b in r["boletos"]}
    paginas = 0
    for r in res:
        w = pypdf.PdfWriter()
        for p in r["fatura"]["paginas"]:
            w.add_page(lf.pages[p])
        for b in r["boletos"]:
            w.add_page(lb.pages[b["pagina"]])
        nome = nome_arquivo(r["fatura"]["nome"]) + ("" if r["boletos"] else " - SEM BOLETO")
        with open(saida / f"{nome}.pdf", "wb") as fh:
            w.write(fh)
        paginas += len(w.pages)

    # ---- conferência
    print("=" * 70)
    print(f" Faturas: {len(faturas)} | Boletos: {len(boletos)} | Arquivos gerados: {len(res)}")
    print(f" Páginas: {len(lf.pages)} fatura + {len(lb.pages)} boleto = "
          f"{len(lf.pages) + len(lb.pages)} | nos arquivos: {paginas}")
    print("=" * 70)
    for r in res:
        f = r["fatura"]
        alerta = []
        if r["boletos"] and abs(r["valor_fat"] - r["valor_bol"]) > 0.009:
            alerta.append(f"VALOR fatura {r['valor_fat']:.2f} x boletos {r['valor_bol']:.2f}")
        if r["boletos"] and not r["dup_ok"]:
            alerta.append("duplicata diferente do nº do documento do boleto")
        if r["boletos"] and r["nf_so_fatura"]:
            alerta.append(f"NFs fora da ref. do boleto: {', '.join(r['nf_so_fatura'])}")
        if r["nf_so_boleto"]:
            alerta.append(f"NFs no boleto e não na fatura: {', '.join(r['nf_so_boleto'])}")
        status = "SEM BOLETO" if not r["boletos"] else f"{len(r['boletos'])} boleto(s) [{r['como']}]"
        print(f" {f['nome'][:45]:45} {status}")
        for a in alerta:
            print(f"    ! {a}")
    sobra = [b for b in boletos if b["pagina"] not in usados]
    if sobra:
        print("\n BOLETOS SEM FATURA:")
        for b in sobra:
            print(f"    - pág. {b['pagina'] + 1}: {b['nome']} ({b['valor']})")
    print("=" * 70)
    print(f" Arquivos em: {saida}")


if __name__ == "__main__":
    main()
