"""
Gerador de DANFSe v2.0 — Nota Técnica Nº 008 (05/05/2026)
Uso: python gerar_danfse.py NFSE_xxx.xml
     python gerar_danfse.py xmls/
"""
import os, sys, glob
from xml.etree import ElementTree as ET
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import qrcode
from io import BytesIO

NS  = {"n": "http://www.sped.fazenda.gov.br/nfse"}
W, H = A4   # 595.27 x 841.89 pts
PT  = 72 / 2.54   # pontos por cm

def cm(v):  return v * PT
def rl(y_cm):  return H - cm(y_cm)   # converte Y de cima-pra-baixo para ReportLab

# ── XML ────────────────────────────────────────────────────
def get(el, tag):
    if el is None: return None
    return el.find(f"n:{tag}", NS)

def val(el, tag):
    f = get(el, tag)
    return (f.text or "").strip() if f is not None else ""

def cnpj(v):
    c = "".join(d for d in str(v or "") if d.isdigit())
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c)==14 else (v or "-")

def cep(v):
    c = "".join(d for d in str(v or "") if d.isdigit())
    return f"{c[:5]}-{c[5:]}" if len(c)==8 else (v or "-")

def data(v):
    if not v: return "-"
    v = v[:10]
    if len(v)==10 and v[4]=="-":
        return f"{v[8:10]}/{v[5:7]}/{v[0:4]}"
    return v

def dh(v):
    if not v: return "-"
    try:
        d2, t = v[:19].split("T")
        return f"{data(d2)} {t}"
    except:
        return v[:19]

def moeda(v):
    if not v: return "R$ 0,00"
    try:
        n = float(v)
        s = f"{n:,.2f}".replace(",","X").replace(".",",").replace("X",".")
        return f"R$ {s}"
    except: return str(v)

def aliq(v):
    if not v: return "-"
    try: return f"{float(v):.2f}%"
    except: return str(v)

CSTAT  = {"100":"NFS-e Gerada","101":"NFS-e Cancelada","102":"Decisão Judicial/Adm.",
           "103":"NFS-e Avulsa","107":"Gerada (Município Conv.)","108":"Cancelada (Município Conv.)"}
TEMIT  = {"1":"Prestador","2":"Tomador","3":"Intermediário"}
TISSQN = {"1":"Operação Tributável","2":"Imune","3":"Isenta","4":"Não Incidência"}
TRET   = {"1":"Não Retido","2":"Retido pelo Tomador","3":"Retido pelo Intermediário"}
OSIMP  = {"1":"Não Optante","2":"Optante pelo Simples","3":"MEI"}

def parse(path):
    root = ET.parse(path).getroot()
    inf     = root.find("n:infNFSe", NS)
    emit    = get(inf,"emit")
    vi      = get(inf,"valores")
    dps     = get(inf,"DPS")
    idps    = get(dps,"infDPS")
    prest   = get(idps,"prest")
    toma    = get(idps,"toma")
    serv    = get(idps,"serv")
    cserv   = get(serv,"cServ")
    vd      = get(idps,"valores")
    trib    = get(vd,"trib")
    tmun    = get(trib,"tribMun")
    tfed    = get(trib,"tribFed")
    tt      = get(trib,"totTrib")
    vtot    = get(tt,"vTotTrib")
    vsp     = get(vd,"vServPrest")
    rt      = get(prest,"regTrib")
    ee      = get(emit,"enderNac")

    et = None
    if toma is not None:
        e2 = get(toma,"end")
        if e2 is not None:
            et = get(e2,"endNac")
            if et is None:
                et = get(e2,"endExt")

    uf = val(ee,"UF")
    mun_e = val(inf,"xLocEmi")
    cstat = val(inf,"cStat")

    def end_str(lgr,nro,cpl,bairro):
        return ", ".join(p for p in [lgr,nro,cpl,bairro] if p) or "-"

    return {
        "chave_id": inf.get("Id","") if inf is not None else "",
        "chave":    (inf.get("Id","") if inf is not None else "").replace("NFS",""),
        "n_nfse":   val(inf,"nNFSe"),
        "d_compet": data(val(idps,"dCompet")),
        "dh_proc":  dh(val(inf,"dhProc")),
        "n_dps":    val(idps,"nDPS"),
        "serie":    val(idps,"serie"),
        "dh_emi":   dh(val(idps,"dhEmi")),
        "tp_emit":  TEMIT.get(val(idps,"tpEmit"),"Prestador"),
        "situacao": CSTAT.get(cstat,f"Código {cstat}"),
        "finalidade":"NFS-e Regular",
        "amb":      "1 - Produção" if val(inf,"ambGer")=="1" else "2 - Homologação",
        "tp_amb":   "1 - Produção" if val(idps,"tpAmb")=="1" else "2 - Homologação",
        "mun_emit": f"{mun_e} / {uf}",
        "homolog":  val(idps,"tpAmb")=="2",
        "cancelada":cstat in ("101","108"),
        # prestador
        "p_cnpj":   cnpj(val(emit,"CNPJ")),
        "p_im":     val(emit,"IM"),
        "p_fone":   val(emit,"fone"),
        "p_nome":   val(emit,"xNome"),
        "p_mun":    f"{mun_e} / {uf}",
        "p_cep":    cep(val(ee,"CEP")),
        "p_end":    end_str(val(ee,"xLgr"),val(ee,"nro"),val(ee,"xCpl"),val(ee,"xBairro")),
        "p_simp":   OSIMP.get(val(rt,"opSimpNac"),"-"),
        # tomador
        "t_cnpj":   cnpj(val(toma,"CNPJ")) if toma is not None else "-",
        "t_nome":   val(toma,"xNome") if toma is not None else "-",
        "t_end":    end_str(val(et,"xLgr"),val(et,"nro"),"",val(et,"xBairro")) if et is not None else "-",
        "t_cep":    cep(val(et,"CEP")) if et is not None else "-",
        # serviço
        "s_cod":    val(cserv,"cTribNac"),
        "s_nbs":    val(cserv,"cNBS") or "-",
        "s_loc":    val(inf,"xLocPrestacao") or "-",
        "s_dtrib":  val(inf,"xTribNac"),
        "s_desc":   val(cserv,"xDescServ").replace("\\n","\n"),
        # issqn
        "i_tp":     TISSQN.get(val(tmun,"tribISSQN"),"-"),
        "i_mun":    val(inf,"xLocIncid") or "-",
        "i_ret":    TRET.get(val(tmun,"tpRetISSQN"),"-"),
        "i_bc":     moeda(val(vi,"vBC")),
        "i_aliq":   aliq(val(vi,"pAliqAplic")),
        "i_iss":    moeda(val(vi,"vISSQN")),
        # federal
        "f_irrf":   moeda(val(tfed,"vRetIRRF")),
        "f_inss":   moeda(val(tfed,"vRetCP")),
        "f_csll":   moeda(val(tfed,"vRetCSLL")),
        # totais
        "v_serv":   moeda(val(vsp,"vServ")),
        "v_ret":    moeda(val(vi,"vTotalRet")),
        "v_liq":    moeda(val(vi,"vLiq")),
        "v_tfed":   moeda(val(vtot,"vTotTribFed")),
        "v_test":   moeda(val(vtot,"vTotTribEst")),
        "v_tmun":   moeda(val(vtot,"vTotTribMun")),
    }

# ── Desenho ────────────────────────────────────────────────
CINZA = colors.Color(.92,.92,.92)
BORDA = colors.Color(.55,.55,.55)
PRETO = colors.black
VERDE = colors.Color(0,.48,0)
MAR   = 0.40   # margem lateral cm
FW    = 20.20  # largura conteúdo cm
LH    = 0.58   # altura linha cm
BH    = 0.50   # altura título bloco cm
LBL   = 5.5    # pt label
VAL   = 7.0    # pt valor
LH    = 0.75   # altura linha cm (sobrescreve a anterior)

def gerar_pdf(d, saida):
    c = canvas.Canvas(saida, pagesize=A4)
    cur = [0.35]   # cursor Y em cm (mutable para closure)

    def adv(h): cur[0] += h

    # Rect: x,y em cm (y = topo da caixa a partir do topo da página)
    def box(x, y, w, h, shade=False, stroke=True):
        rx = cm(x)
        ry = rl(y + h)
        rw = cm(w)
        rh = cm(h)
        if shade:
            c.setFillColor(CINZA)
            c.setStrokeColor(BORDA)
            c.setLineWidth(0.5)
            c.rect(rx, ry, rw, rh, fill=1, stroke=1 if stroke else 0)
            c.setFillColor(PRETO)
        elif stroke:
            c.setFillColor(PRETO)
            c.setStrokeColor(BORDA)
            c.setLineWidth(0.5)
            c.rect(rx, ry, rw, rh, fill=0, stroke=1)

    # Texto: x,y cm (y = distância do topo da página até a baseline)
    def txt(t, x, y, sz=VAL, bold=False, col=PRETO, maxw=None):
        if not t: t = "-"
        c.setFillColor(col)
        fn = "Helvetica-Bold" if bold else "Helvetica"
        c.setFont(fn, sz)
        if maxw:
            while len(t) > 4 and c.stringWidth(t, fn, sz) > cm(maxw) - cm(0.10):
                t = t[:-4] + "..."
        c.drawString(cm(x) + 2, rl(y) + 1, t)

    # Campo = box + label no topo + valor centralizado abaixo
    def campo(lbl, v, x, w, h=None):
        if h is None: h = LH
        y = cur[0]
        box(x, y, w, h)
        # label no topo da caixa
        txt(lbl, x+0.06, y+0.06, sz=LBL, bold=True)
        # valor centralizado verticalmente
        txt(str(v or "-"), x+0.06, y+0.38, sz=VAL, maxw=w-0.12)

    # Linha de campos com avanço automático
    def linha(campos, h=LH):
        for lbl, v, x, w in campos:
            campo(lbl, v, x, w, h)
        adv(h)

    # Título de bloco
    def bloco(titulo):
        y = cur[0]
        box(MAR, y, FW, BH, shade=True)
        txt(titulo.upper(), MAR+0.06, y+0.10, sz=6.5, bold=True)
        adv(BH)

    # ── CABEÇALHO ──────────────────────────────────────────
    CH = 1.70
    box(MAR, cur[0], FW, CH, shade=True)

    # Logo
    c.setFillColor(VERDE)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(cm(MAR+0.12), rl(cur[0]+0.58), "NFS")
    c.setFillColor(PRETO)
    c.setFont("Helvetica", 6)
    c.drawString(cm(MAR+0.12), rl(cur[0]+0.82), "Nota Fiscal de")
    c.drawString(cm(MAR+0.12), rl(cur[0]+1.00), "Serviço Eletrônica")

    # Centro
    cx = cm(MAR + FW/2)
    c.setFillColor(PRETO)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(cx, rl(cur[0]+0.55), "DANFSe v2.0")
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(cx, rl(cur[0]+0.78), "Documento Auxiliar da NFS-e")
    if d["homolog"]:
        c.setFillColor(colors.red)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(cx, rl(cur[0]+1.02), "NFS-e SEM VALIDADE JURÍDICA")
        c.setFillColor(PRETO)

    # Direita (município e ambiente)
    rx = MAR + FW - 5.80
    c.setFont("Helvetica", 7)
    c.drawString(cm(rx), rl(cur[0]+0.48), f"Município: {d['mun_emit']}")
    c.setFont("Helvetica", 6)
    c.drawString(cm(rx), rl(cur[0]+0.68), f"Ambiente Gerador: {d['amb']}")
    c.drawString(cm(rx), rl(cur[0]+0.86), f"Tipo de Ambiente: {d['tp_amb']}")

    # QR Code
    url = f"https://www.nfse.gov.br/ConsultaPublica/?tpc=1&chave={d['chave_id']}"
    qr = qrcode.QRCode(version=2, box_size=4, border=1,
                       error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url); qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
    qsz = cm(1.55)
    c.drawImage(ImageReader(buf), cm(MAR+FW-1.65), rl(cur[0]+1.60), width=qsz, height=qsz)
    c.setFont("Helvetica", 5)
    for i, ln in enumerate(["A autenticidade desta NFS-e pode ser verificada",
                             "pela leitura deste código QR ou pela consulta da",
                             "chave de acesso no portal nacional da NFS-e"]):
        c.drawString(cm(rx), rl(cur[0]+1.10+i*0.18), ln)

    adv(CH + 0.10)

    # ── DADOS DA NFS-E ─────────────────────────────────────
    bloco("DADOS DA NFS-E")

    # Chave acesso
    chave_h = 0.52
    box(MAR, cur[0], FW-4.0, chave_h)
    txt("CHAVE DE ACESSO DA NFS-E", MAR+0.05, cur[0]+0.10, sz=LBL, bold=True)
    txt(d["chave"], MAR+0.05, cur[0]+chave_h-0.14, sz=7.5, maxw=FW-4.10)
    adv(chave_h)

    C = (FW - 0.0) / 3
    linha([
        ("NÚMERO DA NFS-E",                d["n_nfse"],   MAR,       C),
        ("COMPETÊNCIA DA NFS-E",           d["d_compet"], MAR+C,     C),
        ("DATA E HORA DA EMISSÃO DA NFS-E",d["dh_proc"],  MAR+2*C,   C),
    ])
    linha([
        ("NÚMERO DA DPS",                  d["n_dps"],    MAR,       C),
        ("SÉRIE DA DPS",                   d["serie"],    MAR+C,     C),
        ("DATA E HORA DA EMISSÃO DA DPS",  d["dh_emi"],   MAR+2*C,   C),
    ])
    linha([
        ("EMITENTE DA NFS-E",              d["tp_emit"],  MAR,       C),
        ("SITUAÇÃO DA NFS-E",              d["situacao"], MAR+C,     C),
        ("FINALIDADE",                     d["finalidade"],MAR+2*C,  C),
    ])

    # ── PRESTADOR ──────────────────────────────────────────
    bloco("PRESTADOR / FORNECEDOR")
    C1, C2, C3 = 5.10, 5.10, FW-10.20
    linha([
        ("CNPJ / CPF / NIF",              d["p_cnpj"],   MAR,       C1),
        ("INDICADOR MUNICIPAL",           d["p_im"],     MAR+C1,    C2),
        ("TELEFONE",                      d["p_fone"],   MAR+C1+C2, C3),
    ])
    linha([
        ("NOME / NOME EMPRESARIAL",       d["p_nome"],   MAR,       C1+C2),
        ("MUNICÍPIO / SIGLA UF",          d["p_mun"],    MAR+C1+C2, C3),
    ])
    linha([
        ("ENDEREÇO",                      d["p_end"],    MAR,       C1+C2),
        ("CÓDIGO IBGE / CEP",             d["p_cep"],    MAR+C1+C2, C3),
    ])
    linha([
        ("SIMPLES NACIONAL NA DATA DE COMPETÊNCIA", d["p_simp"], MAR, FW),
    ])

    # ── TOMADOR ────────────────────────────────────────────
    bloco("TOMADOR / ADQUIRENTE")
    linha([("CNPJ / CPF / NIF",           d["t_cnpj"],   MAR,       C1+C2)])
    linha([("NOME / NOME EMPRESARIAL",    d["t_nome"],   MAR,       FW)])
    linha([
        ("ENDEREÇO",                      d["t_end"],    MAR,       C1+C2),
        ("CEP",                           d["t_cep"],    MAR+C1+C2, C3),
    ])

    # ── SERVIÇO ────────────────────────────────────────────
    bloco("SERVIÇO PRESTADO")
    linha([
        ("CÓD. DE TRIBUTAÇÃO NACIONAL",   d["s_cod"],    MAR,       C1),
        ("CÓDIGO DA NBS",                 d["s_nbs"],    MAR+C1,    C2),
        ("LOCAL DA PRESTAÇÃO / UF",       d["s_loc"],    MAR+C1+C2, C3),
    ])
    # Descrição tributação
    dth = 0.52
    box(MAR, cur[0], FW, dth)
    txt("DESCRIÇÃO DO CÓDIGO DE TRIBUTAÇÃO NACIONAL", MAR+0.05, cur[0]+0.10, sz=LBL, bold=True)
    txt(d["s_dtrib"][:165], MAR+0.05, cur[0]+dth-0.14, sz=6.5, maxw=FW-0.10)
    adv(dth)
    # Descrição serviço
    linhas_s = [l for l in d["s_desc"].replace("\r","").split("\n") if l.strip()]
    n = max(2, min(6, len(linhas_s)))
    sh = 0.22 * n + 0.28
    box(MAR, cur[0], FW, sh)
    txt("DESCRIÇÃO DO SERVIÇO", MAR+0.05, cur[0]+0.10, sz=LBL, bold=True)
    for i, ln in enumerate(linhas_s[:n]):
        txt(ln, MAR+0.05, cur[0]+sh-0.20-i*0.22, sz=7, maxw=FW-0.10)
    adv(sh)

    # ── ISSQN ──────────────────────────────────────────────
    bloco("TRIBUTAÇÃO MUNICIPAL (ISSQN)")
    linha([
        ("TIPO DE TRIBUTAÇÃO DO ISSQN",   d["i_tp"],     MAR,       C1),
        ("MUNICÍPIO / UF / PAÍS DA INCIDÊNCIA DO ISSQN", d["i_mun"], MAR+C1, C1+C2),
    ])
    linha([
        ("BC ISSQN",                      d["i_bc"],     MAR,       C1),
        ("ALÍQUOTA APLICADA",             d["i_aliq"],   MAR+C1,    C2),
        ("RETENÇÃO DO ISSQN",             d["i_ret"],    MAR+C1+C2, C2),
        ("ISSQN APURADO",                 d["i_iss"],    MAR+C1+2*C2, C3-C2+C2),
    ])

    # ── FEDERAL ────────────────────────────────────────────
    bloco("TRIBUTAÇÃO FEDERAL (EXCETO CBS)")
    linha([
        ("IRRF",                          d["f_irrf"],   MAR,       C1),
        ("CONTRIBUIÇÃO PREVIDENCIÁRIA - RETIDA", d["f_inss"], MAR+C1, C2),
        ("CONTRIBUIÇÕES SOCIAIS - RETIDAS", d["f_csll"], MAR+C1+C2, C3),
    ])

    # ── VALOR TOTAL ────────────────────────────────────────
    bloco("VALOR TOTAL DA NFS-E")
    linha([
        ("VALOR DA OPERAÇÃO / SERVIÇO",   d["v_serv"],   MAR,       C1),
        ("TOTAL DAS RETENÇÕES (ISSQN / FEDERAIS)", d["v_ret"], MAR+C1, C2),
    ])
    # Valor líquido destacado
    vlh = 0.65
    box(MAR, cur[0], C1+C2, vlh, shade=True)
    txt("VALOR LÍQUIDO DA NFS-E", MAR+0.06, cur[0]+0.10, sz=LBL, bold=True)
    txt(d["v_liq"], MAR+0.06, cur[0]+vlh-0.16, sz=11, bold=True, maxw=C1+C2-0.12)
    adv(vlh)

    # ── INFORMAÇÕES COMPLEMENTARES ─────────────────────────
    bloco("INFORMAÇÕES COMPLEMENTARES")
    ith = 0.58
    box(MAR, cur[0], FW, ith)
    trib_txt = (f"Totais Aproximados dos Tributos cfe. Lei nº 12.741/2012: "
                f"Federais: {d['v_tfed']} ; Estaduais: {d['v_test']} ; Municipais: {d['v_tmun']}")
    txt(trib_txt, MAR+0.05, cur[0]+ith-0.14, sz=6.5, maxw=FW-0.10)
    adv(ith)

    # ── CANHOTO ────────────────────────────────────────────
    adv(0.18)
    bloco("CANHOTO (RECEBIMENTO)")
    linha([
        ("DATA DE CIENTIFICAÇÃO",         "",            MAR,       C1),
        ("IDENTIFICAÇÃO E ASSINATURA",    "",            MAR+C1,    C2),
        (f"Nº NFS-E / CHAVE", f"{d['n_nfse']} / {d['chave'][:42]}...", MAR+C1+C2, C3+C1),
    ])

    # Marca d'água
    if d["cancelada"]:
        c.saveState()
        c.setFillColor(colors.Color(.65,.65,.65,alpha=0.4))
        c.setFont("Helvetica-Bold", 72)
        c.translate(W/2, H/2)
        c.rotate(40)
        c.drawCentredString(0, 0, "CANCELADA")
        c.restoreState()

    # Borda da página
    c.setStrokeColor(PRETO)
    c.setLineWidth(1)
    tot_h = cur[0] + 0.15
    c.rect(cm(MAR), rl(tot_h), cm(FW), cm(tot_h - 0.35), fill=0, stroke=1)

    c.save()
    return saida

def processar(caminho):
    if os.path.isdir(caminho):
        xmls = sorted(glob.glob(os.path.join(caminho,"NFSE_*.xml")))
        print(f"{len(xmls)} XMLs em '{caminho}'")
        os.makedirs("danfse_pdfs", exist_ok=True)
        ok = err = 0
        for xml in xmls:
            saida = os.path.join("danfse_pdfs", os.path.splitext(os.path.basename(xml))[0]+".pdf")
            try:
                gerar_pdf(parse(xml), saida); ok += 1
            except Exception as e:
                print(f"  ERRO {xml}: {e}"); err += 1
        print(f"{ok} OK | {err} erros → danfse_pdfs/")
    else:
        saida = os.path.splitext(caminho)[0]+".pdf"
        gerar_pdf(parse(caminho), saida)
        print(f"PDF gerado: {saida}")

if __name__=="__main__":
    processar(sys.argv[1] if len(sys.argv)>1 else "xmls/")
