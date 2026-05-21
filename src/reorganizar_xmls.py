"""
reorganizar_xmls.py
Executa após o ingest para:
1. Cruzar eventos com notas em disco
2. Mover XMLs para as subpastas corretas (Cancelada/, Substituida/, Rejeitadas/)
3. Mover a nota original junto com seu evento de cancelamento
4. Detectar cancelamentos retroativos (evento em competência diferente da nota)
5. Gerar alertas_retroativos.csv para o Power BI

Uso: python reorganizar_xmls.py [pasta_xmls] [empresa]
     python reorganizar_xmls.py xmls/EDB EDB
     python reorganizar_xmls.py xmls/Azulão Azulão
"""

import os
import sys
import glob
import shutil
import csv
import logging
from datetime import datetime
from xml.etree import ElementTree as ET
from eventos_config import info_evento

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.StreamHandler()]
)
log = logging.getLogger(__name__)

NS = {"n": "http://www.sped.fazenda.gov.br/nfse"}
os.makedirs("dados", exist_ok=True)


# ── Helpers ────────────────────────────────────────────────
def slug_seguro(texto: str) -> str:
    for c in r'\/:*?"<>|':
        texto = texto.replace(c, "_")
    return texto.strip()


def val(el, path):
    f = el.find(path, NS) if el is not None else None
    return (f.text or "").strip() if f is not None else ""


def competencia_de(data_str: str) -> str:
    """'2025-10-15' ou '2025-10-15T10:30:00' → '2025-10'"""
    if data_str and len(data_str) >= 7:
        return data_str[:7]
    return ""


def parse_nfse_basico(path: str) -> dict:
    """Extrai chave, competência e cStat da NFS-e."""
    try:
        root  = ET.parse(path).getroot()
        inf   = root.find("n:infNFSe", NS)
        idps  = root.find("n:infNFSe/n:DPS/n:infDPS", NS)
        chave = (inf.get("Id", "") if inf is not None else "").lstrip("NFS")
        if chave.startswith("NFS"):
            chave = chave[3:]
        cstat = val(inf, "n:cStat")
        comp  = competencia_de(val(idps, "n:dCompet"))
        ch_substda = val(root.find("n:infNFSe/n:DPS/n:infDPS/n:subst", NS), "n:chSubstda")
        return {"chave": chave, "cstat": cstat, "comp": comp,
                "ch_substda": ch_substda, "path": path}
    except Exception as e:
        log.debug(f"  Erro ao ler NFS-e {path}: {e}")
        return {}


def parse_evento_basico(path: str) -> dict:
    """Extrai tag, chave da nota vinculada, data e chSubstituta do evento."""
    try:
        root    = ET.parse(path).getroot()
        inf_e   = root.find("n:infEvento", NS)
        inf_ped = root.find("n:infEvento/n:pedRegEvento/n:infPedReg", NS)

        chave_nota = val(inf_ped, "n:chNFSe")
        dh_evento  = val(inf_ped, "n:dhEvento")
        dh_proc    = val(inf_e, "n:dhProc")
        comp_evt   = competencia_de(dh_evento)

        tag_ev = ""
        ch_substituta = ""
        for child in (inf_ped or []):
            t = child.tag.split("}")[-1]
            if t.startswith("e") and len(t) > 1:
                tag_ev = t
                ch_sub = child.find("n:chSubstituta", NS)
                ch_substituta = (ch_sub.text or "").strip() if ch_sub is not None else ""
                break

        descr, situacao, subpasta, cancela = info_evento(tag_ev)

        return {
            "tag": tag_ev, "chave_nota": chave_nota,
            "comp_evento": comp_evt, "dh_evento": dh_evento,
            "dh_proc": dh_proc, "subpasta": subpasta,
            "situacao": situacao, "cancela": cancela,
            "ch_substituta": ch_substituta,
            "descr": descr, "path": path,
        }
    except Exception as e:
        log.debug(f"  Erro ao ler Evento {path}: {e}")
        return {}


def mover(path_atual: str, subpasta: str) -> str:
    """
    Move o arquivo para a subpasta dentro da mesma pasta de competência.
    Retorna o novo caminho.
    """
    pasta_comp = os.path.dirname(path_atual)
    # Se já está numa subpasta (Cancelada/, Substituida/, Rejeitadas/),
    # usa a pasta pai como base
    nome_pasta = os.path.basename(pasta_comp)
    if nome_pasta in ("Cancelada", "Substituida", "Rejeitadas"):
        pasta_comp = os.path.dirname(pasta_comp)

    destino_pasta = os.path.join(pasta_comp, subpasta)
    os.makedirs(destino_pasta, exist_ok=True)
    destino = os.path.join(destino_pasta, os.path.basename(path_atual))

    if os.path.abspath(path_atual) != os.path.abspath(destino):
        shutil.move(path_atual, destino)
    return destino


def subpasta_atual(path: str) -> str:
    """Retorna a subpasta onde o arquivo está ('Cancelada', '' etc.)"""
    nome = os.path.basename(os.path.dirname(path))
    if nome in ("Cancelada", "Substituida", "Rejeitadas"):
        return nome
    return ""


# ── Lógica principal ───────────────────────────────────────
def reorganizar(pasta_raiz: str, empresa: str,
                arquivo_alertas: str = "dados/alertas_retroativos.csv") -> dict:
    log.info("=" * 60)
    log.info(f"Reorganizando XMLs — {empresa}")
    log.info(f"Pasta: {pasta_raiz}")
    log.info("=" * 60)

    # 1) Indexa todas as NFS-e em disco
    notas = {}   # chave → dict
    for path in glob.glob(os.path.join(pasta_raiz, "**", "NFSE_*.xml"), recursive=True):
        n = parse_nfse_basico(path)
        if n and n["chave"]:
            # Mantém a mais recente se duplicada
            if n["chave"] not in notas or \
               os.path.getmtime(path) > os.path.getmtime(notas[n["chave"]]["path"]):
                notas[n["chave"]] = n
    log.info(f"NFS-e indexadas: {len(notas)}")

    # 2) Lê todos os eventos
    eventos = []
    for path in glob.glob(os.path.join(pasta_raiz, "**", "EVENTO_*.xml"), recursive=True):
        e = parse_evento_basico(path)
        if e and e["chave_nota"] and e["subpasta"]:
            eventos.append(e)
    log.info(f"Eventos com efeito de cancelamento: {len(eventos)}")

    # 3) Processa cada evento
    cont = {"movidos": 0, "ja_corretos": 0, "nota_nao_encontrada": 0,
            "retroativos": 0, "erros": 0}
    alertas = []
    nao_encontrados = []

    for ev in eventos:
        chave_nota = ev["chave_nota"]
        nota = notas.get(chave_nota)

        # Move o evento para a subpasta correta
        sub_ev = ev["subpasta"]
        path_ev_atual = ev["path"]
        if subpasta_atual(path_ev_atual) != sub_ev:
            try:
                ev["path"] = mover(path_ev_atual, sub_ev)
                log.debug(f"  Evento movido → {sub_ev}/: {os.path.basename(ev['path'])}")
            except Exception as e:
                log.warning(f"  Erro ao mover evento {path_ev_atual}: {e}")
                cont["erros"] += 1

        if not nota:
            log.warning(f"  Nota não encontrada em disco: {chave_nota[:25]}...")
            cont["nota_nao_encontrada"] += 1
            nao_encontrados.append({
                "Empresa":           empresa,
                "Chave NFS-e":       chave_nota,
                "Competência Evento": ev["comp_evento"],
                "Data Evento":       ev["dh_evento"][:19] if ev["dh_evento"] else "",
                "Tipo Evento":       ev["descr"],
                "Arquivo Evento":    os.path.basename(ev["path"]),
                "Gerado em":         datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            continue

        # Verifica se é retroativo
        comp_nota = nota["comp"]
        comp_evt  = ev["comp_evento"]
        eh_retro  = bool(comp_nota and comp_evt and comp_nota != comp_evt)

        if eh_retro:
            cont["retroativos"] += 1
            alerta = {
                "Empresa":          empresa,
                "Chave NFS-e":      chave_nota,
                "Competência Nota": comp_nota,
                "Competência Evento": comp_evt,
                "Data Evento":      ev["dh_evento"][:19] if ev["dh_evento"] else "",
                "Data Processamento": ev["dh_proc"][:19] if ev["dh_proc"] else "",
                "Tipo Evento":      ev["descr"],
                "Situação":         ev["situacao"],
                "Arquivo NFS-e":    os.path.basename(nota["path"]),
                "Arquivo Evento":   os.path.basename(ev["path"]),
                "Gerado em":        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            alertas.append(alerta)
            log.info(
                f"  ⚠ RETROATIVO: nota {comp_nota} cancelada por evento de {comp_evt} "
                f"| {os.path.basename(nota['path'])}"
            )

        # Move a nota para a subpasta correta se necessário
        path_nota_atual = nota["path"]
        sub_correta = ev["subpasta"]

        if subpasta_atual(path_nota_atual) == sub_correta:
            cont["ja_corretos"] += 1
        else:
            try:
                novo_path = mover(path_nota_atual, sub_correta)
                notas[chave_nota]["path"] = novo_path
                cont["movidos"] += 1
                log.info(
                    f"  Movida → {sub_correta}/: {os.path.basename(novo_path)}"
                    f"{' ⚠ RETROATIVA' if eh_retro else ''}"
                )
            except Exception as e:
                log.warning(f"  Erro ao mover nota {path_nota_atual}: {e}")
                cont["erros"] += 1

    # 4) Grava/atualiza alertas_retroativos.csv
    _gravar_alertas(alertas, arquivo_alertas, empresa)

    # 5) Grava/atualiza notas_nao_encontradas.csv
    arquivo_nf = arquivo_alertas.replace("alertas_retroativos", "notas_nao_encontradas")
    _gravar_nao_encontrados(nao_encontrados, arquivo_nf, empresa)

    log.info("\n" + "=" * 60)
    log.info("Reorganização concluída.")
    log.info(f"  Notas movidas       : {cont['movidos']}")
    log.info(f"  Já na pasta correta : {cont['ja_corretos']}")
    log.info(f"  Nota não encontrada : {cont['nota_nao_encontrada']}")
    log.info(f"  ⚠ Retroativos       : {cont['retroativos']}")
    log.info(f"  Erros               : {cont['erros']}")
    if cont["retroativos"] > 0:
        log.info(f"  Alertas gravados em : {arquivo_alertas}")
    if cont["nota_nao_encontrada"] > 0:
        arquivo_nf = arquivo_alertas.replace("alertas_retroativos", "notas_nao_encontradas")
        log.info(f"  Não encontradas em  : {arquivo_nf}")
    log.info("=" * 60)

    return cont


def _gravar_alertas(novos: list, arquivo: str, empresa: str):
    """
    Atualiza o CSV de alertas — preserva alertas de outras empresas
    e substitui os da empresa atual pelos novos.
    """
    cols = [
        "Empresa", "Chave NFS-e", "Competência Nota", "Competência Evento",
        "Data Evento", "Data Processamento", "Tipo Evento", "Situação",
        "Arquivo NFS-e", "Arquivo Evento", "Gerado em",
    ]

    existentes = []
    if os.path.exists(arquivo):
        with open(arquivo, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            existentes = [r for r in reader if r.get("Empresa") != empresa]

    todos = existentes + novos

    with open(arquivo, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(todos)

    if todos:
        log.info(f"  alertas_retroativos.csv: {len(todos)} alertas total ({len(novos)} novos de {empresa})")


def _gravar_nao_encontrados(novos: list, arquivo: str, empresa: str):
    """
    Atualiza o CSV de notas não encontradas — preserva registros de outras empresas.
    """
    cols = [
        "Empresa", "Chave NFS-e", "Competência Evento",
        "Data Evento", "Tipo Evento", "Arquivo Evento", "Gerado em",
    ]
    existentes = []
    if os.path.exists(arquivo):
        with open(arquivo, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            existentes = [r for r in reader if r.get("Empresa") != empresa]

    todos = existentes + novos
    with open(arquivo, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(todos)

    if novos:
        log.info(f"  notas_nao_encontradas.csv: {len(todos)} registros total ({len(novos)} novos de {empresa})")


# ── Entry point ────────────────────────────────────────────
if __name__ == "__main__":
    pasta  = sys.argv[1] if len(sys.argv) > 1 else "xmls/EDB"
    emp    = sys.argv[2] if len(sys.argv) > 2 else os.path.basename(pasta)
    reorganizar(pasta, emp)
