import os, json, time, gzip, base64, logging, glob
from datetime import datetime
from xml.etree import ElementTree as ET
from dotenv import load_dotenv
from connector import DistribuicaoConnector
from eventos_config import subpasta_evento, info_evento
from reorganizar_xmls import reorganizar

load_dotenv()
os.makedirs("logs", exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(f"logs/ingest_{datetime.now():%Y%m%d_%H%M%S}.log", encoding="utf-8"),
        logging.StreamHandler(),
    ]
)
log = logging.getLogger(__name__)

# Preencha com os CNPJs da sua empresa
# Pode carregar de variável de ambiente, arquivo JSON ou banco de dados
# Exemplo:
#   "00000000000100": "Empresa - Matriz",
#   "00000000000200": "Empresa - Filial SP",
CNPJS_EMPRESA = {}

EMPRESA    = os.getenv("NFSE_EMPRESA", "minha_empresa")
STATE_FILE = os.getenv("NFSE_STATE_FILE", "state.json")
NS         = {"n": "http://www.sped.fazenda.gov.br/nfse"}


# ── Helpers ────────────────────────────────────────────────
def descompactar(b64: str) -> str:
    return gzip.decompress(base64.b64decode(b64)).decode("utf-8")

def slug_seguro(texto: str) -> str:
    for c in r'\/:*?"<>|':
        texto = texto.replace(c, "_")
    return texto.strip()

def competencia_do_xml(xml: str, tipo: str) -> str:
    try:
        root = ET.fromstring(xml)
        if tipo == "NFSE":
            el = root.find("n:infNFSe/n:DPS/n:infDPS/n:dCompet", NS)
        else:
            el = root.find("n:infEvento/n:pedRegEvento/n:infPedReg/n:dhEvento", NS)
        if el is not None and el.text and len(el.text) >= 7:
            return el.text[:7]
    except Exception:
        pass
    return datetime.now().strftime("%Y-%m")

def cstat_nfse(xml: str) -> str:
    try:
        root = ET.fromstring(xml)
        el = root.find("n:infNFSe/n:cStat", NS)
        return el.text.strip() if el is not None else "100"
    except Exception:
        return "100"

def tem_chsubstda(xml: str) -> bool:
    """Retorna True se a nota é substituta (tem <chSubstda> preenchido)."""
    try:
        root = ET.fromstring(xml)
        el = root.find("n:infNFSe/n:DPS/n:infDPS/n:subst/n:chSubstda", NS)
        return el is not None and bool((el.text or "").strip())
    except Exception:
        return False

def tag_e_chave_evento(xml: str) -> tuple:
    """Retorna (tag_evento, chave_nfse_vinculada) do XML de evento."""
    try:
        root = ET.fromstring(xml)
        inf_ped = root.find("n:infEvento/n:pedRegEvento/n:infPedReg", NS)
        if inf_ped is None:
            return "", ""
        ch_el = inf_ped.find("n:chNFSe", NS)
        chave = (ch_el.text or "").strip() if ch_el is not None else ""
        for child in inf_ped:
            tag = child.tag.split("}")[-1]
            if tag.startswith("e") and len(tag) > 1:
                return tag, chave
    except Exception:
        pass
    return "", ""

def buscar_competencia_nota(empresa: str, filial: str, chave: str) -> str:
    base   = os.path.join("xmls", empresa, slug_seguro(filial))
    padrao = os.path.join(base, "**", f"NFSE_*_{chave[:20]}*.xml")
    encontrados = glob.glob(padrao, recursive=True)
    if not encontrados:
        return ""
    try:
        partes = encontrados[0].replace("\\", "/").split("/")
        for parte in reversed(partes[:-1]):
            if len(parte) == 7 and parte[4] == "-":
                return parte
    except Exception:
        pass
    return ""

def pasta_destino(empresa: str, filial: str, competencia: str, subpasta: str = "") -> str:
    partes = ["xmls", empresa, slug_seguro(filial), competencia]
    if subpasta:
        partes.append(subpasta)
    pasta = os.path.join(*partes)
    os.makedirs(pasta, exist_ok=True)
    return pasta


# ── Estado ─────────────────────────────────────────────────
def carregar_estado() -> dict:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_estado(estado: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=2, ensure_ascii=False)


# ── Varredura ──────────────────────────────────────────────
def varrer_cnpj(connector, cnpj, nome_filial, nsu_inicial, estado) -> dict:
    nsu = nsu_inicial
    cont = {"nfse": 0, "substituta": 0, "eventos": 0,
            "canceladas": 0, "substituidas": 0, "rejeitadas": 0,
            "outros": 0, "erros": 0, "desconhecidos": 0}

    while True:
        resultado = connector.buscar_lote(nsu=nsu, cnpj_consulta=cnpj)

        if not resultado["sucesso"]:
            erro = resultado["erro"]
            if "404" in str(erro):
                log.info(f"  [{nome_filial}] NSU {nsu} → fila vazia. Fim.")
            else:
                log.error(f"  [{nome_filial}] NSU {nsu} → erro: {erro}")
            break

        dados  = resultado["dados"]
        status = dados.get("StatusProcessamento")
        log.info(f"  [{nome_filial}] NSU {nsu} → {status}")

        if status == "NENHUM_DOCUMENTO_LOCALIZADO":
            log.info(f"  [{nome_filial}] Fila vazia. Fim.")
            break
        if status == "REJEICAO":
            log.warning(f"  [{nome_filial}] Rejeição: {dados.get('Erros', [])}")
            break

        lote      = dados.get("LoteDFe") or []
        maior_nsu = nsu

        for doc in lote:
            nsu_doc = doc.get("NSU") or 0
            if nsu_doc > maior_nsu:
                maior_nsu = nsu_doc

            tipo    = doc.get("TipoDocumento", "DESCONHECIDO")
            chave   = doc.get("ChaveAcesso") or "sem_chave"
            xml_b64 = doc.get("ArquivoXml")

            log.info(f"    NSU {nsu_doc} | {tipo} | {chave[:20]}...")

            if not xml_b64:
                continue

            try:
                xml      = descompactar(xml_b64)
                nome_arq = f"{tipo}_{nsu_doc}_{chave[:20]}.xml"

                if tipo == "NFSE":
                    competencia = competencia_do_xml(xml, "NFSE")
                    cstat       = cstat_nfse(xml)
                    eh_subst    = (cstat == "101") or tem_chsubstda(xml)
                    # Notas substituta ficam na raiz — são válidas
                    pasta   = pasta_destino(EMPRESA, nome_filial, competencia)
                    caminho = os.path.join(pasta, nome_arq)
                    with open(caminho, "w", encoding="utf-8") as f:
                        f.write(xml)
                    cont["nfse"] += 1
                    if eh_subst:
                        cont["substituta"] += 1
                        log.info(f"      → {caminho} [SUBSTITUTA]")
                    else:
                        log.info(f"      → {caminho}")

                elif tipo == "EVENTO":
                    tag_ev, chave_nota = tag_e_chave_evento(xml)
                    descr, situacao, sub, cancela = info_evento(tag_ev)

                    if tag_ev and tag_ev not in __import__("eventos_config").EVENTOS:
                        log.warning(f"    ⚠ Evento desconhecido: {tag_ev} — salvo na raiz para investigação")
                        cont["desconhecidos"] += 1

                    if sub and chave_nota:
                        # Evento de cancelamento → usa competência da nota original
                        comp_nota   = buscar_competencia_nota(EMPRESA, nome_filial, chave_nota)
                        competencia = comp_nota or competencia_do_xml(xml, "EVENTO")
                    else:
                        competencia = competencia_do_xml(xml, "EVENTO")

                    pasta   = pasta_destino(EMPRESA, nome_filial, competencia, sub)
                    caminho = os.path.join(pasta, nome_arq)
                    with open(caminho, "w", encoding="utf-8") as f:
                        f.write(xml)
                    cont["eventos"] += 1

                    if sub == "Cancelada":      cont["canceladas"]   += 1
                    elif sub == "Substituida":  cont["substituidas"] += 1
                    elif sub == "Rejeitadas":   cont["rejeitadas"]   += 1

                    log.info(f"      → {caminho} [{descr}]")

                else:
                    competencia = competencia_do_xml(xml, tipo)
                    pasta   = pasta_destino(EMPRESA, nome_filial, competencia)
                    caminho = os.path.join(pasta, nome_arq)
                    with open(caminho, "w", encoding="utf-8") as f:
                        f.write(xml)
                    cont["outros"] += 1
                    log.info(f"      → {caminho}")

            except Exception as e:
                log.warning(f"    Falha NSU {nsu_doc}: {e}")
                cont["erros"] += 1

        nsu = maior_nsu + 1
        estado[cnpj] = nsu
        salvar_estado(estado)
        time.sleep(0.3)

    return cont


# ── Main ───────────────────────────────────────────────────
def main():
    connector = DistribuicaoConnector(
        pfx_path=os.getenv("NFSE_A1_PATH"),
        pfx_password=os.getenv("NFSE_A1_PASSWORD"),
        base_url=os.getenv("NFSE_CONTRIBUINTE_URL"),
    )
    estado = carregar_estado()
    tot = {k: 0 for k in ["nfse","substituta","eventos","canceladas",
                           "substituidas","rejeitadas","outros","erros","desconhecidos"]}

    log.info("=" * 60)
    log.info(f"Iniciando varredura — {EMPRESA}")
    log.info(f"Estrutura: xmls/{EMPRESA}/<filial>/YYYY-MM/[Cancelada|Substituida|Rejeitadas]/")
    log.info("=" * 60)

    for cnpj, nome_filial in CNPJS_EMPRESA.items():
        nsu_ini = estado.get(cnpj, 0)
        log.info(f"\n[{nome_filial}] CNPJ {cnpj} — NSU inicial: {nsu_ini}")
        r = varrer_cnpj(connector, cnpj, nome_filial, nsu_ini, estado)
        for k in tot: tot[k] += r[k]
        log.info(
            f"  [{nome_filial}] "
            f"{r['nfse']} NFS-e ({r['substituta']} substitutas) | "
            f"{r['eventos']} eventos: "
            f"{r['canceladas']} cancel. | {r['substituidas']} subst. | {r['rejeitadas']} rejeit. | "
            f"{r['desconhecidos']} desconhecidos | {r['erros']} erros"
        )
        time.sleep(0.5)

    log.info("\n" + "=" * 60)
    log.info("Varredura concluída.")
    log.info(f"  NFS-e normais  : {tot['nfse'] - tot['substituta']}")
    log.info(f"  NFS-e substitutas: {tot['substituta']}")
    log.info(f"  Eventos total  : {tot['eventos']}")
    log.info(f"    Canceladas   : {tot['canceladas']}")
    log.info(f"    Substituídas : {tot['substituidas']}")
    log.info(f"    Rejeitadas   : {tot['rejeitadas']}")
    log.info(f"  Desconhecidos  : {tot['desconhecidos']}  ← investigar se > 0")
    log.info(f"  Erros          : {tot['erros']}")
    log.info(f"  State          : {STATE_FILE}")
    log.info("=" * 60)

if __name__ == "__main__":
    main()
