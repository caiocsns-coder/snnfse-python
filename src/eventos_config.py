"""
Configuração centralizada de eventos NFS-e.
Fonte: Anexo II - Leiaute RN PedRegEvt Evt SNNFSe v1-01-20260122
       + XMLs reais de produção Elecnor

Estrutura de cada entrada:
    "tag": (
        "Descrição legível",    # conforme xDesc do XML
        "Situação CSV/Excel",   # valor da coluna Situação na nota vinculada
        "subpasta",             # "" = raiz | "Cancelada" | "Substituida" | "Rejeitadas"
        cancela_nota,           # True = efeito definitivo de cancelamento
    )

Para adicionar novos eventos: incluir no dicionário EVENTOS abaixo.
"""

EVENTOS = {

    # ══════════════════════════════════════════════════════
    # CATEGORIA 1 — CANCELAMENTOS
    # ══════════════════════════════════════════════════════

    # Cancelamento simples — autor: Emitente (01)
    "e101101": (
        "Cancelamento de NFS-e",
        "Cancelada",
        "Cancelada",
        True,
    ),

    # Cancelamento por substituição — autor: Município Emissor (05)
    # Contém <chSubstituta> com a chave da nota que substituiu
    "e105102": (
        "Cancelamento de NFS-e por Substituição",
        "Cancelada",
        "Substituida",
        True,
    ),

    # Solicitação de análise fiscal — autor: Emitente (01)
    # Pedido pendente; pode ser deferido (e105104) ou indeferido (e105105)
    "e101103": (
        "Solicitação de Análise Fiscal para Cancelamento",
        "Cancelada",
        "Cancelada",
        False,   # não é definitivo — aguarda decisão municipal
    ),

    # Cancelamento deferido por análise fiscal — autor: Município Emissor (05)
    # Mesmo efeito do e101101
    "e105104": (
        "Cancelamento de NFS-e Deferido por Análise Fiscal",
        "Cancelada",
        "Cancelada",
        True,
    ),

    # Cancelamento indeferido por análise fiscal — autor: Município Emissor (05)
    # NÃO cancela a nota — rejeita o pedido de cancelamento
    "e105105": (
        "Cancelamento de NFS-e Indeferido por Análise Fiscal",
        "Gerada",   # nota continua válida
        "",
        False,
    ),

    # ══════════════════════════════════════════════════════
    # CATEGORIA 2 — MANIFESTAÇÕES
    # ══════════════════════════════════════════════════════

    # Confirmação do Prestador — autor: Prestador (02)
    "e202201": (
        "Manifestação - Confirmação do Prestador",
        "Gerada",
        "",
        False,
    ),

    # Confirmação do Tomador — autor: Tomador (03)
    "e203202": (
        "Manifestação - Confirmação do Tomador",
        "Gerada",
        "",
        False,
    ),

    # Confirmação do Intermediário — autor: Intermediário (04)
    "e204203": (
        "Manifestação - Confirmação do Intermediário",
        "Gerada",
        "",
        False,
    ),

    # Confirmação Tácita — autor: Município Emissor (05)
    "e205204": (
        "Manifestação - Confirmação Tácita",
        "Gerada",
        "",
        False,
    ),

    # Rejeição do Prestador — autor: Prestador (02)
    "e202205": (
        "Manifestação - Rejeição do Prestador",
        "Cancelada",
        "Rejeitadas",
        False,
    ),

    # Rejeição do Tomador — autor: Tomador (03)
    "e203206": (
        "Manifestação - Rejeição do Tomador",
        "Cancelada",
        "Rejeitadas",
        False,
    ),

    # Rejeição do Intermediário — autor: Intermediário (04)
    "e204207": (
        "Manifestação - Rejeição do Intermediário",
        "Cancelada",
        "Rejeitadas",
        False,
    ),

    # Anulação da Rejeição — autor: Município Emissor (05)
    # Reverte um evento de rejeição anterior — nota volta a ser válida
    "e205208": (
        "Manifestação - Anulação da Rejeição",
        "Gerada",   # nota volta a ser válida
        "",
        False,
    ),

    # ══════════════════════════════════════════════════════
    # CATEGORIA 3 — OFÍCIOS (emitidos pelo município)
    # ══════════════════════════════════════════════════════

    # Cancelamento por ofício — autor: Município Emissor (05)
    # Mesmo efeito do e101101, mas sem solicitação do contribuinte
    "e305101": (
        "Cancelamento de NFS-e por Ofício",
        "Cancelada",
        "Cancelada",
        True,
    ),

    # Bloqueio por ofício — autor: Município Emissor (05)
    # Bloqueia temporariamente um tipo de evento; não cancela a nota
    "e305102": (
        "Bloqueio de NFS-e por Ofício",
        "Gerada",
        "",
        False,
    ),

    # Desbloqueio por ofício — autor: Município Emissor (05)
    "e305103": (
        "Desbloqueio de NFS-e por Ofício",
        "Gerada",
        "",
        False,
    ),

    # ══════════════════════════════════════════════════════
    # NOVOS EVENTOS
    # Adicionar aqui novos códigos identificados em produção.
    # Formato: "eXXXXXX": ("Descrição", "Situação", "subpasta", cancela_nota)
    # ══════════════════════════════════════════════════════
}


def info_evento(tag: str) -> tuple:
    """
    Retorna (descricao, situacao, subpasta, cancela_nota) para a tag do evento.
    Tags desconhecidas são salvas na raiz com situação genérica para investigação.
    """
    if tag in EVENTOS:
        return EVENTOS[tag]
    return (
        f"Evento não mapeado ({tag})",
        f"Evento Desconhecido ({tag})",
        "",      # raiz — não move sem certeza do efeito
        False,
    )


def subpasta_evento(tag: str) -> str:
    return info_evento(tag)[2]

def situacao_evento(tag: str) -> str:
    return info_evento(tag)[1]

def cancela_nota(tag: str) -> bool:
    return info_evento(tag)[3]
