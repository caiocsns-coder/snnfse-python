# snnfse-python

Coleta, organização e análise de NFS-e diretamente da API oficial do Sistema Nacional NFS-e (gov.br), com geração de relatórios em Excel, CSV para Power BI e DANFSe em PDF.

---

## O que esse projeto faz

O [Sistema Nacional NFS-e](https://www.nfse.gov.br) disponibiliza uma API que centraliza **todas as notas fiscais de serviço** vinculadas a um CNPJ — emitidas e tomadas — em um único ponto de acesso. Este projeto consome essa API e automatiza o ciclo completo:

```
API ADN (gov.br)
  ↓  mTLS com certificado A1
ingest.py
  ↓  XMLs organizados por empresa/filial/competência
reorganizar_xmls.py
  ↓  classifica por eventos (canceladas, substituídas, rejeitadas)
  ↓  detecta cancelamentos retroativos → alertas_retroativos.csv
xml_para_excel.py
  ↓  Excel por filial/competência + CSV consolidado
Power BI
  ↓  reconciliação com ERP, alertas fiscais, cobertura por competência
```

---

## Funcionalidades

- **Coleta incremental** — controla o NSU por CNPJ e retoma de onde parou
- **Estrutura de pastas espelhada** — `xmls/<empresa>/<filial>/<competência>/[Cancelada|Substituida|Rejeitadas]/`
- **Classificação por eventos** — a situação real da nota é determinada pelos eventos vinculados, não pelo `cStat`
- **Detecção de cancelamentos retroativos** — identifica notas canceladas em competências diferentes da original com nível de urgência (Normal / Alta / Crítica)
- **Excel por competência** — uma aba para serviços tomados, outra para prestados, outra para eventos
- **CSV para Power BI** — 47 colunas com todos os campos fiscais e tributários
- **DANFSe em PDF** — geração local conforme leiaute da NT 008 (sem depender da API que será suspensa em 01/07/2026)

---

## Pré-requisitos

- Python 3.10+
- Certificado digital A1 (`.pfx`) com acesso ao Portal Nacional NFS-e
- Acesso à API do ADN: `https://adn.nfse.gov.br/contribuintes`

```bash
pip install requests python-dotenv reportlab qrcode openpyxl
```

---

## Configuração

### 1. Variáveis de ambiente

Copie o arquivo de exemplo e preencha com seus dados:

```bash
cp .env.example .env
```

```env
NFSE_A1_PATH=/caminho/para/certificado.pfx
NFSE_A1_PASSWORD=sua_senha
NFSE_CONTRIBUINTE_URL=https://adn.nfse.gov.br/contribuintes
NFSE_EMPRESA=minha_empresa
```

### 2. CNPJs da empresa

Copie o arquivo de exemplo e preencha com os CNPJs da sua empresa:

```bash
cp exemplos/cnpjs_exemplo.py src/cnpjs.py
```

Edite `src/cnpjs.py`:

```python
CNPJS_EMPRESA = {
    "00000000000100": "Empresa - Matriz",
    "00000000000200": "Empresa - Filial SP",
}
```

No `ingest.py`, importe o dicionário:

```python
from cnpjs import CNPJS_EMPRESA
```

---

## Como usar

### Coletar XMLs

```bash
python src/ingest.py
```

O script:
1. Lê o `state.json` para saber o último NSU processado por CNPJ
2. Baixa os documentos da fila do ADN (até 50 por chamada)
3. Salva os XMLs nas pastas corretas
4. Ao final, executa o `reorganizar_xmls.py` automaticamente

### Gerar Excel e CSV

```bash
python src/xml_para_excel.py
```

Gera:
- `excel_competencias/<filial>/NFS-e_YYYY-MM.xlsx`
- `dados/nfse_consolidado.csv`
- `dados/eventos_consolidado.csv`
- `dados/alertas_retroativos.csv`
- `dados/notas_nao_encontradas.csv`

### Gerar DANFSe PDF

```bash
# Um XML específico
python src/gerar_danfse.py xmls/empresa/filial/2026-01/NFSE_xxx.xml

# Pasta inteira (recursivo)
python src/gerar_danfse.py xmls/
```

---

## Estrutura de pastas gerada

```
xmls/
  minha_empresa/
    Empresa - Matriz/
      2026-01/
        NFSE_103_...xml          ← nota normal (cStat=100)
        NFSE_210_...xml          ← nota substituta (cStat=101)
        Cancelada/
          NFSE_105_...xml        ← nota cancelada (evento e101101)
          EVENTO_106_...xml      ← evento de cancelamento
        Substituida/
          NFSE_107_...xml        ← nota cancelada por substituição (evento e105102)
        Rejeitadas/
          EVENTO_108_...xml      ← rejeição do tomador (evento e203206)

dados/
  nfse_consolidado.csv           ← 47 colunas, todas as notas
  eventos_consolidado.csv        ← eventos vinculados
  alertas_retroativos.csv        ← cancelamentos fora da competência
  notas_nao_encontradas.csv      ← eventos sem nota correspondente em disco
```

---

## Lógica de eventos

A situação real da nota é determinada pelos **eventos vinculados**, não pelo campo `cStat`:

| cStat | Significado |
|---|---|
| `100` | NFS-e gerada normalmente |
| `101` | NFS-e emitida em substituição de outra (nota válida) |

| Evento | Situação | Pasta |
|---|---|---|
| `e101101` | Cancelada | `Cancelada/` |
| `e105102` | Cancelada por Substituição | `Substituida/` |
| `e101103` | Cancelada — Em Análise Fiscal | `Cancelada/` |
| `e105104` | Cancelamento Deferido | `Cancelada/` |
| `e203206` | Cancelada — Rejeição do Tomador | `Rejeitadas/` |
| `e305101` | Cancelada por Ofício | `Cancelada/` |

Novos eventos são adicionados em `src/eventos_config.py`.

---

## Power BI

Os arquivos `powerquery_*.txt` na pasta `docs/` contêm o código M completo para carregar os CSVs no Power BI, incluindo:

- Tipagem correta com locale `en-US` (evita erro de separador decimal)
- Filtro por papel da empresa (Tomadora / Intermediária)
- Chave composta para reconciliação com ERP: `CNPJ|NúmeroNota|DataEmissão`
- Coluna `Alerta` com os 4 tipos de inconsistência fiscal
- Medidas DAX para contagem, valores e score de risco por filial

---

## Aviso sobre o DANFSe

A API `/danfse/{chaveAcesso}` do Portal Nacional NFS-e será **suspensa em 01/07/2026** conforme a Nota Técnica Nº 008. O script `gerar_danfse.py` gera o PDF localmente a partir do XML, sem depender da API.


---

## Power BI

A pasta `docs/` contém os arquivos de código M (Power Query) e DAX prontos para uso:

| Arquivo | Descrição |
|---|---|
| `powerquery_nfse.txt` | Consulta principal — tipagem, filtros, chave de reconciliação |
| `powerquery_eventos.txt` | Eventos vinculados às notas |
| `powerquery_alertas.txt` | Cancelamentos retroativos com urgência |
| `powerquery_dcalendario.txt` | Tabela de datas otimizada |
| `dax_nfse.txt` | Medidas DAX completas com variantes Card |

### Chave de reconciliação com ERP

A consulta NFS-e gera a coluna `Chave_SISAP` no formato:

```
CNPJ (14 dígitos) | Número da nota | Data de emissão (YYYY-MM-DD)
Exemplo: 00000000000100|170|2026-01-08
```

Para cruzar com seu ERP, crie a mesma chave na consulta do ERP e faça o join:

```powerquery
MergeERP = Table.NestedJoin(
    NFS_e, {"Chave_SISAP"},
    MinhaConsultaERP, {"Chave_ERP"},
    "Info_ERP", JoinKind.LeftOuter
)
```

### Coluna Alerta

Após o merge, a coluna `Alerta` classifica cada nota em:

| Alerta | Situação |
|---|---|
| `OK` | Cancelada e não lançada, ou gerada e já no ERP |
| `Cancelada e paga — verificar` | Risco máximo: pagamento após cancelamento |
| `Cancelada e contabilizada — estornar` | Passivo indevido antes do pagamento |
| `Cancelada no ERP — baixar` | Entrada sem baixa no sistema |
| `Gerada — lançar no ERP` | Nota válida sem registro |

---

## Contribuições

Pull requests são bem-vindos. Para mudanças maiores, abra uma issue primeiro.

Ao contribuir, certifique-se de que nenhum dado real (CNPJs, XMLs, certificados) está sendo versionado.

---

## Licença

MIT
