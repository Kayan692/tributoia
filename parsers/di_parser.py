"""
Parser universal para DI e DUIMP da Receita Federal do Brasil.

ABORDAGEM: busca pelo NOME DO CAMPO e captura o numero ao lado.
Nao busca por padrao de valor — qualquer DI com o mesmo campo vai funcionar.
Suporta PDFs digitais e escaneados (OCR).
"""
import re
import pdfplumber
from dataclasses import dataclass, field
from typing import Tuple
from collections import Counter


@dataclass
class DIData:
    di_number:         str   = ""
    register_date:     str   = ""
    doc_type:          str   = "DI"
    importador_cnpj:   str   = ""
    importador_nome:   str   = ""
    uf_desembaraco:    str   = ""
    recinto:           str   = ""
    ncm:               str   = ""
    produto_desc:      str   = ""
    ncms_lista:        list  = field(default_factory=list)
    exportador_nome:   str   = ""
    exportador_pais:   str   = ""
    incoterm:          str   = ""
    vmle_usd:          float = 0.0
    frete_usd:         float = 0.0
    seguro_usd:        float = 0.0
    vmld_usd:          float = 0.0
    moeda:             str   = "USD"
    taxa_cambio:       float = 0.0
    ii:                float = 0.0
    ipi:               float = 0.0
    pis:               float = 0.0
    cofins:            float = 0.0
    antidumping:       float = 0.0
    siscomex:          float = 0.0
    afrmm:             float = 0.0
    icms_aliq:         float = 0.0
    icms_base_doc:     float = 0.0
    icms_valor_doc:    float = 0.0
    multiplas_adicoes: bool  = False
    adicoes_count:     int   = 1
    alerts:            list  = field(default_factory=list)
    raw_text:          str   = ""


def _clean_num(s):
    """BR number to float. '1.234,56'->1234.56  '5,22880'->5.228  '124.952,05'->124952.05"""
    if not s:
        return 0.0
    s = str(s).strip()
    for tok in ["R$","USD","US$","CNY","EUR","BRL","%"]:
        s = s.replace(tok,"")
    s = s.strip()
    if not s:
        return 0.0
    has_dot   = "." in s
    has_comma = "," in s
    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".","").replace(",",".")
        else:
            s = s.replace(",","")
    elif has_comma:
        parts = s.split(",")
        if len(parts)==2 and len(parts[0])<=5:
            s = s.replace(",",".")
        else:
            s = s.replace(",","")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _f(pattern, text, group=1):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(group).strip() if m else ""


def _lookup(text, patterns):
    """
    Tenta cada (regex, estrategia) em ordem.
    Estrategias:
      direct  -> group(1)
      second  -> group(2)
      sum_all -> soma todas as ocorrencias de group(1)
    """
    for pat, strategy in patterns:
        if strategy == "sum_all":
            matches = re.findall(pat, text, re.IGNORECASE | re.MULTILINE)
            if matches:
                total = sum(_clean_num(v) for v in matches)
                if total > 0:
                    return total
        elif strategy == "second":
            m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
            if m and m.lastindex and m.lastindex >= 2:
                val = _clean_num(m.group(2))
                if val > 0:
                    return val
        else:  # direct
            m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
            if m:
                val = _clean_num(m.group(1))
                if val > 0:
                    return val
    return 0.0


# ============================================================
# MAPA DE CAMPOS: (label_regex, estrategia)
# Captura o NOME DO CAMPO e pega o numero ao lado.
# Nao depende do valor — funciona em qualquer DI/DUIMP.
# ============================================================
FIELDS = {

    "taxa_cambio": [
        # "220 - USD - DOLAR DOS EUA: R$ 5,57370"   (DI padrao e DUIMP)
        (r"220\s*[-]\s*USD[^:]*:\s*R\$\s*([\d.,]+)",            "direct"),
        # "TX. DOLAR : 5,22880"                      (DI Santos/outros despachantes)
        (r"TX\.\s*DOLAR\s*:\s*([\d.,]+)",                        "direct"),
        # "TX. FRETE USD : 5,2288000"
        (r"TX\.\s*FRETE\s+USD\s*:\s*([\d.,]+)",                  "direct"),
        # "TX. SEGURO USD : 5,22880"
        (r"TX\.\s*SEGURO\s+USD\s*:\s*([\d.,]+)",                 "direct"),
        # "DOLAR DOS EUA: R$ 5,16820"
        (r"DOLAR\s+DOS?\s+EUA[:\s]+R\$\s*([\d.,]+)",             "direct"),
    ],

    "vmle_usd": [
        # "VMLE: DOLAR DOS ESTADOS UNIDOS  31.020,41"
        (r"VMLE\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$",               "direct"),
        # "VMLE: USD 31.020,41"  (DUIMP)
        (r"VMLE\s+USD\s*:\s*([\d.,]+)",                          "direct"),
        # "VMLE: USD 31.020,41 = R$ ..."
        (r"VMLE\s*:\s*USD\s+([\d.,]+)",                          "direct"),
    ],

    "frete_usd": [
        # "Frete: DOLAR DOS EUA 397,00"
        (r"Frete\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$",              "direct"),
        # "FRETE USD: 1.862,50"  (DUIMP)
        (r"FRETE\s+USD\s*:\s*([\d.,]+)",                         "direct"),
        # "FRETE EXTERNO: USD 1.740,00"
        (r"FRETE\s+EXTERNO\s*:\s*USD\s*([\d.,]+)",               "direct"),
        # "FRETE ORIGEM / FRONTEIRA...........US$ 1.740,00"
        (r"FRETE\s+ORIGEM\s*/\s*FRONTEIRA[.\s]+US\$\s*([\d.,]+)","direct"),
        # "TOTAL FRETE R$: 2.075,83 USD 397,00"  -> pega o USD
        (r"TOTAL\s+FRETE\s+R\$\s*:\s*[\d.,]+\s+USD\s+([\d.,]+)", "direct"),
    ],

    "seguro_usd": [
        # "Seguro: DOLAR DOS EUA 50,00"
        (r"Seguro\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$",             "direct"),
        # "SEGURO USD: 110,91"  (DUIMP)
        (r"SEGURO\s+USD\s*:\s*([\d.,]+)",                        "direct"),
        # "TOTAL SEGURO R$: 261,44 USD 50,00"
        (r"TOTAL\s+SEGURO\s+R\$\s*:\s*[\d.,]+\s+USD\s+([\d.,]+)","direct"),
    ],

    "vmld_usd": [
        # "VMLD: DOLAR DOS ESTADOS UNIDOS  23.896,78"
        (r"VMLD\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$",               "direct"),
        # "VMLD: 32.760,41"
        (r"VMLD\s*:\s*([\d.,]+)",                                 "direct"),
    ],

    # Valor Aduaneiro R$ — usado como fallback para calcular cambio implicito
    "va_brl": [
        # "VALOR ADUANEIRO R$: 124.952,05"  (Santos)
        (r"VALOR\s+ADUANEIRO\s+R\$\s*:\s*([\d.,]+)",              "direct"),
        # "VALOR ADUANEIRO ........... R$ 182.596,69"  (padrao)
        (r"VALOR\s+ADUANEIRO\s*[.\s]+R\$\s*([\d.,]+)",            "direct"),
        # "VALOR ADUANEIRO . .. R$ 183.853,55"  (DUIMP)
        (r"VALOR\s+ADUANEIRO\s*\.\s*\.\s*R\$\s*([\d.,]+)",        "direct"),
    ],

    # ── TRIBUTOS ─────────────────────────────────────────────────────────────
    # Pagina 1 sumario: "I.I.:  0,00  22.558,25" -> col Suspenso / Recolhido
    # Dados Complementares: "II R$: 22.558,25"
    # DUIMP OCR: "IN. . R$ 36.770,71"

    "ii": [
        (r"^II\s+R\$\s*:\s*([\d.,]+)",                            "direct"),
        (r"^I\.I\.\s*:\s*([\d.,]+)\s+([\d.,]+)",                  "second"),
        (r"\bIN\b[.\s]+R\$\s*([\d.,]+)",                          "direct"),
        (r"II\s*[.\s]+R\$\s*([\d.,]+)",                           "direct"),
    ],

    "ipi": [
        (r"^IPI\s+R\$\s*:\s*([\d.,]+)",                           "direct"),
        (r"^I\.P\.I\.\s*:\s*([\d.,]+)\s+([\d.,]+)",               "second"),
        (r"IPI\s*[.\s]+R\$\s*([\d.,]+)",                          "direct"),
    ],

    "pis": [
        (r"^PIS\s+R\$\s*:\s*([\d.,]+)",                           "direct"),
        (r"^Pis/Pasep\s*:\s*([\d.,]+)\s+([\d.,]+)",               "second"),
        (r"PIS\s*[.\s]+R\$\s*([\d.,]+)",                          "direct"),
    ],

    "cofins": [
        (r"^COFINS\s+R\$\s*:\s*([\d.,]+)",                        "direct"),
        (r"^Cofins\s*:\s*([\d.,]+)\s+([\d.,]+)",                  "second"),
        (r"COFINS\s*[.\s]+R\$\s*([\d.,]+)",                       "direct"),
    ],

    "antidumping": [
        (r"^Direitos\s+Antidumping\s*:\s*([\d.,]+)\s+([\d.,]+)",  "second"),
        (r"ANTIDUMPING\s*[.\s]+R\$\s*([\d.,]+)",                  "direct"),
    ],

    # ── SISCOMEX ──────────────────────────────────────────────────────────────
    # Total: "TX. SISCOMEX R$: 331,62"
    # DI padrao: "TAXA UTILIZAÇÃO ........ R$ 154,23"
    # Por adicao: "TAXA SISCOMEX R$: 221,50" (soma todas)
    "siscomex": [
        (r"TX\.\s*SISCOMEX\s+R\$\s*:\s*([\d.,]+)",                "direct"),
        (r"TAXA\s+UTILIZA[CcÇç][AaÃã]O\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        (r"TAXA\s+SISCOMEX\s+R\$\s*:\s*([\d.,]+)",                "sum_all"),
    ],

    # ── AFRMM ─────────────────────────────────────────────────────────────────
    # "A.F.R.M.M: R$: 208,06"
    # "AFRMM (MARINHA MERCANTE) ....... R$ 807,01"
    # Por adicao: "A.F.R.M.M: R$: 154,41" (soma todas)
    "afrmm": [
        (r"A\.F\.R\.M\.M\s*:\s*R\$\s*:\s*([\d.,]+)",              "direct"),
        (r"AFRMM\s*\(?MARINHA\s+MERCANTE\)?\s*[.\s]+R\$\s*([\d.,]+)","direct"),
        (r"AFRMM\s*[.\s]+R\$\s*([\d.,]+)",                        "direct"),
        (r"A\.F\.R\.M\.M\s*R\$\s*:\s*([\d.,]+)",                  "sum_all"),
    ],

    # ── ICMS ──────────────────────────────────────────────────────────────────
    # Total: "VALOR ICMS R$: 34.538,81"
    # DI padrao: "ICMS A RECOLHER (6,00%) ........ R$ 13.034,43"
    # Por adicao: "ICMS A RECOLHER (18,00%): R$ 24.329,70"
    "icms_valor_doc": [
        (r"VALOR\s+ICMS\s+R\$\s*:\s*([\d.,]+)",                   "direct"),
        (r"ICMS\s+A\s+RECOLHER\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)","direct"),
        (r"ICMS\s+[Aa\xc0\xe0]\s+RECOLHER\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)","sum_all"),
        (r"ICMS\s+CALCULADO\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        (r"ICMS\s+CALCULADO\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)",   "direct"),
    ],

    "icms_base_doc": [
        (r"BASE\s+ICMS\s+FINAL\s+R\$\s*:\s*([\d.,]+)",            "direct"),
        (r"BASE\s+C[AaÁá]LCULO\s+ICMS\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)","direct"),
        (r"BASE\s+CALCULO\s+ICMS\s+R\$\s*:\s*([\d.,]+)",          "direct"),
    ],
}

ICMS_ALIQ_PATS = [
    r"ICMS\s+CALCULADO\s*\(([\d.,]+)%?\)",
    r"ICMS\s+A\s+RECOLHER\s*\(([\d.,]+)%?\)",
    r"ICMS\s+[Aa\xc0\xe0]\s+RECOLHER\s*\(([\d.,]+)%?\)",
    r"BASE\s+C[AaÁá]LCULO\s+ICMS\s*\(([\d.,]+)%?\)",
]


def _extract_digital(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def _extract_ocr(pdf_path):
    try:
        from pdf2image import convert_from_path
        import pytesseract
        pages = convert_from_path(pdf_path, dpi=200)
        return "\n".join(pytesseract.image_to_string(p, lang="por") for p in pages)
    except Exception:
        return ""


def _get_text(pdf_path):
    text = _extract_digital(pdf_path)
    if len(text.strip().replace("\n","")) < 100:
        return _extract_ocr(pdf_path), True
    return text, False


def _parse(text, data):

    # Tipo e numero
    duimp = _f(r"(\d{2}BR\d{10}-\d)", text)
    if duimp:
        data.di_number = duimp
        data.doc_type  = "DUIMP"
    else:
        di = _f(r"Declara[cç][aã]o[:\s]+(\d{2}/\d{7}-\d)", text) or _f(r"(\d{2}/\d{7}-\d)", text)
        data.di_number = di
        data.doc_type  = "DI"
    if "DUIMP" in text.upper() and data.doc_type == "DI":
        data.doc_type = "DUIMP"
    if not data.di_number:
        data.alerts.append(("ERROR","Numero da DI/DUIMP nao encontrado"))

    # Data
    data.register_date = (
        _f(r"Data do Registro[:\s]+(\d{2}/\d{2}/\d{4})", text) or
        _f(r"gera[cç][aã]o do PDF[:\s]+(\d{2}/\d{2}/\d{4})", text) or
        _f(r"(\d{2}/\d{2}/\d{4})", text)
    )

    # Importador
    data.importador_cnpj = (
        _f(r"CNPJ do importador[:\s]*([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})", text) or
        _f(r"CNPJ[:\s]+([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})", text)
    )
    if data.importador_cnpj:
        m = re.search(re.escape(data.importador_cnpj) + r"\s+([A-Z\xc0-\xff][^\n]+)", text)
        if m:
            data.importador_nome = m.group(1).strip()
    if not data.importador_nome:
        data.importador_nome = _f(r"Nome do importador[:\s]*\n?([A-Z\xc0-\xff][^\n]{5,80})", text)

    # Todos os campos via FIELD_MAP
    data.taxa_cambio    = _lookup(text, FIELDS["taxa_cambio"])
    data.vmle_usd       = _lookup(text, FIELDS["vmle_usd"])
    data.frete_usd      = _lookup(text, FIELDS["frete_usd"])
    data.seguro_usd     = _lookup(text, FIELDS["seguro_usd"])
    data.vmld_usd       = _lookup(text, FIELDS["vmld_usd"])
    if data.vmld_usd == 0 and data.vmle_usd > 0:
        data.vmld_usd = data.vmle_usd + data.frete_usd + data.seguro_usd

    # Fallback cambio
    if data.taxa_cambio == 0:
        va_brl = _lookup(text, FIELDS["va_brl"])
        if va_brl > 0 and data.vmld_usd > 0:
            data.taxa_cambio = va_brl / data.vmld_usd
            data.alerts.append(("INFO",
                f"Cambio calculado: VA(R${va_brl:,.2f}) / VMLD(USD{data.vmld_usd:,.2f}) = R${data.taxa_cambio:.4f}"))
        else:
            data.alerts.append(("ERROR","Taxa de cambio nao encontrada"))

    if data.vmld_usd == 0:
        data.alerts.append(("ERROR","VMLD nao encontrado"))

    # Tributos
    data.ii          = _lookup(text, FIELDS["ii"])
    data.ipi         = _lookup(text, FIELDS["ipi"])
    data.pis         = _lookup(text, FIELDS["pis"])
    data.cofins      = _lookup(text, FIELDS["cofins"])
    data.antidumping = _lookup(text, FIELDS["antidumping"])
    data.siscomex    = _lookup(text, FIELDS["siscomex"])
    data.afrmm       = _lookup(text, FIELDS["afrmm"])

    if data.pis == 0 and data.cofins == 0:
        data.alerts.append(("WARN","PIS e COFINS zerados — verificar"))

    # ICMS
    data.icms_valor_doc = _lookup(text, FIELDS["icms_valor_doc"])
    data.icms_base_doc  = _lookup(text, FIELDS["icms_base_doc"])

    # Aliquota ICMS: usa ICMS CALCULADO (taxa nominal) com prioridade
    # Ignora ICMS A RECOLHER (%) pois pode ser percentual de pagamento, nao a aliquota
    aliq_list_calc = []
    for m in re.finditer(r'ICMS\s+CALCULADO\s*\(([\d.,]+)%?\)', text, re.IGNORECASE):
        try:
            aliq_list_calc.append(float(m.group(1).replace(",",".")))
        except ValueError:
            pass
    if aliq_list_calc:
        data.icms_aliq = Counter(aliq_list_calc).most_common(1)[0][0] / 100.0
    else:
        # Fallback: BASE CALCULO ICMS (X%)
        for m in re.finditer(r'BASE\s+C[AaÁá]LCULO\s+ICMS\s*\(([\d.,]+)%?\)', text, re.IGNORECASE):
            try:
                aliq_list_calc.append(float(m.group(1).replace(",",".")))
            except ValueError:
                pass
        if aliq_list_calc:
            data.icms_aliq = Counter(aliq_list_calc).most_common(1)[0][0] / 100.0

    # Aliquota ponderada real: aplica apenas quando ICMS e recolhimento integral
    # (nao aplica quando parte do ICMS e isento/suspenso — ex: DUIMP com ICMS parcial)
    if data.icms_valor_doc > 0 and data.icms_base_doc > 0:
        pond = data.icms_valor_doc / data.icms_base_doc
        # So usa ponderada se a diferenca e pequena (max 3pp) — indica aliquotas mistas
        # Se diferenca e grande (ex: 18% nominal vs 8.8% pond), e ICMS parcial — mantem nominal
        if 0 < abs(pond - data.icms_aliq) <= 0.03:
            data.icms_aliq = pond
            data.alerts.append(("INFO",
                f"Aliquota ponderada ({pond*100:.2f}%) — DI com aliquotas mistas"))
        elif abs(pond - data.icms_aliq) > 0.03:
            data.alerts.append(("INFO",
                f"ICMS parcialmente isento/suspenso — aliquota nominal {data.icms_aliq*100:.1f}% mantida"))

    if data.icms_aliq == 0:
        data.alerts.append(("ERROR","Aliquota ICMS nao encontrada"))

    # NCMs / produto
    ncms = re.findall(r"NCM\s+([\d.]{8,11})", text, re.IGNORECASE)
    if ncms:
        data.ncm       = ncms[0].strip()
        data.ncms_lista = list(dict.fromkeys(n.strip() for n in ncms))
    desc = _f(r"Descri[cç][aã]o\s+Detalhada[^\n]*\n(.+?)(?:Certificado|Imposto de Importa)", text)
    if not desc:
        desc = _f(r"NCM[:\s]+[\d.]+\s*[-]\s*([^\n]+)", text)
    data.produto_desc    = (desc or "").strip()[:200]
    data.exportador_nome = _f(r"Nome\s*:\s*([A-Z][^\n]{5,80})", text)
    data.exportador_pais = _f(r"Pa[ií]s de [Oo]rigem[:\s]+([^\n\r]+)", text)
    data.incoterm        = _f(r"INCOTERM[:\s]+([A-Z]{3})", text)

    # UF
    uf_map = {
        "FOZ DO IGUACU":"PR","FOZ DO IGUA\xc7U":"PR",
        "SANTOS":"SP","S\xc3O PAULO":"SP","SAO PAULO":"SP",
        "VIRACOPOS":"SP","GUARULHOS":"SP",
        "PARANAGU\xc1":"PR","PARANAGUA":"PR",
        "RIO DE JANEIRO":"RJ","GAL\xc3O":"RJ",
        "VIT\xd3RIA":"ES","VITORIA":"ES",
        "ITAJA\xcd":"SC","ITAJAI":"SC","NAVEGANTES":"SC",
        "MANAUS":"AM","FORTALEZA":"CE",
        "SALVADOR":"BA","RECIFE":"PE",
        "PORTO ALEGRE":"RS","URUGUAIANA":"RS","CURITIBA":"PR",
    }
    for key, uf in uf_map.items():
        if key in text.upper():
            data.uf_desembaraco = uf
            break

    # Adicoes
    qtd = _f(r"Quantidade\s+de\s+Adi[cç][oõ]es\s*:\s*(\d+)", text)
    if qtd:
        data.adicoes_count     = int(qtd)
        data.multiplas_adicoes = data.adicoes_count > 1
        if data.multiplas_adicoes:
            data.alerts.append(("INFO",
                f"DI com {data.adicoes_count} adicoes, {len(data.ncms_lista)} NCMs. "
                f"Valores totais consolidados foram usados."))

    if data.taxa_cambio and (data.taxa_cambio < 1.0 or data.taxa_cambio > 20.0):
        data.alerts.append(("WARN", f"Taxa de cambio suspeita: R${data.taxa_cambio:.4f}"))


def parse_pdf(pdf_path):
    data = DIData()
    text, used_ocr = _get_text(pdf_path)
    data.raw_text = text
    if used_ocr:
        data.alerts.append(("INFO","PDF escaneado — OCR ativado"))
    _parse(text, data)
    return data
