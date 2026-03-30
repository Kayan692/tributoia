"""
Parser universal para DI e DUIMP da Receita Federal do Brasil.

ABORDAGEM: extrai o texto bruto do PDF e envia para Claude (IA) que lê os 
campos pelo NOME, independente do formato ou despachante.
Funciona com qualquer DI/DUIMP sem necessidade de ajustes manuais.

Requer: variavel de ambiente ANTHROPIC_API_KEY configurada no sistema,
ou o usuario insere a chave na interface.
"""
import re
import json
import os
import pdfplumber
from dataclasses import dataclass, field
from typing import Tuple


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
    if not s: return 0.0
    s = str(s).strip()
    for tok in ["R$","USD","US$","CNY","EUR","BRL","%"]:
        s = s.replace(tok,"")
    s = s.strip()
    if not s: return 0.0
    has_dot, has_comma = "." in s, "," in s
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


def _extract_text(pdf_path: str) -> Tuple[str, bool]:
    """Extrai texto do PDF. Usa OCR se necessario."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    
    if len(text.strip().replace("\n","")) < 100:
        try:
            from pdf2image import convert_from_path
            import pytesseract
            pages = convert_from_path(pdf_path, dpi=200)
            text = "\n".join(pytesseract.image_to_string(p, lang="por") for p in pages)
            return text, True
        except Exception:
            pass
    return text, False


def _extract_with_ai(text: str, api_key: str) -> dict:
    """
    Envia o texto da DI para Claude e pede extracao estruturada dos campos.
    Retorna dicionario com todos os campos necessarios.
    """
    import urllib.request

    prompt = f"""Voce e um especialista em comercio exterior brasileiro.
Abaixo esta o texto extraido de uma DI (Declaracao de Importacao) ou DUIMP da Receita Federal.

Extraia EXATAMENTE os seguintes campos e retorne APENAS um JSON valido, sem explicacoes:

{{
  "di_number": "numero da declaracao (ex: 25/1583944-3 ou 26BR0000171403-2)",
  "register_date": "data do registro DD/MM/AAAA",
  "doc_type": "DI ou DUIMP",
  "importador_cnpj": "CNPJ do importador",
  "importador_nome": "razao social do importador",
  "exportador_nome": "nome do exportador/fabricante",
  "exportador_pais": "pais de origem",
  "ncm": "codigo NCM principal",
  "produto_desc": "descricao do produto (max 150 chars)",
  "incoterm": "incoterm (ex: FOB, CIF, FCA)",
  "uf_desembaraco": "UF de desembaraco (ex: SP, PR, RJ)",
  "vmle_usd": "valor VMLE em USD (numero)",
  "frete_usd": "valor do frete em USD (numero)",
  "seguro_usd": "valor do seguro em USD (numero)",
  "vmld_usd": "valor VMLD em USD (numero)",
  "taxa_cambio": "taxa de cambio R$/USD (numero com 4 decimais)",
  "ii": "valor II - Imposto de Importacao RECOLHIDO em R$ (numero)",
  "ipi": "valor IPI RECOLHIDO em R$ (numero)",
  "pis": "valor PIS/Pasep RECOLHIDO em R$ (numero)",
  "cofins": "valor COFINS RECOLHIDO em R$ (numero)",
  "antidumping": "valor direitos antidumping em R$ (numero, 0 se nao houver)",
  "siscomex": "valor taxa siscomex/taxa utilizacao em R$ (numero)",
  "afrmm": "valor AFRMM em R$ (numero, 0 se nao houver)",
  "icms_aliq": "aliquota ICMS em percentual (ex: 18.0 para 18%, 6.0 para 6%)",
  "icms_valor_doc": "valor ICMS A RECOLHER em R$ (numero)",
  "adicoes_count": "quantidade de adicoes (numero inteiro)"
}}

REGRAS IMPORTANTES:
- Para tributos com colunas Suspenso/Recolhido: use SEMPRE o valor da coluna RECOLHIDO
- Para ICMS: use o valor de ICMS A RECOLHER (nao o calculado se forem diferentes)
- Para aliquota ICMS: use a aliquota nominal do ICMS CALCULADO (ex: 18.0, nao 48.89)
- Taxa de cambio: procure por "220 - USD", "Tx Dolar", "TX. DOLAR" ou similar
- Se um campo nao existir, use 0 para numeros ou "" para texto
- Retorne APENAS o JSON, sem markdown, sem explicacoes

TEXTO DA DI:
{text[:6000]}"""

    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "messages": [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        },
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        content = result["content"][0]["text"].strip()
        # Remove markdown if present
        content = re.sub(r'^```json\s*', '', content)
        content = re.sub(r'\s*```$', '', content)
        return json.loads(content)


def _fallback_regex(text: str, data: DIData):
    """
    Parser de fallback por regex para quando a API nao esta disponivel.
    Cobre os formatos mais comuns ja mapeados.
    """
    from collections import Counter

    def f(pat, t, g=1):
        m = re.search(pat, t, re.IGNORECASE)
        return m.group(g).strip() if m else ""

    def num(s): return _clean_num(s)

    def lookup(text, patterns):
        for pat, strategy in patterns:
            if strategy == "sum_all":
                matches = re.findall(pat, text, re.IGNORECASE | re.MULTILINE)
                if matches:
                    total = sum(num(v) for v in matches)
                    if total > 0: return total
            elif strategy == "second":
                m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
                if m and m.lastindex and m.lastindex >= 2:
                    val = num(m.group(2))
                    if val > 0: return val
            else:
                m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
                if m:
                    val = num(m.group(1))
                    if val > 0: return val
        return 0.0

    # Numero e data
    data.di_number = (f(r"(\d{2}BR\d{10}-\d)", text) or
                      f(r"Declara[cç][aã]o[:\s]+(\d{2}/\d{7}-\d)", text) or
                      f(r"(\d{2}/\d{7}-\d)", text))
    data.doc_type = "DUIMP" if "DUIMP" in text.upper() or re.search(r"\d{2}BR\d{10}", text) else "DI"
    data.register_date = (f(r"Data do Registro[:\s]+(\d{2}/\d{2}/\d{4})", text) or
                          f(r"(\d{2}/\d{2}/\d{4})", text))

    # Importador
    data.importador_cnpj = (
        f(r"CNPJ do importador[:\s]*([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})", text) or
        f(r"CNPJ[:\s]+([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})", text)
    )
    if data.importador_cnpj:
        m = re.search(re.escape(data.importador_cnpj) + r"\s+([A-Z][^\n]+)", text)
        if m: data.importador_nome = m.group(1).strip()

    # Cambio - todos os formatos conhecidos
    data.taxa_cambio = lookup(text, [
        (r"220\s*[-]\s*USD[^:]*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"Tx\s+Dolar\s*:\s*([\d.,]+)", "direct"),
        (r"TX\.\s*DOLAR\s*:\s*([\d.,]+)", "direct"),
        (r"TX\.\s*FRETE\s+USD\s*:\s*([\d.,]+)", "direct"),
        (r"DOLAR\s+DOS?\s+EUA[:\s]+R\$\s*([\d.,]+)", "direct"),
    ])

    # Valores USD
    data.vmld_usd = lookup(text, [
        (r"VMLD\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$", "direct"),
        (r"VMLD\s*:\s*([\d.,]+)", "direct"),
    ])
    data.vmle_usd = lookup(text, [
        (r"VMLE\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$", "direct"),
        (r"VMLE\s+USD\s*:\s*([\d.,]+)", "direct"),
    ])
    data.frete_usd = lookup(text, [
        (r"Frete\s*:\s*D[O\xd3]LAR[^\n]+?([\d.,]{4,})\s*$", "direct"),
        (r"FRETE\s+USD\s*:\s*([\d.,]+)", "direct"),
        (r"VALOR\s+FRETE\s*:\s*[\d.,]+\s+US\$\s*$", "direct"),
    ])
    if data.vmld_usd == 0 and data.vmle_usd > 0:
        data.vmld_usd = data.vmle_usd + data.frete_usd + data.seguro_usd

    # Fallback cambio
    if data.taxa_cambio == 0:
        va_brl = lookup(text, [
            (r"VALOR\s+ADUANEIRO\s+R\$\s*:\s*([\d.,]+)", "direct"),
            (r"Valor\s+Aduaneiro\s*:\s*R\$\s*([\d.,]+)", "direct"),
            (r"VALOR\s+ADUANEIRO\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        ])
        if va_brl > 0 and data.vmld_usd > 0:
            data.taxa_cambio = va_brl / data.vmld_usd
        else:
            data.alerts.append(("ERROR", "Taxa de cambio nao encontrada"))

    # Tributos
    data.ii = lookup(text, [
        (r"^II\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"^II\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"^I\.I\.\s*:\s*([\d.,]+)\s+([\d.,]+)", "second"),
        (r"\bIN\b[.\s]+R\$\s*([\d.,]+)", "direct"),
    ])
    data.ipi = lookup(text, [
        (r"^IPI\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"^IPI\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"^I\.P\.I\.\s*:\s*([\d.,]+)\s+([\d.,]+)", "second"),
    ])
    data.pis = lookup(text, [
        (r"^PIS\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"^PIS\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"^Pis/Pasep\s*:\s*([\d.,]+)\s+([\d.,]+)", "second"),
    ])
    data.cofins = lookup(text, [
        (r"^COFINS\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"^Cofins\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"^Cofins\s*:\s*([\d.,]+)\s+([\d.,]+)", "second"),
    ])
    data.siscomex = lookup(text, [
        (r"TX\.\s*SISCOMEX\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"TAXA\s+UTILIZA[CcÇç][AaÃã]O\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        (r"Taxa\s+Siscomex\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "sum_all"),
        (r"TAXA\s+SISCOMEX\s+R\$\s*:\s*([\d.,]+)", "sum_all"),
    ])
    data.afrmm = lookup(text, [
        (r"A\.F\.R\.M\.M\s*:\s*R\$\s*:\s*([\d.,]+)", "direct"),
        (r"^AFRMM\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"AFRMM\s*\(?MARINHA[^)]*\)?\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        (r"A\.F\.R\.M\.M\s*R\$\s*:\s*([\d.,]+)", "sum_all"),
    ])
    data.antidumping = lookup(text, [
        (r"Direitos\s+Antidumping\s*:\s*([\d.,]+)\s+([\d.,]+)", "second"),
        (r"ANTIDUMPING\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
    ])

    # ICMS
    data.icms_valor_doc = lookup(text, [
        (r"VALOR\s+ICMS\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"Valor\s+do\s+ICMS\s*:\s*R\$\s*([\d.,]+)", "direct"),
        (r"ICMS\s+A\s+RECOLHER\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
        (r"ICMS\s+[AÀ]\s+RECOLHER\s*\([^)]+\)\s*:\s*R\$\s*([\d.,]+)", "sum_all"),
        (r"ICMS\s+CALCULADO\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
    ])
    data.icms_base_doc = lookup(text, [
        (r"BASE\s+ICMS\s+FINAL\s+R\$\s*:\s*([\d.,]+)", "direct"),
        (r"BASE\s+C[AÁ]LCULO\s+ICMS\s*\([^)]+\)\s*[.\s]+R\$\s*([\d.,]+)", "direct"),
    ])

    # Aliquota ICMS — usa ICMS CALCULADO com prioridade
    aliq_list = []
    for m in re.finditer(r"ICMS\s+CALCULADO\s*\(([\d.,]+)%?\)", text, re.IGNORECASE):
        try: aliq_list.append(float(m.group(1).replace(",",".")))
        except: pass
    # Formato "x 18,0%" 
    for m in re.finditer(r"Base\s+de\s+Calculo\s+ICMS[^\n]+x\s*([\d.,]+)%", text, re.IGNORECASE):
        try: aliq_list.append(float(m.group(1).replace(",",".")))
        except: pass
    if not aliq_list:
        for m in re.finditer(r"BASE\s+C[AÁ]LCULO\s+ICMS\s*\(([\d.,]+)%?\)", text, re.IGNORECASE):
            try: aliq_list.append(float(m.group(1).replace(",",".")))
            except: pass

    if aliq_list:
        from collections import Counter
        data.icms_aliq = Counter(aliq_list).most_common(1)[0][0] / 100.0

    # Aliquota ponderada (DI com aliquotas mistas)
    if data.icms_valor_doc > 0 and data.icms_base_doc > 0:
        pond = data.icms_valor_doc / data.icms_base_doc
        if 0 < abs(pond - data.icms_aliq) <= 0.03:
            data.icms_aliq = pond

    if data.icms_aliq == 0:
        data.alerts.append(("ERROR", "Aliquota ICMS nao encontrada"))

    # NCM e produto
    ncms = re.findall(r"NCM\s+([\d.]{8,11})", text, re.IGNORECASE)
    if ncms:
        data.ncm = ncms[0].strip()
        data.ncms_lista = list(dict.fromkeys(n.strip() for n in ncms))
    desc = f(r"Descri[cç][aã]o\s+Detalhada[^\n]*\n(.+?)(?:Certificado|Imposto de Importa)", text)
    if not desc: desc = f(r"NCM[:\s]+[\d.]+\s*[-]\s*([^\n]+)", text)
    data.produto_desc = (desc or "").strip()[:200]
    data.exportador_nome = f(r"Nome\s*:\s*([A-Z][^\n]{5,80})", text)
    data.exportador_pais = f(r"Pa[ií]s de [Oo]rigem[:\s]+([^\n\r]+)", text)
    data.incoterm = f(r"INCOTERM[:\s]+([A-Z]{3})", text)

    # UF
    uf_map = {
        "FOZ DO IGUACU":"PR","SANTOS":"SP","SAO PAULO":"SP","VIRACOPOS":"SP",
        "GUARULHOS":"SP","PARANAGUA":"PR","RIO DE JANEIRO":"RJ","GALEAO":"RJ",
        "VITORIA":"ES","ITAJAI":"SC","NAVEGANTES":"SC","MANAUS":"AM",
        "FORTALEZA":"CE","SALVADOR":"BA","RECIFE":"PE","PORTO ALEGRE":"RS",
        "URUGUAIANA":"RS","CURITIBA":"PR",
    }
    text_up = text.upper()
    for key, uf in uf_map.items():
        if key in text_up:
            data.uf_desembaraco = uf
            break

    qtd = f(r"Quantidade\s+de\s+Adi[cç][oõ]es\s*:\s*(\d+)", text)
    if qtd:
        data.adicoes_count = int(qtd)
        data.multiplas_adicoes = data.adicoes_count > 1


def parse_pdf(pdf_path: str, api_key: str = None) -> DIData:
    """
    Extrai dados de uma DI/DUIMP.
    
    Se api_key for fornecida, usa Claude (IA) para extrair os campos — 
    funciona com qualquer formato, qualquer despachante.
    
    Se api_key nao for fornecida, usa parser por regex como fallback.
    """
    data = DIData()
    text, used_ocr = _extract_text(pdf_path)
    data.raw_text = text

    if used_ocr:
        data.alerts.append(("INFO", "PDF escaneado — OCR ativado"))

    # Tenta API key do ambiente se nao fornecida
    if not api_key:
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")

    if api_key:
        try:
            fields = _extract_with_ai(text, api_key)

            # Popula o dataclass com o resultado da IA
            data.di_number        = str(fields.get("di_number", ""))
            data.register_date    = str(fields.get("register_date", ""))
            data.doc_type         = str(fields.get("doc_type", "DI"))
            data.importador_cnpj  = str(fields.get("importador_cnpj", ""))
            data.importador_nome  = str(fields.get("importador_nome", ""))
            data.exportador_nome  = str(fields.get("exportador_nome", ""))
            data.exportador_pais  = str(fields.get("exportador_pais", ""))
            data.ncm              = str(fields.get("ncm", ""))
            data.produto_desc     = str(fields.get("produto_desc", ""))[:200]
            data.incoterm         = str(fields.get("incoterm", ""))
            data.uf_desembaraco   = str(fields.get("uf_desembaraco", ""))
            data.vmle_usd         = float(fields.get("vmle_usd", 0) or 0)
            data.frete_usd        = float(fields.get("frete_usd", 0) or 0)
            data.seguro_usd       = float(fields.get("seguro_usd", 0) or 0)
            data.vmld_usd         = float(fields.get("vmld_usd", 0) or 0)
            data.taxa_cambio      = float(fields.get("taxa_cambio", 0) or 0)
            data.ii               = float(fields.get("ii", 0) or 0)
            data.ipi              = float(fields.get("ipi", 0) or 0)
            data.pis              = float(fields.get("pis", 0) or 0)
            data.cofins           = float(fields.get("cofins", 0) or 0)
            data.antidumping      = float(fields.get("antidumping", 0) or 0)
            data.siscomex         = float(fields.get("siscomex", 0) or 0)
            data.afrmm            = float(fields.get("afrmm", 0) or 0)
            data.icms_aliq        = float(fields.get("icms_aliq", 0) or 0) / 100.0
            data.icms_valor_doc   = float(fields.get("icms_valor_doc", 0) or 0)
            data.adicoes_count    = int(fields.get("adicoes_count", 1) or 1)
            data.multiplas_adicoes = data.adicoes_count > 1

            if data.vmld_usd == 0 and data.vmle_usd > 0:
                data.vmld_usd = data.vmle_usd + data.frete_usd + data.seguro_usd

            # Validacoes
            if data.taxa_cambio == 0:
                data.alerts.append(("ERROR", "Taxa de cambio nao encontrada"))
            if data.icms_aliq == 0:
                data.alerts.append(("ERROR", "Aliquota ICMS nao encontrada"))
            if data.vmld_usd == 0:
                data.alerts.append(("ERROR", "VMLD nao encontrado"))

            data.alerts.append(("INFO", "Campos extraidos via IA (Claude)"))
            return data

        except Exception as e:
            data.alerts.append(("WARN", f"IA indisponivel, usando parser padrao: {str(e)[:60]}"))

    # Fallback: parser por regex
    _fallback_regex(text, data)
    return data
