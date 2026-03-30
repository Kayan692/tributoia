"""
Engine de calculo tributario.

LOGICA DO COMPARATIVO:
- Atual: le o ICMS A RECOLHER direto da DI (percentual e valor exatos do documento)
- Custo atual total = Subtotal + ICMS_a_recolher_doc
- Alagoas NF 4%:    NF = Subtotal / (1 - 0.04)   -> ICMS = NF - Subtotal
- Alagoas Dif 1.2%: NF = Subtotal / (1 - 0.012)  -> ICMS = NF - Subtotal
- Economia = Custo_atual - NF_Alagoas
"""
from dataclasses import dataclass, field
from typing import List
from parsers.di_parser import DIData


@dataclass
class CalcResult:
    # Base
    va_brl:           float = 0.0
    subtotal:         float = 0.0

    # Cenario Atual — direto da DI (ICMS A RECOLHER)
    icms_aliq_atual:  float = 0.0   # aliquota lida do documento
    icms_atual:       float = 0.0   # valor ICMS A RECOLHER do documento
    custo_atual:      float = 0.0   # subtotal + icms_atual

    # Cenario Alagoas NF 4%
    aliq_al_nf:       float = 0.04
    nf_al_nf:         float = 0.0
    icms_al_nf:       float = 0.0

    # Cenario Alagoas Diferencial 1.2%
    aliq_al_dif:      float = 0.012
    nf_al_dif:        float = 0.0
    icms_al_dif:      float = 0.0

    # Economia (custo_atual - NF_Alagoas)
    economia_vs_al_nf:  float = 0.0
    economia_vs_al_dif: float = 0.0

    # Reducao de ICMS
    reducao_icms_al_nf:  float = 0.0
    reducao_icms_al_dif: float = 0.0

    # Projecoes
    projections:     dict  = field(default_factory=dict)
    validation_ok:   bool  = True
    warnings:        List[str] = field(default_factory=list)


def calculate(data: DIData) -> CalcResult:
    r = CalcResult()

    # Valor Aduaneiro e Subtotal
    r.va_brl = data.vmld_usd * data.taxa_cambio
    r.subtotal = (r.va_brl + data.ii + data.ipi + data.pis +
                  data.cofins + data.antidumping + data.siscomex + data.afrmm)

    if r.subtotal <= 0:
        r.warnings.append("Subtotal federal zerado — verifique os dados")
        r.validation_ok = False
        return r

    # Cenario Atual: direto da DI
    r.icms_aliq_atual = data.icms_aliq
    r.icms_atual      = data.icms_valor_doc
    r.custo_atual     = r.subtotal + r.icms_atual

    if r.icms_atual == 0:
        r.warnings.append("Valor ICMS A RECOLHER nao encontrado no documento")
        r.validation_ok = False

    # Cenarios Alagoas: por dentro
    def por_dentro(aliq):
        nf = r.subtotal / (1.0 - aliq)
        return nf, nf - r.subtotal

    r.nf_al_nf,  r.icms_al_nf  = por_dentro(r.aliq_al_nf)
    r.nf_al_dif, r.icms_al_dif = por_dentro(r.aliq_al_dif)

    # Economia
    r.economia_vs_al_nf  = r.custo_atual - r.nf_al_nf
    r.economia_vs_al_dif = r.custo_atual - r.nf_al_dif

    # Reducao de ICMS
    r.reducao_icms_al_nf  = r.icms_atual - r.icms_al_nf
    r.reducao_icms_al_dif = r.icms_atual - r.icms_al_dif

    # Projecoes
    for n in [1, 5, 10, 20]:
        r.projections[n] = {
            "custo_atual_total":  r.custo_atual      * n,
            "custo_al_nf_total":  r.nf_al_nf         * n,
            "custo_al_dif_total": r.nf_al_dif        * n,
            "economia_nf":        r.economia_vs_al_nf  * n,
            "economia_dif":       r.economia_vs_al_dif * n,
            "icms_atual_total":   r.icms_atual        * n,
            "icms_al_nf_total":   r.icms_al_nf        * n,
            "icms_al_dif_total":  r.icms_al_dif       * n,
            "reducao_icms_dif":   r.reducao_icms_al_dif * n,
        }

    return r
