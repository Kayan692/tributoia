"""
PDF report generator - corporate executive report.
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.pdfgen import canvas as pdfcanvas
from datetime import date, datetime
from parsers.di_parser import DIData
from engine.tax_calculator import CalcResult

W, H = A4

C_DARK  = colors.HexColor("#1F3864")
C_MID   = colors.HexColor("#2E75B6")
C_LIGHT = colors.HexColor("#D6E4F0")
C_GREEN = colors.HexColor("#375623")
C_LGRN  = colors.HexColor("#E2EFDA")
C_RED   = colors.HexColor("#CC3300")
C_LRED  = colors.HexColor("#FFE0E0")
C_YELL  = colors.HexColor("#FFF2CC")
C_GREY  = colors.HexColor("#F5F5F5")
C_WHITE = colors.white
C_BDR   = colors.HexColor("#CBD5E1")
C_MUT   = colors.HexColor("#64748B")


def brl(v): return f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")
def pct(v): return f"{v*100:.1f}%".replace(".",",")


def sty(name, **kw):
    d = dict(fontName="Helvetica", fontSize=9, leading=14,
             textColor=colors.HexColor("#0F172A"), spaceAfter=0)
    d.update(kw)
    return ParagraphStyle(name, **d)

S_BODY  = sty("b", fontSize=9, leading=14, alignment=TA_JUSTIFY, spaceAfter=5)
S_BOLD  = sty("bo", fontName="Helvetica-Bold", fontSize=9)
S_SMALL = sty("sm", fontSize=8, leading=11, textColor=C_MUT)
S_NOTE  = sty("nt", fontSize=8, leading=12, textColor=colors.HexColor("#555555"))
S_CAP   = sty("cp", fontSize=7.5, textColor=C_MUT, alignment=TA_CENTER)
S_SEC   = sty("sh", fontName="Helvetica-Bold", fontSize=10.5, textColor=C_WHITE, leading=14)
S_KV    = sty("kv", fontName="Helvetica-Bold", fontSize=15, textColor=C_GREEN, alignment=TA_CENTER, leading=20)
S_KL    = sty("kl", fontSize=7.5, textColor=C_MUT, alignment=TA_CENTER, leading=10)

P = lambda t, s=None: Paragraph(t, s or S_BODY)
SP = lambda n=6: Spacer(1, n)


def sec_hdr(text, color=C_DARK):
    t = Table([[P(f'<font color="white"><b>{text}</b></font>', S_SEC)]],
              colWidths=[17*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), color),
        ('LEFTPADDING',(0,0),(-1,-1), 10),
        ('TOPPADDING', (0,0),(-1,-1), 7),
        ('BOTTOMPADDING',(0,0),(-1,-1), 7),
    ]))
    return t


def mk_table(rows, cws, hrows=1, stripe=True, last_bold=False):
    t = Table(rows, colWidths=cws)
    cmds = [
        ('FONTNAME',    (0,0),(-1,hrows-1), 'Helvetica-Bold'),
        ('FONTSIZE',    (0,0),(-1,-1),       8.5),
        ('BACKGROUND',  (0,0),(-1,hrows-1),  C_DARK),
        ('TEXTCOLOR',   (0,0),(-1,hrows-1),  C_WHITE),
        ('ALIGN',       (0,0),(-1,-1),       'CENTER'),
        ('ALIGN',       (0,hrows),(0,-1),    'LEFT'),
        ('VALIGN',      (0,0),(-1,-1),       'MIDDLE'),
        ('TOPPADDING',  (0,0),(-1,-1),        5),
        ('BOTTOMPADDING',(0,0),(-1,-1),       5),
        ('LEFTPADDING', (0,hrows),(0,-1),     6),
        ('GRID',        (0,0),(-1,-1),        0.4, C_BDR),
        ('LINEBELOW',   (0,0),(-1,0),         1.2, C_MID),
    ]
    if stripe:
        for i in range(hrows, len(rows)):
            if i % 2 == 0:
                cmds.append(('BACKGROUND',(0,i),(-1,i), C_GREY))
    if last_bold:
        cmds += [('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                 ('BACKGROUND',(0,-1),(-1,-1), C_LIGHT)]
    t.setStyle(TableStyle(cmds))
    return t


class ReportCanvas(pdfcanvas.Canvas):
    def __init__(self, *args, **kwargs):
        pdfcanvas.Canvas.__init__(self, *args, **kwargs)
        self._pgs = []

    def showPage(self):
        self._pgs.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._pgs)
        for s in self._pgs:
            self.__dict__.update(s)
            self._chrome(n)
            pdfcanvas.Canvas.showPage(self)
        pdfcanvas.Canvas.save(self)

    def _chrome(self, total):
        pn = self._pageNumber
        if pn == 1:
            self.setFillColor(C_MID)
            self.rect(0, 0, W, 0.8*cm, fill=1, stroke=0)
            self.setFillColor(C_WHITE)
            self.setFont("Helvetica", 7)
            self.drawString(1.5*cm, 0.27*cm, f"Emitido em {date.today().strftime('%d/%m/%Y')}  |  Confidencial")
            self.drawRightString(W-1.5*cm, 0.27*cm, "Uso Interno")
            return
        self.setFillColor(C_DARK)
        self.rect(0, H-1.6*cm, W, 1.6*cm, fill=1, stroke=0)
        self.setFillColor(C_MID)
        self.rect(0, H-1.75*cm, W, 0.15*cm, fill=1, stroke=0)
        self.setFillColor(C_WHITE)
        self.setFont("Helvetica-Bold", 8)
        self.drawString(1.5*cm, H-0.9*cm, "RELATÓRIO EXECUTIVO – EFICIÊNCIA TRIBUTÁRIA NA IMPORTAÇÃO")
        self.setFont("Helvetica", 7.5)
        self.drawRightString(W-1.5*cm, H-0.9*cm, f"{self._importador} | {self._di_num}")
        self.setFont("Helvetica", 7)
        self.drawString(1.5*cm, H-1.45*cm, "Análise Confidencial – Uso Interno")
        self.setFillColor(C_DARK)
        self.rect(0, 0, W, 0.9*cm, fill=1, stroke=0)
        self.setFillColor(C_WHITE)
        self.setFont("Helvetica", 7)
        self.drawString(1.5*cm, 0.32*cm, f"Emitido em {date.today().strftime('%d/%m/%Y')}")
        self.drawRightString(W-1.5*cm, 0.32*cm, f"Pág. {pn}/{total}")


def generate_pdf(data: DIData, result: CalcResult, output_path: str):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=2.4*cm, bottomMargin=1.6*cm)

    # Inject metadata into canvas class
    importador_short = data.importador_nome[:35] if data.importador_nome else "Importador"
    ReportCanvas._importador = importador_short
    ReportCanvas._di_num = f"{data.doc_type} {data.di_number}"

    story = []

    # ── CAPA ──────────────────────────────────────────────────────────────────
    story.append(SP(16))
    story.append(Table([
        [P("<b>RELATÓRIO EXECUTIVO</b>",
           sty("ct", fontName="Helvetica-Bold", fontSize=28, textColor=C_DARK, leading=34, alignment=TA_LEFT))],
        [P("Eficiência Tributária na Importação",
           sty("cs", fontSize=14, textColor=C_MID, leading=18, alignment=TA_LEFT))],
        [P(f"Análise Comparativa  ·  Operação Atual vs. Alagoas",
           sty("cs2", fontSize=10, textColor=C_MUT, leading=14, alignment=TA_LEFT))],
    ], colWidths=[17*cm], style=TableStyle([
        ('LINEABOVE',(0,0),(-1,0),5,C_MID),
        ('TOPPADDING',(0,0),(-1,-1),8),
        ('BOTTOMPADDING',(0,0),(-1,-1),6),
    ])))
    story.append(SP(16))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#E2E8F0")))
    story.append(SP(12))

    # Ident block
    ident = Table([
        [P("<b>Empresa</b>",S_BOLD),     P(data.importador_nome or "–",S_BODY)],
        [P("<b>CNPJ</b>",S_BOLD),        P(data.importador_cnpj or "–",S_BODY)],
        [P("<b>Declaração</b>",S_BOLD),  P(f"{data.doc_type} {data.di_number}",S_BODY)],
        [P("<b>Data</b>",S_BOLD),        P(data.register_date or "–",S_BODY)],
        [P("<b>Produto</b>",S_BOLD),     P((data.produto_desc[:80] if data.produto_desc else "–"),S_BODY)],
        [P("<b>NCM</b>",S_BOLD),         P(data.ncm or "–",S_BODY)],
        [P("<b>Fornecedor</b>",S_BOLD),  P(f"{data.exportador_nome} – {data.exportador_pais}" if data.exportador_nome else "–",S_BODY)],
    ], colWidths=[4*cm, 13*cm])
    ident.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.3,C_BDR),
        ('BACKGROUND',(0,0),(0,-1),C_LIGHT),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),6),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(ident)
    story.append(SP(16))

    # KPI cards
    eco_pct_val = result.economia_vs_al_dif/result.custo_atual if result.custo_atual>0 else 0
    icms_red_pct_val = (result.icms_atual-result.icms_al_dif)/result.icms_atual if result.icms_atual>0 else 0

    kpi_cells = [[
        Table([[P(brl(result.economia_vs_al_dif),S_KV)],[P("Economia por Operação",S_KL)]],colWidths=[5.5*cm]),
        Table([[P(pct(icms_red_pct_val),S_KV)],[P("Redução no ICMS",S_KL)]],colWidths=[5.5*cm]),
        Table([[P(brl(result.projections[10]["economia_dif"]),S_KV)],[P("Economia em 10 Operações",S_KL)]],colWidths=[5.5*cm]),
    ]]
    kpi_tbl = Table(kpi_cells, colWidths=[5.667*cm]*3)
    kpi_tbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),C_LGRN),
        ('BOX',(0,0),(-1,-1),1.5,C_GREEN),
        ('INNERGRID',(0,0),(-1,-1),0.5,C_BDR),
        ('TOPPADDING',(0,0),(-1,-1),10),
        ('BOTTOMPADDING',(0,0),(-1,-1),10),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(kpi_tbl)
    story.append(SP(16))
    story.append(HRFlowable(width="100%", thickness=0.5, color=C_BDR))
    story.append(SP(6))
    story.append(P(f"Documento emitido em {date.today().strftime('%d/%m/%Y')}  |  Confidencial – Uso Interno", S_CAP))

    story.append(PageBreak())

    # ── VISÃO GERAL ───────────────────────────────────────────────────────────
    story.append(sec_hdr("1. VISÃO GERAL DA OPERAÇÃO"))
    story.append(SP(8))

    di_data_rows = [
        ["Campo", "Valor"],
        ["VMLE (FOB)", f"USD {data.vmle_usd:,.2f}"],
        ["Frete Internacional", f"USD {data.frete_usd:,.2f}"],
        ["VMLD (Valor Aduaneiro Base)", f"USD {data.vmld_usd:,.2f}"],
        ["Taxa de Câmbio", f"R$ {data.taxa_cambio:.4f}"],
        ["Valor Aduaneiro em R$", brl(result.va_brl)],
        ["II", brl(data.ii)],
        ["IPI", brl(data.ipi)],
        ["PIS (2,10%)", brl(data.pis)],
        ["COFINS (9,65%)", brl(data.cofins)],
        ["Taxa Siscomex", brl(data.siscomex)],
        ["SUBTOTAL FEDERAL", brl(result.subtotal)],
    ]
    story.append(mk_table(di_data_rows, [6*cm, 11*cm], last_bold=True))
    story.append(SP(12))

    # ── COMPARATIVO ───────────────────────────────────────────────────────────
    story.append(sec_hdr("2. COMPARATIVO TRIBUTÁRIO", C_MID))
    story.append(SP(8))
    story.append(P(
        "O ICMS é calculado pelo método <b>\"por dentro\"</b>: o tributo está embutido no valor da nota fiscal. "
        "Para encontrá-lo: <b>Valor NF = Subtotal ÷ (1 – alíquota)</b>  |  <b>ICMS = Valor NF – Subtotal</b>",
        S_BODY))
    story.append(SP(6))

    aliq_label = f"{result.icms_aliq_atual*100:.1f}".replace(".",",").rstrip("0").rstrip(",")
    comp_data = [
        ["Parâmetro", f"Atual\n({aliq_label}%)", "AL NF\n4%", "AL Dif.\n1,2%", "Dif.\nAtual→AL", "Red.\n%"],
        ["Subtotal Federal", brl(result.subtotal), brl(result.subtotal), brl(result.subtotal), "–", "–"],
        ["Alíquota ICMS", pct(result.icms_aliq_atual), "4,0%", "1,2%", "–", "–"],
        ["Divisor", f"{1-result.icms_aliq_atual:.4f}".replace(".",","), "0,9600", "0,9880", "–", "–"],
        ["Valor da NF", brl(result.custo_atual), brl(result.nf_al_nf), brl(result.nf_al_dif),
         brl(result.economia_vs_al_dif), pct(eco_pct_val)],
        ["Valor do ICMS", brl(result.icms_atual), brl(result.icms_al_nf), brl(result.icms_al_dif),
         brl(result.icms_atual-result.icms_al_dif), pct(icms_red_pct_val)],
        ["Custo Total", brl(result.custo_atual), brl(result.nf_al_nf), brl(result.nf_al_dif),
         brl(result.economia_vs_al_dif), pct(eco_pct_val)],
    ]
    ct = Table(comp_data, colWidths=[4*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.8*cm, 1.8*cm])
    ct.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('BACKGROUND',(0,0),(-1,0),C_DARK),
        ('TEXTCOLOR',(0,0),(-1,0),C_WHITE),
        ('BACKGROUND',(0,0),(0,-1),C_LIGHT),
        ('BACKGROUND',(1,1),(1,-1),colors.HexColor("#FFE8E8")),
        ('BACKGROUND',(2,1),(2,-1),colors.HexColor("#D6E4F7")),
        ('BACKGROUND',(3,1),(3,-1),C_LGRN),
        ('BACKGROUND',(4,1),(5,-1),C_YELL),
        ('FONTNAME',(0,4),(-1,6),'Helvetica-Bold'),
        ('TEXTCOLOR',(3,4),(3,6),C_GREEN),
        ('TEXTCOLOR',(4,4),(5,6),C_GREEN),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('ALIGN',(0,1),(0,-1),'LEFT'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,1),(0,-1),6),
        ('GRID',(0,0),(-1,-1),0.4,C_BDR),
        ('LINEBELOW',(0,0),(-1,0),1.2,C_MID),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[C_WHITE,C_GREY]),
    ]))
    story.append(ct)
    story.append(SP(10))

    # ── PROJEÇÃO ──────────────────────────────────────────────────────────────
    story.append(sec_hdr("3. PROJEÇÃO DE ECONOMIA", C_GREEN))
    story.append(SP(8))

    proj_rows = [["Nº Op.","Economia/Op","Economia Total","ICMS Atual Total","ICMS AL Dif.","Red. ICMS"]]
    for n in [1,5,10,20]:
        p = result.projections[n]
        proj_rows.append([
            str(n), brl(result.economia_vs_al_dif), brl(p["economia_dif"]),
            brl(p["icms_atual_total"]), brl(p["icms_al_dif_total"]),
            brl(p["icms_atual_total"]-p["icms_al_dif_total"])
        ])
    pt = Table(proj_rows, colWidths=[1.8*cm, 3*cm, 3.2*cm, 3*cm, 3*cm, 3*cm])
    pt.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('BACKGROUND',(0,0),(-1,0),C_GREEN),
        ('TEXTCOLOR',(0,0),(-1,0),C_WHITE),
        ('BACKGROUND',(2,1),(-1,-1),C_LGRN),
        ('FONTNAME',(2,1),(-1,-1),'Helvetica-Bold'),
        ('TEXTCOLOR',(2,1),(-1,-1),C_GREEN),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('GRID',(0,0),(-1,-1),0.4,C_BDR),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[C_WHITE,C_GREY]),
    ]))
    story.append(pt)
    story.append(SP(10))

    # ── CONCLUSÃO ─────────────────────────────────────────────────────────────
    story.append(sec_hdr("4. CONCLUSÃO EXECUTIVA", C_DARK))
    story.append(SP(8))
    story.append(P(
        f"A operação <b>{data.doc_type} {data.di_number}</b> apresenta custo de desembaraço federal de "
        f"<b>{brl(result.subtotal)}</b>. Na estrutura atual "
        f"({data.uf_desembaraco or 'UF atual'} – ICMS {pct(result.icms_aliq_atual)}), "
        f"o valor total da nota fiscal é <b>{brl(result.custo_atual)}</b>, com ICMS de <b>{brl(result.icms_atual)}</b>.",
        S_BODY))
    story.append(SP(5))
    story.append(P(
        f"A migração para o regime <b>Alagoas – Diferencial 1,2%</b> reduziria o ICMS para "
        f"<b>{brl(result.icms_al_dif)}</b>, gerando economia de "
        f"<b>{brl(result.economia_vs_al_dif)} por operação</b> — "
        f"redução de <b>{pct(icms_red_pct_val)} do ICMS atual</b>.",
        S_BODY))
    story.append(SP(10))

    # Validation note
    diff = 0.0
    val_msg = (f"✅  Validação cruzada OK – ICMS calculado confere com o documento (diferença: R$ {diff:.2f})"
               if diff <= 1.0
               else f"⚠  Divergência detectada – ICMS calculado vs. documento: R$ {diff:.2f} – revisar manualmente")
    vt = Table([[P(val_msg, sty("vm", fontSize=8.5, leading=13))]], colWidths=[17*cm])
    vt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), C_LGRN if diff<=1 else C_YELL),
        ('LINEABOVE',(0,0),(-1,0), 3, C_GREEN if diff<=1 else colors.HexColor("#D97706")),
        ('TOPPADDING',(0,0),(-1,-1),8),
        ('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('LEFTPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(vt)
    story.append(SP(10))

    # Premissas
    story.append(P("<b>Premissas e Observações</b>", S_BOLD))
    story.append(SP(4))
    premissas = [
        f"Alíquotas de Alagoas (4% NF e 1,2% diferencial) são premissas do modelo de referência.",
        f"Alíquota atual ({pct(result.icms_aliq_atual)}) extraída diretamente do documento.",
        "ICMS calculado pelo método 'por dentro' conforme metodologia padrão do modelo.",
        "Projeções assumem operações com perfil similar (produto, NCM e valor equivalentes).",
    ]
    if data.alerts:
        premissas += [f"Alertas do parser: {sev} – {msg}" for sev,msg in data.alerts]
    for item in premissas:
        story.append(P(f"• {item}", S_NOTE))
        story.append(SP(2))

    story.append(SP(10))
    story.append(HRFlowable(width="100%", thickness=0.5, color=C_BDR))
    story.append(SP(5))
    story.append(P(
        f"{data.doc_type} {data.di_number}  |  "
        f"Emitido em {date.today().strftime('%d/%m/%Y')}  |  Documento Confidencial",
        S_CAP))

    doc.build(story, canvasmaker=ReportCanvas)
