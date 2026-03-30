"""
Main FastAPI application - web interface for DI/DUIMP analysis.
"""
import sys
import os
import uuid
import json
import zipfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
import uvicorn

# Add app directory to path
sys.path.insert(0, str(Path(__file__).parent))

from parsers.di_parser import parse_pdf
from engine.tax_calculator import calculate
from generators.excel_generator import generate_excel
from generators.pdf_generator import generate_pdf

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Analisador Tributário DI/DUIMP", version="1.0")

HTML_PAGE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Análise Tributária – Importação via Alagoas</title>
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=Overpass:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  :root {
    --ink:    #001742;
    --steel:  #004194;
    --cobalt: #FF7930;
    --sky:    #C0EFFF;
    --jade:   #004194;
    --mint:   #C0EFFF;
    --ember:  #FF7930;
    --cream:  #FAFAFA;
    --smoke:  #F3F3F3;
    --border: #E2E8F0;
    --muted:  #64748B;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Overpass', sans-serif;
    background: var(--cream);
    color: var(--ink);
    min-height: 100vh;
  }

  /* ── TOP BAR ── */
  .topbar {
    background: #001742;
    padding: 0 2.5rem;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
  }
  .topbar-logo {
    font-family: 'Oswald', sans-serif;
    font-size: 1.15rem;
    color: #fff;
    letter-spacing: 0.02em;
  }
  .topbar-tag {
    font-size: 0.72rem;
    color: #C0EFFF;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }

  /* ── HERO ── */
  .hero {
    background: linear-gradient(135deg, #004194 0%, #001742 100%);
    padding: 4rem 2.5rem 3.5rem;
    text-align: center;
    position: relative;
    overflow: hidden;
  }
  .hero::before {
    content: '';
    position: absolute;
    inset: 0;
    background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(255,121,48,0.15) 0%, transparent 70%);
  }
  .hero-eyebrow {
    font-size: 0.72rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #FF7930;
    margin-bottom: 1rem;
    position: relative;
  }
  .hero h1 {
    font-family: 'Oswald', sans-serif;
    font-size: clamp(2rem, 5vw, 3.2rem);
    color: #fff;
    line-height: 1.15;
    position: relative;
    margin-bottom: 1rem;
  }
  .hero h1 em {
    font-style: italic;
    color: #FF7930;
  }
  .hero-sub {
    font-size: 1rem;
    color: #C0EFFF;
    max-width: 520px;
    margin: 0 auto 2rem;
    line-height: 1.6;
    position: relative;
  }

  /* ── STATS BAR ── */
  .stats {
    display: flex;
    justify-content: center;
    gap: 2.5rem;
    flex-wrap: wrap;
    position: relative;
  }
  .stat { text-align: center; }
  .stat-val {
    font-family: 'Oswald', sans-serif;
    font-size: 1.8rem;
    color: #fff;
    display: block;
  }
  .stat-lbl {
    font-size: 0.72rem;
    color: #64748B;
    text-transform: uppercase;
    letter-spacing: 0.08em;
  }
  .stat-divider {
    width: 1px;
    background: rgba(255,121,48,0.3);
    align-self: stretch;
  }

  /* ── MAIN LAYOUT ── */
  .main {
    max-width: 900px;
    margin: 0 auto;
    padding: 3rem 1.5rem 5rem;
  }

  /* ── UPLOAD CARD ── */
  .upload-card {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 4px 24px rgba(13,27,42,0.06);
    margin-bottom: 2rem;
  }
  .card-eyebrow {
    font-size: 0.7rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #FF7930;
    margin-bottom: 0.5rem;
  }
  .card-title {
    font-family: 'Oswald', sans-serif;
    font-size: 1.5rem;
    color: var(--ink);
    margin-bottom: 0.4rem;
  }
  .card-desc {
    font-size: 0.88rem;
    color: var(--muted);
    margin-bottom: 1.8rem;
    line-height: 1.55;
  }

  /* Drop zone */
  .drop-zone {
    border: 2px dashed var(--border);
    border-radius: 12px;
    padding: 2.5rem 1.5rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s;
    background: var(--smoke);
    position: relative;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: #FF7930;
    background: var(--sky);
  }
  .drop-zone input[type=file] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
    width: 100%;
    height: 100%;
  }
  .drop-icon {
    font-size: 2.2rem;
    margin-bottom: 0.8rem;
    display: block;
    opacity: 0.5;
  }
  .drop-label {
    font-size: 0.92rem;
    font-weight: 500;
    color: var(--ink);
    margin-bottom: 0.3rem;
  }
  .drop-hint {
    font-size: 0.78rem;
    color: var(--muted);
  }
  .file-chosen {
    margin-top: 0.75rem;
    font-size: 0.82rem;
    color: var(--jade);
    font-weight: 500;
  }

  /* Submit button */
  .btn-analyze {
    width: 100%;
    padding: 0.95rem 1.5rem;
    background: var(--cobalt);
    color: #fff;
    border: none;
    border-radius: 10px;
    font-family: 'Overpass', sans-serif;
    font-size: 0.95rem;
    font-weight: 600;
    cursor: pointer;
    margin-top: 1.5rem;
    transition: all 0.2s;
    letter-spacing: 0.02em;
  }
  .btn-analyze:hover:not(:disabled) {
    background: #e66920;
    transform: translateY(-1px);
    box-shadow: 0 6px 20px rgba(255,121,48,0.4);
  }
  .btn-analyze:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  /* ── PROGRESS ── */
  .progress-card {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 4px 24px rgba(13,27,42,0.06);
    display: none;
    margin-bottom: 2rem;
  }
  .progress-title {
    font-family: 'Oswald', sans-serif;
    font-size: 1.3rem;
    margin-bottom: 1.5rem;
    color: var(--ink);
  }
  .step {
    display: flex;
    align-items: center;
    gap: 1rem;
    padding: 0.75rem 0;
    border-bottom: 1px solid var(--border);
  }
  .step:last-child { border-bottom: none; }
  .step-icon {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1rem;
    flex-shrink: 0;
    transition: all 0.3s;
  }
  .step-icon.pending  { background: var(--smoke); color: var(--muted); }
  .step-icon.running  { background: var(--sky);   color: #FF7930; animation: pulse 1s infinite; }
  .step-icon.done     { background: #C0EFFF;  color: #004194; }
  .step-icon.error    { background: #FEE2E2;      color: var(--ember); }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
  .step-label { font-size: 0.9rem; font-weight: 500; }
  .step-detail { font-size: 0.78rem; color: var(--muted); }

  /* ── RESULTS ── */
  .results-card {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 4px 24px rgba(13,27,42,0.06);
    display: none;
    margin-bottom: 2rem;
  }

  .result-header {
    display: flex;
    align-items: center;
    gap: 1rem;
    margin-bottom: 1.5rem;
  }
  .result-checkmark {
    width: 44px; height: 44px;
    background: var(--mint);
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.3rem;
    flex-shrink: 0;
  }
  .result-title {
    font-family: 'Oswald', sans-serif;
    font-size: 1.4rem;
    color: #004194;
  }
  .result-subtitle { font-size: 0.83rem; color: var(--muted); }

  /* KPIs */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
    margin-bottom: 1.8rem;
  }
  .kpi {
    background: var(--smoke);
    border-radius: 12px;
    padding: 1.2rem 1rem;
    text-align: center;
    border: 1px solid var(--border);
  }
  .kpi.highlight {
    background: #C0EFFF;
    border-color: #004194;
  }
  .kpi-value {
    font-family: 'Oswald', sans-serif;
    font-size: 1.6rem;
    color: #004194;
    display: block;
    line-height: 1.2;
  }
  .kpi-label {
    font-size: 0.73rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-top: 0.3rem;
  }

  /* Comparison table */
  .comp-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
    margin-bottom: 1.5rem;
  }
  .comp-table th {
    background: var(--ink);
    color: #fff;
    padding: 0.65rem 0.75rem;
    text-align: center;
    font-size: 0.78rem;
    font-weight: 500;
  }
  .comp-table th:first-child { text-align: left; border-radius: 8px 0 0 0; }
  .comp-table th:last-child { border-radius: 0 8px 0 0; }
  .comp-table td {
    padding: 0.6rem 0.75rem;
    text-align: center;
    border-bottom: 1px solid var(--border);
  }
  .comp-table td:first-child { text-align: left; font-weight: 500; }
  .comp-table tr:nth-child(even) td { background: var(--smoke); }
  .comp-table tr:hover td { background: var(--sky); }
  .comp-table .col-atual { background: #FFF0EE !important; }
  .comp-table .col-al { background: #C0EFFF !important; color: #004194; font-weight: 600; }
  .comp-table .col-eco { background: #FFFBEB !important; }

  /* Download buttons */
  .download-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1rem;
    margin-bottom: 1rem;
  }
  .btn-download {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 1rem 1.2rem;
    border: 1.5px solid var(--border);
    border-radius: 10px;
    cursor: pointer;
    background: #fff;
    text-decoration: none;
    transition: all 0.2s;
    color: var(--ink);
  }
  .btn-download:hover {
    border-color: #FF7930;
    background: var(--sky);
    transform: translateY(-1px);
  }
  .btn-download.primary {
    background: var(--cobalt);
    border-color: #FF7930;
    color: #fff;
  }
  .btn-download.primary:hover {
    background: #e66920;
    border-color: #e66920;
  }
  .dl-icon { font-size: 1.5rem; }
  .dl-title { font-size: 0.88rem; font-weight: 600; }
  .dl-sub { font-size: 0.73rem; opacity: 0.7; }

  /* Alerts */
  .alert {
    background: #FEF3C7;
    border-left: 4px solid #D97706;
    border-radius: 0 8px 8px 0;
    padding: 0.85rem 1rem;
    font-size: 0.83rem;
    color: #92400E;
    margin-bottom: 1rem;
  }

  /* New analysis */
  .btn-new {
    width: 100%;
    padding: 0.8rem;
    border: 1.5px solid var(--border);
    border-radius: 10px;
    background: #fff;
    font-family: 'Overpass', sans-serif;
    font-size: 0.88rem;
    color: var(--muted);
    cursor: pointer;
    transition: all 0.2s;
  }
  .btn-new:hover {
    border-color: #FF7930;
    color: #FF7930;
  }

  /* ── HOW IT WORKS ── */
  .how-card {
    background: #004194;
    border-radius: 16px;
    padding: 2.5rem;
    margin-bottom: 2rem;
  }
  .how-title {
    font-family: 'Oswald', sans-serif;
    font-size: 1.4rem;
    color: #fff;
    margin-bottom: 1.5rem;
  }
  .how-steps {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1.2rem;
  }
  .how-step {
    background: rgba(255,255,255,0.07);
    border-radius: 10px;
    padding: 1.2rem;
  }
  .how-num {
    font-family: 'Oswald', sans-serif;
    font-size: 2rem;
    color: #FF7930;
    line-height: 1;
    margin-bottom: 0.5rem;
  }
  .how-step-title { font-size: 0.88rem; font-weight: 600; color: #fff; margin-bottom: 0.3rem; }
  .how-step-desc { font-size: 0.78rem; color: #C0EFFF; line-height: 1.5; }

  /* ── FOOTER ── */
  footer {
    text-align: center;
    padding: 2rem;
    font-size: 0.78rem;
    color: #C0EFFF;
    border-top: 2px solid #FF7930;
  }

  /* Responsive */
  @media (max-width: 640px) {
    .kpi-grid { grid-template-columns: 1fr; }
    .how-steps { grid-template-columns: 1fr; }
    .download-grid { grid-template-columns: 1fr; }
    .stats { gap: 1.5rem; }
    .stat-divider { display: none; }
  }
</style>
</head>
<body>

<div class="topbar">
  <span class="topbar-logo">SAYGO <span style="color:#FF7930;font-weight:400;font-size:0.85rem;letter-spacing:0.15em;">VISION</span></span>
  <span class="topbar-tag" style="color:#FF7930;letter-spacing:0.1em;">ANÁLISE TRIBUTÁRIA · IMPORTAÇÃO VIA ALAGOAS</span>
</div>

<div class="hero">
  <div class="hero-eyebrow">Saygo Vision · Benefícios Fiscais</div>
  <h1>Suba a DI ou DUIMP.<br><em>A economia aparece na hora.</em></h1>
  <p class="hero-sub">
    Extração automática de dados, cálculo de ICMS, comparativo com Alagoas
    e geração de relatório executivo em segundos.
  </p>
  <div class="stats">
    <div class="stat">
      <span class="stat-val">~90%</span>
      <span class="stat-lbl">Redução ICMS</span>
    </div>
    <div class="stat-divider"></div>
    <div class="stat">
      <span class="stat-val">30s</span>
      <span class="stat-lbl">Tempo de análise</span>
    </div>
    <div class="stat-divider"></div>
    <div class="stat">
      <span class="stat-val">2</span>
      <span class="stat-lbl">Entregáveis gerados</span>
    </div>
    <div class="stat-divider"></div>
    <div class="stat">
      <span class="stat-val">0</span>
      <span class="stat-lbl">Erros manuais</span>
    </div>
  </div>
</div>

<div class="main">

  <!-- UPLOAD -->
  <div class="upload-card" id="uploadCard">
    <div class="card-eyebrow">Passo 1 de 1</div>
    <div class="card-title">Envie o documento aduaneiro</div>
    <p class="card-desc">
      Aceita DI (Declaração de Importação) ou DUIMP em formato PDF.
      O sistema identifica automaticamente o tipo do documento.
    </p>

    <div class="drop-zone" id="dropZone">
      <input type="file" id="fileInput" accept=".pdf" />
      <span class="drop-icon">📄</span>
      <div class="drop-label">Arraste o PDF aqui ou clique para selecionar</div>
      <div class="drop-hint">Somente arquivos .pdf · DI ou DUIMP da Receita Federal</div>
      <div class="file-chosen" id="fileChosen"></div>
    </div>

        <button class="btn-analyze" id="btnAnalyze" disabled onclick="runAnalysis()">
      Analisar documento
    </button>
  </div>

  <!-- PROGRESS -->
  <div class="progress-card" id="progressCard">
    <div class="progress-title">Processando seu documento...</div>
    <div id="steps">
      <div class="step" id="step-0">
        <div class="step-icon pending" id="icon-0">①</div>
        <div>
          <div class="step-label">Leitura e extração de campos</div>
          <div class="step-detail" id="detail-0">Aguardando...</div>
        </div>
      </div>
      <div class="step" id="step-1">
        <div class="step-icon pending" id="icon-1">②</div>
        <div>
          <div class="step-label">Validação dos dados extraídos</div>
          <div class="step-detail" id="detail-1">Aguardando...</div>
        </div>
      </div>
      <div class="step" id="step-2">
        <div class="step-icon pending" id="icon-2">③</div>
        <div>
          <div class="step-label">Cálculo tributário (ICMS por dentro)</div>
          <div class="step-detail" id="detail-2">Aguardando...</div>
        </div>
      </div>
      <div class="step" id="step-3">
        <div class="step-icon pending" id="icon-3">④</div>
        <div>
          <div class="step-label">Comparativo Alagoas</div>
          <div class="step-detail" id="detail-3">Aguardando...</div>
        </div>
      </div>
      <div class="step" id="step-4">
        <div class="step-icon pending" id="icon-4">⑤</div>
        <div>
          <div class="step-label">Geração da planilha (XLSX)</div>
          <div class="step-detail" id="detail-4">Aguardando...</div>
        </div>
      </div>
      <div class="step" id="step-5">
        <div class="step-icon pending" id="icon-5">⑥</div>
        <div>
          <div class="step-label">Geração do relatório executivo (PDF)</div>
          <div class="step-detail" id="detail-5">Aguardando...</div>
        </div>
      </div>
    </div>
  </div>

  <!-- RESULTS -->
  <div class="results-card" id="resultsCard">
    <div class="result-header">
      <div class="result-checkmark">✅</div>
      <div>
        <div class="result-title">Análise concluída</div>
        <div class="result-subtitle" id="resultMeta">–</div>
      </div>
    </div>

    <div id="alertsBox"></div>

    <div class="kpi-grid" id="kpiGrid">
      <div class="kpi">
        <span class="kpi-value" id="kpiSubtotal">–</span>
        <div class="kpi-label">Custo Desembaraço Federal</div>
      </div>
      <div class="kpi">
        <span class="kpi-value" id="kpiIcmsAtual">–</span>
        <div class="kpi-label">ICMS Operação Atual</div>
      </div>
      <div class="kpi highlight">
        <span class="kpi-value" id="kpiEconomia">–</span>
        <div class="kpi-label">Economia por Operação</div>
      </div>
    </div>

    <table class="comp-table" id="compTable">
      <thead>
        <tr>
          <th>Parâmetro</th>
          <th class="col-atual">Atual</th>
          <th>AL NF 4%</th>
          <th class="col-al">AL Dif. 1,2%</th>
          <th class="col-eco">Economia</th>
        </tr>
      </thead>
      <tbody id="compBody"></tbody>
    </table>

    <div class="download-grid">
      <a class="btn-download primary" id="dlZip" href="#" download>
        <span class="dl-icon">📦</span>
        <div>
          <div class="dl-title">Baixar tudo (.zip)</div>
          <div class="dl-sub">Planilha + Relatório PDF</div>
        </div>
      </a>
      <a class="btn-download" id="dlPdf" href="#" download>
        <span class="dl-icon">📋</span>
        <div>
          <div class="dl-title">Relatório PDF</div>
          <div class="dl-sub">Relatório executivo corporativo</div>
        </div>
      </a>
      <a class="btn-download" id="dlXlsx" href="#" download>
        <span class="dl-icon">📊</span>
        <div>
          <div class="dl-title">Planilha XLSX</div>
          <div class="dl-sub">5 abas com análise completa</div>
        </div>
      </a>
    </div>

    <button class="btn-new" onclick="resetForm()">↩  Nova análise</button>
  </div>

  <!-- HOW IT WORKS -->
  <div class="how-card" id="howCard">
    <div class="how-title">Como funciona</div>
    <div class="how-steps">
      <div class="how-step">
        <div class="how-num">01</div>
        <div class="how-step-title">Suba o documento</div>
        <div class="how-step-desc">DI ou DUIMP em PDF da Receita Federal. O sistema detecta o tipo automaticamente.</div>
      </div>
      <div class="how-step">
        <div class="how-num">02</div>
        <div class="how-step-title">Extração e cálculo</div>
        <div class="how-step-desc">21 campos extraídos, 12 regras de cálculo aplicadas, ICMS comparado com o documento.</div>
      </div>
      <div class="how-step">
        <div class="how-num">03</div>
        <div class="how-step-title">Receba os entregáveis</div>
        <div class="how-step-desc">Planilha Excel com 5 abas e relatório PDF corporativo prontos para apresentar ao cliente.</div>
      </div>
    </div>
  </div>

</div>

<footer>
  SAYGO VISION · Análise Tributária DI/DUIMP · Importação via Alagoas
</footer>

<script>
  const fileInput = document.getElementById('fileInput');
  const fileChosen = document.getElementById('fileChosen');
  const btnAnalyze = document.getElementById('btnAnalyze');
  const dropZone = document.getElementById('dropZone');

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) {
      fileChosen.textContent = '✓  ' + fileInput.files[0].name;
      btnAnalyze.disabled = false;
    }
  });

  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f && f.name.endsWith('.pdf')) {
      const dt = new DataTransfer();
      dt.items.add(f);
      fileInput.files = dt.files;
      fileChosen.textContent = '✓  ' + f.name;
      btnAnalyze.disabled = false;
    }
  });

  function brl(v) {
    return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\\B(?=(\\d{3})+(?!\\d))/g, '.');
  }

  function pct(v) {
    return (v * 100).toFixed(1).replace('.', ',') + '%';
  }

  function setStep(i, state, detail) {
    const icon = document.getElementById('icon-' + i);
    const det  = document.getElementById('detail-' + i);
    icon.className = 'step-icon ' + state;
    det.textContent = detail;
    if (state === 'done')   icon.textContent = '✓';
    if (state === 'error')  icon.textContent = '✗';
    if (state === 'running') icon.textContent = ['①','②','③','④','⑤','⑥'][i];
  }

  async function runAnalysis() {
    if (!fileInput.files[0]) return;

    // Hide upload, show progress
    document.getElementById('uploadCard').style.display = 'none';
    document.getElementById('howCard').style.display = 'none';
    document.getElementById('progressCard').style.display = 'block';

    // Reset steps
    for (let i = 0; i < 6; i++) setStep(i, 'pending', 'Aguardando...');

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);


    // Animate steps while waiting
    const stepMessages = [
      'Lendo PDF com pdfplumber...',
      'Validando campos obrigatórios...',
      'Calculando ICMS por dentro...',
      'Comparando com Alagoas...',
      'Montando planilha Excel (5 abas)...',
      'Gerando relatório PDF corporativo...',
    ];

    let stepIdx = 0;
    const stepInterval = setInterval(() => {
      if (stepIdx < 6) {
        if (stepIdx > 0) setStep(stepIdx - 1, 'done', 'Concluído');
        setStep(stepIdx, 'running', stepMessages[stepIdx]);
        stepIdx++;
      }
    }, 400);

    try {
      const resp = await fetch('/analyze', { method: 'POST', body: formData });
      clearInterval(stepInterval);

      if (!resp.ok) {
        const err = await resp.json();
        for (let i = 0; i < 6; i++) setStep(i, stepIdx > i ? 'done' : 'error', stepIdx > i ? 'Concluído' : 'Erro');
        setStep(Math.min(stepIdx, 5), 'error', err.detail || 'Erro no processamento');
        return;
      }

      // All done
      for (let i = 0; i < 6; i++) setStep(i, 'done', 'Concluído');
      await new Promise(r => setTimeout(r, 600));

      const data = await resp.json();
      showResults(data);

    } catch (e) {
      clearInterval(stepInterval);
      setStep(Math.max(stepIdx - 1, 0), 'error', 'Erro de rede: ' + e.message);
    }
  }

  function showResults(d) {
    document.getElementById('progressCard').style.display = 'none';
    const rc = document.getElementById('resultsCard');
    rc.style.display = 'block';

    // Meta
    document.getElementById('resultMeta').textContent =
      `${d.doc_type} ${d.di_number}  ·  ${d.importador_nome}  ·  ${d.register_date}`;

    // KPIs
    document.getElementById('kpiSubtotal').textContent    = brl(d.subtotal);
    document.getElementById('kpiIcmsAtual').textContent   = brl(d.icms_atual);
    document.getElementById('kpiEconomia').textContent    = brl(d.economia_dif);

    // Alerts
    const alertsBox = document.getElementById('alertsBox');
    alertsBox.innerHTML = '';
    if (d.alerts && d.alerts.length > 0) {
      const warnAlerts = d.alerts.filter(a => a[0] !== 'ERROR');
      warnAlerts.forEach(([sev, msg]) => {
        const div = document.createElement('div');
        div.className = 'alert';
        div.textContent = `⚠ ${msg}`;
        alertsBox.appendChild(div);
      });
    }

    // Table
    const body = document.getElementById('compBody');
    const aliqLabel = (d.icms_aliq * 100).toFixed(1).replace('.', ',') + '%';
    const rows = [
      ['Alíquota ICMS',    aliqLabel,       '4,0%',          '1,2%',          '–'],
      ['Valor da NF',       brl(d.custo_atual), brl(d.nf_al_nf), brl(d.nf_al_dif), brl(d.economia_dif)],
      ['Valor do ICMS',    brl(d.icms_atual),brl(d.icms_al_nf),brl(d.icms_al_dif),brl(d.icms_atual - d.icms_al_dif)],
      ['Custo Total',      brl(d.custo_atual), brl(d.nf_al_nf), brl(d.nf_al_dif), brl(d.economia_dif)],
    ];
    body.innerHTML = rows.map(r =>
      `<tr>
        <td>${r[0]}</td>
        <td class="col-atual">${r[1]}</td>
        <td>${r[2]}</td>
        <td class="col-al">${r[3]}</td>
        <td class="col-eco">${r[4]}</td>
      </tr>`
    ).join('');

    // Downloads
    document.getElementById('dlZip').href  = `/download/${d.job_id}/zip`;
    document.getElementById('dlPdf').href  = `/download/${d.job_id}/pdf`;
    document.getElementById('dlXlsx').href = `/download/${d.job_id}/xlsx`;
    document.getElementById('dlZip').download  = `${d.di_number.replace('/','_')}_analise.zip`;
    document.getElementById('dlPdf').download  = `${d.di_number.replace('/','_')}_relatorio.pdf`;
    document.getElementById('dlXlsx').download = `${d.di_number.replace('/','_')}_comparativo.xlsx`;
  }

  function resetForm() {
    document.getElementById('resultsCard').style.display = 'none';
    document.getElementById('uploadCard').style.display = 'block';
    document.getElementById('howCard').style.display = 'block';
    document.getElementById('fileChosen').textContent = '';
    document.getElementById('btnAnalyze').disabled = true;
    fileInput.value = '';
    for (let i = 0; i < 6; i++) setStep(i, 'pending', 'Aguardando...');
  }
</script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def root():
    return HTML_PAGE


@app.post("/analyze")
async def analyze(file: UploadFile = File(...), api_key: str = Form(default="")):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(400, "Somente arquivos PDF são aceitos")

    job_id = str(uuid.uuid4())[:8]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    # Save uploaded file
    pdf_path = job_dir / "input.pdf"
    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    # Parse
    try:
        data = parse_pdf(str(pdf_path), api_key=api_key or None)
    except Exception as e:
        raise HTTPException(500, f"Erro na leitura do PDF: {str(e)}")

    # Check critical fields
    critical_errors = [msg for sev, msg in data.alerts if sev == "ERROR"]
    if data.vmld_usd == 0 or data.taxa_cambio == 0:
        raise HTTPException(422, f"Campos obrigatórios ausentes: {'; '.join(critical_errors)}")

    # Calculate
    result = calculate(data)

    # Generate files
    xlsx_path = job_dir / f"comparativo_{job_id}.xlsx"
    pdf_path_out = job_dir / f"relatorio_{job_id}.pdf"

    generate_excel(data, result, str(xlsx_path))
    generate_pdf(data, result, str(pdf_path_out))

    # Create zip
    zip_path = job_dir / f"analise_{job_id}.zip"
    with zipfile.ZipFile(zip_path, 'w') as zf:
        zf.write(xlsx_path, xlsx_path.name)
        zf.write(pdf_path_out, pdf_path_out.name)

    # Return JSON summary
    return JSONResponse({
        "job_id":        job_id,
        "doc_type":      data.doc_type,
        "di_number":     data.di_number,
        "register_date": data.register_date,
        "importador_nome": data.importador_nome,
        "ncm":           data.ncm,
        "subtotal":      round(result.subtotal, 2),
        "icms_aliq":     round(result.icms_aliq_atual, 4),
        "custo_atual":  round(result.custo_atual, 2),
        "icms_atual":   round(result.icms_atual, 2),
        "nf_al_nf":      round(result.nf_al_nf, 2),
        "icms_al_nf":    round(result.icms_al_nf, 2),
        "nf_al_dif":     round(result.nf_al_dif, 2),
        "icms_al_dif":   round(result.icms_al_dif, 2),
        "economia_nf":  round(result.economia_vs_al_nf, 2),
        "economia_dif": round(result.economia_vs_al_dif, 2),
        "reducao_icms": round(result.reducao_icms_al_dif, 2),
        "projections":   result.projections,
        "alerts":        data.alerts,
        "warnings":      result.warnings,
    })


@app.get("/download/{job_id}/{filetype}")
async def download(job_id: str, filetype: str):
    job_dir = OUTPUT_DIR / job_id
    if not job_dir.exists():
        raise HTTPException(404, "Job não encontrado")

    if filetype == "zip":
        files = list(job_dir.glob("analise_*.zip"))
    elif filetype == "pdf":
        files = list(job_dir.glob("relatorio_*.pdf"))
    elif filetype == "xlsx":
        files = list(job_dir.glob("comparativo_*.xlsx"))
    else:
        raise HTTPException(400, "Tipo inválido")

    if not files:
        raise HTTPException(404, "Arquivo não encontrado")

    return FileResponse(str(files[0]), filename=files[0].name)


if __name__ == "__main__":
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "verifique suas configuracoes de rede"

    print("\n" + "="*54)
    print("  ANALISE TRIBUTARIA DI/DUIMP  —  Servidor Iniciado")
    print("="*54)
    print(f"  Neste computador : http://localhost:8000")
    print(f"  Na rede local    : http://{local_ip}:8000")
    print("-"*54)
    print("  Compartilhe o endereco da rede com a equipe.")
    print("  Para encerrar: Ctrl+C")
    print("="*54 + "\n")

    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="warning")

# ── Network startup info ─────────────────────────────────────────────────────
# (replaces the simple uvicorn.run above at runtime via __main__ guard)
