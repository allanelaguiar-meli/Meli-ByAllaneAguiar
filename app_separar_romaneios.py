import streamlit as st
import pandas as pd
import fitz  # pymupdf
import PyPDF2
import io
import os
import re
import zipfile
import tempfile
import shutil
from datetime import datetime
import pytz

# ── Configuração ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Separador de Romaneios",
    page_icon="📦",
    layout="wide",
)
BRT = pytz.timezone("America/Sao_Paulo")

# ── Estilo ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { background: #f4f6fb; }
    .titulo { font-size:2rem; font-weight:800; color:#1a1a2e; letter-spacing:-1px; }
    .subtitulo { color:#666; font-size:.9rem; margin-bottom:1.5rem; }
    .card { background:#fff; border-radius:14px; padding:18px 22px;
            box-shadow:0 2px 12px rgba(0,0,0,.07); margin-bottom:10px; }
    .card-label { font-size:.72rem; color:#999; text-transform:uppercase; letter-spacing:1px; }
    .card-val { font-size:1.8rem; font-weight:800; color:#1a1a2e; }
    .tip { background:#e8f4fd; border-left:4px solid #2196F3;
           border-radius:8px; padding:10px 16px; margin-bottom:1rem; font-size:.88rem; }
    .step { background:#fff; border-radius:12px; padding:14px 18px;
            border-left:4px solid #e94560; margin-bottom:8px; box-shadow:0 1px 6px rgba(0,0,0,.05); }
    .step-num { font-size:.72rem; color:#e94560; font-weight:700; text-transform:uppercase; letter-spacing:1px; }
    .step-title { font-size:1rem; font-weight:700; color:#1a1a2e; margin-top:2px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES EDITÁVEIS
# ══════════════════════════════════════════════════════════════════════════════
IDX_COL_ROTA    = 3   # Coluna D
IDX_COL_QR      = 4   # Coluna E
IDX_COL_TRANSP  = 15  # Coluna P
IDX_COL_DRIVER  = 16  # Coluna Q

MSG_TOPO   = "ATENÇÃO! DIARIAMENTE AO CHEGAR ESCANEIE O QR CODE NA MESA PARA INFORMAR SUA PRESENÇA"
MSG_RODAPE = "FAVOR NÃO DESCARTAR EM VIA PÚBLICA"

# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════════════════════════════

def agora_str():
    return datetime.now(BRT).strftime("%Y%m%d_%H%M")

def safe_str(s):
    if s is None: return ""
    if isinstance(s, float) and (pd.isna(s) or str(s).lower() == "nan"): return ""
    return str(s).strip()

def sanitize(nome):
    """Remove caracteres inválidos para nome de arquivo/pasta."""
    return re.sub(r'[\\/*?:"<>|]', "-", nome).strip() or "_"

def is_envios_extra(nome):
    n = nome.strip().upper().replace("Ç", "C")
    n = re.sub(r'[\s_\-]+', ' ', n)
    return any(x in n for x in ["ENVIO EXTRA", "ENVIOS EXTRA", "ENVIOS EXTRAS", "ENVIO EXTRAS"])

def carregar_planilha(arq_bytes, nome_arq):
    """Lê xlsx ou csv e retorna dict {rota_id: {TRANSPORTADORA, ROMANEIO, MOTORISTA}}."""
    if nome_arq.lower().endswith('.csv'):
        df = pd.read_csv(io.BytesIO(arq_bytes), header=None)
    else:
        df = pd.read_excel(io.BytesIO(arq_bytes), sheet_name='PLAN', header=None)

    # Pula linhas de cabeçalho até encontrar string na coluna de rota
    skip = 0
    while skip < 10 and not isinstance(df.iloc[skip, IDX_COL_ROTA], str):
        skip += 1
    if skip > 0:
        df = df.iloc[skip:]

    rotas = {}
    for row in df.values.tolist():
        rota       = safe_str(row[IDX_COL_ROTA])   if len(row) > IDX_COL_ROTA   else ""
        romaneio   = safe_str(row[IDX_COL_QR])      if len(row) > IDX_COL_QR     else ""
        transp     = safe_str(row[IDX_COL_TRANSP])  if len(row) > IDX_COL_TRANSP else ""
        motorista  = safe_str(row[IDX_COL_DRIVER])  if len(row) > IDX_COL_DRIVER else ""
        if rota:
            rotas[rota] = {"TRANSPORTADORA": transp, "ROMANEIO": romaneio, "MOTORISTA": motorista}
    return rotas

def coletar_pdfs(arquivos_up):
    """Recebe lista de UploadedFile → dict {nome: bytes}."""
    result = {}
    for arq in arquivos_up:
        dados = arq.read()
        if arq.name.lower().endswith('.pdf'):
            result[arq.name] = dados
        elif arq.name.lower().endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(dados)) as z:
                for nome in z.namelist():
                    if nome.lower().endswith('.pdf'):
                        result[os.path.basename(nome)] = z.read(nome)
    return result

def extrair_texto_pagina(pdf_bytes, pg_num):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = doc.load_page(pg_num).get_text()
    doc.close()
    return texto

def get_rota_id(texto):
    m = re.search(r'Rota\s+([A-Z0-9_]+)', texto)
    if m and m.group(1).count('_') == 1:
        return m.group(1)
    return None

def anotar_pdf(pdf_path, romaneio, motorista, extra=False):
    """Sobrescreve número do romaneio e, se extra, nome do motorista + mensagens."""
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc):
        # Cobre e reescreve número (canto superior direito)
        page.draw_rect([729, 38, 800, 58], color=(1,1,1), fill=(1,1,1))
        page.insert_text((730, 50), str(romaneio), fontsize=18, color=(0,0,0))

        if i == 0 and extra and motorista:
            page.draw_rect([110, 64, 330, 87], color=(1,1,1), fill=(1,1,1))
            page.insert_text((115, 79), str(motorista), fontsize=16, color=(0,0,0))

        if i == 0 and extra:
            if MSG_TOPO:
                page.insert_text((60, 25), MSG_TOPO, fontsize=11, color=(1, 0, 0))
            if MSG_RODAPE:
                page.insert_text((60, page.rect.height - 20), MSG_RODAPE, fontsize=11, color=(1, 0, 0))

    doc.saveIncr()
    doc.close()

def pad_para_multiplo_de_2(pdf_path):
    """
    Garante que o romaneio tenha número PAR de páginas.
    Isso permite impressão 2-em-1 frente/verso com grampeamento individual.
    """
    doc = fitz.open(pdf_path)
    if doc.page_count % 2 != 0:
        w, h = doc[-1].rect.width, doc[-1].rect.height
        blank = fitz.open()
        blank.new_page(width=w, height=h)
        doc.insert_pdf(blank)
        doc.saveIncr()
    doc.close()

def montar_pdf_unificado_2por1(lista_pdfs, destino):
    """
    Junta todos os PDFs de Envios Extra em um único arquivo.
    Cada romaneio já foi padded para múltiplo de 2 páginas.
    O resultado pode ser impresso com "2 páginas por folha + frente e verso"
    e cada romaneio fica grampeável individualmente.
    """
    merged = fitz.open()
    for p in lista_pdfs:
        doc = fitz.open(p)
        merged.insert_pdf(doc)
        doc.close()
    merged.save(destino)
    merged.close()

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSAMENTO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def processar_romaneios(pdfs_dict, rotas_info, opcao, tmpdir, progress_bar):
    """
    Separa PDFs por transportadora.
    Retorna dict com estatísticas e caminhos dos ZIPs gerados.
    """
    pastas = {}       # transportadora → pasta no tmpdir
    extras_pdfs = []  # caminhos dos PDFs de envios extra (em ordem)
    stats = {"total": 0, "extra": 0, "transportadoras": {}}

    todos = list(pdfs_dict.items())
    for idx, (filename, filedata) in enumerate(todos):
        progress_bar.progress((idx + 1) / len(todos), text=f"Processando {filename}…")

        reader = PyPDF2.PdfReader(io.BytesIO(filedata))
        n_pages = len(reader.pages)

        # Detecta início de cada romaneio dentro do PDF
        starts, rota_ids = [], []
        for i in range(n_pages):
            txt = extrair_texto_pagina(filedata, i)
            if "Roteiro" in txt:
                rid = get_rota_id(txt)
                starts.append(i)
                rota_ids.append(rid)
        starts.append(n_pages)

        for j in range(len(starts) - 1):
            rota_id = rota_ids[j]
            if not rota_id:
                continue

            info      = rotas_info.get(rota_id, {"TRANSPORTADORA": "", "ROMANEIO": "", "MOTORISTA": ""})
            transp    = sanitize(info["TRANSPORTADORA"]) or "SEM_TRANSPORTADORA"
            romaneio  = sanitize(info["ROMANEIO"]) or "_"
            motorista = info["MOTORISTA"]
            extra     = is_envios_extra(transp)

            # Filtra por opção escolhida
            if opcao == "mlp" and extra:
                continue
            if opcao == "extra" and not extra:
                continue

            # Cria pasta da transportadora se não existir
            if transp not in pastas:
                pasta = os.path.join(tmpdir, transp)
                os.makedirs(pasta, exist_ok=True)
                pastas[transp] = pasta

            nome_arq = f"{transp}-{rota_id}-{romaneio}.pdf"
            caminho  = os.path.join(pastas[transp], nome_arq)

            # Extrai páginas do romaneio
            writer = PyPDF2.PdfWriter()
            for pg in range(starts[j], starts[j + 1]):
                writer.add_page(reader.pages[pg])
            with open(caminho, "wb") as f:
                writer.write(f)

            # Anota PDF
            anotar_pdf(caminho, romaneio, motorista, extra=extra)

            # Contabiliza
            stats["total"] += 1
            stats["transportadoras"][transp] = stats["transportadoras"].get(transp, 0) + 1
            if extra:
                stats["extra"] += 1
                pad_para_multiplo_de_2(caminho)
                extras_pdfs.append(caminho)

    # Gera PDF unificado de Envios Extra
    extras_zip_path = None
    unificado_path  = None

    for transp, pasta in pastas.items():
        if is_envios_extra(transp) and extras_pdfs:
            unificado_path = os.path.join(pasta, "ENVIOS_EXTRA_TODOS_2por1.pdf")
            montar_pdf_unificado_2por1(extras_pdfs, unificado_path)
            break

    # Cria um ZIP por transportadora
    zips = {}
    for transp, pasta in pastas.items():
        zip_path = pasta + ".zip"
        shutil.make_archive(pasta, "zip", pasta)
        zips[transp] = zip_path

    return stats, zips, unificado_path

# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════════

# Session state
for k, v in [("stats", {}), ("zips", {}), ("unificado", None), ("processado", False)]:
    if k not in st.session_state:
        st.session_state[k] = v

st.markdown('<div class="titulo">📦 Separador de Romaneios</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Separa PDFs por transportadora, anota romaneios e gera PDF unificado Envios Extra</div>', unsafe_allow_html=True)

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuração")

    st.markdown("**1 · Planilha de rotas**")
    st.caption("xlsx (aba PLAN) ou csv")
    arq_plan = st.file_uploader("Planilha", type=["xlsx", "csv"], label_visibility="collapsed")

    st.divider()
    st.markdown("**2 · PDFs ou ZIP**")
    st.caption("Pode enviar múltiplos arquivos ou um ZIP com todos")
    arqs_pdf = st.file_uploader("PDFs / ZIP", type=["pdf", "zip"],
                                accept_multiple_files=True, label_visibility="collapsed")

    st.divider()
    st.markdown("**3 · O que separar?**")
    opcao = st.radio(
        "Separar:",
        options=["ambos", "mlp", "extra"],
        format_func=lambda x: {"ambos": "🔄 MLPs + Envios Extra", "mlp": "📋 Só MLPs", "extra": "📦 Só Envios Extra"}[x],
        label_visibility="collapsed",
    )

    st.divider()
    btn = st.button("▶️ Processar", use_container_width=True, type="primary")

# ── INSTRUÇÕES ────────────────────────────────────────────────────────────────
if not st.session_state.processado:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="step">
            <div class="step-num">Passo 1</div>
            <div class="step-title">📊 Planilha de Rotas</div>
            <div style="font-size:.82rem;color:#555;margin-top:6px;">
                Faça upload da sua planilha (xlsx com aba <b>PLAN</b> ou csv).<br>
                Ela define rota, transportadora, romaneio e motorista.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="step">
            <div class="step-num">Passo 2</div>
            <div class="step-title">📄 PDFs dos Romaneios</div>
            <div style="font-size:.82rem;color:#555;margin-top:6px;">
                Envie os PDFs individuais ou um ZIP com todos.<br>
                O app detecta automaticamente cada romaneio.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="step">
            <div class="step-num">Passo 3</div>
            <div class="step-title">⬇️ Baixar Resultados</div>
            <div style="font-size:.82rem;color:#555;margin-top:6px;">
                Um ZIP por transportadora + PDF unificado Envios Extra
                pronto para impressão 2 páginas/folha frente e verso.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="tip">
    📌 <b>Dica de impressão do PDF unificado Envios Extra:</b>
    Imprima com <b>"2 páginas por folha"</b> + <b>"frente e verso"</b>.
    Cada romaneio tem número par de páginas, então ao dobrar e grampear
    cada bloco de folhas você obtém um romaneio individual completo.
    </div>
    """, unsafe_allow_html=True)

# ── PROCESSAMENTO ─────────────────────────────────────────────────────────────
if btn:
    if not arq_plan:
        st.error("⚠️ Faça upload da planilha de rotas.")
    elif not arqs_pdf:
        st.error("⚠️ Faça upload dos PDFs ou ZIP.")
    else:
        st.session_state.processado = False
        st.session_state.zips = {}
        st.session_state.unificado = None

        with st.spinner("Lendo planilha…"):
            rotas_info = carregar_planilha(arq_plan.read(), arq_plan.name)

        with st.spinner("Coletando PDFs…"):
            pdfs_dict = coletar_pdfs(arqs_pdf)

        if not pdfs_dict:
            st.error("Nenhum PDF encontrado nos arquivos enviados.")
        else:
            progress = st.progress(0, text="Iniciando…")
            tmpdir = tempfile.mkdtemp()
            try:
                stats, zips, unificado = processar_romaneios(
                    pdfs_dict, rotas_info, opcao, tmpdir, progress
                )
                progress.progress(1.0, text="✅ Concluído!")

                # Lê bytes dos ZIPs para armazenar no session_state
                zips_bytes = {}
                for transp, zpath in zips.items():
                    with open(zpath, "rb") as f:
                        zips_bytes[transp] = f.read()

                unificado_bytes = None
                if unificado and os.path.exists(unificado):
                    with open(unificado, "rb") as f:
                        unificado_bytes = f.read()

                st.session_state.stats     = stats
                st.session_state.zips      = zips_bytes
                st.session_state.unificado = unificado_bytes
                st.session_state.processado = True

            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)

        st.rerun()

# ── RESULTADOS ────────────────────────────────────────────────────────────────
if st.session_state.processado and st.session_state.stats:
    stats = st.session_state.stats
    zips  = st.session_state.zips
    unif  = st.session_state.unificado

    st.markdown("---")
    st.markdown("### 📊 Resultado do Processamento")

    # Métricas
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="card"><div class="card-label">Total de Romaneios</div>'
                    f'<div class="card-val">{stats["total"]}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="card"><div class="card-label">Transportadoras</div>'
                    f'<div class="card-val">{len(stats["transportadoras"])}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="card"><div class="card-label">📦 Envios Extra</div>'
                    f'<div class="card-val" style="color:#e94560">{stats["extra"]}</div></div>',
                    unsafe_allow_html=True)

    # Tabela resumo
    st.markdown("#### Romaneios por transportadora")
    df_res = pd.DataFrame(
        [(t, n) for t, n in stats["transportadoras"].items()],
        columns=["Transportadora", "Romaneios"]
    ).sort_values("Romaneios", ascending=False)
    st.dataframe(df_res, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### ⬇️ Downloads")

    # ZIP por transportadora
    st.markdown("**📁 ZIP por transportadora** (cada um com os PDFs anotados)")
    cols = st.columns(min(len(zips), 3))
    for i, (transp, zbytes) in enumerate(zips.items()):
        with cols[i % 3]:
            label = "📦 Envios Extra" if is_envios_extra(transp) else f"📋 {transp}"
            st.download_button(
                label=label,
                data=zbytes,
                file_name=f"{sanitize(transp)}_{agora_str()}.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # PDF unificado Envios Extra
    if unif:
        st.markdown("---")
        st.markdown("**🖨️ PDF Unificado Envios Extra — pronto para impressão 2 por folha**")
        st.markdown("""
        <div class="tip">
        Ao imprimir: selecione <b>"Múltiplas páginas por folha: 2"</b> + <b>"Frente e verso (virar pela borda longa)"</b>.
        Cada romaneio tem número par de páginas — ao dobrar e grampear cada bloco você obtém um romaneio individual.
        </div>
        """, unsafe_allow_html=True)
        st.download_button(
            label="📥 Baixar PDF Unificado Envios Extra",
            data=unif,
            file_name=f"ENVIOS_EXTRA_UNIFICADO_{agora_str()}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

st.divider()
st.caption(f"📦 Separador de Romaneios · {datetime.now(BRT).strftime('%d/%m/%Y %H:%M')} · Horário de Brasília")
