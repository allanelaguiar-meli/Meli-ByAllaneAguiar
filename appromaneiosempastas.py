import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red, black, white
import io, os, re, zipfile
from datetime import datetime
import pytz

st.set_page_config(page_title="Separador de Romaneios", page_icon="📦", layout="wide")
BRT = pytz.timezone("America/Sao_Paulo")

st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { background: #f4f6fb; }
    .titulo { font-size:2rem; font-weight:800; color:#1a1a2e; letter-spacing:-1px; }
    .subtitulo { color:#666; font-size:.9rem; margin-bottom:1rem; }
    .card { background:#fff; border-radius:14px; padding:18px 22px;
            box-shadow:0 2px 12px rgba(0,0,0,.07); margin-bottom:10px; }
    .card-label { font-size:.72rem; color:#999; text-transform:uppercase; letter-spacing:1px; }
    .card-val { font-size:1.8rem; font-weight:800; color:#1a1a2e; }
    .tip { background:#e8f4fd; border-left:4px solid #2196F3;
           border-radius:8px; padding:10px 16px; margin-bottom:1rem; font-size:.88rem; }
    .warn { background:#fff3cd; border-left:4px solid #ffc107;
            border-radius:8px; padding:10px 16px; margin-bottom:1rem; font-size:.88rem; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES DE ANOTAÇÃO (espelha o Colab original)
# ══════════════════════════════════════════════════════════════════════════════
MSG_TOPO   = "ATENÇÃO! DIARIAMENTE AO CHEGAR ESCANEIE O QR CODE NA MESA PARA INFORMAR SUA PRESENÇA"
MSG_RODAPE = "FAVOR NÃO DESCARTAR EM VIA PÚBLICA"

# Posições originais do Colab (pontos PDF, origem bottom-left)
# O Colab usava fitz com origem top-left; aqui convertemos para reportlab (bottom-left)
# Página A4 ≈ 842pt altura. fitz y → reportlab y = altura - fitz_y
# numero_x=730, numero_y=46  → rl_y = page_h - 46
# motorista_x=115, motorista_y=69 → rl_y = page_h - 69
# topo y=25 → rl_y = page_h - 25
# rodape y=page_h-20 → rl_y = 20

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def col_letra_para_idx(letra: str) -> int:
    letra = letra.strip().upper()
    idx = 0
    for c in letra:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1

def agora():
    return datetime.now(BRT)

def safe_str(s):
    if s is None: return ""
    try:
        import math
        if isinstance(s, float) and math.isnan(s): return ""
    except: pass
    return "" if str(s).lower() == "nan" else str(s).strip()

def sanitize(nome: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "-", nome).strip() or "_"

def is_envios_extra(nome: str) -> bool:
    n = nome.strip().upper().replace("Ç","C")
    n = re.sub(r'[\s_\-]+',' ', n)
    return any(x in n for x in ["ENVIO EXTRA","ENVIOS EXTRA","ENVIOS EXTRAS","ENVIO EXTRAS"])

def extrair_texto(reader: PdfReader, pg: int) -> str:
    try: return reader.pages[pg].extract_text() or ""
    except: return ""

def get_rota_id(texto: str):
    m = re.search(r'Rota\s+([A-Z0-9_]+)', texto)
    if m and m.group(1).count('_') == 1:
        return m.group(1)
    return None

def salvar_writer(writer: PdfWriter) -> bytes:
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()

def pad_para_par(pdf_bytes: bytes) -> bytes:
    """Garante número par de páginas — necessário para impressão 2/folha grampeável."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if len(reader.pages) % 2 == 0:
        return pdf_bytes
    try:
        w = float(reader.pages[-1].mediabox.width)
        h = float(reader.pages[-1].mediabox.height)
    except:
        w, h = 595, 842
    writer = PdfWriter()
    for pg in reader.pages:
        writer.add_page(pg)
    blank = PdfWriter()
    blank.add_blank_page(width=w, height=h)
    writer.add_page(blank.pages[0])
    return salvar_writer(writer)

def juntar_pdfs(lista_bytes: list) -> bytes:
    writer = PdfWriter()
    for b in lista_bytes:
        for pg in PdfReader(io.BytesIO(b)).pages:
            writer.add_page(pg)
    return salvar_writer(writer)

def coletar_pdfs(arquivos_up) -> dict:
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

# ══════════════════════════════════════════════════════════════════════════════
# ANOTAÇÃO DO PDF (substitui fitz usando reportlab + pypdf overlay)
# ══════════════════════════════════════════════════════════════════════════════
def criar_overlay(page_w, page_h, romaneio, motorista, is_extra, is_first_page):
    """
    Cria uma página PDF transparente com as anotações usando reportlab.
    Retorna bytes do PDF overlay de 1 página.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    # ── Número do romaneio (todas as páginas, canto superior direito) ──
    # Cobre a área original com branco
    c.setFillColor(white)
    c.rect(729, page_h - 58, 71, 22, fill=1, stroke=0)
    # Escreve o número
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(730, page_h - 50, str(romaneio))

    if is_first_page:
        # ── Nome do motorista (só Envios Extra, só 1ª página) ──
        if is_extra and motorista:
            c.setFillColor(white)
            c.rect(110, page_h - 87, 221, 23, fill=1, stroke=0)
            c.setFillColor(black)
            c.setFont("Helvetica-Bold", 16)
            c.drawString(115, page_h - 79, str(motorista))

        # ── Mensagens Envios Extra (só 1ª página) ──
        if is_extra:
            if MSG_TOPO:
                c.setFillColor(red)
                c.setFont("Helvetica-Bold", 11)
                c.drawString(60, page_h - 20, MSG_TOPO)
            if MSG_RODAPE:
                c.setFillColor(red)
                c.setFont("Helvetica-Bold", 11)
                c.drawString(60, 12, MSG_RODAPE)

    c.save()
    buf.seek(0)
    return buf.read()

def anotar_pdf(pdf_bytes: bytes, romaneio: str, motorista: str, is_extra: bool) -> bytes:
    """
    Aplica overlay de anotações em cada página do romaneio.
    Reproduz exatamente o comportamento do fitz original.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        try:
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
        except:
            w, h = 595, 842

        overlay_bytes = criar_overlay(w, h, romaneio, motorista, is_extra, is_first_page=(i == 0))
        overlay_reader = PdfReader(io.BytesIO(overlay_bytes))
        overlay_page   = overlay_reader.pages[0]

        # Merge: overlay por cima da página original
        page.merge_page(overlay_page)
        writer.add_page(page)

    return salvar_writer(writer)

# ══════════════════════════════════════════════════════════════════════════════
# LEITURA DA PLANILHA
# ══════════════════════════════════════════════════════════════════════════════
def carregar_planilha(arq_bytes, nome_arq, cfg: dict):
    if nome_arq.lower().endswith('.csv'):
        df = pd.read_csv(io.BytesIO(arq_bytes), header=None)
    else:
        try:
            df = pd.read_excel(io.BytesIO(arq_bytes), sheet_name='PLAN', header=None)
        except Exception as e:
            st.error(f"Erro ao ler aba PLAN: {e}")
            return {}

    idx_rota = cfg["idx_rota"]
    # Pula linhas até achar string na coluna de rota
    skip = 0
    while skip < 20:
        val = df.iloc[skip, idx_rota] if idx_rota < df.shape[1] else None
        if isinstance(val, str) and re.search(r'[A-Z]{2}_\d+', val.strip()):
            break
        skip += 1
    if skip > 0:
        df = df.iloc[skip:].reset_index(drop=True)

    rotas = {}
    for row in df.values.tolist():
        def get(idx):
            return safe_str(row[idx]) if idx is not None and len(row) > idx else ""
        rota      = get(cfg["idx_rota"])
        romaneio  = get(cfg["idx_qr"])
        transp    = get(cfg["idx_transp"])
        motorista = get(cfg["idx_driver"])
        ciclo     = get(cfg["idx_ciclo"])
        if rota and re.search(r'[A-Z]{2}_\d+', rota):
            rotas[rota] = {
                "TRANSPORTADORA": transp,
                "ROMANEIO": romaneio,
                "MOTORISTA": motorista,
                "CICLO": ciclo,
            }
    return rotas

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSAMENTO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def processar(pdfs_dict, rotas_info, opcao, progress_bar):
    hoje = agora().strftime("%d-%m-%Y")
    ciclos_encontrados = set()

    # mlps[transp] = [(nome, bytes)]
    mlps   = {}
    extras = []   # [(nome, bytes)]
    stats  = {"total": 0, "extra": 0, "mlp": 0, "transportadoras": {}, "sem_mapa": 0}

    todos = list(pdfs_dict.items())
    for idx, (filename, filedata) in enumerate(todos):
        progress_bar.progress((idx + 1) / len(todos), text=f"Processando {filename}…")
        reader  = PdfReader(io.BytesIO(filedata))
        n_pages = len(reader.pages)

        starts, rota_ids = [], []
        for i in range(n_pages):
            txt = extrair_texto(reader, i)
            if "Roteiro" in txt:
                rid = get_rota_id(txt)
                starts.append(i)
                rota_ids.append(rid)
        starts.append(n_pages)

        for j in range(len(starts) - 1):
            rota_id = rota_ids[j]
            if not rota_id:
                continue

            info = rotas_info.get(rota_id)
            if not info:
                stats["sem_mapa"] += 1
                continue

            transp    = sanitize(info["TRANSPORTADORA"]) or "SEM_TRANSPORTADORA"
            romaneio  = sanitize(info["ROMANEIO"]) or "_"
            motorista = info["MOTORISTA"]
            ciclo     = sanitize(info["CICLO"]) or "SEM_CICLO"
            extra     = is_envios_extra(transp)

            ciclos_encontrados.add(ciclo)

            if opcao == "mlp"   and extra:     continue
            if opcao == "extra" and not extra: continue

            # Extrai páginas do romaneio
            writer = PdfWriter()
            for pg in range(starts[j], starts[j + 1]):
                writer.add_page(reader.pages[pg])
            rom_bytes = salvar_writer(writer)

            # ── Anota: número, motorista, mensagens (igual ao Colab original) ──
            rom_bytes = anotar_pdf(rom_bytes, romaneio, motorista, is_extra=extra)

            # ── Pad para par (só para unificado de impressão) ──
            rom_bytes_padded = pad_para_par(rom_bytes)

            nome_arq = f"{transp}-{rota_id}-{romaneio}.pdf"
            stats["total"] += 1

            if extra:
                # Arquivo individual SEM padding (fiel ao original)
                extras.append((nome_arq, rom_bytes, rom_bytes_padded))
                stats["extra"] += 1
            else:
                if transp not in mlps:
                    mlps[transp] = []
                mlps[transp].append((nome_arq, rom_bytes))
                stats["mlp"] += 1
                stats["transportadoras"][transp] = stats["transportadoras"].get(transp, 0) + 1

    ciclo_str  = "_".join(sorted(ciclos_encontrados)) if ciclos_encontrados else "SEM_CICLO"
    pasta_raiz = f"Romaneios_{ciclo_str}_{hoje}"

    # ── Monta ZIP ────────────────────────────────────────────────────────────
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:

        # MLPs: cada transportadora numa subpasta, tudo dentro de MLPs/
        if opcao in ("ambos", "mlp") and mlps:
            # ZIP interno de MLPs
            mlp_buf = io.BytesIO()
            with zipfile.ZipFile(mlp_buf, "w", zipfile.ZIP_DEFLATED) as zf_mlp:
                for transp, roms in mlps.items():
                    for nome_arq, rom_bytes in roms:
                        zf_mlp.writestr(f"{transp}/{nome_arq}", rom_bytes)
            zf.writestr(f"{pasta_raiz}/MLPs_{ciclo_str}_{hoje}.zip", mlp_buf.getvalue())

        # Envios Extra: individuais + unificado 2por1
        if opcao in ("ambos", "extra") and extras:
            extras_padded = []
            for nome_arq, rom_bytes, rom_padded in extras:
                # Arquivo individual (sem padding, igual ao original)
                zf.writestr(f"{pasta_raiz}/Envios_Extra/{nome_arq}", rom_bytes)
                extras_padded.append(rom_padded)

            # PDF unificado com todos padded (para impressão 2/folha)
            if extras_padded:
                unif = juntar_pdfs(extras_padded)
                zf.writestr(
                    f"{pasta_raiz}/Envios_Extra/ENVIOS_EXTRA_UNIFICADO_2por1.pdf",
                    unif
                )

    return stats, zip_buf.getvalue(), pasta_raiz

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
defaults = {
    "processado": False, "stats": {}, "zip_bytes": None, "pasta_raiz": "",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="titulo">📦 Separador de Romaneios</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Separa PDFs por transportadora · pastas por ciclo e data</div>',
            unsafe_allow_html=True)

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuração")

    st.markdown("**1 · Planilha de rotas**")
    st.caption("xlsx (aba PLAN) ou csv")
    arq_plan = st.file_uploader("Planilha", type=["xlsx","csv"], label_visibility="collapsed")

    if arq_plan:
        try:
            plan_bytes = arq_plan.read()
            arq_plan.seek(0)
            if arq_plan.name.lower().endswith('.csv'):
                df_full = pd.read_csv(io.BytesIO(plan_bytes), header=None)
            else:
                df_full = pd.read_excel(io.BytesIO(plan_bytes), sheet_name='PLAN', header=None)

            # Prévia: mostra da linha 5 em diante (índice 4), primeiras 6 linhas visíveis
            df_prev = df_full.iloc[4:10].reset_index(drop=True)
            letras = []
            for i in range(df_prev.shape[1]):
                if i < 26:
                    letras.append(chr(65 + i))
                else:
                    letras.append(chr(64 + i//26) + chr(65 + i%26))
            df_prev.columns = letras
            st.caption("📋 Prévia (a partir da linha 5):")
            st.dataframe(df_prev, use_container_width=True, hide_index=True)
        except Exception as e:
            st.warning(f"Prévia indisponível: {e}")

    st.divider()
    st.markdown("**2 · Colunas** (letra Excel)")
    c1, c2 = st.columns(2)
    with c1:
        col_rota   = st.text_input("Rota",          value="D", max_chars=3).upper()
        col_qr     = st.text_input("QR/Romaneio",   value="E", max_chars=3).upper()
        col_ciclo  = st.text_input("Ciclo",          value="C", max_chars=3).upper()
    with c2:
        col_transp = st.text_input("Transportadora", value="P", max_chars=3).upper()
        col_driver = st.text_input("Motorista",      value="Q", max_chars=3).upper()

    st.divider()
    st.markdown("**3 · PDFs ou ZIP**")
    arqs_pdf = st.file_uploader("PDFs/ZIP", type=["pdf","zip"],
                                accept_multiple_files=True, label_visibility="collapsed")

    st.divider()
    st.markdown("**4 · O que separar?**")
    opcao = st.radio("", options=["ambos","mlp","extra"],
        format_func=lambda x: {
            "ambos": "🔄 MLPs + Envios Extra",
            "mlp":   "📋 Só MLPs",
            "extra": "📦 Só Envios Extra"
        }[x], label_visibility="collapsed")

    st.divider()
    btn = st.button("▶️ Processar", use_container_width=True, type="primary")

# ── INSTRUÇÕES ────────────────────────────────────────────────────────────────
if not st.session_state.processado:
    st.markdown("""
    <div class="tip">
    📁 <b>Estrutura gerada no ZIP:</b><br>
    <code>Romaneios_SD_29-03-2025/</code><br>
    &nbsp;&nbsp;├── <b>MLPs_SD_29-03-2025.zip</b> → subpastas por transportadora<br>
    &nbsp;&nbsp;└── <b>Envios_Extra/</b><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├── romaneios individuais anotados<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── ENVIOS_EXTRA_UNIFICADO_2por1.pdf
    </div>
    <div class="tip">
    ✏️ <b>Anotações nos PDFs Envios Extra</b> (igual ao Colab):<br>
    • Número do romaneio no canto superior direito<br>
    • Nome do motorista na 1ª página<br>
    • Mensagem de atenção no topo (vermelha)<br>
    • Mensagem de rodapé (vermelha)
    </div>
    <div class="tip">
    🖨️ <b>Impressão do Unificado:</b> "2 páginas por folha" + "frente e verso — borda longa".
    Cada romaneio tem n° par de páginas → dobre e grampeie individualmente.
    </div>
    """, unsafe_allow_html=True)

# ── PROCESSAMENTO ─────────────────────────────────────────────────────────────
if btn:
    erros = []
    if not arq_plan: erros.append("Planilha não enviada.")
    if not arqs_pdf: erros.append("PDFs ou ZIP não enviados.")
    if erros:
        for e in erros: st.error(f"⚠️ {e}")
    else:
        st.session_state.processado = False
        cfg = {
            "idx_rota":   col_letra_para_idx(col_rota),
            "idx_qr":     col_letra_para_idx(col_qr)     if col_qr     else None,
            "idx_transp": col_letra_para_idx(col_transp),
            "idx_driver": col_letra_para_idx(col_driver) if col_driver else None,
            "idx_ciclo":  col_letra_para_idx(col_ciclo)  if col_ciclo  else None,
        }
        with st.spinner("Lendo planilha…"):
            arq_plan.seek(0)
            rotas_info = carregar_planilha(arq_plan.read(), arq_plan.name, cfg)
            st.toast(f"✅ {len(rotas_info)} rotas mapeadas")

        with st.spinner("Coletando PDFs…"):
            pdfs_dict = coletar_pdfs(arqs_pdf)
            st.toast(f"✅ {len(pdfs_dict)} PDFs encontrados")

        if not pdfs_dict:
            st.error("Nenhum PDF encontrado.")
        elif not rotas_info:
            st.error("Nenhuma rota encontrada. Verifique as colunas informadas.")
        else:
            progress = st.progress(0, text="Iniciando…")
            try:
                stats, zip_bytes, pasta_raiz = processar(pdfs_dict, rotas_info, opcao, progress)
                progress.progress(1.0, text="✅ Concluído!")
                st.session_state.stats      = stats
                st.session_state.zip_bytes  = zip_bytes
                st.session_state.pasta_raiz = pasta_raiz
                st.session_state.processado = True
            except Exception as e:
                st.error(f"Erro: {e}")
                raise
        st.rerun()

# ── RESULTADO ─────────────────────────────────────────────────────────────────
if st.session_state.processado and st.session_state.stats:
    stats      = st.session_state.stats
    zip_bytes  = st.session_state.zip_bytes
    pasta_raiz = st.session_state.pasta_raiz

    st.markdown("---")
    st.markdown("### 📊 Resultado")

    c1, c2, c3, c4 = st.columns(4)
    for col, label, val, cor in zip(
        [c1,c2,c3,c4],
        ["Total","📋 MLPs","📦 Envios Extra","❓ Sem mapeamento"],
        [stats["total"], stats["mlp"], stats["extra"], stats.get("sem_mapa",0)],
        ["#1a1a2e","#155724","#e94560","#856404"]
    ):
        with col:
            st.markdown(f'<div class="card"><div class="card-label">{label}</div>'
                        f'<div class="card-val" style="color:{cor}">{val}</div></div>',
                        unsafe_allow_html=True)

    if stats.get("transportadoras"):
        with st.expander("Ver romaneios por transportadora (MLPs)"):
            df_t = pd.DataFrame(
                [(t, n) for t, n in stats["transportadoras"].items()],
                columns=["Transportadora","Qtd"]
            ).sort_values("Qtd", ascending=False)
            st.dataframe(df_t, use_container_width=True, hide_index=True)

    if stats.get("sem_mapa", 0) > 0:
        st.markdown(f'<div class="warn">⚠️ <b>{stats["sem_mapa"]} romaneios</b> do PDF sem correspondência na planilha. '
                    'Verifique se as colunas estão corretas.</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.download_button(
        label=f"📥 Baixar {pasta_raiz}.zip",
        data=zip_bytes,
        file_name=f"{pasta_raiz}.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )

st.divider()
st.caption(f"📦 Separador de Romaneios · {agora().strftime('%d/%m/%Y %H:%M')} · Brasília")
