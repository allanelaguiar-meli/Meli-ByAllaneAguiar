import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
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
    .secao { font-size:1.1rem; font-weight:700; color:#1a1a2e; margin:1rem 0 .4rem 0; }
    div[data-testid="stExpander"] { background:#fff; border-radius:12px; border:1px solid #e0e0e0; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def col_letra_para_idx(letra: str) -> int:
    """Converte letra de coluna Excel (A, B, ... Z, AA...) para índice 0-based."""
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
    """Garante número par de páginas adicionando branca se necessário."""
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
# LEITURA DA PLANILHA
# ══════════════════════════════════════════════════════════════════════════════
def carregar_planilha(arq_bytes, nome_arq, cfg: dict):
    """Retorna dict {rota_id: {TRANSPORTADORA, ROMANEIO, MOTORISTA, CICLO}}"""
    if nome_arq.lower().endswith('.csv'):
        df = pd.read_csv(io.BytesIO(arq_bytes), header=None)
    else:
        try:
            df = pd.read_excel(io.BytesIO(arq_bytes), sheet_name='PLAN', header=None)
        except Exception as e:
            st.error(f"Erro ao ler aba PLAN: {e}")
            return {}

    # Pula linhas de cabeçalho até achar string na coluna de rota
    idx_rota = cfg["idx_rota"]
    skip = 0
    while skip < 15:
        val = df.iloc[skip, idx_rota] if idx_rota < df.shape[1] else None
        if isinstance(val, str) and val.strip():
            break
        skip += 1
    if skip > 0:
        df = df.iloc[skip:].reset_index(drop=True)

    rotas = {}
    for row in df.values.tolist():
        def get(idx):
            return safe_str(row[idx]) if idx is not None and len(row) > idx else ""
        rota     = get(cfg["idx_rota"])
        romaneio = get(cfg["idx_qr"])
        transp   = get(cfg["idx_transp"])
        motorist = get(cfg["idx_driver"])
        ciclo    = get(cfg["idx_ciclo"])
        if rota:
            rotas[rota] = {
                "TRANSPORTADORA": transp,
                "ROMANEIO": romaneio,
                "MOTORISTA": motorist,
                "CICLO": ciclo,
            }
    return rotas

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSAMENTO
# ══════════════════════════════════════════════════════════════════════════════
def processar(pdfs_dict, rotas_info, opcao, progress_bar):
    """
    Estrutura de saída no ZIP:
      Romaneios_<CICLO>_<DATA>/
        MLPs/
          <Transportadora>/
            romaneio.pdf ...
        Envios_Extra/
          romaneio.pdf ...
          ENVIOS_EXTRA_UNIFICADO.pdf
    """
    hoje = agora().strftime("%d-%m-%Y")

    # Coleta ciclos presentes (usa o primeiro encontrado como nome da pasta raiz)
    ciclos_encontrados = set()

    # Estrutura: mlps[transp] = [(nome, bytes)]  /  extras = [(nome, bytes)]
    mlps   = {}   # transp → [(nome_arq, bytes)]
    extras = []   # [(nome_arq, bytes)]
    stats  = {"total": 0, "extra": 0, "mlp": 0, "transportadoras": {}, "sem_mapa": 0}

    todos = list(pdfs_dict.items())
    for idx, (filename, filedata) in enumerate(todos):
        progress_bar.progress((idx + 1) / len(todos), text=f"Lendo {filename}…")
        reader = PdfReader(io.BytesIO(filedata))
        n = len(reader.pages)

        starts, rota_ids = [], []
        for i in range(n):
            txt = extrair_texto(reader, i)
            if "Roteiro" in txt:
                rid = get_rota_id(txt)
                starts.append(i)
                rota_ids.append(rid)
        starts.append(n)

        for j in range(len(starts) - 1):
            rota_id = rota_ids[j]
            if not rota_id:
                continue

            info     = rotas_info.get(rota_id)
            if not info:
                stats["sem_mapa"] += 1
                continue

            transp   = sanitize(info["TRANSPORTADORA"]) or "SEM_TRANSPORTADORA"
            romaneio = sanitize(info["ROMANEIO"]) or "_"
            ciclo    = sanitize(info["CICLO"]) or "SEM_CICLO"
            extra    = is_envios_extra(transp)

            ciclos_encontrados.add(ciclo)

            if opcao == "mlp"   and extra:     continue
            if opcao == "extra" and not extra: continue

            # Extrai páginas
            writer = PdfWriter()
            for pg in range(starts[j], starts[j + 1]):
                writer.add_page(reader.pages[pg])
            rom_bytes = salvar_writer(writer)
            rom_bytes = pad_para_par(rom_bytes)

            nome_arq = f"{transp}-{rota_id}-{romaneio}.pdf"
            stats["total"] += 1

            if extra:
                extras.append((nome_arq, rom_bytes))
                stats["extra"] += 1
            else:
                if transp not in mlps:
                    mlps[transp] = []
                mlps[transp].append((nome_arq, rom_bytes))
                stats["mlp"] += 1
                stats["transportadoras"][transp] = stats["transportadoras"].get(transp, 0) + 1

    # Define nome da pasta raiz
    ciclo_str = "_".join(sorted(ciclos_encontrados)) if ciclos_encontrados else "SEM_CICLO"
    pasta_raiz = f"Romaneios_{ciclo_str}_{hoje}"

    # Monta ZIP em memória
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:

        # MLPs → subpastas por transportadora
        if opcao in ("ambos", "mlp"):
            for transp, roms in mlps.items():
                for nome_arq, rom_bytes in roms:
                    caminho = f"{pasta_raiz}/MLPs/{transp}/{nome_arq}"
                    zf.writestr(caminho, rom_bytes)

        # Envios Extra → pasta própria + unificado
        if opcao in ("ambos", "extra") and extras:
            extras_bytes_lista = []
            for nome_arq, rom_bytes in extras:
                caminho = f"{pasta_raiz}/Envios_Extra/{nome_arq}"
                zf.writestr(caminho, rom_bytes)
                extras_bytes_lista.append(rom_bytes)

            if extras_bytes_lista:
                unif = juntar_pdfs(extras_bytes_lista)
                zf.writestr(
                    f"{pasta_raiz}/Envios_Extra/ENVIOS_EXTRA_UNIFICADO_2por1.pdf",
                    unif
                )

    return stats, zip_buf.getvalue(), pasta_raiz

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
defaults = {
    "cfg_ok": False,
    "cfg": {},
    "rotas_info": {},
    "stats": {},
    "zip_bytes": None,
    "pasta_raiz": "",
    "processado": False,
    "preview_cols": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="titulo">📦 Separador de Romaneios</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Separa PDFs por transportadora com pastas organizadas por ciclo e data</div>',
            unsafe_allow_html=True)

# ════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Configuração")

    # ── PASSO 1: Planilha ──────────────────
    st.markdown("**1 · Planilha de rotas**")
    st.caption("xlsx (aba PLAN) ou csv")
    arq_plan = st.file_uploader("Planilha", type=["xlsx","csv"], label_visibility="collapsed")

    if arq_plan:
        # Preview para ajudar usuário a identificar colunas
        try:
            plan_bytes = arq_plan.read()
            arq_plan.seek(0)
            if arq_plan.name.lower().endswith('.csv'):
                df_prev = pd.read_csv(io.BytesIO(plan_bytes), header=None, nrows=6)
            else:
                df_prev = pd.read_excel(io.BytesIO(plan_bytes), sheet_name='PLAN', header=None, nrows=6)
            # Cria cabeçalho com letras A, B, C...
            letras = [chr(65 + i) if i < 26 else chr(64 + i//26) + chr(65 + i%26)
                      for i in range(df_prev.shape[1])]
            df_prev.columns = letras
            st.caption("📋 Prévia da planilha (primeiras linhas):")
            st.dataframe(df_prev, use_container_width=True, hide_index=True)
            st.session_state.preview_cols = letras
        except Exception as e:
            st.warning(f"Não foi possível gerar prévia: {e}")

    st.divider()

    # ── PASSO 2: Configuração das colunas ──
    st.markdown("**2 · Mapeamento de colunas**")
    st.caption("Informe a letra de cada coluna (ex: A, B, D, P...)")

    c1, c2 = st.columns(2)
    with c1:
        col_rota   = st.text_input("Rota",          value="D", max_chars=3).upper()
        col_qr     = st.text_input("QR/Romaneio",   value="E", max_chars=3).upper()
        col_ciclo  = st.text_input("Ciclo",          value="C", max_chars=3).upper()
    with c2:
        col_transp = st.text_input("Transportadora", value="P", max_chars=3).upper()
        col_driver = st.text_input("Motorista",      value="Q", max_chars=3).upper()

    st.divider()

    # ── PASSO 3: PDFs ──────────────────────
    st.markdown("**3 · PDFs ou ZIP**")
    arqs_pdf = st.file_uploader("PDFs/ZIP", type=["pdf","zip"],
                                accept_multiple_files=True, label_visibility="collapsed")

    st.divider()

    # ── PASSO 4: Opção ─────────────────────
    st.markdown("**4 · O que separar?**")
    opcao = st.radio("", options=["ambos","mlp","extra"],
        format_func=lambda x: {
            "ambos": "🔄 MLPs + Envios Extra",
            "mlp":   "📋 Só MLPs",
            "extra": "📦 Só Envios Extra"
        }[x], label_visibility="collapsed")

    st.divider()
    btn = st.button("▶️ Processar", use_container_width=True, type="primary")

# ════════════════════════════════════════════
# INSTRUÇÕES (antes de processar)
# ════════════════════════════════════════════
if not st.session_state.processado:
    st.markdown("""
    <div class="tip">
    📁 <b>Estrutura gerada no ZIP:</b><br>
    <code>Romaneios_SD_29-03-2025/</code><br>
    &nbsp;&nbsp;├── <b>MLPs/</b><br>
    &nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;├── Transportadora_A/ → romaneios...<br>
    &nbsp;&nbsp;│&nbsp;&nbsp;&nbsp;└── Transportadora_B/ → romaneios...<br>
    &nbsp;&nbsp;└── <b>Envios_Extra/</b><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;├── romaneios individuais...<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── ENVIOS_EXTRA_UNIFICADO_2por1.pdf
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="tip">
    🖨️ <b>Como imprimir o PDF Unificado Envios Extra:</b>
    Selecione <b>"2 páginas por folha"</b> + <b>"frente e verso — borda longa"</b>.
    Cada romaneio tem número <b>par</b> de páginas → dobre e grampeie cada bloco individualmente.
    </div>
    """, unsafe_allow_html=True)

# ════════════════════════════════════════════
# PROCESSAMENTO
# ════════════════════════════════════════════
if btn:
    erros = []
    if not arq_plan:    erros.append("Planilha de rotas não enviada.")
    if not arqs_pdf:    erros.append("PDFs ou ZIP não enviados.")
    if not col_rota:    erros.append("Coluna Rota não informada.")
    if not col_transp:  erros.append("Coluna Transportadora não informada.")

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
            st.toast(f"✅ {len(rotas_info)} rotas carregadas")

        with st.spinner("Coletando PDFs…"):
            pdfs_dict = coletar_pdfs(arqs_pdf)
            st.toast(f"✅ {len(pdfs_dict)} PDFs encontrados")

        if not pdfs_dict:
            st.error("Nenhum PDF encontrado.")
        elif not rotas_info:
            st.error("Nenhuma rota encontrada na planilha. Verifique as colunas informadas.")
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

# ════════════════════════════════════════════
# RESULTADO
# ════════════════════════════════════════════
if st.session_state.processado and st.session_state.stats:
    stats      = st.session_state.stats
    zip_bytes  = st.session_state.zip_bytes
    pasta_raiz = st.session_state.pasta_raiz

    st.markdown("---")
    st.markdown("### 📊 Resultado")

    c1, c2, c3, c4 = st.columns(4)
    cards = [
        ("Total Romaneios",   stats["total"],                    "#1a1a2e"),
        ("📋 MLPs",           stats["mlp"],                      "#155724"),
        ("📦 Envios Extra",   stats["extra"],                    "#e94560"),
        ("❓ Sem mapeamento", stats.get("sem_mapa", 0),          "#856404"),
    ]
    for col, (label, val, cor) in zip([c1,c2,c3,c4], cards):
        with col:
            st.markdown(f'<div class="card"><div class="card-label">{label}</div>'
                        f'<div class="card-val" style="color:{cor}">{val}</div></div>',
                        unsafe_allow_html=True)

    if stats.get("transportadoras"):
        with st.expander("📋 Romaneios por transportadora (MLPs)", expanded=False):
            df_t = pd.DataFrame(
                [(t, n) for t, n in stats["transportadoras"].items()],
                columns=["Transportadora","Romaneios"]
            ).sort_values("Romaneios", ascending=False)
            st.dataframe(df_t, use_container_width=True, hide_index=True)

    if stats.get("sem_mapa", 0) > 0:
        st.markdown(f'<div class="warn">⚠️ <b>{stats["sem_mapa"]} romaneios</b> encontrados no PDF mas sem correspondência na planilha. '
                    'Verifique se as colunas estão corretas.</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(f"### ⬇️ Baixar ZIP — `{pasta_raiz}`")
    st.markdown("""
    <div class="tip">
    O ZIP contém a estrutura completa:<br>
    <b>MLPs/</b> com subpastas por transportadora + <b>Envios_Extra/</b> com PDFs individuais e o unificado 2/folha.
    </div>
    """, unsafe_allow_html=True)

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
