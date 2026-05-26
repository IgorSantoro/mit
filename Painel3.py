import streamlit as st
from openpyxl import load_workbook
import os
import json
import logging
import zipfile
import io
from datetime import datetime
import pandas as pd
import re
from collections import defaultdict

# --------------------------
# CONFIG
# --------------------------
st.set_page_config(
    page_title="MIT — Apuração",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --------------------------
# DESIGN — CSS MODERNO
# --------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* ── Fundo branco padrão Streamlit ── */
.stApp {
    background: #ffffff;
    color: #1a1a2e;
}

/* ── Cabeçalho principal ── */
.mit-header {
    background: linear-gradient(135deg, #f0f4ff 0%, #ffffff 100%);
    border: 1px solid #dde3f0;
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.mit-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #1a56db;
    letter-spacing: -0.5px;
    margin: 0;
}

.mit-subtitle {
    font-size: 0.8rem;
    color: #6b7280;
    margin-top: 4px;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Empresa card ── */
.empresa-card {
    background: #f8faff;
    border: 1px solid #dde3f0;
    border-left: 3px solid #1a56db;
    border-radius: 8px;
    padding: 20px 24px;
    margin-bottom: 20px;
}

.empresa-nome {
    font-size: 1.1rem;
    font-weight: 600;
    color: #111827;
    margin: 0 0 6px 0;
}

.empresa-cnpj {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    color: #6b7280;
}

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    font-family: 'IBM Plex Mono', monospace;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 8px;
}

.badge-sem-mov {
    background: #fffbeb;
    color: #92400e;
    border: 1px solid #f59e0b;
}

.badge-com-mov {
    background: #f0fdf4;
    color: #166534;
    border: 1px solid #22c55e;
}

/* ── Progresso ── */
.progress-bar-container {
    background: #e5e7eb;
    border-radius: 4px;
    height: 4px;
    margin-bottom: 20px;
    overflow: hidden;
}

.progress-bar-fill {
    background: linear-gradient(90deg, #1a56db, #0ea5e9);
    height: 100%;
    border-radius: 4px;
    transition: width 0.3s ease;
}

/* ── Tabela de débitos ── */
.debito-header {
    display: grid;
    grid-template-columns: 40px 1fr 1fr 1fr 1fr;
    gap: 12px;
    padding: 8px 16px;
    background: #f1f5f9;
    border: 1px solid #e2e8f0;
    border-radius: 6px 6px 0 0;
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #6b7280;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}

.codigo-chip {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    color: #1a56db;
    font-weight: 600;
}

.valor-text {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.9rem;
    color: #166534;
}

.periodo-text {
    font-size: 0.85rem;
    color: #6b7280;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Seção responsável ── */
.section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #1a56db;
    margin: 24px 0 16px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid #e2e8f0;
}

/* ── Contador de navegação ── */
.nav-counter {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    color: #6b7280;
    text-align: center;
    padding: 4px 0;
}

/* ── Empty state ── */
.empty-state {
    text-align: center;
    padding: 40px 20px;
    color: #9ca3af;
    font-size: 0.85rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Inputs ── */
.stTextInput label {
    font-size: 0.72rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
    color: #6b7280 !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 500 !important;
}

/* ── Divisor ── */
hr {
    border-color: #e2e8f0 !important;
    margin: 20px 0 !important;
}
</style>
""", unsafe_allow_html=True)

# --------------------------
# LOGGING
# --------------------------
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# --------------------------
# CONSTANTES
# --------------------------
CAMINHO = "base.xlsx"
CAMINHO_PERDCOMP = "PERDCOMP"

MAPEAMENTO_CODIGOS = {
    "IRPJ":    ["0220","0231","0507","1599","2089","2362","2430","2456","3373","5625","5993","7756"],
    "CSLL":    ["2030","2372","2484","6012","6758","6773","7837"],
    "IRRF":    ["7769"],
    "IPI":     ["0668","0676","0821","0838","1020","1097","2401","2410","5110","5123"],
    "IOF":     ["1150","2927","3467","4028","4290","5220","6854","6895","7893"],
    "PIS":     ["0679","0691","0906","1921","3703","4574","5434","6824","6912","7797","8109"],
    "COFINS":  ["0760","0776","0929","1840","2172","5442","5856","6840","7784","7987","8645"],
    "CONTRIB": ["8536","8741","9013","9331","9197"],
    "CPSS":    ["1661","1684","1700","1717","1723","1730","1752","1769","1781","1814","1837","5492","5502","5519"],
    "RET/UNIF":["1068","4095","4112","4138","4153","4166","6177"],
}

MAPA_NOME_JSON = {
    "IRPJ":   "Irpj",
    "CSLL":   "Csll",
    "IRRF":   "Irrf",
    "PIS":    "PisPasep",
    "COFINS": "Cofins",
}

ORDEM_JSON = ["IRPJ", "CSLL", "PIS", "COFINS", "IRRF"]

# --------------------------
# FUNÇÕES UTILITÁRIAS
# --------------------------
def normalizar(x: str) -> str:
    return str(x).strip().lower()

def get_aba(wb, nome: str):
    nome_norm = normalizar(nome)
    for n in wb.sheetnames:
        if nome_norm == normalizar(n):
            return wb[n]
    for n in wb.sheetnames:
        if nome_norm in normalizar(n):
            return wb[n]
    return None

def limpar_cnpj(cnpj) -> str:
    return ''.join(filter(str.isdigit, str(cnpj))).zfill(14)

def normalizar_codigo(cod) -> str:
    return re.sub(r"\D", "", str(cod))

def formatar_valor(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --------------------------
# VALIDAÇÃO
# --------------------------
if not os.path.exists(CAMINHO):
    st.error("⚠️ Arquivo **base.xlsx** não encontrado no diretório de trabalho.")
    st.stop()

# --------------------------
# SESSION STATE
# --------------------------
if "indice_empresa" not in st.session_state:
    st.session_state.indice_empresa = 0

# --------------------------
# CARREGAMENTO — WORKBOOK
# --------------------------
@st.cache_resource(show_spinner=False)
def carregar_workbook(caminho: str):
    return load_workbook(caminho, data_only=True)

wb = carregar_workbook(CAMINHO)

# --------------------------
# CARREGAMENTO — EMPRESAS
# --------------------------
@st.cache_data(show_spinner=False)
def carregar_empresas(caminho: str) -> list:
    wb_ = load_workbook(caminho, data_only=True)
    sheet = get_aba(wb_, "dados")
    empresas = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        empresas.append({
            "empresa":      str(row[1]).strip(),
            "cnpj":         limpar_cnpj(row[2]),
            "tributacao":   row[3],
            "qualificacao": row[4],
            "criterio":     row[5],
            "regime":       row[6],
        })
    return empresas

# --------------------------
# CARREGAMENTO — RESPONSÁVEL
# --------------------------
@st.cache_data(show_spinner=False)
def carregar_responsavel(caminho: str) -> dict:
    wb_ = load_workbook(caminho, data_only=True)
    sheet = get_aba(wb_, "responsavel")
    row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]

    def get(i):
        return str(row[i]).strip() if len(row) > i and row[i] else ""

    return {
        "cpf":      get(1).zfill(11),
        "nome":     get(9),
        "telefone": get(5),
        "email":    get(11),
        "crc":      get(13),
        "uf_crc":   get(7),
        "ddd":      get(3),
    }

# --------------------------
# CARREGAMENTO — SIDAF
# --------------------------
@st.cache_data(show_spinner=False)
def carregar_sidaf(caminho: str) -> list:
    wb_ = load_workbook(caminho, data_only=True)
    sheet = get_aba(wb_, "sidaf")
    sidaf = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        empresa_texto = str(row[0]) if row[0] else ""
        codigo        = normalizar_codigo(row[8]) if row[8] else ""
        valor         = row[20]
        situacao      = str(row[13]).strip().lower() if len(row) > 13 and row[13] else ""

        # Ignora registros cancelados (coluna N)
        if situacao == "cancelada":
            continue

        if not empresa_texto or not codigo or valor is None:
            continue
        try:
            sidaf.append({
                "empresa": empresa_texto,
                "codigo":  codigo,
                "valor":   float(valor),
            })
        except (ValueError, TypeError) as e:
            logger.warning(f"SIDAF — valor inválido na linha: {e}")
    return sidaf

# --------------------------
# CARREGAMENTO — PERDCOMP
# --------------------------
@st.cache_data(show_spinner=False)
def carregar_perdcomp(caminho_dir: str) -> list:
    lista = []
    if not os.path.exists(caminho_dir):
        return lista

    for arq in os.listdir(caminho_dir):
        caminho = os.path.join(caminho_dir, arq)

        # Excel
        if arq.lower().endswith(".xlsx"):
            try:
                df = pd.read_excel(caminho)
                for _, row in df.iterrows():
                    codigo = normalizar_codigo(row.get("codigo_receita", ""))
                    valor_raw = row.get("valor_total")
                    cnpj = limpar_cnpj(row.get("cnpj", ""))
                    if not codigo or valor_raw is None:
                        continue
                    try:
                        valor = float(str(valor_raw).replace(".", "").replace(",", "."))
                        lista.append({"cnpj": cnpj, "codigo": codigo, "valor": valor})
                    except ValueError as e:
                        logger.warning(f"PERDCOMP Excel — valor inválido em {arq}: {e}")
            except Exception as e:
                logger.error(f"PERDCOMP Excel — erro ao ler {arq}: {e}")

        # PDF
        elif arq.lower().endswith(".pdf"):
            try:
                import pdfplumber
                with pdfplumber.open(caminho) as pdf:
                    texto = "".join(page.extract_text() or "" for page in pdf.pages)

                codigos = re.findall(r"(\d{4})-\d{2}", texto)
                valores = re.findall(r"Total\s+([\d\.,]+)", texto, re.IGNORECASE)
                cnpj_match = re.search(r"CNPJ\s+([\d.\-/]+)", texto)
                cnpj_pdf = limpar_cnpj(cnpj_match.group(1)) if cnpj_match else None

                for i, valor_str in enumerate(valores):
                    try:
                        valor = float(valor_str.replace(".", "").replace(",", "."))
                        codigo = codigos[i] if i < len(codigos) else None
                        if codigo and valor:
                            lista.append({"cnpj": cnpj_pdf, "codigo": codigo, "valor": valor})
                    except (ValueError, IndexError) as e:
                        logger.warning(f"PERDCOMP PDF — linha {i} em {arq}: {e}")
            except Exception as e:
                logger.error(f"PERDCOMP PDF — erro ao ler {arq}: {e}")

    return lista

# --------------------------
# CONSOLIDAÇÃO DE DÉBITOS
# --------------------------
def consolidar_debitos(cnpj: str, nome_empresa: str, sidaf: list, perdcomp: list) -> list:
    consolidado: dict[str, float] = defaultdict(float)

    for d in sidaf:
        if normalizar(nome_empresa) in normalizar(d["empresa"]):
            consolidado[d["codigo"]] += d["valor"]

    for d in perdcomp:
        if d["cnpj"] == cnpj:
            consolidado[d["codigo"]] += d["valor"]

    return [{"codigo": k, "valor": v} for k, v in consolidado.items()]

# --------------------------
# MONTAGEM DO JSON
# --------------------------
def montar_debitos_json(debitos_auto: list) -> dict:
    agrupados: dict[str, list] = defaultdict(list)

    for d in debitos_auto:
        for grupo, codigos in MAPEAMENTO_CODIGOS.items():
            if d["codigo"] in codigos:
                agrupados[grupo].append(d)
                break

    contador = 1
    estrutura = {"BalancoLucroReal": False}

    for grupo in ORDEM_JSON:
        if grupo not in agrupados:
            continue

        lista_debitos = []
        for d in sorted(agrupados[grupo], key=lambda x: x["codigo"]):
            cod_base = str(d["codigo"]).zfill(4)
            # Regra especial IRRF
            sufixo = "07" if cod_base == "0561" else "01"
            lista_debitos.append({
                "IdDebito":     contador,
                "CodigoDebito": cod_base + sufixo,
                "ValorDebito":  round(d["valor"], 2),
            })
            contador += 1

        nome_json = MAPA_NOME_JSON.get(grupo)
        if nome_json:
            estrutura[nome_json] = {"ListaDebitos": lista_debitos}

    return estrutura


# --------------------------
# GERAÇÃO DE TODOS OS JSONs
# --------------------------
def gerar_zip_todos(empresas, sidaf, perdcomp, responsavel):
    mes_ = datetime.now().month
    ano_ = datetime.now().year

    _codigos_validos = {cod for codigos in MAPEAMENTO_CODIGOS.values() for cod in codigos}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for emp in empresas:
            deb_bruto = consolidar_debitos(emp["cnpj"], emp["empresa"], sidaf, perdcomp)
            deb_validos = [d for d in deb_bruto if d["codigo"] in _codigos_validos]
            emp_sem_mov = len(deb_validos) == 0 or all(d["valor"] == 0 for d in deb_validos)

            if emp_sem_mov:
                json_emp = {
                    "PeriodoApuracao": {"MesApuracao": mes_, "AnoApuracao": ano_},
                    "DadosIniciais": {
                        "SemMovimento": True,
                        "QualificacaoPj": 11,
                        "ResponsavelApuracao": {"CpfResponsavel": responsavel["cpf"]},
                    },
                }
            else:
                json_emp = {
                    "PeriodoApuracao": {"MesApuracao": mes_, "AnoApuracao": ano_},
                    "DadosIniciais": {
                        "SemMovimento":        False,
                        "QualificacaoPj":      1,
                        "TributacaoLucro":     1,
                        "VariacoesMonetarias": 2,
                        "RegimePisCofins":     1,
                        "ResponsavelApuracao": {"CpfResponsavel": responsavel["cpf"]},
                    },
                    "Debitos": montar_debitos_json(deb_validos),
                }

            json_str = json.dumps(json_emp, indent=2, ensure_ascii=False)
            nome_arquivo = f"{emp['cnpj'][:8]}-MIT-{ano_}{str(mes_).zfill(2)}.json"

            # Nome da pasta = nome da empresa (caracteres inválidos removidos)
            nome_pasta = re.sub(r'[\\/*?:"<>|]', "_", emp["empresa"]).strip()
            caminho_no_zip = f"{nome_pasta}/{nome_arquivo}"

            zf.writestr(caminho_no_zip, json_str)

    buf.seek(0)
    return buf.read()

# --------------------------
# CARREGA DADOS
# --------------------------
empresas    = carregar_empresas(CAMINHO)
responsavel = carregar_responsavel(CAMINHO)
sidaf       = carregar_sidaf(CAMINHO)
perdcomp    = carregar_perdcomp(CAMINHO_PERDCOMP)

if not empresas:
    st.error("Nenhuma empresa encontrada na aba 'dados'.")
    st.stop()

# --------------------------
# EMPRESA ATUAL
# --------------------------
idx    = st.session_state.indice_empresa
dados  = empresas[idx]

debitos_auto = consolidar_debitos(dados["cnpj"], dados["empresa"], sidaf, perdcomp)

# Todos os códigos válidos do mapeamento
_todos_codigos_validos = {cod for codigos in MAPEAMENTO_CODIGOS.values() for cod in codigos}

# Débitos que realmente têm código dentro do mapeamento
debits = [d for d in debitos_auto if d["codigo"] in _todos_codigos_validos]

# Sem movimento = nenhum débito com código reconhecido e com valor > 0
sem_mov = len(debits) == 0 or all(d["valor"] == 0 for d in debits)

mes = datetime.now().month
ano = datetime.now().year
periodo = datetime.now().strftime("%m/%Y")

# =========================================================
# LAYOUT — CABEÇALHO
# =========================================================
pct = int(((idx + 1) / len(empresas)) * 100)

st.markdown(f"""
<div class="mit-header">
    <div>
        <div class="mit-title">MIT · Apuração</div>
        <div class="mit-subtitle">Malha de Informações Tributárias · {ano}</div>
    </div>
    <div style="text-align:right">
        <div style="font-family:'IBM Plex Mono',monospace;font-size:1.2rem;color:#111827;font-weight:600;">
            {idx + 1} <span style="color:#6b7280;font-size:0.9rem">/ {len(empresas)}</span>
        </div>
        <div style="font-size:0.72rem;color:#6b7280;margin-top:2px;">empresas</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Barra de progresso
st.markdown(f"""
<div class="progress-bar-container">
    <div class="progress-bar-fill" style="width:{pct}%"></div>
</div>
""", unsafe_allow_html=True)

# ─ Card empresa ─
badge_class = "badge-sem-mov" if sem_mov else "badge-com-mov"
badge_label = "Sem Movimento" if sem_mov else "Com Movimento"

st.markdown(f"""
<div class="empresa-card">
    <div class="empresa-nome">{dados['empresa']}</div>
    <div class="empresa-cnpj">{dados['cnpj']}</div>
    <span class="badge {badge_class}">{badge_label}</span>
</div>
""", unsafe_allow_html=True)

# ─ Navegação ─
col_ant, col_info, col_prox = st.columns([1, 2, 1])

with col_ant:
    if st.button("← Anterior", use_container_width=True):
        st.session_state.indice_empresa = max(0, idx - 1)
        st.rerun()

with col_info:
    st.markdown(
        f'<div class="nav-counter">{idx + 1} de {len(empresas)} · {pct}% concluído</div>',
        unsafe_allow_html=True,
    )

with col_prox:
    if st.button("Próxima →", use_container_width=True):
        st.session_state.indice_empresa = min(len(empresas) - 1, idx + 1)
        st.rerun()

# ─ Botão gerar todos ─
st.markdown("<br>", unsafe_allow_html=True)

col_ger1, col_ger2, col_ger3 = st.columns([1, 2, 1])
with col_ger2:
    if st.button("⚡ Gerar Todos os Arquivos", use_container_width=True, type="primary"):
        with st.spinner("Gerando JSONs e compactando..."):
            st.session_state["zip_bytes"] = gerar_zip_todos(empresas, sidaf, perdcomp, responsavel)
            st.session_state["zip_nome"]  = f"MIT-{datetime.now().strftime('%Y%m')}-todos.zip"

    if "zip_bytes" in st.session_state:
        st.download_button(
            label="📦 Baixar ZIP com todos os JSONs",
            data=st.session_state["zip_bytes"],
            file_name=st.session_state["zip_nome"],
            mime="application/zip",
            use_container_width=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# TABS
# =========================================================
tabs = st.tabs(["DADOS INICIAIS", "DÉBITOS", "ENCERRAMENTO", "EXPORTAR JSON"])

# ─────────────────────────────
# TAB 1 — DADOS INICIAIS
# ─────────────────────────────
with tabs[0]:

    st.markdown('<div class="section-title">Empresa</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    c1.text_input("Qualificação",  value=dados["qualificacao"] or "—", disabled=True)
    c1.text_input("Critério",      value=dados["criterio"]     or "—", disabled=True)
    c2.text_input("Tributação",    value=dados["tributacao"]   or "—", disabled=True)
    c2.text_input("Regime",        value=dados["regime"]       or "—", disabled=True)

    st.toggle("Apuração Sem Movimento", value=sem_mov, disabled=True)

    st.markdown('<div class="section-title">Responsável</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    c1.text_input("CPF",      value=responsavel["cpf"],      disabled=True)
    c1.text_input("DDD",      value=responsavel["ddd"],      disabled=True)
    c1.text_input("UF CRC",   value=responsavel["uf_crc"],   disabled=True)
    c2.text_input("Nome",     value=responsavel["nome"],     disabled=True)
    c2.text_input("Telefone", value=responsavel["telefone"], disabled=True)
    c2.text_input("CRC",      value=responsavel["crc"],      disabled=True)

    st.text_input("E-mail", value=responsavel["email"], disabled=True)

# ─────────────────────────────
# TAB 2 — DÉBITOS
# ─────────────────────────────
with tabs[1]:

    aba_sel = st.radio(
        "Grupo tributário",
        options=list(MAPEAMENTO_CODIGOS.keys()),
        horizontal=True,
        label_visibility="collapsed",
    )

    lista_grupo = [
        d for d in debits
        if d["codigo"] in MAPEAMENTO_CODIGOS.get(aba_sel, [])
    ]

    st.markdown("<br>", unsafe_allow_html=True)

    # Cabeçalho tabela
    st.markdown("""
    <div class="debito-header">
        <div>✔</div>
        <div>CÓDIGO</div>
        <div>PERÍODO</div>
        <div>VALOR DÉBITO</div>
        <div>SALDO A PAGAR</div>
    </div>
    """, unsafe_allow_html=True)

    if not lista_grupo:
        st.markdown(
            '<div class="empty-state">— Nenhum débito para este grupo —</div>',
            unsafe_allow_html=True,
        )
    else:
        # Checkbox + dados por linha usando colunas do Streamlit
        for i, d in enumerate(lista_grupo):
            codigo_fmt  = f"{d['codigo']}-01"
            valor_fmt   = formatar_valor(d["valor"])
            chk_key     = f"chk_{dados['cnpj']}_{aba_sel}_{d['codigo']}_{i}"

            cols = st.columns([0.4, 2, 2, 2, 2])
            cols[0].checkbox("", key=chk_key, label_visibility="collapsed")
            cols[1].markdown(f'<span class="codigo-chip">{codigo_fmt}</span>', unsafe_allow_html=True)
            cols[2].markdown(f'<span class="periodo-text">{periodo}</span>', unsafe_allow_html=True)
            cols[3].markdown(f'<span class="valor-text">{valor_fmt}</span>', unsafe_allow_html=True)
            cols[4].markdown(f'<span class="valor-text">{valor_fmt}</span>', unsafe_allow_html=True)

        # Total do grupo
        total_grupo = sum(d["valor"] for d in lista_grupo)
        st.markdown("<hr>", unsafe_allow_html=True)
        c_tot1, c_tot2 = st.columns([3, 2])
        c_tot1.markdown(
            '<span style="font-family:\'IBM Plex Mono\',monospace;font-size:0.75rem;'
            'text-transform:uppercase;letter-spacing:1px;color:#6b7280;">Total do grupo</span>',
            unsafe_allow_html=True,
        )
        c_tot2.markdown(
            f'<span style="font-family:\'IBM Plex Mono\',monospace;font-size:1rem;'
            f'font-weight:600;color:#111827;">{formatar_valor(total_grupo)}</span>',
            unsafe_allow_html=True,
        )

# ─────────────────────────────
# TAB 3 — ENCERRAMENTO
# ─────────────────────────────
with tabs[2]:

    st.markdown('<div class="section-title">Resumo dos Débitos Apurados</div>', unsafe_allow_html=True)

    # Totais por grupo
    grupos_resumo = [
        ("IRPJ",         "IRPJ"),
        ("CSLL",         "CSLL"),
        ("IRRF",         "IRRF"),
        ("IPI",          "IPI"),
        ("IOF",          "IOF"),
        ("PIS",          "PIS/PASEP"),
        ("COFINS",       "COFINS"),
        ("CONTRIB",      "Contribuições Diversas"),
        ("CPSS",         "CPSS"),
        ("RET/UNIF",     "RET/Pagamento Unificado"),
    ]

    def total_grupo_enc(grupo_key):
        return sum(
            d["valor"] for d in debits
            if d["codigo"] in MAPEAMENTO_CODIGOS.get(grupo_key, [])
        )

    totais = {label: total_grupo_enc(key) for key, label in grupos_resumo}
    total_apurado = sum(totais.values())

    # Divide em 2 colunas de 5 itens
    metade = len(grupos_resumo) // 2
    esquerda = grupos_resumo[:metade]
    direita   = grupos_resumo[metade:]

    st.markdown("""
    <style>
    .enc-card {
        background: #f8faff;
        border: 1px solid #dde3f0;
        border-radius: 12px;
        padding: 24px 28px;
    }
    .enc-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px 0;
        border-bottom: 1px solid #e2e8f0;
    }
    .enc-row:last-child { border-bottom: none; }
    .enc-label {
        font-size: 0.85rem;
        color: #374151;
        font-family: 'IBM Plex Sans', sans-serif;
        font-weight: 500;
    }
    .enc-valor {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.9rem;
        color: #111827;
        font-weight: 600;
    }
    .enc-valor-zero {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.9rem;
        color: #9ca3af;
    }
    .enc-total-card {
        background: #1a56db;
        border-radius: 10px;
        padding: 20px 24px;
        text-align: center;
        margin-bottom: 12px;
    }
    .enc-total-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #bfdbfe;
        font-family: 'IBM Plex Sans', sans-serif;
        margin-bottom: 6px;
    }
    .enc-total-valor {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.3rem;
        font-weight: 700;
        color: #ffffff;
    }
    .enc-suspensao-card {
        background: #f59e0b;
        border-radius: 10px;
        padding: 20px 24px;
        text-align: center;
    }
    .enc-suspensao-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #78350f;
        font-family: 'IBM Plex Sans', sans-serif;
        margin-bottom: 6px;
    }
    .enc-suspensao-valor {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.3rem;
        font-weight: 700;
        color: #ffffff;
    }
    </style>
    """, unsafe_allow_html=True)

    col_esq, col_dir, col_tot = st.columns([3, 3, 2])

    def render_grupo(col, grupos):
        rows_html = ""
        for key, label in grupos:
            v = totais[label]
            cls = "enc-valor" if v > 0 else "enc-valor-zero"
            rows_html += f"""
            <div class="enc-row">
                <span class="enc-label">{label}</span>
                <span class="{cls}">{formatar_valor(v)}</span>
            </div>"""
        col.markdown(f'<div class="enc-card">{rows_html}</div>', unsafe_allow_html=True)

    render_grupo(col_esq, esquerda)
    render_grupo(col_dir, direita)

    with col_tot:
        col_tot.markdown(f"""
        <div class="enc-total-card">
            <div class="enc-total-label">Total Apurado</div>
            <div class="enc-total-valor">{formatar_valor(total_apurado)}</div>
        </div>
        <div class="enc-suspensao-card">
            <div class="enc-suspensao-label">Suspensões</div>
            <div class="enc-suspensao-valor">{formatar_valor(0.0)}</div>
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────
# TAB 4 — EXPORTAR JSON
# ─────────────────────────────
with tabs[3]:

    st.markdown('<div class="section-title">JSON de Exportação</div>', unsafe_allow_html=True)

    # Monta JSON
    if sem_mov:
        json_final = {
            "PeriodoApuracao": {"MesApuracao": mes, "AnoApuracao": ano},
            "DadosIniciais": {
                "SemMovimento": True,
                "QualificacaoPj": 11,
                "ResponsavelApuracao": {"CpfResponsavel": responsavel["cpf"]},
            },
        }
    else:
        json_final = {
            "PeriodoApuracao": {"MesApuracao": mes, "AnoApuracao": ano},
            "DadosIniciais": {
                "SemMovimento":       False,
                "QualificacaoPj":     1,
                "TributacaoLucro":    1,
                "VariacoesMonetarias": 2,
                "RegimePisCofins":    1,
                "ResponsavelApuracao": {"CpfResponsavel": responsavel["cpf"]},
            },
            "Debitos": montar_debitos_json(debits),
        }

    json_str = json.dumps(json_final, indent=2, ensure_ascii=False)
    nome_arquivo = f"{dados['cnpj'][:8]}-MIT-{ano}{str(mes).zfill(2)}.json"

    # Preview
    st.code(json_str, language="json")

    # Download
    st.download_button(
        label="📥  Baixar JSON",
        data=json_str,
        file_name=nome_arquivo,
        mime="application/json",
        use_container_width=False,
        key=f"dl_json_{dados['cnpj']}",
    )

    # Resumo de débitos consolidados
    if debits:
        st.markdown('<div class="section-title">Resumo de Débitos</div>', unsafe_allow_html=True)
        total_geral = sum(d["valor"] for d in debits)
        rows_html = "".join(
            f"<tr>"
            f"<td style='padding:8px 12px;font-family:IBM Plex Mono,monospace;font-size:0.85rem;"
            f"color:#1a56db;border-bottom:1px solid #e2e8f0;'>{d['codigo']}</td>"
            f"<td style='padding:8px 12px;font-family:IBM Plex Mono,monospace;font-size:0.85rem;"
            f"color:#166534;border-bottom:1px solid #e2e8f0;'>{formatar_valor(d['valor'])}</td>"
            f"</tr>"
            for d in debits
        )
        st.markdown(f"""
        <table style='width:100%;border-collapse:collapse;background:#f8faff;
                      border-radius:8px;overflow:hidden;border:1px solid #e2e8f0;'>
            <thead>
                <tr style='background:#f1f5f9;'>
                    <th style='padding:8px 12px;text-align:left;font-size:0.72rem;
                               text-transform:uppercase;letter-spacing:1px;color:#6b7280;
                               font-family:IBM Plex Sans,sans-serif;border-bottom:1px solid #e2e8f0;'>Código</th>
                    <th style='padding:8px 12px;text-align:left;font-size:0.72rem;
                               text-transform:uppercase;letter-spacing:1px;color:#6b7280;
                               font-family:IBM Plex Sans,sans-serif;border-bottom:1px solid #e2e8f0;'>Valor (R$)</th>
                </tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
        """, unsafe_allow_html=True)
        st.markdown(
            f'<div style="text-align:right;font-family:\'IBM Plex Mono\',monospace;'
            f'font-size:0.85rem;color:#6b7280;margin-top:8px;">'
            f'Total geral: <strong style="color:#111827;">{formatar_valor(total_geral)}</strong></div>',
            unsafe_allow_html=True,
        )
