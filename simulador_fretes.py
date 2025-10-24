import os
import re
import pandas as pd
import streamlit as st
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple  # compatível com 3.8/3.9
import io  # <-- para download em memória

# ==========================================================
# BOOTSTRAP PARA SUBMÓDULO PRIVADO (Streamlit Cloud)
# ==========================================================
# Lê a chave privada de st.secrets["SSH_PRIVATE_KEY"], cria ~/.ssh/id_planilhas
# e executa "git submodule update --init --recursive" para baixar dados_vtex.
import subprocess, pathlib

@st.cache_resource(show_spinner=False)  # cacheia por sessão
def _bootstrap_submodule():
    if "SSH_PRIVATE_KEY" not in st.secrets:
        st.warning("SSH_PRIVATE_KEY não configurada nos Secrets do Streamlit Cloud.")
        return False
    ssh_dir = pathlib.Path.home() / ".ssh"
    ssh_dir.mkdir(parents=True, exist_ok=True)
    key_path = ssh_dir / "id_planilhas"
    key_text = st.secrets["SSH_PRIVATE_KEY"]
    if not key_text.endswith("\n"):
        key_text += "\n"
    key_path.write_text(key_text)
    os.chmod(key_path, 0o600)
    # usa essa chave e ignora verificação de host (evita interactive prompt)
    os.environ["GIT_SSH_COMMAND"] = f"ssh -i {key_path} -o StrictHostKeyChecking=no"
    try:
        subprocess.run(["git", "submodule", "sync"], check=True)
        subprocess.run(["git", "submodule", "update", "--init", "--recursive"], check=True)
        return True
    except Exception as e:
        st.error(f"Falha ao atualizar submódulo privado: {e}")
        return False

_bootstrap_submodule()
# ==========================================================

# ==========================================================
# CONFIGURAÇÕES INICIAIS
# ==========================================================
st.set_page_config(page_title="Simulador de Fretes VTEX", layout="wide")

# Caminho da pasta das planilhas VTEX (AGORA LENDO O SUBMÓDULO)
PASTA_VTEX = "dados_vtex"

# Ocultar apenas na TELA (no Excel salvo continuará presente)
HIDE_COLS_ON_SCREEN: List[str] = ["Arquivo_Origem", "Aba_Origem"]

# ==========================================================
# FUNÇÕES AUXILIARES (NOVAS + AS SUAS)
# ==========================================================
def _normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().upper()
    for chars, rep in (("ÁÀÂÃ", "A"), ("ÉÈÊ", "E"), ("ÍÌÎ", "I"), ("ÓÒÔÕ", "O"), ("ÚÙÛ", "U"), ("Ç", "C")):
        s = re.sub(f"[{chars}]", rep, s)
    s = re.sub(r"\s+", " ", s).replace(" ", "_")
    return s

def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def _to_float_br(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace("\u00a0", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

def _find_col(df: pd.DataFrame, targets: List[str]) -> Optional[str]:
    colmap = {c: _normalize_text(c) for c in df.columns}
    inv = {v: k for k, v in colmap.items()}
    for t in targets:
        if t in inv:
            return inv[t]
    for t in targets:
        for norm, orig in inv.items():
            if norm.startswith(t):
                return orig
    return None

def buscar_uf_cep(cep: int) -> Tuple[str, float]:
    cep_ranges = {
        'SP': range(1_000_000, 20_000_000), 'RJ': range(20_000_000, 29_000_000),
        'ES': range(29_000_000, 30_000_000), 'MG': range(30_000_000, 40_000_000),
        'BA': range(40_000_000, 49_000_000), 'SE': range(49_000_000, 50_000_000),
        'PE': range(50_000_000, 57_000_000), 'AL': range(57_000_000, 58_000_000),
        'PB': range(58_000_000, 59_000_000), 'RN': range(59_000_000, 60_000_000),
        'CE': range(60_000_000, 64_000_000), 'PI': range(64_000_000, 65_000_000),
        'MA': range(65_000_000, 66_000_000), 'PA': range(66_000_000, 68_900_000),
        'AP': range(68_900_000, 69_000_000), 'AM': range(69_000_000, 69_300_000),
        'RR': range(69_300_000, 69_400_000), 'AM2': range(69_400_000, 69_900_000),
        'AC': range(69_900_000, 70_000_000), 'DF': range(70_000_000, 73_700_000),
        'GO': range(73_700_000, 76_800_000), 'RO': range(76_800_000, 77_000_000),
        'TO': range(77_000_000, 78_000_000), 'MT': range(78_000_000, 79_000_000),
        'MS': range(79_000_000, 80_000_000), 'PR': range(80_000_000, 88_000_000),
        'SC': range(88_000_000, 90_000_000), 'RS': range(90_000_000, 100_000_000),
    }
    imposto = {
        'AC': 7,'AL': 7,'AM': 7,'AP': 7,'BA': 7,'CE': 7,'DF': 7,'ES': 7,'GO': 7,'MA': 7,
        'MG': 12,'MS': 7,'MT': 7,'PA': 7,'PB': 7,'PE': 7,'PI': 7,'PR': 12,'RJ': 20,'RN': 7,
        'RO': 7,'RR': 7,'RS': 12,'SC': 12,'SE': 7,'SP': 12,'TO': 7
    }
    for uf, rng in cep_ranges.items():
        if cep in rng:
            key = uf if uf != 'AM2' else 'AM'
            return key, float(imposto.get(key, 0.0))
    return '', 0.0

def hash_arquivos(pasta: str) -> str:
    hash_md5 = hashlib.md5()
    try:
        arquivos = sorted(os.listdir(pasta))
    except Exception:
        return ""
    for arquivo in arquivos:
        if arquivo.startswith("~$"):
            continue
        if arquivo.endswith((".xlsx", ".xls")) and "VTEX" in arquivo.upper():
            try:
                caminho = os.path.join(pasta, arquivo)
                mtime = str(os.path.getmtime(caminho))
                hash_md5.update((arquivo + mtime).encode("utf-8"))
            except Exception:
                continue
    return hash_md5.hexdigest()

@st.cache_data(show_spinner=False)
def carregar_planilhas_vtex(pasta, hash_pasta):
    try:
        # aceitar somente .xlsx (evita engine .xls no Cloud)
        arquivos_excel = [
            f for f in os.listdir(pasta)
            if not f.startswith("~$") and f.lower().endswith(".xlsx") and "VTEX" in f.upper()
        ]
    except Exception as e:
        st.warning(f"⚠️ Não foi possível acessar a pasta: {e}")
        return pd.DataFrame()

    dfs = []
    total_arquivos = len(arquivos_excel)
    if total_arquivos == 0:
        return pd.DataFrame()

    barra_progresso = st.progress(0)
    status_texto = st.empty()

    for i, arquivo in enumerate(sorted(arquivos_excel), start=1):
        caminho_arquivo = os.path.join(pasta, arquivo)
        status_texto.text(f"📄 Lendo {i}/{total_arquivos}: {arquivo}")
        try:
            xls = pd.ExcelFile(caminho_arquivo, engine="openpyxl")  # .xlsx
            for aba in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=aba, engine="openpyxl")
                except Exception as e_aba:
                    st.warning(f"⚠️ Erro ao ler a aba '{aba}' do arquivo '{arquivo}': {e_aba}")
                    continue
                nome_transportadora = arquivo.split("-")[0].strip()
                df["Transportadora"] = nome_transportadora
                df["Arquivo_Origem"] = arquivo
                df["Aba_Origem"] = aba
                dfs.append(df)
        except PermissionError as eperm:
            st.warning(f"⚠️ Sem permissão para ler '{arquivo}'. Feche o arquivo e verifique permissões. Detalhe: {eperm}")
        except Exception as e:
            st.warning(f"⚠️ Erro ao ler '{arquivo}': {e}")
        barra_progresso.progress(i / total_arquivos)

    barra_progresso.empty()
    status_texto.text("✅ Consolidação concluída com sucesso.")
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)

def salvar_resultado(df_resultado: pd.DataFrame) -> str:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    caminho = os.path.join(desktop, f"resultado_fretes_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    df_resultado.to_excel(caminho, index=False)
    return caminho

# versão em memória para usar no Streamlit Cloud
def salvar_resultado_para_download(df_resultado: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False)
    buf.seek(0)
    return buf.getvalue()

# ===================== FORMATAÇÃO + DESTAQUE (FRETE_TOTAL) =====================
def destacar_min_max(df: pd.DataFrame) -> "pd.io.formats.style.Styler":
    """
    Destaca menor (verde) e maior (vermelho) FRETE_TOTAL e
    força a exibição com os formatos solicitados:
      - R$ com 2 casas: FRETE_PESO, FRETE_TOTAL
      - 3 casas sem R$: PESO_INICIAL, PESO_FINAL
      - % com 2 casas: GRIS_ADVALOREM, IMPOSTO
    """
    if df.empty:
        return df.style

    # --- formatadores ---
    def brl(x):
        try:
            if pd.isna(x): return ""
            s = f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"R$ {s}"
        except Exception:
            return x

    def three_dec(x):
        try:
            if pd.isna(x): return ""
            return f"{float(x):.3f}".replace(".", ",")
        except Exception:
            return x

    def pct(x):
        try:
            if pd.isna(x): return ""
            return f"{float(x):.2f}%".replace(".", ",")
        except Exception:
            return x

    # min/max antes de formatar
    try:
        min_idx = df["FRETE_TOTAL"].idxmin() if "FRETE_TOTAL" in df.columns else None
        max_idx = df["FRETE_TOTAL"].idxmax() if "FRETE_TOTAL" in df.columns else None
    except Exception:
        min_idx = max_idx = None

    # cópia para exibição (strings já formatadas)
    display_df = df.copy()

    brl_cols = [c for c in ["FRETE_PESO", "FRETE_TOTAL"] if c in display_df.columns]
    weight_cols = [c for c in ["PESO_INICIAL", "PESO_FINAL"] if c in display_df.columns]
    percent_cols = [c for c in ["GRIS_ADVALOREM", "IMPOSTO"] if c in display_df.columns]

    for c in brl_cols:
        display_df[c] = display_df[c].map(brl)
    for c in weight_cols:
        display_df[c] = display_df[c].map(three_dec)   # 3 casas decimais
    for c in percent_cols:
        display_df[c] = display_df[c].map(pct)

    def _row_style(row):
        if min_idx is not None and row.name == min_idx:
            return ["background-color: #d1fae5"] * len(row)  # verde
        if max_idx is not None and row.name == max_idx:
            return ["background-color: #fee2e2"] * len(row)  # vermelho
        return [""] * len(row)

    return display_df.style.apply(_row_style, axis=1)

# ============== Tabela SEM destaque (para o modo Upload) ==============
def show_results_plain(df_display: pd.DataFrame):
    """
    Tabela SEM destaque (sem verde/vermelho), mas com as mesmas
    formatações: R$, %, e 3 casas nos pesos. Também renomeia POLYGON -> Cidade/UF.
    """
    if df_display.empty:
        st.dataframe(df_display, use_container_width=True)
        return

    # Copia para exibição
    df_view = df_display.copy()
    if "POLYGON" in df_view.columns:
        df_view = df_view.rename(columns={"POLYGON": "Cidade/UF"})

    # --- formatadores iguais aos do destacar_min_max ---
    def brl(x):
        try:
            if pd.isna(x): return ""
            s = f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"R$ {s}"
        except Exception:
            return x

    def three_dec(x):
        try:
            if pd.isna(x): return ""
            return f"{float(x):.3f}".replace(".", ",")
        except Exception:
            return x

    def pct(x):
        try:
            if pd.isna(x): return ""
            return f"{float(x):.2f}%".replace(".", ",")
        except Exception:
            return x

    # Aplica formatação como texto
    if "FRETE_PESO" in df_view.columns:
        df_view["FRETE_PESO"] = df_view["FRETE_PESO"].map(brl)
    if "FRETE_TOTAL" in df_view.columns:
        df_view["FRETE_TOTAL"] = df_view["FRETE_TOTAL"].map(brl)
    for c in ["PESO_INICIAL", "PESO_FINAL"]:
        if c in df_view.columns:
            df_view[c] = df_view[c].map(three_dec)
    for c in ["GRIS_ADVALOREM", "IMPOSTO"]:
        if c in df_view.columns:
            df_view[c] = df_view[c].map(pct)

    st.dataframe(df_view, use_container_width=True)

# ===================== ÚNICO RESULTADO (com destaque) =====================
def show_results_single(df_display: pd.DataFrame):
    """
    Única tabela com:
      - destaque verde/vermelho (menor/maior FRETE_TOTAL)
      - formatação R$ / % / 3 casas nos pesos
      - cabeçalho 'Cidade/UF' no lugar de 'POLYGON'
    """
    if df_display.empty:
        st.dataframe(df_display, use_container_width=True)
        return

    df_view = df_display.copy()
    if "POLYGON" in df_view.columns:
        df_view = df_view.rename(columns={"POLYGON": "Cidade/UF"})

    # destacar_min_max retorna um Styler -> usar st.table para evitar crash no Cloud
    st.table(destacar_min_max(df_view))

# ==========================================================
# NORMALIZAÇÃO DA BASE (KG) – mantém PolygonName
# ==========================================================
@st.cache_data(show_spinner=False)
def normalizar_base_vtex(df_vtex: pd.DataFrame) -> pd.DataFrame:
    if df_vtex.empty:
        return df_vtex.copy()

    df = df_vtex.copy()

    c_cep_ini       = _find_col(df, ["ZIPCODESTART", "CEP_INICIAL", "CEP_INI", "FAIXA_CEP_INICIAL"])
    c_cep_fim       = _find_col(df, ["ZIPCODEEND",   "CEP_FINAL",   "CEP_FIM", "FAIXA_CEP_FINAL"])
    c_peso_ini      = _find_col(df, ["WEIGHTSTART",  "PESO_INICIAL","PESO_INI"])
    c_peso_fim      = _find_col(df, ["WEIGHTEND",    "PESO_FINAL",  "PESO_FIM"])
    c_frete_abs     = _find_col(df, ["ABSOLUTEMONEYCOST","FRETE_BASE","FRETE","VALOR_FRETE","VALOR"])
    c_price_percent = _find_col(df, ["PRICEPERCENT"])
    c_extra_per_kg  = _find_col(df, ["PRICEBYEXTRAWEIGHT"])
    c_prazo         = _find_col(df, ["TIMECOST","PRAZO","PRAZO_ENTREGA","LEAD_TIME"])
    c_polygon       = _find_col(df, ["POLYGONNAME","CIDADE","MUNICIPIO"])

    for c in [c_cep_ini, c_cep_fim, c_peso_ini, c_peso_fim]:
        if c and c in df.columns: df[c] = _to_numeric(df[c])
    if c_frete_abs and c_frete_abs in df.columns: df[c_frete_abs] = _to_float_br(df[c_frete_abs])
    if c_price_percent and c_price_percent in df.columns: df[c_price_percent] = _to_float_br(df[c_price_percent])
    if c_extra_per_kg and c_extra_per_kg in df.columns: df[c_extra_per_kg] = _to_float_br(df[c_extra_per_kg])

    df = df[_to_numeric(df[c_cep_ini]).notna() & _to_numeric(df[c_cep_fim]).notna()].copy()

    mask_gramas = (df[c_peso_fim].between(50, 5000, inclusive="both")) & (df[c_peso_fim] >= df[c_peso_ini])
    df["KG_INI"] = df[c_peso_ini].where(~mask_gramas, df[c_peso_ini] / 1000.0)
    df["KG_FIM"] = df[c_peso_fim].where(~mask_gramas, df[c_peso_fim] / 1000.0)

    cols_keep = {
        "Transportadora": "Transportadora",
        c_cep_ini: "CEP_INICIAL",
        c_cep_fim: "CEP_FINAL",
        "KG_INI": "KG_INI",
        "KG_FIM": "KG_FIM",
        c_frete_abs if c_frete_abs else None: "ABS_COST",
        c_price_percent if c_price_percent else None: "PRICE_PERCENT",
        c_extra_per_kg if c_extra_per_kg else None: "EXTRA_PER_KG",
        c_prazo if c_prazo else None: "PRAZO_ENTREGA",
        c_polygon if c_polygon else None: "POLYGON",
        "Arquivo_Origem": "Arquivo_Origem",
        "Aba_Origem": "Aba_Origem",
    }
    cols_keep = {k: v for k, v in cols_keep.items() if k in df.columns}
    base = df[list(cols_keep.keys())].rename(columns=cols_keep)
    base["ABS_COST"] = _to_numeric(base.get("ABS_COST", 0)).fillna(0)
    base["PRICE_PERCENT"] = _to_numeric(base.get("PRICE_PERCENT", 0.5)).fillna(0.5)
    base["EXTRA_PER_KG"] = _to_numeric(base.get("EXTRA_PER_KG", 0)).fillna(0)
    return base.reset_index(drop=True)

def calcular_frete_vetor(base_norm: pd.DataFrame, cep: int, peso_kg: float,
                         transp_sel: Optional[List[str]]) -> pd.DataFrame:
    if base_norm.empty:
        return base_norm

    df = base_norm
    if transp_sel and "Todas" not in transp_sel and "Transportadora" in df.columns:
        df = df[df["Transportadora"].isin(transp_sel)]
        if df.empty:
            return df

    m = (df["CEP_INICIAL"] <= cep) & (df["CEP_FINAL"] >= cep) & (df["KG_INI"] <= peso_kg) & (df["KG_FIM"] >= peso_kg)
    return df.loc[m].copy()

# ==========================================================
# CÁLCULO DE FRETE (SEU ORIGINAL PARA UNITÁRIO)
# ==========================================================
def calcular_frete(df_vtex: pd.DataFrame, cep_destino: str, valor_nf: float, peso: float,
                   transportadora: Optional[str] = None) -> pd.DataFrame:
    if df_vtex.empty:
        return pd.DataFrame()

    cep_somente_digitos = re.sub(r"\D", "", str(cep_destino)).strip()
    if not cep_somente_digitos.isdigit() or len(cep_somente_digitos) < 8:
        return pd.DataFrame()
    cep_num = int(cep_somente_digitos)

    base = normalizar_base_vtex(df_vtex)

    transp_sel = None if (not transportadora or transportadora == "Todas") else [transportadora]
    df_filtrado = calcular_frete_vetor(base, cep_num, float(peso), transp_sel)
    if df_filtrado.empty:
        return pd.DataFrame()

    excesso = (float(peso) - df_filtrado["KG_FIM"]).clip(lower=0)
    frete_peso = df_filtrado["ABS_COST"] + excesso * df_filtrado["EXTRA_PER_KG"]

    # Percentuais
    gris_percent = df_filtrado["PRICE_PERCENT"].fillna(0.5)  # %
    uf, perc_imp = buscar_uf_cep(cep_num)                    # %

    # Cálculo financeiro
    gris_valor = (gris_percent / 100.0) * float(valor_nf)
    sub_total = frete_peso + gris_valor
    fator = 1.0 - (perc_imp / 100.0)
    if fator <= 0: fator = 1.0
    valor_total = sub_total / fator

    saida = df_filtrado.copy()
    saida["PESO_INICIAL"] = saida["KG_INI"]
    saida["PESO_FINAL"] = saida["KG_FIM"]
    saida["FRETE_PESO"] = frete_peso
    saida["GRIS_ADVALOREM"] = gris_percent          # (%)
    saida["IMPOSTO"] = perc_imp                     # (%)
    saida["FRETE_TOTAL"] = valor_total

    cols = [c for c in [
        "Transportadora", "CEP_INICIAL", "CEP_FINAL", "PESO_INICIAL", "PESO_FINAL",
        "FRETE_PESO", "GRIS_ADVALOREM", "IMPOSTO", "FRETE_TOTAL", "PRAZO_ENTREGA",
        "POLYGON", "Arquivo_Origem", "Aba_Origem"
    ] if c in saida.columns]
    saida = saida[cols].sort_values("FRETE_TOTAL", ascending=True).reset_index(drop=True)
    return saida

# ==========================================================
# INTERFACE STREAMLIT (SEU FLUXO ORIGINAL + MULTISELECT + UF/POLYGON)
# ==========================================================
st.title("🚚 Simulador de Fretes VTEX")

hash_atual = hash_arquivos(PASTA_VTEX)
if not hash_atual:
    st.error("❌ Pasta inválida ou inacessível. Verifique o caminho configurado.")
    st.stop()

df_vtex = carregar_planilhas_vtex(PASTA_VTEX, hash_atual)
if df_vtex.empty:
    st.error("❌ Nenhuma planilha VTEX encontrada na pasta configurada.")
    st.stop()

base_norm = normalizar_base_vtex(df_vtex)

modo = st.radio("Selecione o modo de simulação:", ["Consulta unitária", "Upload em Excel"])

# ==========================================================
# CONSULTA UNITÁRIA
# ==========================================================
if modo == "Consulta unitária":
    st.subheader("Simulação Unitária")

    col1, col2, col3 = st.columns(3)
    with col1:
        cep_destino = st.text_input("CEP Destino", "")
    with col2:
        valor_nf = st.number_input("Valor da Nota Fiscal (R$)", min_value=0.0, step=10.0)
    with col3:
        peso = st.number_input("Peso (kg)", min_value=0.0, step=0.1)

    if "Transportadora" in df_vtex.columns:
        transportadoras_lista = sorted(df_vtex["Transportadora"].dropna().unique().tolist())
    else:
        transportadoras_lista = []
    opcoes_transp = ["Todas"] + transportadoras_lista
    transp_selecionadas = st.multiselect("Transportadoras", options=opcoes_transp, default=["Todas"])

    if st.button("Calcular Frete"):
        if cep_destino and valor_nf > 0 and peso > 0:
            resultado = calcular_frete(df_vtex, cep_destino, valor_nf, peso, transportadora=None)

            # Mostrar UF + cidade (Polygon)
            cep_dig = re.sub(r"\D", "", str(cep_destino))
            if cep_dig.isdigit():
                uf, _ = buscar_uf_cep(int(cep_dig))
            else:
                uf = ""
            if resultado is not None and not resultado.empty and "POLYGON" in resultado.columns:
                try:
                    cidade = resultado["POLYGON"].mode().iloc[0]
                except Exception:
                    cidade = resultado["POLYGON"].iloc[0]
            else:
                cidade = "N/D"
            st.info(f"**Destino:** {cidade} — **UF:** {uf}")

            if resultado is not None and not resultado.empty:
                if transp_selecionadas and "Todas" not in transp_selecionadas and "Transportadora" in resultado.columns:
                    resultado = resultado[resultado["Transportadora"].isin(transp_selecionadas)]

            if resultado is not None and not resultado.empty:
                resultado_display = resultado.drop(columns=HIDE_COLS_ON_SCREEN, errors="ignore")
                st.success(f"✅ {len(resultado_display)} opções encontradas.")
                # Único grid (formatação + destaque) e POLYGON -> Cidade/UF
                show_results_single(resultado_display)
            else:
                st.warning("Nenhum resultado encontrado para os filtros informados.")
        else:
            st.error("Por favor, preencha todos os campos para simular.")

# ==========================================================
# UPLOAD DE EXCEL (linha a linha, com POLYGON e UF no resultado)
# ==========================================================
else:
    st.subheader("📤 Upload de Arquivo Excel")
    st.info("**Principais colunas (nomes exatos):** `ORIGEM`, `CEP DESTINO`, `VALOR DE NFE`, `PESO`.")

    if "Transportadora" in df_vtex.columns:
        transportadoras_lista = sorted(df_vtex["Transportadora"].dropna().unique().tolist())
    else:
        transportadoras_lista = []
    opcoes_transp = ["Todas"] + transportadoras_lista
    transp_selecionadas = st.multiselect("Transportadoras (aplicado ao lote)", options=opcoes_transp, default=["Todas"])

    arquivo = st.file_uploader("Selecione um arquivo (.xlsx ou .xls)", type=["xlsx", "xls"])

    if arquivo:
        try:
            df_upload = pd.read_excel(arquivo)
            colunas_obrigatorias = ["ORIGEM", "CEP DESTINO", "VALOR DE NFE", "PESO"]
            faltantes = [c for c in colunas_obrigatorias if c not in df_upload.columns]

            if faltantes:
                st.error(f"❌ Arquivo inválido. Faltam as colunas obrigatórias: {', '.join(faltantes)}")
            else:
                st.success("✅ Arquivo válido. Processando simulações...")
                resultados = []

                for _, linha in df_upload.iterrows():
                    cep_txt = str(linha["CEP DESTINO"])
                    cep_dig = re.sub(r"\D", "", cep_txt)
                    if not cep_dig.isdigit():
                        continue
                    cep = int(cep_dig)

                    peso_linha = float(linha["PESO"]) if pd.notna(linha["PESO"]) else None
                    valor_nf_linha = float(linha["VALOR DE NFE"]) if pd.notna(linha["VALOR DE NFE"]) else None
                    if peso_linha is None or valor_nf_linha is None:
                        continue

                    df_matches = calcular_frete_vetor(base_norm, cep, peso_linha, transp_selecionadas)
                    if df_matches.empty:
                        continue

                    excesso = (peso_linha - df_matches["KG_FIM"]).clip(lower=0)
                    frete_peso = df_matches["ABS_COST"] + excesso * df_matches["EXTRA_PER_KG"]

                    gris_percent = df_matches["PRICE_PERCENT"].fillna(0.5)
                    uf, perc_imp = buscar_uf_cep(cep)

                    gris_valor = (gris_percent / 100.0) * valor_nf_linha
                    sub_total = frete_peso + gris_valor
                    fator = 1.0 - (perc_imp / 100.0)
                    if fator <= 0: fator = 1.0
                    valor_total = sub_total / fator

                    parcial = df_matches.copy()
                    parcial["PESO_INICIAL"] = parcial["KG_INI"]
                    parcial["PESO_FINAL"] = parcial["KG_FIM"]
                    parcial["FRETE_PESO"] = frete_peso
                    parcial["GRIS_ADVALOREM"] = gris_percent    # (%)
                    parcial["IMPOSTO"] = perc_imp               # (%)
                    parcial["FRETE_TOTAL"] = valor_total
                    parcial["UF_DESTINO"] = uf
                    parcial["ORIGEM"] = linha["ORIGEM"]
                    parcial["CEP_DESTINO_ORIGINAL"] = cep
                    parcial["VALOR_NFE_ORIGINAL"] = valor_nf_linha
                    parcial["PESO_ORIGINAL"] = peso_linha

                    resultados.append(parcial[[
                        "Transportadora", "CEP_INICIAL", "CEP_FINAL", "PESO_INICIAL", "PESO_FINAL",
                        "FRETE_PESO", "GRIS_ADVALOREM", "IMPOSTO", "FRETE_TOTAL", "PRAZO_ENTREGA",
                        "POLYGON", "UF_DESTINO",
                        "Arquivo_Origem", "Aba_Origem", "ORIGEM", "CEP_DESTINO_ORIGINAL",
                        "VALOR_NFE_ORIGINAL", "PESO_ORIGINAL"
                    ]])

                if resultados:
                    df_final = pd.concat(resultados, ignore_index=True)

                    # mantém sua gravação local (para uso desktop)
                    caminho = salvar_resultado(df_final)
                    st.success(f"✅ Simulações concluídas. Resultado salvo em:\n{caminho}")

                    # botão para baixar no Cloud
                    payload = salvar_resultado_para_download(df_final)
                    st.download_button(
                        "⬇️ Baixar resultado (Excel)",
                        data=payload,
                        file_name=f"resultado_fretes_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    df_final_display = df_final.drop(columns=HIDE_COLS_ON_SCREEN, errors="ignore")
                    # Upload: sem destaque (apenas formatação) e POLYGON -> Cidade/UF
                    show_results_plain(df_final_display.head(50))
                else:
                    st.warning("Nenhum resultado encontrado para as linhas enviadas.")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
