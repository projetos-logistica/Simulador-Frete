# boot.py — wrapper para exibir qualquer erro de import/execução no Streamlit Cloud
import streamlit as st

st.set_page_config(page_title="Simulador de Fretes VTEX", layout="wide")

try:
    # Importa seu app “de verdade”. NÃO altere o nome nem a estrutura do seu código.
    import simulador_fretes
    st.caption("BOOT: app importado com sucesso.")
except Exception as e:
    st.error("❌ Falha ao carregar o aplicativo.")
    st.exception(e)  # Mostra o stacktrace completo na página
    st.stop()
