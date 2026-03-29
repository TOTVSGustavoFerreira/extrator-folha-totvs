import streamlit as st

st.set_page_config(
    page_title="Portal TOTVS - Extratores",
    page_icon="🏢",
)

st.title("🏢 Portal de Extratores de Folha")
st.write("Bem-vindo ao portal de processamento e conversão de folhas de pagamento!")
st.write("👈 **Selecione a ferramenta desejada no menu lateral à esquerda:**")

st.info("📊 **1. Extrator Excel:** Extrai dados de folhas em formato .xlsx e .xls.")
st.info("📄 **2. Extrator PDF:** Extrai dados APENAS FOLHAS nativas geradas (IOB) em .pdf.")

st.write("---")
st.caption("Desenvolvido para a equipe TOTVS.")
