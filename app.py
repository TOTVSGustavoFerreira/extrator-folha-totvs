import streamlit as st

st.set_page_config(
    page_title="Portal TOTVS - Extratores",
    page_icon="🏢",
)

st.title("🏢 Portal de Extratores de Folha")
st.write("Bem-vindo ao portal de processamento e conversão de folhas de pagamento!")
st.write("👈 **Selecione a ferramenta desejada no menu lateral à esquerda:**")

st.info("📊 **1. Extrator Excel:** Extrai dados de FOLHAS (FORTES) em formato .xlsx e .xls.")
st.info("📄 **2. Extrator PDF:** Extrai dados APENAS FOLHAS nativas geradas (IOB) em .pdf.")
st.info("🧮 **3. Extrator Bases PDF:** Extrai as BASES DE CÁLCULO de folhas geradas (IOB) em .pdf.")
st.info("🧮 **4. Extrator Bases EXCEL:** Extrai as BASES DE CÁLCULO de folhas geradas (FORTES) em .xlsx e .xls.")

st.write("---")
st.caption("Desenvolvido para a equipe TOTVS.")
