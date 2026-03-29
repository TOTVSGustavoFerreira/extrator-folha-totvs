import streamlit as st
import pdfplumber
import pandas as pd
import re
import zipfile
import io

# =========================
# FUNÇÕES AUXILIARES
# =========================
def normalizar_valor(v):
    if v is None or str(v).strip() == "":
        return None

    if isinstance(v, float):
        return v

    v = str(v).strip()
    
    if "," in v:
        v = v.replace(".", "")
        v = v.replace(",", ".")
    else:
        if v.count(".") > 1:
            partes = v.rsplit(".", 1)
            v = partes[0].replace(".", "") + "." + partes[1]

    return v

# =========================
# MOTOR DE PROCESSAMENTO
# =========================
def processar_bases_pdf(arquivo_enviado):
    dados = []
    registro = {}
    id_processados = set()
    campo_pendente = None

    matricula = ""
    nome = ""
    departamento = ""

    with pdfplumber.open(arquivo_enviado) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto: continue

            linhas = texto.split("\n")

            for linha in linhas:
                m = re.search(r"Funcionário:\s*(\d+)\s*-\s*(.*)", linha)
                if m:
                    matricula = m.group(1)
                    nome = m.group(2).split("Adm:")[0].strip()
                    registro = {}
                    continue

                m_dep = re.search(r"Departamento:\s*(.*)", linha)
                if m_dep:
                    departamento = m_dep.group(1).strip()

                if campo_pendente:
                    valor = re.search(r"([\d\.,]+)", linha)
                    if valor:
                        registro[campo_pendente] = valor.group(1)
                        campo_pendente = None
                        continue

                campos = re.findall(r"([A-Za-zÀ-ÿ0-9\s\(\)\.\-]+):\s*([\d\.,]+)?", linha)
                for campo, valor in campos:
                    campo = campo.strip().lower()
                    chave = None

                    if "salário base" in campo or "salario base" in campo: chave = "salario_base"
                    elif "base bruta de irrf" in campo: chave = "base_bruta_irrf"
                    elif "dedu" in campo: chave = "deducao_irrf"
                    elif "base de líquida" in campo or "base de liquida" in campo: chave = "base_liquida_irrf"
                    elif "total de vencimentos" in campo: chave = "total_vencimentos"
                    elif "base de inss funcionário" in campo or "base de inss funcionario" in campo: chave = "base_inss_funcionario"
                    elif "base inss empresa" in campo: chave = "base_inss_empresa"
                    elif "base terceiros" in campo: chave = "base_terceiros"
                    elif "base rat" in campo: chave = "base_rat"
                    elif "total de descontos" in campo: chave = "total_descontos"
                    elif "base de inss suspensa" in campo: chave = "base_inss_suspensa"
                    elif "horas semanais" in campo or "horas sernanais" in campo: chave = "horas_semanais"
                    elif "base de fgts" in campo: chave = "base_fgts"
                    elif "valor do fgts" in campo: chave = "valor_fgts"
                    elif "líquido a receber" in campo or "liquido a receber" in campo: chave = "liquido"

                    if chave:
                        if valor:
                            registro[chave] = valor
                        else:
                            campo_pendente = chave

                    if chave == "liquido":
                        if matricula not in id_processados:
                            dados.append({
                                "matricula": matricula,
                                "nome": nome,
                                "departamento": departamento,
                                **registro
                            })
                            id_processados.add(matricula)

    df = pd.DataFrame(dados)

    if "deducao_irrf" not in df.columns:
        df["deducao_irrf"] = None

    for col in df.columns:
        if col not in ["matricula", "nome", "departamento"]:
            df[col] = df[col].apply(normalizar_valor)
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Salva na memória para download web
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        
    return output.getvalue(), len(df)

# =========================
# INTERFACE DA PÁGINA WEB
# =========================
st.set_page_config(page_title="Extrator Bases PDF - TOTVS", page_icon="🧮")

st.title("🧮 Extrator de Bases (PDF)")
st.write("Envie os ficheiros **PDF** de Bases de Cálculo. O sistema irá extrair os dados e devolver o formato Excel.")

arquivos_enviados = st.file_uploader("Arraste e solte os ficheiros PDF aqui", type=["pdf"], accept_multiple_files=True)

if arquivos_enviados:
    if st.button("Processar Bases"):
        with st.spinner("A processar PDFs (isto pode demorar alguns segundos)..."):
            
            arquivos_processados = {}
            resumo_texto = ""
            
            for arquivo in arquivos_enviados:
                nome_original = arquivo.name
                nome_base = nome_original.rsplit(".", 1)[0]
                nome_novo = f"{nome_base}_bases_exportado.xlsx"
                
                try:
                    bytes_processados, qtd_func = processar_bases_pdf(arquivo)
                    arquivos_processados[nome_novo] = bytes_processados
                    resumo_texto += f"**{nome_original}**: {qtd_func} Funcionários extraídos.\n\n"
                except Exception as e:
                    st.error(f"Erro ao processar {nome_original}: {e}")

            if arquivos_processados:
                st.success("🎉 Processamento concluído com sucesso!")
                st.info(resumo_texto)
                
                if len(arquivos_processados) == 1:
                    nome_arquivo = list(arquivos_processados.keys())[0]
                    dados = arquivos_processados[nome_arquivo]
                    st.download_button(
                        label=f"⬇️ Descarregar {nome_arquivo}",
                        data=dados,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for nome, dados in arquivos_processados.items():
                            zip_file.writestr(nome, dados)
                    
                    st.download_button(
                        label="⬇️ Descarregar Todos (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="bases_pdf_exportadas.zip",
                        mime="application/zip"
                    )
