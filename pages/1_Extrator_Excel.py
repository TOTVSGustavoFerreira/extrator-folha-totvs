import streamlit as st
import pandas as pd
import re
import zipfile
import io

# =========================
# FUNÇÕES AUXILIARES (EXCEL)
# =========================
def eh_numero(valor):
    if pd.isna(valor): return False
    try:
        float(str(valor).replace(",", "."))
        return True
    except: return False

def converter_numero(valor):
    if pd.isna(valor): return None
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan": return None
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try: return float(texto)
    except: return None

def extrair_estab(texto):
    match = re.search(r"ESTAB:\s*(\d+)", texto.upper())
    return match.group(1) if match else None

# =========================
# MOTOR DE PROCESSAMENTO
# =========================
def processar_arquivo(arquivo_enviado):
    df = pd.read_excel(arquivo_enviado, header=None)
    
    dados_eventos = []
    dados_funcionarios = []
    estab_atual = None
    
    i = 0
    while i < len(df):
        linha = " ".join([str(x) for x in df.iloc[i].values])
        linha_upper = linha.upper()

        if "ESTAB:" in linha_upper:
            novo = extrair_estab(linha)
            if novo: estab_atual = novo
            i += 1
            continue

        if "-" in linha:
            parte_antes_hifen = linha.split("-")[0].replace("nan", "").replace("NAN", "").strip()
            if parte_antes_hifen.isdigit():
                texto = linha
                matricula = parte_antes_hifen
                
                try: nome = texto.split("-")[1].split("  ")[0].strip()
                except: nome = ""

                salario = re.search(r"\d{1,3}(\.\d{3})*,\d{2}", texto)
                salario = salario.group() if salario else ""

                data_adm = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
                data_adm = data_adm.group(1) if data_adm else ""

                secao = re.search(r"\d+\.\d+\.\d+\.\d+", texto)
                secao = secao.group() if secao else ""

                try: funcao = texto.split(data_adm)[-1].split(secao)[0].strip()
                except: funcao = ""

                try:
                    linha2 = " ".join([str(x) for x in df.iloc[i+1].values])
                    cpf_match = re.search(r"CPF:\s*([\d\.\-]+)", linha2)
                    cpf = cpf_match.group(1) if cpf_match else ""
                except: cpf = ""

                dados_funcionarios.append({
                    "estab": estab_atual, "matricula": matricula, "nome": nome,
                    "cpf": cpf, "funcao": funcao, "secao": secao,
                    "data_admissao": data_adm, "salario_contratual": converter_numero(salario)
                })

                j = i + 3
                while j < len(df):
                    linha_ev = df.iloc[j].values
                    texto_ev = " ".join([str(x) for x in linha_ev]).upper()
                    linha_ev_str = " ".join([str(x) for x in linha_ev])

                    if "-" in linha_ev_str:
                        parte_ev = linha_ev_str.split("-")[0].replace("nan", "").replace("NAN", "").strip()
                        if parte_ev.isdigit(): break
                            
                    if "ESTAB:" in texto_ev: break
                    if any(x in texto_ev for x in ["T O T A L   G E R A L", "LÍQUIDO", ">>>>", "BASES >>", "OCOR"]):
                        j += 1
                        continue

                    try:
                        if eh_numero(linha_ev[0]) or eh_numero(linha_ev[1]):
                            if eh_numero(linha_ev[0]):
                                evento = int(converter_numero(linha_ev[0]))
                                descricao = str(linha_ev[1]).strip()
                                referencia = converter_numero(linha_ev[5])
                                valor = converter_numero(linha_ev[6])
                            else:
                                evento = int(converter_numero(linha_ev[1]))
                                descricao = str(linha_ev[2]).strip()
                                referencia = converter_numero(linha_ev[6])
                                valor = converter_numero(linha_ev[7])

                            if valor is not None:
                                dados_eventos.append({"estab": estab_atual, "matricula": matricula, "nome": nome, "cpf": cpf, "funcao": funcao, "secao": secao, "evento": evento, "descricao": descricao, "refer": referencia, "valor": valor, "tipo": "PROVENTO"})
                    except: pass

                    try:
                        if eh_numero(linha_ev[7]) or eh_numero(linha_ev[8]):
                            if eh_numero(linha_ev[7]):
                                evento = int(converter_numero(linha_ev[7]))
                                descricao = str(linha_ev[8]).strip()
                                referencia = converter_numero(linha_ev[12])
                                valor = converter_numero(linha_ev[13])
                            else:
                                evento = int(converter_numero(linha_ev[8]))
                                descricao = str(linha_ev[9]).strip()
                                referencia = converter_numero(linha_ev[13])
                                valor = converter_numero(linha_ev[14])

                            if valor is not None:
                                dados_eventos.append({"estab": estab_atual, "matricula": matricula, "nome": nome, "cpf": cpf, "funcao": funcao, "secao": secao, "evento": evento, "descricao": descricao, "refer": referencia, "valor": valor, "tipo": "DESCONTO"})
                    except: pass
                    j += 1
                i = j
            else: i += 1 
        else: i += 1

    df_eventos_final = pd.DataFrame(dados_eventos)
    df_func_final = pd.DataFrame(dados_funcionarios).drop_duplicates(subset=["estab", "matricula"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_eventos_final.empty: df_eventos_final.to_excel(writer, sheet_name="Eventos", index=False)
        if not df_func_final.empty: df_func_final.to_excel(writer, sheet_name="Cadastro_Funcionarios", index=False)
    
    return output.getvalue()

# =========================
# INTERFACE DA PÁGINA WEB
# =========================
st.set_page_config(page_title="Extrator Excel - TOTVS", page_icon="📊")

st.title("📊 Extrator de Folha (Excel)")
st.write("Envie arquivos no formato **.xlsx ou .xls**.")

arquivos_enviados = st.file_uploader("Arraste e solte os arquivos Excel aqui", type=["xlsx", "xls"], accept_multiple_files=True)

if arquivos_enviados:
    if st.button("Processar Arquivos"):
        with st.spinner("A processar os dados..."):
            
            arquivos_processados = {}
            for arquivo in arquivos_enviados:
                nome_original = arquivo.name
                nome_base = nome_original.rsplit(".", 1)[0]
                nome_novo = f"{nome_base}_exportado.xlsx"
                
                try:
                    bytes_processados = processar_arquivo(arquivo)
                    arquivos_processados[nome_novo] = bytes_processados
                except Exception as e:
                    st.error(f"Erro ao processar {nome_original}: {e}")

            if arquivos_processados:
                st.success("🎉 Processamento concluído com sucesso!")
                
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
                        file_name="arquivos_exportados.zip",
                        mime="application/zip"
                    )
