import streamlit as st
import pandas as pd
import re
import zipfile
import io

# =========================
# FUNÇÕES AUXILIARES
# =========================
def converter_numero(valor):
    if pd.isna(valor): return None
    if isinstance(valor, (int, float)): return float(valor)
    
    texto = str(valor).strip()
    match = re.search(r"([\d\.,]+)", texto)
    if match:
        num_str = match.group(1)
        if num_str in [".", ","]: return None
        if "," in num_str:
            num_str = num_str.replace(".", "").replace(",", ".")
        try: return float(num_str)
        except: return None
    return None

def extrair_valor_linha(linha_valores, palavras_chave):
    for idx, celula in enumerate(linha_valores):
        if pd.isna(celula): continue
        cel_str = str(celula).upper().strip()
        
        for p in palavras_chave:
            if p in cel_str:
                resto = cel_str.split(p)[1]
                val_inline = converter_numero(resto)
                if val_inline is not None:
                    return val_inline
                
                for prox_celula in linha_valores[idx+1:]:
                    val = converter_numero(prox_celula)
                    if val is not None:
                        return val
    return None

# =========================
# MOTOR DE PROCESSAMENTO
# =========================
def processar_bases_excel(arquivo_enviado):
    df = pd.read_excel(arquivo_enviado, header=None)
    
    dados = []
    registro_atual = {}
    
    for i in range(len(df)):
        linha_valores = df.iloc[i].values
        
        encontrou_func = False
        for celula in linha_valores:
            if pd.isna(celula): continue
            cel_str = str(celula).strip().upper()
            
            if re.match(r"^\d{1,8}\s*-", cel_str):
                if registro_atual.get("matricula"):
                    dados.append(registro_atual.copy())
                
                matricula = cel_str.split("-")[0].strip()
                resto = cel_str.split("-", 1)[1].strip()
                resto = resto.split("ADM:")[0].strip()
                
                nome_match = re.match(r"([A-ZÀ-Ÿ\s]+)", resto)
                nome = nome_match.group(1).strip() if nome_match else resto
                
                registro_atual = {
                    "matricula": matricula,
                    "nome": nome,
                    "cpf": "",
                    "total_proventos": None,
                    "total_descontos": None,
                    "salario_liquido": None,
                    "base_inss": None,
                    "base_fgts": None,
                    "base_irf": None
                }
                encontrou_func = True
                break
        
        if encontrou_func: continue
            
        if registro_atual.get("matricula"):
            for idx, celula in enumerate(linha_valores):
                if pd.isna(celula): continue
                cel_str = str(celula).strip().upper()
                if "CPF:" in cel_str:
                    match_cpf = re.search(r"CPF:\s*([\d\.\-]+)", cel_str)
                    if match_cpf:
                        registro_atual["cpf"] = match_cpf.group(1)
                    else:
                        if idx + 1 < len(linha_valores):
                            prox = str(linha_valores[idx+1]).strip()
                            match_cpf_prox = re.search(r"([\d\.\-]+)", prox)
                            if match_cpf_prox: registro_atual["cpf"] = match_cpf_prox.group(1)
                    break
        
        if registro_atual.get("matricula"):
            v = extrair_valor_linha(linha_valores, ["TOTAL DE PROVENTOS"])
            if v is not None: registro_atual["total_proventos"] = v
            
            v = extrair_valor_linha(linha_valores, ["TOTAL DE DESCONTOS"])
            if v is not None: registro_atual["total_descontos"] = v
            
            v = extrair_valor_linha(linha_valores, ["SALARIO LIQUIDO", "SALÁRIO LÍQUIDO"])
            if v is not None: registro_atual["salario_liquido"] = v
            
            v = extrair_valor_linha(linha_valores, ["BASE DO INSS", "BASE DE INSS"])
            if v is not None: registro_atual["base_inss"] = v
            
            v = extrair_valor_linha(linha_valores, ["BASE DO FGTS", "BASE DE FGTS"])
            if v is not None: registro_atual["base_fgts"] = v
            
            v = extrair_valor_linha(linha_valores, ["BASE DO IRF", "BASE DE IRF"])
            if v is not None: registro_atual["base_irf"] = v

    if registro_atual.get("matricula"):
        dados.append(registro_atual)
        
    df_final = pd.DataFrame(dados)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="Totais_e_Bases")
        
    return output.getvalue(), len(df_final)

# =========================
# INTERFACE DA PÁGINA WEB
# =========================
st.set_page_config(page_title="Extrator Totais Excel - TOTVS", page_icon="🧾")

st.title("🧾 Extrator de Bases e Totais (Excel)")
st.write("Envie os ficheiros Excel (.xlsx). O sistema irá extrair a consolidação final (Bases) numa única linha por funcionário.")

arquivos_enviados = st.file_uploader("Arraste e solte os ficheiros Excel aqui", type=["xlsx", "xls"], accept_multiple_files=True)

if arquivos_enviados:
    if st.button("Processar Totais"):
        with st.spinner("A processar os dados..."):
            arquivos_processados = {}
            resumo_texto = ""
            
            for arquivo in arquivos_enviados:
                nome_original = arquivo.name
                nome_base = nome_original.rsplit(".", 1)[0]
                nome_novo = f"{nome_base}_totais_exportado.xlsx"
                
                try:
                    bytes_processados, qtd_func = processar_bases_excel(arquivo)
                    arquivos_processados[nome_novo] = bytes_processados
                    resumo_texto += f"**{nome_original}**: {qtd_func} Funcionários consolidados.\n\n"
                except Exception as e:
                    st.error(f"Erro ao processar {nome_original}: {e}")

            if arquivos_processados:
                st.success("🎉 Processamento concluído com sucesso!")
                st.info(resumo_texto)
                
                if len(arquivos_processados) == 1:
                    nome_arquivo = list(arquivos_processados.keys())[0]
                    st.download_button(label=f"⬇️ Descarregar {nome_arquivo}", data=arquivos_processados[nome_arquivo], file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for nome, dados in arquivos_processados.items(): zip_file.writestr(nome, dados)
                    st.download_button(label="⬇️ Descarregar Todos (ZIP)", data=zip_buffer.getvalue(), file_name="totais_excel_exportados.zip", mime="application/zip")
