import streamlit as st
import pdfplumber
import pandas as pd
import re
import zipfile
import io

# =========================
# FUNÇÕES AUXILIARES (PDF)
# =========================
def eh_valor(token):
    return bool(re.match(r"\d{1,3}(\.\d{3})*,\d{2}$", token))

def eh_referencia(token):
    if "/" in token or "%" in token or ":" in token:
        return True
    if re.match(r"^\d+$", token):
        return True
    return False

def extrair_eventos(linha):
    eventos = re.findall(r"\d{5}.*?(?=\d{5}|$)", linha)
    resultados = []
    
    for ev in eventos:
        partes = ev.split()
        if len(partes) < 2:
            continue
            
        codigo = partes[0]
        valor = ""
        referencia = ""
        evento_tokens = []
        
        for token in partes[1:]:
            if eh_valor(token):
                valor = token
                break
            elif eh_referencia(token):
                referencia = token
            else:
                evento_tokens.append(token)
                
        evento = " ".join(evento_tokens)
        resultados.append((codigo, evento, referencia, valor))
        
    return resultados

# =========================
# MOTOR DE PROCESSAMENTO
# =========================
def processar_pdf(arquivo_enviado):
    dados = []
    controle_eventos = set()
    funcionarios = {}
    
    matricula = ""
    nome = ""
    admissao = ""
    funcao = ""
    
    ignorar_totalizacao = False
    x_venc = None
    x_desc = None
    x_mid = None

    with pdfplumber.open(arquivo_enviado) as pdf:
        for pagina in pdf.pages:
            palavras = pagina.extract_words()
            for w in palavras:
                txt = w["text"].upper()
                if txt.startswith("VENCIMENT"):
                    x_venc = w["x0"]
                if txt.startswith("DESCONT"):
                    x_desc = w["x0"]
                    
            if x_venc and x_desc:
                x_mid = (x_venc + x_desc) / 2
                break

        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
                
            linhas = texto.split("\n")
            palavras = pagina.extract_words()
            
            for linha in linhas:
                if "TOTALIZAÇÃO DA FOLHA" in linha.upper():
                    ignorar_totalizacao = True
                    
                if ignorar_totalizacao:
                    continue
                    
                m = re.search(r"Funcionário:\s*(\d+)\s*-\s*(.*?)\s+Adm:\s*([\d/]+).*?Função:\s*(.*)", linha)
                
                if m:
                    matricula = m.group(1)
                    nome = m.group(2).strip()
                    admissao = m.group(3)
                    funcao = m.group(4).strip()
                    funcionarios[matricula] = nome
                    continue
                    
                eventos = extrair_eventos(linha)
                
                for codigo, evento, referencia, valor in eventos:
                    tipo = "DESCONHECIDO"
                    
                    for w in palavras:
                        if w["text"] == codigo:
                            if w["x0"] < x_mid:
                                tipo = "PROVENTO"
                            else:
                                tipo = "DESCONTO"
                            break
                            
                    chave = (matricula, codigo, valor)
                    if chave in controle_eventos:
                        continue
                        
                    controle_eventos.add(chave)
                    
                    dados.append({
                        "matricula": matricula,
                        "nome": nome,
                        "admissao": admissao,
                        "funcao": funcao,
                        "codigo_evento": codigo,
                        "evento": evento,
                        "tipo": tipo,
                        "referencia": referencia,
                        "valor": valor,
                    })

    df_eventos = pd.DataFrame(dados)
    
    if not df_eventos.empty:
        df_eventos["valor"] = (
            df_eventos["valor"]
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_eventos["valor"] = pd.to_numeric(df_eventos["valor"], errors="coerce")

    df_funcionarios = pd.DataFrame(
        [(k, v) for k, v in funcionarios.items()],
        columns=["matricula", "nome"]
    ).sort_values("matricula")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_eventos.empty:
            df_eventos.to_excel(writer, sheet_name="EVENTOS", index=False)
        if not df_funcionarios.empty:
            df_funcionarios.to_excel(writer, sheet_name="FUNCIONARIOS", index=False)
            
    return output.getvalue(), len(df_funcionarios), len(df_eventos)

# =========================
# INTERFACE DA PÁGINA WEB
# =========================
st.set_page_config(page_title="Extrator PDF - TOTVS", page_icon="📄")

st.title("📄 Extrator de Folha (PDF)")
st.write("Envie arquivos no formato **.pdf**.")

arquivos_enviados = st.file_uploader("Arraste e solte os arquivos PDF aqui", type=["pdf"], accept_multiple_files=True)

if arquivos_enviados:
    if st.button("Processar Arquivos"):
        with st.spinner("A processar PDFs (isto pode demorar alguns segundos)..."):
            
            arquivos_processados = {}
            resumo_texto = ""
            
            for arquivo in arquivos_enviados:
                nome_original = arquivo.name
                nome_base = nome_original.rsplit(".", 1)[0]
                nome_novo = f"{nome_base}_exportado.xlsx"
                
                try:
                    bytes_processados, qtd_func, qtd_ev = processar_pdf(arquivo)
                    arquivos_processados[nome_novo] = bytes_processados
                    resumo_texto += f"**{nome_original}**: {qtd_func} Funcionários | {qtd_ev} Registos extraídos.\n\n"
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
                        file_name="folhas_pdf_exportadas.zip",
                        mime="application/zip"
                    )
