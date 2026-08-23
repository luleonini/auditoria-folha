import io
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
from google import genai
from google.genai import types

# Configuração da página
st.set_page_config(
    page_title="Auditoria de Folha de Pagamento",
    page_icon="🔍",
    layout="wide"
)

# Estilização visual (Status de Auditoria)
def destacar_divergencias(val):
    if val != "OK":
        return 'background-color: #FCE4D6; color: #C00000; font-weight: bold;'
    return ''

# Interface Sidebar - Privacidade & Configuração
st.sidebar.title("🔒 Segurança & Configuração")
st.sidebar.info(
    "**Conformidade com a LGPD:**\n"
    "Esta aplicação processa os arquivos exclusivamente em memória RAM. "
    "Nenhum dado financeiro ou pessoal é armazenado no servidor."
)

# Gerenciamento da Chave de API
api_key = st.sidebar.text_input("Gemini API Key", type="password", help="Chave corporativa da Gemini API")
if not api_key:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]

# Botão de Limpeza Manual de Dados
if st.sidebar.button("🧹 Limpar Dados em Memória"):
    st.session_state.clear()
    st.rerun()

st.title("🔍 Auditoria Automática de Folha de Pagamento")
st.markdown("Suba a planilha base em Excel e os PDFs dos holerites para realizar a conciliação automática.")

# Carregamento de Arquivos em Memória
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. Planilha Base (Excel)", type=["xlsx", "xls"])
with col2:
    pdf_files = st.file_uploader("2. Holerites / Folha (PDFs)", type=["pdf"], accept_multiple_files=True)

# Leitura do Prompt de Instrução
try:
    with open("prompt.txt", "r", encoding="utf-8") as f:
        prompt_instrucoes = f.read()
except FileNotFoundError:
    st.error("Erro: Arquivo 'prompt.txt' não encontrado na raiz da aplicação.")
    st.stop()

# Processamento da Auditoria
if st.button("🚀 Processar Auditoria", type="primary"):
    if not api_key:
        st.error("Por favor, informe a chave da Gemini API na barra lateral ou configure nos Secrets.")
        st.stop()
    if not excel_file or not pdf_files:
        st.warning("É necessário carregar a planilha Excel e pelo menos um arquivo PDF.")
        st.stop()

    with st.spinner("Processando arquivos exclusivamente em memória..."):
        try:
            # Inicializa Cliente Gemini
            client = genai.Client(api_key=api_key)

            # Prepara os arquivos diretamente em memória para envio à API
            contents = [prompt_instrucoes]
            
            # Adiciona a planilha Excel (lida como bytes)
            contents.append(
                types.Part.from_bytes(
                    data=excel_file.getvalue(),
                    mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            )

            # Adiciona os PDFs (lidos como bytes)
            for pdf in pdf_files:
                contents.append(
                    types.Part.from_bytes(
                        data=pdf.getvalue(),
                        mime_type="application/pdf"
                    )
                )

            # Chamada ao Modelo Gemini 2.5 Flash
            response = client.models.generate_content(
                model="gemini-2.5-flash", 
                contents=contents
            )

            # Salva o resultado do texto na sessão
            st.session_state["resultado_ia"] = response.text
            st.success("Auditoria concluída com sucesso!")

        except Exception as e:
            st.error(f"Erro durante o processamento: {str(e)}")

# Exibição dos Resultados (se existirem na sessão)
if "resultado_ia" in st.session_state:
    st.markdown("### 📊 Relatório Sintético da Auditoria")
    st.markdown(st.session_state["resultado_ia"])
