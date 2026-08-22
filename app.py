import streamlit as st
import pandas as pd
import json
import os
from google import genai
from google.genai import types
from pydantic import BaseModel

st.set_page_config(page_title="Auditoria de Folha com IA", page_icon="🤖", layout="wide")

st.title("🤖 Auditoria Automática de Folha de Pagamento (com IA)")
st.markdown("Auditoria inteligente: A IA lê **qualquer formato de holerite em PDF** e cruza os dados com o Excel.")

# Entrada da Chave de API do Gemini (Permite ao usuário colar sua API Key ou usar uma do ambiente)
api_key = st.sidebar.text_input("Cole sua Gemini API Key:", type="password")

class DadosHolerite(BaseModel):
    nome: str
    valor_liquido: float
    mes_referencia: str

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

if st.button("🚀 Processar Auditoria com IA", type="primary"):
    if not api_key:
        st.error("⚠️ Insira uma chave de API do Gemini na barra lateral para continuar.")
    elif not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Envie os arquivos PDF e a planilha Excel antes de processar.")
    else:
        client = genai.Client(api_key=api_key)
        dados_pdfs = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, pdf_file in enumerate(uploaded_pdfs):
            status_text.text(f"Analizando com IA: {pdf_file.name}...")
            pdf_bytes = pdf_file.read()

            prompt = (
                "Analise este holerite/folha de pagamento. Extraia o nome completo do colaborador, "
                "o valor líquido a receber (como float numérico puro) e o mês de referência."
            )

            try:
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[
                        types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                        prompt
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=DadosHolerite,
                        temperature=0.0
                    )
                )
                
                dados_json = json.loads(response.text)
                dados_pdfs.append({
                    "nome_clean": dados_json["nome"].strip().upper(),
                    "nome_pdf": dados_json["nome"].strip(),
                    "valor_pdf": float(dados_json["valor_liquido"]),
                    "mes_ref": dados_json["mes_referencia"]
                })
            except Exception as e:
                dados_pdfs.append({
                    "nome_clean": pdf_file.name.upper(),
                    "nome_pdf": f"Erro de Leitura ({pdf_file.name})",
                    "valor_pdf": 0.0,
                    "mes_ref": "N/A"
                })
            
            progress_bar.progress((idx + 1) / len(uploaded_pdfs))

        status_text.text("Concluindo cruzamento de dados...")

        # Processar Excel
        df_excel_raw = pd.read_excel(uploaded_excel)
        
        # Identificação de colunas
        col_nome = [c for c in df_excel_raw.columns if "nome" in str(c).lower() or "colaborador" in str(c).lower()][0]
        col_valor = [c for c in df_excel_raw.columns if "valor" in str(c).lower() or "liquido" in str(c).lower() or "líquido" in str(c).lower()][0]

        df_excel = df_excel_raw.copy()
        df_excel["nome_excel_orig"] = df_excel[col_nome]
        
        # Tratamento correto de valores monetários no Excel
        def limpar_valor_excel(v):
            if pd.isna(v): return 0.0
            if isinstance(v, (int, float)): return float(v)
            v_str = str(v).replace("R$", "").strip()
            if "," in v_str and "." in v_str:
                v_str = v_str.replace(".", "").replace(",", ".")
            elif "," in v_str:
                v_str = v_str.replace(",", ".")
            return float(v_str)

        df_excel["valor_excel_orig"] = df_excel[col_valor].apply(limpar_valor_excel)
        df_excel["nome_clean"] = df_excel[col_nome].astype(str).str.strip().str.upper()

        # Cruzamento
        df_final = pd.merge(pd.DataFrame(dados_pdfs), df_excel, on="nome_clean", how="outer")

        def checar_status(row):
            if pd.isna(row.get("nome_pdf")) or "Erro de Leitura" in str(row.get("nome_pdf")):
                return "FALTANDO PDF"
            if pd.isna(row.get("nome_excel_orig")):
                return "NÃO ENCONTRADO NO EXCEL"
            
            v_pdf = row.get("valor_pdf", 0.0)
            v_excel = row.get("valor_excel_orig", 0.0)
            
            if abs(v_pdf - v_excel) > 0.01:
                return f"DIVERGÊNCIA (PDF: R$ {v_pdf:,.2f} | EXCEL: R$ {v_excel:,.2f})"
            return "OK"

        df_final["Status_Auditoria"] = df_final.apply(checar_status, axis=1)

        df_exibicao = pd.DataFrame({
            "Nome (PDF)": df_final["nome_pdf"].fillna("-"),
            "Mês Ref.": df_final["mes_ref"].fillna("-"),
            "Valor (PDF)": df_final["valor_pdf"].fillna(0.0),
            "Nome (Excel)": df_final["nome_excel_orig"].fillna("-"),
            "Valor (Excel)": df_final["valor_excel_orig"].fillna(0.0),
            "Status da Auditoria": df_final["Status_Auditoria"]
        })

        st.success("✅ Auditoria realizada com sucesso via IA!")
        st.dataframe(df_exibicao, use_container_width=True)

        csv_data = df_exibicao.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="📥 Baixar Relatório de Divergências (.csv)",
            data=csv_data,
            file_name="resultado_auditoria.csv",
            mime="text/csv"
        )
