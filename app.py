import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

# Regex para captura dos dados nos PDFs
PATRAO_NOME = r"Nome:\s*([A-Za-zÀ-ÿ\s]+)"
PATRAO_VALOR = r"(?:Líquido|Total Líquido|Líquido a Receber):\s*R?\$\s*([\d\.\,]+)"
PATRAO_MES = r"(?:Mês/Ano|Referência):\s*(\d{2}/\d{4})"

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

if st.button("🚀 Processar Auditoria", type="primary"):
    if not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Por favor, envie tanto os arquivos PDF quanto a planilha Excel antes de processar.")
    else:
        with st.spinner("Processando arquivos..."):
            dados_pdfs = []
            
            for pdf_file in uploaded_pdfs:
                reader = PdfReader(pdf_file)
                texto = "".join([page.extract_text() or "" for page in reader.pages])
                
                nome = re.search(PATRAO_NOME, texto)
                valor = re.search(PATRAO_VALOR, texto)
                mes = re.search(PATRAO_MES, texto)

                if nome and valor:
                    val_clean = float(valor.group(1).replace(".", "").replace(",", "."))
                    dados_pdfs.append({
                        "nome_clean": nome.group(1).strip().upper(),
                        "nome_pdf": nome.group(1).strip(),
                        "valor_pdf": val_clean,
                        "mes_ref": mes.group(1) if mes else "N/A"
                    })

            df_pdf = pd.DataFrame(dados_pdfs)
            df_excel = pd.read_excel(uploaded_excel)
            df_excel["nome_clean"] = df_excel["Nome"].astype(str).str.strip().str.upper()

            df_final = pd.merge(df_pdf, df_excel, on="nome_clean", how="outer")

            def checar_status(row):
                if pd.isna(row.get("nome_pdf")):
                    return "FALTANDO PDF"
                if pd.isna(row.get("Nome")):
                    return "NÃO ENCONTRADO NO EXCEL"
                if abs(row["valor_pdf"] - row["Valor"]) > 0.01:
                    return f"DIVERGÊNCIA (PDF: R$ {row['valor_pdf']:.2f} | EXCEL: R$ {row['Valor']:.2f})"
                return "OK"

            df_final["Status_Auditoria"] = df_final.apply(checar_status, axis=1)

            st.success("✅ Auditoria realizada com sucesso!")
            
            # Exibição dos resultados
            st.subheader("Resultado do Cruzamento")
            st.dataframe(df_final[["nome_clean", "valor_pdf", "Valor", "Status_Auditoria"]], use_container_width=True)

            # Botão de Download
            csv_data = df_final.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button(
                label="📥 Baixar Relatório de Divergências (.csv)",
                data=csv_data,
                file_name="resultado_auditoria.csv",
                mime="text/csv"
            )
