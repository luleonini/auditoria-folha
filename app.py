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

def encontrar_coluna(df, palavras_chave):
    """Localiza automaticamente uma coluna na planilha ignorando maiúsculas e espaços extras."""
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if any(keyword in col_clean for keyword in palavras_chave):
            return col
    return None

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
            # 1. Leitura e identificação das colunas do Excel
            df_excel_raw = pd.read_excel(uploaded_excel)
            
            col_nome_excel = encontrar_coluna(df_excel_raw, ["nome", "colaborador", "funcionario"])
            col_valor_excel = encontrar_coluna(df_excel_raw, ["valor", "liquido", "líquido", "pago", "total"])

            if not col_nome_excel or not col_valor_excel:
                st.error(f"❌ Não foi possível identificar as colunas no Excel. Certifique-se de que existem colunas com 'Nome' e 'Valor'. Colunas encontradas: {list(df_excel_raw.columns)}")
                st.stop()

            # Normaliza o Excel encontrado
            df_excel = df_excel_raw.copy()
            df_excel["nome_excel_orig"] = df_excel[col_nome_excel]
            df_excel["valor_excel_orig"] = pd.to_numeric(df_excel[col_valor_excel].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
            df_excel["nome_clean"] = df_excel[col_nome_excel].astype(str).str.strip().str.upper()

            # 2. Leitura dos PDFs
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
                else:
                    # Registra se o PDF estiver fora do padrão esperado
                    dados_pdfs.append({
                        "nome_clean": pdf_file.name.upper(),
                        "nome_pdf": f"Erro de Leitura ({pdf_file.name})",
                        "valor_pdf": 0.0,
                        "mes_ref": "N/A"
                    })

            df_pdf = pd.DataFrame(dados_pdfs)

            # 3. Cruzamento dos Dados
            df_final = pd.merge(df_pdf, df_excel, on="nome_clean", how="outer")

            def checar_status(row):
                if pd.isna(row.get("nome_pdf")) or "Erro de Leitura" in str(row.get("nome_pdf")):
                    return "FALTANDO PDF OU FORMATO INVÁLIDO"
                if pd.isna(row.get("nome_excel_orig")):
                    return "NÃO ENCONTRADO NO EXCEL"
                
                v_pdf = row.get("valor_pdf", 0.0)
                v_excel = row.get("valor_excel_orig", 0.0)
                
                if abs(v_pdf - v_excel) > 0.01:
                    return f"DIVERGÊNCIA (PDF: R$ {v_pdf:.2f} | EXCEL: R$ {v_excel:.2f})"
                return "OK"

            df_final["Status_Auditoria"] = df_final.apply(checar_status, axis=1)

            # Formatação final da tabela exibida
            df_exibicao = pd.DataFrame({
                "Nome (PDF)": df_final["nome_pdf"].fillna("-"),
                "Mês Ref.": df_final["mes_ref"].fillna("-"),
                "Valor (PDF)": df_final["valor_pdf"].fillna(0.0),
                "Nome (Excel)": df_final["nome_excel_orig"].fillna("-"),
                "Valor (Excel)": df_final["valor_excel_orig"].fillna(0.0),
                "Status da Auditoria": df_final["Status_Auditoria"]
            })

            st.success("✅ Auditoria realizada com sucesso!")
            
            # Exibição dos resultados
            st.subheader("Resultado do Cruzamento")
            st.dataframe(df_exibicao, use_container_width=True)

            # Botão de Download
            csv_data = df_exibicao.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button(
                label="📥 Baixar Relatório de Divergências (.csv)",
                data=csv_data,
                file_name="resultado_auditoria.csv",
                mime="text/csv"
            )
