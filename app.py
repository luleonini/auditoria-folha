import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

# Expressões regulares aprimoradas para capturar Nome, Valor e Mês/Ano
PATRAO_NOME = r"(?:Nome|Colaborador|Funcionário):\s*([A-Za-zÀ-ÿ\s]+)"
PATRAO_VALOR = r"(?:Líquido|Total Líquido|Líquido a Receber|Valor Líquido):\s*R?\$\s*([\d\.\,]+)"
PATRAO_MES = r"(?:Mês/Ano|Referência|Ref\.|Período):\s*(\d{2}/\d{4})"

def extrair_dados_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    texto = "".join([page.extract_text() or "" for page in reader.pages])
    
    # Busca via Regex
    nome = re.search(PATRAO_NOME, texto, re.IGNORECASE)
    valor = re.search(PATRAO_VALOR, texto, re.IGNORECASE)
    mes = re.search(PATRAO_MES, texto, re.IGNORECASE)

    # Fallback simples caso os rótulos exatos não sejam encontrados
    if not mes:
        mes_match = re.search(r"\b(\d{2}/\d{4})\b", texto)
        mes_val = mes_match.group(1) if mes_match else "N/A"
    else:
        mes_val = mes.group(1)

    nome_val = nome.group(1).strip() if nome else "NÃO IDENTIFICADO"
    
    val_clean = 0.0
    if valor:
        val_str = valor.group(1).replace(".", "").replace(",", ".")
        try:
            val_clean = float(val_str)
        except ValueError:
            val_clean = 0.0

    return {
        "nome_pdf": nome_val,
        "nome_clean": nome_val.upper(),
        "valor_pdf": val_clean,
        "mes_ref": mes_val,
        "chave_cruzamento": f"{nome_val.upper()}_{mes_val}"
    }

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

if st.button("🚀 Processar Auditoria", type="primary"):
    if not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Por favor, envie tanto os arquivos PDF quanto a planilha Excel antes de processar.")
    else:
        # 1. Processar PDFs
        dados_pdfs = [extrair_dados_pdf(pdf) for pdf in uploaded_pdfs]
        df_pdf = pd.DataFrame(dados_pdfs)

        # 2. Processar Excel
        df_excel_raw = pd.read_excel(uploaded_excel)
        
        # Identificar colunas no Excel
        col_nome = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["nome", "colaborador", "funcionario"])][0]
        col_valor = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["valor", "liquido", "líquido", "total"])][0]
        col_mes = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["mês", "mes", "referencia", "referência", "periodo"])][0]

        df_excel = df_excel_raw.copy()
        df_excel["nome_excel_orig"] = df_excel[col_nome]
        df_excel["valor_excel_orig"] = pd.to_numeric(
            df_excel[col_valor].astype(str).str.replace("R$", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).str.strip(),
            errors="coerce"
        ).fillna(0.0)
        
        df_excel["mes_excel"] = df_excel[col_mes].astype(str).str.strip()
        df_excel["chave_cruzamento"] = df_excel[col_nome].astype(str).str.strip().str.upper() + "_" + df_excel["mes_excel"]

        # 3. Cruzamento por Chave Composta (Nome + Mês)
        df_final = pd.merge(df_pdf, df_excel, on="chave_cruzamento", how="outer")

        def checar_status(row):
            if pd.isna(row.get("nome_pdf")) or row.get("nome_pdf") == "NÃO IDENTIFICADO":
                return "FALTANDO PDF"
            if pd.isna(row.get("nome_excel_orig")):
                return "NÃO ENCONTRADO NO EXCEL"
            
            v_pdf = row.get("valor_pdf", 0.0)
            v_excel = row.get("valor_excel_orig", 0.0)
            dif = abs(v_pdf - v_excel)
            
            if dif > 0.01:
                return f"Divergência de Valor (Dif: R$ {dif:,.2f})"
            return "OK"

        df_final["Status_Auditoria"] = df_final.apply(checar_status, axis=1)

        # Montar a tabela idêntica ao resumo
        df_exibicao = pd.DataFrame({
            "Mês Ref.": df_final["mes_ref"].fillna(df_final["mes_excel"]),
            "Nome": df_final["nome_pdf"].fillna(df_final["nome_excel_orig"]),
            "Valor (PDF)": df_final["valor_pdf"].map("R$ {:,.2f}".format),
            "Valor (Excel)": df_final["valor_excel_orig"].map("R$ {:,.2f}".format),
            "Status": df_final["Status_Auditoria"]
        })

        st.success("✅ Auditoria realizada com sucesso!")
        
        # Métrica consolidada (como na Imagem 1)
        total_regs = len(df_final)
        conciliados = len(df_final[df_final["Status_Auditoria"] == "OK"])
        inconsistentes = total_regs - conciliados

        st.markdown(f"**Total de registros auditados:** {total_regs} colaboradores/períodos")
        st.markdown(f"**Registros conciliados (OK):** {conciliados}")
        st.markdown(f"**Inconsistências encontradas:** {inconsistentes}")

        st.subheader("Resumo do Cruzamento:")
        st.dataframe(df_exibicao, use_container_width=True)

        # Download em CSV
        csv_data = df_exibicao.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="📥 Baixar Relatório Completo (.csv)",
            data=csv_data,
            file_name="resultado_auditoria.csv",
            mime="text/csv"
        )
