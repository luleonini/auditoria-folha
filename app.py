import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

PATRAO_NOME = r"(?:Nome|Colaborador|Funcionário):\s*([A-Za-zÀ-ÿ\s]+)"
PATRAO_VALOR = r"(?:Líquido|Total Líquido|Líquido a Receber|Valor Líquido):\s*R?\$\s*([\d\.\,]+)"
PATRAO_MES = r"(?:Mês/Ano|Referência|Ref\.|Período):\s*(\d{2}/\d{4})"

def extrair_dados_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    texto = "".join([page.extract_text() or "" for page in reader.pages])
    
    nome = re.search(PATRAO_NOME, texto, re.IGNORECASE)
    valor = re.search(PATRAO_VALOR, texto, re.IGNORECASE)
    mes = re.search(PATRAO_MES, texto, re.IGNORECASE)

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

def buscar_coluna_segura(colunas, palavras_chave):
    for col in colunas:
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
        # 1. Leitura dos PDFs
        dados_pdfs = [extrair_dados_pdf(pdf) for pdf in uploaded_pdfs]
        df_pdf = pd.DataFrame(dados_pdfs)

        # 2. Leitura do Excel
        df_excel_raw = pd.read_excel(uploaded_excel)
        
        col_nome = buscar_coluna_segura(df_excel_raw.columns, ["nome", "colaborador", "funcionario"])
        col_valor = buscar_coluna_segura(df_excel_raw.columns, ["valor", "liquido", "líquido", "total", "pago"])
        col_mes = buscar_coluna_segura(df_excel_raw.columns, ["mês", "mes", "referencia", "referência", "periodo", "data", "ano"])

        if not col_nome or not col_valor:
            st.error(f"❌ Não foi possível encontrar as colunas de Nome e Valor no Excel. Colunas encontradas: {list(df_excel_raw.columns)}")
            st.stop()

        df_excel = df_excel_raw.copy()
        df_excel["nome_excel_orig"] = df_excel[col_nome]
        df_excel["valor_excel_orig"] = pd.to_numeric(
            df_excel[col_valor].astype(str).str.replace("R$", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).str.strip(),
            errors="coerce"
        ).fillna(0.0)
        
        if col_mes:
            df_excel["mes_excel"] = df_excel[col_mes].astype(str).str.strip()
            df_excel["chave_cruzamento"] = df_excel[col_nome].astype(str).str.strip().str.upper() + "_" + df_excel["mes_excel"]
            modo_cruzamento = "chave_cruzamento"
        else:
            df_excel["nome_clean"] = df_excel[col_nome].astype(str).str.strip().str.upper()
            modo_cruzamento = "nome_clean"

        # 3. Cruzamento
        df_final = pd.merge(df_pdf, df_excel, on=modo_cruzamento, how="outer")

        # Tratamento de campos nulos após o merge
        df_final["valor_pdf"] = df_final["valor_pdf"].fillna(0.0)
        df_final["valor_excel_orig"] = df_final["valor_excel_orig"].fillna(0.0)
        
        if "mes_ref" in df_final.columns:
            df_final["mes_ref"] = df_final["mes_ref"].fillna(df_final.get("mes_excel", "N/A"))
        else:
            df_final["mes_ref"] = df_final.get("mes_excel", "N/A")

        df_final["nome_exibicao"] = df_final["nome_pdf"].fillna(df_final["nome_excel_orig"])

        # Cálculo de divergências
        df_final["diferenca"] = df_final["valor_pdf"] - df_final["valor_excel_orig"]
        df_final["abs_dif"] = df_final["diferenca"].abs()

        def definir_status(row):
            if pd.isna(row.get("nome_pdf")) or row.get("nome_pdf") == "NÃO IDENTIFICADO":
                return "Faltando PDF"
            if pd.isna(row.get("nome_excel_orig")):
                return "Não encontrado no Excel"
            if row["abs_dif"] > 0.01:
                return "Divergência de Valor"
            return "OK"

        df_final["status"] = df_final.apply(definir_status, axis=1)

        # Totais para os marcadores superiores
        total_regs = len(df_final)
        df_ok = df_final[df_final["status"] == "OK"]
        df_inconsistentes = df_final[df_final["status"] != "OK"]

        conciliados_count = len(df_ok)
        inconsistentes_count = len(df_inconsistentes)

        # Construção da lista de texto formatada exatamente como a imagem
        st.markdown(f"* **Total de registros auditados:** {total_regs} colaboradores/períodos")
        
        if conciliados_count > 0:
            refs_ok = ", ".join([f"Referência {m}" for m in df_ok["mes_ref"].unique()])
            st.markdown(f"* **Registros conciliados (OK):** {conciliados_count} ({refs_ok})")
        else:
            st.markdown(f"* **Registros conciliados (OK):** 0")

        if inconsistentes_count > 0:
            refs_inc = ", ".join([f"Referência {row['mes_ref']} com {row['status'].lower()}" for _, row in df_inconsistentes.iterrows()])
            st.markdown(f"* **Inconsistências encontradas:** {inconsistentes_count} ({refs_inc})")
        else:
            st.markdown(f"* **Inconsistências encontradas:** 0")

        st.subheader("Resumo do Cruzamento:")

        # Formatação das linhas idêntica à imagem
        for _, row in df_final.iterrows():
            mes = row["mes_ref"]
            nome = row["nome_exibicao"]
            v_pdf = f"R$ {row['valor_pdf']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            v_excel = f"R$ {row['valor_excel_orig']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            dif_val = f"R$ {row['abs_dif']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            if row["status"] == "OK":
                status_txt = "*(OK)*"
            else:
                status_txt = f"*({row['status']})*"

            st.markdown(
                f"* **{mes}**: {nome} — PDF: **{v_pdf}** | Excel: **{v_excel}** | Diferença: **{dif_val}** {status_txt}"
            )

        # Botão de Download do CSV
        csv_data = df_final.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="📥 Baixar Relatório Completo (.csv)",
            data=csv_data,
            file_name="resultado_auditoria.csv",
            mime="text/csv"
        )
