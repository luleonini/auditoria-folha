import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

def extrair_dados_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    texto = "\n".join([page.extract_text() or "" for page in reader.pages])

    # 1. Busca de Nome
    pats_nome = [
        r"(?:Nome|Colaborador|Funcionário):\s*([A-Za-zÀ-ÿ\s]+)",
        r"([A-ZÀ-ÿ\s]{8,})",
    ]
    nome_val = "NÃO IDENTIFICADO"
    for pat in pats_nome:
        m = re.search(pat, texto, re.IGNORECASE)
        if m and len(m.group(1).strip()) > 5:
            nome_val = m.group(1).strip().split("\n")[0]
            break

    # 2. Busca de Mês/Ano (MM/AAAA)
    mes_match = re.search(r"\b(0[1-9]|1[0-2])/(20\d{2})\b", texto)
    mes_val = mes_match.group(0) if mes_match else "N/A"

    # 3. Extração Inteligente do Valor Líquido
    pats_valor = [
        r"(?:Líquido|Total Líquido|Líquido a Receber|Valor Líquido)[:\s]*R?\$?\s*([\d\.\,]+)",
        r"R\$\s*([\d\.\,]+)",
    ]
    
    val_clean = 0.0
    encontrou_valor = False
    
    for pat in pats_valor:
        matches = re.findall(pat, texto, re.IGNORECASE)
        if matches:
            for m in reversed(matches):
                v_str = m.replace(".", "").replace(",", ".")
                try:
                    v_float = float(v_str)
                    if v_float > 100:  # Ignora valores irrelevantes
                        val_clean = v_float
                        encontrou_valor = True
                        break
                except ValueError:
                    continue
        if encontrou_valor:
            break

    return {
        "nome_pdf": nome_val,
        "nome_clean": nome_val.upper(),
        "valor_pdf": val_clean,
        "mes_ref": mes_val,
        "chave_cruzamento": f"{nome_val.upper()}_{mes_val}"
    }

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

if st.button("🚀 Processar Auditoria", type="primary"):
    if not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Envie os arquivos PDF e a planilha Excel antes de processar.")
    else:
        # 1. Processar PDFs
        dados_pdfs = [extrair_dados_pdf(pdf) for pdf in uploaded_pdfs]
        df_pdf = pd.DataFrame(dados_pdfs)

        # 2. Processar Excel
        df_excel_raw = pd.read_excel(uploaded_excel)

        # Identificar colunas no Excel
        col_nome = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["nome", "colaborador", "funcionario"])][0]
        col_valor = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["valor", "liquido", "líquido", "total"])][0]
        col_mes_search = [c for c in df_excel_raw.columns if any(k in str(c).lower() for k in ["mês", "mes", "referencia", "referência", "periodo"])]
        
        col_mes = col_mes_search[0] if col_mes_search else None

        df_excel = df_excel_raw.copy()
        df_excel["nome_excel_orig"] = df_excel[col_nome]

        def converter_valor_excel(val):
            if pd.isna(val): return 0.0
            if isinstance(val, (int, float)): return float(val)
            v_str = str(val).replace("R$", "").strip()
            if "," in v_str and "." in v_str:
                v_str = v_str.replace(".", "").replace(",", ".")
            elif "," in v_str:
                v_str = v_str.replace(",", ".")
            return float(v_str)

        df_excel["valor_excel_orig"] = df_excel[col_valor].apply(converter_valor_excel)

        if col_mes:
            df_excel["mes_excel"] = df_excel[col_mes].astype(str).str.strip()
            df_excel["chave_cruzamento"] = df_excel[col_nome].astype(str).str.strip().str.upper() + "_" + df_excel["mes_excel"]
            df_final = pd.merge(df_pdf, df_excel, on="chave_cruzamento", how="outer")
        else:
            df_excel["nome_clean"] = df_excel[col_nome].astype(str).str.strip().str.upper()
            df_final = pd.merge(df_pdf, df_excel, on="nome_clean", how="outer")

        # Ajustes de dados pós-merge
        df_final["valor_pdf"] = df_final["valor_pdf"].fillna(0.0)
        df_final["valor_excel_orig"] = df_final["valor_excel_orig"].fillna(0.0)

        if "mes_ref" in df_final.columns and "mes_excel" in df_final.columns:
            df_final["mes_final"] = df_final["mes_ref"].fillna(df_final["mes_excel"])
        elif "mes_ref" in df_final.columns:
            df_final["mes_final"] = df_final["mes_ref"]
        else:
            df_final["mes_final"] = df_final.get("mes_excel", "N/A")

        df_final["nome_final"] = df_final["nome_pdf"].fillna(df_final["nome_excel_orig"])
        df_final["diferenca"] = (df_final["valor_pdf"] - df_final["valor_excel_orig"]).abs()

        def checar_status(row):
            if pd.isna(row.get("nome_pdf")) or row.get("nome_pdf") == "NÃO IDENTIFICADO":
                return "Faltando PDF"
            if pd.isna(row.get("nome_excel_orig")):
                return "Não encontrado no Excel"
            if row["diferenca"] > 0.01:
                return "Divergência de Valor"
            return "OK"

        df_final["status"] = df_final.apply(checar_status, axis=1)

        # Totais para exibição
        total_regs = len(df_final)
        df_ok = df_final[df_final["status"] == "OK"]
        df_inc = df_final[df_final["status"] != "OK"]

        # Formatação idêntica à Imagem Solicitada
        st.markdown(f"* **Total de registros auditados:** {total_regs} colaboradores/períodos")

        if len(df_ok) > 0:
            refs_ok = ", ".join([f"Referência {m}" for m in df_ok["mes_final"].unique()])
            st.markdown(f"* **Registros conciliados (OK):** {len(df_ok)} ({refs_ok})")
        else:
            st.markdown(f"* **Registros conciliados (OK):** 0")

        if len(df_inc) > 0:
            detalhes_inc = []
            for _, r in df_inc.iterrows():
                detalhes_inc.append(f"Referência {r['mes_final']} com divergência de valor")
            st.markdown(f"* **Inconsistências encontradas:** {len(df_inc)} ({', '.join(detalhes_inc)})")
        else:
            st.markdown(f"* **Inconsistências encontradas:** 0")

        st.subheader("Resumo do Cruzamento:")

        for _, row in df_final.iterrows():
            mes = row["mes_final"]
            nome = row["nome_final"]
            v_pdf = formatar_moeda(row["valor_pdf"])
            v_excel = formatar_moeda(row["valor_excel_orig"])
            dif = formatar_moeda(row["diferenca"])

            if row["status"] == "OK":
                status_str = "*(OK)*"
            else:
                status_str = "*(Divergência de Valor)*"

            st.markdown(
                f"* **{mes}**: {nome} — PDF: **{v_pdf}** | Excel: **{v_excel}** | Diferença: **{dif}** {status_str}"
            )

        csv_data = df_final.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="📥 Baixar Relatório Completo (.csv)",
            data=csv_data,
            file_name="resultado_auditoria.csv",
            mime="text/csv"
        )
