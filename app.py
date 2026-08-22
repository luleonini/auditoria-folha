import streamlit as st
import pandas as pd
import json
from google import genai
from google.genai import types
from pydantic import BaseModel

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

# Configuração da API Key do Gemini no menu lateral
st.sidebar.header("Configurações")
api_key = st.sidebar.text_input("Cole sua Gemini API Key:", type="password")

# Estrutura esperada de retorno da IA
class DadosHolerite(BaseModel):
    nome: str
    valor_liquido: float
    mes_referencia: str

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

if st.button("🚀 Processar Auditoria", type="primary"):
    if not api_key:
        st.error("⚠️ Insira uma chave de API do Gemini na barra lateral para prosseguir.")
    elif not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Envie os arquivos PDF e a planilha Excel antes de processar.")
    else:
        client = genai.Client(api_key=api_key)
        dados_pdfs = []

        with st.spinner("Analisando holerites com IA..."):
            for pdf_file in uploaded_pdfs:
                pdf_bytes = pdf_file.read()

                prompt = (
                    "Analise este holerite. Extraia o nome completo do colaborador, "
                    "o valor líquido a receber (como número float) e o mês de referência no formato MM/AAAA."
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
                    nome_limpo = dados_json["nome"].strip().upper()
                    mes_ref = dados_json["mes_referencia"].strip()
                    
                    dados_pdfs.append({
                        "nome_pdf": dados_json["nome"].strip(),
                        "nome_clean": nome_limpo,
                        "valor_pdf": float(dados_json["valor_liquido"]),
                        "mes_ref": mes_ref,
                        "chave_cruzamento": f"{nome_limpo}_{mes_ref}"
                    })
                except Exception as e:
                    st.error(f"Erro ao ler o arquivo {pdf_file.name}: {e}")

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
        
        # Função para limpar e converter moeda do Excel
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

        # Tratamentos pós-merge
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
            if pd.isna(row.get("nome_pdf")):
                return "Faltando PDF"
            if pd.isna(row.get("nome_excel_orig")):
                return "Não encontrado no Excel"
            if row["diferenca"] > 0.01:
                return "Divergência de Valor"
            return "OK"

        df_final["status"] = df_final.apply(checar_status, axis=1)

        # Totais para os marcadores superiores
        total_regs = len(df_final)
        df_ok = df_final[df_final["status"] == "OK"]
        df_inc = df_final[df_final["status"] != "OK"]

        # Construção da lista idêntica à Imagem 1
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

        # Impressão item por item no formato da Imagem 1
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

        # Botão de download
        csv_data = df_final.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label="📥 Baixar Relatório Completo (.csv)",
            data=csv_data,
            file_name="resultado_auditoria.csv",
            mime="text/csv"
        )
