import streamlit as st
import pandas as pd
import json
import os
from google import genai
from google.genai import types
from pydantic import BaseModel

st.set_page_config(page_title="Auditoria de Folha de Pagamento", page_icon="📊", layout="wide")

st.title("📊 Auditoria Automática de Folha de Pagamento")
st.markdown("Faça o upload dos **PDFs dos holerites** e da **planilha base Excel** para cruzar os valores e identificar divergências.")

# Tenta carregar a chave salva nos Secrets do Streamlit Cloud
api_key_secret = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else os.environ.get("GEMINI_API_KEY", "")

st.sidebar.header("Configurações")
api_key_input = st.sidebar.text_input("Gemini API Key:", value=api_key_secret, type="password")

# Schema Pydantic para forçar a IA a responder em JSON estrito
class DadosHolerite(BaseModel):
    nome: str
    valor_liquido: float
    mes_referencia: str

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

col1, col2 = st.columns(2)

with col1:
    uploaded_pdfs = st.file_uploader("1. Selecione os PDFs dos Holerites", type=["pdf"], accept_multiple_files=True)

with col2:
    uploaded_excel = st.file_uploader("2. Selecione a Planilha Base (.xlsx)", type=["xlsx"])

if st.button("🚀 Processar Auditoria", type="primary"):
    final_key = api_key_input.strip()
    
    if not final_key:
        st.error("⚠️ Insira uma chave de API do Gemini válida na barra lateral ou nos Secrets.")
    elif not uploaded_pdfs or not uploaded_excel:
        st.error("⚠️ Envie os arquivos PDF e a planilha Excel antes de processar.")
    else:
        try:
            # Inicializa o cliente Gemini com a chave limpa
            client = genai.Client(api_key=final_key)
            dados_pdfs = []

            with st.spinner("🤖 Agente de IA analisando holerites e extraindo dados..."):
                for pdf_file in uploaded_pdfs:
                    pdf_bytes = pdf_file.read()

                    prompt = (
                        "Analise este holerite. Extraia o nome completo do colaborador, "
                        "o valor líquido a receber (como número float numérico, ex: 3622.84 ou 4092.10) "
                        "e o mês de referência no formato MM/AAAA (ex: 07/2021 ou 08/2021)."
                    )

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
                    val_pdf = float(dados_json["valor_liquido"])
                    
                    dados_pdfs.append({
                        "nome_pdf": dados_json["nome"].strip(),
                        "nome_clean": nome_limpo,
                        "valor_pdf": val_pdf,
                        "mes_ref": mes_ref,
                        "chave_cruzamento": f"{nome_limpo}_{mes_ref}"
                    })

            df_pdf = pd.DataFrame(dados_pdfs)

            # Processar Excel
            df_excel_raw = pd.read_excel(uploaded_excel)
            
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

            # Cruzamento por Chave Única (Nome + Mês)
            if col_mes:
                df_excel["mes_excel"] = df_excel[col_mes].astype(str).str.strip()
                df_excel["chave_cruzamento"] = df_excel[col_nome].astype(str).str.strip().str.upper() + "_" + df_excel["mes_excel"]
                df_final = pd.merge(df_pdf, df_excel, on="chave_cruzamento", how="outer")
            else:
                df_excel["nome_clean"] = df_excel[col_nome].astype(str).str.strip().str.upper()
                df_final = pd.merge(df_pdf, df_excel, on="nome_clean", how="outer")

            # Pós-Merge e Cálculo de Diferenças
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

            # Totais
            total_holerites = len(df_final)
            df_ok = df_final[df_final["status"] == "OK"]
            df_inc = df_final[df_final["status"] != "OK"]

            # -------------------------------------------------------------
            # RENDERIZAÇÃO DO RELATÓRIO
            # -------------------------------------------------------------
            st.markdown("## Resumo da Auditoria")
            st.markdown(f"**Total de holerites auditados:** {total_holerites}")
            st.markdown(f"**Registros conciliados (OK):** {len(df_ok)}")
            st.markdown(f"**Inconsistências encontradas:** {len(df_inc)} (Divergência de Valor)")

            st.markdown("---")
            st.markdown("## Detalhamento dos Resultados")

            df_final = df_final.sort_values(by="mes_final", ascending=False)

            for _, row in df_final.iterrows():
                nome = row["nome_final"]
                mes = row["mes_final"]
                v_pdf = formatar_moeda(row["valor_pdf"])
                v_excel = formatar_moeda(row["valor_excel_orig"])
                dif = formatar_moeda(row["diferenca"])
                status = row["status"]

                if status == "OK":
                    st.markdown(f"### **{nome} (Ref: {mes}):**")
                    st.markdown(f"**Valor Líquido (PDF):** {v_pdf}  \n"
                                f"**Valor Registrado (Excel):** {v_excel}  \n"
                                f"**Diferença:** {dif}  \n"
                                f"**Status:** {status}")
                    st.markdown("")
                else:
                    conteudo_erro = f"""
                    <div style="background-color: #ffe6e6; border-left: 5px solid #ff4d4d; padding: 15px; border-radius: 5px; margin-bottom: 15px; color: #333333;">
                        <h3 style="margin-top:0; color: #cc0000;"><strong>{nome} (Ref: {mes}):</strong></h3>
                        <p style="margin: 3px 0;"><strong>Valor Líquido (PDF):</strong> {v_pdf}</p>
                        <p style="margin: 3px 0;"><strong>Valor Registrado (Excel):</strong> {v_excel}</p>
                        <p style="margin: 3px 0;"><strong>Diferença:</strong> {dif}</p>
                        <p style="margin: 3px 0;"><strong>Status:</strong> <span style="color: #cc0000; font-weight: bold;">{status}</span></p>
                    </div>
                    """
                    st.markdown(conteudo_erro, unsafe_allow_html=True)

            csv_data = df_final.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button(
                label="📥 Baixar Relatório Completo (.csv)",
                data=csv_data,
                file_name="relatorio_auditoria.csv",
                mime="text/csv"
            )

        except Exception as e:
            st.error(f"❌ Erro ao validar a chave ou processar com a IA: {e}")
