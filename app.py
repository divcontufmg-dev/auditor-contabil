import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Auditor Patrimonial", page_icon="🤖", layout="wide")

st.title("🤖 Auditor Patrimonial Automático")
st.markdown("""
**Instruções:**
1. Arraste **TODOS** os arquivos (PDFs do RMB e Excel/CSV do SIAFI) para a área abaixo.
2. O sistema vai parear automaticamente pelo código da UG.
3. Baixe o relatório final consolidado.
""")

# --- UPLOAD DE ARQUIVOS ---
uploaded_files = st.file_uploader("Solte os arquivos aqui (PDFs e Planilhas)", accept_multiple_files=True)

if st.button("Processar Conciliação"):
    if not uploaded_files:
        st.warning("Por favor, faça o upload dos arquivos primeiro.")
    else:
        st.info("Iniciando processamento...")
        
        # Separa arquivos em memória
        pdfs = {f.name: f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excels = {f.name: f for f in uploaded_files if (f.name.lower().endswith('.xlsx') or f.name.lower().endswith('.csv'))}
        
        pares = []
        logs = []

        # Pareamento
        for name_ex, file_ex in excels.items():
            match = re.match(r'^(\d+)', name_ex)
            if match:
                ug = match.group(1)
                # Procura PDF correspondente
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match:
                    pares.append({'ug': ug, 'excel': file_ex, 'pdf': pdf_match})
                else:
                    logs.append(f"⚠️ UG {ug}: Planilha encontrada, mas falta o PDF.")
        
        if not pares:
            st.error("Nenhum par (Excel + PDF) foi identificado. Verifique os nomes dos arquivos.")
        else:
            # --- FUNÇÕES AUXILIARES ---
            def limpar_valor(v):
                if v is None or pd.isna(v): return 0.0
                if isinstance(v, (int, float)): return float(v)
                v = str(v).replace('"', '').replace("'", "").strip()
                if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
                elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
                try: return float(re.sub(r'[^\d.-]', '', v))
                except: return 0.0

            def limpar_codigo_bruto(v):
                try:
                    s = str(v).strip()
                    if s.endswith('.0'): s = s[:-2]
                    return s
                except: return ""

            def extrair_chave_vinculo(codigo_str):
                try: return int(codigo_str[-2:])
                except: return 0

            # --- PREPARAÇÃO DO PDF ---
            class PDF_Report(FPDF):
                def header(self):
                    self.set_font('helvetica', 'B', 12)
                    self.cell(0, 10, 'Relatório de Auditoria Patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    self.ln(5)
                def footer(self):
                    self.set_y(-15); self.set_font('helvetica', 'I', 8)
                    self.cell(0, 10, f'Página {self.page_no()}', align='C')

            pdf_out = PDF_Report()
            pdf_out.add_page()
            
            # --- PROCESSAMENTO ---
            progress_bar = st.progress(0)
            
            for idx, par in enumerate(pares):
                ug = par['ug']
                logs.append(f"✅ Processando UG {ug}...")
                
                # === LEITURA EXCEL (SIAFI) ===
                df_padrao = pd.DataFrame()
                saldo_2042 = 0.0
                tem_2042_com_saldo = False
                
                try:
                    par['excel'].seek(0) # Reinicia ponteiro do arquivo
                    try:
                        df = pd.read_csv(par['excel'], header=None, encoding='latin1', sep=',', engine='python')
                    except:
                        df = pd.read_excel(par['excel'], header=None)
                    
                    if len(df.columns) >= 5:
                        df['Codigo_Limpo'] = df.iloc[:, 1].apply(limpar_codigo_bruto)
                        df['Descricao_Excel'] = df.iloc[:, 3].astype(str).str.strip().str.upper()
                        df['Valor_Limpo'] = df.iloc[:, 4].apply(limpar_valor)
                        
                        # Conta 2042
                        mask_2042 = df['Codigo_Limpo'] == '2042'
                        if mask_2042.any():
                            saldo_2042 = df.loc[mask_2042, 'Valor_Limpo'].sum()
                            if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                        
                        # Contas Padrão (449...)
                        mask_padrao = df['Codigo_Limpo'].str.startswith('449')
                        df_dados = df[mask_padrao].copy()
                        df_dados['Chave_Vinculo'] = df_dados['Codigo_Limpo'].apply(extrair_chave_vinculo)
                        
                        # Agrupa somando e pegando a primeira descrição
                        df_padrao = df_dados.groupby('Chave_Vinculo').agg({
                            'Valor_Limpo': 'sum',
                            'Descricao_Excel': 'first'
                        }).reset_index()
                        df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                    else:
                        logs.append(f"❌ Erro Excel UG {ug}: Arquivo com colunas insuficientes.")

                except Exception as e:
                    logs.append(f"❌ Erro Excel UG {ug}: {e}")

                # === LEITURA PDF (RMB) ===
                df_pdf_final = pd.DataFrame()
                dados_pdf = []
                try:
                    with pdfplumber.open(par['pdf']) as p_doc:
                        for page in p_doc.pages:
                            txt = page.extract_text()
                            if not txt: continue
                            if "SINTÉTICO PATRIMONIAL" not in txt.upper(): continue
                            if "DE ENTRADAS" in txt.upper() or "DE SAÍDAS" in txt.upper(): continue

                            for line in txt.split('\n'):
                                if re.match(r'^"?\d+"?\s+', line):
                                    vals = re.findall(r'([0-9]{1,3}(?:[.,][0-9]{3})*[.,]\d{2})', line)
                                    if len(vals) >= 6:
                                        chave_raw = re.match(r'^"?(\d+)', line).group(1)
                                        dados_pdf.append({
                                            'Chave_Vinculo': int(chave_raw),
                                            'Saldo_PDF': limpar_valor(vals[5])
                                        })
                    if dados_pdf:
                        df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
                except Exception as e:
                    logs.append(f"❌ Erro PDF UG {ug}: {e}")

                # === CRUZAMENTO ===
                if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if x['Descricao_Completa'] != 0 else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
                final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                # === GERAÇÃO DO PDF ===
                pdf_out.set_font("helvetica", 'B', 11)
                pdf_out.set_fill_color(240, 240, 240)
                pdf_out.cell(0, 10, text=f"Unidade Gestora: {ug}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
                
                if not divergencias.empty:
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(255, 200, 200)
                    pdf_out.cell(15, 8, "Item", 1, fill=True)
                    pdf_out.cell(85, 8, "Descrição da Conta", 1, fill=True)
                    pdf_out.cell(30, 8, "SALDO RMB", 1, fill=True)
                    pdf_out.cell(30, 8, "SALDO SIAFI", 1, fill=True)
                    pdf_out.cell(30, 8, "Diferença", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    
                    pdf_out.set_font("helvetica", '', 8)
                    for _, row in divergencias.iterrows():
                        pdf_out.cell(15, 7, str(int(row['Chave_Vinculo'])), 1)
                        pdf_out.cell(85, 7, str(row['Descricao'])[:48], 1)
                        pdf_out.cell(30, 7, f"{row['Saldo_PDF']:,.2f}", 1)
                        pdf_out.cell(30, 7, f"{row['Saldo_Excel']:,.2f}", 1)
                        pdf_out.set_text_color(200, 0, 0)
                        pdf_out.cell(30, 7, f"{row['Diferenca']:,.2f}", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        pdf_out.set_text_color(0, 0, 0)
                else:
                    pdf_out.set_font("helvetica", 'I', 9)
                    pdf_out.cell(0, 8, "Nenhuma divergência encontrada entre RMB e SIAFI.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                if tem_2042_com_saldo:
                    pdf_out.ln(2)
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(255, 255, 200)
                    pdf_out.cell(100, 8, "SALDO NA CONTA DE ESTOQUE INTERNO", 1, fill=True)
                    pdf_out.cell(90, 8, f"R$ {saldo_2042:,.2f}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)

                soma_pdf = final['Saldo_PDF'].sum()
                soma_excel = final['Saldo_Excel'].sum()
                dif_total = soma_pdf - soma_excel
                
                pdf_out.ln(2)
                pdf_out.set_font("helvetica", 'B', 9)
                pdf_out.set_fill_color(220, 230, 241)
                pdf_out.cell(100, 8, "TOTAIS (CONTAS PADRÃO)", 1, fill=True)
                pdf_out.cell(30, 8, f"{soma_pdf:,.2f}", 1, fill=True)
                pdf_out.cell(30, 8, f"{soma_excel:,.2f}", 1, fill=True)
                if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
                pdf_out.cell(30, 8, f"{dif_total:,.2f}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf_out.set_text_color(0, 0, 0)
                pdf_out.ln(5)
                
                progress_bar.progress((idx + 1) / len(pares))

            # --- EXIBIÇÃO ---
            st.success("Processamento concluído!")
            
            with st.expander("Ver Logs do Processamento"):
                for log in logs:
                    st.write(log)
            
            try:
                # CORREÇÃO APLICADA: Converte bytearray para bytes
                pdf_bytes = bytes(pdf_out.output())
                
                st.download_button(
                    label="📥 Baixar Relatório de Auditoria (PDF)",
                    data=pdf_bytes,
                    file_name="RELATORIO_GERAL_AUDITORIA.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Erro ao gerar arquivo para download: {e}")
