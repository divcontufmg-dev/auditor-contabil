import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS para limpar a tela
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Função para ler os arquivos de texto das macros
def carregar_macro(nome_arquivo):
    try:
        with open(nome_arquivo, "r", encoding="utf-8") as f:
            return f.read()
    except:
        # Tenta encoding latin-1 se utf-8 falhar (caso tenha acentos do windows)
        try:
            with open(nome_arquivo, "r", encoding="latin-1") as f:
                return f.read()
        except:
            return "Erro: Arquivo da macro não encontrado no repositório."

# ==========================================
# CABEÇALHO E TUTORIAL
# ==========================================
st.title("📊 Ferramenta de conciliação RMBxSIAFI")
st.markdown("---")

with st.expander("📘 GUIA DE USO E MACROS (Clique para abrir)", expanded=False):
    st.markdown("### 🚀 Passo a Passo Completo")
    
    col_tut1, col_tut2 = st.columns(2)
    
    with col_tut1:
        st.info("💻 **Fase 1: No Excel (Preparação)**")
        st.markdown("""
        O arquivo original do Tesouro precisa ser tratado antes de entrar aqui.
        
        **Passo A: Preparar**
        1. Baixe a **Macro 1 (Preparação)**.
        2. No Excel, aperte `ALT + F11`, insira um Módulo e cole.
        3. Execute para formatar a planilha.
        
        NOTA: A planilha MATRIZ deve estar aberta pra que a macro funcione
        """)
        
        # Lê o arquivo txt que você subiu no GitHub
        macro1_content = carregar_macro("macro_preparar.txt")
        st.download_button(
            label="📥 Baixar Macro 1: Preparar (.txt)",
            data=macro1_content,
            file_name="Macro_1_Preparar.txt",
            mime="text/plain"
        )
        
        st.markdown("---")
        
        st.markdown("""
        **Passo B: Dividir**
        1. Baixe a **Macro 2 (Divisão)**.
        2. Cole no Excel e execute.
        3. Isso vai gerar vários arquivos Excel (um por UG).
        """)
        
        # Lê o arquivo txt que você subiu no GitHub
        macro2_content = carregar_macro("macro_dividir.txt")
        st.download_button(
            label="📥 Baixar Macro 2: Dividir (.txt)",
            data=macro2_content,
            file_name="Macro_2_Dividir.txt",
            mime="text/plain"
        )

    with col_tut2:
        st.success("🤖 **Fase 2: Na ferramenta (Aqui)**")
        st.markdown("""
        Agora que você tem os arquivos separados:
        
        1. Gere o **Relatório em PDF** no sistema RMB (Sintético Patrimonial).

        NOTA: É necessário que o PDF esteja com caracteres selecionáveis, ou seja que seja possível copiar e colar um dado. (Por vezes o relatório é retirado como imagem, dessa forma não funcionará).
        
        2. Arraste **TODOS** os arquivos para a área abaixo:
           * Os PDFs do RMB.
           * Os Excels separados que a Macro 2 gerou.
        3. O sistema vai casar os pares (PDF + Excel) automaticamente.
        4. Clique em **Iniciar Auditoria**.
        """)

# --- ÁREA DE UPLOAD ---
st.subheader("📂 Área de Arquivos")
uploaded_files = st.file_uploader(
    "Arraste seus arquivos PDF (RMB) e Excel/CSV (SIAFI já separados) para esta área:", 
    accept_multiple_files=True,
    help="Selecione os PDFs e as Planilhas de todas as UGs."
)

# --- BOTÃO DE AÇÃO ---
if st.button("▶️ Iniciar", use_container_width=True, type="primary"):
    
    if not uploaded_files:
        st.warning("⚠️ Por favor, adicione os arquivos antes de processar.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        pdfs = {f.name: f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excels = {f.name: f for f in uploaded_files if (f.name.lower().endswith('.xlsx') or f.name.lower().endswith('.csv'))}
        
        pares = []
        logs = []

        for name_ex, file_ex in excels.items():
            match = re.match(r'^(\d+)', name_ex)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match:
                    pares.append({'ug': ug, 'excel': file_ex, 'pdf': pdf_match})
                else:
                    logs.append(f"⚠️ UG {ug}: Planilha encontrada, mas falta o PDF correspondente.")
        
        if not pares:
            st.error("❌ Nenhum par completo (Excel + PDF) foi identificado. Verifique se os nomes dos arquivos Excel começam com o número da UG.")
        else:
            # --- FUNÇÕES INTERNAS ---
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
            st.markdown("---")
            st.subheader("🔍 Resultados da Análise")

            for idx, par in enumerate(pares):
                ug = par['ug']
                status_text.text(f"Processando Unidade Gestora: {ug}...")
                
                with st.container():
                    st.info(f"🏢 **Unidade Gestora: {ug}**")
                    
                    # === LEITURA EXCEL ===
                    df_padrao = pd.DataFrame()
                    saldo_2042 = 0.0
                    tem_2042_com_saldo = False
                    
                    try:
                        par['excel'].seek(0)
                        try:
                            df = pd.read_csv(par['excel'], header=None, encoding='latin1', sep=',', engine='python')
                        except:
                            df = pd.read_excel(par['excel'], header=None)
                        
                        if len(df.columns) >= 5:
                            df['Codigo_Limpo'] = df.iloc[:, 1].apply(limpar_codigo_bruto)
                            df['Descricao_Excel'] = df.iloc[:, 3].astype(str).str.strip().str.upper()
                            df['Valor_Limpo'] = df.iloc[:, 4].apply(limpar_valor)
                            
                            mask_2042 = df['Codigo_Limpo'] == '2042'
                            if mask_2042.any():
                                saldo_2042 = df.loc[mask_2042, 'Valor_Limpo'].sum()
                                if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                            
                            mask_padrao = df['Codigo_Limpo'].str.startswith('449')
                            df_dados = df[mask_padrao].copy()
                            df_dados['Chave_Vinculo'] = df_dados['Codigo_Limpo'].apply(extrair_chave_vinculo)
                            
                            df_padrao = df_dados.groupby('Chave_Vinculo').agg({
                                'Valor_Limpo': 'sum',
                                'Descricao_Excel': 'first'
                            }).reset_index()
                            df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                        else:
                            logs.append(f"❌ Erro Excel UG {ug}: Colunas insuficientes.")
                    except Exception as e:
                        logs.append(f"❌ Erro Leitura Excel UG {ug}: {e}")

                  # === PDF ===
                    df_pdf_final = pd.DataFrame()
                    dados_pdf = []
                    
                    try:
                        # Carrega o PDF em bytes na memória para usar em ambas as bibliotecas sem conflito
                        par['pdf'].seek(0)
                        pdf_bytes = par['pdf'].read()
                        
                        # Abre com pdfplumber usando BytesIO
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                            for page in p_doc.pages:
                                txt = page.extract_text()
                                
                                # --- MODIFICAÇÃO: OCR SE NÃO HOUVER TEXTO ---
                                if not txt or len(txt.strip()) < 10:
                                    try:
                                        # Converte APENAS a página atual para imagem (pdf2image usa índice base-1)
                                        imagens = convert_from_bytes(
                                            pdf_bytes, 
                                            first_page=page.page_number, 
                                            last_page=page.page_number
                                        )
                                        if imagens:
                                            # --psm 6 assume um bloco único de texto (bom para tabelas/listas)
                                            txt = pytesseract.image_to_string(imagens[0], lang='por', config='--psm 6')
                                    except Exception as e_ocr:
                                        # Se falhar o OCR, segue sem texto
                                        pass
                                # ---------------------------------------------

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
                        logs.append(f"❌ Erro Leitura PDF UG {ug}: {e}")

                    # === CRUZAMENTO ===
                    if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                    if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                    final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                    final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if x['Descricao_Completa'] != 0 else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
                    final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                    divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                    # === DASHBOARD ===
                    soma_pdf = final['Saldo_PDF'].sum()
                    soma_excel = final['Saldo_Excel'].sum()
                    dif_total = soma_pdf - soma_excel

                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total RMB (PDF)", f"R$ {soma_pdf:,.2f}")
                    col2.metric("Total SIAFI (Excel)", f"R$ {soma_excel:,.2f}")
                    col3.metric("Diferença", f"R$ {dif_total:,.2f}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
                    
                    if not divergencias.empty:
                        st.warning(f"⚠️ Atenção: {len(divergencias)} conta(s) com divergência.")
                        with st.expander("Ver Detalhes das Divergências"):
                            st.dataframe(divergencias[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']])
                    else:
                        st.success("✅ Tudo certo! Nenhuma divergência encontrada nas contas.")

                    if tem_2042_com_saldo:
                        st.warning(f"ℹ️ Conta de Estoque Interno tem saldo: R$ {saldo_2042:,.2f}")

                    st.markdown("---")

                   # === GERAÇÃO PDF ===
                    
                    # Função auxiliar para inverter ponto e vírgula apenas na exibição
                    def formatar_real(valor):
                        return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

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
                            
                            # APLICAÇÃO DA CORREÇÃO AQUI
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_PDF']), 1)
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_Excel']), 1)
                            
                            pdf_out.set_text_color(200, 0, 0)
                            pdf_out.cell(30, 7, formatar_real(row['Diferenca']), 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf_out.set_text_color(0, 0, 0)
                    else:
                        pdf_out.set_font("helvetica", 'I', 9)
                        pdf_out.cell(0, 8, "Nenhuma divergência encontrada entre RMB e SIAFI.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                    if tem_2042_com_saldo:
                        pdf_out.ln(2)
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.set_fill_color(255, 255, 200)
                        pdf_out.cell(100, 8, "SALDO NA CONTA DE ESTOQUE INTERNO", 1, fill=True)
                        # APLICAÇÃO DA CORREÇÃO AQUI
                        pdf_out.cell(90, 8, f"R$ {formatar_real(saldo_2042)}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        pdf_out.set_text_color(0, 0, 0)

                    pdf_out.ln(2)
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(220, 230, 241)
                    pdf_out.cell(100, 8, "TOTAIS (CONTAS PADRÃO)", 1, fill=True)
                    
                    # APLICAÇÃO DA CORREÇÃO NOS TOTAIS
                    pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True)
                    pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True)
                    
                    if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(5)

            # --- FIM ---
            status_text.text("Processamento concluído!")
            progresso.empty()
            
            if logs:
                with st.expander("⚠️ Avisos do Sistema (Arquivos não pareados)"):
                    for log in logs: st.write(log)
            
            st.markdown("### 📥 Relatório Consolidado")
            try:
                pdf_bytes = bytes(pdf_out.output())
                st.download_button(
                    label="BAIXAR RELATÓRIO COMPLETO (PDF)",
                    data=pdf_bytes,
                    file_name="RELATORIO_GERAL_AUDITORIA.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro ao gerar download: {e}")



