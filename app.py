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
from pytesseract import Output # Import necessário para rotação

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES E CLASSES (NO TOPO PARA EVITAR ERROS)
# ==========================================

def carregar_macro(nome_arquivo):
    try:
        with open(nome_arquivo, "r", encoding="utf-8") as f:
            return f.read()
    except:
        try:
            with open(nome_arquivo, "r", encoding="latin-1") as f:
                return f.read()
        except:
            return "Erro: Arquivo da macro não encontrado."

def limpar_valor(v):
    if v is None or pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).replace('"', '').replace("'", "").strip()
    # Remove pontos de milhar e troca vírgula por ponto
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

def formatar_real(valor):
    # Formata 1000.00 para 1.000,00
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de conferência patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15); self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE
# ==========================================
st.title("📊 Ferramenta de conciliação RMBxSIAFI")
st.markdown("---")

with st.expander("📘 GUIA DE USO E MACROS (Clique para abrir)", expanded=False):
    st.markdown("### 🚀 Passo a Passo Completo")
    col_tut1, col_tut2 = st.columns(2)
    with col_tut1:
        st.info("💻 **Fase 1: No Excel (Preparação)**")
        macro1_content = carregar_macro("macro_preparar.txt")
        st.download_button("📥 Baixar Macro 1: Preparar (.txt)", macro1_content, "Macro_1_Preparar.txt")
        st.markdown("---")
        macro2_content = carregar_macro("macro_dividir.txt")
        st.download_button("📥 Baixar Macro 2: Dividir (.txt)", macro2_content, "Macro_2_Dividir.txt")
    with col_tut2:
        st.success("🤖 **Fase 2: Na ferramenta (Aqui)**")
        st.markdown("1. Gere o PDF do RMB.\n2. Arraste PDFs e Excels abaixo.\n3. Clique em Iniciar.")

st.subheader("📂 Área de Arquivos")
uploaded_files = st.file_uploader(
    "Arraste seus arquivos PDF (RMB) e Excel/CSV (SIAFI já separados) para esta área:", 
    accept_multiple_files=True
)

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
            st.error("❌ Nenhum par completo (Excel + PDF) foi identificado.")
        else:
            pdf_out = PDF_Report()
            pdf_out.add_page()
            
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
                    except Exception as e:
                        logs.append(f"❌ Erro Excel UG {ug}: {e}")

                    # === LEITURA PDF (MODIFICADA: SEM OCR DESNECESSÁRIO + ROTAÇÃO) ===
                    df_pdf_final = pd.DataFrame()
                    dados_pdf = []
                    
                    try:
                        par['pdf'].seek(0)
                        pdf_bytes = par['pdf'].read()
                        
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                            for page in p_doc.pages:
                                # 1. Tenta extrair texto normal
                                txt = page.extract_text()
                                is_ocr = False
                                
                                # Verifica se o texto é útil (contém padrão monetário 0,00)
                                # Se tiver texto legível, NÃO entra no OCR
                                tem_dados_validos = False
                                if txt:
                                    if re.search(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}', txt):
                                        tem_dados_validos = True
                                
                                # 2. Só aplica OCR se NÃO tiver texto válido
                                if not txt or not tem_dados_validos or len(txt) < 50:
                                    is_ocr = True
                                    try:
                                        imagens = convert_from_bytes(
                                            pdf_bytes, 
                                            first_page=page.page_number, 
                                            last_page=page.page_number,
                                            dpi=300
                                        )
                                        if imagens:
                                            img = imagens[0]
                                            
                                            # --- ROTAÇÃO AUTOMÁTICA (NOVA FUNCIONALIDADE) ---
                                            try:
                                                osd = pytesseract.image_to_osd(img, output_type=Output.DICT)
                                                if osd['rotate'] != 0:
                                                    img = img.rotate(-osd['rotate'], expand=True)
                                            except:
                                                pass
                                            # ------------------------------------------------
                                            
                                            # OCR Padrão (Sem limpeza complexa, apenas leitura)
                                            txt = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                                    except Exception:
                                        pass

                                # 3. Processamento dos dados
                                if not txt: continue
                                if "SINTÉTICO PATRIMONIAL" not in txt.upper(): continue
                                if "DE ENTRADAS" in txt.upper() or "DE SAÍDAS" in txt.upper(): continue

                                for line in txt.split('\n'):
                                    if re.match(r'^"?\d+"?\s+', line):
                                        vals = []
                                        # Se for OCR, usa regex tolerante a espaços
                                        if is_ocr:
                                            vals_raw = re.findall(r'([\d\.\s]+,\d{2})', line)
                                            vals = [v.replace(' ', '') for v in vals_raw]
                                        # Se for texto normal, usa regex rígido
                                        else:
                                            vals = re.findall(r'([0-9]{1,3}(?:[.,][0-9]{3})*[.,]\d{2})', line)
                                        
                                        # Usa lógica de leitura reversa (segurança contra colunas vazias)
                                        if len(vals) >= 4:
                                            chave_match = re.match(r'^"?(\d+)', line)
                                            if chave_match:
                                                chave_raw = chave_match.group(1)
                                                dados_pdf.append({
                                                    'Chave_Vinculo': int(chave_raw),
                                                    'Saldo_PDF': limpar_valor(vals[-4])
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

                    # === EXIBIÇÃO ===
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
                        st.success("✅ Tudo certo! Nenhuma divergência encontrada.")

                    if tem_2042_com_saldo:
                        st.warning(f"ℹ️ Conta de Estoque Interno tem saldo: R$ {saldo_2042:,.2f}")

                    st.markdown("---")

                    # === GERAÇÃO PDF FINAL ===
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
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_PDF']), 1)
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_Excel']), 1)
                            pdf_out.set_text_color(200, 0, 0)
                            pdf_out.cell(30, 7, formatar_real(row['Diferenca']), 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf_out.set_text_color(0, 0, 0)
                    else:
                        pdf_out.set_font("helvetica", 'I', 9)
                        pdf_out.cell(0, 8, "Nenhuma divergência encontrada.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                    if tem_2042_com_saldo:
                        pdf_out.ln(2)
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.set_fill_color(255, 255, 200)
                        pdf_out.cell(100, 8, "SALDO ESTOQUE INTERNO", 1, fill=True)
                        pdf_out.cell(90, 8, f"R$ {formatar_real(saldo_2042)}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                    pdf_out.ln(2)
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(220, 230, 241)
                    pdf_out.cell(100, 8, "TOTAIS", 1, fill=True)
                    pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True)
                    pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True)
                    if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(5)
                
                progresso.progress((idx + 1) / len(pares))

            status_text.text("Processamento concluído!")
            progresso.empty()
            
            if logs:
                with st.expander("⚠️ Avisos"):
                    for log in logs: st.write(log)
            
            try:
                pdf_bytes = bytes(pdf_out.output())
                st.download_button("BAIXAR RELATÓRIO PDF", pdf_bytes, "RELATORIO_FINAL.pdf", "application/pdf", type="primary", use_container_width=True)
            except Exception as e:
                st.error(f"Erro download: {e}")

