import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os
from pdf2image import convert_from_bytes
import pytesseract
from PIL import Image

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI (OCR)",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def carregar_macro(nome_arquivo):
    try:
        with open(nome_arquivo, "r", encoding="utf-8") as f: return f.read()
    except:
        try:
            with open(nome_arquivo, "r", encoding="latin-1") as f: return f.read()
        except: return "Erro: Arquivo não encontrado."

# === FUNÇÃO DE INTELIGÊNCIA HÍBRIDA (TEXTO + OCR) ===
def extrair_dados_pdf_hibrido(arquivo_pdf):
    dados_extraidos = []
    usou_ocr = False
    
    # Lê o arquivo para memória
    pdf_bytes = arquivo_pdf.read()
    
    # TENTATIVA 1: Leitura Direta (Com Layout Físico)
    # Isso resolve o caso do arquivo 153272.pdf onde o texto existe mas está desalinhado
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texto_total = ""
        for page in pdf.pages:
            # layout=True força o robô a ler linha por linha visualmente, não por coluna
            texto_pagina = page.extract_text(layout=True) 
            if texto_pagina:
                texto_total += "\n" + texto_pagina

    # Se leu muito pouco texto (menos de 50 caracteres), assume que é IMAGEM
    if len(texto_total) < 50:
        usou_ocr = True
        try:
            # Converte PDF em Imagens
            images = convert_from_bytes(pdf_bytes)
            texto_total = ""
            for img in images:
                # Usa Tesseract para ler a imagem (OCR)
                texto_pagina = pytesseract.image_to_string(img, lang='por')
                texto_total += "\n" + texto_pagina
        except Exception as e:
            return [], f"Erro no OCR: {str(e)}"

    # PROCESSAMENTO DO TEXTO (COMUM AOS DOIS MÉTODOS)
    if not texto_total: return [], "Vazio"

    # Filtros de cabeçalho
    linhas_validas = []
    for line in texto_total.split('\n'):
        if "SINTÉTICO PATRIMONIAL" in line.upper(): continue
        if "DE ENTRADAS" in line.upper() or "DE SAÍDAS" in line.upper(): continue
        linhas_validas.append(line)

    # Regex poderoso para capturar dados mesmo com sujeira
    # Procura: Começo de linha -> Número (Chave) -> Espaço -> Descrição -> Valores
    for line in linhas_validas:
        line = line.strip()
        # Regex ajustado: Pega numero no inicio, ignora chars estranhos, pega descrição e valores
        match = re.search(r'^"?(\d+)"?\s+(.+?)(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})', line)
        
        if match:
            # Tenta encontrar TODOS os valores monetários na linha
            vals = re.findall(r'(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})', line)
            
            # Precisamos do Saldo Atual.
            # Em PDFs normais é a coluna 6. Em OCR pode variar, mas geralmente é o último ou penúltimo saldo grande.
            # Vamos manter a lógica da 6ª coluna se tiver muitos valores, ou pegar o último se tiver poucos.
            if len(vals) >= 3:
                # Se tiver muitos valores, tentamos pegar o índice 5 (6ª coluna) ou o último
                valor_candidato = vals[5] if len(vals) >= 6 else vals[-1]
                
                chave = int(match.group(1))
                # Limpeza da descrição (remove aspas e sujeira do OCR)
                desc = re.sub(r'[\d.,]+$', '', match.group(2)).strip() 
                
                dados_extraidos.append({
                    'Chave_Vinculo': chave,
                    'Descricao': desc,
                    'Saldo_PDF': valor_candidato
                })

    return dados_extraidos, "OCR Ativado" if usou_ocr else "Leitura Direta"

# ==========================================
# INTERFACE
# ==========================================
st.title("👁️ Auditor Patrimonial (com Leitor Óptico)")
st.markdown("---")

with st.expander("📘 GUIA DE USO E MACROS", expanded=False):
    st.markdown("### 🚀 Passo a Passo")
    col1, col2 = st.columns(2)
    with col1:
        st.info("1. Preparação (Excel)")
        macro1 = carregar_macro("macro_preparar.txt")
        st.download_button("📥 Baixar Macro 1", macro1, "Macro_1.txt")
        macro2 = carregar_macro("macro_dividir.txt")
        st.download_button("📥 Baixar Macro 2", macro2, "Macro_2.txt")
    with col2:
        st.success("2. Auditoria (Aqui)")
        st.write("Agora suporta PDFs escaneados (imagens) e PDFs com formatação quebrada!")

st.subheader("📂 Área de Arquivos")
uploaded_files = st.file_uploader(
    "Arraste PDFs (RMB) e Excels (SIAFI)", 
    accept_multiple_files=True
)

if st.button("▶️ Iniciar Auditoria", use_container_width=True, type="primary"):
    if not uploaded_files:
        st.warning("⚠️ Adicione arquivos.")
    else:
        progresso = st.progress(0)
        logs = []
        
        pdfs = {f.name: f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excels = {f.name: f for f in uploaded_files if (f.name.lower().endswith('.xlsx') or f.name.lower().endswith('.csv'))}
        
        pares = []
        for name_ex, file_ex in excels.items():
            match = re.match(r'^(\d+)', name_ex)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match:
                    pares.append({'ug': ug, 'excel': file_ex, 'pdf': pdf_match})
                else:
                    logs.append(f"⚠️ UG {ug}: Falta PDF.")
        
        if not pares:
            st.error("❌ Nenhum par encontrado.")
        else:
            # Funções de Limpeza
            def limpar_valor(v):
                if v is None or pd.isna(v): return 0.0
                if isinstance(v, (int, float)): return float(v)
                v = str(v).replace('"', '').replace("'", "").strip()
                if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
                elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
                try: return float(re.sub(r'[^\d.-]', '', v))
                except: return 0.0

            def limpar_codigo(v):
                try: 
                    s = str(v).strip()
                    if s.endswith('.0'): s = s[:-2]
                    return s
                except: return ""

            def extrair_chave(v):
                try: return int(v[-2:])
                except: return 0

            # PDF Report Class
            class PDF_Report(FPDF):
                def header(self):
                    self.set_font('helvetica', 'B', 12)
                    self.cell(0, 10, 'Relatório de Auditoria Patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    self.ln(5)
                def footer(self):
                    self.set_y(-15); self.set_font('helvetica', 'I', 8)
                    self.cell(0, 10, f'Pg {self.page_no()}', align='C')

            pdf_out = PDF_Report()
            pdf_out.add_page()
            
            st.subheader("🔍 Resultados")

            for idx, par in enumerate(pares):
                ug = par['ug']
                
                with st.container():
                    # === 1. PROCESSAR PDF (AGORA COM OCR) ===
                    # Reinicia ponteiro do PDF para garantir leitura do zero
                    par['pdf'].seek(0)
                    dados_pdf_raw, metodo_leitura = extrair_dados_pdf_hibrido(par['pdf'])
                    
                    # Log visual do método usado
                    if metodo_leitura == "OCR Ativado":
                        st.caption(f"📸 PDF da UG {ug} processado como IMAGEM (OCR).")
                    
                    df_pdf_final = pd.DataFrame()
                    if dados_pdf_raw:
                        df_temp = pd.DataFrame(dados_pdf_raw)
                        df_temp['Saldo_PDF'] = df_temp['Saldo_PDF'].apply(limpar_valor)
                        df_pdf_final = df_temp.groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()

                    # === 2. PROCESSAR EXCEL ===
                    df_padrao = pd.DataFrame()
                    saldo_2042 = 0.0
                    tem_2042 = False
                    
                    try:
                        par['excel'].seek(0)
                        try: df = pd.read_csv(par['excel'], header=None, encoding='latin1', sep=',', engine='python')
                        except: df = pd.read_excel(par['excel'], header=None)
                        
                        if len(df.columns) >= 5:
                            df['Cod'] = df.iloc[:, 1].apply(limpar_codigo)
                            df['Desc'] = df.iloc[:, 3].astype(str).str.strip().str.upper()
                            df['Val'] = df.iloc[:, 4].apply(limpar_valor)
                            
                            if (df['Cod'] == '2042').any():
                                saldo_2042 = df.loc[df['Cod'] == '2042', 'Val'].sum()
                                if abs(saldo_2042) > 0: tem_2042 = True
                            
                            mask = df['Cod'].str.startswith('449')
                            df_dados = df[mask].copy()
                            df_dados['Chave_Vinculo'] = df_dados['Cod'].apply(extrair_chave)
                            
                            df_padrao = df_dados.groupby('Chave_Vinculo').agg({'Val':'sum', 'Desc':'first'}).reset_index()
                            df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                    except Exception as e:
                        logs.append(f"Erro Excel {ug}: {e}")

                    # === 3. CRUZAMENTO ===
                    if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                    if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                    final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                    final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if x['Descricao_Completa'] != 0 else "ITEM SEM DESCRIÇÃO", axis=1)
                    final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                    divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                    # === VISUALIZAÇÃO ===
                    st.info(f"🏢 **UG: {ug}** ({metodo_leitura})")
                    c1, c2, c3 = st.columns(3)
                    s_pdf = final['Saldo_PDF'].sum()
                    s_ex = final['Saldo_Excel'].sum()
                    dif = s_pdf - s_ex
                    
                    c1.metric("RMB (PDF)", f"R$ {s_pdf:,.2f}")
                    c2.metric("SIAFI (Excel)", f"R$ {s_ex:,.2f}")
                    c3.metric("Diferença", f"R$ {dif:,.2f}", delta_color="inverse" if abs(dif) > 0.05 else "normal")

                    if not divergencias.empty:
                        st.warning(f"⚠️ {len(divergencias)} divergência(s).")
                        with st.expander("Ver Detalhes"):
                            st.dataframe(divergencias[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']])
                    else:
                        st.success("✅ Contas conciliadas.")
                    
                    if tem_2042: st.warning(f"ℹ️ Estoque Interno: R$ {saldo_2042:,.2f}")

                    # === PDF REPORT ===
                    pdf_out.set_font("helvetica", 'B', 11)
                    pdf_out.set_fill_color(240, 240, 240)
                    pdf_out.cell(0, 10, f"Unidade Gestora: {ug}", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
                    
                    if not divergencias.empty:
                        pdf_out.set_font("helvetica", 'B', 8)
                        pdf_out.set_fill_color(255, 200, 200)
                        pdf_out.cell(10, 8, "Item", 1, fill=True)
                        pdf_out.cell(90, 8, "Descrição", 1, fill=True)
                        pdf_out.cell(30, 8, "RMB", 1, fill=True)
                        pdf_out.cell(30, 8, "SIAFI", 1, fill=True)
                        pdf_out.cell(30, 8, "DIF", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        pdf_out.set_font("helvetica", '', 8)
                        for _, row in divergencias.iterrows():
                            pdf_out.cell(10, 7, str(int(row['Chave_Vinculo'])), 1)
                            pdf_out.cell(90, 7, str(row['Descricao'])[:55], 1)
                            pdf_out.cell(30, 7, f"{row['Saldo_PDF']:,.2f}", 1)
                            pdf_out.cell(30, 7, f"{row['Saldo_Excel']:,.2f}", 1)
                            pdf_out.set_text_color(200, 0, 0)
                            pdf_out.cell(30, 7, f"{row['Diferenca']:,.2f}", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf_out.set_text_color(0, 0, 0)
                    else:
                        pdf_out.set_font("helvetica", 'I', 9)
                        pdf_out.cell(0, 8, "Sem divergências.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    
                    if tem_2042:
                        pdf_out.ln(2); pdf_out.set_font("helvetica", 'B', 9); pdf_out.set_fill_color(255, 255, 200)
                        pdf_out.cell(100, 8, "SALDO ESTOQUE INTERNO", 1, fill=True)
                        pdf_out.cell(90, 8, f"R$ {saldo_2042:,.2f}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        pdf_out.set_text_color(0, 0, 0)
                    
                    pdf_out.ln(2); pdf_out.set_font("helvetica", 'B', 9); pdf_out.set_fill_color(220, 230, 241)
                    pdf_out.cell(100, 8, "TOTAIS", 1, fill=True)
                    pdf_out.cell(30, 8, f"{s_pdf:,.2f}", 1, fill=True)
                    pdf_out.cell(30, 8, f"{s_ex:,.2f}", 1, fill=True)
                    if abs(dif) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 8, f"{dif:,.2f}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(5)
                
                progresso.progress((idx + 1) / len(pares))

            st.success("Processamento Finalizado!")
            if logs:
                with st.expander("Logs"): 
                    for l in logs: st.write(l)
            
            try:
                pdf_bytes = bytes(pdf_out.output())
                st.download_button("📥 BAIXAR RELATÓRIO PDF", pdf_bytes, "RELATORIO_AUDITORIA.pdf", "application/pdf", type="primary", use_container_width=True)
            except Exception as e: st.error(f"Erro download: {e}")
                
