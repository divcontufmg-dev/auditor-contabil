import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io

# ==========================================
# ÁREA DE INSERÇÃO DOS CÓDIGOS VBA
# ==========================================

# 1. COLE AQUI O CÓDIGO QUE PREPARA A PLANILHA (Limpeza, formatação, etc)
MACRO_PREPARAR = """
Sub AutomateSpreadsheetTasks()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim foundSheet As Boolean
    Dim celula As Range
    Dim valoresParaExcluir As Variant
    Dim i As Long

    ' Desativa a atualização da tela para acelerar a execução
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Desativa o cálculo automático

    Set targetWorkbook = ThisWorkbook ' Define a planilha onde o código está rodando

    ' 1) Inserir uma nova aba chamada “MATRIZ”, capturando os dados da planilha excel aberta chamada “MATRIZ”
    On Error Resume Next ' Ignora erros se a planilha "MATRIZ" já existir
    Set sourceWorkbook = Workbooks("MATRIZ.xlsx") ' Altere para o nome exato do arquivo da planilha MATRIZ se for diferente
    On Error GoTo 0 ' Restaura o tratamento de erros

    If Not sourceWorkbook Is Nothing Then
        foundSheet = False
        For Each ws In targetWorkbook.Worksheets
            If ws.Name = "MATRIZ" Then
                foundSheet = True
                Exit For
            End If
        Next ws

        If Not foundSheet Then
            sourceWorkbook.Sheets(1).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
            targetWorkbook.ActiveSheet.Name = "MATRIZ"
            MsgBox "A aba 'MATRIZ' foi criada e populada com sucesso!", vbInformation
        Else
            MsgBox "A aba 'MATRIZ' já existe e não será recriada.", vbExclamation
        End If
    Else
        MsgBox "A planilha 'MATRIZ.xlsx' não está aberta. Certifique-se de que ela esteja aberta para copiar os dados.", vbExclamation
        GoTo EndSub ' Sai da sub se a planilha MATRIZ não for encontrada
    End If

    ' Loop através de todas as abas, exceto "MATRIZ"
    For Each ws In targetWorkbook.Worksheets
        If ws.Name <> "MATRIZ" Then

            ' 2) Inserir uma coluna em todas as abas da planilha exceto na planilha denominada “MATRIZ”
            ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

            ' 3) Inserir a formula PROCV(B;Planilha1!$A$1:$B$47;2;0) em todas as linhas da coluna A a partir da célula A8 até a última linha
            lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Encontra a última linha com dados na coluna B
            If lastRow >= 8 Then
                ws.Range("A8:A" & lastRow).Formula = "=VLOOKUP(B8,MATRIZ!$A$1:$B$47,2,FALSE)"
            End If

            ' 4) Converter todas as linhas da coluna B em número a partir de B8 até a ultima linha
            If lastRow >= 8 Then
                For Each celula In ws.Range("B8:B" & lastRow)
                    celula.Value = CDbl(celula.Value) ' Converte para Double (número)
                Next celula
            End If

          ' 5) Excluir linhas cujos valores da coluna B sejam iguais a “123110703”, “123110402” e “44905287”
valoresParaExcluir = Array("123110703", "123110402", "44905287")

For i = lastRow To 8 Step -1 ' Percorre de baixo para cima para não saltar linhas ao apagar
    If Not IsError(ws.Cells(i, "B").Value) Then
        Dim valorCelula As String
        valorCelula = Trim(CStr(ws.Cells(i, "B").Value))

        ' Verifica se o valor da célula está em qualquer uma das posições do Array
        If valorCelula = valoresParaExcluir(0) Or _
           valorCelula = valoresParaExcluir(1) Or _
           valorCelula = valoresParaExcluir(2) Then
            ws.Rows(i).Delete
        End If
    End If
Next i

            ' 6) Inserir uma célula contendo o somatório da coluna D após a ultima linha contendo valores,
            ' sendo que o formato da célula deverá apresentar casa de milhares na formatação.
            ' Inserir a palvra “TOTAL” após a ultima célula contendo valores da coluna C.
            If lastRow > 0 Then ' Garante que há linhas com dados
                ws.Cells(lastRow + 1, "D").Formula = "=SUM(D8:D" & lastRow & ")"
                ws.Cells(lastRow + 1, "D").NumberFormat = "#,##0.00" ' Formato com separador de milhares

                ' Inserir a palavra "TOTAL" na coluna C
                ws.Cells(lastRow + 1, "C").Value = "TOTAL"
            End If

            ' 7) Ajustar todo o conteúdo das células de todas as Abas
            ws.Columns.AutoFit

            ' 8) Classificar a coluna A em ordem crescente a partir de A8 até a ultima célula
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow >= 8 Then
                ws.Sort.SortFields.Clear
                ws.Sort.SortFields.Add Key:=ws.Range("A8:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ws.Sort
                    .SetRange ws.Range("A8:A" & lastRow) ' Ajuste o range de classificação para A8 até a última linha com dados na coluna A
                    .Header = xlNo ' Não há cabeçalho no range de classificação
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
            End If

' =================================================================================
            ' 9) NOVA FUNCIONALIDADE: Realçar linhas específicas de vermelho
            ' =================================================================================
            lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Garante que temos a última linha correta
            If lastRow >= 8 Then
                For i = 8 To lastRow
                    ' Condição 1: Valor na coluna B é 123110801
                    ' Condição 2: Valor na coluna D é diferente de 0 e não está vazio
                    If ws.Cells(i, "B").Value = 123110801 And ws.Cells(i, "D").Value <> 0 And Not IsEmpty(ws.Cells(i, "D").Value) Then
                        ' Pinta o fundo do intervalo de B até D de vermelho
                        ws.Range("B" & i & ":D" & i).Interior.Color = vbRed
                    End If
                    If ws.Cells(i, "B").Value = 123119905 And ws.Cells(i, "D").Value <> 0 And Not IsEmpty(ws.Cells(i, "D").Value) Then
                        ' Pinta o fundo do intervalo de B até D de vermelho
                        ws.Range("B" & i & ":D" & i).Interior.Color = vbBlue
                    End If
                Next i
            End If

        End If
    Next ws

EndSub:
    ' Reativa a atualização da tela e o cálculo automático
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' 10) Exibir ao final do trabalho a mensagem “PLANILHA DE BENS MÓVEIS ATUALIZADA COM ÊXITO!”
    MsgBox "PLANILHA DE BENS MÓVEIS ATUALIZADA COM ÊXITO!", vbInformation

End Sub
"""

# 2. COLE AQUI O CÓDIGO QUE DIVIDE EM ARQUIVOS
MACRO_DIVIDIR = """
Sub SalvarAbasComoArquivos()
    Dim ws As Worksheet
    Dim Caminho As String
    
    'Caminho onde os arquivos serão salvos (mesma pasta do arquivo original)
    Caminho = ThisWorkbook.Path & "\"
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Copy
        ActiveWorkbook.SaveAs Filename:=Caminho & ws.Name & ".xlsx"
        ActiveWorkbook.Close SaveChanges:=True
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo concluído! As abas foram salvas na mesma pasta deste arquivo."
End Sub
"""

# ==========================================
# INÍCIO DO APLICATIVO WEB
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

# --- CABEÇALHO E TUTORIAL ---
st.title("📊 Auditor Patrimonial Inteligente")
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
        """)
        st.download_button(
            label="📥 Baixar Macro 1: Preparar (.txt)",
            data=MACRO_PREPARAR,
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
        st.download_button(
            label="📥 Baixar Macro 2: Dividir (.txt)",
            data=MACRO_DIVIDIR,
            file_name="Macro_2_Dividir.txt",
            mime="text/plain"
        )

    with col_tut2:
        st.success("🤖 **Fase 2: No Auditor (Aqui)**")
        st.markdown("""
        Agora que você tem os arquivos separados:
        
        1. Gere o **Relatório em PDF** no sistema RMB (Sintético Patrimonial).
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
if st.button("▶️ Iniciar Auditoria", use_container_width=True, type="primary"):
    
    if not uploaded_files:
        st.warning("⚠️ Por favor, adicione os arquivos antes de processar.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        # Separação dos arquivos
        pdfs = {f.name: f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excels = {f.name: f for f in uploaded_files if (f.name.lower().endswith('.xlsx') or f.name.lower().endswith('.csv'))}
        
        pares = []
        logs = []

        # Pareamento
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
            st.error("❌ Nenhum par completo (Excel + PDF) foi identificado. Verifique se os arquivos Excel começam com o número da UG.")
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
                            
                            # 2042
                            mask_2042 = df['Codigo_Limpo'] == '2042'
                            if mask_2042.any():
                                saldo_2042 = df.loc[mask_2042, 'Valor_Limpo'].sum()
                                if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                            
                            # Padrão
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
                        st.success("✅ Tudo certo! Nenhuma divergência encontrada nas contas padrão.")

                    if tem_2042_com_saldo:
                        st.warning(f"ℹ️ Conta de Estoque Interno (2042) tem saldo: R$ {saldo_2042:,.2f}")

                    st.markdown("---")

                    # === GERAÇÃO PDF ===
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
                
                progresso.progress((idx + 1) / len(pares))

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
