import streamlit as st
import pandas as pd
from docx import Document
import io
import os

st.set_page_config(page_title="Gerador de Relatórios de Fiscalização", layout="wide")

# --- FUNÇÕES DE NEGÓCIO ---

def determinar_jurisdicao(bacia):
    if bacia == "Bacia de Santos": return "-SP"
    if bacia == "Bacia de Campos": return "-RJ"
    if bacia == "Bacia do Espírito Santo": return "-ES"
    return ""

def processar_grandeza(grandeza):
    if grandeza == "Potencial":
        return "quando as consequências não são evidentes", "5"
    elif grandeza == "Reduzida":
        return "quando os danos ambientais são locais ou temporários", "15"
    elif grandeza == "Fraca":
        return "quando os danos ambientais são de pequena proporção ou de baixa complexidade, gravidade ou magnitude, diante do contexto considerado", "30"
    elif grandeza == "Moderada":
        return "quando os danos ambientais são de proporção intermediária ou de moderada complexidade, gravidade ou magnitude, diante do contexto considerado", "50"
    elif grandeza == "Grave":
        return "quando os danos ambientais são de grande proporção ou de alta complexidade, gravidade ou magnitude, diante do contexto considerado", "70"
    return "", ""

def processar_nivel(nivel):
    niveis = {
        "A": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 150 mil a 10 milhões de reais (Mínimo + 0,3% a 20% do teto)",
        "B": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 5 milhões a 15 milhões de reais (Mínimo + 10% a 30% do teto)",
        "C": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 15,5 milhões a 25 milhões de reais (Mínimo + 31% a 50% do teto)",
        "D": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 25,5 milhões a 37,5 milhões de reais (Mínimo + 51% a 75% do teto)",
        "E": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 38 milhões a 50 milhões de reais (Mínimo + 76% a 100% do teto)"
    }
    return niveis.get(nivel, "")

def formatar_decimal(valor):
    try:
        v = str(valor).replace(",", ".")
        f = float(v)
        return f"{f:g}".replace(".", ",")
    except ValueError:
        return str(valor)

def selecionar_modelo(row):
    class_ol = str(row['class_ol']).strip()
    class_risco = str(row['class_risco']).strip()
    gerar_rel = str(row['gerar_rel']).strip()
    
    try:
        vol_num = float(str(row['vol_char']).replace(',', '.'))
    except (ValueError, TypeError):
        vol_num = 0.0

    if gerar_rel != "Sim": return None
    if class_ol == "Não Classificado" or class_risco == "OS (Não Classificado)": return None

    if class_ol == "Oleoso" and class_risco == "A": return "Rel_Fisc_Oleoso_A.docx"
    if class_ol == "Oleoso" and class_risco == "B": return "Rel_Fisc_Oleoso_B.docx"
    if class_ol == "Oleoso" and class_risco == "D": return "Rel_Fisc_Oleoso_D.docx"
    
    if class_ol == "Não Oleoso":
        if class_risco == "A":
            return "Rel_Fisc_Nao_Oleoso_Art_61_A.docx" if vol_num > 8 else "Rel_Fisc_Nao_Oleoso_Art_62_A.docx"
        if class_risco == "B":
            return "Rel_Fisc_Nao_Oleoso_Art_61_B.docx" if vol_num > 200 else "Rel_Fisc_Nao_Oleoso_Art_62_B.docx"
        if class_risco == "D":
            return "Rel_Fisc_Nao_Oleoso_Art_62_D.docx"
    
    return None

def preencher_documento(caminho_modelo, dicionario_dados):
    doc = Document(caminho_modelo)
    
    for p in doc.paragraphs:
        for chave, valor in dicionario_dados.items():
            if chave in p.text:
                p.text = p.text.replace(chave, str(valor))
                
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for chave, valor in dicionario_dados.items():
                        if chave in p.text:
                            p.text = p.text.replace(chave, str(valor))
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFACE STREAMLIT ---

st.title("⚖️ Força Tarefa - Geração de Autos e Relatórios")

st.markdown("""
Nesta versão, faça o upload da planilha Excel de controle. 
No futuro, este sistema será conectado diretamente ao SharePoint.
""")

arquivo_excel = st.file_uploader("Envie a planilha de controle (.xlsm ou .xlsx)", type=["xlsx", "xlsm"])

if arquivo_excel:
    df = pd.read_excel(arquivo_excel, sheet_name="Processos_FT", header=0)
    
    # --- ATUALIZAÇÃO DOS ÍNDICES DAS COLUNAS ---
    # A coluna 1 virou ID (índice 0). A coluna 2 foi excluída.
    # Todos os índices após o 0 caíram 1 posição em relação à versão anterior.
    colunas_map = {
        0: 'num_doc',      # Antiga coluna 1 (agora "ID")
        1: 'processo_sei', # Era 2, caiu para 1
        2: 'gerar_rel',    # Era 3, caiu para 2
        3: 'siema',        # Era 4, caiu para 3
        6: 'situacao',     # Era 7, caiu para 6
        8: 'laudo_sei',    # Era 9, caiu para 8
        10: 'data_acid',   # Era 11, caiu para 10
        11: 'relat_sei',   # Era 12, caiu para 11
        13: 'instalacao',  # Era 14, caiu para 13
        14: 'campo',       # Era 15, caiu para 14
        15: 'bacia',       # Era 16, caiu para 15
        16: 'empresa',     # Era 17, caiu para 16
        17: 'cnpj',        # Era 18, caiu para 17
        18: 'produto',     # Era 19, caiu para 18
        19: 'class_ol',    # Era 20, caiu para 19
        20: 'class_risco', # Era 21, caiu para 20
        21: 'vol_char',    # Era 22, caiu para 21
        23: 'lat',         # Era 24, caiu para 23
        24: 'lon',         # Era 25, caiu para 24
        28: 'grandeza',    # Era 29, caiu para 28
        31: 'nivel_pontos',# Era 32, caiu para 31
        32: 'nivel',       # Era 33, caiu para 32
        33: 'auto',        # Era 34, caiu para 33
        34: 'multa_char',  # Era 35, caiu para 34
        35: 'data_ai'      # Era 36, caiu para 35
    }
    
    nomes_colunas = df.columns.tolist()
    for indice, nome_novo in colunas_map.items():
        if indice < len(nomes_colunas):
            nomes_colunas[indice] = nome_novo
    df.columns = nomes_colunas
    
    st.subheader("Processos Pendentes")
    
    df_filtrado = df.dropna(subset=['siema'])
    df_filtrado = df_filtrado[
        (df_filtrado['laudo_sei'].notna()) & 
        (df_filtrado['laudo_sei'].astype(str) != '0') &
        (df_filtrado['gerar_rel'] == 'Sim')
    ]
    
    if df_filtrado.empty:
        st.success("Não há processos na fila de geração no momento!")
    else:
        colunas_resumo = ['num_doc', 'siema', 'processo_sei', 'empresa', 'bacia']
        st.dataframe(df_filtrado[colunas_resumo])
        
        st.write("---")
        st.subheader("Gerar Relatório")
        
        opcoes_processos = df_filtrado['siema'].astype(str) + " - " + df_filtrado['empresa'].astype(str)
        processo_selecionado = st.selectbox("Selecione o processo para gerar o relatório:", opcoes_processos)
        
        if st.button("Gerar Documento Word"):
            siema_alvo = processo_selecionado.split(" - ")[0]
            row = df_filtrado[df_filtrado['siema'].astype(str) == siema_alvo].iloc[0]
            
            modelo_arquivo = selecionar_modelo(row)
            
            if not modelo_arquivo:
                st.error("Não foi possível determinar o modelo para este processo. Verifique as classes de risco e tipo na planilha.")
            else:
                caminho_modelo = os.path.join("modelos", modelo_arquivo)
                
                if not os.path.exists(caminho_modelo):
                    st.error(f"O modelo '{modelo_arquivo}' não foi encontrado na pasta 'modelos/'.")
                else:
                    grandeza_texto, grandeza_pontos = processar_grandeza(str(row.get('grandeza', '')))
                    
                    dados_replace = {
                        "<<siema>>": str(row.get('siema', '')),
                        "<<processo_sei>>": str(row.get('processo_sei', '')),
                        "<<laudo_sei>>": str(row.get('laudo_sei', '')),
                        "<<data_acid>>": str(row.get('data_acid', '')),
                        "<<relat_sei>>": str(row.get('relat_sei', '')),
                        "<<instalacao>>": str(row.get('instalacao', '')),
                        "<<campo>>": str(row.get('campo', '')),
                        "<<bacia>>": str(row.get('bacia', '')),
                        "<<empresa>>": str(row.get('empresa', '')),
                        "<<cnpj>>": str(row.get('cnpj', '')),
                        "<<produto>>": str(row.get('produto', '')),
                        "<<class_ol>>": str(row.get('class_ol', '')),
                        "<<class_risco>>": str(row.get('class_risco', '')),
                        "<<vol_char>>": formatar_decimal(row.get('vol_char', '')),
                        "<<auto>>": str(row.get('auto', '')),
                        "<<multa_num>>": str(row.get('multa_char', '')),
                        "<<multa_char>>": str(row.get('multa_char', '')),
                        "<<data_ai>>": str(row.get('data_ai', '')),
                        "<<jurisdicao>>": determinar_jurisdicao(str(row.get('bacia', ''))),
                        "<<lat_auto>>": str(row.get('lat', '')),
                        "<<lon_auto>>": str(row.get('lon', '')),
                        "<<grandeza>>": str(row.get('grandeza', '')),
                        "<<grandeza_texto>>": grandeza_texto,
                        "<<grandeza_pontos>>": grandeza_pontos,
                        "<<nivel>>": str(row.get('nivel', '')),
                        "<<nivel_pontos>>": str(row.get('nivel_pontos', '')),
                        "<<nivel_texto>>": processar_nivel(str(row.get('nivel', '')))
                    }
                    
                    for k, v in dados_replace.items():
                        if v == "nan": dados_replace[k] = ""
                    
                    doc_pronto_io = preencher_documento(caminho_modelo, dados_replace)
                    
                    nome_arquivo_saida = f"Rel_Fisc_{row.get('num_doc', 'X')}_{row.get('siema', '')}.docx"
                    
                    st.success("Relatório gerado com sucesso!")
                    
                    st.download_button(
                        label="📥 Baixar Relatório Word",
                        data=doc_pronto_io,
                        file_name=nome_arquivo_saida,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
