import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import io
import os
import requests
from datetime import datetime, timedelta

st.set_page_config(page_title="Gerador de Relatórios de Fiscalização", layout="wide")

# --- FUNÇÕES AUXILIARES DE TRATAMENTO ---

def converter_data_excel(valor):
    val_str = str(valor).strip()
    if not val_str or val_str in ["nan", "None", "0"]:
        return " [ DATA - EDITAR MANUAL ] "
        
    if val_str.isdigit():
        try:
            dias = int(val_str)
            data_real = datetime(1900, 1, 1) + timedelta(days=dias - 2)
            return data_real.strftime("%d/%m/%Y")
        except:
            pass
    return val_str

def extrair_volume_numerico(valor):
    """ Utilizado para os testes de limiares dos modelos """
    if pd.isna(valor):
        return 0.0
    try:
        return float(str(valor).strip().replace(",", "."))
    except ValueError:
        return 0.0

def extrair_volume_texto(valor):
    """
    VARIÁVEL STRING CORRIGIDA: Trata o float puramente numérico (mesmo científico tipo 1e-06),
    forçando a abertura de até 7 casas decimais e aplicando a vírgula brasileira.
    """
    if pd.isna(valor):
        return " [ VOLUME - EDITAR MANUAL ] "
        
    try:
        num_float = float(str(valor).strip().replace(",", "."))
        
        if num_float == 0.0:
            return "0"
            
        # Força formatação decimal estrita com até 7 casas decimais
        texto_formatado = f"{num_float:.7f}"
        
        # Remove os zeros excedentes à direita que o Python adiciona para preencher as 7 casas
        if "." in texto_formatado:
            texto_formatado = texto_formatado.rstrip('0')
            if texto_formatado.endswith('.'):
                texto_formatado = texto_formatado[:-1]
                
        return texto_formatado.replace(".", ",")
    except ValueError:
        return str(valor).strip().replace(".", ",")

# --- FUNÇÕES DE NEGÓCIO ---

def determinar_jurisdicao(bacia):
    bacia_limpa = str(bacia).lower().strip()
    if "santos" in bacia_limpa: return "-SP"
    if "campos" in bacia_limpa: return "-RJ"
    if "espirito santo" in bacia_limpa: return "-ES"
    return ""

def processar_grandeza(grandeza):
    g = str(grandeza).strip().title()
    if g == "Potencial":
        return "quando as consequências não são evidentes", "5"
    elif g == "Reduzida":
        return "quando os danos ambientais são locais ou temporários", "15"
    elif g == "Fraca":
        return "quando os danos ambientais são de pequena proporção ou de baixa complexidade, gravidade ou magnitude, diante do contexto considerado", "30"
    elif g == "Moderada":
        return "quando os danos ambientais são de proporção intermediária ou de moderada complexidade, gravidade ou magnitude, diante do contexto considerado", "50"
    elif g == "Grave":
        return "quando os danos ambientais são de grande proporção ou de alta complexidade, gravidade ou magnitude, diante do contexto considerado", "70"
    return " [ GRANDEZA TEXTO - EDITAR MANUAL ] ", " [ PONTOS GRANDEZA - EDITAR MANUAL ] "

def processar_nivel(nivel):
    niveis = {
        "A": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 150 mil a 10 milhões de reais (Mínimo + 0,3% a 20% do teto)",
        "B": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 5 milhões a 15 milhões de reais (Mínimo + 10% a 30% do teto)",
        "C": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 15,5 milhões a 25 milhões de reais (Mínimo + 31% a 50% do teto)",
        "D": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 25,5 milhões a 37,5 milhões de reais (Mínimo + 51% a 75% do teto)",
        "E": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 38 milhões a 50 milhões de reais (Mínimo + 76% a 100% do teto)"
    }
    return niveis.get(str(nivel).strip().upper(), " [ NÍVEL TEXTO - EDITAR MANUAL ] ")

def extrair_classe_e_modelo(row):
    class_ol = str(row.get('class_ol', '')).strip().title()
    class_risco_bruto = str(row.get('class_risco', '')).strip().upper()
    
    # Processa o volume puramente numérico
    vol_num = extrair_volume_numerico(row.get('vol_char', '0'))

    letra_risco = "A"
    if "B" in class_risco_bruto: letra_risco = "B"
    elif "C" in class_risco_bruto: letra_risco = "C"
    elif "D" in class_risco_bruto: letra_risco = "D"
    elif "E" in class_risco_bruto: letra_risco = "E"

    if "Oleoso" in class_ol and "Não" not in class_ol:
        if letra_risco in ["A", "B", "D"]:
            return f"Rel_Fisc_Oleoso_{letra_risco}.docx", letra_risco
        return "Rel_Fisc_Oleoso_A.docx", "A" 
    
    if "Não Oleoso" in class_ol or "Nao Oleoso" in class_ol:
        if letra_risco == "A":
            return ("Rel_Fisc_Nao_Oleoso_Art_61_A.docx" if vol_num > 8 else "Rel_Fisc_Nao_Oleoso_Art_62_A.docx"), "A"
        if letra_risco == "B":
            return ("Rel_Fisc_Nao_Oleoso_Art_61_B.docx" if vol_num > 200 else "Rel_Fisc_Nao_Oleoso_Art_62_B.docx"), "B"
        if letra_risco == "D":
            return "Rel_Fisc_Nao_Oleoso_Art_62_D.docx", "D"
        return "Rel_Fisc_Nao_Oleoso_Art_62_A.docx", "A"

    return None, None

def preencher_documento(caminho_modelo, dicionario_dados):
    doc = Document(caminho_modelo)
    
    # 1. Substituição e Realce Total em Parágrafos normais
    for p in doc.paragraphs:
        for chave, valor in dicionario_dados.items():
            if chave in p.text:
                p.text = p.text.replace(chave, str(valor))
        
        # Se contiver marcadores de erro, destaca todo o parágrafo de forma estável
        if "[" in p.text and "]" in p.text:
            for run in p.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                
    # 2. Substituição e Realce Total em Tabelas estruturadas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for chave, valor in dicionario_dados.items():
                        if chave in p.text:
                            p.text = p.text.replace(chave, str(valor))
                    
                    if "[" in p.text and "]" in p.text:
                        for run in p.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFACE STREAMLIT ---

st.title("⚖️ Força Tarefa - Geração de Autos e Relatórios")

@st.cache_data(ttl=300) 
def carregar_dados_sharepoint():
    try:
        url = st.secrets["sharepoint"]["url_planilha"]
        headers = {"Content-Type": "application/json"}
        resposta = requests.post(url, headers=headers, json={})
        resposta.raise_for_status()
        dados_json = resposta.json()
        
        if isinstance(dados_json, dict) and "value" in dados_json:
            lista_registros = dados_json["value"]
        elif isinstance(dados_json, list):
            lista_registros = dados_json
        else:
            st.error("O formato retornado pelo Power Automate não é válido.")
            return None
        return pd.DataFrame(lista_registros)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return None

df_original = carregar_dados_sharepoint()

if df_original is not None and not df_original.empty:
    df = df_original.copy()
    
    # Mapeamento estrito associando as colunas oficiais informadas
    colunas_map = {
        'ID': 'num_doc', 'PROCESSO': 'processo_sei', 'SIEMA': 'siema', 'Situação': 'situacao',
        'Laudo Válido (SEI)': 'laudo_sei', 'DATA ACIDENTE': 'data_acid', 'RAIPO_SEI': 'relat_sei',
        'INSTALAÇÃO': 'instalacao', 'Campo': 'campo', 'Bacia': 'bacia', 'EMPRESA': 'empresa',
        'CNPJ': 'cnpj', 'PRODUTO': 'produto', 'CLASS_OL': 'class_ol', 'CLASS. RISCO': 'class_risco',
        'VOL.': 'vol_char', 'Lat_Auto': 'lat', 'Lon_Auto': 'lon', 'Grandeza': 'grandeza',
        'Nivel_Pontos': 'nivel_pontos', 'Nivel': 'nivel', 'Auto Infração': 'auto',
        'Multa Aplicada': 'multa_char', 'Data_AI': 'data_ai'
    }
    
    for col_real, col_interna in colunas_map.items():
        if col_real in df.columns:
            df[col_interna] = df[col_real]
        else:
            df[col_interna] = ""

    df_filtrado = df.reset_index(drop=True)
    
    st.subheader("Fila de Processos Disponíveis (SharePoint)")
    
    df_exibicao = pd.DataFrame({
        "Selecionar": [False] * len(df_filtrado),
        "ID": df_filtrado['num_doc'].astype(str),
        "SIEMA": df_filtrado['siema'].astype(str),
        "Processo SEI": df_filtrado['processo_sei'].astype(str),
        "Empresa": df_filtrado['empresa'].astype(str),
        "Bacia": df_filtrado['bacia'].astype(str)
    })
    
    tabela_editada = st.data_editor(
        df_exibicao,
        hide_index=True,
        disabled=["ID", "SIEMA", "Processo SEI", "Empresa", "Bacia"],
        column_config={
            "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False)
        }
    )
    
    indices_selecionados = tabela_editada[tabela_editada["Selecionar"] == True].index
    df_selecionados = df_filtrado.iloc[indices_selecionados]
    
    st.write("---")
    st.subheader("Ações de Geração")
    
    if df_selecionados.empty:
        st.info("💡 Marque os processos desejados na coluna **'Selecionar'** para liberar a geração.")
    else:
        st.warning(f"Você selecionou **{len(df_selecionados)}** processo(s).")
        
        if st.button("🚀 Gerar Relatórios dos Processos Selecionados"):
            for _, row in df_selecionados.iterrows():
                modelo_arquivo, letra_risco_detectada = extrair_classe_e_modelo(row)
                
                if not modelo_arquivo:
                    st.error(f"❌ Não foi possível determinar o modelo para o ID {row['num_doc']}. Certifique-se de que a coluna 'CLASS_OL' indique se é Oleoso ou Não Oleoso.")
                    continue
                    
                caminho_modelo = os.path.join("modelos", modelo_arquivo)
                
                if not os.path.exists(caminho_modelo):
                    st.error(f"Arquivo '{modelo_arquivo}' ausente na pasta 'modelos/'.")
                    continue
                
                def tratar_tag(valor, nome_tag):
                    v_str = str(valor).strip()
                    if pd.isna(valor) or v_str in ["", "nan", "None", "0", "Processo Não Encontrado"]:
                        return f" [ {nome_tag.upper()} - EDITAR MANUAL ] "
                    if nome_tag == "siema" and "fora do ar" in v_str.lower():
                        return " [ SIEMA FORA DO AR - EDITAR MANUAL ] "
                    return v_str

                grandeza_texto, grandeza_pontos = processar_grandeza(row.get('grandeza', ''))
                nivel_texto = processar_nivel(row.get('nivel', ''))

                risco_final = letra_risco_detectada if letra_risco_detectada else tratar_tag(row.get('class_risco', ''), "class_risco")

                dados_replace = {
                    "<<siema>>": tratar_tag(row.get('siema', ''), "siema"),
                    "<<processo_sei>>": tratar_tag(row.get('processo_sei', ''), "processo_sei"),
                    "<<laudo_sei>>": tratar_tag(row.get('laudo_sei', ''), "laudo_sei").split('.')[0],
                    "<<data_acid>>": converter_data_excel(row.get('data_acid', '')),
                    "<<relat_sei>>": tratar_tag(row.get('relat_sei', ''), "raipo_sei"),
                    "<<instalacao>>": tratar_tag(row.get('instalacao', ''), "instalacao"),
                    "<<campo>>": tratar_tag(row.get('campo', ''), "campo"),
                    "<<bacia>>": tratar_tag(row.get('bacia', ''), "bacia"),
                    "<<empresa>>": tratar_tag(row.get('empresa', ''), "empresa"),
                    "<<cnpj>>": tratar_tag(row.get('cnpj', ''), "cnpj"),
                    "<<produto>>": tratar_tag(row.get('produto', ''), "produto"),
                    "<<class_ol>>": tratar_tag(row.get('class_ol', ''), "class_ol"),
                    "<<class_risco>>": risco_final,
                    "<<vol_char>>": extrair_volume_texto(row.get('vol_char', '')),
                    "<<auto>>": tratar_tag(row.get('auto', ''), "auto_infracao"),
                    "<<multa_num>>": tratar_tag(row.get('multa_char', ''), "multa_aplicada"),
                    "<<multa_char>>": tratar_tag(row.get('multa_char', ''), "multa_aplicada"),
                    "<<data_ai>>": converter_data_excel(row.get('data_ai', '')),
                    "<<jurisdicao>>": determinar_jurisdicao(str(row.get('bacia', ''))),
                    "<<lat_auto>>": tratar_tag(row.get('lat', ''), "lat"),
                    "<<lon_auto>>": tratar_tag(row.get('lon', ''), "lon"),
                    "<<grandeza>>": tratar_tag(row.get('grandeza', ''), "grandeza"),
                    "<<grandeza_texto>>": grandeza_texto,
                    "<<grandeza_pontos>>": grandeza_pontos,
                    "<<nivel>>": tratar_tag(row.get('nivel', ''), "nivel"),
                    "<<nivel_pontos>>": tratar_tag(row.get('nivel_pontos', ''), "nivel_pontos"),
                    "<<nivel_texto>>": nivel_texto
                }
                
                doc_pronto_io = preencher_documento(caminho_modelo, dados_replace)
                nome_arquivo_saida = f"Rel_Fisc_{row.get('num_doc', 'X')}_{row.get('siema', 'Revisao')}.docx"
                
                with st.container(border=True):
                    st.write(f"📄 **ID:** {row['num_doc']} | **Processo:** {row['processo_sei']} | **Empresa:** {row['empresa']}")
                    st.download_button(
                        label=f"📥 Baixar Documento Word (ID {row['num_doc']})",
                        data=doc_pronto_io,
                        file_name=nome_arquivo_saida,
                        key=f"dl_{row['num_doc']}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
else:
    st.info("Aguardando carregamento e resposta válida do Power Automate.")
