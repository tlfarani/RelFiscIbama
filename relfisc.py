import streamlit as st
import pandas as pd
from docx import Document
import io
import os
import requests
import base64

st.set_page_config(page_title="Gerador de Relatórios de Fiscalização", layout="wide")

# --- FUNÇÕES DE NEGÓCIO ---

def determinar_jurisdicao(bacia):
    bacia_limpa = str(bacia).lower().strip()
    if "santos" in bacia_limpa: return "-SP"
    if "campos" in bacia_limpa: return "-RJ"
    if "espirito santo" in bacia_limpa: return "-ES"
    return ""

def processar_grandeza(grandeza):
    g = str(grandeza).strip()
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
    return "", ""

def processar_nivel(nivel):
    niveis = {
        "A": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 150 mil a 10 milhões de reais (Mínimo + 0,3% a 20% do teto)",
        "B": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 5 milhões a 15 milhões de reais (Mínimo + 10% a 30% do teto)",
        "C": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 15,5 milhões a 25 milhões de reais (Mínimo + 31% a 50% do teto)",
        "D": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 25,5 milhões a 37,5 milhões de reais (Mínimo + 51% a 75% do teto)",
        "E": "Como o incidente envolveu uma empresa de grande porte, a multa irá variar de, aproximadamente, 38 milhões a 50 milhões de reais (Mínimo + 76% a 100% do teto)"
    }
    return niveis.get(str(nivel).strip().upper(), "")

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
    
    try:
        vol_num = float(str(row['vol_char']).replace(',', '.'))
    except (ValueError, TypeError):
        vol_num = 0.0

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

# --- CARREGAMENTO AUTOMÁTICO VIA POWER AUTOMATE (SECRETS) ---
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
            st.error("O formato retornado pelo Power Automate não é uma lista válida de registros.")
            return None
            
        df = pd.DataFrame(lista_registros)
        return df
        
    except requests.exceptions.HTTPError as e_http:
        st.error(f"Erro HTTP na integração com o Power Automate: {e_http}")
        return None
    except Exception as e:
        st.error(f"Erro ao processar o JSON recebido: {e}")
        return None

# Tentando carregar a base de dados
df_original = carregar_dados_sharepoint()

if df_original is not None and not df_original.empty:
    df = df_original.copy()
    
    # Mapeamento Direto baseado EXATAMENTE nas chaves brutas que vieram do seu SharePoint
    colunas_map = {
        'ID': 'num_doc',
        'processo_sei': 'processo_sei',
        'siema': 'siema',
        'situacao': 'situacao',
        'laudo_sei': 'laudo_sei',
        'data_acid': 'data_acid',
        'relat_sei': 'relat_sei',
        'instalacao': 'instalacao',
        'campo': 'campo',
        'bacia': 'bacia',
        'empresa': 'empresa',
        'cnpj': 'cnpj',
        'produto': 'produto',
        'class_ol': 'class_ol',
        'class_risco': 'class_risco',
        'vol_char': 'vol_char',
        'lat': 'lat',
        'lon': 'lon',
        'grandeza': 'grandeza',
        'nivel_pontos': 'nivel_pontos',
        'nivel': 'nivel',
        'auto': 'auto',
        'multa_char': 'multa_char',
        'data_ai': 'data_ai'
    }
    
    # Aplica o mapeamento de colunas de forma segura
    for col_real, col_interna in colunas_map.items():
        if col_real in df.columns:
            df[col_interna] = df[col_real]
        else:
            df[col_interna] = ""

    # --- FILTRAGEM INTELIGENTE ---
    # Limpa valores nulos e garante que campos numéricos ou strings vazias sejam tratados corretamente
    df = df.dropna(subset=['siema'])
    df['siema'] = df['siema'].astype(str).str.strip()
    
    # Converte laudo para string limpando decimais flutuantes se houver (ex: 123.0 -> 123)
    def normalizar_laudo(val):
        if pd.isna(val) or val == "": return ""
        val_str = str(val).split('.')[0].strip()
        return "" if val_str in ["0", "nan", "None"] else val_str

    df['laudo_sei'] = df['laudo_sei'].apply(normalizar_laudo)
    
    # Filtrando apenas linhas com SIEMA preenchido e Laudo existente
    df_filtrado = df[(df['siema'] != "") & (df['siema'] != "nan") & (df['laudo_sei'] != "")].copy()
    
    st.subheader("Fila de Processos Disponíveis (SharePoint)")
    
    if df_filtrado.empty:
        st.success("Não há processos com laudos válidos na base de dados no momento!")
        with st.expander("Ver diagnóstico técnico das linhas originais"):
            st.write("Colunas detectadas:", list(df_original.columns))
            st.write("Dados brutos recebidos:", df_original)
    else:
        # Injeta a coluna de checkmarks na memória do app
        df_filtrado.insert(0, "Selecionar", False)
        
        colunas_resumo = ['Selecionar', 'num_doc', 'siema', 'processo_sei', 'empresa', 'bacia']
        df_exibicao = df_filtrado[colunas_resumo].loc[:, ~df_filtrado[colunas_resumo].columns.duplicated()]
        
        # Renderiza a tabela interativa com checkboxes
        tabela_editada = st.data_editor(
            df_exibicao,
            hide_index=True,
            disabled=['num_doc', 'siema', 'processo_sei', 'empresa', 'bacia'],
            column_config={
                "Selecionar": st.column_config.CheckboxColumn(
                    "Selecionar",
                    help="Marque os processos que deseja gerar o relatório",
                    default=False,
                ),
                "num_doc": "ID"
            }
        )
        
        indices_selecionados = tabela_editada[tabela_editada["Selecionar"] == True].index
        df_selecionados = df_filtrado.iloc[indices_selecionados]
        
        st.write("---")
        st.subheader("Ações de Geração")
        
        if df_selecionados.empty:
            st.info("💡 Marque uma ou mais caixas de seleção acima na coluna **'Selecionar'** para liberar a geração do documento.")
        else:
            st.warning(f"Você selecionou **{len(df_selecionados)}** processo(s) para processamento.")
            
            if st.button("🚀 Gerar Relatórios dos Processos Selecionados"):
                for _, row in df_selecionados.iterrows():
                    modelo_arquivo = selecionar_modelo(row)
                    
                    if not modelo_arquivo:
                        st.error(f"Não foi possível determinar o modelo para o SIEMA {row['siema']}. Verifique as classes de risco e tipo.")
                        continue
                        
                    caminho_modelo = os.path.join("modelos", modelo_arquivo)
                    
                    if not os.path.exists(caminho_modelo):
                        st.error(f"O arquivo de modelo '{modelo_arquivo}' não foi encontrado na pasta 'modelos/'.")
                        continue
                        
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
                        if v in ["nan", "None"]: dados_replace[k] = ""
                    
                    doc_pronto_io = preencher_documento(caminho_modelo, dados_replace)
                    nome_arquivo_saida = f"Rel_Fisc_{row.get('num_doc', 'X')}_{row.get('siema', '')}.docx"
                    
                    with st.container(border=True):
                        st.write(f"📄 **Processo:** {row['processo_sei']} | **SIEMA:** {row['siema']} | **Empresa:** {row['empresa']}")
                        st.download_button(
                            label=f"📥 Baixar Documento Word ({row['siema']})",
                            data=doc_pronto_io,
                            file_name=nome_arquivo_saida,
                            key=f"dl_{row['siema']}",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
else:
    st.info("Aguardando carregamento e resposta válida do Power Automate.")
