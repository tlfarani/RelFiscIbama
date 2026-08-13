import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import io
import os
import requests
import zipfile
from datetime import datetime, timedelta
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="FiscFlow — IBAMA", 
    layout="wide", 
    page_icon="🔄"
)

# --- CUSTOMIZAÇÃO COMPLETA DE INTERFACE (BLINDAGEM VISUAL) ---
st.markdown("""
    <style>
    /* 1. Títulos e Subtítulos em Verde Musgo */
    h1, h2, h3, .stSubheader, [data-testid="stWidgetLabel"] p {
        color: #4E5D30 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
        font-weight: bold !important;
    }

    /* 2. Customização das tags internas escolhidas no Multiselect */
    span[data-baseweb="tag"] {
        background-color: #E9EDDE !important;
        color: #4E5D30 !important;
        border: 1px solid #4E5D30 !important;
    }

    /* 3. Estilização Firme dos Botões (Incluindo Botões de Link) */
    div.stButton > button:first-child, 
    div.stDownloadButton > button:first-child, 
    div.stLinkButton > a {
        background-color: #4E5D30 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: 1px solid #4E5D30 !important;
        padding: 10px 24px !important;
        font-weight: bold !important;
        text-decoration: none !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 100% !important;
    }
    div.stButton > button:first-child p, div.stDownloadButton > button:first-child p {
        color: #FFFFFF !important;
    }
    div.stButton > button:first-child:hover, 
    div.stDownloadButton > button:first-child:hover, 
    div.stLinkButton > a:hover {
        background-color: #3A471E !important;
        border-color: #3A471E !important;
        color: #FFFFFF !important;
    }

    /* 4. BLINDAGEM DOS FILTROS (Força fundo claro e remove o bloco escuro) */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CBD5E1 !important;
    }
    div[data-baseweb="select"] input {
        color: #000000 !important;
    }
    div[data-baseweb="select"] svg {
        fill: #4E5D30 !important;
    }

    /* 5. INJEÇÃO CIRÚRGICA DE CORES ALTERNADAS FIXAS NA TABELA STREAMLIT */
    div[data-testid="stDataEditor"] th {
        background-color: #4E5D30 !important;
        color: #F2F2F2 !important;
        font-weight: bold !important;
    }
    
    div[data-testid="stDataEditor"] tr:nth-child(even) td {
        background-color: #E9EDDE !important;
        color: #000000 !important;
    }
    div[data-testid="stDataEditor"] tr:nth-child(odd) td {
        background-color: #FFFFFF !important;
        color: #000000 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- BARRA LATERAL (NAVEGAÇÃO E ATALHOS) ---
st.sidebar.title("📌 Navegação")
pagina = st.sidebar.radio(
    "Selecione o Módulo:",
    ["👑 Coordenação", "🔬 Análise Técnica", "⚖️ Fiscalização"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.header("🔗 Atalhos Rápidos")

st.sidebar.link_button("⚓ Acessar ProMar", "https://promar.streamlit.app/")

if "sharepoint" in st.secrets and "url_visualizacao" in st.secrets["sharepoint"]:
    st.sidebar.link_button("📊 Planilha de Controle", st.secrets["sharepoint"]["url_visualizacao"])

st.sidebar.markdown("---")

# --- FUNÇÕES DE TRATAMENTO ---

def converter_data_excel(valor):
    val_str = str(valor).strip()
    if not val_str or val_str in ["nan", "None", "0"]:
        return " [ DATA - EDITAR MANUAL ] "
    if val_str.isdigit():
        try:
            dias = int(val_str)
            data_real = datetime(1900, 1, 1) + timedelta(days=dias - 2)
            return data_real.strftime("%d/%m/%Y")
        except: pass
    return val_str

def extrair_volume_numerico(valor):
    if pd.isna(valor): return 0.0
    try: return float(str(valor).strip().replace(",", "."))
    except ValueError: return 0.0

def extrair_volume_texto(valor):
    if pd.isna(valor) or str(valor).strip() == "": return " [ VOLUME - EDITAR MANUAL ] "
    val_str = str(valor).strip().lower()
    try:
        num_float = float(val_str.replace(",", "."))
        if num_float == 0.0: return "0"
        texto_formatado = f"{num_float:.7f}"
        if "." in texto_formatado:
            texto_formatado = texto_formatado.rstrip('0').rstrip('.')
        return texto_formatado.replace(".", ",")
    except ValueError: return val_str.replace(".", ",")

def t_tag(valor, nome_tag):
    v_s = str(valor).strip()
    if v_s in ["", "nan", "None", "0", "Processo Não Encontrado"]: 
        return f" [ {nome_tag.upper()} - EDITAR MANUAL ] "
    if nome_tag == "siema" and "fora do ar" in v_s.lower():
        return " [ SIEMA FORA DO AR - EDITAR MANUAL ] "
    return v_s

# --- FUNÇÕES DE NEGÓCIO ---

def determinar_jurisdicao(bacia):
    bacia_limpa = str(bacia).lower().strip()
    if "santos" in bacia_limpa: return "-SP"
    if "campos" in bacia_limpa: return "-RJ"
    if "espirito santo" in bacia_limpa: return "-ES"
    return ""

def processar_grandeza(grandeza):
    g = str(grandeza).strip().title()
    tabela = {
        "Potencial": ("quando as consequências não são evidentes", "5"),
        "Reduzida": ("quando os danos ambientais são locais ou temporários", "15"),
        "Fraca": ("quando os danos ambientais são de pequena proporção ou de baixa complexidade, gravidade ou magnitude, diante do contexto considerado", "30"),
        "Moderada": ("quando os danos ambientais são de proporção intermediária ou de moderada complexidade, gravidade ou magnitude, diante do contexto considerado", "50"),
        "Grave": ("quando os danos ambientais são de grande proporção ou de alta complexidade, gravidade ou magnitude, diante do contexto considerado", "70")
    }
    return tabela.get(g, (" [ GRANDEZA TEXTO - EDITAR MANUAL ] ", " [ PONTOS GRANDEZA - EDITAR MANUAL ] "))

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
    vol_num = extrair_volume_numerico(row.get('vol_char', '0'))
    letra_risco = "A"
    for r in ["B", "C", "D", "E"]:
        if r in class_risco_bruto: letra_risco = r; break

    if "Oleoso" in class_ol and "Não" not in class_ol:
        return f"Rel_Fisc_Oleoso_{letra_risco}.docx" if letra_risco in ["A", "B", "D"] else "Rel_Fisc_Oleoso_A.docx", letra_risco
    
    if "Não Oleoso" in class_ol or "Nao Oleoso" in class_ol:
        if letra_risco == "A": return ("Rel_Fisc_Nao_Oleoso_Art_61_A.docx" if vol_num > 8 else "Rel_Fisc_Nao_Oleoso_Art_62_A.docx"), "A"
        if letra_risco == "B": return ("Rel_Fisc_Nao_Oleoso_Art_61_B.docx" if vol_num > 200 else "Rel_Fisc_Nao_Oleoso_Art_62_B.docx"), "B"
        if letra_risco == "D": return "Rel_Fisc_Nao_Oleoso_Art_62_D.docx", "D"
        return "Rel_Fisc_Nao_Oleoso_Art_62_A.docx", "A"
    return None, None

def preencher_documento(caminho_modelo, dicionario_dados):
    doc = Document(caminho_modelo)
    def tratar_p(p):
        for chave, valor in dicionario_dados.items():
            if chave in p.text: p.text = p.text.replace(chave, str(valor))
        if "[" in p.text and "]" in p.text:
            for run in p.runs: run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
    for p in doc.paragraphs: tratar_p(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: tratar_p(p)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_previa_texto(modelo, dicionario_dados):
    if "Art_61" in modelo:
        texto = "Causar poluição decorrente do lançamento irregular de <<vol_char>> m³ (metros cúbicos) de <<produto>>, em <<data_acid>>, pela instalação <<instalacao>>, no <<campo>> localizado na <<bacia>> (coordenadas geográficas <<lat>> / <<lon>>), conforme apurado em processo nº <<processo_sei>>."
    elif "Art_62" in modelo:
        texto = "Lançar <<vol_char>> m³ (metros cúbicos) de <<produto>>, em <<data_acid>>, pela instalação <<instalacao>>, no <<campo>> localizado na <<bacia>> (coordenadas geográficas <<lat>> / <<lon>>), em desacordo com o licenciamento ambiental e legislação ambiental vigente, conforme apurado em processo nº <<processo_sei>>."
    elif "Oleoso" in modelo:
        texto = "Efetuar a descarga de <<vol_char>> m³ (metros cúbicos) de <<produto>>, em <<data_acid>>, pela instalação <<instalacao>>, no <<campo>> localizado na <<bacia>> (coordenadas geográficas <<lat>> / <<lon>>), em desacordo com o licenciamento ambiental e legislação ambiental vigente, conforme apurado em processo nº <<processo_sei>>."
    else:
        return "Texto padrão de infração indisponível para este modelo estrutural."

    for chave, valor in dicionario_dados.items():
        texto = texto.replace(chave, str(valor))
    return texto

# --- CARREGAMENTO CENTRALIZADO DE DADOS (FLUXO UNIFICADO) ---
@st.cache_data(ttl=300) 
def carregar_dados_sharepoint():
    try:
        url = st.secrets["sharepoint"]["url_planilha"]
        headers = {"Content-Type": "application/json"}
        # Informa a ação "LER" de forma padronizada
        resposta = requests.post(url, headers=headers, json={"acao": "LER"})
        resposta.raise_for_status()
        dados_json = resposta.json()
        if isinstance(dados_json, dict) and "value" in dados_json: lista = dados_json["value"]
        elif isinstance(dados_json, list): lista = dados_json
        else: return None
        return pd.DataFrame(lista)
    except Exception as e:
        st.error(f"Erro ao carregar dados do SharePoint: {e}")
        return None

df_original = carregar_dados_sharepoint()

if df_original is not None and not df_original.empty:
    df = df_original.copy()
    
    # 1. Limpa espaços invisíveis nos nomes dos cabeçalhos vindos do SharePoint
    df.columns = df.columns.astype(str).str.strip()

    # 2. Função de busca flexível (ignora maiúsculas/minúsculas, espaços e underlines)
    def buscar_coluna_flexivel(df_input, candidatas):
        cols_norm = {c.lower().replace("_", "").replace(" ", ""): c for c in df_input.columns}
        for cand in candidatas:
            cand_norm = cand.lower().replace("_", "").replace(" ", "")
            if cand_norm in cols_norm:
                return cols_norm[cand_norm]
        return None

    # 3. Mapeamento com busca flexível de variações
    mapeamento_flexivel = {
        'num_doc': ['ID', 'Num_Doc', 'NUM_DOC'],
        'processo_sei': ['PROCESSO', 'Processo', 'PROCESSO_SEI'],
        'siema': ['SIEMA', 'Siema'],
        'situacao': ['SITUACAO', 'Situacao', 'Situação'],
        'laudo_sei': ['LAUDO_SEI', 'Laudo_SEI', 'LAUDO', 'Laudo'],
        'data_acid': ['DATA_ACIDENTE', 'Data_Acidente', 'Data_Acid', 'DATA'],
        'relat_sei': ['RAIPO_SEI', 'Raipo_SEI', 'RAIPO', 'RELAT_SEI'],
        'instalacao': ['INSTALACAO', 'Instalacao', 'Instalação'],
        'campo': ['Campo', 'CAMPO'],
        'bacia': ['Bacia', 'BACIA'],
        'empresa': ['EMPRESA', 'Empresa'],
        'cnpj': ['CNPJ', 'Cnpj'],
        'produto': ['PRODUTO', 'Produto'],
        'class_ol': ['CLASS_OL', 'Class_OL', 'Class OL', 'CLASS OL'],
        'class_risco': ['CLASS_RISCO', 'Class_Risco', 'Class Risco'],
        'vol_char': ['VOL', 'Vol', 'VOLUME', 'Volume'],
        'lat': ['Lat', 'LAT', 'Latitude'],
        'lon': ['Lon', 'LON', 'Longitude'],
        'grandeza': ['Grandeza', 'GRANDEZA'],
        'auto': ['AUTO_INFRACAO', 'Auto_Infracao', 'AUTO'],
        'multa_char': ['MULTA_APLICADA', 'Multa_Aplicada', 'MULTA'],
        'data_ai': ['Data_AI', 'DATA_AI'],
        'multa_prevista': ['MULTA_PREVISTA', 'Multa_Prevista', 'MULTA PREVISTA'],
        'fiscal': ['Fiscal', 'FISCAL', 'Fiscal Responsavel'],
        'nivel': ['Nivel', 'NIVEL'],
        'nivel_pontos': ['Nivel_Pontos', 'NIVEL_PONTOS'],
        'lat_auto': ['Lat_Auto', 'LAT_AUTO'],
        'lon_auto': ['Lon_Auto', 'LON_AUTO'],
        'servidor_laudo': ['SERVIDOR_LAUDO', 'Servidor_Laudo', 'Servidor Laudo', 'SERVIDOR LAUDO', 'Servidor_laudo', 'ANALISTA_LAUDO', 'Analista', 'Servidor']
    }

    for col_interna, candidatas in mapeamento_flexivel.items():
        col_encontrada = buscar_coluna_flexivel(df, candidatas)
        if col_encontrada:
            df[col_interna] = df[col_encontrada]
        else:
            df[col_interna] = ""

    # 🎯 FILTRO DO UNIVERSO AMOSTRAL: Apenas processos com número SEI/PROCESSO preenchido
    df = df[~df['processo_sei'].astype(str).str.strip().isin(["", "nan", "None"])].reset_index(drop=True)

    # Tratamentos padronizados para visualização
    df['s_laudo_limpo'] = df['servidor_laudo'].astype(str).str.strip().replace({"": "Não Atribuído", "nan": "Não Atribuído", "None": "Não Atribuído", "0": "Não Atribuído"})
    df['f_limpo'] = df['fiscal'].astype(str).str.strip().replace({"": "Não Atribuído", "nan": "Não Atribuído", "None": "Não Atribuído", "0": "Não Atribuído"})

    # =========================================================================
    # 👑 PÁGINA 1: COORDENAÇÃO (VISÃO GERAL & GESTÃO DA ESTEIRA)
    # =========================================================================
    if pagina == "👑 Coordenação":
        st.title("👑 Coordenação — Gestão da Esteira & Distribuição")
        st.caption("Painel de acompanhamento macro, distribuição de carga e gestão da força-tarefa")

        tab_dash, tab_planilha = st.tabs(["📊 Dashboard", "📋 Planilha Geral & Atribuições"])

        # --- ABA 1: DASHBOARD ---
        with tab_dash:
            tot_ft = len(df)
            laudo_vazio = df['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])
            pend_laudo = len(df[laudo_vazio])
            
            situ_s = df['situacao'].astype(str).str.strip().str.lower()
            auto_s = df['auto'].astype(str).str.strip()
            
            is_auto = (situ_s == 'auto lavrado') | (~auto_s.isin(["", "nan", "none", "0", "processo não encontrado"]))
            is_ai = situ_s == 'processo ai gerado'
            
            pend_auto = len(df[(~laudo_vazio) & (~is_auto)])
            pend_proc_ai = len(df[is_auto & (~is_ai)])
            concluidos = len(df[is_ai])

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total de Processos FT", tot_ft)
            c2.metric("Pendentes de Laudo", pend_laudo)
            c3.metric("Pendentes de Auto", pend_auto)
            c4.metric("Pendentes Proc. AI", pend_proc_ai)
            c5.metric("Proc. AI Gerados", concluidos)

            st.write("---")

            col_g1, col_g2 = st.columns(2)

            with col_g1:
                st.markdown("### 🔬 Carga por Servidor de Laudo")
                df_l = df.copy()
                df_l['Status_Laudo'] = df_l['laudo_sei'].apply(lambda x: 'Concluído' if str(x).strip() not in ["", "nan", "None"] else 'Pendente')
                df_l_g = df_l.groupby(['s_laudo_limpo', 'Status_Laudo']).size().reset_index(name='Quantidade')
                
                fig_laudo = px.bar(
                    df_l_g, 
                    y='s_laudo_limpo', 
                    x='Quantidade', 
                    color='Status_Laudo',
                    orientation='h',
                    text_auto=True,
                    color_discrete_map={'Pendente': '#EAB308', 'Concluído': '#4E5D30'}
                )
                fig_laudo.update_traces(textposition='auto')
                fig_laudo.update_layout(yaxis_title="Servidor Laudo", xaxis_title="Qtd Processos", barmode='stack', margin=dict(l=0, r=0, t=20, b=0))
                st.plotly_chart(fig_laudo, use_container_width=True)

            with col_g2:
                st.markdown("### ⚖️ Carga por Fiscal Responsável")
                df_fisc = df.copy()
                df_f_g = df_fisc.groupby(['f_limpo', 'situacao']).size().reset_index(name='Quantidade')
                
                fig_fisc = px.bar(
                    df_f_g, 
                    y='f_limpo', 
                    x='Quantidade', 
                    color='situacao',
                    orientation='h',
                    text_auto=True
                )
                fig_fisc.update_traces(textposition='auto')
                fig_fisc.update_layout(yaxis_title="Fiscal", xaxis_title="Qtd Processos", barmode='stack', margin=dict(l=0, r=0, t=20, b=0))
                st.plotly_chart(fig_fisc, use_container_width=True)

            st.write("---")

            st.markdown("### 🌊 Distribuição de Processos por Bacia Sedimentar")
            df_bacia = df.groupby(['bacia', 'situacao']).size().reset_index(name='Quantidade')
            fig_bacia = px.bar(
                df_bacia, 
                x='bacia', 
                y='Quantidade', 
                color='situacao',
                barmode='group',
                text_auto=True
            )
            fig_bacia.update_traces(textposition='outside')
            fig_bacia.update_yaxes(showticklabels=False, showgrid=False, zeroline=False, showline=False, title=None)
            fig_bacia.update_xaxes(showgrid=False)
            fig_bacia.update_layout(
                xaxis_title="Bacia Sedimentar", 
                yaxis_title=None,
                bargap=0.3,
                bargroupgap=0.15,
                margin=dict(l=0, r=0, t=20, b=0)
            )
            st.plotly_chart(fig_bacia, use_container_width=True)

        # --- ABA 2: PLANILHA GERAL E ATRIBUIÇÕES ---
        with tab_planilha:
            st.markdown("### 📋 Base Completa de Processos da Força-Tarefa")

            # --- PAINEL DE FILTROS SUPERIORES ---
            with st.container(border=True):
                st.markdown("**🔍 Filtros de Visualização**")
                f_col1, f_col2, f_col3, f_col4 = st.columns(4)
                
                with f_col1:
                    op_bacia_c = ["Todas"] + sorted([b for b in df['bacia'].astype(str).unique() if b and b != "nan"])
                    sel_bacia_c = st.selectbox("Bacia Sedimentar:", op_bacia_c, index=0)
                with f_col2:
                    op_serv_c = ["Todos"] + sorted(df['s_laudo_limpo'].unique())
                    sel_serv_c = st.selectbox("Servidor Laudo:", op_serv_c, index=0)
                with f_col3:
                    op_fisc_c = ["Todos"] + sorted(df['f_limpo'].unique())
                    sel_fisc_c = st.selectbox("Fiscal Responsável:", op_fisc_c, index=0)
                with f_col4:
                    op_situ_c = ["Todas"] + sorted(df['situacao'].astype(str).unique())
                    sel_situ_c = st.selectbox("Situação:", op_situ_c, index=0)

            # Aplicação dos Filtros
            df_coord = df.copy()
            if sel_bacia_c != "Todas":
                df_coord = df_coord[df_coord['bacia'].astype(str) == sel_bacia_c]
            if sel_serv_c != "Todos":
                df_coord = df_coord[df_coord['s_laudo_limpo'].astype(str) == sel_serv_c]
            if sel_fisc_c != "Todos":
                df_coord = df_coord[df_coord['f_limpo'].astype(str) == sel_fisc_c]
            if sel_situ_c != "Todas":
                df_coord = df_coord[df_coord['situacao'].astype(str) == sel_situ_c]

            df_coord = df_coord.reset_index(drop=True)

            # --- PAINEL DE AÇÕES DE ATRIBUIÇÃO EM LOTE ---
            with st.container(border=True):
                st.markdown("**⚡ Painel de Atribuição e Alteração em Lote**")
                a_col1, a_col2, a_col3, a_col4 = st.columns(4)
                
                lista_servidores = sorted([s for s in df['s_laudo_limpo'].unique() if s != "Não Atribuído"])
                lista_fiscais = sorted([f for s in [df['f_limpo'].unique()] for f in s if f != "Não Atribuído"])
                lista_situacoes = sorted([s for s in df['situacao'].astype(str).unique() if s])

                with a_col1:
                    novo_servidor = st.selectbox("Atribuir Servidor Laudo:", ["[ Não Alterar ]"] + lista_servidores)
                with a_col2:
                    novo_fiscal = st.selectbox("Atribuir Fiscal:", ["[ Não Alterar ]"] + lista_fiscais)
                with a_col3:
                    nova_situacao = st.selectbox("Alterar Situação:", ["[ Não Alterar ]"] + lista_situacoes)
                with a_col4:
                    st.write("") # alinhamento vertical do botão
                    st.write("")
                    btn_atualizar = st.button("🔄 Aplicar Alterações no SharePoint", use_container_width=True)

            marcar_coord = st.checkbox("✅ Marcar todos os processos visíveis abaixo", value=False)

            # --- TABELA INTERATIVA COM CHECKBOXES ---
            df_coord_exib = pd.DataFrame({
                "Selecionar": [marcar_coord] * len(df_coord),
                "ID": df_coord['num_doc'].astype(str),
                "PROCESSO": df_coord['processo_sei'].astype(str),
                "SITUAÇÃO": df_coord['situacao'].astype(str),
                "Servidor Laudo": df_coord['s_laudo_limpo'].astype(str),
                "Fiscal": df_coord['f_limpo'].astype(str),
                "Bacia": df_coord['bacia'].astype(str),
                "Empresa": df_coord['empresa'].astype(str),
                "Instalação": df_coord['instalacao'].astype(str),
                "Produto": df_coord['produto'].astype(str),
                "Vol (m³)": [extrair_volume_texto(v) for v in df_coord['vol_char']],
                "Laudo SEI": df_coord['laudo_sei'].astype(str),
                "Auto Infração": df_coord['auto'].astype(str)
            })

            tabela_coord_editada = st.data_editor(
                df_coord_exib,
                hide_index=True,
                use_container_width=True,
                disabled=[col for col in df_coord_exib.columns if col != "Selecionar"],
                column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar")},
                key=f"editor_coord_{marcar_coord}"
            )

            # --- PROCESSAMENTO DO DISPARO PARA O FLUXO UNIFICADO DO POWER AUTOMATE ---
            if btn_atualizar:
                indices_marcados = tabela_coord_editada[tabela_coord_editada["Selecionar"] == True].index
                
                if len(indices_marcados) == 0:
                    st.warning("⚠️ Marque pelo menos um processo na tabela abaixo antes de aplicar as alterações.")
                elif novo_servidor == "[ Não Alterar ]" and novo_fiscal == "[ Não Alterar ]" and nova_situacao == "[ Não Alterar ]":
                    st.info("💡 Escolha ao menos uma alteração (Servidor, Fiscal ou Situação) nos menus acima.")
                else:
                    processos_alvo = [str(p) for p in df_coord.iloc[indices_marcados]['processo_sei'].tolist()]
                    ids_alvo = [str(i) for i in df_coord.iloc[indices_marcados]['num_doc'].tolist()]

                    payload_atualizacao = {
                        "acao": "ATUALIZAR",
                        "processos_sei": processos_alvo,
                        "ids": ids_alvo,
                        "novo_servidor_laudo": novo_servidor if novo_servidor != "[ Não Alterar ]" else "",
                        "novo_fiscal": novo_fiscal if novo_fiscal != "[ Não Alterar ]" else "",
                        "nova_situacao": nova_situacao if nova_situacao != "[ Não Alterar ]" else ""
                    }

                    url_planilha = st.secrets["sharepoint"]["url_planilha"]
                    
                    try:
                        with st.spinner("Atualizando registros no SharePoint..."):
                            resp = requests.post(url_planilha, json=payload_atualizacao, headers={"Content-Type": "application/json"})
                            resp.raise_for_status()
                        
                        st.success(f"✅ Sucesso! {len(processos_alvo)} processos atualizados no SharePoint.")
                        
                        # Limpa o cache local e recarrega a página
                        st.cache_data.clear()
                        st.rerun()

                    except Exception as e:
                        st.error(f"Erro ao enviar atualização para o Power Automate: {e}")

    # =========================================================================
    # 🔬 PÁGINA 2: ANÁLISE TÉCNICA (INSTRUÇÃO DE LAUDOS)
    # =========================================================================
    elif pagina == "🔬 Análise Técnica":
        st.title("🔬 Análise Técnica — Instrução de Laudos")
        st.caption("Acompanhamento da elaboração de laudos técnicos e consolidação de evidências")
        
        # --- FILTROS DE ANÁLISE TÉCNICA ---
        with st.container(border=True):
            st.markdown("**🔍 Painel de Filtros — Análise Técnica**")
            c1, c2, c3 = st.columns(3)
            with c1:
                op_situ = sorted(df['situacao'].astype(str).unique())
                default_situ = ["Fazer Laudo"] if "Fazer Laudo" in op_situ else []
                sel_situ = st.multiselect("SITUAÇÃO:", op_situ, default=default_situ)
            with c2:
                op_serv = sorted(df['s_laudo_limpo'].unique())
                sel_serv = st.multiselect("SERVIDOR LAUDO:", ["Todos"] + op_serv, default=["Todos"])
            with c3:
                apenas_pendentes = st.checkbox("Mostrar apenas pendentes de Laudo SEI", value=False)

        # --- MÉTRICAS DA ANÁLISE TÉCNICA (RESPONDEM APENAS AO FILTRO DE SERVIDOR LAUDO) ---
        df_metricas_laudo = df.copy()
        if sel_serv and "Todos" not in sel_serv:
            df_metricas_laudo = df_metricas_laudo[df_metricas_laudo['s_laudo_limpo'].astype(str).isin(sel_serv)]

        total_analise = len(df_metricas_laudo)
        laudos_pendentes = len(df_metricas_laudo[df_metricas_laudo['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])])
        laudos_concluidos = total_analise - laudos_pendentes
        
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric(label="Processos na Esteira", value=total_analise)
        with m2:
            st.metric(label="Pendentes de Laudo SEI", value=laudos_pendentes)
        with m3:
            st.metric(label="Laudos Elaborados", value=laudos_concluidos)
            
        st.write("---")

        # --- FILTRAGEM DA TABELA ---
        df_laudo = df.copy()
        if sel_situ: 
            df_laudo = df_laudo[df_laudo['situacao'].astype(str).isin(sel_situ)]
        if sel_serv and "Todos" not in sel_serv: 
            df_laudo = df_laudo[df_laudo['s_laudo_limpo'].astype(str).isin(sel_serv)]
        if apenas_pendentes: 
            df_laudo = df_laudo[df_laudo['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])]
            
        df_laudo = df_laudo.reset_index(drop=True)
        
        st.markdown(f"### 📋 Processos em Análise Técnica ({len(df_laudo)} encontrados)")
        
        # --- TABELA DE EXIBIÇÃO DE ANÁLISE TÉCNICA ---
        df_laudo_exib = pd.DataFrame({
            "ID": df_laudo['num_doc'].astype(str),
            "SITUAÇÃO": df_laudo['situacao'].astype(str),
            "Servidor Laudo": df_laudo['s_laudo_limpo'].astype(str),
            "Processo SEI": df_laudo['processo_sei'].astype(str),
            "Laudo SEI": df_laudo['laudo_sei'].astype(str),
            "Data Acidente": [converter_data_excel(d) for d in df_laudo['data_acid']],
            "Empresa": df_laudo['empresa'].astype(str),
            "Instalação": df_laudo['instalacao'].astype(str),
            "Produto": df_laudo['produto'].astype(str),
            "Vol (m³)": [extrair_volume_texto(v) for v in df_laudo['vol_char']],
            "Class OL": df_laudo['class_ol'].astype(str),
            "Risco": df_laudo['class_risco'].astype(str),
            "Fiscal Responsável": df_laudo['f_limpo'].astype(str)
        })
        
        st.dataframe(
            df_laudo_exib,
            hide_index=True,
            use_container_width=True
        )

    # =========================================================================
    # ⚖️ PÁGINA 3: FISCALIZAÇÃO (AUTUAÇÃO & MINUTAS)
    # =========================================================================
    elif pagina == "⚖️ Fiscalização":
        st.title("🔄 FiscFlow — Módulo de Fiscalização")
        st.markdown("### ⚖️ Gestão de Fila e Automação de Relatórios — IBAMA")
        st.caption("Sincronização ativa com o SharePoint | Geração de minutas em lote")

        # --- FILTROS DE FISCALIZAÇÃO ---
        with st.container(border=True):
            st.markdown("**🔍 Painel de Filtros — Fiscalização**")
            c1, c2, c3 = st.columns(3)
            with c1:
                op_situ = sorted(df['situacao'].astype(str).unique())
                sel_situ = st.multiselect("SITUAÇÃO:", op_situ, default=["Autuar"] if "Autuar" in op_situ else [])
            with c2:
                op_fisc = sorted(df['f_limpo'].unique())
                sel_fisc = st.multiselect("FISCAL:", ["Todos"] + op_fisc, default=["Todos"])
            with c3:
                todos_laudos = st.checkbox("Mostrar processos sem LAUDO_SEI", value=False)

        # --- MÉTRICAS DA FISCALIZAÇÃO (RESPONDEM APENAS AO FILTRO DE FISCAL) ---
        df_metricas_fisc = df.copy()
        if sel_fisc and "Todos" not in sel_fisc:
            df_metricas_fisc = df_metricas_fisc[df_metricas_fisc['f_limpo'].astype(str).isin(sel_fisc)]

        total_fila = len(df_metricas_fisc)
        situ_str = df_metricas_fisc['situacao'].astype(str).str.strip()
        auto_str = df_metricas_fisc['auto'].astype(str).str.strip()

        prontos_autuacao = len(df_metricas_fisc[situ_str.str.lower() == 'autuar'])
        
        is_auto_lavrado = (situ_str.str.lower() == 'auto lavrado') | (~auto_str.str.lower().isin(["", "nan", "none", "0", "processo não encontrado"]))
        is_ai_gerado = situ_str.str.lower() == 'processo ai gerado'
        
        autos_lavrados = len(df_metricas_fisc[is_auto_lavrado])
        fisc_pendentes = len(df_metricas_fisc[is_auto_lavrado & (~is_ai_gerado)])
        fisc_gerados = len(df_metricas_fisc[is_ai_gerado])

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.metric(label="Processos no Fluxo", value=total_fila)
        with m2:
            st.metric(label="Prontos p/ Autuação", value=prontos_autuacao)
        with m3:
            st.metric(label="Autos Lavrados", value=autos_lavrados)
        with m4:
            st.metric(label="Proc. Fisc. Pendentes", value=fisc_pendentes)
        with m5:
            st.metric(label="Proc. AI Gerados", value=fisc_gerados)
        
        st.write("---")

        # --- FILTRAGEM DA TABELA E DA GERAÇÃO ---
        df_f = df.copy()
        if sel_situ: df_f = df_f[df_f['situacao'].astype(str).isin(sel_situ)]
        if sel_fisc and "Todos" not in sel_fisc: df_f = df_f[df_f['f_limpo'].astype(str).isin(sel_fisc)]
        if not todos_laudos: df_f = df_f[~df_f['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])]

        df_f = df_f.reset_index(drop=True)

        # --- CONTROLE DE SELEÇÃO EM MASSA ---
        st.markdown("### 📋 Processos para Análise")
        
        marcar_todos = st.checkbox("✅ Marcar todos os processos mostrados abaixo", value=False)
        vetor_selecao_inicial = [marcar_todos] * len(df_f)

        # --- PREPARAÇÃO DA BASE DE EXIBIÇÃO ---
        df_exib = pd.DataFrame({
            "Selecionar": vetor_selecao_inicial,
            "ID": df_f['num_doc'].astype(str),
            "SITUAÇÃO": df_f['situacao'].astype(str),
            "Fiscal": df_f['f_limpo'].astype(str),
            "Data": [converter_data_excel(d) for d in df_f['data_acid']],
            "Produto": df_f['produto'].astype(str),
            "Class OL": df_f['class_ol'].astype(str),
            "Risco": df_f['class_risco'].astype(str),
            "Vol (m³)": [extrair_volume_texto(v) for v in df_f['vol_char']],
            "Multa Prev": df_f['multa_prevista'].astype(str),
            "Lat Auto": df_f['lat_auto'].astype(str),
            "Lon Auto": df_f['lon_auto'].astype(str)
        })

        tabela_editada = st.data_editor(
            df_exib,
            hide_index=True,
            use_container_width=True,
            disabled=[col for col in df_exib.columns if col != "Selecionar"],
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Selecionar")
            },
            key=f"editor_{marcar_todos}"
        )
        
        indices_selecionados = tabela_editada[tabela_editada["Selecionar"] == True].index
        selecionados = df_f.iloc[indices_selecionados]

        if not selecionados.empty:
            st.write("---")
            st.subheader(f"🚀 Geração em Lote ({len(selecionados)} itens)")
            
            # --- LÓGICA DE COMPACTAÇÃO EM ZIP (EM MEMÓRIA) ---
            zip_buffer = io.BytesIO()
            arquivos_para_zipar = 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for _, row in selecionados.iterrows():
                    modelo, risco = extrair_classe_e_modelo(row)
                    if not modelo: continue
                    
                    caminho = os.path.join("modelos", modelo)
                    if not os.path.exists(caminho): continue

                    grandeza_texto, grandeza_pontos = processar_grandeza(row.get('grandeza', ''))
                    nivel_texto = processar_nivel(row.get('nivel', ''))

                    dados = {
                        "<<siema>>": t_tag(row.get('siema', ''), "siema"),
                        "<<processo_sei>>": t_tag(row.get('processo_sei', ''), "processo_sei"),
                        "<<laudo_sei>>": str(row.get('laudo_sei', '')).split('.')[0],
                        "<<data_acid>>": converter_data_excel(row.get('data_acid', '')),
                        "<<relat_sei>>": t_tag(row.get('relat_sei', ''), "raipo_sei"),
                        "<<instalacao>>": t_tag(row.get('instalacao', ''), "instalacao"),
                        "<<campo>>": t_tag(row.get('campo', ''), "campo"),
                        "<<bacia>>": t_tag(row.get('bacia', ''), "bacia"),
                        "<<empresa>>": t_tag(row.get('empresa', ''), "empresa"),
                        "<<cnpj>>": t_tag(row.get('cnpj', ''), "cnpj"),
                        "<<produto>>": t_tag(row.get('produto', ''), "produto"),
                        "<<class_ol>>": t_tag(row.get('class_ol', ''), "class_ol"),
                        "<<class_risco>>": risco,
                        "<<vol_char>>": extrair_volume_texto(row.get('vol_char', '')),
                        "<<lat>>": t_tag(row.get('lat', ''), "lat"),
                        "<<lon>>": t_tag(row.get('lon', ''), "lon"),
                        "<<grandeza>>": t_tag(row.get('grandeza', ''), "grandeza"),
                        "<<grandeza_texto>>": grandeza_texto,
                        "<<grandeza_pontos>>": grandeza_pontos,
                        "<<nivel>>": t_tag(row.get('nivel', ''), "nivel"),
                        "<<nivel_pontos>>": t_tag(row.get('nivel_pontos', ''), "nivel_pontos"),
                        "<<nivel_texto>>": nivel_texto,
                        "<<multa_num>>": t_tag(row.get('multa_char', ''), "multa_aplicada"),
                        "<<multa_char>>": t_tag(row.get('multa_char', ''), "multa_aplicada"),
                        "<<data_ai>>": converter_data_excel(row.get('data_ai', '')),
                        "<<auto>>": t_tag(row.get('auto', ''), "auto_infracao"),
                        "<<jurisdicao>>": determinar_jurisdicao(row.get('bacia', ''))
                    }
                    
                    doc_io = preencher_documento(caminho, dados)
                    nome_arquivo = f"Rel_Fisc_{row['num_doc']}.docx"
                    
                    zip_file.writestr(nome_arquivo, doc_io.getvalue())
                    arquivos_para_zipar += 1

            zip_buffer.seek(0)
            
            if arquivos_para_zipar > 1:
                st.markdown("### 📦 Download Unificado")
                st.download_button(
                    label=f"📥 Baixar Todos os {arquivos_para_zipar} Relatórios (.ZIP)",
                    data=zip_buffer,
                    file_name=f"FiscFlow_pacote_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.write("---")

            st.markdown("### 🔍 Detalhes Individuais dos Itens Selecionados")
            for _, row in selecionados.iterrows():
                modelo, risco = extrair_classe_e_modelo(row)
                if not modelo: continue
                
                caminho = os.path.join("modelos", modelo)
                if not os.path.exists(caminho): continue

                grandeza_texto, grandeza_pontos = processar_grandeza(row.get('grandeza', ''))
                nivel_texto = processar_nivel(row.get('nivel', ''))

                dados_unitarios = {
                    "<<siema>>": t_tag(row.get('siema', ''), "siema"),
                    "<<processo_sei>>": t_tag(row.get('processo_sei', ''), "processo_sei"),
                    "<<laudo_sei>>": str(row.get('laudo_sei', '')).split('.')[0],
                    "<<data_acid>>": converter_data_excel(row.get('data_acid', '')),
                    "<<relat_sei>>": t_tag(row.get('relat_sei', ''), "raipo_sei"),
                    "<<instalacao>>": t_tag(row.get('instalacao', ''), "instalacao"),
                    "<<campo>>": t_tag(row.get('campo', ''), "campo"),
                    "<<bacia>>": t_tag(row.get('bacia', ''), "bacia"),
                    "<<empresa>>": t_tag(row.get('empresa', ''), "empresa"),
                    "<<cnpj>>": t_tag(row.get('cnpj', ''), "cnpj"),
                    "<<produto>>": t_tag(row.get('produto', ''), "produto"),
                    "<<class_ol>>": t_tag(row.get('class_ol', ''), "class_ol"),
                    "<<class_risco>>": risco,
                    "<<vol_char>>": extrair_volume_texto(row.get('vol_char', '')),
                    "<<lat>>": t_tag(row.get('lat', ''), "lat"),
                    "<<lon>>": t_tag(row.get('lon', ''), "lon"),
                    "<<grandeza>>": t_tag(row.get('grandeza', ''), "grandeza"),
                    "<<grandeza_texto>>": grandeza_texto,
                    "<<grandeza_pontos>>": grandeza_pontos,
                    "<<nivel>>": t_tag(row.get('nivel', ''), "nivel"),
                    "<<nivel_pontos>>": t_tag(row.get('nivel_pontos', ''), "nivel_pontos"),
                    "<<nivel_texto>>": nivel_texto,
                    "<<multa_num>>": t_tag(row.get('multa_char', ''), "multa_aplicada"),
                    "<<multa_char>>": t_tag(row.get('multa_char', ''), "multa_aplicada"),
                    "<<data_ai>>": converter_data_excel(row.get('data_ai', '')),
                    "<<auto>>": t_tag(row.get('auto', ''), "auto_infracao"),
                    "<<jurisdicao>>": determinar_jurisdicao(row.get('bacia', ''))
                }
                
                doc_io_unitario = preencher_documento(caminho, dados_unitarios)
                nome = f"Rel_Fisc_{row['num_doc']}.docx"
                
                texto_previa = gerar_previa_texto(modelo, dados_unitarios)
                
                with st.container(border=True):
                    st.write(f"📄 **ID:** {row['num_doc']} | **Processo:** {row['processo_sei']} | **Empresa:** {row['empresa']}")
                    
                    st.markdown("**📝 Descrição da Infração (Prévia do Auto e Relatório de Fiscalização):**")
                    st.markdown(f"> *{texto_previa}*")
                    
                    st.download_button(label="Baixar Relatório Isolado", data=doc_io_unitario, file_name=nome, key=f"dl_{row['num_doc']}")

else:
    st.info("Aguardando carregamento dos dados do SharePoint...")
