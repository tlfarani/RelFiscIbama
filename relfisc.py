import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import io
import os
import requests
import zipfile
import random
import json
import hashlib
import secrets
from datetime import datetime, timedelta
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="FiscFlow — IBAMA", 
    layout="wide", 
    page_icon="⚖️"
)

# --- CUSTOMIZAÇÃO VISUAL INSTITUCIONAL ---
st.markdown("""
    <style>
    h1, h2, h3, .stSubheader, [data-testid="stWidgetLabel"] p {
        color: #4E5D30 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
        font-weight: bold !important;
    }
    span[data-baseweb="tag"] {
        background-color: #E9EDDE !important;
        color: #4E5D30 !important;
        border: 1px solid #4E5D30 !important;
    }
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

# =========================================================================
# 🔒 CRIPTOGRAFIA DE SENHAS (PBKDF2-HMAC-SHA256 COM SALT)
# =========================================================================
def gerar_hash(senha: str) -> str:
    """Gera um hash criptográfico com salt aleatório no formato salt$hash"""
    if not senha: return ""
    salt = secrets.token_hex(16)
    chave = hashlib.pbkdf2_hmac('sha256', senha.strip().encode('utf-8'), bytes.fromhex(salt), 100_000)
    return f"{salt}${chave.hex()}"

def validar_senha(senha_digitada: str, hash_armazenado: str) -> bool:
    """Valida a senha contra o hash armazenado, aceitando formato legado de migração"""
    if not senha_digitada or not hash_armazenado: return False
    senha_digitada = senha_digitada.strip()
    hash_armazenado = str(hash_armazenado).strip()

    if "$" in hash_armazenado:
        try:
            salt, chave_esperada = hash_armazenado.split("$", 1)
            chave_calculada = hashlib.pbkdf2_hmac('sha256', senha_digitada.encode('utf-8'), bytes.fromhex(salt), 100_000).hex()
            return secrets.compare_digest(chave_calculada, chave_esperada)
        except Exception:
            return False

    # Suporte para a senha inicial do coordenador inserida manualmente em texto puro
    if senha_digitada == hash_armazenado: return True
    # Suporte legado SHA-256
    hash_simples = hashlib.sha256(senha_digitada.encode('utf-8')).hexdigest()
    return secrets.compare_digest(hash_simples, hash_armazenado)

# --- FUNÇÕES AUXILIARES DE TRATAMENTO ---
def converter_data_excel(valor):
    val_str = str(valor).strip()
    if not val_str or val_str in ["nan", "None", "0"]: return " [ DATA - EDITAR MANUAL ] "
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
    if v_s in ["", "nan", "None", "0", "Processo Não Encontrado"]: return f" [ {nome_tag.upper()} - EDITAR MANUAL ] "
    if nome_tag == "siema" and "fora do ar" in v_s.lower(): return " [ SIEMA FORA DO AR - EDITAR MANUAL ] "
    return v_s

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

# --- CARREGAMENTO CENTRALIZADO VIA POWER AUTOMATE ---
@st.cache_data(ttl=300) 
def carregar_dados_sharepoint():
    try:
        url = st.secrets["sharepoint"]["url_planilha"]
        headers = {"Content-Type": "application/json"}
        resposta = requests.post(url, headers=headers, json={"acao": "LER"})
        
        # Se retornar erro 400, exibe o corpo exato da resposta da Microsoft
        if not resposta.ok:
            st.error(f"Erro {resposta.status_code} retornado pelo Power Automate:")
            st.code(resposta.text, language="json")
            return None, None
            
        resposta.raise_for_status()
    
        
        texto_resposta = resposta.text.strip()
        try:
            dados_json = json.loads(texto_resposta)
        except json.JSONDecodeError as e:
            decoder = json.JSONDecoder()
            dados_json, _ = decoder.raw_decode(texto_resposta)

        if isinstance(dados_json, dict):
            proc_list = dados_json.get("processos", dados_json.get("value", []))
            equipe_list = dados_json.get("equipe", [])
        elif isinstance(dados_json, list):
            proc_list = dados_json
            equipe_list = []
        else:
            return None, None
            
        return pd.DataFrame(proc_list), pd.DataFrame(equipe_list)
    except Exception as e:
        st.error(f"Erro ao carregar dados do SharePoint: {e}")
        return None, None

df_original, df_equipe_raw = carregar_dados_sharepoint()

# --- DIAGNÓSTICO EM TEMPO REAL NA BARRA LATERAL --- Exibe diagnóstico técnico exclusivamente para a Coordenação
    if is_coordenador:
        with st.sidebar.expander("🛠️ Diagnóstico da Conexão", expanded=False):
            qtd_proc_bruta = len(df_original) if df_original is not None else 0
            qtd_eq_bruta = len(df_equipe_raw) if df_equipe_raw is not None else 0
            st.caption("Visível apenas para Coordenação:")
            st.write(f"📦 Linhas brutas em Processos: **{qtd_proc_bruta}**")
            st.write(f"👥 Linhas brutas em Equipe: **{qtd_eq_bruta}**")
            if df_original is not None and not df_original.empty:
                st.write("Colunas detectadas em Processos:", list(df_original.columns))

if df_original is not None and not df_original.empty:
    df = df_original.copy()
    df.columns = df.columns.astype(str).str.strip()

    def buscar_coluna_flexivel(df_input, candidatas):
        cols_norm = {c.lower().replace("_", "").replace(" ", ""): c for c in df_input.columns}
        for cand in candidatas:
            cand_norm = cand.lower().replace("_", "").replace(" ", "")
            if cand_norm in cols_norm: return cols_norm[cand_norm]
        return None

    mapeamento_flexivel = {
        'num_doc': ['ID', 'Num_Doc', 'NUM_DOC'],
        'processo_sei': ['PROCESSO', 'Processo', 'PROCESSO_SEI', 'Processo_SEI', 'Processo SEI', 'N_PROCESSO', 'NUM_PROCESSO'],
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
        'servidor_laudo': ['SERVIDOR_LAUDO', 'Servidor_Laudo', 'Servidor Laudo', 'SERVIDOR LAUDO', 'ANALISTA_LAUDO', 'Analista']
    }

    for col_interna, candidatas in mapeamento_flexivel.items():
        col_encontrada = buscar_coluna_flexivel(df, candidatas)
        if col_encontrada: df[col_interna] = df[col_encontrada]
        else: df[col_interna] = ""

    df = df[~df['processo_sei'].astype(str).str.strip().isin(["", "nan", "None"])].reset_index(drop=True)
    df['s_laudo_limpo'] = df['servidor_laudo'].astype(str).str.strip().replace({"": "Não Atribuído", "nan": "Não Atribuído", "None": "Não Atribuído", "0": "Não Atribuído"})
    df['f_limpo'] = df['fiscal'].astype(str).str.strip().replace({"": "Não Atribuído", "nan": "Não Atribuído", "None": "Não Atribuído", "0": "Não Atribuído"})

    # --- PROCESSAMENTO DA TABELA EQUIPE (DO ARQUIVO RESTRITO) ---
    df_equipe = pd.DataFrame()
    if df_equipe_raw is not None and not df_equipe_raw.empty:
        df_equipe_raw.columns = df_equipe_raw.columns.astype(str).str.strip()
        c_nome = buscar_coluna_flexivel(df_equipe_raw, ['Nome', 'NOME', 'Servidor'])
        c_email = buscar_coluna_flexivel(df_equipe_raw, ['E_Mail', 'Email', 'E-mail', 'MAIL'])
        c_senha = buscar_coluna_flexivel(df_equipe_raw, ['Senha', 'SENHA', 'Password'])
        c_status = buscar_coluna_flexivel(df_equipe_raw, ['Status', 'STATUS', 'Situacao'])
        
        c_coord = buscar_coluna_flexivel(df_equipe_raw, ['Coordenacao', 'Coordenação', 'Coord'])
        c_analise = buscar_coluna_flexivel(df_equipe_raw, ['Analise', 'Análise', 'Analise_Tecnica'])
        c_fisc = buscar_coluna_flexivel(df_equipe_raw, ['Fiscalizacao', 'Fiscalização', 'Fisc'])
        c_perfil_legado = buscar_coluna_flexivel(df_equipe_raw, ['Perfil', 'PERFIL'])

        nomes = df_equipe_raw[c_nome].astype(str).str.strip() if c_nome else ""
        emails = df_equipe_raw[c_email].astype(str).str.strip() if c_email else ""
        senhas = df_equipe_raw[c_senha].astype(str).str.strip() if c_senha else ""
        statuses = df_equipe_raw[c_status].astype(str).str.strip().str.upper() if c_status else "APROVADO"

        def check_bool(val, perfil_str, tag):
            if str(val).strip().upper() in ["SIM", "TRUE", "1", "S"]: return True
            if perfil_str and tag in str(perfil_str).lower(): return True
            return False

        p_legado = df_equipe_raw[c_perfil_legado].astype(str) if c_perfil_legado else ""
        is_coord = [check_bool(df_equipe_raw[c_coord].iloc[i] if c_coord else False, p_legado.iloc[i] if c_perfil_legado else "", "coord") for i in range(len(df_equipe_raw))]
        is_an = [check_bool(df_equipe_raw[c_analise].iloc[i] if c_analise else False, p_legado.iloc[i] if c_perfil_legado else "", "análise") or "tecnica" in str(p_legado.iloc[i] if c_perfil_legado else "").lower() for i in range(len(df_equipe_raw))]
        is_fi = [check_bool(df_equipe_raw[c_fisc].iloc[i] if c_fisc else False, p_legado.iloc[i] if c_perfil_legado else "", "fisc") for i in range(len(df_equipe_raw))]

        df_equipe = pd.DataFrame({
            'nome': nomes,
            'email': emails,
            'senha': senhas,
            'status': statuses,
            'is_coordenacao': is_coord,
            'is_analise': is_an,
            'is_fiscalizacao': is_fi
        })
        df_equipe = df_equipe[~df_equipe['nome'].isin(["", "nan", "None"])].reset_index(drop=True)

    # =========================================================================
    # 🔒 SISTEMA DE LOGIN & CADASTRO DE USUÁRIOS
    # =========================================================================
    if "usuario_logado" not in st.session_state:
        st.session_state["usuario_logado"] = None

    # TELA DE LOGIN BLOQUEANTE
    if st.session_state["usuario_logado"] is None:
        col_esq, col_card, col_dir = st.columns([1, 1.8, 1])
        with col_card:
            st.markdown("<h1 style='text-align: center; color: #4E5D30;'>⚖️ FiscFlow — IBAMA</h1>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: #666;'>Sistema de Gestão de Esteira, Atribuições e Fiscalização</p>", unsafe_allow_html=True)
            st.write("")

            tab_entrar, tab_cadastrar = st.tabs(["🔑 Entrar no Sistema", "📝 Criar Conta / Solicitar Acesso"])

            # --- ABA 1: ENTRAR ---
            with tab_entrar:
                with st.form("form_login"):
                    login_email = st.text_input("E-mail Cadastrado:", placeholder="seu.email@ibama.gov.br").strip().lower()
                    login_senha = st.text_input("Senha:", type="password")
                    btn_login = st.form_submit_button("Entrar", use_container_width=True)

                    if btn_login:
                        if not login_email or not login_senha:
                            st.warning("⚠️ Preencha seu e-mail e senha.")
                        elif df_equipe.empty:
                            st.error("Erro: A tabela de equipe não foi carregada do SharePoint.")
                        else:
                            # Procura o usuário
                            match_user = df_equipe[df_equipe['email'].str.lower() == login_email]
                            if match_user.empty:
                                st.error("⛔ E-mail não encontrado. Caso seja seu primeiro acesso, solicite seu cadastro na aba ao lado.")
                            else:
                                user_data = match_user.iloc[0]
                                status_u = str(user_data['status']).strip().upper()
                                
                                if status_u == "PENDENTE":
                                    st.warning("⏳ **Solicitação em Análise:** Seu cadastro aguarda aprovação da Coordenação via e-mail.")
                                elif status_u == "REJEITADO":
                                    st.error("⛔ Seu acesso não foi autorizado pela Coordenação.")
                                elif not validar_senha(login_senha, user_data['senha']):
                                    st.error("⛔ Senha incorreta.")
                                else:
                                    st.session_state["usuario_logado"] = {
                                        "nome": str(user_data['nome']),
                                        "email": str(user_data['email']),
                                        "is_coordenacao": bool(user_data['is_coordenacao']),
                                        "is_analise": bool(user_data['is_analise']),
                                        "is_fiscalizacao": bool(user_data['is_fiscalizacao'])
                                    }
                                    st.success(f"✅ Bem-vindo(a), {user_data['nome']}!")
                                    st.rerun()

            # --- ABA 2: CADASTRAR NOVO USUÁRIO ---
            with tab_cadastrar:
                with st.form("form_cadastro"):
                    st.caption("Preencha seus dados funcionais. O pedido de acesso será encaminhado para aprovação da Coordenação.")
                    novo_nome = st.text_input("Nome Completo:", placeholder="Ex: Nome Sobrenome").strip()
                    novo_email = st.text_input("E-mail Institucional:", placeholder="nome.sobrenome@ibama.gov.br").strip().lower()
                    nova_senha = st.text_input("Crie uma Senha:", type="password")
                    confirma_senha = st.text_input("Confirme a Senha:", type="password")
                    
                    st.markdown("**Selecione as suas atribuições:**")
                    cad_analise = st.checkbox("🔬 Análise Técnica (Elaboração de Laudos)", value=True)
                    cad_fisc = st.checkbox("⚖️ Fiscalização (Lavratura de Autos de Infração)", value=False)
                    cad_coord = st.checkbox("👑 Coordenação (Gestão macro da Força-Tarefa)", value=False)

                    btn_cadastrar = st.form_submit_button("Solicitar Acesso", use_container_width=True)

                    if btn_cadastrar:
                        if not novo_nome or not novo_email or not nova_senha:
                            st.warning("⚠️ Todos os campos são obrigatórios.")
                        elif nova_senha != confirma_senha:
                            st.error("⛔ As senhas digitadas não coincidem.")
                        elif len(nova_senha) < 4:
                            st.warning("⚠️ A senha deve conter pelo menos 4 caracteres.")
                        elif not df_equipe.empty and novo_email in df_equipe['email'].str.lower().values:
                            st.error("⛔ Este e-mail já possui cadastro. Acesse a aba 'Entrar no Sistema'.")
                        else:
                            # A senha é criptografada com PBKDF2-HMAC-SHA256 antes de trafegar
                            payload_cadastro = {
                                "acao": "CADASTRAR_USUARIO",
                                "nome": novo_nome,
                                "email": novo_email,
                                "senha": gerar_hash(nova_senha),
                                "status": "PENDENTE",
                                "coordenacao": "SIM" if cad_coord else "NÃO",
                                "analise": "SIM" if cad_analise else "NÃO",
                                "fiscalizacao": "SIM" if cad_fisc else "NÃO"
                            }
                            try:
                                url_planilha = st.secrets["sharepoint"]["url_planilha"]
                                with st.spinner("Enviando solicitação para o SharePoint..."):
                                    resp = requests.post(url_planilha, json=payload_cadastro, headers={"Content-Type": "application/json"})
                                    resp.raise_for_status()

                                st.info("📨 **Solicitação enviada com sucesso!**\n\nUm e-mail de autorização foi encaminhado para a Coordenação. Assim que aprovado, você poderá acessar o FiscFlow.")
                                st.cache_data.clear()
                            except Exception as e:
                                st.error(f"Erro ao conectar com o Power Automate: {e}")

        st.stop() # Interrompe a execução até que o login ocorra com sucesso

    # =========================================================================
    # 📱 APLICATIVO AUTENTICADO
    # =========================================================================
    user = st.session_state["usuario_logado"]
    is_coordenador = user["is_coordenacao"]
    is_analise = user["is_analise"]
    is_fiscal = user["is_fiscalizacao"]

    st.sidebar.title("📌 FiscFlow")
    st.sidebar.markdown(f"👤 **{user['nome']}**\n\n✉️ `{user['email']}`")
    
    perfis_badges = []
    if is_coordenador: perfis_badges.append("👑 Coordenação")
    if is_analise: perfis_badges.append("🔬 Análise")
    if is_fiscal: perfis_badges.append("⚖️ Fiscal")
    st.sidebar.caption("Perfis: " + " | ".join(perfis_badges))

    if st.sidebar.button("🚪 Sair / Logoff", use_container_width=True):
        st.session_state["usuario_logado"] = None
        st.rerun()

    st.sidebar.markdown("---")

    # Módulos disponíveis de acordo com as permissões do usuário
    modulos_disponiveis = []
    if is_coordenador:
        modulos_disponiveis = ["👑 Coordenação", "🔬 Análise Técnica", "⚖️ Fiscalização"]
    else:
        if is_analise: modulos_disponiveis.append("🔬 Análise Técnica")
        if is_fiscal: modulos_disponiveis.append("⚖️ Fiscalização")

    if not modulos_disponiveis:
        st.error("Seu usuário não possui nenhum módulo habilitado no momento.")
        st.stop()

    pagina = st.sidebar.radio("Selecione o Módulo:", modulos_disponiveis, index=0)

    st.sidebar.markdown("---")
    st.sidebar.header("🔗 Atalhos Rápidos")
    st.sidebar.link_button("⚓ Acessar ProMar", "https://promar.streamlit.app/")
    if "sharepoint" in st.secrets and "url_visualizacao" in st.secrets["sharepoint"]:
        st.sidebar.link_button("📊 Planilha de Controle", st.secrets["sharepoint"]["url_visualizacao"])
    st.sidebar.markdown("---")

    # =========================================================================
    # 👑 MÓDULO 1: COORDENAÇÃO (EXCLUSIVO PARA COORDENADORES)
    # =========================================================================
    if pagina == "👑 Coordenação":
        st.title("👑 Coordenação — Gestão da Esteira & Distribuição")
        st.caption("Painel macro, distribuição de carga ativa e governança da equipe")

        tab_dash, tab_planilha, tab_auto, tab_equipe = st.tabs([
            "📊 Dashboard", 
            "📋 Planilha Geral & Atribuições", 
            "🎲 Atribuição Automática",
            "👥 Gestão da Equipe"
        ])

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
                st.markdown("### 🔬 Carga Ativa por Servidor de Laudo")
                df_l = df.copy()
                df_l['is_pendente'] = (df_l['situacao'].astype(str).str.strip().str.lower() == 'fazer laudo') & \
                                      (df_l['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None", "0"]))
                df_l['Status_Laudo'] = df_l['is_pendente'].apply(lambda x: 'Pendente (Fazer Laudo)' if x else 'Concluído')
                df_l_g = df_l.groupby(['s_laudo_limpo', 'Status_Laudo']).size().reset_index(name='Quantidade')
                
                fig_laudo = px.bar(
                    df_l_g, 
                    y='s_laudo_limpo', 
                    x='Quantidade', 
                    color='Status_Laudo',
                    orientation='h',
                    text_auto=True,
                    color_discrete_map={'Pendente (Fazer Laudo)': '#EAB308', 'Concluído': '#4E5D30'}
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
            fig_bacia.update_layout(xaxis_title="Bacia Sedimentar", yaxis_title=None, bargap=0.3, bargroupgap=0.15, margin=dict(l=0, r=0, t=20, b=0))
            st.plotly_chart(fig_bacia, use_container_width=True)

        # --- ABA 2: PLANILHA GERAL E ATRIBUIÇÕES ---
        with tab_planilha:
            st.markdown("### 📋 Base Completa de Processos da Força-Tarefa")

            with st.container(border=True):
                st.markdown("**🔍 Filtros de Visualização**")
                f_col1, f_col2, f_col3, f_col4 = st.columns(4)
                
                with f_col1:
                    op_bacia_c = ["Todas"] + sorted([b for b in df['bacia'].astype(str).unique() if b and b != "nan"])
                    sel_bacia_c = st.selectbox("Bacia Sedimentar:", op_bacia_c, index=0)
                with f_col2:
                    op_serv_c = ["Todos"] + sorted(list(set(df['s_laudo_limpo'].unique().tolist() + (df_equipe[df_equipe['is_analise'] == True]['nome'].tolist() if not df_equipe.empty else []))))
                    sel_serv_c = st.selectbox("Servidor Laudo:", op_serv_c, index=0)
                with f_col3:
                    op_fisc_c = ["Todos"] + sorted(list(set(df['f_limpo'].unique().tolist() + (df_equipe[df_equipe['is_fiscalizacao'] == True]['nome'].tolist() if not df_equipe.empty else []))))
                    sel_fisc_c = st.selectbox("Fiscal Responsável:", op_fisc_c, index=0)
                with f_col4:
                    op_situ_c = ["Todas"] + sorted(df['situacao'].astype(str).unique())
                    sel_situ_c = st.selectbox("Situação:", op_situ_c, index=0)

            df_coord = df.copy()
            if sel_bacia_c != "Todas": df_coord = df_coord[df_coord['bacia'].astype(str) == sel_bacia_c]
            if sel_serv_c != "Todos": df_coord = df_coord[df_coord['s_laudo_limpo'].astype(str) == sel_serv_c]
            if sel_fisc_c != "Todos": df_coord = df_coord[df_coord['f_limpo'].astype(str) == sel_fisc_c]
            if sel_situ_c != "Todas": df_coord = df_coord[df_coord['situacao'].astype(str) == sel_situ_c]

            df_coord = df_coord.reset_index(drop=True)

            with st.container(border=True):
                st.markdown("**⚡ Painel de Atribuição e Alteração em Lote**")
                a_col1, a_col2, a_col3, a_col4 = st.columns(4)
                
                if not df_equipe.empty:
                    eq_analise = df_equipe[(df_equipe['is_analise'] == True) & (df_equipe['status'] == 'APROVADO')]['nome'].tolist()
                    eq_fiscal = df_equipe[(df_equipe['is_fiscalizacao'] == True) & (df_equipe['status'] == 'APROVADO')]['nome'].tolist()
                    lista_servidores = sorted(list(set(eq_analise + [s for s in df['s_laudo_limpo'].unique() if s != "Não Atribuído"])))
                    lista_fiscais = sorted(list(set(eq_fiscal + [f for f in df['f_limpo'].unique() if f != "Não Atribuído"])))
                else:
                    lista_servidores = sorted([s for s in df['s_laudo_limpo'].unique() if s != "Não Atribuído"])
                    lista_fiscais = sorted([f for f in df['f_limpo'].unique() if f != "Não Atribuído"])

                lista_situacoes = sorted([s for s in df['situacao'].astype(str).unique() if s])

                with a_col1: novo_servidor = st.selectbox("Atribuir Servidor Laudo:", ["[ Não Alterar ]"] + lista_servidores)
                with a_col2: novo_fiscal = st.selectbox("Atribuir Fiscal:", ["[ Não Alterar ]"] + lista_fiscais)
                with a_col3: nova_situacao = st.selectbox("Alterar Situação:", ["[ Não Alterar ]"] + lista_situacoes)
                with a_col4:
                    st.write("")
                    st.write("")
                    btn_atualizar = st.button("🔄 Aplicar no SharePoint", use_container_width=True)

            marcar_coord = st.checkbox("✅ Marcar todos os processos visíveis abaixo", value=False)

            if df_coord.empty:
                st.warning("⚠️ Nenhum processo encontrado para os filtros selecionados ou a base de processos retornou vazia.")
            else:
                df_coord_exib = pd.DataFrame({
                    "Selecionar": pd.Series([marcar_coord] * len(df_coord), dtype=bool),
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

                if btn_atualizar:
                    indices_marcados = tabela_coord_editada[tabela_coord_editada["Selecionar"] == True].index
                    if len(indices_marcados) == 0:
                        st.warning("⚠️ Marque pelo menos um processo na tabela abaixo.")
                    elif novo_servidor == "[ Não Alterar ]" and novo_fiscal == "[ Não Alterar ]" and nova_situacao == "[ Não Alterar ]":
                        st.info("💡 Escolha ao menos uma alteração nos menus acima.")
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

                        try:
                            url_planilha = st.secrets["sharepoint"]["url_planilha"]
                            with st.spinner("Atualizando registros no SharePoint..."):
                                resp = requests.post(url_planilha, json=payload_atualizacao, headers={"Content-Type": "application/json"})
                                resp.raise_for_status()
                            st.success(f"✅ Sucesso! {len(processos_alvo)} processos atualizados no SharePoint.")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao enviar atualização: {e}")

        # --- ABA 3: ATRIBUIÇÃO AUTOMÁTICA BALANCEADA ---
        with tab_auto:
            st.markdown("### 🎲 Sorteio e Distribuição Automática de Processos")
            st.caption("Distribuição balanceada considerando a carga ativa em aberto de cada servidor")

            col_a1, col_a2 = st.columns(2)

            with col_a1:
                st.markdown("**1. Parâmetros da Demanda**")
                tipo_tarefa = st.radio("Selecione a Tarefa a Distribuir:", ["Elaboração de Laudo (Análise Técnica)", "Lavratura de Auto (Fiscalização)"])
                bacias_disp = ["Todas as Bacias"] + sorted([b for b in df['bacia'].astype(str).unique() if b and b != "nan"])
                bacia_alvo = st.selectbox("Filtrar Bacia Sedimentar:", bacias_disp)

                if "Laudo" in tipo_tarefa:
                    df_pend = df[df['situacao'].astype(str).str.strip().str.lower() == 'fazer laudo']
                else:
                    df_pend = df[df['situacao'].astype(str).str.strip().str.lower() == 'autuar']

                if bacia_alvo != "Todas as Bacias":
                    df_pend = df_pend[df_pend['bacia'].astype(str) == bacia_alvo]

                qtd_disponivel = len(df_pend)
                st.info(f"📌 Processos pendentes encontrados: **{qtd_disponivel}**")
                qtd_distribuir = st.number_input("Quantidade de processos a atribuir:", min_value=1, max_value=max(1, qtd_disponivel), value=min(5, max(1, qtd_disponivel)))

            with col_a2:
                st.markdown("**2. Equipe Elegível Habilitada**")
                if not df_equipe.empty:
                    if "Laudo" in tipo_tarefa:
                        servidores_base = sorted(df_equipe[(df_equipe['is_analise'] == True) & (df_equipe['status'] == 'APROVADO')]['nome'].unique().tolist())
                    else:
                        servidores_base = sorted(df_equipe[(df_equipe['is_fiscalizacao'] == True) & (df_equipe['status'] == 'APROVADO')]['nome'].unique().tolist())
                else:
                    if "Laudo" in tipo_tarefa:
                        servidores_base = sorted([s for s in df['s_laudo_limpo'].unique() if s != "Não Atribuído"])
                    else:
                        servidores_base = sorted([f for f in df['f_limpo'].unique() if f != "Não Atribuído"])

                servidores_selecionados = st.multiselect("Servidores que participarão do sorteio:", servidores_base, default=servidores_base)

            st.write("---")

            if st.button("🎲 Executar Simulação de Distribuição Balanceada", use_container_width=True):
                if qtd_disponivel == 0:
                    st.error("Não há processos pendentes disponíveis para os critérios selecionados.")
                elif not servidores_selecionados:
                    st.error("Selecione pelo menos um servidor habilitado.")
                else:
                    df_alvo_dist = df_pend.head(qtd_distribuir).copy()
                    
                    # Cálculo de carga ativa (processos pendentes sob responsabilidade)
                    if "Laudo" in tipo_tarefa:
                        cargas = {s: len(df[(df['s_laudo_limpo'] == s) & (df['situacao'].str.strip().str.lower() == 'fazer laudo') & (df['laudo_sei'].isin(["", "nan", "None", "0"]))]) for s in servidores_selecionados}
                    else:
                        cargas = {s: len(df[(df['f_limpo'] == s) & (df['situacao'].str.strip().str.lower() == 'autuar') & (df['auto'].isin(["", "nan", "None", "0", "processo não encontrado"]))]) for s in servidores_selecionados}

                    atribuicoes_resultado = []
                    for _, row in df_alvo_dist.iterrows():
                        menor_c = min(cargas.values())
                        candidatos = [s for s, c in cargas.items() if c == menor_c]
                        escolhido = random.choice(candidatos)
                        cargas[escolhido] += 1
                        atribuicoes_resultado.append(escolhido)

                    df_alvo_dist['Novo_Responsavel'] = atribuicoes_resultado
                    st.session_state['resultado_distribuicao'] = {"df": df_alvo_dist, "tipo": tipo_tarefa}

            if 'resultado_distribuicao' in st.session_state:
                res = st.session_state['resultado_distribuicao']
                df_res = res['df']
                tipo_t = res['tipo']

                st.markdown("### 📋 Prévia da Distribuição")
                st.dataframe(pd.DataFrame({
                    "ID": df_res['num_doc'].astype(str),
                    "Processo SEI": df_res['processo_sei'].astype(str),
                    "Bacia": df_res['bacia'].astype(str),
                    "Empresa": df_res['empresa'].astype(str),
                    "Novo Responsável Sorteado": df_res['Novo_Responsavel'].astype(str)
                }), hide_index=True, use_container_width=True)

                if st.button("🚀 Confirmar e Gravar Atribuições no SharePoint", use_container_width=True):
                    processos_alvo = [str(p) for p in df_res['processo_sei'].tolist()]
                    ids_alvo = [str(i) for i in df_res['num_doc'].tolist()]

                    for idx, row_dist in df_res.iterrows():
                        resp_sorteado = str(row_dist['Novo_Responsavel'])
                        payload_single = {
                            "acao": "ATUALIZAR",
                            "processos_sei": [str(row_dist['processo_sei'])],
                            "ids": [str(row_dist['num_doc'])],
                            "novo_servidor_laudo": resp_sorteado if "Laudo" in tipo_t else "",
                            "novo_fiscal": resp_sorteado if "Auto" in tipo_t else "",
                            "nova_situacao": ""
                        }
                        try:
                            requests.post(st.secrets["sharepoint"]["url_planilha"], json=payload_single, headers={"Content-Type": "application/json"})
                        except: pass

                    st.success(f"✅ {len(processos_alvo)} processos gravados com sucesso!")
                    del st.session_state['resultado_distribuicao']
                    st.cache_data.clear()
                    st.rerun()

        # --- ABA 4: GESTÃO DA EQUIPE ---
        with tab_equipe:
            st.markdown("### 👥 Integrantes da Força-Tarefa")
            st.caption("Visualização das permissões dos servidores (as senhas não são exibidas)")
            
            if not df_equipe.empty:
                df_eq_view = pd.DataFrame({
                    "Nome": df_equipe['nome'],
                    "E-mail": df_equipe['email'],
                    "Status": df_equipe['status'],
                    "Coordenação": df_equipe['is_coordenacao'].apply(lambda x: "SIM" if x else "NÃO"),
                    "Análise Técnica": df_equipe['is_analise'].apply(lambda x: "SIM" if x else "NÃO"),
                    "Fiscalização": df_equipe['is_fiscalizacao'].apply(lambda x: "SIM" if x else "NÃO"),
                })
                st.dataframe(df_eq_view, hide_index=True, use_container_width=True)
            else:
                st.info("Nenhum integrante cadastrado na tabela Equipe.")

    # =========================================================================
    # 🔬 MÓDULO 2: ANÁLISE TÉCNICA (INSTRUÇÃO DE LAUDOS)
    # =========================================================================
    elif pagina == "🔬 Análise Técnica":
        st.title("🔬 Análise Técnica — Instrução de Laudos")
        st.caption("Acompanhamento da elaboração de laudos técnicos e consolidação de evidências")
        
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

        df_metricas_laudo = df.copy()
        if sel_serv and "Todos" not in sel_serv:
            df_metricas_laudo = df_metricas_laudo[df_metricas_laudo['s_laudo_limpo'].astype(str).isin(sel_serv)]

        total_analise = len(df_metricas_laudo)
        laudos_pendentes = len(df_metricas_laudo[df_metricas_laudo['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])])
        laudos_concluidos = total_analise - laudos_pendentes
        
        m1, m2, m3 = st.columns(3)
        with m1: st.metric(label="Processos na Esteira", value=total_analise)
        with m2: st.metric(label="Pendentes de Laudo SEI", value=laudos_pendentes)
        with m3: st.metric(label="Laudos Elaborados", value=laudos_concluidos)
            
        st.write("---")

        df_laudo = df.copy()
        if sel_situ: df_laudo = df_laudo[df_laudo['situacao'].astype(str).isin(sel_situ)]
        if sel_serv and "Todos" not in sel_serv: df_laudo = df_laudo[df_laudo['s_laudo_limpo'].astype(str).isin(sel_serv)]
        if apenas_pendentes: df_laudo = df_laudo[df_laudo['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])]
            
        df_laudo = df_laudo.reset_index(drop=True)
        st.markdown(f"### 📋 Processos em Análise Técnica ({len(df_laudo)} encontrados)")
        
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
        st.dataframe(df_laudo_exib, hide_index=True, use_container_width=True)

    # =========================================================================
    # ⚖️ MÓDULO 3: FISCALIZAÇÃO (AUTUAÇÃO & MINUTAS)
    # =========================================================================
    elif pagina == "⚖️ Fiscalização":
        st.title("🔄 FiscFlow — Módulo de Fiscalização")
        st.markdown("### ⚖️ Gestão de Fila e Automação de Relatórios — IBAMA")
        st.caption("Sincronização ativa com o SharePoint | Geração de minutas em lote")

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
        with m1: st.metric(label="Processos no Fluxo", value=total_fila)
        with m2: st.metric(label="Prontos p/ Autuação", value=prontos_autuacao)
        with m3: st.metric(label="Autos Lavrados", value=autos_lavrados)
        with m4: st.metric(label="Proc. Fisc. Pendentes", value=fisc_pendentes)
        with m5: st.metric(label="Proc. AI Gerados", value=fisc_gerados)
        
        st.write("---")

        df_f = df.copy()
        if sel_situ: df_f = df_f[df_f['situacao'].astype(str).isin(sel_situ)]
        if sel_fisc and "Todos" not in sel_fisc: df_f = df_f[df_f['f_limpo'].astype(str).isin(sel_fisc)]
        if not todos_laudos: df_f = df_f[~df_f['laudo_sei'].astype(str).str.strip().isin(["", "nan", "None"])]

        df_f = df_f.reset_index(drop=True)
        st.markdown("### 📋 Processos para Análise")
        
        marcar_todos = st.checkbox("✅ Marcar todos os processos mostrados abaixo", value=False)
        vetor_selecao_inicial = [marcar_todos] * len(df_f)

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
            column_config={"Selecionar": st.column_config.CheckboxColumn("Selecionar")},
            key=f"editor_{marcar_todos}"
        )
        
        indices_selecionados = tabela_editada[tabela_editada["Selecionar"] == True].index
        selecionados = df_f.iloc[indices_selecionados]

        if not selecionados.empty:
            st.write("---")
            st.subheader(f"🚀 Geração em Lote ({len(selecionados)} itens)")
            
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
