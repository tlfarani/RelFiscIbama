# --- PALETA DE CORES INTEGRAL (CSS BLINDADO CONTRA DARK MODE) ---
st.markdown("""
    <style>
    /* 1. Forçar o Fundo Geral do App para Cinza Claro */
    .stApp {
        background-color: #F8F9F9 !important;
    }
    
    /* 2. Títulos, Subtítulos e Rótulos de Filtros em Verde Musgo */
    h1, h2, h3, .stSubheader, p, span, label, [data-testid="stWidgetLabel"] p {
        color: #4E5D30 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
    }

    /* 3. Forçar o Fundo das Caixas de Seleção (Filtros) para Branco com texto Preto */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border-color: #E9EDDE !important;
    }
    
    /* Input interno do texto digitado nos filtros */
    div[data-baseweb="select"] input {
        color: #000000 !important;
    }
    
    /* Menu suspenso de opções dos filtros */
    ul[role="listbox"] {
        background-color: #FFFFFF !important;
    }
    ul[role="listbox"] li {
        color: #000000 !important;
    }

    /* 4. Customização das tags internas escolhidas no Multiselect */
    span[data-baseweb="tag"] {
        background-color: #E9EDDE !important;
        color: #4E5D30 !important;
        border: 1px solid #4E5D30 !important;
    }

    /* 5. Forçar a visibilidade do Texto do Botão de Geração */
    div.stButton > button:first-child {
        background-color: #4E5D30 !important;
        color: #FFFFFF !important; /* Força o texto a ficar Branco */
        border-radius: 8px !important;
        border: 1px solid #4E5D30 !important;
        padding: 10px 24px !important;
        font-weight: bold !important;
    }
    div.stButton > button:first-child p {
        color: #FFFFFF !important; /* Garante que o parágrafo interno do botão seja branco */
    }
    div.stButton > button:first-child:hover {
        background-color: #3A471E !important;
        border-color: #3A471E !important;
    }

    /* 6. Texto do Checkbox Lateral */
    div[data-testid="stCheckbox"] span {
        color: #4E5D30 !important;
    }
    </style>
""", unsafe_allow_html=True)
