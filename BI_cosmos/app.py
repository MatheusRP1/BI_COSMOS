import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import google.generativeai as genai
import re
from streamlit_option_menu import option_menu
import base64 
import pathlib # Adicionado para corrigir o caminho da logo

st.set_page_config(
    page_title="Studio Cosmos - Análise de Viabilidade",
    page_icon="🌍",
    layout="wide"
)

st.markdown(
    """
    <style>
    :root {
        --color-primary: #9B59B6; 
        --color-secondary: #C2185B; 
        --color-tertiary: #F1C40F; 
        
        --color-background: #0E1117; 
        --color-sidebar: #1E1E2E; 
        --color-container: #262730; 
        
        --color-text-primary: #FAFAFA; 
        --color-text-secondary: #FAFAFA; /* AJUSTE: Alterado de #A0A0A0 para branco */
        --color-border: #3E3E3E; 
    }

    [data-testid="stAppViewContainer"] {
        background-color: var(--color-background) !important;
        color: var(--color-text-primary) !important;
    }
    .main { 
        background-color: var(--color-background) !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: var(--color-sidebar); 
        border-right: 1px solid var(--color-border);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .st-bq,
    [data-testid="stSidebar"] li {
        color: var(--color-text-primary) !important;
    }

    [data-testid="stMetric"], [data-testid="stExpander"], [data-testid="stDataFrame"], [data-testid="stAlert"] {
        background-color: var(--color-container); 
        border: 1px solid var(--color-border);
        border-radius: 8px;
    }
    
    [data-testid="stMetric"] {
        padding: 15px;
        color: var(--color-text-primary) !important; 
    }
    [data-testid="stMetricLabel"] {
        color: var(--color-text-secondary) !important; /* Agora usa o secundário (branco) */
    }
    
    /* --- CORREÇÃO DE COR AQUI --- */
    [data-testid="stExpander"] summary {
        font-weight: bold;
        color: var(--color-text-primary) !important; /* AJUSTE: Alterado de rosa para branco */
    }
    /* --- FIM DA CORREÇÃO --- */

    /* --- HEADER MAIOR --- */
    .header {
        display: flex;
        align-items: center;
        justify-content: center; 
        padding-bottom: 20px;
    }
    .header img {
        width: 100px; 
        height: 100px; 
        margin-right: 25px;
    }
    .header .titles {
        display: flex;
        flex-direction: column;
    }
    .header h1 {
        margin: 0;
        font-size: 3.0em; 
        color: var(--color-text-primary);
    }
    .header h2 {
        margin: 0;
        font-size: 1.4em; 
        color: var(--color-text-secondary); /* Agora usa o secundário (branco) */
    }
    /* --- FIM DO HEADER MAIOR --- */

    div[data-testid="stHorizontalBlock"] > div[data-testid^="element-container-"] > div[data-testid^="stOptionMenu-"] {
        background-color: var(--color-container);
        border: 1px solid var(--color-border);
        border-radius: 8px;
        padding: 5px !important;
        margin-bottom: 10px; 
    }
    div[data-testid^="stOptionMenu-"] > nav > ul > li > a {
        color: var(--color-text-secondary) !important; /* Agora usa o secundário (branco) */
        padding: 10px;
        border-radius: 6px;
    }
    div[data-testid^="stOptionMenu-"] > nav > ul > li > a:hover {
        background-color: var(--color-border);
        color: var(--color-text-primary) !important;
    }
    div[data-testid^="stOptionMenu-"] > nav > ul > li > a.active {
        background-color: var(--color-primary); 
        color: var(--color-text-primary) !important;
    }
    div[data-testid^="stOptionMenu-"] > nav > ul > li > a.active > i {
        color: var(--color-text-primary) !important;
    }
    div[data-testid^="stOptionMenu-"] > nav > ul > li > a:not(.active) > i {
        color: var(--color-primary) !important; 
    }
    </style>
    """,
    unsafe_allow_html=True
)

def encontrar_coluna(df, nomes_possiveis):
    for nome in nomes_possiveis:
        nome_limpo = nome.lower()
        for col in df.columns:
            col_limpa = str(col).lower()
            if nome_limpo in col_limpa:
                return col 
    return None

def extrair_dado_visita(df, aspecto_procurado, col_aspecto, col_obs):
    try:
        resultado = df[df[col_aspecto].str.contains(aspecto_procurado, case=False, na=False)]
        if not resultado.empty:
            valor = resultado.iloc[0][col_obs]
            if pd.isna(valor) or str(valor).strip() == "":
                return "N/A"
            return str(valor)
    except Exception as e:
        pass
    return "N/A"

def extrair_primeiro_numero(texto):
    if not isinstance(texto, str):
        return None
    texto_limpo = str(texto).replace('.', '')
    texto_limpo = texto_limpo.replace(',', '.')
    
    match = re.search(r'(\d+\.?\d*)', texto_limpo)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None

def clean_str(s):
    return str(s).lower().strip()

@st.cache_data
def carregar_dados_excel(ficheiro_carregado):
    try:
        try:
            ficheiro_carregado.seek(0)
            xls = pd.ExcelFile(ficheiro_carregado)
            nomes_das_abas = xls.sheet_names
        except Exception as e:
            st.error(f"Erro ao ler a estrutura do arquivo Excel: {e}")
            return None

        mapeamento = {
            'urbana': 'KPIs (Urbana)',
            'ambiental': 'KPIs (Ambiental)',
            'social': 'KPIs (Social)',
            'economica': 'KPIs (Econômica)',
            'fisica': 'KPIs (Física)',
            'sensorial': 'KPIs (Sensorial)',
            'leg_ade': 'KPIs (Legislativa) - ADE',
            'leg_zr3': 'KPIs (Legislativa) - ZR3',
            'visita': 'Dados de campo (Relatório)',
            'matriz': 'Matriz, pesos e índices',
            'resumo_analitico': 'Resumo analítico' 
        }

        abas_encontradas = {}
        
        mapa_nomes_reais = {}
        for chave, nome_parcial in mapeamento.items():
            nome_encontrado = next((nome_aba for nome_aba in nomes_das_abas if nome_parcial.lower() in nome_aba.lower()), None)
            if nome_encontrado:
                mapa_nomes_reais[chave] = nome_encontrado
            else:
                abas_encontradas[chave] = pd.DataFrame()
                st.warning(f"Aviso: Não foi possível encontrar a aba que contém '{nome_parcial}'")

        chaves_para_ler = [chave for chave in mapa_nomes_reais if chave != 'economica']
        nomes_reais_para_ler = [mapa_nomes_reais[chave] for chave in chaves_para_ler]
        
        if nomes_reais_para_ler:
            try:
                ficheiro_carregado.seek(0)
                dfs_normais = pd.read_excel(ficheiro_carregado, sheet_name=nomes_reais_para_ler)
                for i, chave in enumerate(chaves_para_ler):
                    abas_encontradas[chave] = dfs_normais[nomes_reais_para_ler[i]]
            except Exception as e:
                st.error(f"Erro ao ler abas normais: {e}")

        if 'economica' in mapa_nomes_reais:
            try:
                ficheiro_carregado.seek(0)
                df_eco_normal = pd.read_excel(ficheiro_carregado, sheet_name=mapa_nomes_reais['economica'])
                abas_encontradas['economica'] = df_eco_normal
            except Exception as e:
                st.error(f"Erro ao ler a aba 'Econômica' (leitura normal): {e}")
                abas_encontradas['economica'] = pd.DataFrame()
        

        metricas = {}
        
        for chave in ['fisica', 'social', 'urbana', 'ambiental', 'economica', 'sensorial']:
            df = abas_encontradas[chave]
            if df.empty:
                metricas[f'media_{chave}'] = 0
                continue
                
            col_escala = encontrar_coluna(df, ['ESCALA (0–5)', 'ESCALA'])
            if col_escala:
                df[col_escala] = pd.to_numeric(df[col_escala], errors='coerce')
                df = df.dropna(subset=[col_escala])
                abas_encontradas[chave] = df 
                metricas[f'media_{chave}'] = df[col_escala].mean()
            else:
                metricas[f'media_{chave}'] = 0
        
        df_matriz = abas_encontradas['matriz']
        metricas['it_total'] = 0.0
        
        if not df_matriz.empty:
            try:
                col_dimensao = encontrar_coluna(df_matriz, ['DIMENSÃO'])
                if not col_dimensao: col_dimensao = df_matriz.columns[0]
                
                it_row = df_matriz[df_matriz[col_dimensao].str.contains('Índice Territorial', na=False, case=False)]
                
                if not it_row.empty:
                    col_it_valor = encontrar_coluna(df_matriz, ['ESCALA (0–5)', 'ESCALA'])
                    if not col_it_valor: col_it_valor = df_matriz.columns[2]
                    
                    valor_it = it_row.iloc[0][col_it_valor]
                    metricas['it_total'] = pd.to_numeric(valor_it, errors='coerce')
                else:
                    st.warning("Não foi possível encontrar a linha 'Índice Territorial' na aba Matriz.")
            
            except Exception as e:
                st.error(f"Erro ao processar a aba 'Matriz': {e}")

        abas_encontradas['metricas'] = metricas
        return abas_encontradas

    except Exception as e:
        st.error(f"Erro Crítico ao processar o Excel: {e}")
        st.error("Verifique se o formato do arquivo corresponde ao template original e não está corrompido.")
        return None

def criar_pagina_dimensao(nome_dimensao, df_dimensao, mapas_carregados):
    if df_dimensao.empty:
        st.warning(f"Dados da dimensão '{nome_dimensao}' não encontrados. Verifique a aba correspondente no Excel.")
        return

    col_escala = encontrar_coluna(df_dimensao, ['ESCALA (0–5)', 'ESCALA'])
    col_indicador = encontrar_coluna(df_dimensao, ['INDICADOR'])
    col_analise = encontrar_coluna(df_dimensao, ['ANÁLISE', 'ANALISE'])
    col_projeto = encontrar_coluna(df_dimensao, ['RELAÇÃO COM O PROJETO', 'RELAÇÃO'])
    col_mapa = encontrar_coluna(df_dimensao, ['MAPA CORRESPONDENTE', 'MAPA'])

    if col_escala and col_indicador:
        st.subheader("Perfil da Dimensão")
        df_grafico = df_dimensao.dropna(subset=[col_escala, col_indicador])
        
        if df_grafico.empty:
            st.warning(f"Não há dados válidos para o gráfico de perfil da dimensão '{nome_dimensao}'.")
        else:
            fig = px.bar(
                df_grafico,
                x=col_indicador,
                y=col_escala,
                color=col_escala,
                color_continuous_scale=px.colors.sequential.Blues, # AJUSTE: Alterado de Magma para Blues
                text=col_escala,
                title=f"Notas dos indicadores - {nome_dimensao.capitalize()}",
                range_y=[0, 6], 
                template="plotly_dark" 
            )
            
            fig.update_traces(
                texttemplate='%{y:.1f}', 
                textposition='outside',
            )
            fig.update_layout(
                yaxis_title="Escala (0-5)", 
                xaxis_title="Indicador",
                plot_bgcolor='var(--color-container)',
                paper_bgcolor='var(--color-container)'
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning(f"Não foi possível gerar o gráfico de perfil para '{nome_dimensao}'. Colunas 'INDICADOR' ou 'ESCALA' não encontradas.")
    
    st.divider()

    st.subheader("Análise Detalhada dos Indicadores")
    if col_analise and col_projeto and col_indicador and col_escala:
        df_analise = df_dimensao.dropna(subset=[col_indicador, col_escala])
        
        for index, row in df_analise.iterrows():
            titulo = f"**{row[col_indicador]}** (Nota: {row[col_escala]:.1f})"
            with st.expander(titulo):
                
                nomes_dos_mapas_str = None
                if col_mapa and col_mapa in row and pd.notna(row[col_mapa]):
                    nomes_dos_mapas_str = str(row[col_mapa])
                
                if nomes_dos_mapas_str:
                    lista_de_mapas_excel = nomes_dos_mapas_str.split(',')
                    mapas_encontrados_count = 0
                    
                    for nome_excel_sujo in lista_de_mapas_excel:
                        nome_excel_limpo = nome_excel_sujo.strip().lower() 
                        if not nome_excel_limpo: continue 

                        mapa_encontrado_nome_real = None
                        for nome_arquivo_upload in mapas_carregados.keys():
                            if nome_excel_limpo in nome_arquivo_upload.lower():
                                mapa_encontrado_nome_real = nome_arquivo_upload
                                break 
                        
                        if mapa_encontrado_nome_real:
                            st.image(
                                mapas_carregados[mapa_encontrado_nome_real],
                                caption=f"Mapa: {mapa_encontrado_nome_real}",
                                use_container_width=True
                            )
                            mapas_encontrados_count += 1
                        else:
                            st.warning(f"O mapa '{nome_excel_limpo}' foi referenciado, mas um arquivo correspondente (ex: '{nome_excel_limpo}.png') não foi encontrado nos uploads.")
                    
                    if mapas_encontrados_count > 0:
                        st.markdown("---") 

                if pd.notna(row[col_analise]):
                    st.markdown("#### Análise")
                    st.write(row[col_analise])
                
                if pd.notna(row[col_projeto]):
                    st.markdown("#### Relação com o Projeto")
                    st.write(row[col_projeto])
    else:
        st.error("Não foi possível encontrar as colunas 'Indicador', 'Análise' ou 'Relação' no Excel.")
        st.dataframe(df_dimensao)

# --- INÍCIO DAS NOVAS FUNÇÕES LEGISLATIVAS ---
def find_header_row(df_raw, keywords):
    for i, row in df_raw.head(30).iterrows():
        row_values = [clean_str(val) for val in row.values]
        has_all_keywords = True
        for keyword in keywords:
            if keyword not in row_values:
                has_all_keywords = False
                break
        if has_all_keywords:
            return i
    return -1

def processar_tabela_parametros(df_raw):
    keywords = ['indicador', 'valor indicado']
    header_row_idx = find_header_row(df_raw, keywords)
    
    if header_row_idx == -1:
        return pd.DataFrame() 

    df_processado = df_raw.loc[header_row_idx:].copy()
    new_cols = [clean_str(col) if pd.notna(col) else f"unnamed_{j}" for j, col in enumerate(df_processado.iloc[0])]
    df_processado.columns = new_cols
    df_processado = df_processado.iloc[1:].reset_index(drop=True)
    
    stop_keyword = 'usos'
    stop_row_mask = df_processado.apply(lambda row: row.astype(str).str.contains(stop_keyword, case=False, na=False).any(), axis=1)
    
    if stop_row_mask.any():
        stop_row_idx = stop_row_mask.idxmax()
        df_processado = df_processado.loc[:stop_row_idx-1]
        
    df_processado = df_processado.dropna(how='all')
    return df_processado

def processar_tabela_usos(df_raw):
    header_row_idx = -1
    col_usos_str = 'usos'
    col_adeq_str = 'adequação'
    col_adeq_str_alt = 'adequacao'

    for i, row in df_raw.head(30).iterrows():
        row_values = [clean_str(val) for val in row.values]
        has_usos = col_usos_str in row_values
        has_adeq = any(col_adeq_str in val or col_adeq_str_alt in val for val in row_values)
        
        if has_usos and has_adeq:
            header_row_idx = i
            break
    
    if header_row_idx == -1:
        return pd.DataFrame() 

    df_processado = df_raw.loc[header_row_idx:].copy()
    new_cols = [clean_str(col) if pd.notna(col) else f"unnamed_{j}" for j, col in enumerate(df_processado.iloc[0])]
    df_processado.columns = new_cols
    df_processado = df_processado.iloc[1:].reset_index(drop=True)

    col_adeq_final = encontrar_coluna(df_processado, [col_adeq_str, col_adeq_str_alt])
    col_usos_final = encontrar_coluna(df_processado, [col_usos_str])
    col_indicador_final = encontrar_coluna(df_processado, ['indicador'])


    if col_adeq_final and col_usos_final:
        # Substitui vazios e NaN por "Inadequado"
        df_processado[col_adeq_final] = df_processado[col_adeq_final].replace(r'^\s*$', np.nan, regex=True)
        df_processado[col_adeq_final] = df_processado[col_adeq_final].fillna('Inadequado')
        
        # Remove linhas onde a coluna 'USOS' é vazia
        df_processado = df_processado.dropna(subset=[col_usos_final])
        
        # Remove linhas que são resquícios do cabeçalho ou títulos
        df_processado = df_processado[~df_processado[col_usos_final].str.contains(col_usos_str, case=False, na=False)]
        if col_indicador_final:
             df_processado = df_processado[~df_processado[col_indicador_final].str.contains('adequação dos usos', case=False, na=False)]
    
    df_processado = df_processado.dropna(how='all')
    return df_processado
# --- FIM DAS NOVAS FUNÇÕES LEGISLATIVAS ---


# --- AJUSTE DE CAMINHO DA LOGO (INÍCIO) ---
# Define o caminho para a pasta onde o script app.py está
SCRIPT_DIR = pathlib.Path(__file__).parent
# Define o caminho completo para a logo
LOGO_PATH = SCRIPT_DIR / "logo.jpg"

logo_src = "https://raw.githubusercontent.com/streamlit/templates/main/multipage-apps/assets/dialogue.png"
logo_style_override = "filter: brightness(0) invert(1);" 

try:
    # Use o caminho completo (LOGO_PATH) em vez de só "logo.jpg"
    with open(LOGO_PATH, "rb") as f:
        bytes_data = f.read()
        base64_str = base64.b64encode(bytes_data).decode()
        logo_src = f"data:image/jpeg;base64,{base64_str}"
        logo_style_override = "filter: none;" 
except FileNotFoundError:
    st.sidebar.warning("Arquivo 'logo.jpg' não encontrado. Usando logo padrão.")
except Exception as e:
    st.sidebar.error(f"Erro ao carregar 'logo.jpg': {e}. Usando logo padrão.")
# --- AJUSTE DE CAMINHO DA LOGO (FIM) ---


st.markdown(
    f"""
    <div class="header">
        <img src="{logo_src}" alt="Logo Studio Cosmos" style="{logo_style_override}">
        <div class="titles">
            <h1>Studio Cosmos</h1>
            <h2>Análise de Viabilidade Territorial</h2>
        </div>
    </div>
    """, 
    unsafe_allow_html=True
)

st.sidebar.title("Configuração")

uploaded_file = st.sidebar.file_uploader(
    "1. Carregue o arquivo Excel (KPIs)",
    type=["xlsx"]
)

mapa_files = st.sidebar.file_uploader(
    "2. Carregue os Mapas (PNG, JPG)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

mapas_carregados = {}
if mapa_files:
    for mapa_file in mapa_files:
        mapas_carregados[mapa_file.name] = mapa_file
    st.sidebar.success(f"{len(mapas_carregados)} mapas carregados.", icon="✅")


if uploaded_file is None:
    st.info("Por favor, carregue o arquivo Excel de análise na barra lateral para começar.")
    st.stop()

dados = carregar_dados_excel(uploaded_file)
if dados is None:
    st.stop()

pagina_selecionada = option_menu(
    menu_title=None,
    options=[
        "Resumo Geral",
        "Dimensões",
        "Análise Legislativa",
        "Relatório de Visita",
        "Estratégia e Riscos",
        "🤖 IA Chatbot"
    ],
    icons=[
        "pie-chart-fill",
        "grid-1x2-fill",
        "building",
        "clipboard-data",
        "lightbulb-fill",
        "robot"
    ],
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "transparent"},
        "nav-link-selected": {"background-color": "var(--color-primary)"},
    }
)

dimensao_selecionada = None

if pagina_selecionada == "Resumo Geral":
    st.header("Resumo Geral da Análise")
    
    col1, col2 = st.columns(2)
    
    with col1:
        it_valor = dados['metricas'].get('it_total', 0.0)
        st.metric("Índice Territorial (IT)", f"{it_valor:.2f}", 
            "Alto Potencial" if it_valor > 3.5 else "Potencial Moderado")
    
    with col2:
        media_sensorial = dados['metricas'].get('media_sensorial', 0.0)
        delta_sensorial = media_sensorial - 3.0
        st.metric("Média Sensorial (Ponto de Atenção)", f"{media_sensorial:.2f}", 
            f"{delta_sensorial:.2f} vs. Meta (3.0)",
            delta_color="inverse" if media_sensorial < 3 else "normal")

    st.divider()
    
    st.subheader("Desempenho por Dimensão (Média 0-5)")
    st.caption("✔️ Este gráfico é gerado dinamicamente a partir das médias das abas.")
    
    metricas_df = pd.DataFrame({
        'Dimensão': [
            'Física', 'Econômica', 'Social', 'Urbana', 'Ambiental', 'Sensorial'
        ],
        'Média': [
            dados['metricas']['media_fisica'],
            dados['metricas']['media_economica'],
            dados['metricas']['media_social'],
            dados['metricas']['media_urbana'],
            dados['metricas']['media_ambiental'],
            dados['metricas']['media_sensorial']
        ]
    }).sort_values(by='Média', ascending=False)
    
    fig = px.bar(
        metricas_df,
        x='Média',
        y='Dimensão',
        orientation='h',
        text='Média',
        color='Média',
        color_continuous_scale=px.colors.sequential.Viridis, # AJUSTE: Alterado de Magma para Viridis (azul/verde)
        range_x=[0, 5.5], 
        template="plotly_dark"
    )
    
    fig.update_traces(texttemplate='%{x:.2f}', textposition='outside')
    fig.update_layout(
        yaxis_title="", 
        xaxis_title="Média (0-5)", 
        showlegend=False,
        coloraxis_showscale=False,
        plot_bgcolor='var(--color-container)',
        paper_bgcolor='var(--color-container)'
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Principais Drivers de Impacto (da Matriz de Pesos)")
    df_matriz = dados['matriz']
    
    if not df_matriz.empty:
        col_indicador = encontrar_coluna(df_matriz, ['INDICADOR'])
        col_vp = encontrar_coluna(df_matriz, ['VALOR PONDERADO'])
        
        if col_indicador and col_vp:
            df_drivers = df_matriz.dropna(subset=[col_indicador, col_vp])
            df_drivers = df_drivers[~df_drivers[col_indicador].str.contains('Índice|Interpretação', na=False, case=False)]
            
            col_escala = encontrar_coluna(df_matriz, ['ESCALA'])
            col_peso = encontrar_coluna(df_matriz, ['PESO'])
            cols_to_show = [col_indicador, col_escala, col_peso, col_vp]
            cols_existentes = [col for col in cols_to_show if col in df_drivers.columns]
            
            df_drivers_show = df_drivers[cols_existentes].copy()
            df_drivers_show[col_vp] = pd.to_numeric(df_drivers_show[col_vp], errors='coerce')
            
            st.dataframe(
                df_drivers_show.sort_values(by=col_vp, ascending=False),
                use_container_width=True
            )
        else:
            st.warning("Não foi possível encontrar as colunas 'Indicador' e 'Valor Ponderado' na aba 'Matriz'.")
            st.dataframe(df_matriz)
    else:
        st.error("Aba '3) Matriz, pesos e índices' não encontrada ou está vazia.")

elif pagina_selecionada == "Dimensões":
    dimensao_selecionada = option_menu(
        menu_title=None,
        options=["Urbana", "Ambiental", "Social", "Econômica", "Física", "Sensorial"],
        icons=["bi-building", "bi-tree", "bi-people", "bi-cash-coin", "bi-rulers", "bi-mic"],
        orientation="horizontal",
        styles={
            "container": {
                "padding": "0!important", 
                "background-color": "transparent", 
                "margin-bottom": "20px",
                "border-bottom": f"2px solid var(--color-border)"
            },
            "nav-link": {
                "color": "var(--color-text-secondary)", 
                "--hover-color": "var(--color-container)",
                "border-bottom": "2px solid transparent",
                "padding": "10px 0"
            },
            "nav-link-selected": {
                "background-color": "transparent", 
                "color": "var(--color-secondary)", 
                "border-bottom": f"2px solid var(--color-secondary)"
            },
            "icon": {"display": "none"}
        }
    )

    st.header(f"Dimensão: {dimensao_selecionada.capitalize()}")

    if dimensao_selecionada == "Urbana":
        criar_pagina_dimensao("Urbana", dados['urbana'], mapas_carregados)
    elif dimensao_selecionada == "Ambiental":
        criar_pagina_dimensao("Ambiental", dados['ambiental'], mapas_carregados)
    elif dimensao_selecionada == "Social":
        criar_pagina_dimensao("Social", dados['social'], mapas_carregados)
    elif dimensao_selecionada == "Física":
        criar_pagina_dimensao("Física", dados['fisica'], mapas_carregados)
    elif dimensao_selecionada == "Sensorial":
        criar_pagina_dimensao("Sensorial", dados['sensorial'], mapas_carregados)
        
    elif dimensao_selecionada == "Econômica":
        criar_pagina_dimensao("Econômica", dados['economica'], mapas_carregados)
        
        st.divider()
        st.subheader("Análise de Parceiros Potenciais (Stakeholders)")
        st.caption("Dados extraídos dinamicamente da aba 'KPIs (Econômica)'")
        
        df_eco_raw = pd.DataFrame()
        nome_real_economica = None
        
        try:
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                if "kpis (econômica)" in sheet_name.lower():
                    nome_real_economica = sheet_name
                    break
        except Exception as e:
            st.error(f"Erro ao re-abrir o Excel para encontrar a aba: {e}")

        if not nome_real_economica:
            st.error("Não foi possível encontrar o nome da aba 'KPIs (Econômica)' no arquivo.")
        else:
            try:
                uploaded_file.seek(0)
                df_eco_raw = pd.read_excel(
                    uploaded_file, 
                    sheet_name=nome_real_economica, 
                    header=None 
                )
            except Exception as e:
                st.error(f"Erro ao re-ler a aba 'Econômica' com header=None: {e}")

        if df_eco_raw.empty:
            st.warning("Aba 'Econômica' está vazia ou não pôde ser lida como dados brutos.")
        else:
            header_row_index = -1
            df_stakeholders = pd.DataFrame()

            for i, row in df_eco_raw.head(30).iterrows():
                row_values = [clean_str(val) for val in row.values]
                has_instituicao = any('instituição' in val or 'instituicao' in val for val in row_values)
                has_potencial = any('potencial' in val for val in row_values)
                
                if has_instituicao and has_potencial:
                    header_row_index = i
                    break
                        
            if header_row_index == -1:
                st.error("Não foi possível encontrar a linha de cabeçalho (INSTITUIÇÃO, POTENCIAL) na aba 'Econômica', mesmo lendo os dados brutos.")
            
            else:
                df_stakeholders = df_eco_raw.loc[header_row_index:].copy()
                
                header_row_values = df_stakeholders.iloc[0]
                new_cols = []
                counts = {}
                
                for col in header_row_values:
                    col_name = clean_str(col) 
                    
                    if col_name == "nan" or col_name == "":
                        col_name = "unnamed" 
                    
                    if col_name in counts:
                        counts[col_name] += 1
                        new_cols.append(f"{col_name}_{counts[col_name]}")
                    else:
                        counts[col_name] = 0
                        new_cols.append(col_name)
                        
                df_stakeholders.columns = new_cols
                df_stakeholders = df_stakeholders.iloc[1:]
                
                col_inst_nome = encontrar_coluna(df_stakeholders, ['instituição', 'instituicao']) 
                col_pot_nome = encontrar_coluna(df_stakeholders, ['potencial'])
                col_loc_nome = encontrar_coluna(df_stakeholders, ['localização', 'localizacao'])

                if col_inst_nome and col_pot_nome and col_loc_nome:
                    df_stakeholders = df_stakeholders.dropna(subset=[col_inst_nome, col_pot_nome, col_loc_nome])
                    
                    if df_stakeholders.empty:
                            st.info("A tabela de stakeholders foi encontrada, mas está vazia.")
                    else:
                        
                        st.markdown("#### Análise Detalhada por Parceiro")
                        st.caption("Clique em um parceiro para ver os detalhes.")
                        
                        mapeamento_potencial = {'Alto': 3, 'Médio': 2, 'Baixo': 1}
                        
                        df_stakeholders['Potencial_Nivel_Temp'] = 'Baixo' 
                        if col_pot_nome in df_stakeholders.columns:
                            df_stakeholders['Potencial_Nivel_Temp'] = df_stakeholders[col_pot_nome].astype(str).str.title()
                        
                        df_stakeholders['Potencial_Num'] = df_stakeholders['Potencial_Nivel_Temp'].map(mapeamento_potencial).fillna(0)
                        df_stakeholders = df_stakeholders.sort_values(by='Potencial_Num', ascending=False)


                        for index, row in df_stakeholders.iterrows():
                            nome_parceiro = row[col_inst_nome]
                            localizacao_parceiro = row[col_loc_nome]
                            
                            titulo_expander = f"{nome_parceiro} | {localizacao_parceiro}"
                            
                            with st.expander(titulo_expander):
                                col_potencial_texto = encontrar_coluna(df_stakeholders, ['potencial'])
                                potencial_texto = row[col_potencial_texto]

                                if pd.notna(potencial_texto) and str(potencial_texto).strip():
                                    topicos = [t.strip() for t in str(potencial_texto).split('.') if t.strip()]
                                    
                                    markdown_formatado = ""
                                    for topico in topicos:
                                        markdown_formatado += f"- {topico}.\n"
                                    
                                    if markdown_formatado:
                                        st.markdown(markdown_formatado)
                                    else:
                                        st.write(potencial_texto)
                                else:
                                    st.info("Nenhuma análise de potencial fornecida.")
                
                else:
                    st.error("Encontrei o cabeçalho, mas as colunas 'INSTITUIÇÃO', 'POTENCIAL' ou 'LOCALIZAÇÃO' parecem estar ausentes.")


elif pagina_selecionada == "Análise Legislativa":
    st.header("Análise Legislativa (ADE vs ZR3)")
    st.caption("Comparativo gráfico e de parâmetros-chave para as duas zonas.")

    df_ade_raw = pd.DataFrame()
    df_zr3_raw = pd.DataFrame()
    nome_real_ade = None
    nome_real_zr3 = None
    
    try:
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file)
        for sheet_name in xls.sheet_names:
            if "kpis (legislativa) - ade" in sheet_name.lower():
                nome_real_ade = sheet_name
            if "kpis (legislativa) - zr3" in sheet_name.lower():
                nome_real_zr3 = sheet_name
    except Exception as e:
        st.error(f"Erro ao re-abrir o Excel para encontrar as abas de legislação: {e}")

    if not nome_real_ade or not nome_real_zr3:
        st.error("Não foi possível encontrar as abas 'Legislativa - ADE' ou 'Legislativa - ZR3' no arquivo.")
        if 'st' in locals(): st.stop()
        else: exit()

    try:
        uploaded_file.seek(0)
        df_ade_raw = pd.read_excel(uploaded_file, sheet_name=nome_real_ade, header=None)
        uploaded_file.seek(0)
        df_zr3_raw = pd.read_excel(uploaded_file, sheet_name=nome_real_zr3, header=None)
    except Exception as e:
        st.error(f"Erro ao re-ler as abas de Legislação com header=None: {e}")

    if df_ade_raw.empty or df_zr3_raw.empty:
        st.warning("Abas de Legislação estão vazias ou não puderam ser lidas.")
        if 'st' in locals(): st.stop()
        else: exit()
    
    df_ade_params = processar_tabela_parametros(df_ade_raw)
    df_zr3_params = processar_tabela_parametros(df_zr3_raw)
    df_ade_usos = processar_tabela_usos(df_ade_raw)
    df_zr3_usos = processar_tabela_usos(df_zr3_raw)

    if df_ade_usos.empty or df_zr3_usos.empty:
        st.error("Não foi possível localizar a tabela (cabeçalho 'USOS' e 'ADEQUAÇÃO') nas abas de legislação.")
    else:
        col_usos_ade = encontrar_coluna(df_ade_usos, ['usos'])
        col_adeq_ade = encontrar_coluna(df_ade_usos, ['adequação', 'adequacao'])
        
        col_usos_zr3 = encontrar_coluna(df_zr3_usos, ['usos'])
        col_adeq_zr3 = encontrar_coluna(df_zr3_usos, ['adequação', 'adequacao'])

        if not (col_usos_ade and col_adeq_ade and col_usos_zr3 and col_adeq_zr3):
            st.error("Tabela de Usos encontrada, mas os nomes das colunas 'USOS' ou 'ADEQUAÇÃO' não puderam ser confirmados.")
        else:
            st.subheader("ADEQUAÇÃO DOS USOS ÀS ZONAS")
            col1, col2 = st.columns(2)

            def criar_cards_de_uso(df, col_usos, col_adeq):
                df_usos = df.dropna(subset=[col_usos])
                df_usos = df_usos.copy()
                
                df_usos['Categoria'] = df_usos[col_adeq].fillna('Inadequado').astype(str).str.strip().str.title()
                df_usos['Categoria'] = df_usos['Categoria'].replace(['Nan', 'Não Adequado', ''], 'Inadequado')

                
                adequados = df_usos[df_usos['Categoria'] == 'Adequado'][col_usos]
                proibidos = df_usos[df_usos['Categoria'] == 'Proibido'][col_usos]
                inadequados = df_usos[df_usos['Categoria'] == 'Inadequado'][col_usos]
                
                with st.expander(f"✅ Adequadas ({len(adequados)})"):
                    st.dataframe(adequados, use_container_width=True)
                    
                with st.expander(f"❌ Proibidas ({len(proibidos)})"):
                    st.dataframe(proibidos, use_container_width=True)
                    
                with st.expander(f"⚠️ Inadequadas ({len(inadequados)})"):
                    st.dataframe(inadequados, use_container_width=True)

            with col1:
                st.markdown("#### ADE (Local)")
                criar_cards_de_uso(df_ade_usos, col_usos_ade, col_adeq_ade)

            with col2:
                st.markdown("#### ZR3 (Entorno)")
                criar_cards_de_uso(df_zr3_usos, col_usos_zr3, col_adeq_zr3)

    st.divider()

    col_param_ade = encontrar_coluna(df_ade_params, ['indicador'])
    col_valor_ade = encontrar_coluna(df_ade_params, ['valor indicado'])
    col_param_zr3 = encontrar_coluna(df_zr3_params, ['indicador'])
    col_valor_zr3 = encontrar_coluna(df_zr3_params, ['valor indicado'])

    parametros_numericos_comuns = [
        'taxa de ocupação', 
        'coeficiente de aproveitamento', 
        'taxa de permeabilidade',
        'área mínima de lote',
        'testada mínima',
        'afast. frontal',
        'afast. lateral'
    ]
    dados_grafico = []

    if col_param_ade and col_valor_ade and col_param_zr3 and col_valor_zr3:
        for param_nome in parametros_numericos_comuns:
            val_ade = extrair_dado_visita(df_ade_params, param_nome, col_param_ade, col_valor_ade)
            val_zr3 = extrair_dado_visita(df_zr3_params, param_nome, col_param_zr3, col_valor_zr3)

            num_ade = extrair_primeiro_numero(val_ade)
            num_zr3 = extrair_primeiro_numero(val_zr3)

            param_nome_formatado = param_nome.replace('taxa de ', 'T. ').replace('coeficiente de ', 'C. ').replace('afast. ', 'A. ').replace('área mínima de ', 'Área Mín. ').title()

            if num_ade is not None:
                dados_grafico.append({"Parâmetro": param_nome_formatado, "Zoneamento": "ADE (Local)", "Valor": num_ade})
            if num_zr3 is not None:
                dados_grafico.append({"Parâmetro": param_nome_formatado, "Zoneamento": "ZR3 (Entorno)", "Valor": num_zr3})
        
        if dados_grafico:
            st.subheader("Comparativo de Parâmetros Numéricos")
            st.caption("Eixo Y em escala logarítmica para melhor visualização. Passe o mouse sobre as barras para ver os valores exatos.")
            df_plot = pd.DataFrame(dados_grafico)
            
            fig = px.bar(
                df_plot,
                x="Parâmetro",
                y="Valor",
                color="Zoneamento",
                barmode="group",
                template="plotly_dark", 
                color_discrete_map={ # AJUSTE: Alterado para azul e verde
                    "ADE (Local)": "#007BFF",  # Azul vibrante
                    "ZR3 (Entorno)": "#28A745" # Verde vibrante
                },
                title="Comparativo de Zoneamento: ADE (Local) vs. ZR3 (Entorno)",
                log_y=True,
                hover_data={"Valor": True}
            )
            
            fig.update_yaxes(title="Valor (Escala Log)")
            fig.update_xaxes(title=None) 
            fig.update_layout(
                legend_title="Zoneamento",
                plot_bgcolor='var(--color-container)',
                paper_bgcolor='var(--color-container)'
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Não foram encontrados parâmetros numéricos comuns (Ex: 'Taxa de Ocupação') nas abas de legislação para gerar o gráfico.")
    else:
        st.error("Não foi possível encontrar as colunas de 'Indicador' ou 'Valor Indicado' nas tabelas de parâmetros de legislação.")


elif pagina_selecionada == "Relatório de Visita":
    st.header("Relatório de Visita de Campo")
    
    df_visita = dados['visita']
    if df_visita.empty:
        st.error("Aba '1) Dados de campo (Relatório)' não encontrada.")
    else:
        col_aspecto = encontrar_coluna(df_visita, ['ASPECTO / DADO'])
        col_obs = encontrar_coluna(df_visita, ['OBSERVAÇÕES / RESPOSTAS'])

        if not col_aspecto or not col_obs:
            st.error("Colunas 'Aspecto' ou 'Observações' não encontradas no Relatório.")
            st.dataframe(df_visita)
        else:
            st.subheader("Dashboard da Visita")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                fotos = extrair_dado_visita(df_visita, "Fotografias capturadas", col_aspecto, col_obs)
                st.metric("Fotos Capturadas", str(extrair_primeiro_numero(fotos) or fotos))
            with col2:
                calcada = extrair_dado_visita(df_visita, "Largura da calçada principal", col_aspecto, col_obs)
                st.metric("Largura Calçada (m)", str(extrair_primeiro_numero(calcada) or calcada))
            with col3:
                fluxo_ped = extrair_dado_visita(df_visita, "Fluxo médio de pedestres", col_aspecto, col_obs)
                st.metric("Fluxo Pedestres (10min)", fluxo_ped)
            with col4:
                altura_viz = extrair_dado_visita(df_visita, "Altura média dos edifícios vizinhos", col_aspecto, col_obs)
                st.metric("Altura Vizinhança (m)", altura_viz)

            st.divider()
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.subheader("Análise de Ruídos")
                ruidos_str = extrair_dado_visita(df_visita, "Sons e ruídos predominantes", col_aspecto, col_obs)
                
                fontes_ruido = []
                if "trânsito" in ruidos_str.lower() or "veículos" in ruidos_str.lower():
                    fontes_ruido.append({"Fonte": "Trânsito", "Intensidade": 1})
                if "natureza" in ruidos_str.lower() or "cigarras" in ruidos_str.lower():
                    fontes_ruido.append({"Fonte": "Natureza", "Intensidade": 1})
                if "pessoas" in ruidos_str.lower():
                    fontes_ruido.append({"Fonte": "Pessoas", "Intensidade": 1})
                
                if fontes_ruido:
                    df_ruido = pd.DataFrame(fontes_ruido)
                    fig_ruido = px.bar(
                        df_ruido, 
                        x="Fonte", 
                        y="Intensidade", 
                        title="Fontes de Ruído Predominantes",
                        template="plotly_dark", 
                        color="Fonte",
                        color_discrete_map={ # AJUSTE: Alterado para azul, verde e azul claro
                            "Trânsito": "#007BFF",  # Azul
                            "Natureza": "#28A745", # Verde
                            "Pessoas": "#B0E0E6"  # Azul pálido (próximo do branco)
                        }
                    )
                    fig_ruido.update_layout(
                        yaxis_title=None, 
                        yaxis_visible=False, 
                        showlegend=False,
                        plot_bgcolor='var(--color-container)',
                        paper_bgcolor='var(--color-container)'
                    )
                    st.plotly_chart(fig_ruido, use_container_width=True)
                else:
                    st.info("Fontes de ruído não detalhadas.")
            
            with col2:
                st.subheader("Observações Principais")
                st.info(f"**Condições:** {extrair_dado_visita(df_visita, 'Condições climáticas', col_aspecto, col_obs)}")
                st.info(f"**Topografia:** {extrair_dado_visita(df_visita, 'Topografia e drenagem', col_aspecto, col_obs)}")
                st.info(f"**Vegetação:** {extrair_dado_visita(df_visita, 'Vegetação existente', col_aspecto, col_obs)}")
            
            st.divider()
            st.subheader("Todos os Dados da Visita (Tabela)")
            with st.expander("Clique para ver a tabela de dados brutos da visita"):
                df_visita_display = df_visita.dropna(subset=[col_aspecto])
                col_secao = df_visita.columns[0]
                df_visita_display[col_secao] = df_visita_display[col_secao].ffill()
                df_visita_display = df_visita_display[[col_secao, col_aspecto, col_obs]]
                st.dataframe(df_visita_display, use_container_width=True, height=400)

elif pagina_selecionada == "Estratégia e Riscos":
    st.header("Estratégia e Riscos (Resumo Analítico)")
    
    df_resumo = dados['resumo_analitico']
    
    if df_resumo.empty:
        st.error("Aba '4) Resumo analítico' não encontrada.")
    else:
        st.subheader("Diretrizes Estratégicas por Dimensão")
        
        col_dimensao = encontrar_coluna(df_resumo, ['DIMENSÃO'])
        df_resumo_display = pd.DataFrame() 

        if col_dimensao:
            col_dimensao_nome = encontrar_coluna(df_resumo, ['DIMENSÃO'])
            
            df_resumo_display = df_resumo[~df_resumo[col_dimensao_nome].astype(str).str.contains('➡️|📍|^\s*-', na=False, regex=True)]
            df_resumo_display = df_resumo_display.iloc[:, 0:4].dropna(how='all')
            df_resumo_display = df_resumo_display.dropna(subset=[col_dimensao_nome])
            df_resumo_display = df_resumo_display[df_resumo_display[col_dimensao_nome].str.strip() != '']
            
            if not df_resumo_display.empty:
                df_resumo_display = df_resumo_display.set_index(col_dimensao_nome)
        
        else:
            st.error("Não foi possível encontrar a coluna 'DIMENSÃO' no Resumo Analítico.")
            st.stop()
            
        if not df_resumo_display.empty:
            col_situacao = encontrar_coluna(df_resumo_display, ['SITUAÇÃO'])
            col_potencial = encontrar_coluna(df_resumo_display, ['POTENCIAL'])
            col_estrategia = encontrar_coluna(df_resumo_display, ['ESTRATÉGIA'])

            if col_situacao and col_potencial and col_estrategia:
                for dimensao, row in df_resumo_display.iterrows():
                    potencial_str = row[col_potencial] if pd.notna(row[col_potencial]) else "Potencial não definido"
                    situacao_str = row[col_situacao] if pd.notna(row[col_situacao]) else "Situação não definida"
                    estrategia_str = row[col_estrategia] if pd.notna(row[col_estrategia]) else "Estratégia não definida"
                    
                    with st.expander(f"**{dimensao}**: {potencial_str}"):
                        st.markdown(f"**Situação Atual:** {situacao_str}")
                        st.markdown(f"**Estratégia de Projeto:** {estrategia_str}")
            else:
                st.error("Não foi possível encontrar as colunas 'Situação', 'Potencial' e 'Estratégia' no Resumo.")
                st.dataframe(df_resumo_display)
        else:
            st.warning("Nenhum dado de estratégia encontrado após a limpeza.")
        
        st.divider()
        st.subheader("Análise de Risco (da Visita)")
        
        df_visita = dados['visita']
        if not df_visita.empty:
            col_aspecto = encontrar_coluna(df_visita, ['ASPECTO / DADO'])
            col_obs = encontrar_coluna(df_visita, ['OBSERVAÇÕES / RESPOSTAS'])

            if col_aspecto and col_obs:
                st.warning(f"**Risco de Ruído:** {extrair_dado_visita(df_visita, 'Ruídos e odores', col_aspecto, col_obs)}")
                st.warning(f"**Risco de Segurança:** {extrair_dado_visita(df_visita, 'Nível de segurança', col_aspecto, col_obs)}")
                st.warning(f"**Risco de Topografia:** {extrair_dado_visita(df_visita, 'Topografia e drenagem', col_aspecto, col_obs)}")
            else:
                st.error("Colunas de 'Aspecto' ou 'Observações' não encontradas na aba Visita.")


elif pagina_selecionada == "🤖 IA Chatbot":
    st.header("🤖 Chatbot de Análise")
    st.caption("Faça perguntas sobre os dados da análise (Ex: Qual a estratégia para a dimensão social?)")
    
    # AJUSTE: Código para buscar a chave de API dos Secrets do Streamlit
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except:
        api_key = st.sidebar.text_input(
            "Insira sua Chave de API do Google Gemini:", 
            type="password", 
            help="Obtenha sua chave no Google AI Studio para ativar o chatbot."
        )
    
    if not api_key:
        st.warning("Por favor, insira uma chave de API do Google Gemini na barra lateral (ou configure nos Secrets) para ativar o chatbot.")
        st.stop()
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-pro-latest')
    except Exception as e:
        st.error(f"Erro ao configurar a API do Gemini (verifique sua chave): {e}")
        st.stop()

    if "messages" not in st.session_state:
        st.session_state.messages = []
        st.session_state.messages.append({
            "role": "assistant", 
            "content": "Olá! Sou o assistente de análise do Studio Cosmos. Pergunte-me sobre os dados carregados."
        })

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if prompt := st.chat_input("Qual a sua pergunta?"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        try:
            contexto_resumo = "Nenhum resumo analítico encontrado."
            if not dados['resumo_analitico'].empty:
                col_dimensao_nome = encontrar_coluna(dados['resumo_analitico'], ['DIMENSÃO'])
                if col_dimensao_nome: 
                    df_resumo_limpo = dados['resumo_analitico'][~dados['resumo_analitico'][col_dimensao_nome].astype(str).str.contains('➡️|📍|^\s*-', na=False, regex=True)]
                    df_resumo_limpo = df_resumo_limpo.iloc[:, 0:4].dropna(how='all')
                    df_resumo_limpo = df_resumo_limpo.dropna(subset=[col_dimensao_nome])
                    contexto_resumo = df_resumo_limpo.to_string()
            
            contexto_visita = "Nenhum relatório de visita encontrado."
            if not dados['visita'].empty:
                contexto_visita = dados['visita'].dropna(how='all').to_string()
            
            contexto_stakeholders = "Nenhum stakeholder encontrado."
            df_eco_raw_ia = pd.DataFrame()
            nome_real_economica_ia = None
            
            try:
                uploaded_file.seek(0)
                xls_ia = pd.ExcelFile(uploaded_file)
                for sheet_name in xls_ia.sheet_names:
                    if "kpis (econômica)" in sheet_name.lower():
                        nome_real_economica_ia = sheet_name
                        break
            except Exception:
                pass 

            if nome_real_economica_ia:
                try:
                    uploaded_file.seek(0)
                    df_eco_raw_ia = pd.read_excel(
                        uploaded_file, 
                        sheet_name=nome_real_economica_ia, 
                        header=None 
                    )
                except Exception:
                    pass 

            if not df_eco_raw_ia.empty:
                header_row_index_ia = -1
                for i, row in df_eco_raw_ia.head(30).iterrows():
                    row_values_ia = [clean_str(val) for val in row.values]
                    has_instituicao_ia = any('instituição' in val or 'instituicao' in val for val in row_values_ia)
                    has_potencial_ia = any('potencial' in val for val in row_values_ia)
                    
                    if has_instituicao_ia and has_potencial_ia:
                        header_row_index_ia = i
                        break
                        
                if header_row_index_ia != -1:
                    df_stakeholders_ia = df_eco_raw_ia.loc[header_row_index_ia:].copy()
                    
                    header_values_ia = df_stakeholders_ia.iloc[0]
                    new_cols_ia = []
                    counts_ia = {}
                    for col_ia in header_values_ia:
                        col_name_ia = clean_str(col_ia)
                        if col_name_ia == "nan" or col_name_ia == "":
                            col_name_ia = "unnamed"
                        if col_name_ia in counts_ia:
                            counts_ia[col_name_ia] += 1
                            new_cols_ia.append(f"{col_name_ia}_{counts_ia[col_name_ia]}")
                        else:
                            counts_ia[col_name_ia] = 0
                            new_cols_ia.append(col_name_ia)
                    
                    df_stakeholders_ia.columns = new_cols_ia
                    df_stakeholders_ia = df_stakeholders_ia.iloc[1:]
                    
                    col_inst_ia = encontrar_coluna(df_stakeholders_ia, ['instituição', 'instituicao'])
                    col_pot_ia = encontrar_coluna(df_stakeholders_ia, ['potencial'])
                    
                    if col_inst_ia and col_pot_ia:
                        df_stakeholders_ia = df_stakeholders_ia.dropna(subset=[col_inst_ia, col_pot_ia])
                        contexto_stakeholders = df_stakeholders_ia.to_string()

            contexto_dados = f"""
            DADOS DE CONTEXTO ESTRATÉGICO:
            
            Métricas Chave:
            - Média da Dimensão Física: {dados['metricas']['media_fisica']:.2f}
            - Média da Dimensão Social: {dados['metricas']['media_social']:.2f}
            - Média da Dimensão Ambiental: {dados['metricas']['media_ambiental']:.2f}
            - Média da Dimensão Urbana: {dados['metricas']['media_urbana']:.2f}
            - Média da Dimensão Econômica: {dados['metricas']['media_economica']:.2f}
            - Média da Dimensão Sensorial: {dados['metricas']['media_sensorial']:.2f}
            - Índice Territorial (IT) Total: {dados['metricas']['it_total']:.2f}
            
            Resumo das Estratégias (Aba 'Resumo Analítico'):
            {contexto_resumo}
            
            Observações da Visita de Campo (Aba 'Dados de campo'):
            {contexto_visita}

            Parceiros e Stakeholders (da aba Econômica):
            {contexto_stakeholders}
            """
        except Exception as e:
            st.error(f"Erro ao montar o contexto para a IA. Detalhe: {e}")
            contexto_dados = "Erro ao carregar dados."

        prompt_para_ia = f"""
        Você é um assistente de arquitetura sênior do Studio Cosmos.
        Sua tarefa é responder perguntas sobre uma análise de viabilidade de terreno.
        Use **exclusivamente** os {contexto_dados} para formular sua resposta.
        
        Se a informação não estiver no contexto, diga "Essa informação não foi encontrada nos dados carregados".
        Não invente números ou dados que não estejam no contexto.
        Seja objetivo, profissional e use markdown (como negrito) para destacar os números e pontos-chave.
        
        PERGUNTA DO USUÁRIOS:
        {prompt}
        """

        try:
            with st.spinner("Analisando..."):
                response = model.generate_content(prompt_para_ia)
                resposta_ia = response.text
            
            with st.chat_message("assistant"):
                st.markdown(resposta_ia)
            st.session_state.messages.append({"role": "assistant", "content": resposta_ia})
        
        except Exception as e:
            st.error(f"Erro ao contactar a IA: {e}")
            msg_erro = f"Desculpe, não consegui processar sua pergunta. Erro: {e}"
            st.session_state.messages.append({"role": "assistant", "content": msg_erro})