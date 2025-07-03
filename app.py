import streamlit as st
import json
import os
import sys
import logging
import tempfile
from datetime import datetime, timedelta, timezone
import pandas as pd
from io import BytesIO
from docx import Document
from PyPDF2 import PdfMerger
import plotly.express as px
import folium
from streamlit_folium import st_folium
import unicodedata
from supabase import create_client, Client
from dotenv import load_dotenv
from analyze_curativo import analyze_curativo
import subprocess




# A importação incorreta de 'docx2pdf' foi REMOVIDA daqui.

# Carregar variáveis de ambiente
load_dotenv()

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuração da página
st.set_page_config(
    page_title="Controle de Bombas de Sucção",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilo CSS
st.markdown("""
    <style>
        body { font-family: 'Arial', sans-serif; }
        .main { padding: 20px; }
        .main-title { text-align: center; font-size: 2.5rem; font-weight: bold; color: #1f2937; margin-bottom: 10px; }
        .section-title { text-align: center; font-size: 2rem; font-weight: bold; color: #1f2937; margin-top: 20px; margin-bottom: 20px; }
        .subtitle { text-align: center; font-size: 1rem; color: #4b5563; margin-bottom: 10px; }
        .filial-main { text-align: center; font-size: 1.2rem; font-weight: 500; color: #1e90ff; margin-bottom: 20px; }
        .sidebar .sidebar-content { min-width: 240px; max-width: 240px; background-color: #f9fafb; padding: 20px; }
        .filial-sidebar {
            background-color: #e6f0fa; padding: 10px; border-radius: 8px;
            font-size: 1.1rem; font-weight: bold; color: #1f2937; margin-bottom: 20px;
            text-align: center; border: 1px solid #1e90ff;
        }
        .stButton > button {
            width: 100%; padding: 10px; border-radius: 8px; font-weight: 500; transition: background-color 0.2s;
        }
        .stButton button.filial-button {
            background-color: #1e90ff; color: white;
            width: 50%; margin-left: 25%;
        }
        .stButton button.filial-button:hover {
            background-color: #104e8b;
        }
        .stButton button.general-button {
            background-color: #28a745; color: white;
        }
        .stButton button.general-button:hover {
            background-color: #218838;
        }
        .stButton button.confirm-button {
            background-color: #ff9800; color: white;
        }
        .stButton button.confirm-button:hover {
            background-color: #f57c00;
        }
        .stDownloadButton > button { background-color: #28a745 !important; color: white !important; }
        .stDownloadButton > button:hover { background-color: #218838 !important; }
        .stTextInput > div > input { border: 1px solid #d1d5db; border-radius: 8px; padding: 10px; width: 100%; font-size: 1rem; }
        .stTextInput > div { margin-bottom: 15px; }
        .form-container { display: flex; flex-wrap: wrap; gap: 20px; }
        .form-column { flex: 1; min-width: 300px; }
        .stDataFrame { border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
        .stDataFrame table { width: 100%; table-layout: auto; }
        .stDataFrame th, .stDataFrame td { 
            padding: 8px; border: 1px solid #ddd; text-align: left; 
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis; 
        }
        .stMetric { background-color: #f9fafb; border-radius: 8px; padding: 16px; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05); }
        .stMetric * { color: #000000 !important; }
        h1, h2, h3, h5 { color: #1f2937; font-weight: 600; }
        .status-badge { display: inline-block; padding: 4px 8px; border-radius: 12px; font-size: 0.875rem; font-weight: 500; }
        .status-no-prazo { background-color: #d1fae5; color: #065f46; }
        .status-menos-7-dias { background-color: #fefcbf; color: #92400e; }
        .status-fora-prazo { background-color: #fee2e2; color: #991b1b; }
        .status-indefinido, .status-data-invalida { background-color: #e5e7eb; color: #4b5563; }
        .status-devolvida { background-color: #d1e7dd; color: #0f5132; }
        .status-em-manutencao { background-color: #fefcbf; color: #92400e; }
        .legend-item { display: flex; align-items: center; margin: 5px 0; }
        .legend-dot { height: 15px; width: 15px; border-radius: 50%; margin-right: 10px; }
        .brasilia-dot { background-color: #1E90FF; }
        .goiania-dot { background-color: #32CD32; }
        .cuiaba-dot { background-color: #FFD700; }
        @media (max-width: 768px) {
            .main-title, .section-title { font-size: 1.8rem; }
            .form-column { min-width: 100%; }
            .sidebar .sidebar-content { min-width: 200px; max-width: 200px; }
            .stDataFrame th, .stDataFrame td { font-size: 0.9rem; }
        }
    </style>
""", unsafe_allow_html=True)

# -------------------- CONFIGURAÇÕES GERAIS --------------------
CONFIG_FILE = "config.json"
GENERAL_PWD = "suplen2025"
FILIAIS = ["BRASILIA", "GOIANIA", "CUIABA"]
CONTRATO_LOCAL_PATH = "contrato.docx"
CONTRATO_STORAGE_PATH = "contratos/contrato.docx"
MAINTENANCE_STORAGE_PATH = "nfs/"

# Função para normalizar texto
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII').upper().strip()

# --- INÍCIO DA NOVA FUNÇÃO CENTRALIZADA PARA DATAS ---
def parse_supabase_date(date_string: str | None) -> datetime | None:
    """
    Converte com segurança uma string de data do Supabase para um objeto datetime,
    lidando com o sufixo 'Z' para compatibilidade entre versões do Python.
    """
    if not date_string or not isinstance(date_string, str):
        return None
    try:
        # Garante compatibilidade com Python < 3.11 que não lida com 'Z' nativamente
        if date_string.endswith('Z'):
            date_string = date_string[:-1] + '+00:00'
        return datetime.fromisoformat(date_string)
    except (ValueError, TypeError):
        logging.warning(f"Não foi possível converter a string de data: {date_string}")
        return None
# --- FIM DA NOVA FUNÇÃO ---

# Dicionário de senhas para filiais
FILIAIS_PASSWORDS = {filial: normalize_text(filial) + "123" for filial in FILIAIS}

# Inicializar Supabase
@st.cache_resource
def init_supabase():
    try:
        supabase_url = os.getenv("SUPABASE_URL")
        supabase_key = os.getenv("SUPABASE_KEY")
        if not supabase_url or not supabase_key:
            st.error("Variáveis SUPABASE_URL ou SUPABASE_KEY não encontradas no arquivo .env.")
            st.stop()
        supabase: Client = create_client(supabase_url, supabase_key)
        logging.info("Supabase inicializado com sucesso.")
        return supabase
    except Exception as e:
        logging.error(f"Erro ao inicializar Supabase: {e}")
        st.error(f"Erro ao conectar ao Supabase: {e}")
        st.stop()
supabase = init_supabase()

# -------------------- FUNÇÕES AUXILIARES OTIMIZADAS --------------------
@st.cache_data(ttl=300)
def download_file_from_storage(storage_path):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        response = supabase.storage.from_(bucket_name).download(storage_path)
        logging.info(f"Arquivo {storage_path} baixado com sucesso.")
        return response
    except Exception as e:
        if "Object not found" in str(e):
            logging.warning(f"Arquivo não encontrado no storage: {storage_path}")
            return None
        logging.error(f"Erro ao baixar {storage_path}: {str(e)}")
        return None

@st.cache_data(ttl=300)
def get_dados_bombas_df():
    try:
        response = supabase.table("DADOS_BOMBAS").select("Serial, Modelo, Ultima_Manut, Venc_Manut").execute()
        data = response.data
        if not data:
            logging.warning("A tabela DADOS_BOMBAS está vazia ou não foi encontrada.")
            return pd.DataFrame(columns=['Serial', 'Modelo', 'Ultima_Manut', 'Venc_Manut'])
        df = pd.DataFrame(data)
        df['Serial_Normalized'] = df['Serial'].apply(normalize_text)
        df['Ultima_Manut'] = pd.to_datetime(df['Ultima_Manut'], errors='coerce')
        df['Venc_Manut'] = pd.to_datetime(df['Venc_Manut'], errors='coerce')
        logging.info("Dados da tabela DADOS_BOMBAS (tipo DATE) carregados com sucesso.")
        return df
    except Exception as e:
        logging.error(f"Erro ao carregar dados da tabela DADOS_BOMBAS: {e}")
        st.error(f"Não foi possível carregar os dados de manutenção das bombas: {e}")
        return pd.DataFrame()

# -------------------- CONFIGURAÇÃO INICIAL --------------------
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_config(filial):
    config = {"filial": filial}
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def setup_filial():
    config = load_config()
    if "filial" in config:
        st.sidebar.markdown(f'<div class="filial-sidebar">Filial atual: {config["filial"]}</div>', unsafe_allow_html=True)
        if st.sidebar.button("Mudar Filial", key="change_filial"):
            if os.path.exists(CONFIG_FILE):
                os.remove(CONFIG_FILE)
            st.session_state.clear()
            st.session_state.show_filial = True
            st.session_state.show_general = False
            st.session_state.general_mode = False
            st.rerun()
        return config["filial"]
    if "show_filial" not in st.session_state:
        st.session_state.show_filial = True
    if "show_general" not in st.session_state:
        st.session_state.show_general = False
    if "general_mode" not in st.session_state:
        st.session_state.general_mode = False
    st.sidebar.header("Configuração Inicial")
    if st.session_state.show_filial:
        filial = st.sidebar.selectbox("Selecione a Filial", FILIAIS, key="filial_select")
        pwd = st.sidebar.text_input("Senha", type="password", key="filial_pwd")
        if st.sidebar.button("Confirmar Filial", key="confirm_filial"):
            if pwd == FILIAIS_PASSWORDS.get(filial):
                save_config(filial)
                st.session_state.filial = filial
                st.session_state.general_mode = False
                st.session_state.show_filial = False
                st.session_state.show_general = False
                st.rerun()
            else:
                st.sidebar.error("Senha incorreta!")
        if st.sidebar.button("Dashboard Geral", key="general_button"):
            st.session_state.show_filial = False
            st.session_state.show_general = True
            st.rerun()
    if st.session_state.show_general:
        general_pwd = st.sidebar.text_input("Senha para Dashboard Geral", type="password", key="general_pwd")
        if st.sidebar.button("Acessar Dashboard Geral", key="access_general"):
            if general_pwd == GENERAL_PWD:
                st.session_state.filial = None
                st.session_state.general_mode = True
                st.session_state.show_filial = False
                st.session_state.show_general = False
                st.rerun()
            else:
                st.sidebar.error("Senha do Dashboard Geral incorreta!")
        if st.sidebar.button("Voltar para Filial", key="back_to_filial"):
            st.session_state.show_filial = True
            st.session_state.show_general = False
            st.session_state.general_mode = False
            st.rerun()
    if st.session_state.get("general_mode", False):
        if st.sidebar.button("Sair do Dashboard Geral", key="exit_general"):
            if os.path.exists(CONFIG_FILE):
                os.remove(CONFIG_FILE)
            st.session_state.clear()
            st.session_state.show_filial = True
            st.session_state.show_general = False
            st.session_state.general_mode = False
            st.rerun()
    return None if st.session_state.get("general_mode", False) else config.get("filial")

# -------------------- LÓGICA DE DADOS OTIMIZADA --------------------
def calculate_status(data_saida, periodo):
    logging.info(f"Calculando status com data_saida: {data_saida} (tipo: {type(data_saida)}) e periodo: {periodo} (tipo: {type(periodo)})")
    if not data_saida or not periodo:
        return "Indefinido"
    try:
        data_devolucao = data_saida + timedelta(days=int(periodo))
        # ### CORREÇÃO PRINCIPAL ###
        # Compara data com data, usando .date() para remover a informação de hora
        dias_restantes = (data_devolucao - datetime.now().date()).days
        if dias_restantes < 0:
            return "Fora Prazo"
        elif dias_restantes <= 7:
            return "Menos de 7 dias"
        return "No Prazo"
    except (ValueError, TypeError) as e:
        logging.error(f"Erro em calculate_status: {e}")
        return "Data Inválida"

def register_event(table, record_id, description, filial):
    if "event_buffer" not in st.session_state:
        st.session_state.event_buffer = []
    st.session_state.event_buffer.append({
        "data_evento": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "descricao": description.upper(),
        "filial": filial
    })

def flush_events():
    if st.session_state.get("event_buffer"):
        try:
            supabase.table("historico").insert(st.session_state.event_buffer).execute()
            st.session_state.event_buffer = []
            logging.info("Eventos registrados em lote.")
        except Exception as e:
            logging.error(f"Erro ao registrar eventos em lote: {e}")
            st.error("Erro ao salvar eventos.")

@st.cache_data(ttl=300, show_spinner=False)
def get_bombas(search_term="", filial=None, active_only=True):
    try:
        query = supabase.table("bombas").select("*")
        if active_only:
            query = query.eq("ativo", True)
        if filial:
            query = query.eq("filial", filial)
        bombas = query.execute().data
        if not bombas:
            return []
        if search_term:
            result = []
            search_term_lower = search_term.lower()
            for bomba in bombas:
                if any(search_term_lower in str(bomba.get(field, "")).lower() for field in ["serial", "paciente", "hospital"]):
                    result.append(bomba)
            bombas = result
        for bomba in bombas:
            bomba["id"] = str(bomba["id"])
            for field in ["data_saida", "data_registro", "data_retorno"]:
                # --- USO DA NOVA FUNÇÃO DE DATA ---
                dt_obj = parse_supabase_date(bomba.get(field))
                if dt_obj:
                    bomba[field] = dt_obj.strftime("%d/%m/%Y")
        logging.info(f"Busca de bombas retornou {len(bombas)} registros.")
        return bombas
    except Exception as e:
        logging.error(f"Erro ao buscar bombas: {e}")
        st.error("Erro ao acessar dados das bombas.")
        return []

@st.cache_data(ttl=300, show_spinner=False)
def get_manutencao(search_term="", filial=None):
    try:
        query = supabase.table("manutencao").select("*")
        if filial:
            query = query.eq("filial", filial)
        manutencoes = query.execute().data
        if not manutencoes:
            return []
        manut_df = pd.DataFrame(manutencoes)
        if search_term:
            search_term_lower = search_term.lower()
            mask = manut_df.apply(lambda row: any(search_term_lower in str(row.get(field, "")).lower() for field in ["serial", "defeito", "nf_numero"]), axis=1)
            manut_df = manut_df[mask]
        if manut_df.empty:
            return []
        bombas_df = get_dados_bombas_df()
        if not bombas_df.empty:
            manut_df['Serial_Normalized'] = manut_df['serial'].apply(normalize_text)
            merged_df = pd.merge(manut_df, bombas_df[['Serial_Normalized', 'Modelo', 'Ultima_Manut', 'Venc_Manut']], left_on='Serial_Normalized', right_on='Serial_Normalized', how='left')
        else:
            merged_df = manut_df
            merged_df['Modelo'] = "N/A"
            merged_df['Ultima_Manut'] = pd.NaT
            merged_df['Venc_Manut'] = pd.NaT
        merged_df['modelo'] = merged_df['Modelo'].fillna("N/A")
        merged_df['ultima_manut'] = pd.to_datetime(merged_df['Ultima_Manut'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("N/A")
        merged_df['venc_manut'] = pd.to_datetime(merged_df['Venc_Manut'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("N/A")
        
        # --- USO DA NOVA FUNÇÃO DE DATA ---
        merged_df['data_registro'] = merged_df['data_registro'].apply(lambda x: (parse_supabase_date(x).strftime('%d/%m/%Y') if parse_supabase_date(x) else "N/A"))

        merged_df['id'] = merged_df['id'].astype(str)
        return merged_df.to_dict('records')
    except Exception as e:
        logging.error(f"Erro ao buscar manutenções: {e}")
        st.error("Erro ao acessar dados de manutenção.")
        return []

@st.cache_data(ttl=300, show_spinner=False)
def get_historico_devolvidas(filial=None):
    try:
        query = supabase.table("historico").select("*").ilike("descricao", "%BOMBA DEVOLVIDA%").order("data_evento", desc=True)
        if filial:
            query = query.eq("filial", filial)
        historico = query.execute().data
        for doc in historico:
            doc["id"] = str(doc["id"])
            # --- USO DA NOVA FUNÇÃO DE DATA ---
            dt_obj = parse_supabase_date(doc.get("data_evento"))
            if dt_obj:
                doc["data_evento"] = dt_obj.strftime("%d/%m/%Y, %H:%M")
        return historico
    except Exception as e:
        logging.error(f"Erro ao buscar histórico de devolvidas: {e}")
        st.error(f"Erro ao acessar histórico: {e}")
        return []

@st.cache_data(ttl=300, show_spinner="Carregando dados de curativos...")
def get_saldo_curativo_data():
    try:
        # Busca todos os dados do banco
        response = supabase.table("saldo_curativo").select("*").execute()
        data = response.data
        if not data:
            return pd.DataFrame()
            
        # Converte para DataFrame
        df = pd.DataFrame(data)
        total_registros = len(df)
        
        # Identifica duplicatas baseado em todas as colunas relevantes
        df_sem_duplicatas = df.drop_duplicates(
            subset=['Produto', 'Desc_Produto', 'Referencia', 'Lote', 'Data_Validad', 'Saldo_Lote']
        )
        registros_unicos = len(df_sem_duplicatas)
        
        # Se encontrou duplicatas, mostra informação
        if total_registros > registros_unicos:
            duplicatas = total_registros - registros_unicos
            st.sidebar.warning(f"""
                ⚠️ Atenção: Foram encontradas duplicatas no banco
                - Total de registros: {total_registros}
                - Registros únicos: {registros_unicos}
                - Duplicatas removidas: {duplicatas}
            """)
            
            # Mostra as linhas que estão duplicadas
            linhas_duplicadas = df[df.duplicated(
                subset=['Produto', 'Desc_Produto', 'Referencia', 'Lote', 'Data_Validad', 'Saldo_Lote'],
                keep=False
            )]
            if not linhas_duplicadas.empty:
                with st.sidebar.expander("Ver registros duplicados"):
                    st.dataframe(
                        linhas_duplicadas.sort_values(by=['Produto', 'Lote']),
                        use_container_width=True
                    )
        
        # Retorna o DataFrame sem duplicatas
        return df_sem_duplicatas
        
    except Exception as e:
        logging.error(f"Erro ao carregar dados da tabela saldo_curativo: {e}")
        st.error(f"Não foi possível carregar os dados de saldo de curativos: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def get_dashboard_metrics(filial=None):
    try:
        dados_bombas_df = get_dados_bombas_df()
        if dados_bombas_df.empty:
            logging.warning("DADOS_BOMBAS está vazio, as métricas podem ser imprecisas.")
        active_pumps_list = supabase.table("bombas").select("serial, ativo, status, hospital, filial").eq("ativo", True).execute().data
        active_pumps_df = pd.DataFrame(active_pumps_list if active_pumps_list else [])
        manut_list = supabase.table("manutencao").select("id, filial, serial").eq("status", "Em Manutenção").execute().data
        manut_df = pd.DataFrame(manut_list if manut_list else [])

        if filial and not active_pumps_df.empty:
            scoped_active_pumps_df = active_pumps_df[active_pumps_df['filial'] == filial].copy()
        else:
            scoped_active_pumps_df = active_pumps_df.copy()
        if filial and not manut_df.empty:
            scoped_manut_df = manut_df[manut_df['filial'] == filial].copy()
        else:
            scoped_manut_df = manut_df.copy()

        em_manutencao = len(scoped_manut_df)

        ulta_count = 0
        activac_count = 0
        if not scoped_active_pumps_df.empty and not dados_bombas_df.empty:
            scoped_active_pumps_df['Serial_Normalized'] = scoped_active_pumps_df['serial'].apply(normalize_text)
            merged_df = pd.merge(scoped_active_pumps_df, dados_bombas_df[['Serial_Normalized', 'Modelo']], on='Serial_Normalized', how='left')
            model_counts = merged_df['Modelo'].str.upper().value_counts().to_dict()
            ulta_count = model_counts.get('ULTA', 0)
            activac_count = model_counts.get('ACTIVAC', 0)
        total_comodato_painel = ulta_count + activac_count

        total_bombas_inventario = len(dados_bombas_df)
        seriais_ativos = set(active_pumps_df['serial'].apply(normalize_text)) if not active_pumps_df.empty else set()
        seriais_manut = set(manut_df['serial'].apply(normalize_text)) if not manut_df.empty else set()
        seriais_indisponiveis = seriais_ativos.union(seriais_manut)

        df_estoque = pd.DataFrame()
        if not dados_bombas_df.empty:
            df_estoque = dados_bombas_df[~dados_bombas_df['Serial_Normalized'].isin(seriais_indisponiveis)].copy()

        disponiveis = 0
        modelos_disponiveis_counts = {}
        if not df_estoque.empty:
            hoje = datetime.now().date()
            df_estoque['Venc_Manut_Date'] = pd.to_datetime(df_estoque['Venc_Manut'], errors='coerce').dt.date
            df_disponiveis = df_estoque[df_estoque['Venc_Manut_Date'] >= hoje].copy()
            disponiveis = len(df_disponiveis)
            if not df_disponiveis.empty:
                modelos_disponiveis_counts = df_disponiveis['Modelo'].str.upper().value_counts().to_dict()

        status_counts = scoped_active_pumps_df['status'].value_counts().to_dict() if not scoped_active_pumps_df.empty else {}
        hosp_counts = scoped_active_pumps_df['hospital'].apply(normalize_text).value_counts().to_dict() if not scoped_active_pumps_df.empty else {}
        bombas_por_filial = active_pumps_df['filial'].apply(normalize_text).value_counts().to_dict() if not active_pumps_df.empty else {}
        
        modelos_por_filial_records = []
        if not active_pumps_df.empty and not dados_bombas_df.empty:
            active_pumps_df['Serial_Normalized'] = active_pumps_df['serial'].apply(normalize_text)
            merged_global_df = pd.merge(active_pumps_df, dados_bombas_df[['Serial_Normalized', 'Modelo']], on='Serial_Normalized', how='left')
            if not merged_global_df.empty and 'filial' in merged_global_df.columns and 'Modelo' in merged_global_df.columns:
                merged_global_df.dropna(subset=['filial', 'Modelo'], inplace=True)
                modelos_por_filial_df = merged_global_df.groupby(['filial', 'Modelo']).size().reset_index(name='Quantidade')
                modelos_por_filial_records = modelos_por_filial_df.to_dict('records')
        
        metrics = {
            "ativas": total_comodato_painel, "disponiveis": disponiveis, "em_manutencao": em_manutencao,
            "ulta_count": ulta_count, "activac_count": activac_count, "total_bombas": total_bombas_inventario,
            "status_counts": {**{"No Prazo": 0, "Menos de 7 dias": 0, "Fora Prazo": 0}, **status_counts},
            "hosp_counts": hosp_counts,
            "bombas_por_filial": {**{"BRASILIA": 0, "GOIANIA": 0, "CUIABA": 0}, **bombas_por_filial},
            "modelos_por_filial": modelos_por_filial_records,
            "modelos_disponiveis": modelos_disponiveis_counts,
        }
        return metrics
    except Exception as e:
        logging.error(f"Erro ao obter métricas do dashboard: {e}")
        st.error(f"Erro ao carregar métricas do dashboard: {e}")
        return None

def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    try:
        out_dir = os.path.dirname(pdf_path) or "."; os.makedirs(out_dir, exist_ok=True)
        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path]
        result = subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        converted_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if not os.path.exists(converted_pdf): raise FileNotFoundError(f"PDF não encontrado: {converted_pdf}. Output: {result.stderr.decode()}")
        if os.path.abspath(converted_pdf) != os.path.abspath(pdf_path): os.replace(converted_pdf, pdf_path)
        return True
    except Exception as e:
        st.warning(f"Erro ao converter para PDF: {e}")
        return False

def generate_combined_pdf(bomba_data):
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_docx_path = os.path.join(temp_dir, f"contrato_{bomba_data['serial']}.docx")
            if os.path.exists(CONTRATO_LOCAL_PATH):
                import shutil
                shutil.copyfile(CONTRATO_LOCAL_PATH, temp_docx_path)
            else:
                file_content = download_file_from_storage(CONTRATO_STORAGE_PATH)
                if not file_content: st.error("Modelo 'contrato.docx' não encontrado!"); return None
                with open(temp_docx_path, "wb") as f: f.write(file_content)
            doc = Document(temp_docx_path)
            data_atual = datetime.now()
            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            data_formatada = f"Brasília, {data_atual.day:02d} de {meses[data_atual.month - 1]} de {data_atual.year}"
            replacements = {"{SERIAL}": bomba_data.get("serial", "N/A"), "{PACIENTE}": bomba_data.get("paciente", "N/A"), "{NOTA_FISCAL}": bomba_data.get("nf", "N/A"), "{DATA_ATUAL}": data_formatada}
            for p in doc.paragraphs:
                for key, value in replacements.items():
                    if key in p.text:
                        inline = p.runs
                        for i in range(len(inline)):
                            if key in inline[i].text: inline[i].text = inline[i].text.replace(key, str(value))
            doc.save(temp_docx_path)
            temp_pdf_path = os.path.join(temp_dir, f"contrato_{bomba_data['serial']}.pdf")
            if not convert_docx_to_pdf(temp_pdf_path, temp_pdf_path): return None
            merger = PdfMerger()
            merger.append(temp_pdf_path)
            pdf_serial_content = download_file_from_storage(f"pdfs/{bomba_data['serial']}.pdf")
            if pdf_serial_content: merger.append(BytesIO(pdf_serial_content))
            pdf_buffer = BytesIO(); merger.write(pdf_buffer); merger.close(); pdf_buffer.seek(0)
            return pdf_buffer
    except Exception as e:
        st.error(f"Erro ao gerar PDF: {e}")
        return None

def generate_excel_saldo_curativo(df):
    if df.empty: return None
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df.to_excel(writer, index=False, sheet_name='SaldoCurativo')
    return output.getvalue()

def upload_nf_pdf(serial, data_registro, file):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        data_registro_clean = datetime.strptime(data_registro, "%d/%m/%Y").strftime("%Y%m%d")
        file_name = f"{MAINTENANCE_STORAGE_PATH}{serial}_{data_registro_clean}.pdf"
        supabase.storage.from_(bucket_name).upload(file_name, file.getvalue(), file_options={"content-type": "application/pdf", "upsert": "true"})
        return True
    except Exception as e:
        st.error(f"Erro ao enviar NF: {e}")
        return False

def upload_nf_assinada(bomba_data, file):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        serial = bomba_data.get("serial", "SEM_SERIAL")
        hospital = normalize_text(bomba_data.get('hospital', 'SEM_HOSPITAL')).replace(' ', '')
        paciente = normalize_text(bomba_data.get('paciente', 'SEM_PACIENTE')).replace(' ', '')
        data_registro_str = "DD-MM-AAAA"
        if bomba_data.get('data_registro'):
            try: data_registro_str = datetime.strptime(bomba_data['data_registro'], '%d/%m/%Y').strftime('%d-%m-%Y')
            except (ValueError, TypeError): data_registro_str = str(bomba_data['data_registro']).replace('/', '-')
        file_name = f"nfs_assinadas/{serial}*{hospital}_{paciente}*{data_registro_str}_assinado.pdf"
        supabase.storage.from_(bucket_name).upload(file_name, file.getvalue(), file_options={"content-type": "application/pdf", "cache-control": "3600", "upsert": "true"})
        return True
    except Exception as e:
        st.error(f"Erro ao enviar NF assinada: {e}")
        return False

@st.cache_data(ttl=300)
def get_all_nfs_assinadas_info():
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        path = "nfs_assinadas/"
        files = supabase.storage.from_(bucket_name).list(path)
        serial_to_filename_map = {}
        for file_info in files:
            name = file_info.get('name')
            if name and name.endswith("_assinado.pdf"):
                serial_to_filename_map[name.split('*')[0]] = f"{path}{name}"
        return serial_to_filename_map
    except Exception as e:
        logging.error(f"Erro ao listar NFs assinadas: {e}")
        return {}

def get_nf_assinada_filename(serial, nf_map): return nf_map.get(serial)
def check_nf_assinada(serial, nf_map): return serial in nf_map

@st.cache_data(ttl=300)
def download_nf_assinada(serial):
    try:
        nf_map = get_all_nfs_assinadas_info()
        file_path = get_nf_assinada_filename(serial, nf_map)
        if file_path: return supabase.storage.from_("controle-de-bombas-suplen-files").download(file_path)
        return None
    except Exception as e:
        logging.error(f"Erro ao baixar NF assinada para {serial}: {e}")
        return None

def format_status(status):
    status_map = {"No Prazo": '<span class="status-badge status-no-prazo">🟢 No Prazo</span>', "Menos de 7 dias": '<span class="status-badge status-menos-7-dias">🟡 Menos de 7 dias</span>', "Fora Prazo": '<span class="status-badge status-fora-prazo">🔴 Fora Prazo</span>', "Indefinido": '<span class="status-badge status-indefinido">Indefinido</span>', "Data Inválida": '<span class="status-badge status-data-invalida">Data Inválida</span>', "✅ DEVOLVIDA": '<span class="status-badge status-devolvida">✅ Devolvida</span>', "Em Manutenção": '<span class="status-badge status-em-manutencao">🛠 Em Manutenção</span>'}
    return status_map.get(status, status)

def display_manutencao_table(title, manutencoes):
    st.markdown(f"### {title}");
    if not manutencoes: st.info("Nenhum registro encontrado."); return
    df = pd.DataFrame(manutencoes)
    df_display = df[["data_registro", "serial", "modelo", "ultima_manut", "venc_manut", "defeito", "nf_numero", "status"]].copy()
    df_display.rename(columns={"data_registro": "DATA REGISTRO", "serial": "SERIAL", "modelo": "MODELO", "ultima_manut": "ÚLTIMA MANUT", "venc_manut": "VENC. MANUT", "defeito": "DEFEITO", "nf_numero": "NF EMITIDA", "status": "STATUS"}, inplace=True)
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    st.markdown("---"); st.markdown("#### Ações de Manutenção")
    manutencoes_em_andamento = [m for m in manutencoes if m.get('status') == 'Em Manutenção']
    if manutencoes_em_andamento:
        selected_manut = st.selectbox("Selecione uma manutenção para marcar como devolvida:", options=manutencoes_em_andamento, format_func=lambda m: f"Serial: {m.get('serial')} - Defeito: {m.get('defeito', '')[:30]}...", key="devolver_manut_select", index=None, placeholder="Selecione uma bomba em manutenção")
        if selected_manut and st.button(f"Marcar Devolução da Bomba {selected_manut['serial']}", key=f"devolver_manut_{selected_manut['id']}"):
            try:
                with st.spinner("Atualizando status..."):
                    supabase.table("manutencao").update({"status": "Devolvida"}).eq("id", selected_manut["id"]).execute()
                    register_event("manutencao", selected_manut["id"], f"BOMBA DEVOLVIDA APÓS MANUTENÇÃO (SERIAL: {selected_manut['serial']})", selected_manut['filial'])
                    flush_events(); st.session_state.messages.append({"text": f"Bomba {selected_manut['serial']} marcada como devolvida!", "icon": "✅"}); st.cache_data.clear(); st.rerun()
            except Exception as e: st.error(f"Erro ao marcar como devolvida: {e}")
    else: st.info("Nenhuma bomba 'Em Manutenção' para realizar ações.")

def display_bombas_table(title, bombas_list, bombas_df, nf_map):
    st.markdown(f"### {title}");
    if not bombas_list: st.info("Nenhum registro encontrado."); return
    df = pd.DataFrame(bombas_list)
    df['serial_normalized'] = df['serial'].apply(normalize_text)
    if not bombas_df.empty:
        merged_df = pd.merge(df, bombas_df, left_on='serial_normalized', right_on='Serial_Normalized', how='left')
        df['modelo'] = merged_df['Modelo'].fillna('N/A'); df['ultima_manut'] = merged_df['Ultima_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A'); df['venc_manut'] = merged_df['Venc_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')
    else: df['modelo'] = 'N/A'; df['ultima_manut'] = 'N/A'; df['venc_manut'] = 'N/A'
    df['nf_assinada'] = df['serial'].apply(lambda s: check_nf_assinada(s, nf_map)); df['nf_assinada_str'] = df['nf_assinada'].apply(lambda x: "✅ Sim" if x else "❌ Não")
    df['status_html'] = df['status'].apply(format_status); df['periodo_str'] = df['periodo'].apply(lambda x: f"{x} dias" if pd.notna(x) and x != "N/A" else "N/A")
    display_df = df[["serial", "modelo", "hospital", "paciente", "data_saida", "periodo_str", "status_html", "nf", "ultima_manut", "venc_manut", "nf_assinada_str"]].rename(columns={"serial": "SERIAL", "modelo": "MODELO", "hospital": "HOSPITAL", "paciente": "PACIENTE", "data_saida": "DATA SAÍDA", "periodo_str": "PERÍODO", "status_html": "STATUS", "nf": "NF", "ultima_manut": "ÚLTIMA MANUT", "venc_manut": "VENC MANUT", "nf_assinada_str": "NF ASSINADA"})
    st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

def generate_excel_bombas_ativas(bombas, bombas_df, filial, nf_map):
    if not bombas: return None
    df = pd.DataFrame(bombas); df['serial_normalized'] = df['serial'].apply(normalize_text)
    if not bombas_df.empty:
        merged_df = pd.merge(df, bombas_df, left_on='serial_normalized', right_on='Serial_Normalized', how='left')
        df['modelo'] = merged_df['Modelo'].fillna('N/A'); df['ultima_manut'] = merged_df['Ultima_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A'); df['venc_manut'] = merged_df['Venc_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')
    else: df['modelo'] = 'N/A'; df['ultima_manut'] = 'N/A'; df['venc_manut'] = 'N/A'
    df['nf_assinada'] = df['serial'].apply(lambda s: check_nf_assinada(s, nf_map)); df['NF ASSINADA'] = df['nf_assinada'].apply(lambda x: "Sim" if x else "Não")
    cols = ["serial", "modelo", "hospital", "paciente", "data_saida", "periodo", "status", "nf", "ultima_manut", "venc_manut", "NF ASSINADA"]
    display_cols = ["SERIAL", "MODELO", "HOSPITAL", "PACIENTE", "DATA SAÍDA", "PERÍODO", "STATUS", "NF", "ÚLTIMA MANUT", "VENC MANUT", "NF ASSINADA"]
    df_excel = df[cols]; df_excel.columns = display_cols; df_excel["PERÍODO"] = df_excel["PERÍODO"].apply(lambda x: f"{x} dias" if pd.notna(x) else "N/A")
    buffer = BytesIO(); df_excel.to_excel(buffer, index=False, engine="openpyxl"); buffer.seek(0)
    return buffer

def main():
    st.markdown('<h1 class="main-title">Controle de Bombas de Sucção</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Desenvolvido por Guilherme Rodrigues – Suplen Médical</p>', unsafe_allow_html=True)
    
    if 'bomba_edit_key' not in st.session_state:
        st.session_state.bomba_edit_key = 0

    filial = setup_filial()
    if not filial and not st.session_state.get("general_mode", False):
        st.sidebar.warning("Por favor, selecione e confirme uma filial para continuar.")
        return
    if filial:
        st.markdown(f'<p class="filial-main">Filial: {filial}</p>', unsafe_allow_html=True)
    else:
        st.markdown(f'<p class="filial-main">Dashboard Geral</p>', unsafe_allow_html=True)
    if "messages" not in st.session_state:
        st.session_state.messages = []
    for msg in st.session_state.messages:
        st.toast(msg['text'], icon=msg['icon'])
    st.session_state.messages = []
    menu = ["Dashboard", "Registrar", "Bombas em Comodato", "Devolver", "Manutenção de Bombas", "Histórico Devolvidas", "Saldo Curativo"]
    if st.session_state.get("general_mode", False):
        menu = ["Dashboard Geral"]
    choice = st.sidebar.selectbox("Navegação", menu, format_func=lambda x: f"📋 {x}")

    if choice == "Dashboard":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Dashboard da Filial</h2>', unsafe_allow_html=True)
        with st.spinner("Carregando métricas da filial..."):
            metrics = get_dashboard_metrics(filial)
        if not metrics:
            st.error("Erro ao carregar métricas do dashboard.")
            return
        st.markdown("##### Visão Geral")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("BOMBAS EM COMODATO", metrics["ativas"])
        col2.metric("Disponíveis", metrics["disponiveis"])
        col3.metric("Em Manutenção", metrics["em_manutencao"])
        col4.metric("Bombas ULTA", metrics.get("ulta_count", "N/A"))
        col5.metric("Bombas ACTIVAC", metrics.get("activac_count", "N/A"))
        
        st.markdown("<h5 style='text-align: center; margin-top: 20px;'>Bombas Disponíveis</h5>", unsafe_allow_html=True)
        modelos_disponiveis = metrics.get("modelos_disponiveis", {})
        ulta_disp = modelos_disponiveis.get("ULTA", 0)
        activac_disp = modelos_disponiveis.get("ACTIVAC", 0)
        outros_disp_sum = sum(v for k, v in modelos_disponiveis.items() if k not in ["ULTA", "ACTIVAC"])
        if outros_disp_sum > 0:
            disp_col1, disp_col2, disp_col3 = st.columns(3)
            disp_col1.metric("ULTA Disponíveis", ulta_disp)
            disp_col2.metric("ACTIVAC Disponíveis", activac_disp)
            disp_col3.metric("Outros Modelos", outros_disp_sum)
        else:
            _, disp_col1, disp_col2, _ = st.columns([1, 2, 2, 1])
            disp_col1.metric("ULTA Disponíveis", ulta_disp)
            disp_col2.metric("ACTIVAC Disponíveis", activac_disp)
        
        st.markdown("---")
        col_graph1, col_graph2 = st.columns(2)
        with col_graph1:
            st.markdown("#### Status das Bombas em Comodato")
            df_status = pd.DataFrame(list(metrics["status_counts"].items()), columns=['Status', 'Quantidade'])
            df_status = df_status[df_status['Quantidade'] > 0]
            if not df_status.empty:
                df_status['Status_Label'] = df_status['Status'].apply(lambda x: f"🟢 {x}" if x == "No Prazo" else f"🟡 {x}" if x == "Menos de 7 dias" else f"🔴 {x}" if x == "Fora Prazo" else x)
                fig_pie = px.pie(df_status, names='Status_Label', values='Quantidade', title="Distribuição por Status", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Set2)
                fig_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sem bombas em comodato para exibir.")
        with col_graph2:
            st.markdown("#### Bombas em Comodato por Hospital")
            df_hosp = pd.DataFrame(list(metrics["hosp_counts"].items()), columns=['hospital', 'count'])
            if not df_hosp.empty:
                df_hosp['hospital'] = df_hosp['hospital'].replace('', 'HOSPITAL NÃO ESPECIFICADO')
                df_hosp['hospital_label'] = df_hosp['hospital'].apply(lambda x: x[:30] + '...' if len(x) > 30 else x)
                fig_bar = px.bar(df_hosp.sort_values('count', ascending=False).head(15), x='hospital_label', y='count', title="Top 15 Hospitais com Bombas", color='hospital_label', text='count', template='plotly_dark', labels={'hospital_label': 'Hospital', 'count': 'Nº de Bombas'}, color_discrete_sequence=px.colors.qualitative.Dark24)
                fig_bar.update_layout(showlegend=False, xaxis_tickangle=-45, xaxis=dict(tickfont=dict(size=10), automargin=True))
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Sem bombas em comodato para exibir.")

    elif choice == "Dashboard Geral":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Dashboard Geral (Todas as Filiais)</h2>', unsafe_allow_html=True)
        filial_filter = st.selectbox("Filtrar por Filial", ["Todas"] + FILIAIS, key="filial_filter")
        filial_to_query = filial_filter if filial_filter != "Todas" else None
        with st.spinner("Carregando métricas gerais..."):
            metrics = get_dashboard_metrics(filial_to_query)
        if not metrics:
            st.error("Erro ao carregar métricas do dashboard.")
            return
        st.markdown("#### Métricas Gerais do Inventário")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total de Bombas Registradas", metrics["total_bombas"])
        col2.metric("Bombas em Comodato", metrics["ativas"])
        col3.metric("Disponíveis", metrics["disponiveis"])
        col4.metric("Em Manutenção", metrics["em_manutencao"])
        
        st.markdown("<h5 style='text-align: center; margin-top: 20px;'>Bombas Disponíveis (Geral)</h5>", unsafe_allow_html=True)
        modelos_disponiveis = metrics.get("modelos_disponiveis", {})
        ulta_disp = modelos_disponiveis.get("ULTA", 0)
        activac_disp = modelos_disponiveis.get("ACTIVAC", 0)
        outros_disp_sum = sum(v for k, v in modelos_disponiveis.items() if k not in ["ULTA", "ACTIVAC"])
        if outros_disp_sum > 0:
            disp_col1, disp_col2, disp_col3 = st.columns(3)
            disp_col1.metric("ULTA Disponíveis", ulta_disp)
            disp_col2.metric("ACTIVAC Disponíveis", activac_disp)
            disp_col3.metric("Outros Modelos", outros_disp_sum)
        else:
            _, disp_col1, disp_col2, _ = st.columns([1, 2, 2, 1])
            disp_col1.metric("ULTA Disponíveis", ulta_disp)
            disp_col2.metric("ACTIVAC Disponíveis", activac_disp)

        st.markdown("---")
        st.markdown("### Análise das Bombas em Comodato")
        col_graph1, col_graph2 = st.columns(2)
        with col_graph1:
            st.markdown("##### Distribuição por Status")
            df_status_global = pd.DataFrame(list(metrics["status_counts"].items()), columns=['Status', 'Quantidade'])
            df_status_global = df_status_global[df_status_global['Quantidade'] > 0]
            if not df_status_global.empty:
                df_status_global['Status_Label'] = df_status_global['Status'].apply(lambda x: f"🟢 {x}" if x == "No Prazo" else f"🟡 {x}" if x == "Menos de 7 dias" else f"🔴 {x}" if x == "Fora Prazo" else x)
                fig_pie_global = px.pie(df_status_global, names='Status_Label', values='Quantidade', title="Distribuição por Status", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Set2)
                fig_pie_global.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                st.plotly_chart(fig_pie_global, use_container_width=True)
            else:
                st.info("Nenhuma bomba em comodato para exibir status.")
        with col_graph2:
            st.markdown("##### Distribuição por Hospital")
            df_hosp_global = pd.DataFrame(list(metrics["hosp_counts"].items()), columns=['hospital', 'count'])
            if not df_hosp_global.empty:
                df_hosp_global['hospital'] = df_hosp_global['hospital'].replace('', 'HOSPITAL NÃO ESPECIFICADO')
                df_hosp_global['hospital_label'] = df_hosp_global['hospital'].apply(lambda x: x[:30] + '...' if len(x) > 30 else x)
                fig_bar_global = px.bar(df_hosp_global.sort_values('count', ascending=False).head(15), x='hospital_label', y='count', title="Top 15 Hospitais", color='hospital_label', text='count', template='plotly_dark', labels={'hospital_label': 'Hospital', 'count': 'Nº de Bombas'}, color_discrete_sequence=px.colors.qualitative.Dark24)
                fig_bar_global.update_layout(showlegend=False, xaxis_tickangle=-45, xaxis=dict(tickfont=dict(size=10), automargin=True))
                st.plotly_chart(fig_bar_global, use_container_width=True)
            else:
                st.info("Nenhuma bomba em comodato para exibir por hospital.")
        st.markdown("---")
        st.markdown("### Análise do Inventário e Manutenção")
        with st.spinner("Analisando dados de manutenção de todas as bombas..."):
            dados_bombas_df = get_dados_bombas_df()
        if dados_bombas_df.empty:
            st.warning("Dados de manutenção não encontrados. Gráficos de inventário não podem ser gerados.")
        else:
            current_date = pd.to_datetime(datetime.now())
            overdue_count = dados_bombas_df[dados_bombas_df['Venc_Manut'] < current_date].shape[0]
            non_expired_df = dados_bombas_df[dados_bombas_df['Venc_Manut'] >= current_date]
            non_expired_model_counts = non_expired_df['Modelo'].str.upper().value_counts().to_dict()
            ulta_non_expired = non_expired_model_counts.get('ULTA', 0)
            activac_non_expired = non_expired_model_counts.get('ACTIVAC', 0)
            model_counts = dados_bombas_df['Modelo'].value_counts()
            model_df = pd.DataFrame({'Modelo': model_counts.index, 'Quantidade': model_counts.values})
            st.markdown("##### Manutenção Preventiva")
            m_col1, m_col2, m_col3 = st.columns(3)
            m_col1.metric("Bombas com Manutenção Vencida", overdue_count)
            m_col2.metric("ULTA com Manutenção Válida", ulta_non_expired)
            m_col3.metric("ACTIVAC com Manutenção Válida", activac_non_expired)
            st.markdown("##### Distribuição de Modelos no Inventário Total")
            if not model_df.empty:
                fig_model = px.pie(model_df, names='Modelo', values='Quantidade', title="Distribuição de Bombas por Modelo", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Set2)
                fig_model.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                st.plotly_chart(fig_model, use_container_width=True)
            else:
                st.info("Sem dados de modelos para exibir.")
        st.markdown("---")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown("### Mapa de Bombas por Filial (Todas as Filiais)")
            m = folium.Map(location=[-15.0, -55.0], zoom_start=4, tiles="CartoDB Dark_Matter")
            coordinates = {"Brasília": (-15.7942, -47.8822), "Goiânia": (-16.6869, -49.2648), "Cuiabá": (-15.6014, -56.0979)}
            colors = {"Brasília": "#1E90FF", "Goiânia": "#32CD32", "Cuiabá": "#FFD700"}
            for filial_map, (lat, lon) in coordinates.items():
                normalized_filial = normalize_text(filial_map)
                count = metrics["bombas_por_filial"].get(normalized_filial, 0)
                if count > 0:
                    radius = max(10000, count * 15000)
                    folium.Circle(location=[lat, lon], radius=radius, fill=True, fill_opacity=0.3, color=colors.get(filial_map, 'gray'), fill_color=colors.get(filial_map, 'gray'), popup=folium.Popup(f"{filial_map}: {count} bombas em comodato", max_width=200)).add_to(m)
            st_folium(m, width=800, height=400)
        with col2:
            legend_html = """<div style="height: 100%; display: flex; flex-direction: column; justify-content: flex-start;"><h3>Legenda</h3>"""
            items_added = False
            for filial_map, color in colors.items():
                normalized_filial = normalize_text(filial_map)
                display_count = metrics["bombas_por_filial"].get(normalized_filial, 0)
                if display_count > 0:
                    items_added = True
                    legend_html += (f'<div class="legend-item"><span class="legend-dot" style="background-color: {color};"></span>{filial_map}: {display_count}</div>')
            if not items_added:
                legend_html += "<div>Nenhuma bomba em comodato registrada.</div>"
            legend_html += "</div>"
            st.markdown(legend_html, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### Total de Bombas por Modelo em Cada Filial")
        if metrics.get("modelos_por_filial"):
            df_modelos_filial = pd.DataFrame(metrics["modelos_por_filial"])
            if not df_modelos_filial.empty:
                df_modelos_filial = df_modelos_filial.sort_values(by=['filial', 'Modelo'])
                fig_modelos_filial = px.bar(df_modelos_filial, x='filial', y='Quantidade', color='Modelo', title='Distribuição de Modelos de Bombas por Filial', barmode='group', text_auto=True, template='plotly_dark', labels={'filial': 'Filial', 'Quantidade': 'Nº de Bombas', 'Modelo': 'Modelo da Bomba'}, color_discrete_sequence=px.colors.qualitative.Vivid)
                fig_modelos_filial.update_layout(xaxis_title="Filial", yaxis_title="Total de Bombas", legend_title_text='Modelo')
                st.plotly_chart(fig_modelos_filial, use_container_width=True)
            else:
                st.info("Não há dados de modelos por filial para exibir.")
        else:
            st.info("Métricas de modelos por filial não disponíveis.")
        st.markdown('<h2 class="section-title">Análise de Curativos</h2>', unsafe_allow_html=True)
        with st.spinner("Carregando análise de curativos..."):
            curativo_results = analyze_curativo()
            if curativo_results.get('last_updated'):
                ts_utc = parse_supabase_date(curativo_results.get('last_updated'))
                if ts_utc:
                    ts_local = ts_utc.astimezone(timezone(timedelta(hours=-3)))
                    formatted_ts = ts_local.strftime('%d/%m/%Y às %H:%M')
                    st.caption(f"🗓️ _Dados da planilha atualizados em: **{formatted_ts}**_")
            if curativo_results.get('error'):
                st.error(curativo_results['error'])
            else:
                status_to_exclude = ['Disponível', 'Não Encontrado']
                status_df = curativo_results.get('status_df', pd.DataFrame())
                if not status_df.empty:
                    status_df = status_df[~status_df['Status'].isin(status_to_exclude)]
                revenue_status_df = curativo_results.get('revenue_status_df', pd.DataFrame())
                if not revenue_status_df.empty:
                    revenue_status_df = revenue_status_df[~revenue_status_df['Status'].isin(status_to_exclude)]
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("#### Porcentagem de Vendas por Status")
                    if not status_df.empty:
                        fig_status = px.pie(status_df, names='Status', values='Percentage', title="Distribuição de Vendas por Status", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_status.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                        st.plotly_chart(fig_status, use_container_width=True)
                    else:
                        st.info("Sem dados de status para exibir.")
                with col2:
                    st.markdown("#### Receita por Status")
                    if not revenue_status_df.empty:
                        revenue_status_df['Revenue_Formatted'] = revenue_status_df['Revenue'].apply(lambda x: f"R$ {x:,.2f}")
                        fig_revenue = px.bar(revenue_status_df, x='Status', y='Revenue', title="Receita por Status", text='Revenue_Formatted', template='plotly_dark', color='Status', color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_revenue.update_layout(showlegend=False, xaxis_tickangle=45, yaxis_title="Receita (R$)")
                        st.plotly_chart(fig_revenue, use_container_width=True)
                    else:
                        st.info("Sem dados de receita por status para exibir.")

    elif choice == "Registrar":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Registrar Nova Bomba</h2>', unsafe_allow_html=True)
        with st.form("registrar_form", clear_on_submit=True):
            st.markdown('<div class="form-container">', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                st.markdown('<div class="form-column">', unsafe_allow_html=True)
                serial = st.text_input("🔢 SERIAL*", placeholder="Ex.: ABC123").upper()
                paciente = st.text_input("👤 PACIENTE*", placeholder="Ex.: João da Silva").upper()
                convenio = st.text_input("💳 CONVÊNIO*", placeholder="Ex.: Unimed").upper()
                periodo = st.text_input("⏳ PERÍODO (dias)*", placeholder="Ex.: 30")
                pedido = st.text_input("📝 PEDIDO", placeholder="Ex.: PED789").upper()
                st.markdown('</div>', unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="form-column">', unsafe_allow_html=True)
                hospital = st.text_input("🏥 HOSPITAL*", placeholder="Ex.: Hospital Central").upper()
                medico = st.text_input("🩺 MÉDICO*", placeholder="Ex.: Dr. José").upper()
                data_saida = st.date_input("📅 DATA SAÍDA* (DD/MM/YYYY)", value=None, format="DD/MM/YYYY")
                nf = st.text_input("🧾 NF", placeholder="Ex.: 123456").upper()
                st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
            submit_button = st.form_submit_button("Registrar Bomba")
            if submit_button:
                if not all([serial, hospital, paciente, medico, convenio, data_saida, periodo]):
                    st.error("ERRO: Preencha todos os campos obrigatórios (*).")
                elif not periodo.isdigit() or int(periodo) <= 0:
                    st.error("ERRO: O campo 'PERÍODO' deve ser um número inteiro de dias, maior que zero.")
                else:
                    try:
                        with st.spinner("Verificando e registrando..."):
                            existing = supabase.table("bombas").select("id", count='exact').eq("serial", serial).eq("ativo", True).execute()
                            if existing.count > 0:
                                st.error(f"ERRO: A bomba com o serial '{serial}' já está registrada e ativa.")
                            else:
                                data_saida_fmt = data_saida.strftime("%Y-%m-%d")
                                status_inicial = calculate_status(data_saida, periodo)
                                new_record = {"serial": serial, "hospital": hospital, "paciente": paciente, "medico": medico, "convenio": convenio, "data_registro": datetime.now().strftime("%Y-%m-%d"), "data_saida": data_saida_fmt, "periodo": int(periodo), "status": status_inicial, "nf": nf, "pedido": pedido, "ativo": True, "filial": filial, "nf_devolucao": ""}
                                response = supabase.table("bombas").insert(new_record).execute()
                                if response.data:
                                    bomba_id = response.data[0]["id"]
                                    register_event("bombas", bomba_id, f"BOMBA REGISTRADA (SERIAL: {serial})", filial)
                                    flush_events()
                                    st.session_state.messages.append({"text": "Bomba registrada com sucesso!", "icon": "✅"})
                                    st.cache_data.clear()
                                    st.rerun()
                                else:
                                    st.error(f"Ocorreu um erro ao registrar a bomba. Detalhes: {response.error}")
                    except Exception as e:
                        logging.error(f"Erro no registro da bomba: {e}")
                        st.error(f"ERRO AO REGISTRAR: {e}")

    elif choice == "Bombas em Comodato":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Bombas em Comodato</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar bombas em comodato (por Serial, Paciente ou Hospital)...", key="listagem_search")
        with st.spinner("Carregando bombas..."):
            bombas_ativas = get_bombas(search_term, filial, active_only=True)
            bombas_df = get_dados_bombas_df()
            nf_map = get_all_nfs_assinadas_info()
        display_bombas_table("Listagem de Bombas Ativas", bombas_ativas, bombas_df, nf_map)
        st.markdown("---")
        excel_buffer = generate_excel_bombas_ativas(bombas_ativas, bombas_df, filial, nf_map)
        if excel_buffer:
            st.download_button(label="✅ Baixar Listagem em Comodato (Excel)", data=excel_buffer, file_name=f"bombas_comodato_{filial}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("---")
        st.markdown("### Gerenciamento de Bombas")
        if not bombas_ativas:
            st.info("Não há bombas ativas para gerenciar.")
        else:
            tab_edit, tab_pdf, tab_anexar_nf, tab_download_nf = st.tabs(["📝 Editar Bomba", "🖨️ Gerar Documentos", "📎 Anexar NF Assinada", "📥 Baixar NF Assinada"])
            with tab_edit:
                st.subheader("Editar Dados da Bomba")
                bomba_to_edit = st.selectbox("Selecione uma bomba para editar:", options=bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key=f"edit_bomba_select_{st.session_state.bomba_edit_key}", index=None, placeholder="Selecione uma bomba...")
                if bomba_to_edit:
                    with st.form(key=f"edit_form_{bomba_to_edit['id']}"):
                        try:
                            # Tenta converter a data de string (dd/mm/YYYY) para objeto date
                            dt_obj = datetime.strptime(bomba_to_edit.get('data_saida'), '%d/%m/%Y').date()
                            current_data_saida = dt_obj
                        except (ValueError, TypeError):
                            current_data_saida = None
                        c1, c2 = st.columns(2)
                        with c1:
                            serial = st.text_input("Serial*", value=bomba_to_edit.get('serial', '')).upper()
                            paciente = st.text_input("Paciente", value=bomba_to_edit.get('paciente', '')).upper()
                            hospital = st.text_input("Hospital", value=bomba_to_edit.get('hospital', '')).upper()
                            medico = st.text_input("Médico", value=bomba_to_edit.get('medico', '')).upper()
                        with c2:
                            convenio = st.text_input("Convênio", value=bomba_to_edit.get('convenio', '')).upper()
                            periodo = st.text_input("Período (dias)*", value=bomba_to_edit.get('periodo', ''))
                            pedido = st.text_input("Pedido", value=bomba_to_edit.get('pedido', '')).upper()
                            nf = st.text_input("NF", value=bomba_to_edit.get('nf', '')).upper()
                        data_saida = st.date_input("Data de Saída*", value=current_data_saida, format="DD/MM/YYYY")
                        submitted = st.form_submit_button("Salvar Alterações")
                        if submitted:
                            if not all([serial, data_saida, periodo]) or not str(periodo).isdigit():
                                st.error("ERRO: Preencha os campos obrigatórios (*) e verifique se o período é um número.")
                            else:
                                with st.spinner("Atualizando dados..."):
                                    try:
                                        data_saida_fmt = data_saida.strftime("%Y-%m-%d")
                                        # Passa o objeto 'data_saida' (date) que já é do tipo correto
                                        new_status = calculate_status(data_saida, periodo)
                                        update_data = {"serial": serial, "paciente": paciente, "hospital": hospital, "medico": medico, "convenio": convenio, "periodo": int(periodo), "pedido": pedido, "nf": nf, "data_saida": data_saida_fmt, "status": new_status}
                                        supabase.table("bombas").update(update_data).eq("id", bomba_to_edit['id']).execute()
                                        register_event("bombas", bomba_to_edit['id'], f"DADOS DA BOMBA ATUALIZADOS (SERIAL: {serial})", filial)
                                        flush_events()
                                        st.session_state.messages.append({"text": "Dados da bomba atualizados!", "icon": "📝"})
                                        st.cache_data.clear()
                                        st.session_state.bomba_edit_key += 1
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Ocorreu um erro ao salvar as alterações: {e}")
            with tab_pdf:
                st.subheader("Gerar PDF do Contrato")
                bomba_pdf = st.selectbox("Selecione uma bomba:", bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="pdf_bomba_select", index=None, placeholder="Selecione uma bomba...")
                if bomba_pdf and st.button("Gerar e Baixar PDF", key="generate_pdf_button"):
                    with st.spinner("Gerando PDF... Esta operação pode ser demorada."):
                        pdf_bytes = generate_combined_pdf(bomba_pdf)
                        if pdf_bytes:
                            st.session_state.pdf_to_download = {"data": pdf_bytes, "name": f"documentos_{bomba_pdf['serial']}.pdf"}
                            register_event("bombas", bomba_pdf['id'], "DOCUMENTOS GERADOS", filial)
                            flush_events()
                        else: st.error("Falha ao gerar o PDF.")
                if "pdf_to_download" in st.session_state:
                    st.download_button(label="Clique aqui para Baixar o PDF Gerado", data=st.session_state.pdf_to_download["data"], file_name=st.session_state.pdf_to_download["name"], mime="application/pdf")
                    del st.session_state.pdf_to_download
            with tab_anexar_nf:
                st.subheader("Anexar NF Assinada")
                bombas_sem_nf = [b for b in bombas_ativas if not check_nf_assinada(b['serial'], nf_map)]
                if not bombas_sem_nf: st.info("Todas as bombas nesta listagem já possuem NF assinada.")
                else:
                    bomba_anexar = st.selectbox("Selecione a bomba para anexar a NF:", bombas_sem_nf, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="anexar_nf_select")
                    nf_file = st.file_uploader("Anexar PDF da NF assinada", type=["pdf"], key="nf_assinada_upload")
                    if st.button("Enviar NF Assinada", disabled=(not bomba_anexar or not nf_file)):
                        with st.spinner("Enviando NF..."):
                            if upload_nf_assinada(bomba_anexar, nf_file):
                                register_event("bombas", bomba_anexar["id"], f"NF ASSINADA ENVIADA (SERIAL: {bomba_anexar['serial']})", filial)
                                flush_events(); st.session_state.messages.append({"text": "NF assinada enviada com sucesso!", "icon": "📎"}); st.cache_data.clear(); st.rerun()
            with tab_download_nf:
                st.subheader("Baixar NF Assinada")
                bombas_com_nf = [b for b in bombas_ativas if check_nf_assinada(b['serial'], nf_map)]
                if not bombas_com_nf: st.info("Nenhuma bomba nesta listagem possui NF assinada para download.")
                else:
                    bomba_selecionada = st.selectbox("Selecione a bomba para baixar a NF:", bombas_com_nf, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="download_nf_select")
                    if st.button("Baixar NF Assinada", key="baixar_nf_button_tab", disabled=not bomba_selecionada):
                        with st.spinner("Preparando download..."):
                            nf_data = download_nf_assinada(bomba_selecionada['serial'])
                            if nf_data:
                                nf_filename_path = get_nf_assinada_filename(bomba_selecionada['serial'], nf_map)
                                st.download_button(label="Clique aqui para baixar", data=nf_data, file_name=os.path.basename(nf_filename_path), mime="application/pdf")
                            else: st.error("Erro ao encontrar a NF para download.")

    elif choice == "Devolver":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Devolver Bomba</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar bomba para devolver (Serial, Paciente, Hospital)...", key="devolver_search")
        bombas_ativas = get_bombas(search_term, filial, active_only=True)
        if not bombas_ativas:
            st.info("Nenhuma bomba ativa encontrada com este critério de busca.")
        else:
            with st.form("devolver_form"):
                bomba = st.selectbox("Selecione a bomba a devolver:", bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", index=None, placeholder="Selecione uma bomba...")
                data_retorno = st.date_input("📅 Data de Retorno", value=datetime.now(), format="DD/MM/YYYY")
                nf_devolucao = st.text_input("🧾 NF de Devolução*", placeholder="Ex.: 987654").upper()
                submitted = st.form_submit_button("Confirmar Devolução")
                if submitted:
                    if not nf_devolucao or not bomba:
                        st.error("Selecione uma bomba e preencha a NF de Devolução.")
                    else:
                        try:
                            with st.spinner("Registrando devolução e desvinculando NF antiga..."):
                                serial_devolvido = bomba['serial']
                                nf_map = get_all_nfs_assinadas_info()
                                nf_assinada_path = get_nf_assinada_filename(serial_devolvido, nf_map)
                                if nf_assinada_path:
                                    try:
                                        supabase.storage.from_("controle-de-bombas-suplen-files").remove([nf_assinada_path])
                                        msg_sucesso = f"Bomba {serial_devolvido} devolvida e NF assinada anterior desvinculada!"
                                    except Exception as e_storage:
                                        st.warning(f"Atenção: Erro ao remover a NF assinada anterior: {e_storage}")
                                        msg_sucesso = f"Bomba {serial_devolvido} devolvida (com aviso)."
                                else:
                                    msg_sucesso = "Bomba devolvida com sucesso!"
                                supabase.table("bombas").update({"ativo": False, "status": "✅ DEVOLVIDA", "data_retorno": data_retorno.strftime("%Y-%m-%d"), "nf_devolucao": nf_devolucao}).eq("id", bomba["id"]).execute()
                                register_event("bombas", bomba["id"], f"BOMBA DEVOLVIDA (SERIAL: {bomba['serial']}) NF: {nf_devolucao}", filial)
                                flush_events()
                                st.session_state.messages.append({"text": msg_sucesso, "icon": "✅"})
                                st.cache_data.clear()
                                st.rerun()
                        except Exception as e:
                            logging.error(f"Erro ao devolver bomba: {e}")
                            st.error(f"Erro ao registrar devolução: {e}")

    elif choice == "Manutenção de Bombas":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Manutenção de Bombas</h2>', unsafe_allow_html=True)
        tabs = st.tabs(["Cadastrar Manutenção", "Listagem de Manutenções"])
        with tabs[0]:
            st.markdown("### Cadastrar Bomba em Manutenção")
            with st.form("manutencao_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    serial = st.text_input("🔢 SERIAL*", placeholder="Ex.: ABC123").upper()
                    defeito = st.text_area("📝 DEFEITO APRESENTADO*", placeholder="Descreva o defeito").upper()
                with c2:
                    nf_numero = st.text_input("🧾 NÚMERO DA NF EMITIDA*", placeholder="Ex.: 123456").upper()
                    nf_file = st.file_uploader("📄 UPLOAD NF (PDF)*", type=["pdf"], key="nf_upload")
                submit_button = st.form_submit_button("Registrar Manutenção")
                if submit_button:
                    if not all([serial, defeito, nf_numero, nf_file]):
                        st.error("Preencha todos os campos obrigatórios (*).")
                    else:
                        with st.spinner("Verificando e registrando..."):
                            existing = supabase.table("manutencao").select("id").eq("serial", serial).eq("status", "Em Manutenção").execute()
                            if existing.data:
                                st.error(f"ERRO: Serial '{serial}' já está em manutenção!")
                            else:
                                data_registro_str = datetime.now().strftime("%d/%m/%Y")
                                if upload_nf_pdf(serial, data_registro_str, nf_file):
                                    response = supabase.table("manutencao").insert({"serial": serial, "defeito": defeito, "data_registro": datetime.now().strftime("%Y-%m-%d"), "nf_numero": nf_numero, "nf_status": "Enviada", "status": "Em Manutenção", "filial": filial}).execute()
                                    manutencao_id = response.data[0]["id"]
                                    register_event("manutencao", manutencao_id, f"MANUTENÇÃO REGISTRADA (SERIAL: {serial})", filial)
                                    flush_events()
                                    st.session_state.messages.append({"text": "Manutenção registrada!", "icon": "🛠️"})
                                    st.cache_data.clear()
                                    st.rerun()
        with tabs[1]:
            st.markdown("### Listagem de Bombas em Manutenção")
            search_term_manut = st.text_input("Pesquisar manutenções...", key="manutencao_search")
            with st.spinner("Carregando manutenções..."):
                manutencoes = get_manutencao(search_term_manut, filial)
            display_manutencao_table("Manutenções Registradas", manutencoes)

    elif choice == "Histórico Devolvidas":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Histórico de Bombas Devolvidas</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar no histórico (serial, NF, etc.)...", key="historico_search")
        historico = get_historico_devolvidas(filial if not st.session_state.get("general_mode", False) else None)
        if not historico:
            st.info("Nenhuma bomba devolvida registrada.")
        else:
            df = pd.DataFrame(historico, columns=["data_evento", "descricao", "filial"])
            df.columns = ["Data do Evento", "Descrição", "Filial"]
            if search_term:
                df = df[df.apply(lambda row: search_term.lower() in str(row.values).lower(), axis=1)]
            st.dataframe(df, use_container_width=True, hide_index=True)

    elif choice == "Saldo Curativo":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Saldo de Curativos</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar por Descrição do Produto...", key="saldo_curativo_search")
        st.markdown("---")
        st.markdown("<h6>Legenda de Validade</h6>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.markdown('<div style="background-color: #90ee90; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>NORMAL</strong><br>&gt; 3 meses</div>', unsafe_allow_html=True)
        col2.markdown('<div style="background-color: #FFD580; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>ATENÇÃO</strong><br>2 a 3 meses</div>', unsafe_allow_html=True)
        col3.markdown('<div style="background-color: #F08080; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>CRÍTICO</strong><br>&lt; 2 meses</div>', unsafe_allow_html=True)
        st.markdown("---")
        df_curativos = get_saldo_curativo_data()
        if df_curativos.empty:
            st.info("Nenhum dado de saldo de curativo encontrado.")
        else:
            df_para_exibir = df_curativos.copy()
            if search_term:
                df_para_exibir = df_para_exibir[df_para_exibir['Desc_Produto'].str.contains(search_term, case=False, na=False)]
            if not df_para_exibir.empty:
                # Garante que a Data_Validad seja tratada corretamente
                if 'Data_Validad' in df_para_exibir.columns:
                    df_para_exibir['Data_Validad'] = pd.to_datetime(df_para_exibir['Data_Validad'], errors='coerce')
                    df_para_exibir = df_para_exibir.sort_values(by=['Data_Validad', 'Desc_Produto'])
                    colunas_exibir = ['Produto', 'Desc_Produto', 'Referencia', 'Lote', 'Data_Validad', 'Saldo_Lote']
                    df_display = df_para_exibir[colunas_exibir].copy()
                    df_display.rename(columns={
                        'Desc_Produto': 'Descrição do Produto',
                        'Referencia': 'Referência',
                        'Data_Validad': 'Data de Validade',
                        'Saldo_Lote': 'Saldo do Lote'
                    }, inplace=True)
                else:
                    # Se não tem coluna Data_Validad, ordenar só por Desc_Produto
                    df_para_exibir = df_para_exibir.sort_values(by=['Desc_Produto'])
                    colunas_exibir = ['Produto', 'Desc_Produto', 'Referencia', 'Lote', 'Saldo_Lote']
                    df_display = df_para_exibir[colunas_exibir].copy()
                    df_display.rename(columns={
                        'Desc_Produto': 'Descrição do Produto',
                        'Referencia': 'Referência',
                        'Saldo_Lote': 'Saldo do Lote'
                    }, inplace=True)
                def style_validade(row):
                    if 'Data de Validade' in row.index:
                        hoje = datetime.now()
                        if pd.notna(row['Data de Validade']):
                            diferenca_dias = (row['Data de Validade'] - hoje).days
                            bg_color = 'background-color: #F08080;'  # Crítico (< 2 meses)
                            if diferenca_dias > 90:
                                bg_color = 'background-color: #90ee90;'  # Normal
                            elif 60 <= diferenca_dias <= 90:
                                bg_color = 'background-color: #FFD580;'  # Atenção
                            return [f"color: black; {bg_color}"] * len(row)
                    return [''] * len(row)  # Sem estilo para linhas sem data de validade

                if 'Data de Validade' in df_display.columns:
                    styler = df_display.style.apply(style_validade, axis=1).format({
                        'Data de Validade': lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else '',
                        'Saldo do Lote': '{:.0f}'
                    })
                else:
                    styler = df_display.style.format({'Saldo do Lote': '{:.0f}'})
                    
                st.dataframe(styler, use_container_width=True, hide_index=True)
                excel_data = generate_excel_saldo_curativo(styler.data)
                if excel_data:
                    st.download_button(label="✅ Baixar Saldo em Excel", data=excel_data, file_name=f"saldo_curativo_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Nenhum produto encontrado com os termos da busca.")

if __name__ == "__main__":
    main()