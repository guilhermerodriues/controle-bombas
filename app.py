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
from docx2pdf import convert

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

# -------------------- FUNÇÕES AUXILIARES --------------------
@st.cache_data(ttl=300)
def download_file_from_storage(storage_path):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        response = supabase.storage.from_(bucket_name).download(storage_path)
        logging.info(f"Arquivo {storage_path} baixado com sucesso.")
        return response
    except Exception as e:
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

# -------------------- LÓGICA DE DADOS --------------------
def calculate_status(data_saida, periodo):
    if not data_saida or not periodo:
        return "Indefinido"
    try:
        dt_saida = datetime.strptime(data_saida, "%Y-%m-%d")
        data_devolucao = dt_saida + timedelta(days=int(periodo))
        dias_restantes = (data_devolucao - datetime.now()).days
        if dias_restantes < 0:
            return "Fora Prazo"
        elif dias_restantes <= 7:
            return "Menos de 7 dias"
        return "No Prazo"
    except (ValueError, TypeError):
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
        
        result = []
        search_term = search_term.lower()
        for bomba in bombas:
            bomba["id"] = str(bomba["id"])
            for field in ["data_saida", "data_registro", "data_retorno"]:
                if bomba.get(field):
                    try:
                        bomba[field] = datetime.strptime(bomba[field], "%Y-%m-%d").strftime("%d/%m/%Y")
                    except ValueError:
                        pass
            if search_term and not any(
                search_term in str(bomba.get(field, "")).lower()
                for field in ["serial", "paciente", "hospital"]
            ):
                continue
            result.append(bomba)
        logging.info(f"Busca de bombas retornou {len(result)} registros.")
        return result
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
        bombas_df = get_dados_bombas_df()

        result = []
        search_term = search_term.lower()
        for manut in manutencoes:
            manut["id"] = str(manut["id"])
            if manut.get("data_registro"):
                try:
                    manut["data_registro"] = datetime.strptime(manut["data_registro"], "%Y-%m-%d").strftime("%d/%m/%Y")
                except ValueError:
                    pass
            
            if search_term and not any(
                search_term in str(manut.get(field, "")).lower()
                for field in ["serial", "defeito", "nf_numero"]
            ):
                continue
            
            serial_normalizado = normalize_text(manut.get('serial', ''))
            match = bombas_df[bombas_df['Serial_Normalized'] == serial_normalizado] if not bombas_df.empty and serial_normalizado else pd.DataFrame()
            
            manut['modelo'] = match['Modelo'].iloc[0] if not match.empty else "N/A"
            manut['ultima_manut'] = match['Ultima_Manut'].iloc[0].strftime('%d/%m/%Y') if not match.empty and pd.notna(match['Ultima_Manut'].iloc[0]) else "N/A"
            manut['venc_manut'] = match['Venc_Manut'].iloc[0].strftime('%d/%m/%Y') if not match.empty and pd.notna(match['Venc_Manut'].iloc[0]) else "N/A"
            result.append(manut)
        logging.info(f"Busca de manutenções retornou {len(result)} registros.")
        return result
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
            if doc.get("data_evento"):
                try:
                    doc["data_evento"] = datetime.strptime(doc["data_evento"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y, %H:%M")
                except ValueError:
                    pass
        logging.info(f"Histórico de devolvidas retornou {len(historico)} registros.")
        return historico
    except Exception as e:
        logging.error(f"Erro ao buscar histórico de devolvidas: {e}")
        st.error(f"Erro ao acessar histórico: {e}")
        return []

@st.cache_data(ttl=300, show_spinner="Carregando dados de curativos...")
def get_saldo_curativo_data():
    """Busca todos os dados da tabela saldo_curativo."""
    try:
        response = supabase.table("saldo_curativo").select("*").execute()
        data = response.data
        if not data:
            logging.warning("A tabela saldo_curativo está vazia ou não foi encontrada.")
            return pd.DataFrame()
        
        df = pd.DataFrame(data)
        logging.info(f"Dados da tabela saldo_curativo carregados: {len(df)} registros.")
        return df
    except Exception as e:
        logging.error(f"Erro ao carregar dados da tabela saldo_curativo: {e}")
        st.error(f"Não foi possível carregar os dados de saldo de curativos: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def get_dashboard_metrics(filial=None):
    try:
        # --- LÓGICA DE DADOS ---
        dados_bombas_df = get_dados_bombas_df()
        
        # Obter todas as bombas ativas e em manutenção (geral ou por filial)
        query_bombas_ativas = supabase.table("bombas").select("serial, ativo, status, hospital, filial").eq("ativo", True)
        query_manut = supabase.table("manutencao").select("id, filial").eq("status", "Em Manutenção")

        if filial:
            query_bombas_ativas = query_bombas_ativas.eq("filial", filial)
            query_manut = query_manut.eq("filial", filial)

        ativas_list = query_bombas_ativas.execute().data
        manutencao_list = query_manut.execute().data
        em_manutencao = len(manutencao_list)

        # --- CÁLCULO DOS MODELOS (ULTA/ACTIVAC) ---
        ulta_count = 0
        activac_count = 0
        active_serials = [b['serial'] for b in ativas_list]
        if not dados_bombas_df.empty and active_serials:
            active_serials_normalized = [normalize_text(s) for s in active_serials]
            active_pumps_details = dados_bombas_df[dados_bombas_df['Serial_Normalized'].isin(active_serials_normalized)]
            model_counts = active_pumps_details['Modelo'].str.upper().value_counts().to_dict()
            ulta_count = model_counts.get('ULTA', 0)
            activac_count = model_counts.get('ACTIVAC', 0)

        # --- CORREÇÃO DA SOMA: O total "Em Comodato" no painel será a soma dos modelos exibidos ---
        total_comodato_painel = ulta_count + activac_count
        
        # --- CÁLCULO DE DISPONÍVEIS (SEMPRE GLOBAL) ---
        total_bombas_inventario = len(dados_bombas_df)
        total_ativas_geral = supabase.table("bombas").select("id", count='exact').eq("ativo", True).execute().count
        total_manut_geral = supabase.table("manutencao").select("id", count='exact').eq("status", "Em Manutenção").execute().count
        disponiveis = total_bombas_inventario - total_ativas_geral - total_manut_geral

        # --- Montagem final das métricas ---
        metrics = {
            "ativas": total_comodato_painel,
            "disponiveis": disponiveis,
            "em_manutencao": em_manutencao,
            "ulta_count": ulta_count,
            "activac_count": activac_count,
            "total_bombas": total_bombas_inventario,
            "status_counts": {"No Prazo": 0, "Menos de 7 dias": 0, "Fora Prazo": 0, "Indefinido": 0, "Data Inválida": 0},
            "hosp_counts": {},
            "bombas_por_filial": {"BRASILIA": 0, "GOIANIA": 0, "CUIABA": 0}
        }
        
        # Outros cálculos (status, hospital) usam a lista real de bombas ativas
        for bomba in ativas_list:
            status = bomba.get("status", "Indefinido")
            if status in metrics["status_counts"]:
                metrics["status_counts"][status] += 1
            
            hospital_original = bomba.get("hospital", "Desconhecido")
            hospital_normalizado = normalize_text(hospital_original)
            if hospital_normalizado:
                metrics["hosp_counts"][hospital_normalizado] = metrics["hosp_counts"].get(hospital_normalizado, 0) + 1
        
        # O total de bombas por filial para o mapa precisa ser da lista real
        bombas_geral_ativo = supabase.table("bombas").select("filial, ativo").eq("ativo", True).execute().data
        for bomba in bombas_geral_ativo:
            filial_norm = normalize_text(bomba.get("filial", ""))
            if filial_norm in metrics["bombas_por_filial"]:
                metrics["bombas_por_filial"][filial_norm] += 1
        
        logging.info("Métricas do dashboard carregadas com nova lógica.")
        return metrics
    except Exception as e:
        logging.error(f"Erro ao obter métricas do dashboard: {e}")
        return None

# -------------------- GERAÇÃO DE DOCUMENTOS --------------------
def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        # A biblioteca docx2pdf não é compatível com Linux sem MS Office.
        # Trocamos para 'docx-to-pdf', que usa LibreOffice em Linux.
        # É necessário instalar a biblioteca: pip install docx-to-pdf
        # E garantir que o LibreOffice esteja no ambiente: sudo apt-get install libreoffice
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        logging.info(f"Convertido '{docx_path}' -> '{pdf_path}'")
        return True
    except Exception as e:
        logging.error(f"Erro na conversão DOCX -> PDF: {e}")
        # A mensagem de erro é mais explícita sobre a real necessidade em Linux.
        st.warning(f"Erro ao converter para PDF. Em ambientes Linux, é necessário ter o LibreOffice instalado. Detalhe do erro: {e}")
        return False

def generate_combined_pdf(bomba_data):
    try:
        temp_docx_path = os.path.join(tempfile.gettempdir(), f"contrato_{bomba_data['serial']}.docx")
        
        if os.path.exists(CONTRATO_LOCAL_PATH):
            import shutil
            shutil.copyfile(CONTRATO_LOCAL_PATH, temp_docx_path)
        else:
            file_content = download_file_from_storage(CONTRATO_STORAGE_PATH)
            if not file_content:
                st.error("Modelo 'contrato.docx' não encontrado localmente ou no Supabase Storage!")
                return None
            with open(temp_docx_path, "wb") as f:
                f.write(file_content)
        
        doc = Document(temp_docx_path)
        data_atual = datetime.now()
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        data_formatada = f"Brasília, {data_atual.day:02d} de {meses[data_atual.month - 1]} de {data_atual.year}"
        
        replacements = {
            "{SERIAL}": bomba_data.get("serial", "N/A"),
            "{PACIENTE}": bomba_data.get("paciente", "N/A"),
            "{NOTA_FISCAL}": bomba_data.get("nf", "N/A"),
            "{DATA_ATUAL}": data_formatada,
        }

        for p in doc.paragraphs:
            for key, value in replacements.items():
                if key in p.text:
                    inline = p.runs
                    for i in range(len(inline)):
                        if key in inline[i].text:
                            inline[i].text = inline[i].text.replace(key, value)

        doc.save(temp_docx_path)
        temp_pdf_path = os.path.join(tempfile.gettempdir(), f"contrato_{bomba_data['serial']}.pdf")
        if not convert_docx_to_pdf(temp_docx_path, temp_pdf_path):
            os.remove(temp_docx_path)
            return None

        merger = PdfMerger()
        merger.append(temp_pdf_path)

        pdf_serial_path = os.path.join(tempfile.gettempdir(), f"{bomba_data['serial']}.pdf")
        file_content = download_file_from_storage(f"pdfs/{bomba_data['serial']}.pdf")
        if file_content:
            with open(pdf_serial_path, "wb") as f:
                f.write(file_content)
            merger.append(pdf_serial_path)
        else:
            logging.warning(f"Nenhum PDF associado ao serial {bomba_data['serial']} no Supabase Storage.")

        pdf_buffer = BytesIO()
        merger.write(pdf_buffer)
        merger.close()
        
        os.remove(temp_docx_path)
        os.remove(temp_pdf_path)
        if os.path.exists(pdf_serial_path):
            os.remove(pdf_serial_path)
        
        return pdf_buffer
    except Exception as e:
        logging.error(f"Erro ao gerar PDF Contrato: {e}")
        st.error("Erro ao gerar PDF.")
        return None

def generate_excel_saldo_curativo(df):
    """Gera um buffer de Excel a partir de um DataFrame do saldo de curativos."""
    if df.empty:
        return None
    
    output = BytesIO()
    # Usar 'with' garante que o writer seja fechado corretamente
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='SaldoCurativo')
    
    processed_data = output.getvalue()
    return processed_data

# -------------------- MÓDULO DE MANUTENÇÃO --------------------
def upload_nf_pdf(serial, data_registro, file):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        data_registro_clean = data_registro.replace("/", "")
        file_name = f"{MAINTENANCE_STORAGE_PATH}{serial}_{data_registro_clean}.pdf"

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(file.read())
            tmp_path = tmp.name
        
        with open(tmp_path, "rb") as f:
            supabase.storage.from_(bucket_name).upload(file_name, f, file_options={"content-type": "application/pdf"})

        os.remove(tmp_path)
        logging.info(f"NF assinada para {serial} enviada com sucesso.")
        return True
    except Exception as e:
        logging.error(f"Erro ao enviar NF assinada: {e}")
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
            try:
                dt_obj = datetime.strptime(bomba_data['data_registro'], '%d/%m/%Y')
                data_registro_str = dt_obj.strftime('%d-%m-%Y')
            except (ValueError, TypeError):
                data_registro_str = str(bomba_data['data_registro']).replace('/', '-')
        
        file_name = f"nfs_assinadas/{serial}*{hospital}_{paciente}*{data_registro_str}_assinado.pdf"
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(file.read())
            tmp_path = tmp.name
        
        with open(tmp_path, "rb") as f:
            supabase.storage.from_(bucket_name).upload(file_name, f, file_options={"content-type": "application/pdf", "cache-control": "3600", "upsert": "true"})
        
        os.remove(tmp_path)
        logging.info(f"NF assinada para {serial} enviada com sucesso como {file_name}.")
        return True
    except Exception as e:
        logging.error(f"Erro ao enviar NF assinada: {e}")
        st.error(f"Erro ao enviar NF assinada: {e}")
        return False

def get_nf_assinada_filename(serial):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        path = "nfs_assinadas/"
        files = supabase.storage.from_(bucket_name).list(path)
        
        search_prefix = f"{serial}*"
        search_suffix = "_assinado.pdf"
        
        for file_info in files:
            if file_info['name'].startswith(search_prefix.split('*')[0]) and file_info['name'].endswith(search_suffix):
                return f"{path}{file_info['name']}"
        return None
    except Exception as e:
        logging.error(f"Erro ao buscar nome do arquivo de NF assinada para {serial}: {e}")
        return None

def check_nf_assinada(serial):
    return get_nf_assinada_filename(serial) is not None

@st.cache_data(ttl=300)
def download_nf_assinada(serial):
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        file_path = get_nf_assinada_filename(serial)
        if file_path:
            response = supabase.storage.from_(bucket_name).download(file_path)
            return response
        return None
    except Exception as e:
        logging.error(f"Erro ao baixar NF assinada para {serial}: {e}")
        return None

def format_status(status):
    if status == "No Prazo":
        return '<span class="status-badge status-no-prazo">🟢 No Prazo</span>'
    elif status == "Menos de 7 dias":
        return '<span class="status-badge status-menos-7-dias">🟡 Menos de 7 dias</span>'
    elif status == "Fora Prazo":
        return '<span class="status-badge status-fora-prazo">🔴 Fora Prazo</span>'
    elif status == "Indefinido":
        return '<span class="status-badge status-indefinido">Indefinido</span>'
    elif status == "Data Inválida":
        return '<span class="status-badge status-data-invalida">Data Inválida</span>'
    elif status == "✅ DEVOLVIDA":
        return '<span class="status-badge status-devolvida">✅ Devolvida</span>'
    elif status == "Em Manutenção":
        return '<span class="status-badge status-em-manutencao">🛠 Em Manutenção</span>'
    return status

def display_manutencao_table(title, manutencoes, bombas_df):
    st.markdown(f"### {title}")
    if not manutencoes:
        st.info("Nenhum registro encontrado.")
        return

    display_cols = ["DATA REGISTRO", "SERIAL", "MODELO", "ÚLTIMA MANUT", "VENC. MANUT", "DEFEITO", "NF EMITIDA", "STATUS"]
    cols_db = ["data_registro", "serial", "modelo", "ultima_manut", "venc_manut", "defeito", "nf_numero", "status"]
    
    df = pd.DataFrame(manutencoes)
    for col in cols_db:
        if col not in df.columns:
            df[col] = "N/A"
    df = df[cols_db]
    df.columns = display_cols
    df['STATUS'] = df['STATUS'].apply(format_status)

    table_html = "<table style='width:100%; border-collapse: collapse;'><thead><tr>"
    for col in df.columns:
        table_html += f"<th style='border: 1px solid #ddd; padding: 8px;'>{col}</th>"
    table_html += "<th style='border: 1px solid #ddd; padding: 8px;'>AÇÕES</th>"
    table_html += "</tr></thead><tbody>"
    
    for _, row in df.iterrows():
        table_html += "<tr>"
        for col in df.columns:
            table_html += f"<td style='border: 1px solid #ddd; padding: 8px;'>{row[col]}</td>"
        
        table_html += "<td style='border: 1px solid #ddd; padding: 8px;'>"
        data_registro_clean = row['DATA REGISTRO'].replace("/", "")
        file_name = f"{row['SERIAL']}_{data_registro_clean}.pdf"
        table_html += f'<a href="#" onclick="alert(\'Download iniciado\');" download="{file_name}">📥 Download NF</a>'
        if row['STATUS'] == '<span class="status-badge status-em-manutencao">🛠 Em Manutenção</span>':
            table_html += f' <button onclick="alert(\'Marcar como Devolvida\');">✅ Marcar Devolvida</button>'
        table_html += "</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)

def display_bombas_table(title, bombas_list, bombas_df):
    st.markdown(f"### {title}")
    if not bombas_list:
        st.info("Nenhum registro encontrado.")
        return

    df = pd.DataFrame(bombas_list)
    df['serial_normalized'] = df['serial'].apply(normalize_text)
    
    df['modelo'] = 'N/A'
    df['ultima_manut'] = 'N/A'
    df['venc_manut'] = 'N/A'
    
    if not bombas_df.empty:
        merged_df = pd.merge(
            df,
            bombas_df,
            left_on='serial_normalized',
            right_on='Serial_Normalized',
            how='left'
        )
        
        df['modelo'] = merged_df['Modelo'].fillna('N/A')
        df['ultima_manut'] = merged_df['Ultima_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')
        df['venc_manut'] = merged_df['Venc_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')

    df['nf_assinada'] = df['serial'].apply(check_nf_assinada)

    df['status_html'] = df['status'].apply(format_status)
    df['periodo_str'] = df['periodo'].apply(lambda x: f"{x} dias" if pd.notna(x) and x != "N/A" else "N/A")
    df['nf_assinada_str'] = df['nf_assinada'].apply(lambda x: "✅ Sim" if x else "❌ Não")
    
    display_cols_map = {
        "serial": "SERIAL",
        "modelo": "MODELO",
        "hospital": "HOSPITAL",
        "paciente": "PACIENTE",
        "data_saida": "DATA SAÍDA",
        "periodo_str": "PERÍODO",
        "status_html": "STATUS",
        "nf": "NF",
        "ultima_manut": "ÚLTIMA MANUT",
        "venc_manut": "VENC MANUT",
        "nf_assinada_str": "NF ASSINADA"
    }
    
    table_html = "<table style='width:100%; border-collapse: collapse;'><thead><tr>"
    for header in display_cols_map.values():
        table_html += f"<th style='border: 1px solid #ddd; padding: 8px;'>{header}</th>"
    table_html += "</tr></thead><tbody>"

    for _, row in df.iterrows():
        table_html += "<tr>"
        for col_key in display_cols_map.keys():
            table_html += f"<td style='border: 1px solid #ddd; padding: 8px;'>{row.get(col_key, 'N/A')}</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)


def generate_excel_bombas_ativas(bombas, bombas_df, filial):
    if not bombas:
        return None

    cols = ["serial", "modelo", "hospital", "paciente", "data_saida", "periodo", "status", "nf", "ultima_manut", "venc_manut", "nf_assinada"]
    display_cols = ["SERIAL", "MODELO", "HOSPITAL", "PACIENTE", "DATA SAÍDA", "PERÍODO", "STATUS", "NF", "ÚLTIMA MANUT", "VENC MANUT", "NF ASSINADA"]
    
    df = pd.DataFrame(bombas)
    df['serial_normalized'] = df['serial'].apply(normalize_text)
    
    df['modelo'] = 'N/A'
    df['ultima_manut'] = 'N/A'
    df['venc_manut'] = 'N/A'
    df['nf_assinada'] = False

    if not bombas_df.empty:
        merged_df = pd.merge(
            df,
            bombas_df,
            left_on='serial_normalized',
            right_on='Serial_Normalized',
            how='left'
        )
        df['modelo'] = merged_df['Modelo'].fillna('N/A')
        df['ultima_manut'] = merged_df['Ultima_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')
        df['venc_manut'] = merged_df['Venc_Manut'].dt.strftime('%d/%m/%Y').fillna('N/A')
        df['nf_assinada'] = df['serial'].apply(check_nf_assinada)
    
    df = df[cols]
    df.columns = display_cols
    df["PERÍODO"] = df["PERÍODO"].apply(lambda x: f"{x} dias" if pd.notna(x) else "N/A")
    df["NF ASSINADA"] = df["NF ASSINADA"].apply(lambda x: "Sim" if x else "Não")
    
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer

# -------------------- INTERFACE --------------------
def main():
    st.markdown('<h1 class="main-title">Controle de Bombas de Sucção</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Desenvolvido por Guilherme Rodrigues – Suplen Médical</p>', unsafe_allow_html=True)
    
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

        st.markdown("---")
        
        col_graph1, col_graph2 = st.columns(2)
        with col_graph1:
            st.markdown("#### Status das Bombas em Comodato")
            df_status = pd.DataFrame(list(metrics["status_counts"].items()), columns=['Status', 'Quantidade'])
            df_status = df_status[df_status['Quantidade'] > 0]
            df_status['Status'] = df_status['Status'].apply(
                lambda x: f"🟢 {x}" if x == "No Prazo" else f"🟡 {x}" if x == "Menos de 7 dias" else f"🔴 {x}" if x == "Fora Prazo" else x
            )
            if not df_status.empty:
                fig_pie = px.pie(
                    df_status, names='Status', values='Quantidade',
                    title="Distribuição por Status", template="plotly_dark",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sem bombas em comodato para exibir.")
        
        with col_graph2:
            st.markdown("#### Bombas em Comodato por Hospital")
            df_hosp = pd.DataFrame(list(metrics["hosp_counts"].items()), columns=['hospital', 'count'])
            if not df_hosp.empty:
                df_hosp['hospital'] = df_hosp['hospital'].apply(lambda x: x[:30] + '...' if len(x) > 30 else x)
                fig_bar = px.bar(
                    df_hosp, x='hospital', y='count',
                    title="Top Hospitais com Bombas em Comodato", color='hospital',
                    text='count', template='plotly_dark',
                    labels={'hospital': 'Hospital', 'count': 'Nº de Bombas'},
                    color_discrete_sequence=px.colors.qualitative.Dark24
                )
                fig_bar.update_layout(
                    showlegend=False, xaxis_tickangle=45, margin=dict(t=50, b=100, l=200),
                    xaxis=dict(tickfont=dict(size=10), automargin=True)
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Sem bombas em comodato para exibir.")

    elif choice == "Dashboard Geral":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Dashboard Geral (Todas as Filiais)</h2>', unsafe_allow_html=True)
        filial_filter = st.selectbox("Filtrar por Filial", ["Todas"] + FILIAIS, key="filial_filter")
        filial_to_query = filial_filter if filial_filter != "Todas" else None

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
        st.markdown("---")

        st.markdown("### Análise das Bombas em Comodato")
        col_graph1, col_graph2 = st.columns(2)
        with col_graph1:
            st.markdown("##### Distribuição por Status")
            df_status_global = pd.DataFrame(list(metrics["status_counts"].items()), columns=['Status', 'Quantidade'])
            df_status_global = df_status_global[df_status_global['Quantidade'] > 0]
            if not df_status_global.empty:
                df_status_global['Status'] = df_status_global['Status'].apply(lambda x: f"🟢 {x}" if x == "No Prazo" else f"🟡 {x}" if x == "Menos de 7 dias" else f"🔴 {x}" if x == "Fora Prazo" else x)
                fig_pie_global = px.pie(
                    df_status_global, names='Status', values='Quantidade',
                    title="Distribuição por Status (Global)", template="plotly_dark",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig_pie_global.update_traces(textposition='inside', textinfo='percent+label', textfont_size=14)
                st.plotly_chart(fig_pie_global, use_container_width=True)
            else:
                st.info("Nenhuma bomba em comodato para exibir status.")
                
        with col_graph2:
            st.markdown("##### Distribuição por Hospital")
            df_hosp_global = pd.DataFrame(list(metrics["hosp_counts"].items()), columns=['hospital', 'count'])
            if not df_hosp_global.empty:
                df_hosp_global['hospital'] = df_hosp_global['hospital'].apply(lambda x: x[:30] + '...' if len(x) > 30 else x)
                fig_bar_global = px.bar(
                    df_hosp_global, x='hospital', y='count',
                    title="Top Hospitais (Global)", color='hospital', text='count',
                    template='plotly_dark', labels={'hospital': 'Hospital', 'count': 'Nº de Bombas'},
                    color_discrete_sequence=px.colors.qualitative.Dark24
                )
                fig_bar_global.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(t=50, b=100, l=200), xaxis=dict(tickfont=dict(size=10), automargin=True))
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

            active_filials = []
            for filial_map, (lat, lon) in coordinates.items():
                normalized_filial = normalize_text(filial_map)
                count = metrics["bombas_por_filial"].get(normalized_filial, 0)
                if count > 0:
                    radius = max(5000, count * 10000)
                    folium.Circle(location=[lat, lon], radius=radius, fill=True, fill_opacity=0.2, color=colors[filial_map], fill_color=colors[filial_map], popup=folium.Popup(f"{filial_map}: {count} bombas em comodato", max_width=200)).add_to(m)
                    active_filials.append((filial_map, count))
            st_folium(m, width=800, height=400)
        with col2:
            # --- INÍCIO DA CORREÇÃO ---
            # Envolvemos a legenda em um único bloco de markdown com um contêiner flexível.
            # O spacer div com 'flex-grow: 1' empurrará todo o conteúdo para cima.
            legend_html = """
            <div style="height: 100%; display: flex; flex-direction: column;">
                <h3>Legenda</h3>
            """
            items_added = False
            for filial_map, color in colors.items():
                normalized_filial = normalize_text(filial_map)
                display_count = metrics["bombas_por_filial"].get(normalized_filial, 0)
                if display_count > 0:
                    items_added = True
                    legend_html += (f'<div class="legend-item"><span class="legend-dot" style="background-color: {color};"></span>{filial_map}: {display_count} bombas</div>')
            
            if not items_added:
                legend_html += "<div>Nenhuma bomba em comodato registrada.</div>"
            
            legend_html += "<div style='flex-grow: 1;'></div></div>" # Este div empurra o conteúdo para cima
            
            st.markdown(legend_html, unsafe_allow_html=True)
            # --- FIM DA CORREÇÃO ---
            
        st.markdown('<h2 class="section-title">Análise de Curativos</h2>', unsafe_allow_html=True)
        with st.spinner("Carregando análise de curativos..."):
            curativo_results = analyze_curativo()
            
            if curativo_results.get('last_updated'):
                try:
                    ts_utc_str = curativo_results['last_updated'].replace("Z", "+00:00")
                    ts_utc = datetime.fromisoformat(ts_utc_str)
                    ts_local = ts_utc.astimezone(timezone(timedelta(hours=-3)))
                    formatted_ts = ts_local.strftime('%d/%m/%Y às %H:%M')
                    st.caption(f"🗓️ _Dados da planilha atualizados em: **{formatted_ts}**_")
                except Exception as e:
                    logging.error(f"Erro ao formatar o timestamp da planilha: {e}")

            if curativo_results.get('error'):
                st.error(curativo_results['error'])
            else:
                # Filtra status indesejados dos DataFrames
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
                        # Criar uma coluna para o texto formatado para não alterar o tipo de dado da receita
                        revenue_status_df['Revenue_Formatted'] = revenue_status_df['Revenue'].apply(lambda x: f"R$ {x:,.2f}")
                        fig_revenue = px.bar(revenue_status_df, x='Status', y='Revenue', title="Receita por Status", text='Revenue_Formatted', template='plotly_dark', color='Status', color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_revenue.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(t=50, b=100, l=200), xaxis=dict(tickfont=dict(size=10), automargin=True), yaxis_title="Receita (R$)")
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
                elif not periodo.isdigit():
                    st.error("ERRO: O campo 'PERÍODO' deve ser um número inteiro de dias.")
                else:
                    try:
                        with st.spinner("Verificando serial..."):
                            existing = supabase.table("bombas").select("id").eq("serial", serial).eq("ativo", True).execute()
                        if existing.data:
                            st.error(f"ERRO: A bomba com o serial '{serial}' já está registrada e ativa.")
                        else:
                            data_saida_fmt = data_saida.strftime("%Y-%m-%d")
                            data_registro = datetime.now().strftime("%Y-%m-%d")
                            status_inicial = calculate_status(data_saida_fmt, periodo)
                            with st.spinner("Registrando bomba..."):
                                response = supabase.table("bombas").insert({"serial": serial, "hospital": hospital, "paciente": paciente, "medico": medico, "convenio": convenio, "data_registro": data_registro, "data_saida": data_saida_fmt, "periodo": int(periodo), "status": status_inicial, "nf": nf, "pedido": pedido, "ativo": True, "filial": filial, "nf_devolucao": ""}).execute()
                                if response.data:
                                    bomba_id = response.data[0]["id"]
                                    register_event("bombas", bomba_id, f"BOMBA REGISTRADA (SERIAL: {serial})", filial)
                                    flush_events()
                                    st.session_state.messages.append({"text": "Bomba registrada com sucesso!", "icon": "✅"})
                                    st.cache_data.clear()
                                    st.rerun()
                                else:
                                    st.error("Ocorreu um erro desconhecido ao registrar a bomba.")
                    except Exception as e:
                        logging.error(f"Erro no registro da bomba: {e}")
                        st.error(f"ERRO AO REGISTRAR: {e}")

    elif choice == "Bombas em Comodato":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Bombas em Comodato</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar bombas em comodato (por Serial, Paciente ou Hospital)...", key="listagem_search")
        with st.spinner("Carregando bombas..."):
            bombas_ativas = get_bombas(search_term, filial, active_only=True)
            bombas_df = get_dados_bombas_df()
        display_bombas_table("Listagem de Bombas Ativas", bombas_ativas, bombas_df)
        st.markdown("---")
        excel_buffer = generate_excel_bombas_ativas(bombas_ativas, bombas_df, filial)
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
                st.caption("Use esta seção para corrigir informações de uma bomba que já foi registrada.")
                bomba_to_edit = st.selectbox("Selecione uma bomba para editar os dados:", options=bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="edit_bomba_select", index=None, placeholder="Selecione uma bomba...")
                if bomba_to_edit:
                    with st.form(key=f"edit_form_{bomba_to_edit['id']}"):
                        try:
                            current_data_saida = datetime.strptime(bomba_to_edit.get('data_saida', ''), '%d/%m/%Y').date()
                        except (ValueError, TypeError):
                            current_data_saida = None
                        c1, c2 = st.columns(2)
                        with c1:
                            serial = st.text_input("Serial*", value=bomba_to_edit.get('serial', '')).upper()
                            paciente = st.text_input("Paciente", value=bomba_to_edit.get('paciente', '')).upper()
                            hospital = st.text_input("Hospital", value=bomba_to_edit.get('hospital', '')).upper()
                            medico = st.text_input("Médico", value=bomba_to_edit.get('medico', '')).upper()
                            data_saida = st.date_input("Data de Saída*", value=current_data_saida, format="DD/MM/YYYY")
                        with c2:
                            convenio = st.text_input("Convênio", value=bomba_to_edit.get('convenio', '')).upper()
                            periodo = st.text_input("Período (dias)*", value=bomba_to_edit.get('periodo', ''))
                            pedido = st.text_input("Pedido", value=bomba_to_edit.get('pedido', '')).upper()
                            nf = st.text_input("NF", value=bomba_to_edit.get('nf', '')).upper()
                        submitted = st.form_submit_button("Salvar Alterações")
                        if submitted:
                            if not all([serial, data_saida, periodo]):
                                st.error("ERRO: Os campos Serial, Data de Saída e Período são obrigatórios.")
                            elif not str(periodo).isdigit():
                                st.error("ERRO: O campo 'Período' deve ser um número inteiro.")
                            else:
                                with st.spinner("Atualizando dados..."):
                                    try:
                                        data_saida_fmt = data_saida.strftime("%Y-%m-%d")
                                        new_status = calculate_status(data_saida_fmt, periodo)
                                        update_data = {"serial": serial, "paciente": paciente, "hospital": hospital, "medico": medico, "convenio": convenio, "periodo": int(periodo), "pedido": pedido, "nf": nf, "data_saida": data_saida_fmt, "status": new_status}
                                        supabase.table("bombas").update(update_data).eq("id", bomba_to_edit['id']).execute()
                                        register_event("bombas", bomba_to_edit['id'], f"DADOS DA BOMBA ATUALIZADOS (SERIAL: {serial})", filial)
                                        flush_events()
                                        st.session_state.messages.append({"text": "Dados da bomba atualizados com sucesso!", "icon": "📝"})
                                        st.cache_data.clear()
                                        st.rerun()
                                    except Exception as e:
                                        logging.error(f"Erro ao atualizar bomba: {e}")
                                        st.error(f"Ocorreu um erro ao salvar as alterações: {e}")
            with tab_pdf:
                st.subheader("Gerar PDF do Contrato")
                st.caption("Use esta seção para gerar o contrato de comodato e a ficha técnica da bomba em um único PDF.")
                bomba_pdf = st.selectbox("Selecione uma bomba para gerar o contrato:", bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="pdf_bomba_select", index=None, placeholder="Selecione uma bomba...")
                if bomba_pdf:
                    if st.button("Gerar PDF do Contrato", disabled=st.session_state.get("generating_pdf", False), key="generate_pdf_button"):
                        st.session_state.generating_pdf = True
                        try:
                            with st.spinner("Gerando PDF... Esta operação pode levar alguns segundos."):
                                pdf_bytes = generate_combined_pdf(bomba_pdf)
                                if pdf_bytes:
                                    st.download_button(label="📥 Baixar PDF do Contrato", data=pdf_bytes, file_name=f"documentos_{bomba_pdf['serial']}.pdf", mime="application/pdf", key=f"download_pdf_{bomba_pdf['id']}")
                                    register_event("bombas", bomba_pdf['id'], "DOCUMENTOS GERADOS", filial)
                                    flush_events()
                                else:
                                    st.error("Falha ao gerar o PDF. Verifique se o arquivo 'contrato.docx' existe e se o PDF da bomba está no sistema.")
                        finally:
                            st.session_state.generating_pdf = False
            with tab_anexar_nf:
                st.subheader("Anexar NF Assinada")
                st.caption("Anexe o PDF da Nota Fiscal de comodato assinada pelo cliente.")
                bombas_sem_nf = [bomba for bomba in bombas_ativas if not check_nf_assinada(bomba['serial'])]
                if not bombas_sem_nf:
                    st.info("Todas as bombas em comodato nesta listagem já possuem NF assinada.")
                else:
                    bomba_anexar = st.selectbox("Selecione a bomba para anexar a NF assinada:", bombas_sem_nf, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="anexar_nf_select")
                    if bomba_anexar:
                        with st.form(key=f"anexar_nf_form_{bomba_anexar['id']}", clear_on_submit=True):
                            nf_file = st.file_uploader("📄 ANEXAR NF ASSINADA (PDF)*", type=["pdf"], key=f"nf_assinada_upload_{bomba_anexar['id']}")
                            submitted = st.form_submit_button("Enviar NF Assinada")
                            if submitted:
                                if not nf_file:
                                    st.error("Por favor, faça o upload do arquivo PDF da NF assinada.")
                                else:
                                    try:
                                        with st.spinner("Enviando NF..."):
                                            if upload_nf_assinada(bomba_anexar, nf_file):
                                                register_event("bombas", bomba_anexar["id"], f"NF ASSINADA ENVIADA PARA BOMBA (SERIAL: {bomba_anexar['serial']})", filial)
                                                flush_events()
                                                st.session_state.messages.append({"text": "NF assinada enviada com sucesso!", "icon": "📎"})
                                                st.cache_data.clear()
                                                st.rerun()
                                            else:
                                                st.error("Erro ao enviar NF assinada.")
                                    except Exception as e:
                                        st.error(f"Erro: {e}")
            with tab_download_nf:
                st.subheader("Download de NF Assinada")
                st.caption("Baixe a Nota Fiscal de comodato assinada que foi anexada previamente.")
                bombas_com_nf = [bomba for bomba in bombas_ativas if check_nf_assinada(bomba['serial'])]
                if not bombas_com_nf:
                    st.info("Nenhuma bomba nesta listagem possui NF assinada para download.")
                else:
                    bomba_selecionada = st.selectbox("Selecione a bomba para baixar a NF assinada:", bombas_com_nf, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} | PACIENTE: {b.get('paciente', 'N/A')} | NF: {b.get('nf', 'N/A')}", key="download_nf_select_tab")
                    if st.button("Baixar NF Assinada", key="baixar_nf_button_tab"):
                        try:
                            with st.spinner("Preparando download..."):
                                nf_data = download_nf_assinada(bomba_selecionada['serial'])
                                nf_filename_path = get_nf_assinada_filename(bomba_selecionada['serial'])
                                if nf_data and nf_filename_path:
                                    nf_filename = os.path.basename(nf_filename_path)
                                    st.download_button(label="📥 Clique aqui para baixar a NF Assinada", data=nf_data, file_name=nf_filename, mime="application/pdf")
                                    register_event("bombas", bomba_selecionada["id"], f"NF ASSINADA BAIXADA (SERIAL: {bomba_selecionada['serial']})", filial)
                                    flush_events()
                                    st.session_state.messages.append({"text": "Download da NF assinada iniciado!", "icon": "📥"})
                                else:
                                    st.error("Erro ao baixar NF assinada.")
                        except Exception as e:
                            st.error(f"Erro: {e}")

    elif choice == "Devolver":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Devolver Bomba</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar bomba consignada...", key="devolver_search")
        bombas_ativas = get_bombas(search_term, filial, active_only=True)
        if not bombas_ativas:
            st.info("Nenhuma bomba consignada encontrada para devolução.")
        else:
            with st.form("devolver_form", clear_on_submit=True):
                bomba = st.selectbox("Selecione a bomba a devolver:", bombas_ativas, format_func=lambda b: f"SERIAL: {b.get('serial', 'N/A')} - PACIENTE: {b.get('paciente', 'N/A')} - HOSPITAL: {b.get('hospital', 'N/A')}")
                data_retorno = st.date_input("📅 Data de Retorno", value=datetime.now(), format="DD/MM/YYYY")
                nf_devolucao = st.text_input("🧾 NF de Devolução*", placeholder="Ex.: 987654").upper()
                submitted = st.form_submit_button("Confirmar Devolução")
                
                if submitted:
                    if not nf_devolucao:
                        st.error("Preencha o campo NF de Devolução.")
                    elif not bomba:
                        st.error("Selecione uma bomba para devolver.")
                    else:
                        try:
                            with st.spinner("Registrando devolução..."):
                                data_retorno_fmt = data_retorno.strftime("%Y-%m-%d")
                                supabase.table("bombas").update({"ativo": False, "status": "✅ DEVOLVIDA", "data_retorno": data_retorno_fmt, "nf_devolucao": nf_devolucao}).eq("id", bomba["id"]).execute()
                                register_event("bombas", bomba["id"], f"BOMBA DEVOLVIDA (SERIAL: {bomba['serial']}) EM {data_retorno.strftime('%d/%m/%Y')} (NF DEVOLUÇÃO: {nf_devolucao})", filial)
                                flush_events()
                                st.session_state.messages.append({"text": "Bomba devolvida com sucesso!", "icon": "✅"})
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
                st.markdown('<div class="form-container">', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown('<div class="form-column">', unsafe_allow_html=True)
                    serial = st.text_input("🔢 SERIAL*", placeholder="Ex.: ABC123").upper()
                    defeito = st.text_area("📝 DEFEITO APRESENTADO*", placeholder="Descreva o defeito apresentado").upper()
                    nf_file = st.file_uploader("📄 UPLOAD NF (PDF)*", type=["pdf"], key="nf_upload")
                    st.markdown('</div>', unsafe_allow_html=True)
                with c2:
                    st.markdown('<div class="form-column">', unsafe_allow_html=True)
                    data_registro = datetime.now().strftime("%d/%m/%Y")
                    st.text_input("📅 DATA DE REGISTRO", value=data_registro, disabled=True)
                    nf_numero = st.text_input("🧾 NÚMERO DA NF EMITIDA*", placeholder="Ex.: 123456").upper()
                    st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                submit_button = st.form_submit_button("Registrar Manutenção")
                if submit_button:
                    if not all([serial, defeito, nf_numero, nf_file]):
                        st.error("Preencha todos os campos obrigatórios (*), incluindo o upload da NF.")
                    else:
                        try:
                            with st.spinner("Verificando e registrando..."):
                                existing = supabase.table("manutencao").select("id").eq("serial", serial).eq("status", "Em Manutenção").execute()
                                if existing.data:
                                    st.error(f"ERRO: Serial '{serial}' já está em manutenção!")
                                else:
                                    if upload_nf_pdf(serial, data_registro, nf_file):
                                        data_registro_fmt = datetime.now().strftime("%Y-%m-%d")
                                        response = supabase.table("manutencao").insert({"serial": serial, "defeito": defeito, "data_registro": data_registro_fmt, "nf_numero": nf_numero, "nf_status": "Enviada", "status": "Em Manutenção", "filial": filial}).execute()
                                        manutencao_id = response.data[0]["id"]
                                        register_event("manutencao", manutencao_id, f"MANUTENÇÃO REGISTRADA (SERIAL: {serial}, NF: {nf_numero})", filial)
                                        flush_events()
                                        st.session_state.messages.append({"text": "Manutenção registrada com sucesso!", "icon": "🛠️"})
                                        st.cache_data.clear()
                                        st.rerun()
                                    else:
                                        st.error("Erro ao enviar a NF.")
                        except Exception as e:
                            st.error(f"Erro: {e}")
        with tabs[1]:
            st.markdown("### Listagem de Bombas em Manutenção")
            search_term = st.text_input("Pesquisar manutenções...", key="manutencao_search")
            manutencoes = get_manutencao(search_term, filial)
            bombas_df = get_dados_bombas_df()
            display_manutencao_table("Manutenções Registradas", manutencoes, bombas_df)
            for manut in manutencoes:
                if manut['status'] == "Em Manutenção":
                    if st.button(f"Marcar como Devolvida: {manut['serial']}", key=f"devolver_manut_{manut['id']}"):
                        try:
                            supabase.table("manutencao").update({"status": "Devolvida"}).eq("id", manut["id"]).execute()
                            register_event("manutencao", manut["id"], f"BOMBA DEVOLVIDA APÓS MANUTENÇÃO (SERIAL: {manut['serial']})", filial)
                            flush_events()
                            st.session_state.messages.append({"text": f"Bomba {manut['serial']} marcada como devolvida!", "icon": "✅"})
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro: {e}")

    elif choice == "Histórico Devolvidas":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Histórico de Bombas Devolvidas</h2>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar histórico...", key="historico_search")
        historico = get_historico_devolvidas(filial if not st.session_state.get("general_mode", False) else None)
        if not historico:
            st.info("Nenhuma bomba devolvida registrada.")
        else:
            df = pd.DataFrame(historico, columns=["data_evento", "descricao", "filial"])
            df.columns = ["Data do Evento", "Descrição", "Filial"]
            if search_term:
                df = df[df.apply(lambda row: search_term.lower() in str(row).lower(), axis=1)]
            st.dataframe(df, use_container_width=True, hide_index=True)
            
    elif choice == "Saldo Curativo":
        st.markdown('<h2 style="font-size: 1.25rem; margin-bottom: 1rem;">Saldo de Curativos</h2>', unsafe_allow_html=True)
        
        search_term = st.text_input("Pesquisar por Descrição do Produto...", key="saldo_curativo_search")

        # Legenda de Cores
        st.markdown("---")
        st.markdown("<h6>Legenda de Validade</h6>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown('<div style="background-color: #90ee90; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>NORMAL</strong><br>> 3 meses</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div style="background-color: #FFD580; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>ATENÇÃO</strong><br>2 a 3 meses</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div style="background-color: #F08080; color: black; text-align: center; padding: 10px; border-radius: 5px;"><strong>CRÍTICO</strong><br>< 2 meses</div>', unsafe_allow_html=True)
        st.markdown("---",)
        
        df_curativos = get_saldo_curativo_data()

        if df_curativos.empty:
            st.info("Nenhum dado de saldo de curativo encontrado.")
        else:
            df_para_exibir = df_curativos.copy()
            
            # Lógica de busca inteligente
            if search_term:
                search_keywords = normalize_text(search_term).split()
                
                def fuzzy_search(row):
                    if not isinstance(row.get('Desc_Produto'), str):
                        return False
                    
                    desc_words = normalize_text(row['Desc_Produto']).split()
                    if not desc_words:
                        return False
                    
                    for skey in search_keywords:
                        match_found = any(skey.startswith(dkey) or dkey.startswith(skey) for dkey in desc_words)
                        if not match_found:
                            return False
                    return True

                mask = df_para_exibir.apply(fuzzy_search, axis=1)
                df_para_exibir = df_para_exibir[mask]

            # Continuar com o dataframe filtrado
            if not df_para_exibir.empty:
                df_para_exibir['Data_Validad'] = pd.to_datetime(df_para_exibir['Data_Validad'], errors='coerce')
                df_para_exibir.dropna(subset=['Data_Validad'], inplace=True)
                df_para_exibir = df_para_exibir.sort_values(by=['Data_Validad', 'Desc_Produto'], ascending=[True, True])

                df_display = df_para_exibir[['Produto', 'Desc_Produto', 'Referencia', 'Lote', 'Data_Validad', 'Saldo_Lote']].copy()
                df_display.rename(columns={
                    'Desc_Produto': 'Descrição do Produto',
                    'Referencia': 'Referência',
                    'Data_Validad': 'Data de Validade',
                    'Saldo_Lote': 'Saldo do Lote'
                }, inplace=True)

                def style_validade(row):
                    hoje = datetime.now()
                    data_validade = row['Data de Validade']
                    diferenca_dias = (data_validade - hoje).days
                    text_color = "color: black;"
                    
                    if diferenca_dias > 90:
                        bg_color = 'background-color: #90ee90;'
                    elif 60 < diferenca_dias <= 90:
                        bg_color = 'background-color: #FFD580;'
                    else:
                        bg_color = 'background-color: #F08080;'
                    
                    style = bg_color + " " + text_color
                    return [style] * len(row)

                styler = df_display.style.apply(style_validade, axis=1)
                styler.format({'Data de Validade': '{:%d/%m/%Y}'})
                styler.set_properties(subset=['Saldo do Lote'], **{'text-align': 'center'})

                st.dataframe(styler, use_container_width=True, hide_index=True)

                # Usar o styler.data para obter o DataFrame com os dados corretos antes da formatação de exibição
                df_for_download = styler.data
                excel_data = generate_excel_saldo_curativo(df_for_download)
                if excel_data:
                    st.download_button(
                        label="✅ Baixar Saldo em Excel",
                        data=excel_data,
                        file_name=f"saldo_curativo_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Nenhum produto encontrado com os termos da busca.")

if __name__ == "__main__":
    main()