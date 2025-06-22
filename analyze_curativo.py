import pandas as pd
import logging
import os
import sys
import tempfile
from datetime import datetime, timezone, timedelta
from supabase import create_client, Client
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Inicializar Supabase
def init_supabase():
    """Inicializa o cliente Supabase e retorna o cliente de armazenamento."""
    try:
        supabase_url = os.getenv("SUPABASE_URL")
        supabase_key = os.getenv("SUPABASE_KEY")
        if not supabase_url or not supabase_key:
            logging.error("Variáveis SUPABASE_URL ou SUPABASE_KEY não encontradas.")
            return None

        supabase: Client = create_client(supabase_url, supabase_key)
        logging.info("Supabase inicializado com sucesso.")
        return supabase
    except Exception as e:
        logging.error(f"Erro ao inicializar Supabase: {e}")
        return None

def download_file_from_storage(supabase, storage_path, local_path):
    """Baixa um arquivo do Supabase Storage para um caminho local."""
    try:
        bucket_name = "controle-de-bombas-suplen-files"
        response = supabase.storage.from_(bucket_name).download(storage_path)
        with open(local_path, "wb") as f:
            f.write(response)
        logging.info(f"Arquivo {storage_path} baixado para {local_path}")
        return True
    except Exception as e:
        logging.error(f"Erro ao baixar {storage_path}: {e}")
        return False

def analyze_curativo(file_path='analise/bdcurativo.xlsx'):
    """
    Analisa a planilha bdcurativo.xlsx do Supabase Storage e retorna KPIs para uso em dashboards.
    Args:
        file_path (str): Caminho do arquivo no Supabase Storage.
    Returns:
        dict: Dicionário com DataFrames e KPIs.
    """
    try:
        # Inicializar Supabase
        supabase = init_supabase()
        if not supabase:
            logging.error("Falha ao inicializar Supabase Storage.")
            return {"error": "Erro ao conectar ao Supabase Storage. Verifique as credenciais."}

        # **NOVA LÓGICA: Buscar metadados do arquivo para obter data de atualização**
        last_updated_timestamp = None
        bucket_name = "controle-de-bombas-suplen-files"
        try:
            file_list = supabase.storage.from_(bucket_name).list(path='analise')
            for file_obj in file_list:
                if file_obj['name'] == 'bdcurativo.xlsx':
                    last_updated_timestamp = file_obj['updated_at']
                    logging.info(f"Timestamp do arquivo encontrado: {last_updated_timestamp}")
                    break
        except Exception as e:
            logging.warning(f"Não foi possível obter os metadados do arquivo: {e}")


        # Criar arquivo temporário para armazenar o download
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            temp_path = temp_file.name

        # Baixar o arquivo do Storage
        if not download_file_from_storage(supabase, file_path, temp_path):
            logging.error(f"Arquivo {file_path} não encontrado no Supabase Storage.")
            return {"error": f"Arquivo {file_path} não encontrado no Supabase Storage."}

        # Ler a planilha
        df = pd.read_excel(temp_path)

        # Remover arquivo temporário
        try:
            os.remove(temp_path)
            logging.info(f"Arquivo temporário {temp_path} removido.")
        except Exception as e:
            logging.warning(f"Erro ao remover arquivo temporário {temp_path}: {e}")

        # Limpeza de dados
        df['Valor Cotado'] = df['Valor Cotado'].replace('- 0', 0).replace('', 0)
        df['Valor Cotado'] = pd.to_numeric(df['Valor Cotado'], errors='coerce').fillna(0)
        date_columns = ['Dt Procedime', 'Dt Apont Uti']
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

        # KPI 1: Porcentagem de vendas por status
        status_counts = df['Status Utili'].value_counts()
        status_percentages = (status_counts / status_counts.sum() * 100).round(2)
        status_df = pd.DataFrame({
            'Status': status_percentages.index,
            'Percentage': status_percentages.values
        })

        # KPI 2: Top 5 produtos mais vendidos
        sales_df = df[df['Status Utili'].isin(['Utilizado', 'Finalizado'])]
        product_counts = sales_df['Desc Produto'].value_counts().head(5)
        product_df = pd.DataFrame({
            'Product': product_counts.index, 
            'Count': product_counts.values
        })

        # KPI 3: Receita por status
        revenue_by_status = df.groupby('Status Utili')['Valor Cotado'].sum().round(2)
        revenue_status_df = pd.DataFrame({
            'Status': revenue_by_status.index,
            'Revenue': revenue_by_status.values
        })

        # KPI 4: Receita por produto (Finalizado)
        finalized_df = df[df['Status Utili'] == 'Finalizado']
        revenue_by_product = finalized_df.groupby('Desc Produto')['Valor Cotado'].sum().nlargest(5).round(2)
        revenue_product_df = pd.DataFrame({
            'Product': revenue_by_product.index,
            'Revenue': revenue_by_product.values
        })

        # KPI 5: Desempenho por cliente
        client_sales = sales_df.groupby('Nome Cli').agg({
            'Desc Produto': 'count',
            'Valor Cotado': 'sum'
        }).rename(columns={'Desc Produto': 'Sales Count', 'Valor Cotado': 'Revenue'}).nlargest(5, 'Sales Count')
        client_sales['Revenue'] = client_sales['Revenue'].round(2)

        # KPI 6: Taxa de perda de estoque
        non_found_count = df[df['Status Utili'] == 'Não Encontrado'].shape[0]
        total_count = df.shape[0]
        loss_rate = round(non_found_count / total_count * 100, 2) if total_count > 0 else 0

        # KPI 7: Tempo médio de faturamento
        finalized_df = finalized_df.copy()
        finalized_df['Days to Invoice'] = (finalized_df['Dt Apont Uti'] - finalized_df['Dt Procedime']).dt.days
        avg_days_to_invoice = finalized_df['Days to Invoice'].mean().round(2) if not finalized_df.empty else 0

        # KPI 8: Vendas por mês
        sales_df = sales_df.copy()
        sales_df['Month'] = sales_df['Dt Procedime'].dt.to_period('M')
        sales_by_month = sales_df.groupby('Month').agg({
            'Desc Produto': 'count',
            'Valor Cotado': 'sum'
        }).rename(columns={'Desc Produto': 'Sales Count', 'Valor Cotado': 'Revenue'})
        sales_by_month['Revenue'] = sales_by_month['Revenue'].round(2)

        logging.info("Análise da planilha concluída com sucesso.")
        return {
            'status_df': status_df,
            'product_df': product_df,
            'revenue_status_df': revenue_status_df,
            'revenue_product_df': revenue_product_df,
            'client_sales': client_sales,
            'loss_rate': loss_rate,
            'avg_days_to_invoice': avg_days_to_invoice,
            'sales_by_month': sales_by_month,
            'error': None,
            'last_updated': last_updated_timestamp  # **NOVA INFORMAÇÃO RETORNADA**
        }

    except Exception as e:
        logging.error(f"Erro na análise da planilha: {e}")
        return {"error": f"Erro na análise: {str(e)}"}

if __name__ == "__main__":
    results = analyze_curativo()
    if not results.get('error'):
        print("KPI 1: Porcentagem de Vendas por Status:")
        print(results['status_df'].to_string(index=False))
        print("\nKPI 2: Top 5 Produtos Mais Vendidos:")
        print(results['product_df'].to_string(index=False))
        print("\nKPI 3: Receita por Status:")
        print(results['revenue_status_df'].to_string(index=False))
        print("\nKPI 4: Receita por Produto:")
        print(results['revenue_product_df'].to_string(index=False))
        print("\nKPI 5: Desempenho por Cliente:")
        print(results['client_sales'].to_string())
        print(f"\nKPI 6: Taxa de Perda de Estoque: {results['loss_rate']}%")
        print(f"\nKPI 7: Tempo Médio de Faturamento: {results['avg_days_to_invoice']} dias")
        print("\nKPI 8: Vendas por Mês:")
        print(results['sales_by_month'].to_string())
    else:
        print(results['error'])