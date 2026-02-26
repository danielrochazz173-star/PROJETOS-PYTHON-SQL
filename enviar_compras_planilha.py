"""
Sistema de Envio de Compras para Planilha
Envia dados de compras para planilha Excel ou Google Sheets
"""

import oracledb
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# Importa gspread para Google Sheets
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False
    print("⚠️  Aviso: gspread não está instalado. Funcionalidade Google Sheets desabilitada.")

# Caminho do arquivo de credenciais (busca no diretório raiz, igual ao enviar_estoque_bembrasil.py)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)
GOOGLE_CREDENTIALS_FILE = os.path.join(ROOT_DIR, 'tensile-proxy-468412-t2-dd576fcb8c22.json')

# Função para conectar ao Google Sheets
def _get_worksheet(spreadsheet_name, worksheet_name):
    """Função auxiliar para conectar e abrir uma aba específica."""
    if not GOOGLE_SHEETS_AVAILABLE:
        return None
    
    try:
        # Verifica se o arquivo de credenciais existe
        if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
            print(f"❌ Arquivo de credenciais não encontrado: {GOOGLE_CREDENTIALS_FILE}")
            return None
        
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ERRO CRÍTICO: Planilha '{spreadsheet_name}' não encontrada.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        print(f"ERRO CRÍTICO: Aba '{worksheet_name}' não encontrada na planilha '{spreadsheet_name}'.")
        return None
    except Exception as e:
        print(f"ATENÇÃO: Erro inesperado ao conectar ao Google Sheets: {e}")
        import traceback
        traceback.print_exc()
        return None

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_CONFIG = dict(
    host='10.0.0.10',
    port=1521,
    service='PROD',
    user='powerbi',
    password='cbjc4xp3nlq6'
)

SPREADSHEET_NAME = 'COMPRA'
ABA_COMPRAS = 'compras'

# =========================================================================
# FUNÇÕES DE BANCO DE DADOS
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    dsn = f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['service']}"
    try:
        conn = oracledb.connect(user=DB_CONFIG['user'], password=DB_CONFIG['password'], dsn=dsn)
        return conn
    except Exception as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

def buscar_compras(conn, data_inicio, data_fim):
    """
    Busca dados de compras do banco Oracle
    
    Retorna DataFrame com:
    - DATA (dia, mês, ano)
    - CODPROD
    - NOME_PRODUTO
    - QUANTIDADE_CAIXAS (quantidade comprada em caixas)
    - CUSTO
    - NF (número da nota fiscal)
    """
    query = """
    SELECT 
        TO_CHAR(N.DTENT, 'DD/MM/YYYY') AS DATA,
        M.CODPROD,
        P.DESCRICAO AS NOME_PRODUTO,
        CASE 
            WHEN NVL(P.QTUNITCX, 0) > 0 THEN 
                ROUND(M.QT / P.QTUNITCX, 2)
            ELSE 
                M.QT
        END AS QUANTIDADE_CAIXAS,
        NVL(M.PUNIT, 0) AS CUSTO,
        N.NUMNOTA AS NF
    FROM PCNFENT N
    JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
    JOIN PCPRODUT P ON M.CODPROD = P.CODPROD
    WHERE N.DTENT >= TO_DATE(:data_inicio, 'DD/MM/YYYY')
      AND N.DTENT < TO_DATE(:data_fim, 'DD/MM/YYYY')
      AND N.CODFILIAL = '1'
    ORDER BY N.DTENT DESC, N.NUMNOTA, M.CODPROD
    """
    
    try:
        df = pd.read_sql(query, conn, params={
            'data_inicio': data_inicio.strftime('%d/%m/%Y'),
            'data_fim': data_fim.strftime('%d/%m/%Y')
        })
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"❌ Erro ao buscar compras: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# FUNÇÕES DE ESCRITA - EXCEL
# =========================================================================
def escrever_excel(df, caminho_arquivo):
    """
    Escreve dados no Excel
    Cria ou atualiza a planilha COMPRA, aba compras
    """
    try:
        # Se o arquivo existe, abre e atualiza a aba
        if Path(caminho_arquivo).exists():
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=ABA_COMPRAS, index=False)
        else:
            # Cria novo arquivo
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=ABA_COMPRAS, index=False)
        
        return True
    except Exception as e:
        print(f"❌ Erro ao escrever Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

# =========================================================================
# FUNÇÕES DE ESCRITA - GOOGLE SHEETS
# =========================================================================
def preparar_cabecalho_google_sheets(worksheet):
    """Prepara o cabeçalho na planilha Google Sheets"""
    try:
        # Verifica se já existe cabeçalho
        existing_data = worksheet.get_all_values()
        if existing_data and len(existing_data) > 0:
            # Verifica se o cabeçalho está correto
            if existing_data[0] == ['DATA', 'CODPROD', 'NOME_PRODUTO', 'QUANTIDADE_CAIXAS', 'CUSTO', 'NF']:
                return  # Cabeçalho já existe e está correto
        
        # Cria cabeçalho
        cabecalho = ['DATA', 'CODPROD', 'NOME_PRODUTO', 'QUANTIDADE_CAIXAS', 'CUSTO', 'NF']
        worksheet.clear()  # Limpa a planilha
        worksheet.append_row(cabecalho)
        print("✅ Cabeçalho criado na aba compras.")
    except Exception as e:
        print(f"⚠️  Erro ao preparar cabeçalho: {e}")

def escrever_google_sheets(df):
    """
    Escreve dados no Google Sheets
    Escreve na planilha COMPRA, aba compras
    """
    if not GOOGLE_SHEETS_AVAILABLE:
        print("❌ Google Sheets não está disponível. Instale gspread: pip install gspread")
        return False
    
    print(f"📝 Escrevendo dados na aba '{ABA_COMPRAS}' da planilha '{SPREADSHEET_NAME}'...")
    worksheet = _get_worksheet(SPREADSHEET_NAME, ABA_COMPRAS)
    
    if not worksheet:
        print(f"❌ Não foi possível acessar a aba '{ABA_COMPRAS}'")
        return False
    
    # Prepara cabeçalho se necessário
    preparar_cabecalho_google_sheets(worksheet)
    
    # Prepara dados para escrita
    rows_to_append = []
    for _, row in df.iterrows():
        # Formata quantidade (com vírgula)
        qtde_cx = row.get('quantidade_caixas', 0)
        qtde_str = f"{qtde_cx:.2f}".replace('.', ',')
        
        # Formata custo (com vírgula)
        custo = row.get('custo', 0)
        custo_str = f"{custo:.2f}".replace('.', ',')
        
        row_data = [
            str(row.get('data', '')),  # DATA
            str(int(row.get('codprod', 0))),  # CODPROD
            str(row.get('nome_produto', '')),  # NOME_PRODUTO
            qtde_str,  # QUANTIDADE_CAIXAS
            custo_str,  # CUSTO
            str(row.get('nf', ''))  # NF
        ]
        rows_to_append.append(row_data)
    
    # Escreve todas as linhas de uma vez
    if rows_to_append:
        try:
            print(f"\n💾 Tentando escrever {len(rows_to_append)} linha(s)...")
            worksheet.append_rows(rows_to_append)
            print(f"✅ {len(rows_to_append)} linha(s) adicionada(s) com sucesso!")
            return True
        except Exception as e:
            print(f"❌ Erro ao escrever no Google Sheets: {e}")
            print("   Tentando escrever linha por linha...")
            
            # Tenta escrever linha por linha como fallback
            sucesso = 0
            for idx, row in enumerate(rows_to_append, start=1):
                try:
                    worksheet.append_row(row)
                    sucesso += 1
                    if idx % 100 == 0:
                        print(f"   ✅ {sucesso} linha(s) escrita(s)...")
                except Exception as e2:
                    print(f"   ❌ Erro ao escrever linha {idx}: {e2}")
            
            if sucesso > 0:
                print(f"✅ {sucesso} de {len(rows_to_append)} linha(s) escrita(s) com sucesso.")
                return True
            else:
                print(f"❌ Nenhuma linha foi escrita.")
                return False
    else:
        print("⚠️  Nenhum dado para adicionar.")
        return False

# =========================================================================
# INTERFACE GRÁFICA
# =========================================================================
class EnviarComprasPlanilha:
    def __init__(self, root):
        self.root = root
        self.root.title("Enviar Compras para Planilha")
        self.root.geometry("500x400")
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="📊 Enviar Compras para Planilha", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Frame de datas
        date_frame = ttk.LabelFrame(main_frame, text="Período", padding="10")
        date_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Data início
        ttk.Label(date_frame, text="Data Início:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.data_inicio = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        ttk.Entry(date_frame, textvariable=self.data_inicio, width=15).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(date_frame, text="(DD/MM/YYYY)").grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # Data fim
        ttk.Label(date_frame, text="Data Fim:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.data_fim = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        ttk.Entry(date_frame, textvariable=self.data_fim, width=15).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(date_frame, text="(DD/MM/YYYY)").grid(row=1, column=2, sticky=tk.W, padx=5)
        
        # Frame de opções
        options_frame = ttk.LabelFrame(main_frame, text="Destino", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.destino = tk.StringVar(value="excel")
        
        ttk.Radiobutton(options_frame, text="📄 Excel", variable=self.destino, 
                       value="excel").pack(anchor=tk.W, pady=5)
        
        # Opção Google Sheets (só se estiver disponível)
        if GOOGLE_SHEETS_AVAILABLE:
            ttk.Radiobutton(options_frame, text="☁️  Google Sheets", variable=self.destino, 
                           value="sheets").pack(anchor=tk.W, pady=5)
        else:
            ttk.Radiobutton(options_frame, text="☁️  Google Sheets (não disponível)", 
                           variable=self.destino, value="sheets", state=tk.DISABLED).pack(anchor=tk.W, pady=5)
        
        # Frame de caminho (só aparece para Excel)
        self.path_frame = ttk.LabelFrame(main_frame, text="Caminho do Arquivo Excel", padding="10")
        self.path_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.caminho_excel = tk.StringVar()
        ttk.Entry(self.path_frame, textvariable=self.caminho_excel, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(self.path_frame, text="📁", command=self.selecionar_arquivo).pack(side=tk.LEFT)
        
        # Atualiza visibilidade do frame de caminho
        self.destino.trace('w', self.atualizar_interface)
        self.atualizar_interface()
        
        # Botão executar
        self.btn_executar = ttk.Button(main_frame, text="▶️  Executar", command=self.executar)
        self.btn_executar.pack(pady=10)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para executar", 
                                     foreground="green")
        self.status_label.pack(pady=5)
        
        # Texto de informações
        info_text = ttk.Label(main_frame, 
                             text="Os dados serão enviados para:\n"
                                  "• Excel: Arquivo selecionado, aba 'compras'\n"
                                  "• Google Sheets: Planilha 'COMPRA', aba 'compras'",
                             font=("Arial", 9), foreground="gray")
        info_text.pack(pady=10)
    
    def atualizar_interface(self, *args):
        """Atualiza a interface baseado na opção selecionada"""
        if self.destino.get() == "excel":
            self.path_frame.pack(fill=tk.X, pady=(0, 15))
        else:
            self.path_frame.pack_forget()
    
    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo Excel"""
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar arquivo Excel"
        )
        if arquivo:
            self.caminho_excel.set(arquivo)
    
    def validar_datas(self):
        """Valida as datas informadas"""
        try:
            data_inicio = datetime.strptime(self.data_inicio.get(), '%d/%m/%Y')
            data_fim = datetime.strptime(self.data_fim.get(), '%d/%m/%Y')
            
            if data_fim < data_inicio:
                messagebox.showerror("Erro", "Data fim deve ser maior ou igual à data início!")
                return None, None
            
            return data_inicio, data_fim
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido! Use DD/MM/YYYY")
            return None, None
    
    def executar(self):
        """Executa o processo de busca e envio"""
        # Valida datas
        data_inicio, data_fim = self.validar_datas()
        if not data_inicio or not data_fim:
            return
        
        # Valida destino
        destino = self.destino.get()
        if destino == "excel":
            if not self.caminho_excel.get():
                messagebox.showerror("Erro", "Selecione o caminho do arquivo Excel!")
                return
        elif destino == "sheets":
            if not GOOGLE_SHEETS_AVAILABLE:
                messagebox.showerror("Erro", "Google Sheets não está disponível!\n\nInstale: pip install gspread")
                return
        
        # Desabilita botão durante execução
        self.btn_executar.config(state=tk.DISABLED)
        self.status_label.config(text="⏳ Processando...", foreground="blue")
        self.root.update()
        
        def executar_thread():
            try:
                # Conecta ao banco
                conn = get_db_connection()
                if not conn:
                    self.status_label.config(text="❌ Erro ao conectar ao banco", foreground="red")
                    self.btn_executar.config(state=tk.NORMAL)
                    return
                
                # Busca dados
                self.status_label.config(text="⏳ Buscando dados...", foreground="blue")
                self.root.update()
                
                df = buscar_compras(conn, data_inicio, data_fim)
                conn.close()
                
                if df.empty:
                    self.status_label.config(text="⚠️  Nenhum dado encontrado", foreground="orange")
                    self.btn_executar.config(state=tk.NORMAL)
                    messagebox.showwarning("Aviso", "Nenhuma compra encontrada no período informado!")
                    return
                
                # Prepara DataFrame para escrita
                df_final = df[[
                    'data', 'codprod', 'nome_produto', 
                    'quantidade_caixas', 'custo', 'nf'
                ]].copy()
                
                df_final.columns = ['DATA', 'CODPROD', 'NOME_PRODUTO', 
                                   'QUANTIDADE_CAIXAS', 'CUSTO', 'NF']
                
                # Escreve no destino escolhido
                self.status_label.config(text="⏳ Escrevendo dados...", foreground="blue")
                self.root.update()
                
                sucesso = False
                if destino == "excel":
                    sucesso = escrever_excel(df_final, self.caminho_excel.get())
                    if sucesso:
                        self.status_label.config(
                            text=f"✅ Excel gerado com sucesso! ({len(df_final)} registros)", 
                            foreground="green"
                        )
                        messagebox.showinfo("Sucesso", 
                                          f"Excel gerado com sucesso!\n\n"
                                          f"Arquivo: {self.caminho_excel.get()}\n"
                                          f"Total de registros: {len(df_final)}")
                        # Abre o arquivo
                        try:
                            os.startfile(self.caminho_excel.get())
                        except:
                            pass
                else:  # Google Sheets
                    sucesso = escrever_google_sheets(df_final)
                    if sucesso:
                        self.status_label.config(
                            text=f"✅ Dados enviados com sucesso! ({len(df_final)} registros)", 
                            foreground="green"
                        )
                        messagebox.showinfo("Sucesso", 
                                          f"Dados enviados para Google Sheets com sucesso!\n\n"
                                          f"Planilha: {SPREADSHEET_NAME}\n"
                                          f"Aba: {ABA_COMPRAS}\n"
                                          f"Total de registros: {len(df_final)}")
                    else:
                        self.status_label.config(text="❌ Erro ao enviar dados", foreground="red")
                        messagebox.showerror("Erro", "Erro ao enviar dados para Google Sheets!")
                
                if not sucesso:
                    self.status_label.config(text="❌ Erro ao processar", foreground="red")
                
            except Exception as e:
                self.status_label.config(text="❌ Erro ao processar", foreground="red")
                messagebox.showerror("Erro", f"Erro ao processar:\n{e}")
                import traceback
                traceback.print_exc()
            finally:
                self.btn_executar.config(state=tk.NORMAL)
        
        threading.Thread(target=executar_thread, daemon=True).start()

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = EnviarComprasPlanilha(root)
    root.mainloop()

