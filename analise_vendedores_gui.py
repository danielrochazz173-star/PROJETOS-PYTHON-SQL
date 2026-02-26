"""
Sistema de Análise de Vendedores - Interface Tkinter
Permite selecionar mês e gerar Excel com RCA, Nome, Faturamento, Positivação, Margem, Qtd Caixas e Peso
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime, date, timedelta, timedelta
from calendar import monthrange
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook
import os
import threading
try:
    from tkcalendar import DateEntry
except ImportError:
    messagebox.showerror("Erro", "Biblioteca tkcalendar não encontrada!\n\nExecute: pip install tkcalendar")
    DateEntry = None
# PDF: tenta WeasyPrint, senão ReportLab (puro Python)
try:
    from weasyprint import HTML, CSS  # type: ignore
    PDF_ENGINE = "weasyprint"
except Exception:
    HTML = None
    CSS = None
    PDF_ENGINE = "reportlab"
    try:
        from reportlab.lib import colors  # type: ignore
        from reportlab.lib.pagesizes import A4, landscape  # type: ignore
        from reportlab.lib.styles import getSampleStyleSheet  # type: ignore
        from reportlab.lib.units import mm  # type: ignore
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle  # type: ignore
    except Exception:
        PDF_ENGINE = None

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# =========================================================================
# FUNÇÕES DE BANCO
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        return None

def buscar_departamentos(conn):
    """Busca lista de departamentos"""
    try:
        query = "SELECT CODEPTO, DESCRICAO FROM PCDEPTO ORDER BY DESCRICAO"
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]  # Converter para minúsculas
        return df
    except Exception as e:
        print(f"Erro ao buscar departamentos: {e}")
        return pd.DataFrame()

def buscar_vendedores_periodo(conn, data_inicio, data_fim, coddepto=None, rcas_list=None, codprod_list=None):
    """Busca dados de vendedores para o período especificado
    
    Retorna DataFrame com:
    - RCA (código vendedor)
    - Nome do vendedor
    - Faturamento (líquido)
    - Positivação (quantidade de clientes diferentes que compraram)
    - SKU (quantidade de produtos distintos vendidos)
    - Margem (valor)
    - Qtd vendida (em caixas)
    - Peso vendido
    
    Args:
        conn: Conexão com o banco
        data_inicio: Data de início (datetime ou date)
        data_fim: Data de fim (datetime ou date) - exclusiva (não inclui este dia)
        coddepto: Código do departamento (opcional, None para todos)
        rcas_list: Lista de códigos RCA (opcional, None para todos)
        codprod_list: Lista de códigos de produto (opcional, None para todos)
    """
    # Converter para datetime se necessário
    if isinstance(data_inicio, date):
        data_inicio = datetime.combine(data_inicio, datetime.min.time())
    if isinstance(data_fim, date):
        data_fim = datetime.combine(data_fim, datetime.min.time())
    
    # Adicionar 1 dia à data_fim para incluir o último dia
    data_fim_inclusiva = data_fim + timedelta(days=1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim_inclusiva.strftime('%d/%m/%Y')
    
    # Filtro de departamento
    filtro_depto = ""
    if coddepto:
        filtro_depto = f"AND P.CODEPTO = {coddepto}"
    
    # Filtro de RCAs (para vendas - usa C.CODUSUR)
    filtro_rca = ""
    if rcas_list:
        # Dividir em chunks de 1000 se necessário
        chunks = [rcas_list[i:i+1000] for i in range(0, len(rcas_list), 1000)]
        if len(chunks) == 1:
            filtro_rca = f"AND C.CODUSUR IN ({', '.join(map(str, rcas_list))})"
        else:
            # Múltiplos chunks - usar OR
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"C.CODUSUR IN ({', '.join(map(str, chunk))})")
            filtro_rca = f"AND ({' OR '.join(condicoes)})"
    
    # Filtro de RCAs para devoluções (usa D.CODUSUR)
    filtro_rca_devol = ""
    if rcas_list:
        # Dividir em chunks de 1000 se necessário
        chunks = [rcas_list[i:i+1000] for i in range(0, len(rcas_list), 1000)]
        if len(chunks) == 1:
            filtro_rca_devol = f"AND D.CODUSUR IN ({', '.join(map(str, rcas_list))})"
        else:
            # Múltiplos chunks - usar OR
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"D.CODUSUR IN ({', '.join(map(str, chunk))})")
            filtro_rca_devol = f"AND ({' OR '.join(condicoes)})"
    
    # Filtro de produtos
    filtro_produto = ""
    if codprod_list:
        # Dividir em chunks de 1000 se necessário
        chunks = [codprod_list[i:i+1000] for i in range(0, len(codprod_list), 1000)]
        if len(chunks) == 1:
            filtro_produto = f"AND I.CODPROD IN ({', '.join(map(str, codprod_list))})"
        else:
            # Múltiplos chunks - usar OR
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"I.CODPROD IN ({', '.join(map(str, chunk))})")
            filtro_produto = f"AND ({' OR '.join(condicoes)})"
    
    # Filtro de produtos para devoluções
    filtro_produto_devol = ""
    if codprod_list:
        # Dividir em chunks de 1000 se necessário
        chunks = [codprod_list[i:i+1000] for i in range(0, len(codprod_list), 1000)]
        if len(chunks) == 1:
            filtro_produto_devol = f"AND D.CODPROD IN ({', '.join(map(str, codprod_list))})"
        else:
            # Múltiplos chunks - usar OR
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"D.CODPROD IN ({', '.join(map(str, chunk))})")
            filtro_produto_devol = f"AND ({' OR '.join(condicoes)})"
    
    query = f"""
    WITH CLIENTES_POSITIVADOS AS (
        SELECT DISTINCT
            C.CODUSUR,
            C.CODCLI
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          {filtro_depto}
          {filtro_rca}
          {filtro_produto}
    ),
    PRODUTOS_VENDIDOS AS (
        SELECT DISTINCT
            C.CODUSUR,
            I.CODPROD
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          {filtro_depto}
          {filtro_rca}
          {filtro_produto}
    ),
    VENDAS_VALIDAS AS (
        SELECT 
            C.CODUSUR,
            U.NOME AS NOME_VENDEDOR,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            6, 0, 11, 0, 12, 0, 
                            ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                        )
                    ELSE 0 
                END
            ) AS VALOR_BRUTO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0)
                        )
                    ELSE 0 
                END
            ) AS CUSTO_BRUTO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            6, 0, 11, 0, 12, 0, 
                            CASE 
                                WHEN NVL(P.QTUNITCX, 1) > 0 THEN 
                                    ROUND(NVL(I.QT, 0) / NVL(P.QTUNITCX, 1), 2)
                                ELSE NVL(I.QT, 0)
                            END
                        )
                    ELSE 0 
                END
            ) AS QTD_CAIXAS,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(P.PESOBRUTO, 0)
                        )
                    ELSE 0 
                END
            ) AS PESO_TOTAL
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          {filtro_depto}
          {filtro_rca}
          {filtro_produto}
        GROUP BY C.CODUSUR, U.NOME
    ),
    POSITIVACAO_POR_VENDEDOR AS (
        SELECT 
            CODUSUR,
            COUNT(*) AS POSITIVACAO
        FROM CLIENTES_POSITIVADOS
        GROUP BY CODUSUR
    ),
    SKU_POR_VENDEDOR AS (
        SELECT 
            CODUSUR,
            COUNT(*) AS SKU
        FROM PRODUTOS_VENDIDOS
        GROUP BY CODUSUR
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODUSUR,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          {filtro_depto}
          {filtro_rca_devol}
          {filtro_produto_devol}
        GROUP BY D.CODUSUR
    )
    SELECT 
        V.CODUSUR AS RCA,
        V.NOME_VENDEDOR,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO,
        NVL(P.POSITIVACAO, 0) AS POSITIVACAO,
        NVL(S.SKU, 0) AS SKU,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS MARGEM_VALOR,
        NVL(V.QTD_CAIXAS, 0) AS QTD_CAIXAS,
        NVL(V.PESO_TOTAL, 0) AS PESO_VENDIDO
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODUSUR = D.CODUSUR
    LEFT JOIN POSITIVACAO_POR_VENDEDOR P ON V.CODUSUR = P.CODUSUR
    LEFT JOIN SKU_POR_VENDEDOR S ON V.CODUSUR = S.CODUSUR
    WHERE V.VALOR_BRUTO > 0
    ORDER BY FATURAMENTO DESC
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Calcular margem percentual
        if not df.empty:
            df['margem_percentual'] = df.apply(
                lambda row: 0.0 if row['faturamento'] <= 0 
                else round(((row['margem_valor'] / row['faturamento']) * 100), 2),
                axis=1
            )
        
        return df
    except Exception as e:
        print(f"Erro ao buscar vendedores: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class AnaliseVendedoresGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Análise de Vendedores")
        self.root.geometry("500x650")
        
        # Variáveis
        self.conn = None
        self.df_departamentos = None
        
        # Criar interface
        self.criar_interface()
        
        # Conectar ao banco em thread separada
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Análise de Vendedores", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de seleção de período
        frame_periodo = ttk.LabelFrame(main_frame, text="Período de Análise", padding="15")
        frame_periodo.pack(fill=tk.X, pady=10)
        
        if DateEntry is None:
            ttk.Label(frame_periodo, text="Erro: tkcalendar não instalado. Execute: pip install tkcalendar", 
                     foreground='red').pack()
            return
        
        # Data inicial - padrão: primeiro dia do mês atual
        hoje = datetime.now()
        data_inicio_padrao = datetime(hoje.year, hoje.month, 1).date()
        data_fim_padrao = hoje.date()
        
        frame_data_inicio = ttk.Frame(frame_periodo)
        frame_data_inicio.pack(fill=tk.X, pady=5)
        ttk.Label(frame_data_inicio, text="Data Início:", width=15).pack(side=tk.LEFT)
        self.date_inicio = DateEntry(frame_data_inicio, width=12, background='darkblue',
                                     foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_inicio.set_date(data_inicio_padrao)
        self.date_inicio.pack(side=tk.LEFT, padx=5)
        
        frame_data_fim = ttk.Frame(frame_periodo)
        frame_data_fim.pack(fill=tk.X, pady=5)
        ttk.Label(frame_data_fim, text="Data Fim:", width=15).pack(side=tk.LEFT)
        self.date_fim = DateEntry(frame_data_fim, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_fim.set_date(data_fim_padrao)
        self.date_fim.pack(side=tk.LEFT, padx=5)
        
        # Frame de filtros
        frame_filtros = ttk.LabelFrame(main_frame, text="Filtros", padding="15")
        frame_filtros.pack(fill=tk.X, pady=10)
        
        # Departamento
        frame_depto_sel = ttk.Frame(frame_filtros)
        frame_depto_sel.pack(fill=tk.X, pady=5)
        ttk.Label(frame_depto_sel, text="Departamento:", width=15).pack(side=tk.LEFT)
        
        self.combo_depto = ttk.Combobox(frame_depto_sel, width=40, state="readonly")
        self.combo_depto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # RCAs
        frame_rca_sel = ttk.Frame(frame_filtros)
        frame_rca_sel.pack(fill=tk.X, pady=5)
        ttk.Label(frame_rca_sel, text="RCAs (vírgula):", width=15).pack(side=tk.LEFT)
        
        self.entry_rca = ttk.Entry(frame_rca_sel, width=50)
        self.entry_rca.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(frame_rca_sel, text="(deixe vazio para todos)").pack(side=tk.LEFT, padx=5)
        
        # Produtos
        frame_prod_sel = ttk.Frame(frame_filtros)
        frame_prod_sel.pack(fill=tk.X, pady=5)
        ttk.Label(frame_prod_sel, text="Produtos (vírgula):", width=15).pack(side=tk.LEFT)
        
        self.entry_produto = ttk.Entry(frame_prod_sel, width=50)
        self.entry_produto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(frame_prod_sel, text="(deixe vazio para todos)").pack(side=tk.LEFT, padx=5)
        
        # Botão de gerar Excel
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.btn_gerar = ttk.Button(btn_frame, text="📊 Gerar Excel", 
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pdf = ttk.Button(btn_frame, text="🖨️ Gerar PDF", 
                                  command=self.gerar_pdf, state=tk.DISABLED)
        self.btn_pdf.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    
    def conectar_banco(self):
        """Conecta ao banco e carrega dados em thread separada"""
        def carregar():
            try:
                self.status_label.config(text="Conectando ao banco...")
                self.conn = get_db_connection()
                
                if not self.conn:
                    self.status_label.config(text="❌ Erro ao conectar")
                    messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")
                    return
                
                self.status_label.config(text="Carregando departamentos...")
                self.df_departamentos = buscar_departamentos(self.conn)
                if not self.df_departamentos.empty:
                    deptos = ["Todos"] + [f"{int(row['codepto'])} - {row['descricao']}" 
                                         for _, row in self.df_departamentos.iterrows()]
                    self.combo_depto['values'] = deptos
                    self.combo_depto.current(0)
                else:
                    self.combo_depto['values'] = ["Todos"]
                    self.combo_depto.current(0)
                
                self.status_label.config(text="✅ Pronto para gerar relatório")
                self.btn_gerar.config(state=tk.NORMAL)
                if PDF_ENGINE:
                    self.btn_pdf.config(state=tk.NORMAL)
                
            except Exception as e:
                self.status_label.config(text=f"❌ Erro: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao conectar: {e}")
        
        threading.Thread(target=carregar, daemon=True).start()
    
    def gerar_excel(self):
        """Gera o arquivo Excel"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        # Obter datas selecionadas
        data_inicio = self.date_inicio.get_date()
        data_fim = self.date_fim.get_date()
        
        if data_fim < data_inicio:
            messagebox.showwarning("Aviso", "Data fim deve ser maior ou igual à data início!")
            return
        
        # Obter departamento selecionado
        depto_sel = self.combo_depto.get()
        coddepto = None
        if depto_sel and depto_sel != "Todos":
            coddepto = int(depto_sel.split(" - ")[0])
        
        # Obter RCAs selecionados
        rcas_str = self.entry_rca.get().strip()
        rcas_list = None
        if rcas_str:
            rcas_list = []
            for rca in rcas_str.split(','):
                rca_limpo = rca.strip()
                if rca_limpo.isdigit():
                    rcas_list.append(int(rca_limpo))
            
            if not rcas_list:
                rcas_list = None
        
        # Obter Produtos selecionados
        produtos_str = self.entry_produto.get().strip()
        codprod_list = None
        if produtos_str:
            codprod_list = []
            for prod in produtos_str.split(','):
                prod_limpo = prod.strip()
                if prod_limpo.isdigit():
                    codprod_list.append(int(prod_limpo))
            
            if not codprod_list:
                codprod_list = None
        
        def executar():
            try:
                self.status_label.config(text="Buscando dados...")
                self.root.update()
                
                # Nome do arquivo
                data_inicio_str = data_inicio.strftime('%Y%m%d')
                data_fim_str = data_fim.strftime('%Y%m%d')
                nome_arquivo = f"analise_vendedores_{data_inicio_str}_{data_fim_str}_{datetime.now().strftime('%H%M%S')}.xlsx"
                if coddepto:
                    nome_arquivo = f"analise_vendedores_{data_inicio_str}_{data_fim_str}_depto_{coddepto}_{datetime.now().strftime('%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                # Buscar dados do período
                df = buscar_vendedores_periodo(self.conn, data_inicio, data_fim, coddepto, rcas_list, codprod_list)
                
                if df.empty:
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado para o período selecionado")
                    self.status_label.config(text="❌ Nenhum dado encontrado")
                    return
                
                # Preparar DataFrame para Excel
                df_excel = pd.DataFrame({
                    'RCA': df['rca'],
                    'Nome do Vendedor': df['nome_vendedor'],
                    'Faturamento': df['faturamento'],
                    'Positivação': df['positivacao'],
                    'SKU': df['sku'],
                    'Margem (%)': df['margem_percentual'] / 100,
                    'Qtd Vendida (Caixas)': df['qtd_caixas'],
                    'Peso Vendido (kg)': df['peso_vendido']
                })
                
                # Ordenar por faturamento (decrescente)
                df_excel = df_excel.sort_values('Faturamento', ascending=False)
                
                # Salvar Excel
                self.status_label.config(text="Gerando Excel...")
                self.root.update()
                
                with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, sheet_name='Vendedores', index=False)
                
                self.status_label.config(text="Formatando Excel...")
                self.root.update()
                
                # Formatar Excel
                periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
                self.formatar_excel(caminho_arquivo, periodo_str, depto_sel if coddepto else "Todos")
                
                self.status_label.config(text="✅ Excel gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"Excel gerado com sucesso!\n\nArquivo: {caminho_arquivo}\n\nPeríodo: {periodo_str}")
                
                # Abrir Excel
                try:
                    os.startfile(str(caminho_arquivo))
                except:
                    pass
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar Excel:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar")
                import traceback
                traceback.print_exc()
        
        threading.Thread(target=executar, daemon=True).start()

    def gerar_pdf(self):
        """Gera PDF paginado (WeasyPrint ou ReportLab)."""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        if PDF_ENGINE is None:
            messagebox.showerror(
                "Dependência faltando",
                "Nenhum gerador de PDF disponível.\n\nOpções:\n- pip install weasyprint (requer GTK/Pango/Cairo)\n- ou: pip install reportlab (puro Python)"
            )
            return

        data_inicio = self.date_inicio.get_date()
        data_fim = self.date_fim.get_date()
        if data_fim < data_inicio:
            messagebox.showwarning("Aviso", "Data fim deve ser maior ou igual à data início!")
            return

        depto_sel = self.combo_depto.get()
        coddepto = None
        if depto_sel and depto_sel != "Todos":
            coddepto = int(depto_sel.split(" - ")[0])

        rcas_str = self.entry_rca.get().strip()
        rcas_list = None
        if rcas_str:
            rcas_list = [int(rca.strip()) for rca in rcas_str.split(',') if rca.strip().isdigit()] or None

        produtos_str = self.entry_produto.get().strip()
        codprod_list = None
        if produtos_str:
            codprod_list = [int(prod.strip()) for prod in produtos_str.split(',') if prod.strip().isdigit()] or None

        def executar():
            try:
                self.status_label.config(text="Buscando dados...")
                self.root.update()

                df = buscar_vendedores_periodo(self.conn, data_inicio, data_fim, coddepto, rcas_list, codprod_list)
                if df.empty:
                    self.status_label.config(text="❌ Nenhum dado")
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return

                df_pdf = pd.DataFrame({
                    'rca': df['rca'],
                    'nome_vendedor': df['nome_vendedor'],
                    'faturamento': df['faturamento'],
                    'positivacao': df['positivacao'],
                    'sku': df['sku'],
                    'margem_percentual': df['margem_percentual'],
                    'qtd_caixas': df['qtd_caixas'],
                    'peso_vendido': df['peso_vendido']
                }).sort_values('faturamento', ascending=False)

                data_inicio_str = data_inicio.strftime('%Y%m%d')
                data_fim_str = data_fim.strftime('%Y%m%d')
                nome_pdf = f"analise_vendedores_{data_inicio_str}_{data_fim_str}_{datetime.now().strftime('%H%M%S')}.pdf"
                caminho_pdf = Path(nome_pdf).resolve()
                periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"

                self.status_label.config(text="Gerando PDF...")
                self.root.update()
                if PDF_ENGINE == "weasyprint":
                    html_str = self._montar_html_pdf(df_pdf, periodo_str, depto_sel if coddepto else "Todos")
                    css_str = self._css_pdf()
                    HTML(string=html_str, base_url=str(Path.cwd())).write_pdf(
                        str(caminho_pdf),
                        stylesheets=[CSS(string=css_str)]
                    )
                elif PDF_ENGINE == "reportlab":
                    self._gerar_pdf_reportlab(df_pdf, caminho_pdf, periodo_str, depto_sel if coddepto else "Todos")
                else:
                    raise RuntimeError("Nenhum engine de PDF disponível")

                engine_str = "WeasyPrint" if PDF_ENGINE == "weasyprint" else "ReportLab"
                self.status_label.config(text=f"✅ PDF gerado ({engine_str})")
                messagebox.showinfo("Sucesso", f"PDF gerado com sucesso via {engine_str}!\n\nArquivo: {caminho_pdf}")
                try:
                    os.startfile(str(caminho_pdf))
                except Exception:
                    pass
            except Exception as e:
                self.status_label.config(text="❌ Erro ao gerar PDF")
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")
                import traceback
                traceback.print_exc()

        threading.Thread(target=executar, daemon=True).start()
    
    def formatar_excel(self, arquivo, periodo_str, depto_nome="Todos"):
        """Formata o arquivo Excel"""
        wb = load_workbook(arquivo)
        
        # Estilos
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )
        
        # Formatar aba
        ws = wb['Vendedores']
        
        # Formatar cabeçalho
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        # Ajustar larguras das colunas
        col_widths = {
            'A': 10,  # RCA
            'B': 35,  # Nome do Vendedor
            'C': 15,  # Faturamento
            'D': 12,  # Positivação
            'E': 10,  # SKU
            'F': 12,  # Margem (%)
            'G': 18,  # Qtd Vendida (Caixas)
            'H': 18   # Peso Vendido (kg)
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width
        
        # Formatar dados
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                col_letter = cell.column_letter
                
                if col_letter == 'A':  # RCA
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col_letter == 'B':  # Nome do Vendedor
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                elif col_letter == 'C':  # Faturamento
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'D':  # Positivação
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'E':  # SKU
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'F':  # Margem (%)
                    cell.number_format = '0.00%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'G':  # Qtd Vendida (Caixas)
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'H':  # Peso Vendido (kg)
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Adicionar título
        ws.insert_rows(1)
        ws.merge_cells(f'A1:H1')
        titulo_cell = ws['A1']
        titulo_texto = f"Análise de Vendedores - {periodo_str}"
        if depto_nome != "Todos":
            titulo_texto += f" - {depto_nome}"
        titulo_cell.value = titulo_texto
        titulo_cell.font = Font(bold=True, size=14, color="366092")
        titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Congelar primeira linha de dados (linha 2 após título)
        ws.freeze_panes = 'A3'
        
        wb.save(arquivo)

    def _css_pdf(self):
        """CSS para PDF paginado via WeasyPrint."""
        return """
        @page { 
            size: A4 landscape;
            margin: 20mm 15mm 15mm 15mm;
        }
        body { font-family: Arial, sans-serif; font-size: 11px; color: #333; }
        h1 { color: #1f4e79; margin-bottom: 5px; }
        .meta { margin-bottom: 15px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #d9d9d9; padding: 6px 8px; }
        th { background: #1f4e79; color: #fff; text-align: center; }
        td.num { text-align: right; }
        tr:nth-child(even) td { background: #f5f8fb; }
        .small { font-size: 10px; color: #555; }
        """

    def _montar_html_pdf(self, df, periodo_str, depto_sel):
        """Monta HTML para WeasyPrint."""
        linhas = ""
        for _, row in df.iterrows():
            linhas += f"""
            <tr>
                <td>{int(row['rca'])}</td>
                <td>{row['nome_vendedor']}</td>
                <td class="num">{row['faturamento']:,.2f}</td>
                <td class="num">{int(row['positivacao'])}</td>
                <td class="num">{int(row['sku'])}</td>
                <td class="num">{row['margem_percentual']:,.2f}%</td>
                <td class="num">{row['qtd_caixas']:,.2f}</td>
                <td class="num">{row['peso_vendido']:,.2f}</td>
            </tr>
            """

        depto_texto = depto_sel if depto_sel != "Todos" else "Todos os departamentos"

        return f"""
        <html>
        <head><meta charset="utf-8"></head>
        <body>
            <h1>Análise de Vendedores</h1>
            <div class="meta">
                <div><strong>Período:</strong> {periodo_str}</div>
                <div><strong>Departamento:</strong> {depto_texto}</div>
                <div class="small">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>RCA</th>
                        <th>Nome do Vendedor</th>
                        <th>Faturamento (R$)</th>
                        <th>Positivação</th>
                        <th>SKU</th>
                        <th>Margem (%)</th>
                        <th>Qtd Vendida (Caixas)</th>
                        <th>Peso Vendido (kg)</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas}
                </tbody>
            </table>
        </body>
        </html>
        """

    def _gerar_pdf_reportlab(self, df, caminho_pdf, periodo_str, depto_sel):
        """Gera PDF usando ReportLab (puro Python)."""
        doc = SimpleDocTemplate(
            str(caminho_pdf),
            pagesize=landscape(A4),
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=15 * mm,
            bottomMargin=15 * mm,
        )

        styles = getSampleStyleSheet()
        elementos = []

        titulo = "Análise de Vendedores"
        depto_texto = depto_sel if depto_sel != "Todos" else "Todos os departamentos"
        meta = f"Período: {periodo_str} | Departamento: {depto_texto}"
        gerado = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"

        elementos.append(Paragraph(f"<para align='center'><b>{titulo}</b></para>", styles['Title']))
        elementos.append(Spacer(1, 6))
        elementos.append(Paragraph(f"<para align='center'>{meta}</para>", styles['Normal']))
        elementos.append(Paragraph(f"<para align='center'><font size=9 color='#666666'>{gerado}</font></para>", styles['Normal']))
        elementos.append(Spacer(1, 12))

        header = [
            "RCA", "Nome do Vendedor", "Faturamento (R$)", "Positivação", "SKU",
            "Margem (%)", "Qtd Caixas", "Peso Vendido (kg)"
        ]

        dados = [header]
        for _, row in df.iterrows():
            dados.append([
                int(row['rca']),
                str(row['nome_vendedor']),
                f"{row['faturamento']:,.2f}",
                int(row['positivacao']),
                int(row['sku']),
                f"{row['margem_percentual']:,.2f}%",
                f"{row['qtd_caixas']:,.2f}",
                f"{row['peso_vendido']:,.2f}",
            ])

        tabela = Table(dados, repeatRows=1)
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#f5f8fb")]),
        ]))

        elementos.append(tabela)
        doc.build(elementos)

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = AnaliseVendedoresGUI(root)
    root.mainloop()
