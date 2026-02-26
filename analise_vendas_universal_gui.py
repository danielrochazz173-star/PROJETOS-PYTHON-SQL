"""
Sistema Universal de Análise de Vendas - Interface Tkinter
Permite filtrar por múltiplos critérios e gerar Excel formatado
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import oracledb
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import threading
# PDF: tenta WeasyPrint; fallback para ReportLab (puro Python)
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

def buscar_categorias(conn):
    """Busca lista de categorias"""
    try:
        query = """
        SELECT DISTINCT CODCATEGORIA, CATEGORIA 
        FROM PCCATEGORIA 
        WHERE CODCATEGORIA IS NOT NULL
        ORDER BY CATEGORIA
        """
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]  # Converter para minúsculas
        return df
    except Exception as e:
        print(f"Erro ao buscar categorias: {e}")
        return pd.DataFrame()

def buscar_vendedores(conn):
    """Busca lista de vendedores"""
    try:
        query = """
        SELECT DISTINCT CODUSUR, NOME 
        FROM PCUSUARI 
        WHERE CODUSUR IS NOT NULL
        ORDER BY NOME
        """
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]  # Converter para minúsculas
        return df
    except Exception as e:
        print(f"Erro ao buscar vendedores: {e}")
        return pd.DataFrame()

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class AnaliseVendasUniversal:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Análise Universal de Vendas")
        self.root.geometry("900x800")
        
        # Variáveis
        self.conn = None
        self.df_departamentos = None
        self.df_categorias = None
        self.df_vendedores = None
        
        # Criar interface
        self.criar_interface()
        
        # Conectar ao banco em thread separada
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        # Frame principal com scroll
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Análise Universal de Vendas", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Notebook para abas
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Aba 1: Filtros Básicos
        aba_filtros = ttk.Frame(notebook, padding="10")
        notebook.add(aba_filtros, text="Filtros Básicos")
        self.criar_aba_filtros(aba_filtros)
        
        # Aba 2: Filtros Avançados
        aba_avancados = ttk.Frame(notebook, padding="10")
        notebook.add(aba_avancados, text="Filtros Avançados")
        self.criar_aba_avancados(aba_avancados)
        
        # Botão de gerar Excel
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.btn_gerar = ttk.Button(btn_frame, text="📊 Gerar Excel", 
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pdf = ttk.Button(btn_frame, text="🖨️ Gerar PDF",
                                 command=self.gerar_pdf, state=tk.DISABLED)
        self.btn_pdf.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    def criar_aba_filtros(self, parent):
        """Cria aba de filtros básicos"""
        # Data Início
        frame_data_inicio = ttk.Frame(parent)
        frame_data_inicio.pack(fill=tk.X, pady=5)
        ttk.Label(frame_data_inicio, text="Data Início (DD/MM/AAAA):", width=25).pack(side=tk.LEFT)
        self.entry_data_inicio = ttk.Entry(frame_data_inicio, width=15)
        self.entry_data_inicio.pack(side=tk.LEFT, padx=5)
        self.entry_data_inicio.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # Data Fim
        frame_data_fim = ttk.Frame(parent)
        frame_data_fim.pack(fill=tk.X, pady=5)
        ttk.Label(frame_data_fim, text="Data Fim (DD/MM/AAAA):", width=25).pack(side=tk.LEFT)
        self.entry_data_fim = ttk.Entry(frame_data_fim, width=15)
        self.entry_data_fim.pack(side=tk.LEFT, padx=5)
        self.entry_data_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # Produtos
        frame_produtos = ttk.Frame(parent)
        frame_produtos.pack(fill=tk.X, pady=5)
        ttk.Label(frame_produtos, text="Códigos Produtos (vírgula):", width=25).pack(side=tk.LEFT)
        self.entry_produtos = ttk.Entry(frame_produtos, width=50)
        self.entry_produtos.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(frame_produtos, text="(deixe vazio para todos)").pack(side=tk.LEFT, padx=5)
        
        # Preço Mínimo
        frame_preco_min = ttk.Frame(parent)
        frame_preco_min.pack(fill=tk.X, pady=5)
        ttk.Label(frame_preco_min, text="Preço Mínimo:", width=25).pack(side=tk.LEFT)
        self.entry_preco_min = ttk.Entry(frame_preco_min, width=15)
        self.entry_preco_min.pack(side=tk.LEFT, padx=5)
        
        # Preço Máximo
        frame_preco_max = ttk.Frame(parent)
        frame_preco_max.pack(fill=tk.X, pady=5)
        ttk.Label(frame_preco_max, text="Preço Máximo:", width=25).pack(side=tk.LEFT)
        self.entry_preco_max = ttk.Entry(frame_preco_max, width=15)
        self.entry_preco_max.pack(side=tk.LEFT, padx=5)
        
        # Preço Específico
        frame_preco_esp = ttk.Frame(parent)
        frame_preco_esp.pack(fill=tk.X, pady=5)
        ttk.Label(frame_preco_esp, text="Preço Específico:", width=25).pack(side=tk.LEFT)
        self.entry_preco_esp = ttk.Entry(frame_preco_esp, width=15)
        self.entry_preco_esp.pack(side=tk.LEFT, padx=5)
        ttk.Label(frame_preco_esp, text="(sobrescreve min/max)").pack(side=tk.LEFT, padx=5)
    
    def criar_aba_avancados(self, parent):
        """Cria aba de filtros avançados"""
        # Departamento
        frame_depto = ttk.Frame(parent)
        frame_depto.pack(fill=tk.X, pady=5)
        ttk.Label(frame_depto, text="Departamento:", width=25).pack(side=tk.LEFT)
        self.combo_depto = ttk.Combobox(frame_depto, width=40, state="readonly")
        self.combo_depto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Categoria
        frame_cat = ttk.Frame(parent)
        frame_cat.pack(fill=tk.X, pady=5)
        ttk.Label(frame_cat, text="Categoria:", width=25).pack(side=tk.LEFT)
        self.combo_cat = ttk.Combobox(frame_cat, width=40, state="readonly")
        self.combo_cat.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Vendedor
        frame_vendedor = ttk.Frame(parent)
        frame_vendedor.pack(fill=tk.X, pady=5)
        ttk.Label(frame_vendedor, text="Vendedor:", width=25).pack(side=tk.LEFT)
        self.combo_vendedor = ttk.Combobox(frame_vendedor, width=40, state="readonly")
        self.combo_vendedor.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Cliente (código)
        frame_cliente = ttk.Frame(parent)
        frame_cliente.pack(fill=tk.X, pady=5)
        ttk.Label(frame_cliente, text="Código Cliente:", width=25).pack(side=tk.LEFT)
        self.entry_cliente = ttk.Entry(frame_cliente, width=15)
        self.entry_cliente.pack(side=tk.LEFT, padx=5)
        
        # Condição de Venda
        frame_cond = ttk.Frame(parent)
        frame_cond.pack(fill=tk.X, pady=5)
        ttk.Label(frame_cond, text="Condição Venda:", width=25).pack(side=tk.LEFT)
        self.combo_cond = ttk.Combobox(frame_cond, width=40, state="readonly",
                                      values=["Todas", "Excluir 4,8,10,13,20,98,99"])
        self.combo_cond.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.combo_cond.current(1)  # Padrão: excluir
        
        # Filial
        frame_filial = ttk.Frame(parent)
        frame_filial.pack(fill=tk.X, pady=5)
        ttk.Label(frame_filial, text="Filial:", width=25).pack(side=tk.LEFT)
        self.combo_filial = ttk.Combobox(frame_filial, width=40, state="readonly",
                                        values=["Todas", "1", "98", "1 e 98"])
        self.combo_filial.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.combo_filial.current(3)  # Padrão: 1 e 98
    
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
                
                self.status_label.config(text="Carregando categorias...")
                self.df_categorias = buscar_categorias(self.conn)
                if not self.df_categorias.empty:
                    cats = ["Todas"] + [f"{int(row['codcategoria'])} - {row['categoria']}" 
                                       for _, row in self.df_categorias.iterrows()]
                    self.combo_cat['values'] = cats
                    self.combo_cat.current(0)
                else:
                    self.combo_cat['values'] = ["Todas"]
                    self.combo_cat.current(0)
                
                self.status_label.config(text="Carregando vendedores...")
                self.df_vendedores = buscar_vendedores(self.conn)
                if not self.df_vendedores.empty:
                    vends = ["Todos"] + [f"{int(row['codusur'])} - {row['nome']}" 
                                        for _, row in self.df_vendedores.iterrows()]
                    self.combo_vendedor['values'] = vends
                    self.combo_vendedor.current(0)
                else:
                    self.combo_vendedor['values'] = ["Todos"]
                    self.combo_vendedor.current(0)
                
                self.status_label.config(text="✅ Pronto para gerar relatório")
                self.btn_gerar.config(state=tk.NORMAL)
                if PDF_ENGINE:
                    self.btn_pdf.config(state=tk.NORMAL)
                
            except Exception as e:
                self.status_label.config(text=f"❌ Erro: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
        
        threading.Thread(target=carregar, daemon=True).start()
    
    def montar_query(self):
        """Monta a query SQL baseada nos filtros"""
        # Datas
        data_inicio = self.entry_data_inicio.get().strip()
        data_fim = self.entry_data_fim.get().strip()
        
        try:
            datetime.strptime(data_inicio, "%d/%m/%Y")
            datetime.strptime(data_fim, "%d/%m/%Y")
        except:
            raise ValueError("Datas inválidas. Use formato DD/MM/AAAA")
        
        # Produtos
        produtos_str = self.entry_produtos.get().strip()
        filtro_produtos = ""
        if produtos_str:
            codprod_list = []
            for cod in produtos_str.split(','):
                cod_limpo = cod.strip()
                if cod_limpo.isdigit():
                    codprod_list.append(int(cod_limpo))
            
            if codprod_list:
                # Dividir em chunks de 1000
                chunks = [codprod_list[i:i+1000] for i in range(0, len(codprod_list), 1000)]
                if len(chunks) == 1:
                    filtro_produtos = f"AND I.CODPROD IN ({', '.join(map(str, codprod_list))})"
                else:
                    # Múltiplos chunks - usar OR
                    condicoes = []
                    for chunk in chunks:
                        condicoes.append(f"I.CODPROD IN ({', '.join(map(str, chunk))})")
                    filtro_produtos = f"AND ({' OR '.join(condicoes)})"
        
        # Preço
        filtro_preco = ""
        preco_esp = self.entry_preco_esp.get().strip()
        if preco_esp:
            try:
                preco_val = float(preco_esp.replace(',', '.'))
                filtro_preco = f"AND ROUND(I.PVENDA, 2) = {preco_val:.2f}"
            except:
                pass
        else:
            preco_min = self.entry_preco_min.get().strip()
            preco_max = self.entry_preco_max.get().strip()
            if preco_min:
                try:
                    min_val = float(preco_min.replace(',', '.'))
                    filtro_preco += f"AND I.PVENDA >= {min_val:.2f}"
                except:
                    pass
            if preco_max:
                try:
                    max_val = float(preco_max.replace(',', '.'))
                    filtro_preco += f"AND I.PVENDA <= {max_val:.2f}"
                except:
                    pass
        
        # Departamento
        filtro_depto = ""
        depto_sel = self.combo_depto.get()
        if depto_sel and depto_sel != "Todos":
            coddepto = depto_sel.split(" - ")[0]
            filtro_depto = f"AND P.CODEPTO = {coddepto}"
        
        # Categoria
        filtro_cat = ""
        cat_sel = self.combo_cat.get()
        if cat_sel and cat_sel != "Todas":
            codcat = cat_sel.split(" - ")[0]
            filtro_cat = f"AND P.CODCATEGORIA = {codcat}"
        
        # Vendedor
        filtro_vendedor = ""
        vend_sel = self.combo_vendedor.get()
        if vend_sel and vend_sel != "Todos":
            codvend = vend_sel.split(" - ")[0]
            filtro_vendedor = f"AND C.CODUSUR = {codvend}"
        
        # Cliente
        filtro_cliente = ""
        cliente_str = self.entry_cliente.get().strip()
        if cliente_str and cliente_str.isdigit():
            filtro_cliente = f"AND C.CODCLI = {cliente_str}"
        
        # Condição de Venda
        filtro_cond = ""
        cond_sel = self.combo_cond.get()
        if cond_sel == "Excluir 4,8,10,13,20,98,99":
            filtro_cond = "AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)"
        
        # Filial
        filtro_filial = ""
        filial_sel = self.combo_filial.get()
        if filial_sel == "1":
            filtro_filial = "AND C.CODFILIAL = '1'"
        elif filial_sel == "98":
            filtro_filial = "AND C.CODFILIAL = '98'"
        elif filial_sel == "1 e 98":
            filtro_filial = "AND C.CODFILIAL IN ('1', '98')"
        
        # Montar query
        query = f"""
        SELECT 
            C.NUMNOTA AS NF,
            C.DATA,
            I.CODPROD,
            P.DESCRICAO AS PRODUTO,
            I.PVENDA AS PRECO,
            I.QT AS QUANTIDADE,
            NVL(P.QTUNITCX, 1) AS QTUNITCX,
            CASE 
                WHEN NVL(P.QTUNITCX, 1) > 1 THEN ROUND(I.QT * NVL(P.QTUNITCX, 1), 2)
                ELSE I.QT
            END AS QUANTIDADE_UNIDADE,
            P.UNIDADE,
            P.EMBALAGEM,
            ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL,
            ROUND(I.QT * NVL(I.VLCUSTOFIN, 0), 2) AS CUSTO_TOTAL,
            CL.CODCLI,
            CL.CLIENTE,
            C.CODUSUR,
            U.NOME AS VENDEDOR,
            D.DESCRICAO AS DEPARTAMENTO,
            CAT.CATEGORIA
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        LEFT JOIN PCCATEGORIA CAT ON P.CODCATEGORIA = CAT.CODCATEGORIA AND P.CODSEC = CAT.CODSEC
        WHERE C.DATA BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
            AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
            {filtro_produtos}
            {filtro_preco}
            {filtro_depto}
            {filtro_cat}
            {filtro_vendedor}
            {filtro_cliente}
            {filtro_cond}
            {filtro_filial}
            AND C.POSICAO = 'F'
            AND C.DTCANCEL IS NULL
            AND NVL(I.BONIFIC, 'N') = 'N'
        ORDER BY C.DATA DESC, C.NUMNOTA, I.CODPROD
        """
        
        return query
    
    def gerar_excel(self):
        """Gera o arquivo Excel"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        def executar():
            try:
                self.status_label.config(text="Montando query...")
                query = self.montar_query()
                
                self.status_label.config(text="Executando query...")
                df = pd.read_sql_query(query, self.conn)
                
                if df.empty:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado com os filtros especificados")
                    self.status_label.config(text="❌ Nenhum registro encontrado")
                    return
                
                self.status_label.config(text=f"Processando {len(df)} registros...")
                
                # Ajustar coluna de caixas:
                # Coluna G passa a ser "CAIXAS" = quantidade em unidades / unidades por caixa
                if 'QUANTIDADE_UNIDADE' in df.columns and 'QTUNITCX' in df.columns:
                    def calcular_caixas(row):
                        qtunitcx = row['QTUNITCX'] or 0
                        if qtunitcx == 0:
                            # Se não tiver unidades por caixa, considera a própria quantidade em unidades
                            return row['QUANTIDADE_UNIDADE']
                        return round(row['QUANTIDADE_UNIDADE'] / qtunitcx, 2)
                    
                    df['CAIXAS'] = df.apply(calcular_caixas, axis=1)
                    
                    cols = list(df.columns)
                    # Garante que CAIXAS só apareça uma vez
                    if 'CAIXAS' in cols:
                        cols.remove('CAIXAS')
                    # Insere CAIXAS exatamente na posição onde estava QTUNITCX (coluna G)
                    if 'QTUNITCX' in cols:
                        idx_qtunit = cols.index('QTUNITCX')
                    else:
                        idx_qtunit = 6  # posição G como fallback
                    cols.insert(idx_qtunit, 'CAIXAS')
                    df = df[cols]
                
                # Converter data
                if 'DATA' in df.columns:
                    df['DATA'] = pd.to_datetime(df['DATA'])
                
                # Calcular margem por item
                if 'CUSTO_TOTAL' in df.columns:
                    df['MARGEM_VALOR'] = df['VALOR_TOTAL'] - df['CUSTO_TOTAL']
                    df['MARGEM_PERCENT'] = df.apply(
                        lambda row: 0.0 if row['VALOR_TOTAL'] <= 0 
                        else round(((row['MARGEM_VALOR'] / row['VALOR_TOTAL']) * 100), 2),
                        axis=1
                    )
                
                # Resumo por NF
                resumo_nf = df.groupby(['NF', 'DATA', 'CLIENTE'], as_index=False).agg(
                    QUANTIDADE=('QUANTIDADE', 'sum'),
                    VALOR_TOTAL=('VALOR_TOTAL', 'sum'),
                    CUSTO_TOTAL=('CUSTO_TOTAL', 'sum'),
                )
                # Calcular margem no resumo
                resumo_nf['MARGEM_VALOR'] = resumo_nf['VALOR_TOTAL'] - resumo_nf['CUSTO_TOTAL']
                resumo_nf['MARGEM_PERCENT'] = resumo_nf.apply(
                    lambda row: 0.0 if row['VALOR_TOTAL'] <= 0 
                    else round(((row['MARGEM_VALOR'] / row['VALOR_TOTAL']) * 100), 2),
                    axis=1
                )
                # Reordenar colunas: NF, DATA, CLIENTE, QUANTIDADE, VALOR_TOTAL, MARGEM_PERCENT, CUSTO_TOTAL, MARGEM_VALOR
                resumo_nf = resumo_nf[['NF', 'DATA', 'CLIENTE', 'QUANTIDADE', 'VALOR_TOTAL', 'MARGEM_PERCENT', 'CUSTO_TOTAL', 'MARGEM_VALOR']]
                resumo_nf = resumo_nf.sort_values('VALOR_TOTAL', ascending=False)
                
                # Resumo por Produto
                resumo_produto = df.groupby(['CODPROD', 'PRODUTO', 'DEPARTAMENTO'], as_index=False).agg(
                    VALOR_TOTAL=('VALOR_TOTAL', 'sum'),
                    CUSTO_TOTAL=('CUSTO_TOTAL', 'sum'),
                )
                # Calcular margem no resumo por produto
                resumo_produto['MARGEM_VALOR'] = resumo_produto['VALOR_TOTAL'] - resumo_produto['CUSTO_TOTAL']
                resumo_produto['MARGEM_PERCENT'] = resumo_produto.apply(
                    lambda row: 0.0 if row['VALOR_TOTAL'] <= 0 
                    else round(((row['MARGEM_VALOR'] / row['VALOR_TOTAL']) * 100), 2),
                    axis=1
                )
                # Reordenar colunas: CODPROD, PRODUTO, DEPARTAMENTO, VALOR_TOTAL, CUSTO_TOTAL, MARGEM_PERCENT, MARGEM_VALOR
                resumo_produto = resumo_produto[['CODPROD', 'PRODUTO', 'DEPARTAMENTO', 'VALOR_TOTAL', 'CUSTO_TOTAL', 'MARGEM_PERCENT', 'MARGEM_VALOR']]
                resumo_produto = resumo_produto.sort_values('VALOR_TOTAL', ascending=False)
                
                # Salvar Excel
                nome_arquivo = f"vendas_universal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                self.status_label.config(text="Gerando Excel...")
                self.formatar_excel(caminho_arquivo, df, resumo_nf, resumo_produto)
                
                self.status_label.config(text="✅ Excel gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"Excel gerado:\n{caminho_arquivo}")
                
                # Abrir Excel
                try:
                    os.startfile(str(caminho_arquivo))
                except:
                    pass
                    
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
                self.status_label.config(text="❌ Erro na validação")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar Excel:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar")
        
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

        def executar():
            try:
                self.status_label.config(text="Montando query...")
                query = self.montar_query()

                self.status_label.config(text="Executando query...")
                df = pd.read_sql_query(query, self.conn)

                if df.empty:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado com os filtros especificados")
                    self.status_label.config(text="❌ Nenhum registro encontrado")
                    return

                self.status_label.config(text=f"Processando {len(df)} registros...")

                # Converter data
                if 'DATA' in df.columns:
                    df['DATA'] = pd.to_datetime(df['DATA'])

                # Calcular margem por item
                if 'CUSTO_TOTAL' in df.columns:
                    df['MARGEM_VALOR'] = df['VALOR_TOTAL'] - df['CUSTO_TOTAL']
                    df['MARGEM_PERCENT'] = df.apply(
                        lambda row: 0.0 if row['VALOR_TOTAL'] <= 0 
                        else round(((row['MARGEM_VALOR'] / row['VALOR_TOTAL']) * 100), 2),
                        axis=1
                    )

                resumo_nf = df.groupby(['NF', 'DATA', 'CLIENTE'], as_index=False).agg(
                    QUANTIDADE=('QUANTIDADE', 'sum'),
                    VALOR_TOTAL=('VALOR_TOTAL', 'sum'),
                    CUSTO_TOTAL=('CUSTO_TOTAL', 'sum'),
                )
                resumo_nf['MARGEM_VALOR'] = resumo_nf['VALOR_TOTAL'] - resumo_nf['CUSTO_TOTAL']
                resumo_nf['MARGEM_PERCENT'] = resumo_nf.apply(
                    lambda row: 0.0 if row['VALOR_TOTAL'] <= 0 
                    else round(((row['MARGEM_VALOR'] / row['VALOR_TOTAL']) * 100), 2),
                    axis=1
                )
                resumo_nf = resumo_nf[['NF', 'DATA', 'CLIENTE', 'QUANTIDADE', 'VALOR_TOTAL', 'MARGEM_PERCENT', 'CUSTO_TOTAL', 'MARGEM_VALOR']]
                resumo_nf = resumo_nf.sort_values('VALOR_TOTAL', ascending=False)

                nome_pdf = f"vendas_universal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                caminho_pdf = Path(nome_pdf).resolve()
                periodo_str = f"{self.entry_data_inicio.get().strip()} a {self.entry_data_fim.get().strip()}"

                self.status_label.config(text="Gerando PDF...")
                self.root.update()
                if PDF_ENGINE == "weasyprint":
                    html_str = self._montar_html_pdf(resumo_nf, periodo_str, len(df))
                    css_str = self._css_pdf()
                    HTML(string=html_str, base_url=str(Path.cwd())).write_pdf(
                        str(caminho_pdf),
                        stylesheets=[CSS(string=css_str)]
                    )
                elif PDF_ENGINE == "reportlab":
                    self._gerar_pdf_reportlab(resumo_nf, caminho_pdf, periodo_str, len(df))
                else:
                    raise RuntimeError("Nenhum engine de PDF disponível")

                engine_str = "WeasyPrint" if PDF_ENGINE == "weasyprint" else "ReportLab"
                self.status_label.config(text=f"✅ PDF gerado ({engine_str})")
                messagebox.showinfo("Sucesso", f"PDF gerado com sucesso via {engine_str}!\n\nArquivo: {caminho_pdf}")
                try:
                    os.startfile(str(caminho_pdf))
                except Exception:
                    pass
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
                self.status_label.config(text="❌ Erro na validação")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar PDF")

        threading.Thread(target=executar, daemon=True).start()
    
    def formatar_excel(self, arquivo, df, resumo_nf, resumo_produto):
        """Formata o arquivo Excel"""
        from openpyxl import load_workbook
        
        with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
            # Aba Detalhado
            df.to_excel(writer, sheet_name='Detalhado', index=False)
            ws = writer.sheets['Detalhado']
            
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin'),
            )
            
            # Formatar cabeçalho
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            # Ajustar larguras
            col_widths = {
                'A': 12, 'B': 12, 'C': 10, 'D': 40, 'E': 12,
                'F': 10, 'G': 15, 'H': 10, 'I': 40, 'J': 10,
                'K': 25, 'L': 25, 'M': 25
            }
            for col, width in col_widths.items():
                ws.column_dimensions[col].width = width
            
            # Formatar dados
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.border = border
                    col_letter = get_column_letter(cell.column)
                    col_name = ws.cell(row=1, column=cell.column).value
                    
                    # Formatar PRECO (coluna E = 5)
                    if cell.column == 5 or col_name == 'PRECO':
                        cell.number_format = 'R$ #,##0.00'
                    # Formatar VALOR_TOTAL (coluna K = 11)
                    elif cell.column == 11 or col_name == 'VALOR_TOTAL':
                        cell.number_format = 'R$ #,##0.00'
                    # Formatar CUSTO_TOTAL (coluna L = 12)
                    elif cell.column == 12 or col_name == 'CUSTO_TOTAL':
                        cell.number_format = 'R$ #,##0.00'
                    # Formatar MARGEM_VALOR (coluna S = 19)
                    elif cell.column == 19 or col_name == 'MARGEM_VALOR':
                        cell.number_format = 'R$ #,##0.00'
                    # Formatar MARGEM_PERCENT (coluna T = 20) - formato número sem %
                    elif cell.column == 20 or col_name == 'MARGEM_PERCENT':
                        cell.number_format = '#,##0.00'
                    # Formatar DATA (coluna B = 2)
                    elif cell.column == 2 or col_name == 'DATA':
                        cell.number_format = 'DD/MM/YYYY'
            
            # Aba Resumo
            resumo_nf.to_excel(writer, sheet_name='Resumo por NF', index=False)
            ws_resumo = writer.sheets['Resumo por NF']
            
            for cell in ws_resumo[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            ws_resumo.column_dimensions['A'].width = 12  # NF
            ws_resumo.column_dimensions['B'].width = 12  # DATA
            ws_resumo.column_dimensions['C'].width = 40  # CLIENTE
            ws_resumo.column_dimensions['D'].width = 12  # QUANTIDADE
            ws_resumo.column_dimensions['E'].width = 15  # VALOR_TOTAL
            ws_resumo.column_dimensions['F'].width = 12  # MARGEM_PERCENT
            ws_resumo.column_dimensions['G'].width = 15  # CUSTO_TOTAL
            ws_resumo.column_dimensions['H'].width = 15  # MARGEM_VALOR
            
            for row in ws_resumo.iter_rows(min_row=2, max_row=ws_resumo.max_row):
                for cell in row:
                    cell.border = border
                    col_letter = get_column_letter(cell.column)
                    if col_letter == 'B':  # DATA
                        cell.number_format = 'DD/MM/YYYY'
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col_letter == 'D':  # QUANTIDADE
                        cell.number_format = '#,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'E':  # VALOR_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'F':  # MARGEM_PERCENT - formato número sem %
                        cell.number_format = '#,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'G':  # CUSTO_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'H':  # MARGEM_VALOR
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'C':  # CLIENTE
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Aba Resumo por Produto
            resumo_produto.to_excel(writer, sheet_name='Resumo por Produto', index=False)
            ws_produto = writer.sheets['Resumo por Produto']
            
            for cell in ws_produto[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            ws_produto.column_dimensions['A'].width = 12  # CODPROD
            ws_produto.column_dimensions['B'].width = 50  # PRODUTO
            ws_produto.column_dimensions['C'].width = 30  # DEPARTAMENTO
            ws_produto.column_dimensions['D'].width = 15  # VALOR_TOTAL
            ws_produto.column_dimensions['E'].width = 15  # CUSTO_TOTAL
            ws_produto.column_dimensions['F'].width = 12  # MARGEM_PERCENT
            ws_produto.column_dimensions['G'].width = 15  # MARGEM_VALOR
            
            for row in ws_produto.iter_rows(min_row=2, max_row=ws_produto.max_row):
                for cell in row:
                    cell.border = border
                    col_letter = get_column_letter(cell.column)
                    if col_letter == 'A':  # CODPROD
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col_letter == 'B':  # PRODUTO
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    elif col_letter == 'C':  # DEPARTAMENTO
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    elif col_letter == 'D':  # VALOR_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'E':  # CUSTO_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'F':  # MARGEM_PERCENT
                        cell.number_format = '#,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif col_letter == 'G':  # MARGEM_VALOR
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
            
            # Congelar primeira linha
            ws.freeze_panes = 'A2'
            ws_resumo.freeze_panes = 'A2'
            ws_produto.freeze_panes = 'A2'

    def _css_pdf(self):
        """CSS para PDF (WeasyPrint)."""
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

    def _montar_html_pdf(self, resumo_df, periodo_str, total_itens):
        """Monta HTML para PDF (WeasyPrint) usando resumo por NF."""
        linhas = ""
        for _, row in resumo_df.iterrows():
            data_fmt = row['DATA'].strftime('%d/%m/%Y') if not pd.isna(row['DATA']) else ""
            linhas += f"""
            <tr>
                <td>{row['NF']}</td>
                <td>{data_fmt}</td>
                <td>{row['CLIENTE']}</td>
                <td class="num">{row['QUANTIDADE']:,.2f}</td>
                <td class="num">{row['VALOR_TOTAL']:,.2f}</td>
                <td class="num">{row['MARGEM_PERCENT']:,.2f}%</td>
                <td class="num">{row['CUSTO_TOTAL']:,.2f}</td>
                <td class="num">{row['MARGEM_VALOR']:,.2f}</td>
            </tr>
            """

        return f"""
        <html>
        <head><meta charset="utf-8"></head>
        <body>
            <h1>Análise Universal de Vendas</h1>
            <div class="meta">
                <div><strong>Período:</strong> {periodo_str}</div>
                <div><strong>Pedidos (linhas detalhadas):</strong> {total_itens}</div>
                <div class="small">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>NF</th>
                        <th>Data</th>
                        <th>Cliente</th>
                        <th>Qtd</th>
                        <th>Valor Total (R$)</th>
                        <th>Margem (%)</th>
                        <th>Custo Total (R$)</th>
                        <th>Margem (R$)</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas}
                </tbody>
            </table>
        </body>
        </html>
        """

    def _gerar_pdf_reportlab(self, resumo_df, caminho_pdf, periodo_str, total_itens):
        """Gera PDF via ReportLab (puro Python)."""
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

        titulo = "Análise Universal de Vendas"
        meta = f"Período: {periodo_str} | Linhas detalhadas: {total_itens}"
        gerado = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"

        elementos.append(Paragraph(f"<para align='center'><b>{titulo}</b></para>", styles['Title']))
        elementos.append(Spacer(1, 6))
        elementos.append(Paragraph(f"<para align='center'>{meta}</para>", styles['Normal']))
        elementos.append(Paragraph(f"<para align='center'><font size=9 color='#666666'>{gerado}</font></para>", styles['Normal']))
        elementos.append(Spacer(1, 12))

        header = ["NF", "Data", "Cliente", "Qtd", "Valor Total (R$)", "Margem (%)", "Custo Total (R$)", "Margem (R$)"]
        dados = [header]
        for _, row in resumo_df.iterrows():
            data_fmt = row['DATA'].strftime('%d/%m/%Y') if not pd.isna(row['DATA']) else ""
            dados.append([
                row['NF'],
                data_fmt,
                str(row['CLIENTE']),
                f"{row['QUANTIDADE']:,.2f}",
                f"{row['VALOR_TOTAL']:,.2f}",
                f"{row['MARGEM_PERCENT']:,.2f}%",
                f"{row['CUSTO_TOTAL']:,.2f}",
                f"{row['MARGEM_VALOR']:,.2f}",
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
            ('ALIGN', (3, 1), (-1, -1), 'RIGHT'),
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
    app = AnaliseVendasUniversal(root)
    root.mainloop()

# portfolio-commit-ready: analise_vendas_universal_gui
