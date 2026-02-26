import oracledb
import pandas as pd
from pathlib import Path
from datetime import datetime
import locale
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer
from dateutil.relativedelta import relativedelta
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

from google_sheets_api import get_vilog_stock_data

# ======= CONFIGURAÇÃO DE ACESSO ORACLE =======
DB_CONFIG = dict(
    host='10.0.0.10',
    port=1521,
    service='PROD',
    user='powerbi',
    password='cbjc4xp3nlq6'
)

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '')

SQL_PRODUTOS = """
SELECT 
    P.CODPROD, 
    P.DESCRICAO,
    P.EMBALAGEM,
    P.QTUNITCX,
    D.DESCRICAO AS DEPARTAMENTO,
    COALESCE((E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA), 0) AS DISPONIVEL_UN,
    E.CUSTOFIN AS CUSTO_FINAL_UNITARIO
FROM PCPRODUT P
LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
WHERE P.DTEXCLUSAO IS NULL
  AND (P.OBS2 <> 'FL' OR P.OBS2 IS NULL)
  AND P.REVENDA = 'S'
ORDER BY D.DESCRICAO ASC, P.DESCRICAO ASC
"""

SQL_COMPRAS = """
SELECT
    M.CODPROD,
    TRUNC(N.DTENT, 'MM') AS MES_REF,
    SUM(M.QT) AS QT_COMPRADA
FROM PCNFENT N
JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
WHERE N.DTENT >= TRUNC(ADD_MONTHS(SYSDATE, -4), 'MM')
GROUP BY M.CODPROD, TRUNC(N.DTENT, 'MM')
"""


def get_db_connection():
    dsn = f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['service']}"
    try:
        conn = oracledb.connect(user=DB_CONFIG['user'], password=DB_CONFIG['password'], dsn=dsn)
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco Oracle: {e}")
        sys.exit(1)


def fetch_products():
    conn = get_db_connection()
    try:
        df = pd.read_sql(SQL_PRODUTOS, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    finally:
        conn.close()


def fetch_compras():
    conn = get_db_connection()
    try:
        comp = pd.read_sql(SQL_COMPRAS, conn)
        comp.columns = [x.lower() for x in comp.columns]
        return comp
    finally:
        conn.close()


def get_tres_meses_anteriores():
    hoje = datetime.now()
    meses = []
    for retro in range(3, 0, -1):
        datames = hoje - relativedelta(months=retro)
        meses.append((datames.strftime('%Y%m'), datames.strftime('%b').capitalize()))
    return meses


def format_real(valor):
    try:
        return f"{float(valor):,.2f}".replace('.', ',')
    except:
        return "0,00"


class FaltaProdutosInterface:
    def __init__(self, master):
        self.master = master
        
        # Variáveis
        self.df_produtos = None
        self.df_compras = None
        self.estoque_vilog = None
        self.produtos_filtrados = None
        self.produtos_selecionados = []
        
        # Criar frame principal
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)
        
        # Carregar dados
        self.carregar_dados()
        
        # Criar interface
        self.criar_interface()
        
    def get_frame(self):
        """Retorna o frame principal da interface"""
        return self.main_frame
        
    def carregar_dados(self):
        """Carrega os dados do banco e planilha"""
        try:
            # Mostrar progresso
            self.master.config(cursor="wait")
            
            # Carregar dados
            self.df_produtos = fetch_products()
            self.df_compras = fetch_compras()
            self.estoque_vilog = get_vilog_stock_data()
            self.estoque_vilog = {str(int(str(k).strip())): v for k, v in self.estoque_vilog.items() if k}
            
            self.df_produtos['departamento'] = self.df_produtos['departamento'].fillna('Sem Departamento')
            
            messagebox.showinfo("Sucesso", "Dados carregados com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
        finally:
            self.master.config(cursor="")
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        # Título
        titulo = ttk.Label(self.main_frame, text="Sistema de Relatórios de Produtos", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Frame de busca
        self.busca_frame = ttk.LabelFrame(self.main_frame, text="Busca de Produtos", padding="10")
        self.busca_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.busca_frame.columnconfigure(1, weight=1)
        
        # Opções de busca
        ttk.Label(self.busca_frame, text="Tipo de busca:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.tipo_busca = tk.StringVar(value="departamento")
        ttk.Radiobutton(self.busca_frame, text="Por Departamento", variable=self.tipo_busca, 
                       value="departamento", command=self.atualizar_busca).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(self.busca_frame, text="Por Nome do Produto", variable=self.tipo_busca, 
                       value="nome", command=self.atualizar_busca).grid(row=0, column=2, sticky=tk.W)
        
        # Campo de busca
        ttk.Label(self.busca_frame, text="Buscar:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.campo_busca = ttk.Entry(self.busca_frame, width=40)
        self.campo_busca.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        self.campo_busca.bind('<KeyRelease>', self.filtrar_produtos)
        
        # Botão buscar
        ttk.Button(self.busca_frame, text="Buscar", command=self.filtrar_produtos).grid(row=1, column=2, padx=(10, 0), pady=(10, 0))
        
        # Frame da lista de produtos
        lista_frame = ttk.LabelFrame(self.main_frame, text="Produtos Encontrados", padding="10")
        lista_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
        
        # Treeview para produtos
        colunas = ('selecionado', 'codigo', 'produto', 'departamento', 'embalagem', 'estoque_dicon', 'estoque_vilog', 'total')
        self.tree = ttk.Treeview(lista_frame, columns=colunas, show='headings', height=15)
        
        # Configurar colunas
        self.tree.heading('selecionado', text='✓')
        self.tree.heading('codigo', text='Código')
        self.tree.heading('produto', text='Produto')
        self.tree.heading('departamento', text='Departamento')
        self.tree.heading('embalagem', text='Embalagem')
        self.tree.heading('estoque_dicon', text='Est. DICON')
        self.tree.heading('estoque_vilog', text='Est. VILOG')
        self.tree.heading('total', text='Total')
        
        self.tree.column('selecionado', width=30, anchor='center')
        self.tree.column('codigo', width=80)
        self.tree.column('produto', width=300)
        self.tree.column('departamento', width=150)
        self.tree.column('embalagem', width=100)
        self.tree.column('estoque_dicon', width=80)
        self.tree.column('estoque_vilog', width=80)
        self.tree.column('total', width=80)
        
        # Bind para clique na coluna de seleção
        self.tree.bind('<Button-1>', self.on_click)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(lista_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Frame de controles
        controles_frame = ttk.Frame(self.main_frame)
        controles_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        # Botões
        ttk.Button(controles_frame, text="Selecionar Todos", command=self.selecionar_todos).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(controles_frame, text="Limpar Seleção", command=self.limpar_selecao).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(controles_frame, text="Gerar PDF", command=self.gerar_pdf).pack(side=tk.LEFT, padx=(0, 10))
        
        # Status
        self.status_label = ttk.Label(self.main_frame, text="Pronto para buscar produtos")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        # Inicializar busca por departamento
        self.atualizar_busca()
    
    def atualizar_busca(self):
        """Atualiza a interface baseada no tipo de busca selecionado"""
        # Destruir campo atual se existir
        if hasattr(self, 'campo_busca'):
            self.campo_busca.destroy()
        
        if self.tipo_busca.get() == "departamento":
            # Mostrar combobox com departamentos
            departamentos = sorted(self.df_produtos['departamento'].unique().tolist())
            self.campo_busca = ttk.Combobox(self.busca_frame, values=departamentos, width=40)
            self.campo_busca.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
            self.campo_busca.bind('<<ComboboxSelected>>', self.filtrar_produtos)
            
        else:
            # Mostrar entry para busca por nome
            self.campo_busca = ttk.Entry(self.busca_frame, width=40)
            self.campo_busca.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
            self.campo_busca.bind('<KeyRelease>', self.filtrar_produtos)
    
    def filtrar_produtos(self, event=None):
        """Filtra produtos baseado na busca"""
        try:
            termo = self.campo_busca.get().strip()
            if not termo:
                self.limpar_lista()
                return
            
            if self.tipo_busca.get() == "departamento":
                # Busca por departamento
                self.produtos_filtrados = self.df_produtos[self.df_produtos['departamento'] == termo].copy()
            else:
                # Busca por nome
                self.produtos_filtrados = self.df_produtos[self.df_produtos['descricao'].str.upper().str.contains(termo.upper(), na=False)].copy()
            
            self.atualizar_lista()
            self.status_label.config(text=f"Encontrados {len(self.produtos_filtrados)} produtos")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao filtrar produtos: {e}")
    
    def limpar_lista(self):
        """Limpa a lista de produtos"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.produtos_filtrados = None
        self.status_label.config(text="Pronto para buscar produtos")
    
    def atualizar_lista(self):
        """Atualiza a lista de produtos na interface"""
        # Limpar lista atual
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.produtos_filtrados is None or self.produtos_filtrados.empty:
            return
        
        # Adicionar produtos filtrados
        for _, produto in self.produtos_filtrados.iterrows():
            codprod = str(produto['codprod']).strip()
            
            # Calcular estoques
            qtunitcx = int(float(produto.get('qtunitcx') or 0))
            estoque_sql_units = int(produto['disponivel_un']) if pd.notna(produto['disponivel_un']) else 0
            estoque_sql_caixas = int(round(estoque_sql_units / qtunitcx)) if qtunitcx > 0 else 0
            estoque_vilog_caixas = self.estoque_vilog.get(codprod, 0)
            estoque_total = estoque_sql_caixas + estoque_vilog_caixas
            
            self.tree.insert('', 'end', values=(
                '',  # Coluna de seleção (vazia inicialmente)
                codprod,  # Usar o código já convertido para string
                produto['descricao'][:50] + "..." if len(produto['descricao']) > 50 else produto['descricao'],
                produto['departamento'],
                produto['embalagem'] or '',
                estoque_sql_caixas,
                estoque_vilog_caixas,
                estoque_total
            ))
    
    def on_click(self, event):
        """Gerencia cliques na coluna de seleção"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == '#1':  # Coluna de seleção
                item = self.tree.identify_row(event.y)
                if item:
                    valores = list(self.tree.item(item, 'values'))
                    if valores[0] == '':  # Não selecionado
                        valores[0] = '✓'
                        self.tree.selection_add(item)
                    else:  # Já selecionado
                        valores[0] = ''
                        self.tree.selection_remove(item)
                    self.tree.item(item, values=valores)
    
    def selecionar_todos(self):
        """Seleciona todos os produtos da lista"""
        for item in self.tree.get_children():
            valores = list(self.tree.item(item, 'values'))
            valores[0] = '✓'
            self.tree.item(item, values=valores)
            self.tree.selection_add(item)
    
    def limpar_selecao(self):
        """Limpa a seleção atual"""
        for item in self.tree.get_children():
            valores = list(self.tree.item(item, 'values'))
            valores[0] = ''
            self.tree.item(item, values=valores)
            self.tree.selection_remove(item)
    
    def gerar_pdf(self):
        """Gera o PDF com os produtos selecionados"""
        try:
            # Obter produtos selecionados (baseado na coluna de checkbox)
            produtos_selecionados = []
            for item in self.tree.get_children():
                valores = self.tree.item(item, 'values')
                if valores[0] == '✓':  # Se tem checkbox marcado
                    produtos_selecionados.append(str(valores[1]).strip())  # Adiciona o código do produto como string
            
            if not produtos_selecionados:
                messagebox.showwarning("Aviso", "Selecione pelo menos um produto!")
                return
            
            # Converter códigos do DataFrame para string para comparação
            codigos_df = self.produtos_filtrados['codprod'].astype(str).str.strip()
            produtos_pdf = self.produtos_filtrados[codigos_df.isin(produtos_selecionados)].copy()
            
            if produtos_pdf.empty:
                # Debug: mostrar informações sobre o problema
                print(f"Produtos selecionados: {produtos_selecionados}")
                print(f"Tipos dos produtos selecionados: {[type(p) for p in produtos_selecionados]}")
                print(f"Primeiros códigos do DataFrame: {codigos_df.head().tolist()}")
                print(f"Tipos dos códigos do DataFrame: {[type(c) for c in codigos_df.head().tolist()]}")
                messagebox.showerror("Erro", "Nenhum produto válido selecionado!")
                return
            
            # Gerar PDF
            self.criar_pdf(produtos_pdf)
            
            messagebox.showinfo("Sucesso", "PDF gerado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF: {e}")
    
    def criar_pdf(self, produtos):
        """Cria o PDF com os produtos selecionados"""
        # Processar dados como no script original
        meses_ref = get_tres_meses_anteriores()
        meses_ano = [m[0] for m in meses_ref]
        meses_nome = [m[1] for m in meses_ref]
        meses_nome_compra = [f"{m} (compra)" for m in meses_nome]
        
        # Processar compras
        self.df_compras['yyyymm'] = self.df_compras['mes_ref'].dt.strftime('%Y%m')
        compras_grouped = self.df_compras.groupby(['codprod', 'yyyymm'], as_index=False)['qt_comprada'].sum()
        
        compras_dict = {}
        for _, row in compras_grouped.iterrows():
            codprod = str(row['codprod']).strip()
            mes = row['yyyymm']
            if mes in meses_ano:
                mes_idx = meses_ano.index(mes)
                mes_nome = meses_nome[mes_idx]
                if codprod not in compras_dict:
                    compras_dict[codprod] = {mn: 0 for mn in meses_nome}
                compras_dict[codprod][mes_nome] = int(row['qt_comprada'])
        
        # Preparar dados para PDF
        rows_temp = []
        for _, p in produtos.iterrows():
            codprod = str(p['codprod']).strip()
            nome = p['descricao']
            embalagem = p['embalagem'] or ''
            qtunitcx = int(float(p.get('qtunitcx') or 0))
            custo = format_real(p.get('custo_final_unitario', 0))
            
            # Compras
            compras_meses_units = compras_dict.get(codprod, {mn: 0 for mn in meses_nome})
            compras_meses_caixas = []
            compras_soma = 0
            for mn in meses_nome:
                qtqtd = compras_meses_units[mn]
                if qtunitcx > 0:
                    caixas = int(round(qtqtd / qtunitcx))
                else:
                    caixas = 0
                compras_meses_caixas.append(caixas)
                compras_soma += caixas
            
            # Estoques
            estoque_sql_units = int(p['disponivel_un']) if pd.notna(p['disponivel_un']) else 0
            estoque_sql_caixas = int(round(estoque_sql_units / qtunitcx)) if qtunitcx > 0 else 0
            qtd_vilog = self.estoque_vilog.get(codprod, 0)
            estoque_total = estoque_sql_caixas + qtd_vilog
            
            rows_temp.append([
                compras_soma,
                estoque_total,
                nome.lower(),
                codprod,
                nome,
                embalagem,
                compras_meses_caixas[0],
                compras_meses_caixas[1],
                compras_meses_caixas[2],
                estoque_sql_caixas,
                qtd_vilog,
                estoque_total,
                custo
            ])
        
        rows_sorted = sorted(rows_temp, key=lambda x: (-x[0], -x[1], x[2]))
        
        data = [
            ['Código', 'Produto', 'Embalagem'] + meses_nome_compra + ['Qtd DICON', 'Qtd VILOG', 'Total', 'Custo']
        ] + [row[3:] for row in rows_sorted]
        
        # Configurar PDF
        colwidths = [
            20 * mm, 85 * mm, 20 * mm, 25 * mm, 25 * mm,
            25 * mm, 20 * mm, 20 * mm, 20 * mm, 20 * mm,
        ]
        
        # Nome do arquivo
        nome_arquivo = f'produtos_selecionados_{datetime.today().strftime("%Y%m%d_%H%M")}.pdf'
        pdf_path = Path(nome_arquivo)
        
        doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(A4), 
                              rightMargin=8 * mm, leftMargin=8 * mm,
                              topMargin=10 * mm, bottomMargin=10 * mm)
        
        styleSheet = getSampleStyleSheet()
        titulo = Paragraph(
            f"<b>Relatório de Falta de Produtos</b>"
            f"<br/><font size=10>Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</font>",
            styleSheet['Title']
        )
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#124377")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.15, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('PADDING', (0, 0), (-1, -1), 3),
            ('WORDWRAP', (0, 0), (-1, -1), True),
        ])
        
        table = Table(data, repeatRows=1, colWidths=colwidths)
        table.setStyle(style)
        
        elements = [titulo, Spacer(1, 6 * mm), table]
        doc.build(elements)
        
        # Abrir o PDF
        import os
        os.startfile(str(pdf_path.resolve()))


def main():
    root = tk.Tk()
    root.title("Sistema de Relatórios de Produtos")
    root.geometry("1200x800")
    app = FaltaProdutosInterface(root)
    root.mainloop()


if __name__ == "__main__":
    main()
