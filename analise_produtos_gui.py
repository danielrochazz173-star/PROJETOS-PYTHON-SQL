"""
Sistema Tkinter para análise de produtos.
- Calcula faturamento líquido (venda - devolução), margem, positivação,
  quantidade de caixas e peso vendido.
- Exporta para Excel e PDF paginado (WeasyPrint).
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook
import threading
import os

try:
    from tkcalendar import DateEntry
except ImportError:
    messagebox.showerror(
        "Erro",
        "Biblioteca tkcalendar não encontrada!\n\nExecute: pip install tkcalendar"
    )
    DateEntry = None

# WeasyPrint é opcional: avisamos no momento da geração do PDF
# Tentamos WeasyPrint; se não estiver disponível, caímos para ReportLab (puro Python).
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
DB_HOST = os.getenv('DB_HOST', '10.0.0.10')
DB_PORT = int(os.getenv('DB_PORT', '1521'))
DB_SERVICE = os.getenv('DB_SERVICE', 'PROD')
DB_USER = os.getenv('DB_USER', 'powerbi')
DB_PASSWORD = os.getenv('DB_PASSWORD', '')


# =========================================================================
# FUNÇÕES DE BANCO
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle."""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError:
        return None


def buscar_departamentos(conn):
    """Busca lista de departamentos."""
    try:
        query = "SELECT CODEPTO, DESCRICAO FROM PCDEPTO ORDER BY DESCRICAO"
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def buscar_produtos_periodo(conn, data_inicio, data_fim, coddepto=None, codprod_list=None):
    """
    Busca dados consolidados por produto no período.

    Retorna DataFrame com:
    - codprod, descricao, departamento
    - faturamento (venda líquida)
    - custo líquido
    - positivação (clientes únicos)
    - qtd_caixas, peso_vendido
    """
    data_fim_inclusiva = data_fim + timedelta(days=1)
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim_inclusiva.strftime('%d/%m/%Y')

    filtro_depto = ""
    if coddepto:
        filtro_depto = f"AND P.CODEPTO = {coddepto}"

    # Filtros de produtos (vendas)
    filtro_produto = ""
    filtro_produto_devol = ""
    if codprod_list:
        chunks = [codprod_list[i:i + 1000] for i in range(0, len(codprod_list), 1000)]
        cond_vendas = []
        cond_devol = []
        for chunk in chunks:
            cond_vendas.append(f"I.CODPROD IN ({', '.join(map(str, chunk))})")
            cond_devol.append(f"D.CODPROD IN ({', '.join(map(str, chunk))})")
        filtro_produto = f"AND ({' OR '.join(cond_vendas)})"
        filtro_produto_devol = f"AND ({' OR '.join(cond_devol)})"

    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            I.CODPROD,
            P.DESCRICAO,
            P.CODEPTO,
            D.DESCRICAO AS NOME_DEPTO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                        )
                    ELSE 0 
                END
            ) AS VALOR_BRUTO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0)
                        )
                    ELSE 0 
                END
            ) AS CUSTO_BRUTO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
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
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(P.PESOBRUTO, 0)
                        )
                    ELSE 0 
                END
            ) AS PESO_TOTAL,
            COUNT(DISTINCT C.CODCLI) AS POSITIVACAO
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          {filtro_depto}
          {filtro_produto}
        GROUP BY I.CODPROD, P.DESCRICAO, P.CODEPTO, D.DESCRICAO
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          {filtro_depto}
          {filtro_produto_devol}
        GROUP BY D.CODPROD
    )
    SELECT 
        V.CODPROD,
        V.DESCRICAO,
        NVL(V.NOME_DEPTO, 'Sem Depto') AS DEPARTAMENTO,
        NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0) AS FATURAMENTO,
        NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0) AS CUSTO_LIQUIDO,
        NVL(V.POSITIVACAO, 0) AS POSITIVACAO,
        NVL(V.QTD_CAIXAS, 0) AS QTD_CAIXAS,
        NVL(V.PESO_TOTAL, 0) AS PESO_VENDIDO
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODPROD = D.CODPROD
    WHERE NVL(V.VALOR_BRUTO, 0) > 0 OR NVL(D.VALOR_DEVOLVIDO, 0) > 0
    ORDER BY FATURAMENTO DESC
    """

    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]

    if df.empty:
        return df

    df['margem_valor'] = df['faturamento'] - df['custo_liquido']
    df['margem_percentual'] = df.apply(
        lambda row: 0.0 if row['faturamento'] <= 0 else round((row['margem_valor'] / row['faturamento']) * 100, 2),
        axis=1
    )

    return df


# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class AnaliseProdutosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 Análise de Produtos")
        self.root.geometry("560x700")

        self.conn = None
        self.df_departamentos = None

        self.criar_interface()
        self.conectar_banco()

    def criar_interface(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        titulo = ttk.Label(main_frame, text="📦 Análise de Produtos", font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))

        frame_periodo = ttk.LabelFrame(main_frame, text="Período de Análise", padding="15")
        frame_periodo.pack(fill=tk.X, pady=10)

        if DateEntry is None:
            ttk.Label(
                frame_periodo,
                text="Erro: tkcalendar não instalado. Execute: pip install tkcalendar",
                foreground='red'
            ).pack()
            return

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

        frame_filtros = ttk.LabelFrame(main_frame, text="Filtros", padding="15")
        frame_filtros.pack(fill=tk.X, pady=10)

        frame_depto_sel = ttk.Frame(frame_filtros)
        frame_depto_sel.pack(fill=tk.X, pady=5)
        ttk.Label(frame_depto_sel, text="Departamento:", width=15).pack(side=tk.LEFT)
        self.combo_depto = ttk.Combobox(frame_depto_sel, width=40, state="readonly")
        self.combo_depto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        frame_prod_sel = ttk.Frame(frame_filtros)
        frame_prod_sel.pack(fill=tk.X, pady=5)
        ttk.Label(frame_prod_sel, text="Produtos (vírgula):", width=15).pack(side=tk.LEFT)
        self.entry_produto = ttk.Entry(frame_prod_sel, width=50)
        self.entry_produto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(frame_prod_sel, text="(deixe vazio para todos)").pack(side=tk.LEFT, padx=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))

        self.btn_excel = ttk.Button(btn_frame, text="📊 Gerar Excel",
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_excel.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pdf = ttk.Button(btn_frame, text="🖨️ Gerar PDF",
                                 command=self.gerar_pdf, state=tk.DISABLED)
        self.btn_pdf.pack(side=tk.LEFT, padx=(0, 10))

        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)

    def conectar_banco(self):
        """Conecta ao banco e carrega departamentos em thread."""
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
                    deptos = ["Todos"] + [
                        f"{int(row['codepto'])} - {row['descricao']}"
                        for _, row in self.df_departamentos.iterrows()
                    ]
                    self.combo_depto['values'] = deptos
                    self.combo_depto.current(0)
                else:
                    self.combo_depto['values'] = ["Todos"]
                    self.combo_depto.current(0)

                self.status_label.config(text="✅ Pronto para gerar relatórios")
                self.btn_excel.config(state=tk.NORMAL)
                self.btn_pdf.config(state=tk.NORMAL)

            except Exception as exc:
                self.status_label.config(text=f"❌ Erro: {str(exc)}")
                messagebox.showerror("Erro", f"Erro ao conectar: {exc}")

        threading.Thread(target=carregar, daemon=True).start()

    def _obter_filtros(self):
        """Lê filtros da interface e valida datas."""
        data_inicio = self.date_inicio.get_date()
        data_fim = self.date_fim.get_date()
        if data_fim < data_inicio:
            messagebox.showwarning("Aviso", "Data fim deve ser maior ou igual à data início!")
            return None

        depto_sel = self.combo_depto.get()
        coddepto = None
        if depto_sel and depto_sel != "Todos":
            coddepto = int(depto_sel.split(" - ")[0])

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

        return data_inicio, data_fim, coddepto, codprod_list

    def gerar_excel(self):
        """Gera Excel em thread."""
        filtros = self._obter_filtros()
        if not filtros:
            return
        data_inicio, data_fim, coddepto, codprod_list = filtros

        def executar():
            try:
                self.status_label.config(text="Buscando dados...")
                self.root.update()

                df = buscar_produtos_periodo(self.conn, data_inicio, data_fim, coddepto, codprod_list)
                if df.empty:
                    self.status_label.config(text="❌ Nenhum dado")
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return

                df_excel = pd.DataFrame({
                    'Cód Produto': df['codprod'],
                    'Descrição': df['descricao'],
                    'Departamento': df['departamento'],
                    'Faturamento (R$)': df['faturamento'],
                    'Margem (R$)': df['margem_valor'],
                    'Margem (%)': df['margem_percentual'] / 100,
                    'Positivação': df['positivacao'],
                    'Qtd Caixas': df['qtd_caixas'],
                    'Peso Vendido (kg)': df['peso_vendido']
                }).sort_values('Faturamento (R$)', ascending=False)

                data_inicio_str = data_inicio.strftime('%Y%m%d')
                data_fim_str = data_fim.strftime('%Y%m%d')
                nome_arquivo = f"analise_produtos_{data_inicio_str}_{data_fim_str}_{datetime.now().strftime('%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()

                self.status_label.config(text="Gerando Excel...")
                self.root.update()
                with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, sheet_name='Produtos', index=False)

                periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
                self.status_label.config(text="Formatando Excel...")
                self.root.update()
                self.formatar_excel(caminho_arquivo, periodo_str, depto_sel=depto_sel if coddepto else "Todos")

                self.status_label.config(text="✅ Excel gerado")
                messagebox.showinfo("Sucesso", f"Excel gerado com sucesso!\n\nArquivo: {caminho_arquivo}")
                try:
                    import os
                    os.startfile(str(caminho_arquivo))
                except Exception:
                    pass
            except Exception as exc:
                self.status_label.config(text="❌ Erro ao gerar Excel")
                messagebox.showerror("Erro", f"Erro ao gerar Excel:\n{exc}")

        data_inicio, data_fim, coddepto, codprod_list = filtros
        depto_sel = self.combo_depto.get()
        threading.Thread(target=executar, daemon=True).start()

    def gerar_pdf(self):
        """Gera PDF via WeasyPrint em thread."""
        if PDF_ENGINE is None:
            messagebox.showerror(
                "Dependência faltando",
                "Nenhum gerador de PDF disponível.\n\nOpções:\n- pip install weasyprint (requer GTK/Pango/Cairo)\n- ou: pip install reportlab (puro Python)"
            )
            return

        filtros = self._obter_filtros()
        if not filtros:
            return
        data_inicio, data_fim, coddepto, codprod_list = filtros
        depto_sel = self.combo_depto.get()

        def executar():
            try:
                self.status_label.config(text="Buscando dados...")
                self.root.update()

                df = buscar_produtos_periodo(self.conn, data_inicio, data_fim, coddepto, codprod_list)
                if df.empty:
                    self.status_label.config(text="❌ Nenhum dado")
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return

                df_pdf = pd.DataFrame({
                    'codprod': df['codprod'],
                    'descricao': df['descricao'],
                    'departamento': df['departamento'],
                    'faturamento': df['faturamento'],
                    'margem_valor': df['margem_valor'],
                    'margem_percentual': df['margem_percentual'],
                    'positivacao': df['positivacao'],
                    'qtd_caixas': df['qtd_caixas'],
                    'peso_vendido': df['peso_vendido']
                }).sort_values('faturamento', ascending=False)

                data_inicio_str = data_inicio.strftime('%Y%m%d')
                data_fim_str = data_fim.strftime('%Y%m%d')
                nome_pdf = f"analise_produtos_{data_inicio_str}_{data_fim_str}_{datetime.now().strftime('%H%M%S')}.pdf"
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
                    import os
                    os.startfile(str(caminho_pdf))
                except Exception:
                    pass
            except Exception as exc:
                self.status_label.config(text="❌ Erro ao gerar PDF")
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{exc}")

        threading.Thread(target=executar, daemon=True).start()

    def formatar_excel(self, arquivo, periodo_str, depto_sel="Todos"):
        """Formata a planilha de produtos."""
        wb = load_workbook(arquivo)
        ws = wb['Produtos']

        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        col_widths = {
            'A': 12,   # Cód Produto
            'B': 40,   # Descrição
            'C': 25,   # Departamento
            'D': 18,   # Faturamento
            'E': 15,   # Margem R$
            'F': 12,   # Margem %
            'G': 12,   # Positivação
            'H': 14,   # Qtd Caixas
            'I': 16    # Peso Vendido
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                col_letter = cell.column_letter
                if col_letter == 'A':
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col_letter in ['B', 'C']:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                elif col_letter in ['D', 'E', 'H', 'I']:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'F':
                    cell.number_format = '0.00%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_letter == 'G':
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')

        ws.insert_rows(1)
        ws.merge_cells('A1:I1')
        titulo_cell = ws['A1']
        titulo_texto = f"Análise de Produtos - {periodo_str}"
        if depto_sel != "Todos":
            titulo_texto += f" - {depto_sel}"
        titulo_cell.value = titulo_texto
        titulo_cell.font = Font(bold=True, size=14, color="366092")
        titulo_cell.alignment = Alignment(horizontal='center', vertical='center')

        ws.freeze_panes = 'A3'
        wb.save(arquivo)

    def _css_pdf(self):
        """CSS para PDF paginado."""
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
        """Monta HTML para o PDF."""
        linhas = ""
        for _, row in df.iterrows():
            linhas += f"""
            <tr>
                <td>{int(row['codprod'])}</td>
                <td>{row['descricao']}</td>
                <td>{row['departamento']}</td>
                <td class="num">{row['faturamento']:,.2f}</td>
                <td class="num">{row['margem_valor']:,.2f}</td>
                <td class="num">{row['margem_percentual']:,.2f}%</td>
                <td class="num">{int(row['positivacao'])}</td>
                <td class="num">{row['qtd_caixas']:,.2f}</td>
                <td class="num">{row['peso_vendido']:,.2f}</td>
            </tr>
            """

        depto_texto = depto_sel if depto_sel != "Todos" else "Todos os departamentos"

        return f"""
        <html>
        <head><meta charset="utf-8"></head>
        <body>
            <h1>Análise de Produtos</h1>
            <div class="meta">
                <div><strong>Período:</strong> {periodo_str}</div>
                <div><strong>Departamento:</strong> {depto_texto}</div>
                <div class="small">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>Cód Produto</th>
                        <th>Descrição</th>
                        <th>Departamento</th>
                        <th>Faturamento (R$)</th>
                        <th>Margem (R$)</th>
                        <th>Margem (%)</th>
                        <th>Positivação</th>
                        <th>Qtd Caixas</th>
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

        titulo = "Análise de Produtos"
        depto_texto = depto_sel if depto_sel != "Todos" else "Todos os departamentos"
        meta = f"Período: {periodo_str} | Departamento: {depto_texto}"
        gerado = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"

        elementos.append(Paragraph(f"<para align='center'><b>{titulo}</b></para>", styles['Title']))
        elementos.append(Spacer(1, 6))
        elementos.append(Paragraph(f"<para align='center'>{meta}</para>", styles['Normal']))
        elementos.append(Paragraph(f"<para align='center'><font size=9 color='#666666'>{gerado}</font></para>", styles['Normal']))
        elementos.append(Spacer(1, 12))

        header = [
            "Cód Produto", "Descrição", "Departamento",
            "Faturamento (R$)", "Margem (R$)", "Margem (%)",
            "Positivação", "Qtd Caixas", "Peso Vendido (kg)"
        ]

        dados = [header]
        for _, row in df.iterrows():
            dados.append([
                int(row['codprod']),
                str(row['descricao']),
                str(row['departamento']),
                f"{row['faturamento']:,.2f}",
                f"{row['margem_valor']:,.2f}",
                f"{row['margem_percentual']:,.2f}%",
                int(row['positivacao']),
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
    app = AnaliseProdutosGUI(root)
    root.mainloop()

# portfolio-commit-ready: analise_produtos_gui
