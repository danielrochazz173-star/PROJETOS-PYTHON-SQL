"""
Sistema de Relatório de Vendedores - Excel
Gera Excel com RCA, Nome e colunas organizadas: FATs, Margens, MIX, Pesos
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime
from calendar import monthrange
from pathlib import Path
import os
import threading
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

def buscar_vendas_vendedores_mes(conn, mes, ano):
    """Busca vendas por vendedor no mês especificado
    Usa a mesma lógica do compraxvenda_totais.py
    """
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    WITH PRODUTOS_VENDIDOS AS (
        SELECT DISTINCT
            C.CODUSUR,
            I.CODPROD
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 6, 8, 10, 11, 12, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
    ),
    VENDAS_VALIDAS AS (
        SELECT 
            C.CODUSUR,
            U.NOME AS NOME_VENDEDOR,
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
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY C.CODUSUR, U.NOME
    ),
    MIX_POR_VENDEDOR AS (
        SELECT 
            CODUSUR,
            COUNT(*) AS MIX_VENDIDO
        FROM PRODUTOS_VENDIDOS
        GROUP BY CODUSUR
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODUSUR,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
        GROUP BY D.CODUSUR
    )
    SELECT 
        V.CODUSUR AS RCA,
        V.NOME_VENDEDOR,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS MARGEM,
        NVL(M.MIX_VENDIDO, 0) AS MIX_VENDIDO,
        NVL(V.PESO_TOTAL, 0) AS PESO
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODUSUR = D.CODUSUR
    LEFT JOIN MIX_POR_VENDEDOR M ON V.CODUSUR = M.CODUSUR
    WHERE V.VALOR_BRUTO > 0
    ORDER BY FATURAMENTO DESC
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar vendas: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# FUNÇÕES DE EXCEL
# =========================================================================
def formatar_excel(caminho_arquivo, meses_selecionados):
    """Formata o arquivo Excel com estilos"""
    wb = openpyxl.load_workbook(caminho_arquivo)
    ws = wb.active
    
    # Estilos
    header_fill = PatternFill(start_color="124377", end_color="124377", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # PRIMEIRO: Remover QUALQUER preenchimento de TODAS as linhas de dados
    # Isso garante que não fique nenhum azul nas linhas de dados
    no_fill = PatternFill()  # Sem preenchimento (transparente)
    for row_num in range(3, ws.max_row + 1):
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            # FORÇAR remoção de qualquer preenchimento
            cell.fill = no_fill
    
    # Formatar cabeçalho (linha 2 após título) - APENAS linha 2
    linha_cabecalho = 2
    for cell in ws[linha_cabecalho]:
        cell.fill = header_fill  # Azul apenas aqui
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_style
    
    # Agora formatar as células de dados (a partir da linha 3)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            # FORÇAR sem preenchimento novamente (garantir que não ficou azul)
            cell.fill = no_fill
            cell.border = border_style
            if cell.column <= 2:  # RCA e Nome
                if cell.column == 1:  # RCA
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:  # Nome
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            else:  # Colunas de dados
                cell.alignment = Alignment(horizontal='right', vertical='center')
                # Nova ordem: RCA(1), Nome(2), FATs(3 até 2+N), Margens(3+N até 2+2N), MIX(3+2N até 2+3N), Pesos(3+3N até 2+4N)
                num_meses = len(meses_selecionados)
                col_relativa = cell.column - 2  # Coluna relativa (1 = primeira coluna de dados)
                
                if col_relativa <= num_meses:  # FATs
                    cell.number_format = '#,##0.00'
                elif col_relativa <= num_meses * 2:  # Margens
                    cell.number_format = '0.00%'
                elif col_relativa <= num_meses * 3:  # MIX
                    cell.number_format = '#,##0'
                else:  # Pesos
                    cell.number_format = '#,##0.00'
    
    # Ajustar largura das colunas
    ws.column_dimensions['A'].width = 10  # RCA
    ws.column_dimensions['B'].width = 30  # Nome Vendedor
    
    # Ajustar largura das colunas
    num_meses = len(meses_selecionados)
    # FATs (colunas 3 até 2+num_meses)
    for i in range(num_meses):
        col_letter = get_column_letter(3 + i)
        ws.column_dimensions[col_letter].width = 15
    # Margens (colunas 3+num_meses até 2+2*num_meses)
    for i in range(num_meses):
        col_letter = get_column_letter(3 + num_meses + i)
        ws.column_dimensions[col_letter].width = 15
    # MIX (colunas 3+2*num_meses até 2+3*num_meses)
    for i in range(num_meses):
        col_letter = get_column_letter(3 + 2 * num_meses + i)
        ws.column_dimensions[col_letter].width = 12
    # Pesos (colunas 3+3*num_meses até 2+4*num_meses)
    for i in range(num_meses):
        col_letter = get_column_letter(3 + 3 * num_meses + i)
        ws.column_dimensions[col_letter].width = 15
    
    # Adicionar título
    ws.insert_rows(1)
    num_cols = 2 + len(meses_selecionados) * 4  # RCA, Nome, FATs, Margens, MIX, Pesos
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    titulo_cell = ws['A1']
    
    meses_nomes = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    
    meses_texto = ', '.join([f"{meses_nomes[mes]}/{str(ano)[-2:]}" for mes, ano in meses_selecionados])
    titulo_cell.value = f"Relatório de Vendedores - {meses_texto}"
    titulo_cell.font = Font(bold=True, size=14, color="124377")
    titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # LIMPEZA FINAL: Garantir que nenhuma linha de dados tenha azul
    no_fill = PatternFill()
    for row_num in range(3, ws.max_row + 1):
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            # Se tiver qualquer preenchimento azul, remover
            if cell.fill and hasattr(cell.fill, 'start_color'):
                if hasattr(cell.fill.start_color, 'rgb') and cell.fill.start_color.rgb:
                    if '124377' in str(cell.fill.start_color.rgb) or cell.fill.start_color.rgb == 'FF124377':
                        cell.fill = no_fill
                elif cell.fill.fill_type == 'solid' and cell.fill.start_color:
                    cell.fill = no_fill
    
    wb.save(caminho_arquivo)

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class RelatorioVendedoresExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Relatório de Vendedores - Excel")
        self.root.geometry("500x500")
        
        # Variáveis
        self.conn = None
        
        # Criar interface
        self.criar_interface()
        
        # Conectar ao banco
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Relatório de Vendedores - Excel", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de meses
        frame_meses = ttk.LabelFrame(main_frame, text="Selecionar Meses", padding="15")
        frame_meses.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Scrollable frame para meses
        canvas = tk.Canvas(frame_meses, height=250)
        scrollbar = ttk.Scrollbar(frame_meses, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Variáveis para checkboxes
        self.meses_vars = {}
        
        # Gerar últimos 12 meses
        hoje = datetime.now()
        meses_lista = []
        meses_nomes = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        
        for i in range(12):
            # Calcular mês e ano
            mes_atual = hoje.month - i
            ano_atual = hoje.year
            
            while mes_atual <= 0:
                mes_atual += 12
                ano_atual -= 1
            
            mes_nome = f"{meses_nomes[mes_atual]}/{str(ano_atual)[-2:]}"
            meses_lista.append((mes_atual, ano_atual, mes_nome))
        
        # Criar checkboxes
        for mes, ano, mes_nome in meses_lista:
            var = tk.BooleanVar()
            self.meses_vars[(mes, ano)] = var
            cb = ttk.Checkbutton(scrollable_frame, text=mes_nome, variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões de seleção rápida
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill=tk.X, pady=10)
        
        ttk.Button(frame_botoes, text="Selecionar Todos", 
                  command=self.selecionar_todos).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Limpar Seleção", 
                  command=self.limpar_selecao).pack(side=tk.LEFT, padx=5)
        
        # Botão gerar Excel
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        self.btn_gerar = ttk.Button(btn_frame, text="📄 Gerar Excel", 
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    def selecionar_todos(self):
        """Seleciona todos os meses"""
        for var in self.meses_vars.values():
            var.set(True)
    
    def limpar_selecao(self):
        """Limpa seleção de meses"""
        for var in self.meses_vars.values():
            var.set(False)
    
    def conectar_banco(self):
        """Conecta ao banco"""
        def conectar():
            try:
                self.status_label.config(text="Conectando ao banco...")
                self.conn = get_db_connection()
                
                if not self.conn:
                    self.status_label.config(text="❌ Erro ao conectar")
                    messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")
                    return
                
                self.status_label.config(text="✅ Pronto para gerar relatório")
                self.btn_gerar.config(state=tk.NORMAL)
                
            except Exception as e:
                self.status_label.config(text=f"❌ Erro: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao conectar: {e}")
        
        threading.Thread(target=conectar, daemon=True).start()
    
    def gerar_excel(self):
        """Gera o Excel com vendas dos vendedores"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        # Obter meses selecionados
        meses_selecionados = [(mes, ano) for (mes, ano), var in self.meses_vars.items() if var.get()]
        
        if not meses_selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um mês!")
            return
        
        # Ordenar meses por data
        meses_selecionados.sort(key=lambda x: (x[1], x[0]))
        
        def executar():
            try:
                self.status_label.config(text="Buscando dados...")
                self.root.update()
                
                # Buscar dados de todos os vendedores primeiro (para ter lista completa)
                todos_vendedores = {}
                
                # Buscar vendas de cada mês
                dados_por_mes = {}
                for mes, ano in meses_selecionados:
                    df_mes = buscar_vendas_vendedores_mes(self.conn, mes, ano)
                    dados_por_mes[(mes, ano)] = df_mes
                    
                    # Adicionar vendedores à lista completa
                    if not df_mes.empty:
                        for _, row in df_mes.iterrows():
                            rca = row['rca']
                            nome = row['nome_vendedor']
                            if rca not in todos_vendedores:
                                todos_vendedores[rca] = nome
                
                if not todos_vendedores:
                    messagebox.showinfo("Informação", "Nenhuma venda encontrada para os meses selecionados.")
                    self.status_label.config(text="Nenhuma venda encontrada")
                    return
                
                self.status_label.config(text="Gerando Excel...")
                self.root.update()
                
                # Criar DataFrame para Excel
                df_excel = pd.DataFrame({
                    'RCA': list(todos_vendedores.keys()),
                    'Nome do Vendedor': list(todos_vendedores.values())
                })
                
                # Adicionar colunas para cada mês
                meses_nomes = {
                    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
                    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
                    9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
                }
                
                # Lista para calcular faturamento acumulado
                colunas_fat = []
                
                for mes, ano in meses_selecionados:
                    mes_abrev = meses_nomes[mes]
                    ano_abrev = str(ano)[-2:]
                    
                    df_mes = dados_por_mes[(mes, ano)]
                    
                    # Criar dicionário para lookup rápido
                    dict_mes = {}
                    if not df_mes.empty:
                        for _, row in df_mes.iterrows():
                            fat = row['faturamento']
                            margem_valor = row['margem']
                            # Calcular margem percentual
                            margem_percentual = (margem_valor / fat * 100) if fat > 0 else 0
                            dict_mes[row['rca']] = {
                                'faturamento': fat,
                                'margem_percentual': margem_percentual,
                                'mix_vendido': row['mix_vendido'],
                                'peso': row['peso']
                            }
                    
                    # Adicionar colunas (vamos adicionar todas e depois reorganizar)
                    col_fat = f'FAT {mes_abrev}/{ano_abrev}'
                    colunas_fat.append(col_fat)
                    df_excel[col_fat] = df_excel['RCA'].map(
                        lambda x: dict_mes.get(x, {}).get('faturamento', 0)
                    )
                    # Margem como percentual (dividido por 100 para formato decimal do Excel)
                    df_excel[f'Margem {mes_abrev}/{ano_abrev}'] = df_excel['RCA'].map(
                        lambda x: dict_mes.get(x, {}).get('margem_percentual', 0) / 100
                    )
                    df_excel[f'MIX {mes_abrev}/{ano_abrev}'] = df_excel['RCA'].map(
                        lambda x: dict_mes.get(x, {}).get('mix_vendido', 0)
                    )
                    df_excel[f'Peso {mes_abrev}/{ano_abrev}'] = df_excel['RCA'].map(
                        lambda x: dict_mes.get(x, {}).get('peso', 0)
                    )
                
                # Calcular faturamento acumulado
                df_excel['FAT_ACUMULADO'] = df_excel[colunas_fat].sum(axis=1)
                
                # Ordenar por faturamento acumulado (decrescente)
                df_excel = df_excel.sort_values('FAT_ACUMULADO', ascending=False)
                
                # Remover coluna auxiliar
                df_excel = df_excel.drop(columns=['FAT_ACUMULADO'])
                
                # Reorganizar colunas: RCA, Nome, depois todos FATs, depois todas Margens, depois todos MIX, depois todos Pesos
                colunas_ordenadas = ['RCA', 'Nome do Vendedor']
                
                # Adicionar todos os FATs
                for mes, ano in meses_selecionados:
                    mes_abrev = meses_nomes[mes]
                    ano_abrev = str(ano)[-2:]
                    colunas_ordenadas.append(f'FAT {mes_abrev}/{ano_abrev}')
                
                # Adicionar todas as Margens
                for mes, ano in meses_selecionados:
                    mes_abrev = meses_nomes[mes]
                    ano_abrev = str(ano)[-2:]
                    colunas_ordenadas.append(f'Margem {mes_abrev}/{ano_abrev}')
                
                # Adicionar todos os MIX
                for mes, ano in meses_selecionados:
                    mes_abrev = meses_nomes[mes]
                    ano_abrev = str(ano)[-2:]
                    colunas_ordenadas.append(f'MIX {mes_abrev}/{ano_abrev}')
                
                # Adicionar todos os Pesos
                for mes, ano in meses_selecionados:
                    mes_abrev = meses_nomes[mes]
                    ano_abrev = str(ano)[-2:]
                    colunas_ordenadas.append(f'Peso {mes_abrev}/{ano_abrev}')
                
                # Reordenar DataFrame
                df_excel = df_excel[colunas_ordenadas]
                
                # Nome do arquivo
                meses_str = '_'.join([f"{mes:02d}{str(ano)[-2:]}" for mes, ano in meses_selecionados])
                nome_arquivo = f"relatorio_vendedores_{meses_str}_{datetime.now().strftime('%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                # Salvar Excel
                df_excel.to_excel(caminho_arquivo, index=False, sheet_name='Vendedores')
                
                # Formatar Excel
                formatar_excel(caminho_arquivo, meses_selecionados)
                
                self.status_label.config(text="✅ Excel gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"Excel gerado com sucesso!\n\nArquivo: {caminho_arquivo}\n\nTotal de vendedores: {len(df_excel)}")
                
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

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = RelatorioVendedoresExcel(root)
    root.mainloop()
