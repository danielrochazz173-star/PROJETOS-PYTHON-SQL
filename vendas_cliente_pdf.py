"""
Sistema de Vendas por Cliente - Interface Tkinter
Gera PDF com vendas do cliente separadas por período (mês)
Mostra valor total, margem e quantidade por período
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime, timedelta
from calendar import monthrange
from pathlib import Path
import os
import threading
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

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

def buscar_dados_cliente(conn, codcli):
    """Busca dados do cliente"""
    try:
        query = f"""
        SELECT CODCLI, CLIENTE, CGCENT AS CNPJ_CPF
        FROM PCCLIENT
        WHERE CODCLI = {codcli}
        """
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df.iloc[0] if not df.empty else None
    except:
        return None

def buscar_vendas_cliente_periodo(conn, codcli, data_inicio, data_fim):
    """Busca vendas do cliente no período (faturamento - devoluções)"""
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            TO_CHAR(TRUNC(C.DATA, 'MM'), 'MM/YYYY') AS MES_REF,
            TRUNC(C.DATA, 'MM') AS MES_DATA,
            C.NUMNOTA AS NF,
            C.DATA,
            I.CODPROD,
            P.DESCRICAO AS PRODUTO,
            I.PVENDA AS PRECO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS QUANTIDADE,
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
            ) AS CUSTO_BRUTO
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.CODCLI = {codcli}
          AND C.DATA BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
            AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY 
            TO_CHAR(TRUNC(C.DATA, 'MM'), 'MM/YYYY'),
            TRUNC(C.DATA, 'MM'),
            C.NUMNOTA,
            C.DATA,
            I.CODPROD,
            P.DESCRICAO,
            I.PVENDA
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD,
            TRUNC(D.DTENT, 'MM') AS MES_DATA,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.CODCLI = {codcli}
          AND D.DTENT BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
            AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
        GROUP BY D.CODPROD, TRUNC(D.DTENT, 'MM')
    ),
    VENDAS_TOTAIS_MES AS (
        SELECT 
            I.CODPROD,
            TRUNC(C.DATA, 'MM') AS MES_DATA,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                        )
                    ELSE 0 
                END
            ) AS VALOR_BRUTO_TOTAL,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0)
                        )
                    ELSE 0 
                END
            ) AS CUSTO_BRUTO_TOTAL
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        WHERE C.CODCLI = {codcli}
          AND C.DATA BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
            AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY I.CODPROD, TRUNC(C.DATA, 'MM')
    )
    SELECT 
        V.MES_REF,
        V.MES_DATA,
        V.NF,
        V.DATA,
        V.CODPROD,
        V.PRODUTO,
        V.PRECO,
        V.QUANTIDADE,
        ROUND(
            CASE 
                WHEN NVL(VT.VALOR_BRUTO_TOTAL, 0) > 0 THEN
                    V.VALOR_BRUTO * ((NVL(VT.VALOR_BRUTO_TOTAL, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) / VT.VALOR_BRUTO_TOTAL)
                ELSE V.VALOR_BRUTO
            END, 2
        ) AS VALOR_TOTAL,
        ROUND(
            CASE 
                WHEN NVL(VT.CUSTO_BRUTO_TOTAL, 0) > 0 AND NVL(D.CUSTO_DEVOLVIDO, 0) > 0 THEN
                    V.CUSTO_BRUTO * ((NVL(VT.CUSTO_BRUTO_TOTAL, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) / VT.CUSTO_BRUTO_TOTAL)
                ELSE V.CUSTO_BRUTO
            END, 2
        ) AS CUSTO_TOTAL
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODPROD = D.CODPROD AND V.MES_DATA = D.MES_DATA
    LEFT JOIN VENDAS_TOTAIS_MES VT ON V.CODPROD = VT.CODPROD AND V.MES_DATA = VT.MES_DATA
    WHERE V.VALOR_BRUTO > 0
    ORDER BY V.DATA DESC, V.NF, V.CODPROD
    """
    
    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]
    
    # Calcular margem após buscar os dados
    if not df.empty:
        df['margem_valor'] = df['valor_total'] - df['custo_total']
        df['margem_percentual'] = (df['margem_valor'] / df['valor_total'] * 100).round(2)
        df['margem_percentual'] = df['margem_percentual'].fillna(0)
    
    return df

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class VendasClientePDF:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Vendas por Cliente - PDF")
        self.root.geometry("700x600")
        
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
        titulo = ttk.Label(main_frame, text="📊 Vendas por Cliente - PDF", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de cliente
        frame_cliente = ttk.LabelFrame(main_frame, text="Cliente", padding="15")
        frame_cliente.pack(fill=tk.X, pady=10)
        
        ttk.Label(frame_cliente, text="Código(s) do Cliente (separados por vírgula):").pack(anchor=tk.W)
        self.entry_cliente = ttk.Entry(frame_cliente, width=50)
        self.entry_cliente.pack(fill=tk.X, pady=5)
        
        # Frame de período
        frame_periodo = ttk.LabelFrame(main_frame, text="Selecionar Período", padding="15")
        frame_periodo.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Scrollable frame para meses
        canvas = tk.Canvas(frame_periodo, height=200)
        scrollbar = ttk.Scrollbar(frame_periodo, orient="vertical", command=canvas.yview)
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
        for i in range(12):
            mes_data = hoje - timedelta(days=30*i)
            mes_nome = mes_data.strftime("%B/%Y").title()
            mes_key = mes_data.strftime("%m/%Y")
            meses_lista.append((mes_key, mes_nome))
        
        # Criar checkboxes
        for mes_key, mes_nome in meses_lista:
            var = tk.BooleanVar()
            self.meses_vars[mes_key] = var
            cb = ttk.Checkbutton(scrollable_frame, text=mes_nome, variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botão gerar PDF
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        self.btn_gerar = ttk.Button(btn_frame, text="📄 Gerar PDF", 
                                   command=self.gerar_pdf, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
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
    
    def gerar_pdf(self):
        """Gera o PDF com vendas do cliente"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        # Validar cliente
        cliente_str = self.entry_cliente.get().strip()
        if not cliente_str:
            messagebox.showerror("Erro", "Informe pelo menos um código de cliente")
            return
        
        # Processar códigos de cliente
        codcli_list = []
        for cod in cliente_str.split(','):
            cod_limpo = cod.strip()
            if cod_limpo.isdigit():
                codcli_list.append(int(cod_limpo))
        
        if not codcli_list:
            messagebox.showerror("Erro", "Nenhum código de cliente válido encontrado")
            return
        
        # Validar meses selecionados
        meses_selecionados = [mes for mes, var in self.meses_vars.items() if var.get()]
        if not meses_selecionados:
            messagebox.showerror("Erro", "Selecione pelo menos um mês")
            return
        
        def executar():
            try:
                self.status_label.config(text="Gerando PDF...")
                
                # Nome do arquivo
                cliente_nome = "_".join(map(str, codcli_list))
                nome_arquivo = f"vendas_cliente_{cliente_nome}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                # Gerar PDF
                self.criar_pdf(caminho_arquivo, codcli_list, meses_selecionados)
                
                self.status_label.config(text="✅ PDF gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"PDF gerado:\n{caminho_arquivo}")
                
                # Abrir PDF
                try:
                    os.startfile(str(caminho_arquivo))
                except:
                    pass
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar")
                import traceback
                traceback.print_exc()
        
        threading.Thread(target=executar, daemon=True).start()
    
    def criar_pdf(self, caminho, codcli_list, meses_selecionados):
        """Cria o arquivo PDF"""
        doc = SimpleDocTemplate(str(caminho), pagesize=A4,
                               rightMargin=1.5*cm, leftMargin=1.5*cm,
                               topMargin=2*cm, bottomMargin=1.5*cm)
        
        story = []
        styles = getSampleStyleSheet()
        
        # Processar cada cliente
        for codcli in codcli_list:
            # Buscar dados do cliente
            dados_cliente = buscar_dados_cliente(self.conn, codcli)
            if dados_cliente is None:
                continue
            
            # Título do cliente
            titulo_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#124377'),
                spaceAfter=20,
                alignment=0  # Esquerda
            )
            
            cliente_nome = str(dados_cliente.get('cliente', 'N/A'))
            cnpj_cpf_val = dados_cliente.get('cnpj_cpf', None)
            if cnpj_cpf_val is None or (isinstance(cnpj_cpf_val, float) and pd.isna(cnpj_cpf_val)):
                cnpj_cpf = 'N/A'
            else:
                cnpj_cpf = str(cnpj_cpf_val)
            
            # Formatar CNPJ/CPF
            if len(cnpj_cpf) == 14:
                cnpj_cpf = f"{cnpj_cpf[:2]}.{cnpj_cpf[2:5]}.{cnpj_cpf[5:8]}/{cnpj_cpf[8:12]}-{cnpj_cpf[12:]}"
            elif len(cnpj_cpf) == 11:
                cnpj_cpf = f"{cnpj_cpf[:3]}.{cnpj_cpf[3:6]}.{cnpj_cpf[6:9]}-{cnpj_cpf[9:]}"
            
            story.append(Paragraph(f"<b>Cliente:</b> {cliente_nome}", titulo_style))
            story.append(Paragraph(f"<b>Código:</b> {codcli} | <b>CNPJ/CPF:</b> {cnpj_cpf}", styles['Normal']))
            story.append(Spacer(1, 0.5*cm))
            
            # Processar cada mês
            for mes_key in sorted(meses_selecionados):
                mes, ano = mes_key.split('/')
                mes_int = int(mes)
                ano_int = int(ano)
                
                # Calcular datas do mês
                data_inicio = f"01/{mes}/{ano}"
                ultimo_dia = monthrange(ano_int, mes_int)[1]
                data_fim = f"{ultimo_dia:02d}/{mes}/{ano}"
                
                # Buscar vendas do mês
                df_vendas = buscar_vendas_cliente_periodo(self.conn, codcli, data_inicio, data_fim)
                
                if df_vendas is None or len(df_vendas) == 0:
                    continue
                
                # Título do mês
                meses_nomes = {
                    '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março', '04': 'Abril',
                    '05': 'Maio', '06': 'Junho', '07': 'Julho', '08': 'Agosto',
                    '09': 'Setembro', '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
                }
                mes_nome = meses_nomes.get(mes, mes)
                
                mes_style = ParagraphStyle(
                    'MesTitle',
                    parent=styles['Heading2'],
                    fontSize=14,
                    textColor=colors.HexColor('#1a5a9a'),
                    spaceAfter=10,
                    spaceBefore=15
                )
                
                story.append(Paragraph(f"<b>{mes_nome}/{ano}</b>", mes_style))
                
                # Calcular totais do mês
                total_valor = float(df_vendas['valor_total'].sum())
                total_custo = float(df_vendas['custo_total'].sum())
                total_margem = float(df_vendas['margem_valor'].sum())
                total_quantidade = float(df_vendas['quantidade'].sum())
                margem_percentual = (total_margem / total_valor * 100) if total_valor > 0 else 0
                
                # Resumo do mês
                resumo_text = f"""
                <b>Resumo do Mês:</b><br/>
                Valor Total: R$ {total_valor:,.2f}<br/>
                Custo Total: R$ {total_custo:,.2f}<br/>
                Margem: R$ {total_margem:,.2f} ({margem_percentual:.2f}%)<br/>
                Quantidade Total: {total_quantidade:,.0f} unidades
                """
                story.append(Paragraph(resumo_text, styles['Normal']))
                story.append(Spacer(1, 0.3*cm))
                
                # Agrupar por produto e pegar top 5
                df_produtos = df_vendas.groupby(['codprod', 'produto']).agg({
                    'quantidade': 'sum',
                    'valor_total': 'sum',
                    'custo_total': 'sum',
                    'margem_valor': 'sum'
                }).reset_index()
                
                # Calcular margem percentual por produto
                df_produtos['margem_percentual'] = (df_produtos['margem_valor'] / df_produtos['valor_total'] * 100).round(2)
                df_produtos = df_produtos.fillna(0)
                
                # Ordenar por valor total e pegar top 5
                df_produtos = df_produtos.sort_values('valor_total', ascending=False).head(5)
                
                # Preparar dados da tabela (top 5 produtos)
                dados_tabela = [['#', 'Produto', 'Quantidade', 'Valor Total', 'Custo', 'Margem', 'Margem %']]
                
                for idx, (_, row) in enumerate(df_produtos.iterrows(), start=1):
                    produto_nome = str(row['produto'])[:40] if len(str(row['produto'])) > 40 else str(row['produto'])
                    
                    dados_tabela.append([
                        str(idx),
                        produto_nome,
                        f"{float(row['quantidade']):,.0f}",
                        f"R$ {float(row['valor_total']):,.2f}",
                        f"R$ {float(row['custo_total']):,.2f}",
                        f"R$ {float(row['margem_valor']):,.2f}",
                        f"{float(row['margem_percentual']):.2f}%"
                    ])
                
                # Criar tabela
                tabela = Table(dados_tabela, repeatRows=1)
                
                # Estilo da tabela
                estilo_tabela = TableStyle([
                    # Cabeçalho
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124377')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('TOPPADDING', (0, 0), (-1, 0), 8),
                    
                    # Dados
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Número centralizado
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),  # Produto à esquerda
                    ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),  # Números à direita
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f5f5f5')]),
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                    ('TOPPADDING', (0, 1), (-1, -1), 6),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ])
                
                tabela.setStyle(estilo_tabela)
                
                # Ajustar larguras
                col_widths = [0.8*cm, 6*cm, 2*cm, 2.2*cm, 2*cm, 2*cm, 1.8*cm]
                tabela._argW = col_widths
                
                story.append(Paragraph("<b>Top 5 Produtos (Maior Venda):</b>", styles['Normal']))
                story.append(Spacer(1, 0.2*cm))
                story.append(tabela)
                story.append(Spacer(1, 0.5*cm))
            
            # Quebra de página entre clientes
            if codcli != codcli_list[-1]:
                story.append(PageBreak())
        
        # Gerar PDF
        doc.build(story)

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = VendasClientePDF(root)
    root.mainloop()

# portfolio-commit-ready: vendas_cliente_pdf
