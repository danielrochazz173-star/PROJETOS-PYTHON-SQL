"""
Sistema de Curva ABC de Clientes - Interface Tkinter
Gera PDF com relação de vendas por cliente (razão social, CNPJ/CPF)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import os
import threading
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        return None

def buscar_vendas_clientes(conn, data_inicio, data_fim):
    """Busca vendas por cliente no período"""
    query = f"""
    SELECT 
        CL.CODCLI,
        CL.CLIENTE,
        CL.CGCENT AS CNPJ_CPF,
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
        COUNT(DISTINCT C.NUMPED) AS QTD_PEDIDOS
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    WHERE C.DATA BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
        AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
        AND C.CODFILIAL IN ('1', '98')
        AND C.POSICAO = 'F'
        AND C.DTCANCEL IS NULL
        AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
    GROUP BY CL.CODCLI, CL.CLIENTE, CL.CGCENT
    HAVING SUM(
        CASE 
            WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                DECODE(C.CONDVENDA, 
                    5, 0, 6, 0, 11, 0, 12, 0, 
                    ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                )
            ELSE 0 
        END
    ) > 0
    ORDER BY VALOR_BRUTO DESC
    """
    
    df = pd.read_sql_query(query, conn)
    return df

def calcular_curva_abc(df):
    """Calcula a Curva ABC dos clientes"""
    df = df.copy()
    df = df.sort_values('VALOR_BRUTO', ascending=False)
    
    # Calcular totais
    total_faturamento = df['VALOR_BRUTO'].sum()
    
    # Calcular percentual acumulado
    df['VALOR_ACUMULADO'] = df['VALOR_BRUTO'].cumsum()
    df['PERCENTUAL'] = (df['VALOR_BRUTO'] / total_faturamento * 100).round(2)
    df['PERCENTUAL_ACUMULADO'] = (df['VALOR_ACUMULADO'] / total_faturamento * 100).round(2)
    
    # Classificar ABC
    df['CLASSIFICACAO'] = df['PERCENTUAL_ACUMULADO'].apply(
        lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
    )
    
    return df

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class CurvaABCClientes:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Curva ABC de Clientes")
        self.root.geometry("600x400")
        
        # Variáveis
        self.conn = None
        self.tipo_periodo = tk.StringVar(value="60_dias")
        
        # Criar interface
        self.criar_interface()
        
        # Conectar ao banco
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Curva ABC de Clientes", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de período
        frame_periodo = ttk.LabelFrame(main_frame, text="Selecionar Período", padding="15")
        frame_periodo.pack(fill=tk.X, pady=10)
        
        # Opções de período
        ttk.Radiobutton(frame_periodo, text="Últimos 30 dias", 
                       variable=self.tipo_periodo, value="30_dias").pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(frame_periodo, text="Últimos 60 dias", 
                       variable=self.tipo_periodo, value="60_dias").pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(frame_periodo, text="Mês atual completo", 
                       variable=self.tipo_periodo, value="mes_atual").pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(frame_periodo, text="Data específica (início e fim)", 
                       variable=self.tipo_periodo, value="especifico").pack(anchor=tk.W, pady=5)
        
        # Frame de datas específicas
        self.frame_datas = ttk.Frame(frame_periodo)
        self.frame_datas.pack(fill=tk.X, pady=10)
        
        ttk.Label(self.frame_datas, text="Data Início (DD/MM/AAAA):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_data_inicio = ttk.Entry(self.frame_datas, width=15)
        self.entry_data_inicio.grid(row=0, column=1, padx=5, pady=5)
        self.entry_data_inicio.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        ttk.Label(self.frame_datas, text="Data Fim (DD/MM/AAAA):").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.entry_data_fim = ttk.Entry(self.frame_datas, width=15)
        self.entry_data_fim.grid(row=0, column=3, padx=5, pady=5)
        self.entry_data_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # Inicialmente desabilitar datas específicas
        self.entry_data_inicio.config(state=tk.DISABLED)
        self.entry_data_fim.config(state=tk.DISABLED)
        
        # Atualizar quando mudar tipo de período
        self.tipo_periodo.trace('w', self.atualizar_campos_data)
        
        # Botão gerar PDF
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        self.btn_gerar = ttk.Button(btn_frame, text="📄 Gerar PDF", 
                                   command=self.gerar_pdf, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    def atualizar_campos_data(self, *args):
        """Atualiza campos de data baseado no tipo de período"""
        if self.tipo_periodo.get() == "especifico":
            self.entry_data_inicio.config(state=tk.NORMAL)
            self.entry_data_fim.config(state=tk.NORMAL)
        else:
            self.entry_data_inicio.config(state=tk.DISABLED)
            self.entry_data_fim.config(state=tk.DISABLED)
    
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
    
    def calcular_periodo(self):
        """Calcula as datas do período selecionado"""
        hoje = datetime.now()
        tipo = self.tipo_periodo.get()
        
        if tipo == "30_dias":
            data_fim = hoje
            data_inicio = hoje - timedelta(days=30)
        elif tipo == "60_dias":
            data_fim = hoje
            data_inicio = hoje - timedelta(days=60)
        elif tipo == "mes_atual":
            data_fim = hoje
            data_inicio = datetime(hoje.year, hoje.month, 1)
        else:  # específico
            try:
                data_inicio_str = self.entry_data_inicio.get().strip()
                data_fim_str = self.entry_data_fim.get().strip()
                data_inicio = datetime.strptime(data_inicio_str, "%d/%m/%Y")
                data_fim = datetime.strptime(data_fim_str, "%d/%m/%Y")
            except:
                raise ValueError("Datas inválidas. Use formato DD/MM/AAAA")
        
        return data_inicio.strftime("%d/%m/%Y"), data_fim.strftime("%d/%m/%Y")
    
    def gerar_pdf(self):
        """Gera o PDF com Curva ABC"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        def executar():
            try:
                self.status_label.config(text="Calculando período...")
                data_inicio, data_fim = self.calcular_periodo()
                
                self.status_label.config(text="Buscando dados...")
                df = buscar_vendas_clientes(self.conn, data_inicio, data_fim)
                
                if df.empty:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado no período")
                    self.status_label.config(text="❌ Nenhum registro encontrado")
                    return
                
                self.status_label.config(text="Calculando Curva ABC...")
                df_abc = calcular_curva_abc(df)
                
                self.status_label.config(text="Gerando PDF...")
                
                # Nome do arquivo
                periodo_nome = self.tipo_periodo.get().replace("_", "_")
                nome_arquivo = f"curva_abc_clientes_{periodo_nome}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                # Gerar PDF
                self.criar_pdf(caminho_arquivo, df_abc, data_inicio, data_fim)
                
                self.status_label.config(text="✅ PDF gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"PDF gerado:\n{caminho_arquivo}")
                
                # Abrir PDF
                try:
                    os.startfile(str(caminho_arquivo))
                except:
                    pass
                    
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
                self.status_label.config(text="❌ Erro na validação")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar")
                import traceback
                traceback.print_exc()
        
        threading.Thread(target=executar, daemon=True).start()
    
    def criar_pdf(self, caminho, df, data_inicio, data_fim):
        """Cria o arquivo PDF"""
        doc = SimpleDocTemplate(str(caminho), pagesize=landscape(A4),
                               rightMargin=1.5*cm, leftMargin=1.5*cm,
                               topMargin=2.5*cm, bottomMargin=1.5*cm)
        
        story = []
        styles = getSampleStyleSheet()
        
        # Título
        titulo_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#124377'),
            spaceAfter=30,
            alignment=1  # Centralizado
        )
        
        story.append(Paragraph("📊 CURVA ABC DE CLIENTES", titulo_style))
        story.append(Spacer(1, 0.8*cm))
        
        # Período
        periodo_text = f"<b>Período:</b> {data_inicio} a {data_fim}"
        story.append(Paragraph(periodo_text, styles['Normal']))
        story.append(Spacer(1, 0.2*cm))
        story.append(Paragraph(f"<b>Data de geração:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 0.8*cm))
        
        # Resumo
        total_clientes = len(df)
        total_faturamento = df['VALOR_BRUTO'].sum()
        clientes_a = len(df[df['CLASSIFICACAO'] == 'A'])
        clientes_b = len(df[df['CLASSIFICACAO'] == 'B'])
        clientes_c = len(df[df['CLASSIFICACAO'] == 'C'])
        fat_a = df[df['CLASSIFICACAO'] == 'A']['VALOR_BRUTO'].sum()
        fat_b = df[df['CLASSIFICACAO'] == 'B']['VALOR_BRUTO'].sum()
        fat_c = df[df['CLASSIFICACAO'] == 'C']['VALOR_BRUTO'].sum()
        
        resumo_text = f"""
        <b>RESUMO:</b><br/>
        Total de Clientes: {total_clientes}<br/>
        Faturamento Total: R$ {total_faturamento:,.2f}<br/><br/>
        <b>Classificação ABC:</b><br/>
        Classe A: {clientes_a} clientes (R$ {fat_a:,.2f} - {fat_a/total_faturamento*100:.1f}%)<br/>
        Classe B: {clientes_b} clientes (R$ {fat_b:,.2f} - {fat_b/total_faturamento*100:.1f}%)<br/>
        Classe C: {clientes_c} clientes (R$ {fat_c:,.2f} - {fat_c/total_faturamento*100:.1f}%)
        """
        story.append(Paragraph(resumo_text, styles['Normal']))
        story.append(Spacer(1, 0.8*cm))
        
        # Preparar dados da tabela
        dados_tabela = [['#', 'Código', 'Cliente', 'CNPJ/CPF', 'Faturamento', '%', '% Acum.', 'Classificação', 'Pedidos']]
        
        for idx, row in df.iterrows():
            cnpj_cpf = str(row['CNPJ_CPF']) if pd.notna(row['CNPJ_CPF']) else 'N/A'
            # Formatar CNPJ/CPF
            if len(cnpj_cpf) == 14:
                cnpj_cpf = f"{cnpj_cpf[:2]}.{cnpj_cpf[2:5]}.{cnpj_cpf[5:8]}/{cnpj_cpf[8:12]}-{cnpj_cpf[12:]}"
            elif len(cnpj_cpf) == 11:
                cnpj_cpf = f"{cnpj_cpf[:3]}.{cnpj_cpf[3:6]}.{cnpj_cpf[6:9]}-{cnpj_cpf[9:]}"
            
            # Limitar tamanho do nome do cliente
            nome_cliente = str(row['CLIENTE'])[:35] if len(str(row['CLIENTE'])) > 35 else str(row['CLIENTE'])
            
            dados_tabela.append([
                str(len(dados_tabela)),
                str(int(row['CODCLI'])),
                nome_cliente,
                cnpj_cpf,
                f"R$ {row['VALOR_BRUTO']:,.2f}",
                f"{row['PERCENTUAL']:.2f}%",
                f"{row['PERCENTUAL_ACUMULADO']:.2f}%",
                row['CLASSIFICACAO'],
                str(int(row['QTD_PEDIDOS']))
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
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('LEFTPADDING', (0, 0), (-1, 0), 5),
            ('RIGHTPADDING', (0, 0), (-1, 0), 5),
            
            # Dados
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('ALIGN', (4, 1), (7, -1), 'RIGHT'),  # Números à direita
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f5f5f5')]),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('LEFTPADDING', (0, 1), (-1, -1), 4),
            ('RIGHTPADDING', (0, 1), (-1, -1), 4),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Classificação ABC com cores - aplicar por linha
            ('TEXTCOLOR', (7, 1), (7, -1), colors.black),
        ])
        
        # Adicionar cores de fundo para classificação ABC
        for i, row in enumerate(df.iterrows(), start=1):  # i começa em 1 (pula cabeçalho)
            classificacao = row[1]['CLASSIFICACAO']
            if classificacao == 'A':
                estilo_tabela.add('BACKGROUND', (7, i), (7, i), colors.HexColor('#90EE90'))
            elif classificacao == 'B':
                estilo_tabela.add('BACKGROUND', (7, i), (7, i), colors.HexColor('#FFE4B5'))
            else:  # C
                estilo_tabela.add('BACKGROUND', (7, i), (7, i), colors.HexColor('#FFB6C1'))
        
        tabela.setStyle(estilo_tabela)
        
        # Ajustar larguras das colunas (distribuídas melhor)
        col_widths = [0.8*cm, 1.2*cm, 5.5*cm, 2.8*cm, 2.2*cm, 1.3*cm, 1.5*cm, 1.3*cm, 1.2*cm]
        tabela._argW = col_widths
        
        story.append(tabela)
        
        # Gerar PDF
        doc.build(story)

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = CurvaABCClientes(root)
    root.mainloop()

# portfolio-commit-ready: curva_abc_clientes_pdf
