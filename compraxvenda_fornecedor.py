"""
Sistema de Análise: Compra x Venda por Departamento
Gera relatório HTML com análise mensal de compra e venda por departamento
"""

import oracledb
import pandas as pd
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import os

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# ======= CONFIGURAÇÃO DE ACESSO ORACLE =======
DB_CONFIG = dict(
    host=os.getenv('DB_HOST', '10.0.0.10'),
    port=int(os.getenv('DB_PORT', '1521')),
    service=os.getenv('DB_SERVICE', 'PROD'),
    user=os.getenv('DB_USER', 'powerbi'),
    password=os.getenv('DB_PASSWORD', '')
)

def get_db_connection():
    dsn = f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['service']}"
    try:
        conn = oracledb.connect(user=DB_CONFIG['user'], password=DB_CONFIG['password'], dsn=dsn)
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco Oracle: {e}")
        sys.exit(1)

def buscar_departamentos(conn):
    """Busca lista de departamentos do banco"""
    query = """
    SELECT 
        D.CODEPTO,
        D.DESCRICAO AS DEPARTAMENTO
    FROM PCDEPTO D
    WHERE D.DESCRICAO IS NOT NULL
    ORDER BY D.DESCRICAO
    """
    try:
        df = pd.read_sql(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar departamentos: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_compras_depto_mes(conn, codepto, mes, ano):
    """Busca compras do departamento no mês especificado"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    query = """
    SELECT 
        SUM(NVL(M.QT, 0) * NVL(M.PUNIT, 0)) AS VALOR_COMPRA
    FROM PCNFENT N
    JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
    JOIN PCPRODUT P ON M.CODPROD = P.CODPROD
    JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
    WHERE N.DTENT >= :data_inicio
      AND N.DTENT < :data_fim
      AND D.CODEPTO = :codepto
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim, 'codepto': codepto})
        valor = float(df['VALOR_COMPRA'].iloc[0] if not df.empty and pd.notna(df['VALOR_COMPRA'].iloc[0]) else 0)
        return valor
    except Exception as e:
        print(f"Erro ao buscar compras: {e}")
        return 0

def buscar_vendas_depto_mes(conn, codepto, mes, ano):
    """Busca vendas do departamento no mês especificado"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    query = """
    WITH VENDAS_VALIDAS AS (
        SELECT 
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                        )
                    ELSE 0 
                END
            ) AS VALOR_BRUTO
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= :data_inicio
          AND C.DATA < :data_fim
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          AND P.CODEPTO = :codepto
    ),
    DEVOLUCOES AS (
        SELECT 
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE D.DTENT >= :data_inicio
          AND D.DTENT < :data_fim
          AND P.CODEPTO = :codepto
    )
    SELECT 
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS VALOR_VENDA
    FROM VENDAS_VALIDAS V
    CROSS JOIN DEVOLUCOES D
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim, 'codepto': codepto})
        valor = float(df['VALOR_VENDA'].iloc[0] if not df.empty and pd.notna(df['VALOR_VENDA'].iloc[0]) else 0)
        return valor
    except Exception as e:
        print(f"Erro ao buscar vendas: {e}")
        return 0

def analisar_depto_periodos(conn, codepto, departamento_nome, meses_selecionados, anos_selecionados):
    """Analisa compra e venda do departamento para meses e anos selecionados"""
    resultados = []
    
    # Criar lista de períodos (mes, ano) ordenados
    periodos = []
    for ano in sorted(anos_selecionados):
        for mes in sorted(meses_selecionados):
            periodos.append((mes, ano))
    
    for mes, ano in periodos:
        print(f"📊 Analisando {mes:02d}/{ano}...")
        
        valor_compra = buscar_compras_depto_mes(conn, codepto, mes, ano)
        valor_venda = buscar_vendas_depto_mes(conn, codepto, mes, ano)
        diferenca = valor_venda - valor_compra
        
        resultados.append({
            'mes': mes,
            'ano': ano,
            'mes_ano': f"{mes:02d}/{ano}",
            'mes_abrev': get_mes_abrev(mes),
            'valor_compra': valor_compra,
            'valor_venda': valor_venda,
            'diferenca': diferenca
        })
    
    return resultados

def get_mes_abrev(mes):
    """Retorna abreviação do mês"""
    meses_abrev = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                   'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    return meses_abrev[mes - 1]

def formatar_valor(valor):
    """Formata valor para exibição (sem R$)"""
    try:
        return f"{float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "0,00"

def formatar_valor_completo(valor):
    """Formata valor completo em R$ com pontuação brasileira"""
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

def gerar_html_depto(resultados, codepto, departamento_nome, meses_selecionados, anos_selecionados):
    """Gera HTML bonito com análise mensal do departamento"""
    
    # Título
    titulo = f"COMPRA X VENDA - {departamento_nome.upper()}"
    
    # Criar subtítulo com períodos selecionados
    meses_nomes = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                   'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    meses_sel = [meses_nomes[m-1] for m in sorted(meses_selecionados)]
    anos_sel = sorted(anos_selecionados)
    
    if len(meses_sel) == 12:
        subtitulo = f"Meses: Todos | Anos: {', '.join(map(str, anos_sel))}"
    else:
        subtitulo = f"Meses: {', '.join(meses_sel)} | Anos: {', '.join(map(str, anos_sel))}"
    
    # Nome do arquivo
    nome_arquivo = f'compraxvenda_depto_{codepto}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
    html_path = Path(nome_arquivo)
    
    # Calcular totais
    total_compra = sum(r['valor_compra'] for r in resultados)
    total_venda = sum(r['valor_venda'] for r in resultados)
    total_diferenca = total_venda - total_compra
    
    # Gerar HTML
    html_content = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{titulo}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Arial', sans-serif;
            background: linear-gradient(135deg, #0a1a2e 0%, #1a2a3a 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 40px;
        }}
        
        .titulo {{
            color: #ffffff;
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 2px;
            text-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
        }}
        
        .subtitulo {{
            color: #ffd700;
            font-size: 20px;
            font-weight: normal;
            letter-spacing: 1px;
        }}
        
        .totais-cards {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 30px;
            margin-bottom: 40px;
        }}
        
        @media (max-width: 1200px) {{
            .totais-cards {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .card-total {{
            background: rgba(18, 67, 119, 0.3);
            border-left: 4px solid #ffd700;
            border-radius: 8px;
            padding: 30px;
            backdrop-filter: blur(10px);
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
            text-align: center;
        }}
        
        .card-total .label {{
            color: #ffffff;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .card-total .valor {{
            color: #ffd700;
            font-size: 32px;
            font-weight: bold;
            text-shadow: 0 2px 10px rgba(255, 215, 0, 0.3);
        }}
        
        .card-total.diferenca {{
            border-left-color: #4CAF50;
        }}
        
        .card-total.diferenca .valor {{
            color: #4CAF50;
        }}
        
        .card-total.diferenca.negativa {{
            border-left-color: #F44336;
        }}
        
        .card-total.diferenca.negativa .valor {{
            color: #F44336;
        }}
        
        .tabela-container {{
            background: rgba(18, 67, 119, 0.2);
            border-radius: 12px;
            padding: 30px;
            backdrop-filter: blur(10px);
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
            overflow-x: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 8px;
            overflow: hidden;
        }}
        
        thead {{
            background: rgba(18, 67, 119, 0.5);
        }}
        
        th {{
            color: #ffd700;
            font-size: 16px;
            font-weight: bold;
            padding: 20px 15px;
            text-align: center;
            text-transform: uppercase;
            letter-spacing: 1px;
            border-bottom: 2px solid rgba(255, 215, 0, 0.3);
        }}
        
        td {{
            color: #ffffff;
            font-size: 15px;
            padding: 18px 15px;
            text-align: center;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }}
        
        tbody tr:hover {{
            background: rgba(255, 215, 0, 0.1);
            transition: background 0.3s ease;
        }}
        
        tbody tr:last-child td {{
            border-bottom: none;
        }}
        
        .mes-col {{
            font-weight: bold;
            color: #ffd700;
        }}
        
        .valor-positivo {{
            color: #4CAF50;
        }}
        
        .valor-negativo {{
            color: #F44336;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 40px;
        }}
        
        .logo {{
            max-width: 150px;
            height: auto;
            display: block;
            margin: 0 auto;
            opacity: 0.9;
        }}
        
        @media print {{
            body {{
                background: #0a1a2e;
                padding: 20px;
            }}
            .tabela-container {{
                page-break-inside: avoid;
            }}
            .logo {{
                max-width: 120px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1 class="titulo">{titulo}</h1>
            <p class="subtitulo">{subtitulo}</p>
        </div>
        
        <div class="totais-cards">
            <div class="card-total">
                <div class="label">Total Comprado</div>
                <div class="valor">{formatar_valor_completo(total_compra)}</div>
            </div>
            <div class="card-total">
                <div class="label">Total Vendido</div>
                <div class="valor">{formatar_valor_completo(total_venda)}</div>
            </div>
            <div class="card-total diferenca {'negativa' if total_diferenca < 0 else ''}">
                <div class="label">Diferença (Venda - Compra)</div>
                <div class="valor">{formatar_valor_completo(total_diferenca)}</div>
            </div>
        </div>
        
        <div class="tabela-container">
            <table>
                <thead>
                    <tr>
                        <th>Mês/Ano</th>
                        <th>Compra</th>
                        <th>Venda</th>
                        <th>Diferença</th>
                    </tr>
                </thead>
                <tbody>
"""
    
    # Adicionar linhas da tabela (já ordenadas por ano e mês)
    for resultado in resultados:
        mes_ano = resultado['mes_ano']
        mes_abrev = resultado['mes_abrev']
        ano = resultado['ano']
        compra = resultado['valor_compra']
        venda = resultado['valor_venda']
        diferenca = resultado['diferenca']
        
        classe_diferenca = 'valor-positivo' if diferenca >= 0 else 'valor-negativo'
        
        html_content += f"""
                    <tr>
                        <td class="mes-col">{mes_abrev}/{ano}</td>
                        <td>{formatar_valor_completo(compra)}</td>
                        <td>{formatar_valor_completo(venda)}</td>
                        <td class="{classe_diferenca}">{formatar_valor_completo(diferenca)}</td>
                    </tr>
"""
    
    # Caminho da logo
    logo_path = Path(__file__).parent / 'logo-dicon.png'
    logo_path_str = str(logo_path.resolve()).replace('\\', '/')
    
    html_content += f"""
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <img src="file:///{logo_path_str}" alt="DICON" class="logo" />
        </div>
    </div>
</body>
</html>
"""
    
    # Salvar arquivo
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # Abrir no navegador
    import os
    os.startfile(str(html_path.resolve()))
    
    return str(html_path.resolve())

class InterfaceDepto:
    def __init__(self, master):
        self.master = master
        self.master.title("Análise: Compra x Venda por Departamento")
        self.master.geometry("700x700")
        
        # Frame principal com scroll
        canvas = tk.Canvas(self.master)
        scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Análise por Departamento", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Seleção de departamento
        ttk.Label(main_frame, text="Departamento:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.depto_var = tk.StringVar()
        self.combo_depto = ttk.Combobox(main_frame, textvariable=self.depto_var,
                                        state="readonly", width=50)
        self.combo_depto.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Frame para meses
        frame_meses = ttk.LabelFrame(main_frame, text="Selecionar Meses", padding="10")
        frame_meses.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        meses_nomes = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                      'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
        self.meses_vars = {}
        
        # Criar checkboxes para meses em 3 colunas
        for idx, (mes_num, mes_nome) in enumerate(zip(range(1, 13), meses_nomes)):
            var = tk.BooleanVar(value=True)  # Por padrão, todos selecionados
            self.meses_vars[mes_num] = var
            
            row = idx // 4
            col = idx % 4
            
            checkbox = ttk.Checkbutton(
                frame_meses,
                text=mes_nome,
                variable=var
            )
            checkbox.grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
        
        # Botões rápidos para meses
        frame_botoes_meses = ttk.Frame(frame_meses)
        frame_botoes_meses.grid(row=3, column=0, columnspan=4, pady=(10, 0))
        
        ttk.Button(frame_botoes_meses, text="Selecionar Todos", 
                  command=self.selecionar_todos_meses).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_meses, text="Limpar Seleção", 
                  command=self.limpar_meses).pack(side=tk.LEFT, padx=5)
        
        # Frame para anos
        frame_anos = ttk.LabelFrame(main_frame, text="Selecionar Anos", padding="10")
        frame_anos.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Criar lista de anos (últimos 5 anos + ano atual)
        ano_atual = datetime.now().year
        anos_lista = list(range(ano_atual - 4, ano_atual + 2))
        self.anos_vars = {}
        
        # Criar checkboxes para anos
        for idx, ano in enumerate(anos_lista):
            var = tk.BooleanVar(value=True if ano == ano_atual or ano == ano_atual - 1 else False)
            self.anos_vars[ano] = var
            
            checkbox = ttk.Checkbutton(
                frame_anos,
                text=str(ano),
                variable=var
            )
            checkbox.grid(row=idx // 5, column=idx % 5, sticky=tk.W, padx=10, pady=5)
        
        # Botões rápidos para anos
        frame_botoes_anos = ttk.Frame(frame_anos)
        frame_botoes_anos.grid(row=1, column=0, columnspan=5, pady=(10, 0))
        
        ttk.Button(frame_botoes_anos, text="Selecionar Todos", 
                  command=self.selecionar_todos_anos).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_anos, text="Limpar Seleção", 
                  command=self.limpar_anos).pack(side=tk.LEFT, padx=5)
        
        # Botão carregar departamentos
        btn_carregar = ttk.Button(main_frame, text="Carregar Departamentos", 
                                 command=self.carregar_departamentos)
        btn_carregar.grid(row=4, column=0, columnspan=2, pady=15)
        
        # Botão gerar
        btn_gerar = ttk.Button(main_frame, text="Gerar Relatório", command=self.gerar_relatorio)
        btn_gerar.grid(row=5, column=0, columnspan=2, pady=20)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Clique em 'Carregar Departamentos' para começar")
        self.status_label.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Dados
        self.departamentos_data = []
        
        # Pack canvas e scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Carregar departamentos automaticamente ao iniciar
        self.master.after(100, self.carregar_departamentos)
    
    def selecionar_todos_meses(self):
        """Seleciona todos os meses"""
        for var in self.meses_vars.values():
            var.set(True)
    
    def limpar_meses(self):
        """Limpa seleção de meses"""
        for var in self.meses_vars.values():
            var.set(False)
    
    def selecionar_todos_anos(self):
        """Seleciona todos os anos"""
        for var in self.anos_vars.values():
            var.set(True)
    
    def limpar_anos(self):
        """Limpa seleção de anos"""
        for var in self.anos_vars.values():
            var.set(False)
    
    def get_meses_selecionados(self):
        """Retorna lista de meses selecionados"""
        return [mes for mes, var in self.meses_vars.items() if var.get()]
    
    def get_anos_selecionados(self):
        """Retorna lista de anos selecionados"""
        return [ano for ano, var in self.anos_vars.items() if var.get()]
    
    def carregar_departamentos(self):
        """Carrega lista de departamentos do banco"""
        try:
            self.status_label.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            conn = get_db_connection()
            try:
                self.status_label.config(text="Buscando departamentos...")
                self.master.update()
                
                df_departamentos = buscar_departamentos(conn)
                
                if df_departamentos.empty:
                    messagebox.showwarning("Aviso", "Nenhum departamento encontrado!")
                    self.status_label.config(text="Nenhum departamento encontrado")
                    return
                
                self.departamentos_data = df_departamentos.to_dict('records')
                
                # Criar lista de strings para o combobox
                deptos_list = [f"{d['codepto']} - {d['departamento']}" for d in self.departamentos_data]
                self.combo_depto['values'] = deptos_list
                
                if deptos_list:
                    self.combo_depto.current(0)
                
                self.status_label.config(text=f"{len(deptos_list)} departamentos carregados")
                messagebox.showinfo("Sucesso", f"{len(deptos_list)} departamentos carregados com sucesso!")
                
            finally:
                conn.close()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar departamentos:\n{str(e)}")
            self.status_label.config(text="Erro ao carregar departamentos")
            import traceback
            traceback.print_exc()
    
    def gerar_relatorio(self):
        """Gera relatório de análise do departamento"""
        try:
            # Validar departamento selecionado
            depto_str = self.depto_var.get()
            if not depto_str:
                messagebox.showwarning("Aviso", "Selecione um departamento!")
                return
            
            # Extrair código do departamento
            codepto = int(depto_str.split(' - ')[0])
            depto_nome = ' - '.join(depto_str.split(' - ')[1:])
            
            # Obter meses e anos selecionados
            meses_selecionados = self.get_meses_selecionados()
            anos_selecionados = self.get_anos_selecionados()
            
            if not meses_selecionados:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês!")
                return
            
            if not anos_selecionados:
                messagebox.showwarning("Aviso", "Selecione pelo menos um ano!")
                return
            
            self.status_label.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            # Conectar ao banco
            conn = get_db_connection()
            
            try:
                total_periodos = len(meses_selecionados) * len(anos_selecionados)
                self.status_label.config(text=f"Analisando {total_periodos} períodos... Isso pode levar alguns minutos...")
                self.master.update()
                
                # Analisar departamento
                resultados = analisar_depto_periodos(conn, codepto, depto_nome, meses_selecionados, anos_selecionados)
                
                if not resultados:
                    messagebox.showinfo("Informação", 
                                      "Nenhum dado encontrado para o departamento no período selecionado.")
                    self.status_label.config(text="Nenhum resultado encontrado")
                    return
                
                self.status_label.config(text="Gerando relatório HTML...")
                self.master.update()
                
                # Gerar HTML
                html_path = gerar_html_depto(resultados, codepto, depto_nome, meses_selecionados, anos_selecionados)
                
                messagebox.showinfo("Sucesso", 
                                  f"Relatório gerado com sucesso!\n\nArquivo: {html_path}\n\n"
                                  f"Departamento: {depto_nome}\n"
                                  f"Períodos analisados: {len(resultados)}\n\n"
                                  f"O arquivo foi aberto no navegador. Você pode tirar print ou salvar como PDF.")
                self.status_label.config(text="Relatório gerado com sucesso!")
                
            finally:
                conn.close()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{str(e)}")
            self.status_label.config(text="Erro ao gerar relatório")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = InterfaceDepto(root)
    root.mainloop()

if __name__ == "__main__":
    main()

# portfolio-commit-ready: compraxvenda_fornecedor
