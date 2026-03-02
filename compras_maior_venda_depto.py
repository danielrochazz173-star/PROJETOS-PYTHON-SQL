"""
Sistema de Análise: Departamentos com Compra Maior que Venda
Gera PDF com os 10 principais departamentos onde a compra em R$ é maior que a venda
"""

import oracledb
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from calendar import month_abbr
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

def format_real(valor):
    """Formata valor em Real brasileiro"""
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

def get_mes_ano_formatado(mes, ano):
    """Retorna formato NOV/25 para o título"""
    meses_abrev = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                   'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    return f"{meses_abrev[mes-1]}/{str(ano)[-2:]}"

def buscar_compras_por_depto(conn, mes, ano):
    """Busca compras por departamento no mês especificado"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    query = """
    SELECT 
        D.CODEPTO,
        D.DESCRICAO AS DEPARTAMENTO,
        SUM(NVL(M.QT, 0) * NVL(M.PUNIT, 0)) AS VALOR_COMPRA
    FROM PCNFENT N
    JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
    JOIN PCPRODUT P ON M.CODPROD = P.CODPROD
    JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
    WHERE N.DTENT >= :data_inicio
      AND N.DTENT < :data_fim
      AND D.DESCRICAO IS NOT NULL
    GROUP BY D.CODEPTO, D.DESCRICAO
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar compras: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_supervisores(conn):
    """Busca lista de supervisores do banco"""
    query = """
    SELECT 
        CODSUPERVISOR,
        NOME
    FROM PCSUPERV
    ORDER BY CODSUPERVISOR
    """
    try:
        df = pd.read_sql(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar supervisores: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_vendas_por_depto(conn, mes, ano, supervisores=None):
    """Busca vendas por departamento no mês especificado"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    # Construir filtro de supervisores
    if supervisores and len(supervisores) > 0:
        supervisores_str = ','.join(map(str, supervisores))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    else:
        filtro_supervisor = ""
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            P.CODEPTO,
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
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE C.DATA >= :data_inicio
          AND C.DATA < :data_fim
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          {filtro_supervisor}
        GROUP BY P.CODEPTO
    ),
    DEVOLUCOES AS (
        SELECT 
            P.CODEPTO,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
        WHERE D.DTENT >= :data_inicio
          AND D.DTENT < :data_fim
          {filtro_supervisor}
        GROUP BY P.CODEPTO
    )
    SELECT 
        DEP.CODEPTO,
        DEP.DESCRICAO AS DEPARTAMENTO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS VALOR_VENDA
    FROM PCDEPTO DEP 
    LEFT JOIN VENDAS_VALIDAS V ON DEP.CODEPTO = V.CODEPTO 
    LEFT JOIN DEVOLUCOES D ON DEP.CODEPTO = D.CODEPTO
    WHERE (NVL(V.VALOR_BRUTO, 0) > 0 OR NVL(D.VALOR_DEVOLVIDO, 0) > 0)
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar vendas: {e}")
        return pd.DataFrame()

def buscar_totais(conn, mes, ano, supervisores=None):
    """Busca totais de compra e venda"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    # Construir filtro de supervisores
    if supervisores and len(supervisores) > 0:
        supervisores_str = ','.join(map(str, supervisores))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    else:
        filtro_supervisor = ""
    
    # Total de compras
    query_compras = """
    SELECT 
        SUM(NVL(M.QT, 0) * NVL(M.PUNIT, 0)) AS TOTAL_COMPRA
    FROM PCNFENT N
    JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
    WHERE N.DTENT >= :data_inicio
      AND N.DTENT < :data_fim
    """
    
    # Total de vendas
    query_vendas = f"""
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
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE C.DATA >= :data_inicio
          AND C.DATA < :data_fim
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          {filtro_supervisor}
    ),
    DEVOLUCOES AS (
        SELECT 
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
        WHERE D.DTENT >= :data_inicio
          AND D.DTENT < :data_fim
          {filtro_supervisor}
    )
    SELECT 
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS TOTAL_VENDA
    FROM VENDAS_VALIDAS V
    CROSS JOIN DEVOLUCOES D
    """
    
    try:
        df_compra = pd.read_sql(query_compras, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        df_venda = pd.read_sql(query_vendas, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        
        total_compra = float(df_compra['TOTAL_COMPRA'].iloc[0] if not df_compra.empty and pd.notna(df_compra['TOTAL_COMPRA'].iloc[0]) else 0)
        total_venda = float(df_venda['TOTAL_VENDA'].iloc[0] if not df_venda.empty and pd.notna(df_venda['TOTAL_VENDA'].iloc[0]) else 0)
        
        return total_compra, total_venda
    except Exception as e:
        print(f"Erro ao buscar totais: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0

def analisar_deptos_compras_maior_venda(conn, mes, ano, supervisores=None):
    """Analisa departamentos onde compra > venda"""
    print(f"📊 Buscando compras do mês {mes}/{ano}...")
    df_compras = buscar_compras_por_depto(conn, mes, ano)
    
    print(f"📊 Buscando vendas do mês {mes}/{ano}...")
    df_vendas = buscar_vendas_por_depto(conn, mes, ano, supervisores)
    
    # Merge dos dados
    df_merge = pd.merge(
        df_compras[['codepto', 'departamento', 'valor_compra']],
        df_vendas[['codepto', 'valor_venda']],
        on='codepto',
        how='outer'
    )
    
    # Preencher valores nulos com 0
    df_merge['valor_compra'] = df_merge['valor_compra'].fillna(0)
    df_merge['valor_venda'] = df_merge['valor_venda'].fillna(0)
    
    # Filtrar onde compra > venda
    df_resultado = df_merge[df_merge['valor_compra'] > df_merge['valor_venda']].copy()
    
    # Calcular diferença
    df_resultado['diferenca'] = df_resultado['valor_compra'] - df_resultado['valor_venda']
    
    # Ordenar por diferença (maior primeiro)
    df_resultado = df_resultado.sort_values('diferenca', ascending=False)
    
    # Pegar top 12 (4 linhas de 3)
    df_resultado = df_resultado.head(12)
    
    return df_resultado

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

def gerar_html(df_resultado, mes, ano):
    """Gera HTML bonito no estilo da imagem"""
    mes_formatado = get_mes_ano_formatado(mes, ano)
    titulo = f"FORNECEDORES COM A COMPRA MAIOR QUE A VENDA {mes_formatado}"
    
    # Nome do arquivo
    nome_arquivo = f'compras_maior_venda_{mes:02d}_{ano}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
    html_path = Path(nome_arquivo)
    
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
            max-width: 1200px;
            margin: 0 auto;
        }}
        
        .titulo {{
            text-align: center;
            color: #ffffff;
            font-size: 28px;
            font-weight: bold;
            margin-bottom: 40px;
            text-transform: uppercase;
            letter-spacing: 2px;
        }}
        
        .cards-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 30px;
            margin-bottom: 30px;
        }}
        
        @media (max-width: 1200px) {{
            .cards-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
        
        @media (max-width: 768px) {{
            .cards-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .card {{
            background: rgba(18, 67, 119, 0.3);
            border-left: 4px solid #ffd700;
            border-radius: 8px;
            padding: 30px;
            backdrop-filter: blur(10px);
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4);
        }}
        
        .departamento {{
            color: #ffffff;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 25px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .valor-compra {{
            color: #ffd700;
            font-size: 32px;
            font-weight: bold;
            margin-bottom: 8px;
            text-shadow: 0 2px 10px rgba(255, 215, 0, 0.3);
        }}
        
        .label-compra {{
            color: #ffffff;
            font-size: 16px;
            margin-bottom: 20px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .valor-venda {{
            color: #ffd700;
            font-size: 32px;
            font-weight: bold;
            margin-bottom: 8px;
            text-shadow: 0 2px 10px rgba(255, 215, 0, 0.3);
        }}
        
        .label-venda {{
            color: #ffffff;
            font-size: 16px;
            text-transform: uppercase;
            letter-spacing: 1px;
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
            .card {{
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
        <h1 class="titulo">{titulo}</h1>
        
        <div class="cards-grid">
"""
    
    # Adicionar cards para cada departamento
    for _, row in df_resultado.iterrows():
        depto = str(row['departamento']).upper()
        valor_compra = formatar_valor(row['valor_compra'])
        valor_venda = formatar_valor(row['valor_venda'])
        
        html_content += f"""
            <div class="card">
                <div class="departamento">{depto}</div>
                <div class="valor-compra">{valor_compra}</div>
                <div class="label-compra">COMPRA</div>
                <div class="valor-venda">{valor_venda}</div>
                <div class="label-venda">VENDA</div>
            </div>
"""
    
    # Caminho da logo
    logo_path = Path(__file__).parent / 'logo-dicon.png'
    logo_path_str = str(logo_path.resolve()).replace('\\', '/')
    
    html_content += f"""
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

def gerar_html_totais(total_venda, total_compra, mes, ano):
    """Gera HTML com totais no estilo da imagem"""
    mes_formatado = get_mes_ano_formatado(mes, ano)
    
    # Nome do arquivo
    nome_arquivo = f'totais_{mes:02d}_{ano}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
    html_path = Path(nome_arquivo)
    
    valor_venda = formatar_valor_completo(total_venda)
    valor_compra = formatar_valor_completo(total_compra)
    
    # Caminho da logo
    logo_path = Path(__file__).parent / 'logo-dicon.png'
    logo_path_str = str(logo_path.resolve()).replace('\\', '/')
    
    # Gerar HTML
    html_content = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Totais - {mes_formatado}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Arial', sans-serif;
            background: linear-gradient(135deg, #0a1a2e 0%, #1a2a3a 50%, #0a1a2e 100%);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 40px 20px;
            position: relative;
            overflow: hidden;
        }}
        
        body::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: 
                radial-gradient(circle at 20% 50%, rgba(18, 67, 119, 0.3) 0%, transparent 50%),
                radial-gradient(circle at 80% 50%, rgba(18, 67, 119, 0.3) 0%, transparent 50%);
            pointer-events: none;
        }}
        
        .titulo {{
            color: #ffffff;
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 60px;
            text-transform: uppercase;
            letter-spacing: 3px;
            text-align: center;
            text-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
            position: relative;
            z-index: 1;
        }}
        
        .container {{
            display: flex;
            gap: 80px;
            max-width: 1400px;
            width: 100%;
            justify-content: center;
            align-items: center;
            position: relative;
            z-index: 1;
        }}
        
        .card {{
            text-align: center;
            flex: 1;
            max-width: 500px;
            min-width: 400px;
            padding: 50px 40px;
            background: rgba(18, 67, 119, 0.15);
            border-radius: 20px;
            backdrop-filter: blur(10px);
            border: 2px solid rgba(255, 255, 255, 0.1);
            box-shadow: 
                0 8px 32px rgba(0, 0, 0, 0.3),
                inset 0 1px 0 rgba(255, 255, 255, 0.1);
            position: relative;
            overflow: visible;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .card::before {{
            content: '';
            position: absolute;
            top: -50%;
            left: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255, 215, 0, 0.1) 0%, transparent 70%);
            animation: pulse 4s ease-in-out infinite;
        }}
        
        .card::after {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            border-radius: 20px;
            padding: 2px;
            background: linear-gradient(135deg, rgba(255, 215, 0, 0.3), rgba(255, 215, 0, 0.1));
            -webkit-mask: linear-gradient(#fff 0 0) content-box, linear-gradient(#fff 0 0);
            -webkit-mask-composite: xor;
            mask-composite: exclude;
            pointer-events: none;
        }}
        
        @keyframes pulse {{
            0%, 100% {{
                opacity: 0.3;
                transform: scale(1);
            }}
            50% {{
                opacity: 0.6;
                transform: scale(1.1);
            }}
        }}
        
        .card-content {{
            position: relative;
            z-index: 2;
        }}
        
        .label {{
            color: #ffffff;
            font-size: 22px;
            font-weight: bold;
            margin-bottom: 30px;
            text-transform: uppercase;
            letter-spacing: 2px;
            text-shadow: 0 2px 10px rgba(0, 0, 0, 0.5);
        }}
        
        .valor {{
            color: #ffffff;
            font-size: 48px;
            font-weight: bold;
            line-height: 1.2;
            text-shadow: 
                0 4px 20px rgba(0, 0, 0, 0.5),
                0 0 30px rgba(255, 255, 255, 0.2);
            white-space: nowrap;
            display: inline-block;
            max-width: 100%;
        }}
        
        .footer {{
            margin-top: 60px;
            text-align: center;
            position: relative;
            z-index: 1;
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
            }}
            .card::before {{
                animation: none;
            }}
            .footer {{
                margin-top: 40px;
            }}
            .logo {{
                max-width: 120px;
            }}
        }}
    </style>
</head>
<body>
    <h1 class="titulo">Totais - {mes_formatado}</h1>
    
    <div class="container">
        <div class="card">
            <div class="card-content">
                <div class="label">Total Vendido</div>
                <div class="valor">{valor_venda}</div>
            </div>
        </div>
        
        <div class="card">
            <div class="card-content">
                <div class="label">Total Comprado</div>
                <div class="valor">{valor_compra}</div>
            </div>
        </div>
    </div>
    
    <div class="footer">
        <img src="file:///{logo_path_str}" alt="DICON" class="logo" />
    </div>
</body>
</html>
"""
    
    # Caminho da logo para o HTML de totais também
    logo_path = Path(__file__).parent / 'logo-dicon.png'
    logo_path_str = str(logo_path.resolve()).replace('\\', '/')
    
    # Salvar arquivo
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # Abrir no navegador
    import os
    os.startfile(str(html_path.resolve()))
    
    return str(html_path.resolve())

class InterfaceAnalise:
    def __init__(self, master):
        self.master = master
        self.master.title("Análise: Compra > Venda por Departamento")
        self.master.geometry("650x600")
        
        # Notebook para abas
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Aba 1: Departamentos
        self.frame_departamentos = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.frame_departamentos, text="Departamentos")
        
        # Aba 2: Totais
        self.frame_totais = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.frame_totais, text="Totais")
        
        # Configurar frames
        self.frame_departamentos.columnconfigure(1, weight=1)
        self.frame_totais.columnconfigure(1, weight=1)
        
        # Criar interface da aba departamentos
        self.criar_aba_departamentos()
        
        # Criar interface da aba totais
        self.criar_aba_totais()
    
    def criar_aba_departamentos(self):
        """Cria interface da aba de departamentos"""
        main_frame = self.frame_departamentos
        
        # Título
        titulo = ttk.Label(main_frame, text="Análise de Departamentos", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Seleção de mês
        ttk.Label(main_frame, text="Mês:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.mes_var = tk.IntVar(value=datetime.now().month)
        meses = list(range(1, 13))
        meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        self.combo_mes = ttk.Combobox(main_frame, textvariable=self.mes_var, 
                                     values=[f"{m:02d} - {n}" for m, n in zip(meses, meses_nomes)],
                                     state="readonly", width=30)
        self.combo_mes.current(datetime.now().month - 1)
        self.combo_mes.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Seleção de ano
        ttk.Label(main_frame, text="Ano:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.ano_var = tk.IntVar(value=datetime.now().year)
        anos = list(range(2020, datetime.now().year + 2))
        self.combo_ano = ttk.Combobox(main_frame, textvariable=self.ano_var,
                                     values=anos, state="readonly", width=30)
        self.combo_ano.current(len(anos) - 1)
        self.combo_ano.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Frame de supervisores
        frame_supervisores = ttk.LabelFrame(main_frame, text="Supervisores", padding="10")
        frame_supervisores.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        frame_supervisores.columnconfigure(0, weight=1)
        
        # Frame com scroll para checkboxes
        canvas_frame = tk.Frame(frame_supervisores)
        canvas_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Canvas para scroll
        canvas = tk.Canvas(canvas_frame, height=150)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Variáveis para checkboxes
        self.supervisores_vars = {}
        self.supervisores_data = []
        
        # Carregar supervisores
        self.carregar_supervisores(scrollable_frame)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Botões de seleção rápida
        frame_botoes = ttk.Frame(frame_supervisores)
        frame_botoes.grid(row=1, column=0, pady=(10, 0))
        
        ttk.Button(frame_botoes, text="Selecionar Todos", 
                  command=self.selecionar_todos_supervisores).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Limpar Seleção", 
                  command=self.limpar_supervisores).pack(side=tk.LEFT, padx=5)
        
        # Botão gerar
        btn_gerar = ttk.Button(main_frame, text="Gerar Relatório", command=self.gerar_relatorio)
        btn_gerar.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para gerar relatório")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=10)
    
    def criar_aba_totais(self):
        """Cria interface da aba de totais"""
        main_frame = self.frame_totais
        
        # Título
        titulo = ttk.Label(main_frame, text="Totais de Venda e Compra", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Seleção de mês
        ttk.Label(main_frame, text="Mês:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.mes_var_totais = tk.IntVar(value=datetime.now().month)
        meses = list(range(1, 13))
        meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        self.combo_mes_totais = ttk.Combobox(main_frame, textvariable=self.mes_var_totais, 
                                     values=[f"{m:02d} - {n}" for m, n in zip(meses, meses_nomes)],
                                     state="readonly", width=30)
        self.combo_mes_totais.current(datetime.now().month - 1)
        self.combo_mes_totais.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Seleção de ano
        ttk.Label(main_frame, text="Ano:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.ano_var_totais = tk.IntVar(value=datetime.now().year)
        anos = list(range(2020, datetime.now().year + 2))
        self.combo_ano_totais = ttk.Combobox(main_frame, textvariable=self.ano_var_totais,
                                     values=anos, state="readonly", width=30)
        self.combo_ano_totais.current(len(anos) - 1)
        self.combo_ano_totais.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Frame de supervisores
        frame_supervisores_totais = ttk.LabelFrame(main_frame, text="Supervisores", padding="10")
        frame_supervisores_totais.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        frame_supervisores_totais.columnconfigure(0, weight=1)
        
        # Frame com scroll para checkboxes
        canvas_frame_totais = tk.Frame(frame_supervisores_totais)
        canvas_frame_totais.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Canvas para scroll
        canvas_totais = tk.Canvas(canvas_frame_totais, height=150)
        scrollbar_totais = ttk.Scrollbar(canvas_frame_totais, orient="vertical", command=canvas_totais.yview)
        scrollable_frame_totais = ttk.Frame(canvas_totais)
        
        scrollable_frame_totais.bind(
            "<Configure>",
            lambda e: canvas_totais.configure(scrollregion=canvas_totais.bbox("all"))
        )
        
        canvas_totais.create_window((0, 0), window=scrollable_frame_totais, anchor="nw")
        canvas_totais.configure(yscrollcommand=scrollbar_totais.set)
        
        # Variáveis para checkboxes (compartilhadas)
        # Reutilizar os mesmos supervisores da aba departamentos
        self.supervisores_vars_totais = {}
        
        # Carregar supervisores
        self.carregar_supervisores_totais(scrollable_frame_totais)
        
        canvas_totais.pack(side="left", fill="both", expand=True)
        scrollbar_totais.pack(side="right", fill="y")
        
        # Botões de seleção rápida
        frame_botoes_totais = ttk.Frame(frame_supervisores_totais)
        frame_botoes_totais.grid(row=1, column=0, pady=(10, 0))
        
        ttk.Button(frame_botoes_totais, text="Selecionar Todos", 
                  command=self.selecionar_todos_supervisores_totais).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_totais, text="Limpar Seleção", 
                  command=self.limpar_supervisores_totais).pack(side=tk.LEFT, padx=5)
        
        # Botão gerar
        btn_gerar_totais = ttk.Button(main_frame, text="Gerar Totais", command=self.gerar_totais)
        btn_gerar_totais.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Status
        self.status_label_totais = ttk.Label(main_frame, text="Pronto para gerar totais")
        self.status_label_totais.grid(row=5, column=0, columnspan=2, pady=10)
    
    def carregar_supervisores_totais(self, parent):
        """Carrega lista de supervisores para aba de totais"""
        try:
            conn = get_db_connection()
            try:
                df_supervisores = buscar_supervisores(conn)
                supervisores_data = df_supervisores.to_dict('records')
                
                # Criar checkboxes
                for idx, sup in enumerate(supervisores_data):
                    var = tk.BooleanVar(value=True)  # Por padrão, todos selecionados
                    self.supervisores_vars_totais[sup['codsupervisor']] = var
                    
                    checkbox = ttk.Checkbutton(
                        parent,
                        text=f"{sup['codsupervisor']} - {sup['nome']}",
                        variable=var
                    )
                    checkbox.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar supervisores:\n{str(e)}")
    
    def selecionar_todos_supervisores_totais(self):
        """Seleciona todos os supervisores na aba totais"""
        for var in self.supervisores_vars_totais.values():
            var.set(True)
    
    def limpar_supervisores_totais(self):
        """Limpa seleção de supervisores na aba totais"""
        for var in self.supervisores_vars_totais.values():
            var.set(False)
    
    def get_supervisores_selecionados_totais(self):
        """Retorna lista de códigos de supervisores selecionados na aba totais"""
        selecionados = []
        for cod, var in self.supervisores_vars_totais.items():
            if var.get():
                selecionados.append(cod)
        return selecionados
    
    def gerar_totais(self):
        """Gera relatório de totais"""
        try:
            # Obter mês e ano selecionados
            mes_ano_str = self.combo_mes_totais.get()
            mes = int(mes_ano_str.split(' - ')[0])
            ano = int(self.combo_ano_totais.get())
            
            # Obter supervisores selecionados
            supervisores = self.get_supervisores_selecionados_totais()
            
            if not supervisores:
                messagebox.showwarning("Aviso", "Selecione pelo menos um supervisor!")
                return
            
            self.status_label_totais.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            # Conectar ao banco
            conn = get_db_connection()
            
            try:
                self.status_label_totais.config(text="Buscando totais...")
                self.master.update()
                
                # Buscar totais
                total_compra, total_venda = buscar_totais(conn, mes, ano, supervisores)
                
                self.status_label_totais.config(text="Gerando relatório...")
                self.master.update()
                
                # Gerar HTML
                html_path = gerar_html_totais(total_venda, total_compra, mes, ano)
                
                messagebox.showinfo("Sucesso", 
                                  f"Relatório de totais gerado com sucesso!\n\nArquivo: {html_path}\n\n"
                                  f"Total Vendido: {formatar_valor_completo(total_venda)}\n"
                                  f"Total Comprado: {formatar_valor_completo(total_compra)}\n\n"
                                  f"O arquivo foi aberto no navegador. Você pode tirar print ou salvar como PDF.")
                self.status_label_totais.config(text="Relatório gerado com sucesso!")
                
            finally:
                conn.close()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar totais:\n{str(e)}")
            self.status_label_totais.config(text="Erro ao gerar totais")
            import traceback
            traceback.print_exc()
    
    def carregar_supervisores(self, parent):
        """Carrega lista de supervisores do banco"""
        try:
            conn = get_db_connection()
            try:
                df_supervisores = buscar_supervisores(conn)
                self.supervisores_data = df_supervisores.to_dict('records')
                
                # Criar checkboxes
                for idx, sup in enumerate(self.supervisores_data):
                    var = tk.BooleanVar(value=True)  # Por padrão, todos selecionados
                    self.supervisores_vars[sup['codsupervisor']] = var
                    
                    checkbox = ttk.Checkbutton(
                        parent,
                        text=f"{sup['codsupervisor']} - {sup['nome']}",
                        variable=var
                    )
                    checkbox.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar supervisores:\n{str(e)}")
    
    def selecionar_todos_supervisores(self):
        """Seleciona todos os supervisores"""
        for var in self.supervisores_vars.values():
            var.set(True)
    
    def limpar_supervisores(self):
        """Limpa seleção de supervisores"""
        for var in self.supervisores_vars.values():
            var.set(False)
    
    def get_supervisores_selecionados(self):
        """Retorna lista de códigos de supervisores selecionados"""
        selecionados = []
        for cod, var in self.supervisores_vars.items():
            if var.get():
                selecionados.append(cod)
        return selecionados
    
    def gerar_relatorio(self):
        try:
            # Obter mês e ano selecionados
            mes_ano_str = self.combo_mes.get()
            mes = int(mes_ano_str.split(' - ')[0])
            ano = int(self.combo_ano.get())
            
            # Obter supervisores selecionados
            supervisores = self.get_supervisores_selecionados()
            
            if not supervisores:
                messagebox.showwarning("Aviso", "Selecione pelo menos um supervisor!")
                return
            
            self.status_label.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            # Conectar ao banco
            conn = get_db_connection()
            
            try:
                self.status_label.config(text="Analisando dados...")
                self.master.update()
                
                # Analisar departamentos
                df_resultado = analisar_deptos_compras_maior_venda(conn, mes, ano, supervisores)
                
                if df_resultado.empty:
                    messagebox.showinfo("Informação", 
                                      "Nenhum departamento encontrado onde compra > venda no período selecionado.")
                    self.status_label.config(text="Nenhum resultado encontrado")
                    return
                
                self.status_label.config(text="Gerando relatório...")
                self.master.update()
                
                # Gerar HTML
                html_path = gerar_html(df_resultado, mes, ano)
                
                messagebox.showinfo("Sucesso", 
                                  f"Relatório gerado com sucesso!\n\nArquivo: {html_path}\n\n"
                                  f"Total de departamentos: {len(df_resultado)}\n"
                                  f"Supervisores: {len(supervisores)} selecionados\n\n"
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
    app = InterfaceAnalise(root)
    root.mainloop()

if __name__ == "__main__":
    main()

# portfolio-commit-ready: compras_maior_venda_depto
