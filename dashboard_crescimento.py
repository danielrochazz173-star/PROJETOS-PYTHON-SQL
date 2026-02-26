"""
Dashboard de Análise de Crescimento - Comparação 2024 vs 2025
Sistema web com exportação PDF
"""

from flask import Flask, render_template, jsonify
import oracledb
import pandas as pd
from datetime import datetime
import json
import io
import base64
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager

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

app = Flask(__name__)
app.config['SECRET_KEY'] = 'dicon-crescimento-2024-2025'

# =========================================================================
# FUNÇÕES DE BANCO
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}"
        )
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

def buscar_faturamento_ano(conn, ano):
    """Busca faturamento líquido (vendas - devoluções) de um ano"""
    data_inicio = f"01/01/{ano}"
    data_fim = f"31/12/{ano}"
    
    query = f"""
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
        WHERE C.DATA >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND C.DATA <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
    ),
    DEVOLUCOES AS (
        SELECT 
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND D.DTENT <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
    )
    SELECT 
        NVL(V.VALOR_BRUTO, 0) AS FATURAMENTO_BRUTO,
        NVL(D.VALOR_DEVOLVIDO, 0) AS DEVOLUCOES,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
        NVL(V.CUSTO_BRUTO, 0) AS CUSTO_BRUTO,
        NVL(D.CUSTO_DEVOLVIDO, 0) AS CUSTO_DEVOLVIDO,
        (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO,
        ((NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - 
         (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0))) AS LUCRO_LIQUIDO
    FROM VENDAS_VALIDAS V
    CROSS JOIN DEVOLUCOES D
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        if df.empty:
            return {
                'faturamento_bruto': 0,
                'devolucoes': 0,
                'faturamento_liquido': 0,
                'custo_bruto': 0,
                'custo_devolvido': 0,
                'custo_liquido': 0,
                'lucro_liquido': 0
            }
        
        row = df.iloc[0]
        return {
            'faturamento_bruto': float(row['FATURAMENTO_BRUTO']),
            'devolucoes': float(row['DEVOLUCOES']),
            'faturamento_liquido': float(row['FATURAMENTO_LIQUIDO']),
            'custo_bruto': float(row['CUSTO_BRUTO']),
            'custo_devolvido': float(row['CUSTO_DEVOLVIDO']),
            'custo_liquido': float(row['CUSTO_LIQUIDO']),
            'lucro_liquido': float(row['LUCRO_LIQUIDO'])
        }
    except Exception as e:
        print(f"Erro ao buscar faturamento: {e}")
        return {
            'faturamento_bruto': 0,
            'devolucoes': 0,
            'faturamento_liquido': 0,
            'custo_bruto': 0,
            'custo_devolvido': 0,
            'custo_liquido': 0,
            'lucro_liquido': 0
        }

def buscar_faturamento_mensal(conn, ano):
    """Busca faturamento mensal de um ano"""
    data_inicio = f"01/01/{ano}"
    data_fim = f"31/12/{ano}"
    
    query = f"""
    WITH VENDAS_MES AS (
        SELECT 
            TO_CHAR(C.DATA, 'MM') AS MES,
            TO_NUMBER(TO_CHAR(C.DATA, 'MM')) AS MES_NUM,
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
        WHERE C.DATA >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND C.DATA <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY TO_CHAR(C.DATA, 'MM'), TO_NUMBER(TO_CHAR(C.DATA, 'MM'))
    ),
    DEVOLUCOES_MES AS (
        SELECT 
            TO_CHAR(D.DTENT, 'MM') AS MES,
            TO_NUMBER(TO_CHAR(D.DTENT, 'MM')) AS MES_NUM,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND D.DTENT <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
        GROUP BY TO_CHAR(D.DTENT, 'MM'), TO_NUMBER(TO_CHAR(D.DTENT, 'MM'))
    )
    SELECT 
        COALESCE(V.MES_NUM, D.MES_NUM) AS MES,
        NVL(V.VALOR_BRUTO, 0) AS FATURAMENTO_BRUTO,
        NVL(D.VALOR_DEVOLVIDO, 0) AS DEVOLUCOES,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO
    FROM VENDAS_MES V
    FULL OUTER JOIN DEVOLUCOES_MES D ON V.MES = D.MES
    ORDER BY COALESCE(V.MES_NUM, D.MES_NUM)
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        
        resultado = []
        for idx, row in df.iterrows():
            mes_num = int(row['MES']) if pd.notna(row['MES']) else 0
            if 1 <= mes_num <= 12:
                resultado.append({
                    'mes': meses_nomes[mes_num - 1],
                    'mes_num': mes_num,
                    'faturamento_liquido': float(row['FATURAMENTO_LIQUIDO'])
                })
        
        return resultado
    except Exception as e:
        print(f"Erro ao buscar faturamento mensal: {e}")
        return []

def buscar_top_produtos(conn, ano, limite=10):
    """Busca top produtos por faturamento líquido"""
    data_inicio = f"01/01/{ano}"
    data_fim = f"31/12/{ano}"
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            I.CODPROD,
            P.DESCRICAO,
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
        WHERE C.DATA >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND C.DATA <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY I.CODPROD, P.DESCRICAO
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND D.DTENT <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
        GROUP BY D.CODPROD
    )
    SELECT 
        V.CODPROD,
        V.DESCRICAO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
        (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO,
        ((NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - 
         (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0))) AS LUCRO_LIQUIDO
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODPROD = D.CODPROD
    WHERE (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) > 0
    ORDER BY FATURAMENTO_LIQUIDO DESC
    FETCH FIRST {limite} ROWS ONLY
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        resultado = []
        for idx, row in df.iterrows():
            resultado.append({
                'codprod': int(row['CODPROD']),
                'descricao': str(row['DESCRICAO'])[:50],  # Limita tamanho
                'faturamento_liquido': float(row['FATURAMENTO_LIQUIDO']),
                'custo_liquido': float(row['CUSTO_LIQUIDO']),
                'lucro_liquido': float(row['LUCRO_LIQUIDO'])
            })
        return resultado
    except Exception as e:
        print(f"Erro ao buscar top produtos: {e}")
        return []

def buscar_faturamento_departamentos(conn, ano):
    """Busca faturamento por departamento (fornecedor)"""
    data_inicio = f"01/01/{ano}"
    data_fim = f"31/12/{ano}"
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            P.CODEPTO,
            D.DESCRICAO AS DEPARTAMENTO,
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
        LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        WHERE C.DATA >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND C.DATA <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY P.CODEPTO, D.DESCRICAO
    ),
    DEVOLUCOES AS (
        SELECT 
            P.CODEPTO,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE D.DTENT >= TO_DATE('{data_inicio}', 'DD/MM/YYYY')
          AND D.DTENT <= TO_DATE('{data_fim}', 'DD/MM/YYYY')
        GROUP BY P.CODEPTO
    )
    SELECT 
        V.CODEPTO,
        NVL(V.DEPARTAMENTO, 'Sem Departamento') AS DEPARTAMENTO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
        (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO,
        ((NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - 
         (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0))) AS LUCRO_LIQUIDO
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODEPTO = D.CODEPTO
    WHERE (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) > 0
    ORDER BY FATURAMENTO_LIQUIDO DESC
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        resultado = []
        for idx, row in df.iterrows():
            resultado.append({
                'codepto': int(row['CODEPTO']) if pd.notna(row['CODEPTO']) else 0,
                'departamento': str(row['DEPARTAMENTO'])[:50],
                'faturamento_liquido': float(row['FATURAMENTO_LIQUIDO']),
                'custo_liquido': float(row['CUSTO_LIQUIDO']),
                'lucro_liquido': float(row['LUCRO_LIQUIDO'])
            })
        return resultado
    except Exception as e:
        print(f"Erro ao buscar faturamento por departamento: {e}")
        return []

# =========================================================================
# FUNÇÕES DE GRÁFICOS
# =========================================================================
def gerar_grafico_comparativo_mensal(dados_2024, dados_2025):
    """Gera gráfico comparativo mensal"""
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    
    # Criar dicionários para facilitar busca
    dict_2024 = {item['mes']: item['faturamento_liquido'] for item in dados_2024}
    dict_2025 = {item['mes']: item['faturamento_liquido'] for item in dados_2025}
    
    valores_2024 = [dict_2024.get(mes, 0) for mes in meses]
    valores_2025 = [dict_2025.get(mes, 0) for mes in meses]
    
    fig, ax = plt.subplots(figsize=(12, 6))
    x = range(len(meses))
    width = 0.35
    
    bars1 = ax.bar([i - width/2 for i in x], valores_2024, width, label='2024', color='#0049A8', alpha=0.8)
    bars2 = ax.bar([i + width/2 for i in x], valores_2025, width, label='2025', color='#D40000', alpha=0.8)
    
    ax.set_xlabel('Mês', fontsize=11, fontweight='bold')
    ax.set_ylabel('Faturamento Líquido (R$)', fontsize=11, fontweight='bold')
    ax.set_title('Comparativo Mensal de Faturamento - 2024 vs 2025', fontsize=13, fontweight='bold', pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(meses)
    ax.legend()
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Formatar valores no eixo Y
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M' if x >= 1e6 else f'R$ {x/1e3:.0f}K'))
    
    plt.tight_layout()
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('ascii')

def gerar_grafico_crescimento(dados_2024, dados_2025):
    """Gera gráfico de crescimento percentual"""
    total_2024 = sum(item['faturamento_liquido'] for item in dados_2024)
    total_2025 = sum(item['faturamento_liquido'] for item in dados_2025)
    
    if total_2024 == 0:
        crescimento = 0
    else:
        crescimento = ((total_2025 - total_2024) / total_2024) * 100
    
    fig, ax = plt.subplots(figsize=(8, 6))
    
    anos = ['2024', '2025']
    valores = [total_2024, total_2025]
    cores = ['#0049A8', '#D40000']
    
    bars = ax.bar(anos, valores, color=cores, alpha=0.8, width=0.5)
    
    # Adicionar valores nas barras
    for bar, valor in zip(bars, valores):
        altura = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., altura,
                f'R$ {valor/1e6:.2f}M',
                ha='center', va='bottom', fontsize=11, fontweight='bold')
    
    # Adicionar linha de crescimento
    ax.text(0.5, max(valores) * 1.1, 
            f'Crescimento: {crescimento:+.2f}%',
            ha='center', fontsize=12, fontweight='bold',
            bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7))
    
    ax.set_ylabel('Faturamento Líquido (R$)', fontsize=11, fontweight='bold')
    ax.set_title('Comparativo Anual de Faturamento', fontsize=13, fontweight='bold', pad=15)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M'))
    
    plt.tight_layout()
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('ascii')

def gerar_grafico_lucro(dados_2024, dados_2025):
    """Gera gráfico comparativo de lucro"""
    lucro_2024 = dados_2024.get('lucro_liquido', 0)
    lucro_2025 = dados_2025.get('lucro_liquido', 0)
    
    fig, ax = plt.subplots(figsize=(8, 6))
    
    anos = ['2024', '2025']
    valores = [lucro_2024, lucro_2025]
    cores = ['#0049A8', '#D40000']
    
    bars = ax.bar(anos, valores, color=cores, alpha=0.8, width=0.5)
    
    # Adicionar valores nas barras
    for bar, valor in zip(bars, valores):
        altura = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., altura,
                f'R$ {valor/1e6:.2f}M',
                ha='center', va='bottom', fontsize=11, fontweight='bold')
    
    if lucro_2024 != 0:
        crescimento_lucro = ((lucro_2025 - lucro_2024) / abs(lucro_2024)) * 100
        ax.text(0.5, max(valores) * 1.1 if max(valores) > 0 else min(valores) * 1.1,
                f'Variação: {crescimento_lucro:+.2f}%',
                ha='center', fontsize=12, fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7))
    
    ax.set_ylabel('Lucro Líquido (R$)', fontsize=11, fontweight='bold')
    ax.set_title('Comparativo de Lucro - 2024 vs 2025', fontsize=13, fontweight='bold', pad=15)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M'))
    
    plt.tight_layout()
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('ascii')

def gerar_grafico_top_produtos(produtos_2024, produtos_2025, limite=10):
    """Gera gráfico dos top produtos"""
    # Pegar top produtos combinados
    todos_produtos = {}
    for p in produtos_2024[:limite]:
        cod = p['codprod']
        todos_produtos[cod] = {'descricao': p['descricao'], '2024': p['faturamento_liquido'], '2025': 0}
    
    for p in produtos_2025[:limite]:
        cod = p['codprod']
        if cod in todos_produtos:
            todos_produtos[cod]['2025'] = p['faturamento_liquido']
        else:
            todos_produtos[cod] = {'descricao': p['descricao'], '2024': 0, '2025': p['faturamento_liquido']}
    
    # Ordenar por maior valor em 2025
    produtos_ordenados = sorted(todos_produtos.items(), key=lambda x: x[1]['2025'], reverse=True)[:limite]
    
    descricoes = [p[1]['descricao'][:30] + '...' if len(p[1]['descricao']) > 30 else p[1]['descricao'] 
                  for p in produtos_ordenados]
    valores_2024 = [p[1]['2024'] for p in produtos_ordenados]
    valores_2025 = [p[1]['2025'] for p in produtos_ordenados]
    
    fig, ax = plt.subplots(figsize=(12, 8))
    x = range(len(descricoes))
    width = 0.35
    
    bars1 = ax.barh([i - width/2 for i in x], valores_2024, width, label='2024', color='#0049A8', alpha=0.8)
    bars2 = ax.barh([i + width/2 for i in x], valores_2025, width, label='2025', color='#D40000', alpha=0.8)
    
    ax.set_xlabel('Faturamento Líquido (R$)', fontsize=11, fontweight='bold')
    ax.set_ylabel('Produtos', fontsize=11, fontweight='bold')
    ax.set_title(f'Top {limite} Produtos - Comparativo 2024 vs 2025', fontsize=13, fontweight='bold', pad=15)
    ax.set_yticks(x)
    ax.set_yticklabels(descricoes)
    ax.legend()
    ax.grid(axis='x', alpha=0.3, linestyle='--')
    ax.invert_yaxis()
    
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M' if x >= 1e6 else f'R$ {x/1e3:.0f}K'))
    
    plt.tight_layout()
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('ascii')

def gerar_grafico_departamentos(deptos_2024, deptos_2025, limite=10):
    """Gera gráfico dos top departamentos"""
    # Pegar top departamentos combinados
    todos_deptos = {}
    for d in deptos_2024[:limite]:
        cod = d['codepto']
        todos_deptos[cod] = {'departamento': d['departamento'], '2024': d['faturamento_liquido'], '2025': 0}
    
    for d in deptos_2025[:limite]:
        cod = d['codepto']
        if cod in todos_deptos:
            todos_deptos[cod]['2025'] = d['faturamento_liquido']
        else:
            todos_deptos[cod] = {'departamento': d['departamento'], '2024': 0, '2025': d['faturamento_liquido']}
    
    # Ordenar por maior valor em 2025
    deptos_ordenados = sorted(todos_deptos.items(), key=lambda x: x[1]['2025'], reverse=True)[:limite]
    
    nomes = [d[1]['departamento'][:30] + '...' if len(d[1]['departamento']) > 30 else d[1]['departamento'] 
             for d in deptos_ordenados]
    valores_2024 = [d[1]['2024'] for d in deptos_ordenados]
    valores_2025 = [d[1]['2025'] for d in deptos_ordenados]
    
    fig, ax = plt.subplots(figsize=(12, 8))
    x = range(len(nomes))
    width = 0.35
    
    bars1 = ax.barh([i - width/2 for i in x], valores_2024, width, label='2024', color='#0049A8', alpha=0.8)
    bars2 = ax.barh([i + width/2 for i in x], valores_2025, width, label='2025', color='#D40000', alpha=0.8)
    
    ax.set_xlabel('Faturamento Líquido (R$)', fontsize=11, fontweight='bold')
    ax.set_ylabel('Departamentos', fontsize=11, fontweight='bold')
    ax.set_title(f'Top {limite} Departamentos - Comparativo 2024 vs 2025', fontsize=13, fontweight='bold', pad=15)
    ax.set_yticks(x)
    ax.set_yticklabels(nomes)
    ax.legend()
    ax.grid(axis='x', alpha=0.3, linestyle='--')
    ax.invert_yaxis()
    
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M' if x >= 1e6 else f'R$ {x/1e3:.0f}K'))
    
    plt.tight_layout()
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('ascii')

# =========================================================================
# ROTAS
# =========================================================================
@app.route('/')
def index():
    """Página principal do dashboard"""
    conn = get_db_connection()
    if not conn:
        return render_template('erro.html', mensagem='Erro ao conectar ao banco de dados')
    
    try:
        # Buscar dados de 2024
        dados_2024 = buscar_faturamento_ano(conn, 2024)
        mensal_2024 = buscar_faturamento_mensal(conn, 2024)
        produtos_2024 = buscar_top_produtos(conn, 2024, 10)
        deptos_2024 = buscar_faturamento_departamentos(conn, 2024)
        
        # Buscar dados de 2025
        dados_2025 = buscar_faturamento_ano(conn, 2025)
        mensal_2025 = buscar_faturamento_mensal(conn, 2025)
        produtos_2025 = buscar_top_produtos(conn, 2025, 10)
        deptos_2025 = buscar_faturamento_departamentos(conn, 2025)
        
        # Calcular crescimento
        fat_2024 = dados_2024['faturamento_liquido']
        fat_2025 = dados_2025['faturamento_liquido']
        crescimento = ((fat_2025 - fat_2024) / fat_2024 * 100) if fat_2024 > 0 else 0
        
        # Gerar gráficos
        grafico_mensal = gerar_grafico_comparativo_mensal(mensal_2024, mensal_2025)
        grafico_crescimento = gerar_grafico_crescimento(mensal_2024, mensal_2025)
        grafico_lucro = gerar_grafico_lucro(dados_2024, dados_2025)
        grafico_produtos = gerar_grafico_top_produtos(produtos_2024, produtos_2025, 10)
        grafico_deptos = gerar_grafico_departamentos(deptos_2024[:15], deptos_2025[:15], 10)
        
        return render_template('dashboard_crescimento.html',
                             dados_2024=dados_2024,
                             dados_2025=dados_2025,
                             mensal_2024=mensal_2024,
                             mensal_2025=mensal_2025,
                             produtos_2024=produtos_2024[:10],
                             produtos_2025=produtos_2025[:10],
                             deptos_2024=deptos_2024[:15],
                             deptos_2025=deptos_2025[:15],
                             crescimento=crescimento,
                             grafico_mensal=grafico_mensal,
                             grafico_crescimento=grafico_crescimento,
                             grafico_lucro=grafico_lucro,
                             grafico_produtos=grafico_produtos,
                             grafico_deptos=grafico_deptos,
                             data_geracao=datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
    except Exception as e:
        print(f"Erro ao gerar dashboard: {e}")
        import traceback
        traceback.print_exc()
        return render_template('erro.html', mensagem=f'Erro ao gerar dashboard: {str(e)}')
    finally:
        conn.close()

if __name__ == '__main__':
    print("=" * 80)
    print("📊 DASHBOARD DE ANÁLISE DE CRESCIMENTO - 2024 vs 2025")
    print("=" * 80)
    print("🌐 Acesse: http://localhost:3235")
    print("=" * 80)
    app.run(host='0.0.0.0', port=3235, debug=True)

