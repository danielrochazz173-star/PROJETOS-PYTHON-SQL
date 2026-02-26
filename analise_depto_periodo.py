"""
Sistema de Análise por Departamento com Período Selecionável
Permite selecionar um departamento e um período (ex: últimos 6 meses)
e gera análise completa: compras, vendas, vendas abaixo do custo, margem, etc.
Gera PDF paginado com gráficos usando WeasyPrint
"""

import oracledb
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Backend não-interativo para gerar imagens
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import base64
import io
import os

try:
    from weasyprint import HTML, CSS
    WEASYPRINT_AVAILABLE = True
except (ImportError, OSError) as e:
    WEASYPRINT_AVAILABLE = False
    print(f"AVISO: WeasyPrint não disponível ({e}). Usando ReportLab como alternativa.")

# ReportLab como alternativa (já está no projeto)
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    print("AVISO: ReportLab não está instalado. Instale com: pip install reportlab")

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# ======= CONFIGURAÇÃO DE ACESSO ORACLE =======
DB_CONFIG = dict(
    host='10.0.0.10',
    port=1521,
    service='PROD',
    user='powerbi',
    password='cbjc4xp3nlq6'
)

def get_db_connection():
    dsn = f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['service']}"
    try:
        conn = oracledb.connect(user=DB_CONFIG['user'], password=DB_CONFIG['password'], dsn=dsn)
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco Oracle: {e}")
        sys.exit(1)

def get_mes_ano_formatado(mes, ano):
    """Retorna formato NOV/25 para o título"""
    meses_abrev = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                   'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    return f"{meses_abrev[mes-1]}/{str(ano)[-2:]}"

def get_nome_mes(mes):
    """Retorna nome completo do mês"""
    meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                   'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    return meses_nomes[mes - 1]

def buscar_departamentos(conn):
    """Busca lista de departamentos do banco"""
    query = """
    SELECT 
        CODEPTO,
        DESCRICAO
    FROM PCDEPTO
    WHERE DESCRICAO IS NOT NULL
    ORDER BY DESCRICAO
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

def analisar_depto_periodo(conn, codepto, meses_periodo):
    """
    Analisa departamento para os últimos N meses
    Retorna DataFrame com análise mensal: compra, venda, vendas abaixo do custo, margem, etc.
    """
    data_fim = datetime.now()
    data_inicio = data_fim - relativedelta(months=meses_periodo)
    
    # Gerar lista de meses para análise
    meses_analise = []
    data_atual = data_inicio.replace(day=1)
    while data_atual < data_fim:
        meses_analise.append((data_atual.month, data_atual.year))
        if data_atual.month == 12:
            data_atual = data_atual.replace(year=data_atual.year + 1, month=1)
        else:
            data_atual = data_atual.replace(month=data_atual.month + 1)
    
    resultados = []
    
    for mes, ano in meses_analise:
        data_inicio_mes = datetime(ano, mes, 1)
        if mes == 12:
            data_fim_mes = datetime(ano + 1, 1, 1)
        else:
            data_fim_mes = datetime(ano, mes + 1, 1)
        
        # Buscar compras do mês
        query_compras = """
        SELECT 
            SUM(NVL(M.QT, 0) * NVL(M.PUNIT, 0)) AS VALOR_COMPRA
        FROM PCNFENT N
        JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
        JOIN PCPRODUT P ON M.CODPROD = P.CODPROD
        WHERE N.DTENT >= :data_inicio
          AND N.DTENT < :data_fim
          AND P.CODEPTO = :codepto
        """
        
        # Buscar vendas do mês (com custo)
        query_vendas = """
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
                                ROUND(NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS CUSTO_BRUTO,
                SUM(
                    CASE 
                        WHEN NVL(I.BONIFIC, 'N') = 'N' 
                         AND I.PVENDA < NVL(I.VLCUSTOFIN, 0) THEN
                            DECODE(C.CONDVENDA, 
                                5, 0, 6, 0, 11, 0, 12, 0, 
                                ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS VALOR_ABAIXO_CUSTO,
                SUM(
                    CASE 
                        WHEN NVL(I.BONIFIC, 'N') = 'N' 
                         AND I.PVENDA < NVL(I.VLCUSTOFIN, 0) THEN
                            DECODE(C.CONDVENDA, 
                                5, 0, 6, 0, 11, 0, 12, 0, 
                                ROUND((NVL(I.VLCUSTOFIN, 0) - NVL(I.PVENDA, 0)) * NVL(I.QT, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS PREJUIZO_ABAIXO_CUSTO
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
                SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
                SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
            FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
            JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
            WHERE D.DTENT >= :data_inicio
              AND D.DTENT < :data_fim
              AND P.CODEPTO = :codepto
        )
        SELECT 
            NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0) AS VALOR_VENDA,
            NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0) AS CUSTO_VENDA,
            NVL(V.VALOR_ABAIXO_CUSTO, 0) AS VALOR_ABAIXO_CUSTO,
            NVL(V.PREJUIZO_ABAIXO_CUSTO, 0) AS PREJUIZO_ABAIXO_CUSTO
        FROM VENDAS_VALIDAS V
        CROSS JOIN DEVOLUCOES D
        """
        
        try:
            df_compra = pd.read_sql(query_compras, conn, 
                                   params={'data_inicio': data_inicio_mes, 'data_fim': data_fim_mes, 'codepto': codepto})
            df_venda = pd.read_sql(query_vendas, conn, 
                                  params={'data_inicio': data_inicio_mes, 'data_fim': data_fim_mes, 'codepto': codepto})
            
            valor_compra = float(df_compra['VALOR_COMPRA'].iloc[0] if not df_compra.empty and pd.notna(df_compra['VALOR_COMPRA'].iloc[0]) else 0)
            valor_venda = float(df_venda['VALOR_VENDA'].iloc[0] if not df_venda.empty and pd.notna(df_venda['VALOR_VENDA'].iloc[0]) else 0)
            custo_venda = float(df_venda['CUSTO_VENDA'].iloc[0] if not df_venda.empty and pd.notna(df_venda['CUSTO_VENDA'].iloc[0]) else 0)
            valor_abaixo_custo = float(df_venda['VALOR_ABAIXO_CUSTO'].iloc[0] if not df_venda.empty and pd.notna(df_venda['VALOR_ABAIXO_CUSTO'].iloc[0]) else 0)
            prejuizo_abaixo_custo = float(df_venda['PREJUIZO_ABAIXO_CUSTO'].iloc[0] if not df_venda.empty and pd.notna(df_venda['PREJUIZO_ABAIXO_CUSTO'].iloc[0]) else 0)
            
            # Calcular margem
            if valor_venda > 0:
                margem_bruta = valor_venda - custo_venda
                margem_percentual = (margem_bruta / valor_venda) * 100
            else:
                margem_bruta = 0
                margem_percentual = 0
            
            resultados.append({
                'mes': mes,
                'ano': ano,
                'nome_mes': get_nome_mes(mes),
                'mes_formatado': get_mes_ano_formatado(mes, ano),
                'valor_compra': valor_compra,
                'valor_venda': valor_venda,
                'custo_venda': custo_venda,
                'margem_bruta': margem_bruta,
                'margem_percentual': margem_percentual,
                'valor_abaixo_custo': valor_abaixo_custo,
                'prejuizo_abaixo_custo': prejuizo_abaixo_custo
            })
            
        except Exception as e:
            print(f"Erro ao analisar mês {mes}/{ano}: {e}")
            import traceback
            traceback.print_exc()
    
    return pd.DataFrame(resultados)

def analisar_produtos_depto(conn, codepto, meses_periodo):
    """
    Analisa produtos do departamento para os últimos N meses
    Retorna DataFrame com análise mensal por produto: compra, venda, margem, etc.
    """
    data_fim = datetime.now()
    data_inicio = data_fim - relativedelta(months=meses_periodo)
    
    # Gerar lista de meses para análise
    meses_analise = []
    data_atual = data_inicio.replace(day=1)
    while data_atual < data_fim:
        meses_analise.append((data_atual.month, data_atual.year))
        if data_atual.month == 12:
            data_atual = data_atual.replace(year=data_atual.year + 1, month=1)
        else:
            data_atual = data_atual.replace(month=data_atual.month + 1)
    
    resultados = []
    
    # Buscar informações dos produtos do departamento uma vez
    query_produtos_info = """
    SELECT CODPROD, DESCRICAO, EMBALAGEM
    FROM PCPRODUT
    WHERE CODEPTO = :codepto
    """
    df_produtos_info = pd.read_sql(query_produtos_info, conn, params={'codepto': codepto})
    df_produtos_info.columns = [x.lower() for x in df_produtos_info.columns]
    dict_produtos = df_produtos_info.set_index('codprod').to_dict('index')
    
    for mes, ano in meses_analise:
        data_inicio_mes = datetime(ano, mes, 1)
        if mes == 12:
            data_fim_mes = datetime(ano + 1, 1, 1)
        else:
            data_fim_mes = datetime(ano, mes + 1, 1)
        
        # Buscar compras por produto no mês
        query_compras = """
        SELECT 
            P.CODPROD,
            SUM(NVL(M.QT, 0) * NVL(M.PUNIT, 0)) AS VALOR_COMPRA,
            SUM(NVL(M.QT, 0)) AS QT_COMPRA
        FROM PCNFENT N
        JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
        JOIN PCPRODUT P ON M.CODPROD = P.CODPROD
        WHERE N.DTENT >= :data_inicio
          AND N.DTENT < :data_fim
          AND P.CODEPTO = :codepto
        GROUP BY P.CODPROD
        """
        
        # Buscar vendas por produto no mês
        query_vendas = """
        WITH VENDAS_VALIDAS AS (
            SELECT 
                P.CODPROD,
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
                                ROUND(NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS CUSTO_BRUTO,
                SUM(
                    CASE 
                        WHEN NVL(I.BONIFIC, 'N') = 'N' 
                         AND I.PVENDA < NVL(I.VLCUSTOFIN, 0) THEN
                            DECODE(C.CONDVENDA, 
                                5, 0, 6, 0, 11, 0, 12, 0, 
                                ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS VALOR_ABAIXO_CUSTO,
                SUM(
                    CASE 
                        WHEN NVL(I.BONIFIC, 'N') = 'N' 
                         AND I.PVENDA < NVL(I.VLCUSTOFIN, 0) THEN
                            DECODE(C.CONDVENDA, 
                                5, 0, 6, 0, 11, 0, 12, 0, 
                                ROUND((NVL(I.VLCUSTOFIN, 0) - NVL(I.PVENDA, 0)) * NVL(I.QT, 0), 2)
                            )
                        ELSE 0 
                    END
                ) AS PREJUIZO_ABAIXO_CUSTO
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
            GROUP BY P.CODPROD
        ),
        DEVOLUCOES AS (
            SELECT 
                P.CODPROD,
                SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
                SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
            FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
            JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
            WHERE D.DTENT >= :data_inicio
              AND D.DTENT < :data_fim
              AND P.CODEPTO = :codepto
            GROUP BY P.CODPROD
        )
        SELECT 
            V.CODPROD,
            NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0) AS VALOR_VENDA,
            NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0) AS CUSTO_VENDA,
            NVL(V.VALOR_ABAIXO_CUSTO, 0) AS VALOR_ABAIXO_CUSTO,
            NVL(V.PREJUIZO_ABAIXO_CUSTO, 0) AS PREJUIZO_ABAIXO_CUSTO
        FROM VENDAS_VALIDAS V
        LEFT JOIN DEVOLUCOES D ON V.CODPROD = D.CODPROD
        """
        
        try:
            df_compras = pd.read_sql(query_compras, conn, 
                                    params={'data_inicio': data_inicio_mes, 'data_fim': data_fim_mes, 'codepto': codepto})
            df_vendas = pd.read_sql(query_vendas, conn, 
                                   params={'data_inicio': data_inicio_mes, 'data_fim': data_fim_mes, 'codepto': codepto})
            
            # Merge dos dados
            df_merge = pd.merge(
                df_compras,
                df_vendas,
                on='CODPROD',
                how='outer'
            )
            
            # Preencher valores nulos
            df_merge['VALOR_COMPRA'] = df_merge['VALOR_COMPRA'].fillna(0)
            df_merge['QT_COMPRA'] = df_merge['QT_COMPRA'].fillna(0)
            df_merge['VALOR_VENDA'] = df_merge['VALOR_VENDA'].fillna(0)
            df_merge['CUSTO_VENDA'] = df_merge['CUSTO_VENDA'].fillna(0)
            df_merge['VALOR_ABAIXO_CUSTO'] = df_merge['VALOR_ABAIXO_CUSTO'].fillna(0)
            df_merge['PREJUIZO_ABAIXO_CUSTO'] = df_merge['PREJUIZO_ABAIXO_CUSTO'].fillna(0)
            
            # Adicionar informações do produto e mês
            for _, row in df_merge.iterrows():
                codprod = int(row['CODPROD']) if pd.notna(row['CODPROD']) else None
                if codprod and codprod in dict_produtos:
                    descricao = dict_produtos[codprod].get('descricao', 'Produto sem descrição')
                    embalagem = dict_produtos[codprod].get('embalagem', '')
                else:
                    descricao = 'Produto sem descrição'
                    embalagem = ''
                
                # Calcular margem
                valor_venda = float(row['VALOR_VENDA'])
                custo_venda = float(row['CUSTO_VENDA'])
                margem_bruta = valor_venda - custo_venda
                margem_percentual = (margem_bruta / valor_venda * 100) if valor_venda > 0 else 0
                
                resultados.append({
                    'codprod': codprod,
                    'descricao': descricao,
                    'embalagem': embalagem,
                    'mes': mes,
                    'ano': ano,
                    'nome_mes': get_nome_mes(mes),
                    'mes_formatado': get_mes_ano_formatado(mes, ano),
                    'valor_compra': float(row['VALOR_COMPRA']),
                    'valor_venda': valor_venda,
                    'custo_venda': custo_venda,
                    'margem_bruta': margem_bruta,
                    'margem_percentual': margem_percentual,
                    'valor_abaixo_custo': float(row['VALOR_ABAIXO_CUSTO']),
                    'prejuizo_abaixo_custo': float(row['PREJUIZO_ABAIXO_CUSTO'])
                })
                
        except Exception as e:
            print(f"Erro ao analisar produtos do mês {mes}/{ano}: {e}")
            import traceback
            traceback.print_exc()
    
    df_resultado = pd.DataFrame(resultados)
    
    # Se não houver resultados, retornar DataFrame vazio
    if df_resultado.empty:
        return df_resultado
    
    # Converter nomes das colunas para minúsculas
    df_resultado.columns = [x.lower() for x in df_resultado.columns]
    
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

def formatar_percentual(valor):
    """Formata percentual"""
    try:
        return f"{float(valor):.2f}%".replace('.', ',')
    except:
        return "0,00%"

def criar_grafico_compras_vendas_margem(df_resultado):
    """Cria gráfico de barras mostrando compras, vendas e margem por mês"""
    # Preparar dados
    df_resultado = df_resultado.sort_values(['ano', 'mes'])
    meses_labels = [f"{row['mes_formatado']}" for _, row in df_resultado.iterrows()]
    compras = df_resultado['valor_compra'].values
    vendas = df_resultado['valor_venda'].values
    margens = df_resultado['margem_percentual'].values  # Usar margem percentual
    
    # Criar figura
    fig, ax = plt.subplots(figsize=(12, 6))
    
    x = range(len(meses_labels))
    width = 0.35
    
    # Criar barras
    bars1 = ax.bar([i - width/2 for i in x], compras, width, label='Compra', color='#1f77b4', alpha=0.8)
    bars2 = ax.bar([i + width/2 for i in x], vendas, width, label='Venda', color='#2ca02c', alpha=0.8)
    
    # Adicionar linha de margem (em percentual) - cor vermelha para melhor visibilidade na impressão
    ax2 = ax.twinx()
    line = ax2.plot(x, margens, color='#d32f2f', marker='o', linewidth=2.5, markersize=8, label='Margem %')
    ax2.set_ylabel('Margem (%)', color='#d32f2f', fontsize=10, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#d32f2f')
    
    # Adicionar valores nas barras
    for i, (bar1, bar2, margem) in enumerate(zip(bars1, bars2, margens)):
        height1 = bar1.get_height()
        height2 = bar2.get_height()
        if height1 > 0:
            ax.text(bar1.get_x() + bar1.get_width()/2., height1,
                   formatar_valor(height1), ha='center', va='bottom', fontsize=8, rotation=90)
        if height2 > 0:
            ax.text(bar2.get_x() + bar2.get_width()/2., height2,
                   formatar_valor(height2), ha='center', va='bottom', fontsize=8, rotation=90)
        # Adicionar margem percentual acima da barra de venda - cor vermelha
        if margem != 0:
            ax2.text(i, margem, formatar_percentual(margem), ha='center', 
                    va='bottom' if margem > 0 else 'top', fontsize=8, 
                    color='#d32f2f', fontweight='bold')
    
    # Configurar eixos
    ax.set_xlabel('Mês', fontsize=11, fontweight='bold')
    ax.set_ylabel('Valor (R$)', fontsize=11, fontweight='bold')
    ax.set_title('Compras vs Vendas com Margem por Mês', fontsize=14, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(meses_labels, rotation=45, ha='right')
    ax.legend(loc='upper left', fontsize=10)
    ax2.legend(loc='upper right', fontsize=10)
    ax.grid(True, alpha=0.3, axis='y')
    
    plt.tight_layout()
    
    # Se for para WeasyPrint, retornar base64
    # Se for para ReportLab, a função criar_grafico_imagem salva o arquivo
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.read()).decode()
    plt.close()
    
    return image_base64

def gerar_html_pdf(df_resultado, nome_depto, meses_periodo, tipo_relatorio='resumido'):
    """Gera HTML otimizado para PDF com WeasyPrint"""
    titulo = f"ANÁLISE - {nome_depto.upper()}"
    subtitulo = f"Período: Últimos {meses_periodo} meses"
    
    # Calcular totais
    total_compra = df_resultado['valor_compra'].sum()
    total_venda = df_resultado['valor_venda'].sum()
    total_custo = df_resultado['custo_venda'].sum()
    total_margem = df_resultado['margem_bruta'].sum()
    total_abaixo_custo = df_resultado['valor_abaixo_custo'].sum()
    total_prejuizo = df_resultado['prejuizo_abaixo_custo'].sum()
    
    margem_total_percentual = (total_margem / total_venda * 100) if total_venda > 0 else 0
    
    # Criar gráfico
    grafico_base64 = criar_grafico_compras_vendas_margem(df_resultado)
    
    # Caminho da logo
    logo_path = Path(__file__).parent / 'logo-dicon.png'
    logo_base64 = ""
    if logo_path.exists():
        with open(logo_path, 'rb') as f:
            logo_base64 = base64.b64encode(f.read()).decode()
    
    # Gerar HTML
    html_content = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>{titulo}</title>
    <style>
        @page {{
            size: A4;
            margin: 2cm;
            @top-center {{
                content: "{titulo}";
                font-size: 10pt;
                color: #666;
            }}
            @bottom-center {{
                content: "Página " counter(page) " de " counter(pages);
                font-size: 9pt;
                color: #666;
            }}
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Arial', sans-serif;
            font-size: 10pt;
            color: #333;
            line-height: 1.4;
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 20px;
            page-break-after: avoid;
        }}
        
        .header h1 {{
            font-size: 18pt;
            font-weight: bold;
            color: #124377;
            margin-bottom: 5px;
        }}
        
        .header .subtitulo {{
            font-size: 12pt;
            color: #666;
            margin-bottom: 10px;
        }}
        
        .header .data {{
            font-size: 9pt;
            color: #999;
        }}
        
        .logo {{
            max-width: 120px;
            height: auto;
            margin: 0 auto 10px;
        }}
        
        .resumo {{
            display: table;
            width: 100%;
            margin-bottom: 20px;
            page-break-inside: avoid;
        }}
        
        .card-resumo {{
            display: table-cell;
            width: 33.33%;
            padding: 10px;
            border: 1px solid #ddd;
            text-align: center;
            vertical-align: top;
        }}
        
        .card-resumo .label {{
            font-size: 9pt;
            font-weight: bold;
            color: #666;
            margin-bottom: 5px;
            text-transform: uppercase;
        }}
        
        .card-resumo .valor {{
            font-size: 14pt;
            font-weight: bold;
            color: #124377;
        }}
        
        .card-resumo .valor.negativo {{
            color: #d32f2f;
        }}
        
        .card-resumo .valor.positivo {{
            color: #388e3c;
        }}
        
        .grafico-container {{
            text-align: center;
            margin: 20px 0;
            page-break-inside: avoid;
        }}
        
        .grafico-container img {{
            max-width: 100%;
            height: auto;
        }}
        
        .tabela-container {{
            margin-top: 20px;
            page-break-inside: avoid;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 9pt;
            page-break-inside: auto;
        }}
        
        thead {{
            display: table-header-group;
        }}
        
        tbody {{
            display: table-row-group;
        }}
        
        tr {{
            page-break-inside: avoid;
            page-break-after: auto;
        }}
        
        th {{
            background-color: #124377;
            color: white;
            padding: 8px 5px;
            text-align: left;
            font-weight: bold;
            font-size: 9pt;
            border: 1px solid #0d2f5a;
        }}
        
        td {{
            padding: 6px 5px;
            border: 1px solid #ddd;
            text-align: right;
        }}
        
        td:first-child {{
            text-align: left;
            font-weight: bold;
        }}
        
        tbody tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        
        .negativo {{
            color: #d32f2f;
            font-weight: bold;
        }}
        
        .positivo {{
            color: #388e3c;
            font-weight: bold;
        }}
        
        .footer {{
            margin-top: 30px;
            text-align: center;
            font-size: 8pt;
            color: #999;
            page-break-inside: avoid;
        }}
        
        @media print {{
            .page-break {{
                page-break-before: always;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        {"<img src='data:image/png;base64," + logo_base64 + "' class='logo' alt='Logo' />" if logo_base64 else ""}
        <h1>{titulo}</h1>
        <div class="subtitulo">{subtitulo}</div>
        <div class="data">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
    </div>
    
    <div class="resumo">
        <div class="card-resumo">
            <div class="label">Total Comprado</div>
            <div class="valor">{formatar_valor_completo(total_compra)}</div>
        </div>
        <div class="card-resumo">
            <div class="label">Total Vendido</div>
            <div class="valor">{formatar_valor_completo(total_venda)}</div>
        </div>
        <div class="card-resumo">
            <div class="label">Margem Bruta</div>
            <div class="valor {'positivo' if total_margem >= 0 else 'negativo'}">{formatar_valor_completo(total_margem)}</div>
        </div>
    </div>
    
    <div class="resumo">
        <div class="card-resumo">
            <div class="label">Margem %</div>
            <div class="valor {'positivo' if margem_total_percentual >= 0 else 'negativo'}">{formatar_percentual(margem_total_percentual)}</div>
        </div>
        <div class="card-resumo">
            <div class="label">Vendas Abaixo Custo</div>
            <div class="valor negativo">{formatar_valor_completo(total_abaixo_custo)}</div>
        </div>
        <div class="card-resumo">
            <div class="label">Prejuízo</div>
            <div class="valor negativo">{formatar_valor_completo(total_prejuizo)}</div>
        </div>
    </div>
    
    <div class="grafico-container">
        <img src="data:image/png;base64,{grafico_base64}" alt="Gráfico Compras vs Vendas" />
    </div>
"""
    
    if tipo_relatorio == 'completo':
        html_content += """
    <div class="tabela-container">
        <table>
            <thead>
                <tr>
                    <th>Mês</th>
                    <th>Compra</th>
                    <th>Venda</th>
                    <th>Custo</th>
                    <th>Margem Bruta</th>
                    <th>Margem %</th>
                    <th>Vendas Abaixo Custo</th>
                    <th>Prejuízo</th>
                </tr>
            </thead>
            <tbody>
"""
        
        # Adicionar linhas da tabela
        for _, row in df_resultado.iterrows():
            margem_class = 'positivo' if row['margem_bruta'] >= 0 else 'negativo'
            margem_pct_class = 'positivo' if row['margem_percentual'] >= 0 else 'negativo'
            
            html_content += f"""
                <tr>
                    <td><strong>{row['nome_mes']} {row['ano']}</strong></td>
                    <td>{formatar_valor_completo(row['valor_compra'])}</td>
                    <td>{formatar_valor_completo(row['valor_venda'])}</td>
                    <td>{formatar_valor_completo(row['custo_venda'])}</td>
                    <td class="{margem_class}">{formatar_valor_completo(row['margem_bruta'])}</td>
                    <td class="{margem_pct_class}">{formatar_percentual(row['margem_percentual'])}</td>
                    <td class="negativo">{formatar_valor_completo(row['valor_abaixo_custo'])}</td>
                    <td class="negativo">{formatar_valor_completo(row['prejuizo_abaixo_custo'])}</td>
                </tr>
"""
        
        html_content += """
            </tbody>
        </table>
    </div>
"""
    
    html_content += """
    <div class="footer">
        Relatório gerado automaticamente pelo sistema de análise de departamentos
    </div>
</body>
</html>
"""
    
    return html_content

def criar_grafico_imagem(df_resultado, caminho_imagem):
    """Salva gráfico como imagem PNG"""
    # Garantir que o diretório existe
    caminho_path = Path(caminho_imagem)
    caminho_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Preparar dados
    df_resultado = df_resultado.sort_values(['ano', 'mes'])
    meses_labels = [f"{row['mes_formatado']}" for _, row in df_resultado.iterrows()]
    compras = df_resultado['valor_compra'].values
    vendas = df_resultado['valor_venda'].values
    margens = df_resultado['margem_percentual'].values  # Usar margem percentual
    
    # Criar figura
    fig, ax = plt.subplots(figsize=(12, 6))
    
    x = range(len(meses_labels))
    width = 0.35
    
    # Criar barras
    bars1 = ax.bar([i - width/2 for i in x], compras, width, label='Compra', color='#1f77b4', alpha=0.8)
    bars2 = ax.bar([i + width/2 for i in x], vendas, width, label='Venda', color='#2ca02c', alpha=0.8)
    
    # Adicionar linha de margem (em percentual) - cor vermelha para melhor visibilidade na impressão
    ax2 = ax.twinx()
    line = ax2.plot(x, margens, color='#d32f2f', marker='o', linewidth=2.5, markersize=8, label='Margem %')
    ax2.set_ylabel('Margem (%)', color='#d32f2f', fontsize=10, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#d32f2f')
    
    # Adicionar valores nas barras
    for i, (bar1, bar2, margem) in enumerate(zip(bars1, bars2, margens)):
        height1 = bar1.get_height()
        height2 = bar2.get_height()
        if height1 > 0:
            ax.text(bar1.get_x() + bar1.get_width()/2., height1,
                   formatar_valor(height1), ha='center', va='bottom', fontsize=8, rotation=90)
        if height2 > 0:
            ax.text(bar2.get_x() + bar2.get_width()/2., height2,
                   formatar_valor(height2), ha='center', va='bottom', fontsize=8, rotation=90)
        # Adicionar margem percentual acima da barra de venda - cor vermelha
        if margem != 0:
            ax2.text(i, margem, formatar_percentual(margem), ha='center', 
                    va='bottom' if margem > 0 else 'top', fontsize=8, 
                    color='#d32f2f', fontweight='bold')
    
    # Configurar eixos
    ax.set_xlabel('Mês', fontsize=11, fontweight='bold')
    ax.set_ylabel('Valor (R$)', fontsize=11, fontweight='bold')
    ax.set_title('Compras vs Vendas com Margem por Mês', fontsize=14, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(meses_labels, rotation=45, ha='right')
    ax.legend(loc='upper left', fontsize=10)
    ax2.legend(loc='upper right', fontsize=10)
    ax.grid(True, alpha=0.3, axis='y')
    
    plt.tight_layout()
    plt.savefig(str(caminho_imagem), format='png', dpi=150, bbox_inches='tight')
    plt.close()
    
    # Verificar se o arquivo foi criado
    if not Path(caminho_imagem).exists():
        raise Exception(f"Erro: Arquivo de gráfico não foi criado em {caminho_imagem}")

def gerar_pdf_reportlab(df_resultado, df_produtos, nome_depto, meses_periodo, tipo_relatorio, nome_arquivo):
    """Gera PDF usando ReportLab (alternativa compatível com Windows)"""
    if not REPORTLAB_AVAILABLE:
        raise Exception("ReportLab não está instalado. Instale com: pip install reportlab")
    
    pdf_path = Path(nome_arquivo)
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4,
                           rightMargin=1.2*cm, leftMargin=1.2*cm,
                           topMargin=1.5*cm, bottomMargin=1.2*cm)  # Reduzir margens para aproveitar espaço
    
    story = []
    styles = getSampleStyleSheet()
    
    # Calcular totais
    total_compra = df_resultado['valor_compra'].sum()
    total_venda = df_resultado['valor_venda'].sum()
    total_custo = df_resultado['custo_venda'].sum()
    total_margem = df_resultado['margem_bruta'].sum()
    total_abaixo_custo = df_resultado['valor_abaixo_custo'].sum()
    total_prejuizo = df_resultado['prejuizo_abaixo_custo'].sum()
    margem_total_percentual = (total_margem / total_venda * 100) if total_venda > 0 else 0
    
    # Título
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#124377'),
        spaceAfter=10,
        alignment=1  # Centralizado
    )
    
    story.append(Paragraph(f"<b>ANÁLISE - {nome_depto.upper()}</b>", titulo_style))
    story.append(Paragraph(f"Período: Últimos {meses_periodo} meses", styles['Normal']))
    story.append(Paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", 
                          ParagraphStyle('Data', parent=styles['Normal'], fontSize=8, textColor=colors.grey)))
    story.append(Spacer(1, 0.3*cm))
    
    # Cards de resumo
    resumo_data = [
        ['Total Comprado', 'Total Vendido', 'Margem Bruta'],
        [formatar_valor_completo(total_compra), 
         formatar_valor_completo(total_venda),
         formatar_valor_completo(total_margem)],
        ['Margem %', 'Vendas Abaixo Custo', 'Prejuízo'],
        [formatar_percentual(margem_total_percentual),
         formatar_valor_completo(total_abaixo_custo),
         formatar_valor_completo(total_prejuizo)]
    ]
    
    resumo_table = Table(resumo_data, colWidths=[6*cm, 6*cm, 6*cm])
    resumo_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124377')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 2), (-1, 2), 9),
        ('FONTSIZE', (0, 3), (-1, 3), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, 1), [colors.white]),
        ('ROWBACKGROUNDS', (0, 3), (-1, 3), [colors.lightgrey]),
    ]))
    
    story.append(resumo_table)
    story.append(Spacer(1, 0.3*cm))
    
    # Gráfico - usar caminho absoluto no diretório do PDF (mesmo diretório)
    grafico_path = pdf_path.parent / f'temp_grafico_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
    grafico_path = grafico_path.resolve()  # Garantir caminho absoluto
    
    try:
        criar_grafico_imagem(df_resultado, str(grafico_path))
        
        # Pequeno delay para garantir que o arquivo foi escrito
        import time
        time.sleep(0.1)
        
        # Verificar novamente se existe antes de adicionar ao PDF
        if grafico_path.exists() and grafico_path.stat().st_size > 0:
            # Usar caminho absoluto como string
            img = Image(str(grafico_path), width=16*cm, height=7*cm)  # Reduzir altura do gráfico
            story.append(img)
            story.append(Spacer(1, 0.3*cm))
        else:
            print(f"Aviso: Gráfico não foi criado ou está vazio: {grafico_path}")
    except Exception as e:
        print(f"Erro ao criar gráfico: {e}")
        import traceback
        traceback.print_exc()
        # Continuar sem gráfico se houver erro
    
    # Tabela detalhada (se completo) - sem quebra de página, aproveitar espaço
    if tipo_relatorio == 'completo':
        story.append(Spacer(1, 0.6*cm))  # Aumentar espaçamento antes da tabela detalhada
        story.append(Paragraph("<b>Tabela Detalhada por Mês</b>", styles['Heading2']))
        story.append(Spacer(1, 0.3*cm))
        
        # Preparar dados da tabela
        tabela_data = [['Mês', 'Compra', 'Venda', 'Custo', 'Margem Bruta', 'Margem %', 'Vendas Abaixo Custo', 'Prejuízo']]
        
        for _, row in df_resultado.iterrows():
            tabela_data.append([
                f"{row['nome_mes']} {row['ano']}",
                formatar_valor_completo(row['valor_compra']),
                formatar_valor_completo(row['valor_venda']),
                formatar_valor_completo(row['custo_venda']),
                formatar_valor_completo(row['margem_bruta']),
                formatar_percentual(row['margem_percentual']),
                formatar_valor_completo(row['valor_abaixo_custo']),
                formatar_valor_completo(row['prejuizo_abaixo_custo'])
            ])
        
        # Calcular larguras das colunas
        col_widths = [3*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2.5*cm, 2.5*cm]
        
        tabela = Table(tabela_data, colWidths=col_widths, repeatRows=1)
        tabela_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124377')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]

        # Cores para margens: negativo vermelho, abaixo de 6% em âmbar
        cor_negativa = colors.HexColor('#d32f2f')
        cor_baixa = colors.HexColor('#b45309')

        for idx, row in df_resultado.reset_index(drop=True).iterrows():
            linha = idx + 1  # tabela começa na linha 1 (0 é cabeçalho)
            if row['margem_bruta'] < 0:
                tabela_style.append(('TEXTCOLOR', (4, linha), (4, linha), cor_negativa))
            if row['margem_percentual'] < 0:
                tabela_style.append(('TEXTCOLOR', (5, linha), (5, linha), cor_negativa))
            elif row['margem_percentual'] < 6:
                tabela_style.append(('TEXTCOLOR', (5, linha), (5, linha), cor_baixa))

        tabela.setStyle(TableStyle(tabela_style))
        
        story.append(tabela)
        story.append(Spacer(1, 0.4*cm))  # Espaçamento após tabela detalhada
    
    # Análise por Produto - sem quebra de página, aproveitar espaço
    if not df_produtos.empty:
        story.append(Spacer(1, 0.8*cm))  # Aumentar espaçamento antes da seção de produtos
        story.append(Paragraph("<b>Análise por Produto (Mensal)</b>", styles['Heading2']))
        story.append(Spacer(1, 0.4*cm))
        
        # Agrupar por produto e calcular totais, depois ordenar
        df_produtos_totais = df_produtos.groupby(['codprod', 'descricao', 'embalagem']).agg({
            'valor_compra': 'sum',
            'valor_venda': 'sum',
            'custo_venda': 'sum',
            'margem_bruta': 'sum',
            'valor_abaixo_custo': 'sum'
        }).reset_index()
        
        # Recalcular margem percentual total
        df_produtos_totais['margem_percentual'] = (df_produtos_totais['margem_bruta'] / df_produtos_totais['valor_venda'] * 100).fillna(0)
        
        # Ordenar por valor de venda (maior primeiro) e pegar top 30
        df_produtos_totais = df_produtos_totais.sort_values('valor_venda', ascending=False).head(30)
        
        # Contador para limitar produtos
        contador_produtos = 0
        total_produtos = len(df_produtos.groupby('codprod'))
        
        # Para cada produto top, mostrar detalhamento mensal
        for idx, produto in df_produtos_totais.iterrows():
            codprod = produto['codprod']
            descricao = str(produto['descricao'])
            embalagem = str(produto['embalagem']) if pd.notna(produto['embalagem']) else ''
            
            # Buscar dados mensais deste produto
            df_produto_mensal = df_produtos[df_produtos['codprod'] == codprod].sort_values(['ano', 'mes'])
            
            if df_produto_mensal.empty:
                continue
            
            # Espaçamento ANTES do bloco do produto (maior para evitar sobreposição)
            if contador_produtos == 0:
                story.append(Spacer(1, 0.7*cm))  # Primeiro produto: mais espaço
            else:
                story.append(Spacer(1, 1.0*cm))  # Produtos seguintes: espaço maior
            
            # Montar bloco do produto em conjunto (título + tabela)
            bloco_produto = []
            produto_titulo = f"{int(codprod)} - {descricao[:40]}{'...' if len(descricao) > 40 else ''}"
            if embalagem:
                produto_titulo += f" ({embalagem})"
            bloco_produto.append(Paragraph(
                f"<b>{produto_titulo}</b>",
                ParagraphStyle(
                    'ProdutoTitulo',
                    parent=styles['Normal'],
                    fontSize=9,
                    textColor=colors.HexColor('#124377'),
                    alignment=1,
                    spaceAfter=6
                )
            ))
            bloco_produto.append(Spacer(1, 0.25*cm))  # Espaço entre título e tabela
            
            # Tabela mensal do produto
            tabela_produto_data = [['Mês', 'Compra', 'Venda', 'Custo', 'Margem', 'Margem %', 'Abaixo Custo']]
            
            for _, row_mes in df_produto_mensal.iterrows():
                tabela_produto_data.append([
                    row_mes['mes_formatado'],
                    formatar_valor_completo(row_mes['valor_compra']),
                    formatar_valor_completo(row_mes['valor_venda']),
                    formatar_valor_completo(row_mes['custo_venda']),
                    formatar_valor_completo(row_mes['margem_bruta']),
                    formatar_percentual(row_mes['margem_percentual']),
                    formatar_valor_completo(row_mes['valor_abaixo_custo'])
                ])
            
            # Adicionar linha de totais
            tabela_produto_data.append([
                '<b>TOTAL</b>',
                formatar_valor_completo(produto['valor_compra']),
                formatar_valor_completo(produto['valor_venda']),
                formatar_valor_completo(produto['custo_venda']),
                formatar_valor_completo(produto['margem_bruta']),
                formatar_percentual(produto['margem_percentual']),
                formatar_valor_completo(produto['valor_abaixo_custo'])
            ])
            
            # Larguras das colunas (mais largas, com folga)
            col_widths_produto = [3*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2.5*cm]
            
            tabela_produto = Table(tabela_produto_data, colWidths=col_widths_produto, repeatRows=1)
            tabela_style_prod = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124377')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),  # Mesmo tamanho do detalhado
                ('FONTSIZE', (0, 1), (-1, -2), 7),  # Mesmo tamanho do detalhado
                ('FONTSIZE', (0, -1), (-1, -1), 7),  # Total igual ao corpo
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),  # Linha de total em negrito
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),  # Linha de total com fundo
                ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.lightgrey]),
                ('TOPPADDING', (0, -1), (-1, -1), 6),  # Padding superior na linha de total
                ('BOTTOMPADDING', (0, -1), (-1, -1), 6),  # Padding inferior na linha de total
            ]

            # Cores para margens: negativo vermelho, abaixo de 6% em âmbar
            cor_negativa = colors.HexColor('#d32f2f')
            cor_baixa = colors.HexColor('#b45309')

            # Linhas mensais
            for idx_linha, row_mes in df_produto_mensal.reset_index(drop=True).iterrows():
                linha = idx_linha + 1  # cabeçalho é 0
                if row_mes['margem_bruta'] < 0:
                    tabela_style_prod.append(('TEXTCOLOR', (4, linha), (4, linha), cor_negativa))
                if row_mes['margem_percentual'] < 0:
                    tabela_style_prod.append(('TEXTCOLOR', (5, linha), (5, linha), cor_negativa))
                elif row_mes['margem_percentual'] < 6:
                    tabela_style_prod.append(('TEXTCOLOR', (5, linha), (5, linha), cor_baixa))

            # Linha TOTAL (última)
            linha_total = len(tabela_produto_data) - 1
            if produto['margem_bruta'] < 0:
                tabela_style_prod.append(('TEXTCOLOR', (4, linha_total), (4, linha_total), cor_negativa))
            if produto['margem_percentual'] < 0:
                tabela_style_prod.append(('TEXTCOLOR', (5, linha_total), (5, linha_total), cor_negativa))
            elif produto['margem_percentual'] < 6:
                tabela_style_prod.append(('TEXTCOLOR', (5, linha_total), (5, linha_total), cor_baixa))

            tabela_produto.setStyle(TableStyle(tabela_style_prod))
            
            bloco_produto.append(tabela_produto)
            bloco_produto.append(Spacer(1, 0.5*cm))  # Espaçamento após a tabela
            story.append(KeepTogether(bloco_produto))
            
            contador_produtos += 1
            
            # Limitar a 30 produtos para não sobrecarregar o PDF
            if contador_produtos >= 30:
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph(f"<i>Mostrando top 30 produtos de {total_produtos} totais (ordenados por valor de venda)</i>", 
                                      ParagraphStyle('Info', parent=styles['Normal'], fontSize=6, textColor=colors.grey)))
                break
    
    # Rodapé
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("Relatório gerado automaticamente pelo sistema de análise de departamentos", 
                          ParagraphStyle('Footer', parent=styles['Normal'], fontSize=7, textColor=colors.grey, alignment=1)))
    
    # Construir PDF
    try:
        doc.build(story)
    finally:
        # Remover arquivo temporário do gráfico após construir o PDF
        if 'grafico_path' in locals() and grafico_path.exists():
            try:
                grafico_path.unlink()
            except:
                pass
    
    return str(pdf_path.resolve())

def gerar_pdf(html_content, nome_arquivo):
    """Gera PDF a partir do HTML usando WeasyPrint"""
    if not WEASYPRINT_AVAILABLE:
        raise Exception("WeasyPrint não está instalado. Instale com: pip install weasyprint")
    
    pdf_path = Path(nome_arquivo)
    HTML(string=html_content).write_pdf(pdf_path)
    return str(pdf_path.resolve())

class InterfaceAnaliseDepto:
    def __init__(self, master):
        self.master = master
        self.master.title("Análise por Departamento - Período Selecionável")
        self.master.geometry("600x500")
        
        main_frame = ttk.Frame(self.master, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Análise Completa por Departamento", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Seleção de departamento
        ttk.Label(main_frame, text="Departamento:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.depto_var = tk.StringVar()
        self.combo_depto = ttk.Combobox(main_frame, textvariable=self.depto_var,
                                        state="readonly", width=40)
        self.combo_depto.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Carregar departamentos
        self.carregar_departamentos()
        
        # Seleção de período (últimos N meses)
        ttk.Label(main_frame, text="Período (últimos meses):").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.meses_var = tk.IntVar(value=6)
        meses_opcoes = [3, 6, 9, 12, 18, 24]
        self.combo_meses = ttk.Combobox(main_frame, textvariable=self.meses_var,
                                        values=meses_opcoes, state="readonly", width=40)
        self.combo_meses.current(1)  # 6 meses por padrão
        self.combo_meses.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=10)
        
        # Seleção de tipo de relatório
        ttk.Label(main_frame, text="Tipo de Relatório:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.tipo_relatorio_var = tk.StringVar(value="resumido")
        frame_tipo = ttk.Frame(main_frame)
        frame_tipo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Radiobutton(frame_tipo, text="Resumido (Gráficos + Resumo)", 
                        variable=self.tipo_relatorio_var, value="resumido").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(frame_tipo, text="Completo (Gráficos + Tabela Detalhada)", 
                        variable=self.tipo_relatorio_var, value="completo").pack(side=tk.LEFT, padx=5)
        
        # Botão gerar
        btn_gerar = ttk.Button(main_frame, text="Gerar PDF", command=self.gerar_analise)
        btn_gerar.grid(row=4, column=0, columnspan=2, pady=30)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para gerar análise")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=10)
        
        # Aviso sobre bibliotecas
        if not WEASYPRINT_AVAILABLE and not REPORTLAB_AVAILABLE:
            aviso = ttk.Label(main_frame, 
                            text="⚠️ Nenhuma biblioteca PDF disponível. Instale ReportLab: pip install reportlab",
                            foreground="red", font=("Arial", 9))
            aviso.grid(row=6, column=0, columnspan=2, pady=5)
        elif not WEASYPRINT_AVAILABLE and REPORTLAB_AVAILABLE:
            aviso = ttk.Label(main_frame, 
                            text="ℹ️ Usando ReportLab (compatível com Windows). Para WeasyPrint, veja TUTORIAL_INSTALAR_GTK.md",
                            foreground="blue", font=("Arial", 9))
            aviso.grid(row=6, column=0, columnspan=2, pady=5)
    
    def carregar_departamentos(self):
        """Carrega lista de departamentos do banco"""
        try:
            conn = get_db_connection()
            try:
                df_deptos = buscar_departamentos(conn)
                if not df_deptos.empty:
                    deptos_list = [f"{row['codepto']} - {row['descricao']}" 
                                  for _, row in df_deptos.iterrows()]
                    self.combo_depto['values'] = deptos_list
                    self.deptos_data = df_deptos.to_dict('records')
                else:
                    messagebox.showwarning("Aviso", "Nenhum departamento encontrado!")
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar departamentos:\n{str(e)}")
    
    def gerar_analise(self):
        """Gera análise completa do departamento"""
        try:
            if not WEASYPRINT_AVAILABLE and not REPORTLAB_AVAILABLE:
                messagebox.showerror("Erro", 
                    "Nenhuma biblioteca PDF disponível!\n\n"
                    "Instale ReportLab: pip install reportlab\n"
                    "Ou instale WeasyPrint: pip install weasyprint")
                return
            
            # Obter departamento selecionado
            depto_str = self.combo_depto.get()
            if not depto_str:
                messagebox.showwarning("Aviso", "Selecione um departamento!")
                return
            
            codepto = int(depto_str.split(' - ')[0])
            nome_depto = depto_str.split(' - ', 1)[1]
            
            # Obter período
            meses_periodo = int(self.meses_var.get())
            
            # Obter tipo de relatório
            tipo_relatorio = self.tipo_relatorio_var.get()
            
            self.status_label.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            # Conectar ao banco
            conn = get_db_connection()
            
            try:
                self.status_label.config(text=f"Analisando {nome_depto}...")
                self.master.update()
                
                # Analisar departamento
                df_resultado = analisar_depto_periodo(conn, codepto, meses_periodo)
                
                if df_resultado.empty:
                    messagebox.showinfo("Informação", 
                                      f"Nenhum dado encontrado para o departamento {nome_depto} no período selecionado.")
                    self.status_label.config(text="Nenhum resultado encontrado")
                    return
                
                self.status_label.config(text="Gerando gráficos...")
                self.master.update()
                
                # Nome do arquivo
                tipo_str = "resumido" if tipo_relatorio == "resumido" else "completo"
                nome_arquivo = f'analise_depto_{nome_depto.replace(" ", "_")}_{tipo_str}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
                
                # Buscar análise por produtos
                self.status_label.config(text="Analisando produtos...")
                self.master.update()
                df_produtos = analisar_produtos_depto(conn, codepto, meses_periodo)
                
                # Gerar PDF (usar ReportLab se WeasyPrint não disponível)
                if WEASYPRINT_AVAILABLE:
                    html_content = gerar_html_pdf(df_resultado, nome_depto, meses_periodo, tipo_relatorio)
                    self.status_label.config(text="Gerando PDF (WeasyPrint)...")
                    self.master.update()
                    pdf_path = gerar_pdf(html_content, nome_arquivo)
                elif REPORTLAB_AVAILABLE:
                    self.status_label.config(text="Gerando PDF (ReportLab)...")
                    self.master.update()
                    pdf_path = gerar_pdf_reportlab(df_resultado, df_produtos, nome_depto, meses_periodo, tipo_relatorio, nome_arquivo)
                else:
                    raise Exception("Nenhuma biblioteca PDF disponível")
                
                messagebox.showinfo("Sucesso", 
                                  f"PDF gerado com sucesso!\n\nArquivo: {pdf_path}\n\n"
                                  f"Departamento: {nome_depto}\n"
                                  f"Período: Últimos {meses_periodo} meses\n"
                                  f"Tipo: {tipo_str.title()}\n\n"
                                  f"O arquivo foi aberto automaticamente.")
                self.status_label.config(text="PDF gerado com sucesso!")
                
                # Abrir PDF
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                
            finally:
                conn.close()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{str(e)}")
            self.status_label.config(text="Erro ao gerar PDF")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = InterfaceAnaliseDepto(root)
    root.mainloop()

if __name__ == "__main__":
    main()
