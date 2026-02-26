"""
Sistema de Exclusão de NFs - Análise de Totais e Margem
Permite excluir NFs específicas e ver o impacto nos totais de venda e margem
"""

import oracledb
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

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

def formatar_valor_completo(valor):
    """Formata valor completo em R$ com pontuação brasileira"""
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

def get_mes_ano_formatado(mes, ano):
    """Retorna formato NOV/25 para o título"""
    meses_abrev = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                   'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    return f"{meses_abrev[mes-1]}/{str(ano)[-2:]}"

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

def buscar_totais(conn, mes, ano, supervisores=None, incluir_pedidos=False, excluir_nfs=None, 
                 codcli_margem=None, margem_minima=None):
    """
    Busca totais de venda e margem de lucro
    excluir_nfs: lista de números de notas fiscais para excluir (ex: ['123456', '789012'])
    codcli_margem: código do cliente para filtrar por margem
    margem_minima: porcentagem mínima de margem (ex: 6.0 para 6%)
    Retorna: (total_venda, total_margem)
    """
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    # Construir filtro de supervisores
    if supervisores and len(supervisores) > 0:
        supervisores_str = ','.join(map(str, supervisores))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    else:
        filtro_supervisor = ""
    
    # Construir filtro de posição (F para faturamento, F ou L se incluir pedidos)
    if incluir_pedidos:
        filtro_posicao = "AND C.POSICAO IN ('F', 'L')"
    else:
        filtro_posicao = "AND C.POSICAO = 'F'"
    
    # Construir filtro de exclusão de NFs
    if excluir_nfs and len(excluir_nfs) > 0:
        # Limpar e validar NFs
        nfs_limpas = []
        for nf in excluir_nfs:
            nf_limpa = str(nf).strip()
            if nf_limpa:
                nfs_limpas.append(nf_limpa)
        
        if nfs_limpas:
            # Converter para números se possível, senão usar como string
            try:
                nfs_str = ','.join([f"'{nf}'" for nf in nfs_limpas])
            except:
                nfs_str = ','.join([f"'{str(nf)}'" for nf in nfs_limpas])
            filtro_excluir_nf = f"AND C.NUMNOTA NOT IN ({nfs_str})"
        else:
            filtro_excluir_nf = ""
    else:
        filtro_excluir_nf = ""
    
    # Construir filtro de exclusão por margem baixa do cliente
    filtro_margem_cliente = ""
    if codcli_margem and margem_minima is not None:
        # Excluir itens do cliente onde a margem percentual é menor que o mínimo
        # Excluir quando: cliente = codcli_margem E margem < margem_minima
        # Incluir quando: cliente != codcli_margem OU margem >= margem_minima
        filtro_margem_cliente = f"""
          AND (
            C.CODCLI != {codcli_margem}
            OR I.PVENDA = 0
            OR ((I.PVENDA - NVL(I.VLCUSTOFIN, 0)) / I.PVENDA * 100) >= {margem_minima}
          )
        """
    
    # Query para buscar faturamento e margem (igual ao relatorio_vendedores_excel.py)
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
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          {filtro_posicao}
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          {filtro_supervisor}
          {filtro_excluir_nf}
          {filtro_margem_cliente}
    ),
    DEVOLUCOES AS (
        SELECT 
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          {filtro_supervisor}
    )
    SELECT 
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS MARGEM
    FROM VENDAS_VALIDAS V
    CROSS JOIN DEVOLUCOES D
    """
    
    try:
        df = pd.read_sql(query, conn)
        
        total_venda = float(df['FATURAMENTO'].iloc[0] if not df.empty and pd.notna(df['FATURAMENTO'].iloc[0]) else 0)
        total_margem = float(df['MARGEM'].iloc[0] if not df.empty and pd.notna(df['MARGEM'].iloc[0]) else 0)
        
        return total_venda, total_margem
    except Exception as e:
        print(f"Erro ao buscar totais: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0

def buscar_detalhes_nfs(conn, mes, ano, nfs, supervisores=None):
    """
    Busca os detalhes das NFs especificadas
    Retorna DataFrame com: NF, Data, Cliente, Produto, Quantidade, Preço, Valor Total, Vendedor
    """
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    # Construir filtro de supervisores
    if supervisores and len(supervisores) > 0:
        supervisores_str = ','.join(map(str, supervisores))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    else:
        filtro_supervisor = ""
    
    # Limpar e validar NFs
    nfs_limpas = []
    for nf in nfs:
        nf_limpa = str(nf).strip()
        if nf_limpa:
            nfs_limpas.append(nf_limpa)
    
    if not nfs_limpas:
        return pd.DataFrame()
    
    # Converter para string SQL
    try:
        nfs_str = ','.join([f"'{nf}'" for nf in nfs_limpas])
    except:
        nfs_str = ','.join([f"'{str(nf)}'" for nf in nfs_limpas])
    
    query = f"""
    SELECT 
        C.NUMNOTA AS NF,
        TO_CHAR(C.DATA, 'DD/MM/YYYY') AS DATA,
        C.CODCLI,
        CL.CLIENTE,
        I.CODPROD,
        P.DESCRICAO AS PRODUTO,
        I.QT AS QUANTIDADE,
        I.PVENDA AS PRECO,
        ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL,
        I.VLCUSTOFIN AS CUSTO_UNIT,
        ROUND(I.QT * I.VLCUSTOFIN, 2) AS CUSTO_TOTAL,
        ROUND((I.QT * I.PVENDA) - (I.QT * I.VLCUSTOFIN), 2) AS MARGEM_ITEM,
        C.CODUSUR,
        U.NOME AS VENDEDOR
    FROM PCPEDC C 
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
      AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
      AND NVL(I.BONIFIC, 'N') = 'N'
      AND C.NUMNOTA IN ({nfs_str})
      {filtro_supervisor}
    ORDER BY C.NUMNOTA, I.CODPROD
    """
    
    try:
        df = pd.read_sql(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar detalhes das NFs: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_detalhes_margem_cliente(conn, mes, ano, codcli, margem_minima, supervisores=None):
    """
    Busca os detalhes das vendas do cliente com margem abaixo do percentual especificado
    Retorna DataFrame com: NF, Data, Cliente, Produto, Quantidade, Preço, Valor Total, Margem %, Vendedor
    """
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    # Construir filtro de supervisores
    if supervisores and len(supervisores) > 0:
        supervisores_str = ','.join(map(str, supervisores))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    else:
        filtro_supervisor = ""
    
    query = f"""
    SELECT 
        C.NUMNOTA AS NF,
        TO_CHAR(C.DATA, 'DD/MM/YYYY') AS DATA,
        C.CODCLI,
        CL.CLIENTE,
        I.CODPROD,
        P.DESCRICAO AS PRODUTO,
        I.QT AS QUANTIDADE,
        I.PVENDA AS PRECO,
        ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL,
        I.VLCUSTOFIN AS CUSTO_UNIT,
        ROUND(I.QT * I.VLCUSTOFIN, 2) AS CUSTO_TOTAL,
        ROUND((I.QT * I.PVENDA) - (I.QT * I.VLCUSTOFIN), 2) AS MARGEM_ITEM,
        CASE 
            WHEN I.PVENDA > 0 THEN 
                ROUND(((I.PVENDA - NVL(I.VLCUSTOFIN, 0)) / I.PVENDA * 100), 2)
            ELSE 0 
        END AS MARGEM_PERCENT,
        C.CODUSUR,
        U.NOME AS VENDEDOR
    FROM PCPEDC C 
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
      AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
      AND NVL(I.BONIFIC, 'N') = 'N'
      AND C.CODCLI = {codcli}
      AND I.PVENDA > 0
      AND ((I.PVENDA - NVL(I.VLCUSTOFIN, 0)) / I.PVENDA * 100) < {margem_minima}
      {filtro_supervisor}
    ORDER BY C.NUMNOTA, I.CODPROD
    """
    
    try:
        df = pd.read_sql(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar detalhes das vendas com margem baixa: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_valor_margem_cliente(conn, mes, ano, codcli, margem_minima, supervisores=None):
    """
    Busca o valor total das vendas do cliente com margem abaixo do percentual especificado
    Retorna o valor total dessas vendas
    """
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
        ) AS VALOR_VENDAS
    FROM PCPEDC C 
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA >= :data_inicio
      AND C.DATA < :data_fim
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
      AND C.CODCLI = {codcli}
      AND I.PVENDA > 0
      AND ((I.PVENDA - NVL(I.VLCUSTOFIN, 0)) / I.PVENDA * 100) < {margem_minima}
      {filtro_supervisor}
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        valor = float(df['VALOR_VENDAS'].iloc[0] if not df.empty and pd.notna(df['VALOR_VENDAS'].iloc[0]) else 0)
        return valor
    except Exception as e:
        print(f"Erro ao buscar valor das vendas com margem baixa: {e}")
        import traceback
        traceback.print_exc()
        return 0

def buscar_valor_nfs(conn, mes, ano, nfs, supervisores=None):
    """
    Busca o valor total das NFs especificadas
    Retorna o valor total dessas NFs
    """
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
    
    # Limpar e validar NFs
    nfs_limpas = []
    for nf in nfs:
        nf_limpa = str(nf).strip()
        if nf_limpa:
            nfs_limpas.append(nf_limpa)
    
    if not nfs_limpas:
        return 0
    
    # Converter para string SQL
    try:
        nfs_str = ','.join([f"'{nf}'" for nf in nfs_limpas])
    except:
        nfs_str = ','.join([f"'{str(nf)}'" for nf in nfs_limpas])
    
    query = f"""
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
        ) AS VALOR_NFS
    FROM PCPEDC C 
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA >= :data_inicio
      AND C.DATA < :data_fim
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
      AND C.NUMNOTA IN ({nfs_str})
      {filtro_supervisor}
    """
    
    try:
        df = pd.read_sql(query, conn, params={'data_inicio': data_inicio, 'data_fim': data_fim})
        valor = float(df['VALOR_NFS'].iloc[0] if not df.empty and pd.notna(df['VALOR_NFS'].iloc[0]) else 0)
        return valor
    except Exception as e:
        print(f"Erro ao buscar valor das NFs: {e}")
        import traceback
        traceback.print_exc()
        return 0

def gerar_html_comparacao(total_venda_original, total_venda_sem_nfs, 
                         margem_original, margem_sem_nfs, valor_nfs_excluidas,
                         mes, ano, nfs_excluidas, df_detalhes=None,
                         codcli_margem=None, margem_minima=None, 
                         valor_margem_excluida=0, df_detalhes_margem=None,
                         valor_total_excluido_real=None):
    """Gera HTML com comparação antes e depois da exclusão"""
    mes_formatado = get_mes_ano_formatado(mes, ano)
    
    # Nome do arquivo
    nome_arquivo = f'exclusao_nf_totais_{mes:02d}_{ano}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
    html_path = Path(nome_arquivo)
    
    # Calcular margem percentual (margem sobre venda)
    margem_percent_original = (margem_original / total_venda_original * 100) if total_venda_original > 0 else 0
    margem_percent_sem_nfs = (margem_sem_nfs / total_venda_sem_nfs * 100) if total_venda_sem_nfs > 0 else 0
    
    # Formatação
    valor_venda_original = formatar_valor_completo(total_venda_original)
    valor_venda_sem_nfs = formatar_valor_completo(total_venda_sem_nfs)
    valor_margem_original = formatar_valor_completo(margem_original)
    valor_margem_sem_nfs = formatar_valor_completo(margem_sem_nfs)
    valor_nfs_excluidas_fmt = formatar_valor_completo(valor_nfs_excluidas)
    
    # Formatação de porcentagem
    margem_percent_original_fmt = f"{margem_percent_original:.2f}%"
    margem_percent_sem_nfs_fmt = f"{margem_percent_sem_nfs:.2f}%"
    
    # Lista de NFs excluídas
    nfs_lista = ', '.join(nfs_excluidas) if nfs_excluidas else 'Nenhuma'
    
    # Valor total excluído: usar o valor real (diferença) se fornecido, senão usar a soma
    # O valor real evita duplicação quando há sobreposição entre NFs excluídas e vendas por margem
    if valor_total_excluido_real is not None:
        valor_total_excluido = valor_total_excluido_real
    else:
        valor_total_excluido = valor_nfs_excluidas + valor_margem_excluida
    
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
    <title>Exclusão de NFs - {mes_formatado}</title>
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
            padding: 40px 20px;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
        }}
        
        .titulo {{
            color: #ffffff;
            font-size: 32px;
            font-weight: bold;
            margin-bottom: 20px;
            text-align: center;
            text-transform: uppercase;
            letter-spacing: 2px;
            text-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
        }}
        
        .subtitulo {{
            color: #ffffff;
            font-size: 18px;
            margin-bottom: 40px;
            text-align: center;
            opacity: 0.9;
        }}
        
        .nfs-excluidas {{
            background: rgba(244, 67, 54, 0.2);
            border: 2px solid rgba(244, 67, 54, 0.5);
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 40px;
            text-align: center;
        }}
        
        .nfs-excluidas-label {{
            color: #ffffff;
            font-size: 16px;
            margin-bottom: 10px;
            font-weight: bold;
        }}
        
        .nfs-excluidas-valor {{
            color: #ffd700;
            font-size: 20px;
            font-weight: bold;
        }}
        
        .nfs-lista {{
            color: #ffffff;
            font-size: 14px;
            margin-top: 10px;
            opacity: 0.8;
        }}
        
        .grid-comparacao {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 40px;
            margin-bottom: 40px;
        }}
        
        @media (max-width: 1200px) {{
            .grid-comparacao {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .card {{
            background: rgba(18, 67, 119, 0.15);
            border-radius: 20px;
            padding: 40px;
            backdrop-filter: blur(10px);
            border: 2px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        }}
        
        .card-titulo {{
            color: #ffffff;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 35px;
            text-transform: uppercase;
            letter-spacing: 1px;
            text-align: center;
            border-bottom: 3px solid rgba(255, 215, 0, 0.5);
            padding-bottom: 15px;
        }}
        
        .item {{
            margin-bottom: 30px;
        }}
        
        .item-label {{
            color: #ffffff;
            font-size: 14px;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 1px;
            opacity: 0.8;
        }}
        
        .item-valor {{
            color: #ffd700;
            font-size: 42px;
            font-weight: bold;
            text-shadow: 0 2px 10px rgba(255, 215, 0, 0.3);
            margin-bottom: 5px;
        }}
        
        .item-margem {{
            color: #4CAF50;
            font-size: 42px;
            font-weight: bold;
            text-shadow: 0 2px 10px rgba(76, 175, 80, 0.3);
            margin-bottom: 5px;
        }}
        
        .item-percent {{
            color: #4CAF50;
            font-size: 28px;
            font-weight: bold;
            opacity: 0.9;
            margin-top: 5px;
        }}
        
        .detalhes-section {{
            margin-top: 50px;
            margin-bottom: 40px;
        }}
        
        .detalhes-titulo {{
            color: #ffffff;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 20px;
            text-align: center;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .tabela-container {{
            overflow-x: auto;
            background: rgba(18, 67, 119, 0.15);
            border-radius: 10px;
            padding: 20px;
            backdrop-filter: blur(10px);
            border: 2px solid rgba(255, 255, 255, 0.1);
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            color: #ffffff;
            font-size: 12px;
        }}
        
        th {{
            background: rgba(255, 215, 0, 0.2);
            color: #ffd700;
            padding: 12px 8px;
            text-align: left;
            font-weight: bold;
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.5px;
            border-bottom: 2px solid rgba(255, 215, 0, 0.3);
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        td {{
            padding: 10px 8px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }}
        
        tr:hover {{
            background: rgba(255, 255, 255, 0.05);
        }}
        
        .text-right {{
            text-align: right;
        }}
        
        .text-center {{
            text-align: center;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 60px;
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
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1 class="titulo">Análise de Exclusão de NFs</h1>
        <div class="subtitulo">{mes_formatado}</div>
        
        <div class="nfs-excluidas">
            <div class="nfs-excluidas-label">Valor Total Excluído</div>
            <div class="nfs-excluidas-valor">{formatar_valor_completo(valor_total_excluido)}</div>
"""
    
    # Adicionar informações sobre NFs excluídas
    if nfs_excluidas:
        html_content += f'            <div class="nfs-lista">NFs Excluídas: {nfs_lista}</div>\n'
    
    # Adicionar informações sobre vendas excluídas por margem
    if codcli_margem and margem_minima is not None:
        html_content += f'            <div class="nfs-lista">Vendas Cliente {codcli_margem} (margem < {margem_minima}%): {formatar_valor_completo(valor_margem_excluida)}</div>\n'
    
    html_content += f"""        </div>
        
        <div class="grid-comparacao">
            <div class="card">
                <div class="card-titulo">Antes da Exclusão</div>
                
                <div class="item">
                    <div class="item-label">Total Vendido</div>
                    <div class="item-valor">{valor_venda_original}</div>
                </div>
                
                <div class="item">
                    <div class="item-label">Margem</div>
                    <div class="item-margem">{valor_margem_original}</div>
                    <div class="item-percent">{margem_percent_original_fmt}</div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-titulo">Após Exclusão</div>
                
                <div class="item">
                    <div class="item-label">Total Vendido</div>
                    <div class="item-valor">{valor_venda_sem_nfs}</div>
                </div>
                
                <div class="item">
                    <div class="item-label">Margem</div>
                    <div class="item-margem">{valor_margem_sem_nfs}</div>
                    <div class="item-percent">{margem_percent_sem_nfs_fmt}</div>
                </div>
            </div>
        </div>
"""
    
    # Adicionar tabela de detalhes se houver dados
    if df_detalhes is not None and not df_detalhes.empty:
        html_content += """
        <div class="detalhes-section">
            <div class="detalhes-titulo">Detalhes das NFs Excluídas</div>
            <div class="tabela-container">
                <table>
                    <thead>
                        <tr>
                            <th>NF</th>
                            <th>Data</th>
                            <th>Cliente</th>
                            <th>Produto</th>
                            <th class="text-center">Qtd</th>
                            <th class="text-right">Preço Unit.</th>
                            <th class="text-right">Valor Total</th>
                            <th class="text-right">Custo Unit.</th>
                            <th class="text-right">Custo Total</th>
                            <th class="text-right">Margem Item</th>
                            <th>Vendedor</th>
                        </tr>
                    </thead>
                    <tbody>
"""
        # Adicionar linhas da tabela
        for _, row in df_detalhes.iterrows():
            nf = str(row.get('nf', ''))
            data = str(row.get('data', ''))
            cliente = str(row.get('cliente', ''))[:30]  # Limitar tamanho
            produto = str(row.get('produto', ''))[:40]  # Limitar tamanho
            qtd = f"{float(row.get('quantidade', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            preco = formatar_valor_completo(row.get('preco', 0))
            valor_total = formatar_valor_completo(row.get('valor_total', 0))
            custo_unit = formatar_valor_completo(row.get('custo_unit', 0))
            custo_total = formatar_valor_completo(row.get('custo_total', 0))
            margem_item = formatar_valor_completo(row.get('margem_item', 0))
            vendedor = str(row.get('vendedor', ''))[:25]  # Limitar tamanho
            
            html_content += f"""
                        <tr>
                            <td>{nf}</td>
                            <td>{data}</td>
                            <td>{cliente}</td>
                            <td>{produto}</td>
                            <td class="text-center">{qtd}</td>
                            <td class="text-right">{preco}</td>
                            <td class="text-right">{valor_total}</td>
                            <td class="text-right">{custo_unit}</td>
                            <td class="text-right">{custo_total}</td>
                            <td class="text-right">{margem_item}</td>
                            <td>{vendedor}</td>
                        </tr>
"""
        
        html_content += """
                    </tbody>
                </table>
            </div>
        </div>
"""
    
    # Adicionar tabela de detalhes de vendas excluídas por margem baixa se houver dados
    if df_detalhes_margem is not None and not df_detalhes_margem.empty:
        html_content += f"""
        <div class="detalhes-section">
            <div class="detalhes-titulo">Vendas Excluídas - Cliente {codcli_margem} (Margem < {margem_minima}%)</div>
            <div class="tabela-container">
                <table>
                    <thead>
                        <tr>
                            <th>NF</th>
                            <th>Data</th>
                            <th>Cliente</th>
                            <th>Produto</th>
                            <th class="text-center">Qtd</th>
                            <th class="text-right">Preço Unit.</th>
                            <th class="text-right">Valor Total</th>
                            <th class="text-right">Custo Unit.</th>
                            <th class="text-right">Custo Total</th>
                            <th class="text-right">Margem Item</th>
                            <th class="text-right">Margem %</th>
                            <th>Vendedor</th>
                        </tr>
                    </thead>
                    <tbody>
"""
        # Adicionar linhas da tabela
        for _, row in df_detalhes_margem.iterrows():
            nf = str(row.get('nf', ''))
            data = str(row.get('data', ''))
            cliente = str(row.get('cliente', ''))[:30]  # Limitar tamanho
            produto = str(row.get('produto', ''))[:40]  # Limitar tamanho
            qtd = f"{float(row.get('quantidade', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            preco = formatar_valor_completo(row.get('preco', 0))
            valor_total = formatar_valor_completo(row.get('valor_total', 0))
            custo_unit = formatar_valor_completo(row.get('custo_unit', 0))
            custo_total = formatar_valor_completo(row.get('custo_total', 0))
            margem_item = formatar_valor_completo(row.get('margem_item', 0))
            margem_percent = f"{float(row.get('margem_percent', 0)):.2f}%"
            vendedor = str(row.get('vendedor', ''))[:25]  # Limitar tamanho
            
            html_content += f"""
                        <tr>
                            <td>{nf}</td>
                            <td>{data}</td>
                            <td>{cliente}</td>
                            <td>{produto}</td>
                            <td class="text-center">{qtd}</td>
                            <td class="text-right">{preco}</td>
                            <td class="text-right">{valor_total}</td>
                            <td class="text-right">{custo_unit}</td>
                            <td class="text-right">{custo_total}</td>
                            <td class="text-right">{margem_item}</td>
                            <td class="text-right">{margem_percent}</td>
                            <td>{vendedor}</td>
                        </tr>
"""
        
        html_content += """
                    </tbody>
                </table>
            </div>
        </div>
"""
    
    html_content += f"""
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

class InterfaceExclusaoNF:
    def __init__(self, master):
        self.master = master
        self.master.title("Exclusão de NFs - Análise de Totais e Margem")
        self.master.geometry("700x850")
        
        main_frame = ttk.Frame(self.master, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Exclusão de NFs", 
                          font=("Arial", 18, "bold"))
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
        canvas = tk.Canvas(canvas_frame, height=120)
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
        
        # Campo de NFs para excluir
        ttk.Label(main_frame, text="NFs para Excluir (separadas por vírgula):").grid(
            row=4, column=0, columnspan=2, sticky=tk.W, pady=(20, 5))
        
        self.nfs_text = scrolledtext.ScrolledText(main_frame, height=4, wrap=tk.WORD)
        self.nfs_text.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(main_frame, 
                 text="Exemplo: 123456, 789012, 345678",
                 font=("Arial", 9), foreground="gray").grid(
            row=6, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Checkbox para incluir pedidos
        self.incluir_pedidos_var = tk.BooleanVar(value=False)
        checkbox_pedidos = ttk.Checkbutton(
            main_frame,
            text="Incluir Pedidos (status 'L') além do Faturamento (status 'F')",
            variable=self.incluir_pedidos_var
        )
        checkbox_pedidos.grid(row=7, column=0, columnspan=2, pady=10, sticky=tk.W)
        
        # Frame para filtro de margem do cliente
        frame_margem_cliente = ttk.LabelFrame(main_frame, text="Filtro por Cliente e Margem", padding="10")
        frame_margem_cliente.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        frame_margem_cliente.columnconfigure(1, weight=1)
        
        # Campo código do cliente
        ttk.Label(frame_margem_cliente, text="Código do Cliente:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.codcli_entry = ttk.Entry(frame_margem_cliente, width=20)
        self.codcli_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Campo margem mínima
        ttk.Label(frame_margem_cliente, text="Margem Mínima (%):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.margem_minima_entry = ttk.Entry(frame_margem_cliente, width=20)
        self.margem_minima_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        ttk.Label(frame_margem_cliente, 
                 text="Excluir vendas deste cliente com margem abaixo do percentual informado",
                 font=("Arial", 9), foreground="gray").grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Botão gerar
        btn_gerar = ttk.Button(main_frame, text="Calcular Totais e Margem", 
                              command=self.calcular_totais)
        btn_gerar.grid(row=9, column=0, columnspan=2, pady=20)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para calcular")
        self.status_label.grid(row=10, column=0, columnspan=2, pady=10)
    
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
    
    def calcular_totais(self):
        """Calcula totais com e sem as NFs excluídas"""
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
            
            # Obter NFs para excluir
            nfs_text = self.nfs_text.get("1.0", tk.END).strip()
            nfs_lista = []
            if nfs_text:
                # Separar por vírgula e limpar
                nfs_lista = [nf.strip() for nf in nfs_text.split(',') if nf.strip()]
            
            # Obter opção de incluir pedidos
            incluir_pedidos = self.incluir_pedidos_var.get()
            
            # Obter código do cliente e margem mínima
            codcli_text = self.codcli_entry.get().strip()
            margem_text = self.margem_minima_entry.get().strip()
            
            codcli_margem = None
            margem_minima = None
            
            if codcli_text and margem_text:
                try:
                    codcli_margem = int(codcli_text)
                    margem_minima = float(margem_text)
                except ValueError:
                    messagebox.showwarning("Aviso", "Código do cliente e margem mínima devem ser números válidos!")
                    return
            
            self.status_label.config(text="Conectando ao banco de dados...")
            self.master.update()
            
            # Conectar ao banco
            conn = get_db_connection()
            
            try:
                # Buscar totais originais (sem exclusão)
                self.status_label.config(text="Buscando totais originais...")
                self.master.update()
                total_venda_original, margem_original = buscar_totais(
                    conn, mes, ano, supervisores, incluir_pedidos, 
                    excluir_nfs=None, codcli_margem=None, margem_minima=None)
                
                # Buscar totais sem as exclusões (NFs + margem baixa)
                df_detalhes = None
                df_detalhes_margem = None
                valor_nfs_excluidas = 0
                valor_margem_excluida = 0
                
                # Buscar totais com todas as exclusões aplicadas
                self.status_label.config(text="Buscando totais com exclusões aplicadas...")
                self.master.update()
                total_venda_sem_nfs, margem_sem_nfs = buscar_totais(
                    conn, mes, ano, supervisores, incluir_pedidos, 
                    excluir_nfs=nfs_lista if nfs_lista else None,
                    codcli_margem=codcli_margem, margem_minima=margem_minima)
                
                # Buscar detalhes das NFs excluídas
                if nfs_lista:
                    self.status_label.config(text="Buscando valor das NFs excluídas...")
                    self.master.update()
                    valor_nfs_excluidas = buscar_valor_nfs(conn, mes, ano, nfs_lista, supervisores)
                    
                    self.status_label.config(text="Buscando detalhes das NFs excluídas...")
                    self.master.update()
                    df_detalhes = buscar_detalhes_nfs(conn, mes, ano, nfs_lista, supervisores)
                
                # Buscar detalhes das vendas excluídas por margem baixa
                if codcli_margem and margem_minima is not None:
                    self.status_label.config(text=f"Buscando vendas do cliente {codcli_margem} com margem < {margem_minima}%...")
                    self.master.update()
                    valor_margem_excluida = buscar_valor_margem_cliente(
                        conn, mes, ano, codcli_margem, margem_minima, supervisores)
                    
                    self.status_label.config(text="Buscando detalhes das vendas excluídas por margem...")
                    self.master.update()
                    df_detalhes_margem = buscar_detalhes_margem_cliente(
                        conn, mes, ano, codcli_margem, margem_minima, supervisores)
                
                # Calcular margem percentual
                margem_percent_original = (margem_original / total_venda_original * 100) if total_venda_original > 0 else 0
                margem_percent_sem_nfs = (margem_sem_nfs / total_venda_sem_nfs * 100) if total_venda_sem_nfs > 0 else 0
                
                # Calcular valor total excluído (diferença real entre original e após exclusão)
                # Isso evita duplicação quando há sobreposição entre NFs excluídas e vendas por margem
                valor_total_excluido_real = total_venda_original - total_venda_sem_nfs
                
                self.status_label.config(text="Gerando relatório...")
                self.master.update()
                
                # Gerar HTML
                html_path = gerar_html_comparacao(
                    total_venda_original, total_venda_sem_nfs,
                    margem_original, margem_sem_nfs, valor_nfs_excluidas,
                    mes, ano, nfs_lista, df_detalhes,
                    codcli_margem, margem_minima, valor_margem_excluida, df_detalhes_margem,
                    valor_total_excluido_real)
                
                messagebox.showinfo("Sucesso", 
                                  f"Análise gerada com sucesso!\n\nArquivo: {html_path}\n\n"
                                  f"Total Vendido Original: {formatar_valor_completo(total_venda_original)}\n"
                                  f"Total Vendido Sem NFs: {formatar_valor_completo(total_venda_sem_nfs)}\n"
                                  f"Margem Original: {formatar_valor_completo(margem_original)} ({margem_percent_original:.2f}%)\n"
                                  f"Margem Sem NFs: {formatar_valor_completo(margem_sem_nfs)} ({margem_percent_sem_nfs:.2f}%)\n\n"
                                  f"O arquivo foi aberto no navegador.")
                self.status_label.config(text="Análise gerada com sucesso!")
                
            finally:
                conn.close()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao calcular totais:\n{str(e)}")
            self.status_label.config(text="Erro ao calcular totais")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = InterfaceExclusaoNF(root)
    root.mainloop()

if __name__ == "__main__":
    main()

