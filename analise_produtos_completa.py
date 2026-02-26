"""
Sistema de análise completa de produtos
- Aceita múltiplos CODPROD separados por vírgula
- Mostra compras dos últimos 3 meses
- Mostra estoque atual (Oracle + Vilog)
- Analisa lucro/prejuízo dos últimos 12 meses
"""

import oracledb
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import sys
import os

# Adiciona o diretório raiz ao path para importar google_sheets_api
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importa função para buscar estoque da Vilog
try:
    from google_sheets_api import get_vilog_stock_data
except ImportError:
    print("⚠️  Aviso: Não foi possível importar get_vilog_stock_data. Continuando sem estoque Vilog.")
    get_vilog_stock_data = lambda: {}

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
# FUNÇÕES
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

def buscar_dados_produtos(conn, codprod_list):
    """Busca dados básicos dos produtos"""
    codprod_str = ', '.join(map(str, codprod_list))
    
    query = f"""
    SELECT 
        P.CODPROD,
        P.DESCRICAO,
        P.CODFAB,
        P.CODAUXILIAR AS EAN,
        P.EMBALAGEM,
        P.PESOLIQ,
        P.QTUNITCX,
        D.DESCRICAO AS DEPARTAMENTO,
        D.CODEPTO,
        E.CUSTOFIN AS CUSTO_FINAL,
        COALESCE((E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA) / NULLIF(P.QTUNITCX, 1), 
                 (E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA), 0) AS ESTOQUE_CX_ORACLE,
        COALESCE((E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA), 0) AS ESTOQUE_UN_ORACLE
    FROM PCPRODUT P
    LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
    LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
    WHERE P.CODPROD IN ({codprod_str})
      AND P.DTEXCLUSAO IS NULL
      AND (P.OBS2 <> 'FL' OR P.OBS2 IS NULL)
      AND P.REVENDA = 'S'
    ORDER BY P.CODPROD
    """
    
    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]
    return df

def buscar_compras_ultimos_3_meses(conn, codprod_list):
    """Busca compras dos últimos 3 meses"""
    codprod_str = ', '.join(map(str, codprod_list))
    
    # Calcula data de 3 meses atrás
    data_inicio = datetime.now() - relativedelta(months=3)
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    
    query = f"""
    SELECT
        M.CODPROD,
        TO_CHAR(TRUNC(N.DTENT, 'MM'), 'MM/YYYY') AS MES_REF,
        TRUNC(N.DTENT, 'MM') AS MES_DATA,
        SUM(M.QT) AS QT_COMPRADA
    FROM PCNFENT N
    JOIN PCMOV M ON N.NUMTRANSENT = M.NUMTRANSENT AND M.CODOPER = 'E'
    WHERE N.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
      AND M.CODPROD IN ({codprod_str})
      AND N.CODFILIAL = '1'
    GROUP BY M.CODPROD, TRUNC(N.DTENT, 'MM')
    ORDER BY M.CODPROD, TRUNC(N.DTENT, 'MM') DESC
    """
    
    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]
    return df

def buscar_vendas_ultimos_12_meses(conn, codprod_list):
    """Busca vendas dos últimos 12 meses para análise de lucro/prejuízo"""
    codprod_str = ', '.join(map(str, codprod_list))
    
    # Calcula data de 12 meses atrás
    data_inicio = datetime.now() - relativedelta(months=12)
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = datetime.now().strftime('%d/%m/%Y')
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            I.CODPROD,
            TO_CHAR(TRUNC(C.DATA, 'MM'), 'MM/YYYY') AS MES_REF,
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
                            NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS QT_VENDIDA
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA <= TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND I.CODPROD IN ({codprod_str})
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        GROUP BY I.CODPROD, TRUNC(C.DATA, 'MM')
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD,
            TO_CHAR(TRUNC(D.DTENT, 'MM'), 'MM/YYYY') AS MES_REF,
            TRUNC(D.DTENT, 'MM') AS MES_DATA,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT <= TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND D.CODPROD IN ({codprod_str})
        GROUP BY D.CODPROD, TRUNC(D.DTENT, 'MM')
    )
    SELECT 
        COALESCE(V.CODPROD, D.CODPROD) AS CODPROD,
        COALESCE(V.MES_REF, D.MES_REF) AS MES_REF,
        COALESCE(V.MES_DATA, D.MES_DATA) AS MES_DATA,
        NVL(V.VALOR_BRUTO, 0) AS VALOR_BRUTO,
        NVL(D.VALOR_DEVOLVIDO, 0) AS VALOR_DEVOLVIDO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
        NVL(V.CUSTO_BRUTO, 0) AS CUSTO_BRUTO,
        NVL(D.CUSTO_DEVOLVIDO, 0) AS CUSTO_DEVOLVIDO,
        (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO,
        NVL(V.QT_VENDIDA, 0) AS QT_VENDIDA
    FROM VENDAS_VALIDAS V
    FULL OUTER JOIN DEVOLUCOES D ON V.CODPROD = D.CODPROD AND V.MES_DATA = D.MES_DATA
    WHERE (NVL(V.VALOR_BRUTO, 0) > 0 OR NVL(D.VALOR_DEVOLVIDO, 0) > 0)
    ORDER BY COALESCE(V.CODPROD, D.CODPROD), COALESCE(V.MES_DATA, D.MES_DATA) DESC
    """
    
    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]
    return df

def format_real(valor):
    """Formata valor como moeda brasileira"""
    if pd.isna(valor) or valor == 0:
        return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def format_numero_br(numero, decimais=0):
    """Formata número com pontuação brasileira (ponto para milhar, vírgula para decimal)"""
    if pd.isna(numero) or numero == 0:
        return "0" if decimais == 0 else "0," + "0" * decimais
    
    if decimais == 0:
        # Número inteiro
        return f"{int(numero):,}".replace(',', '.')
    else:
        # Número com decimais
        format_str = f"{{:,.{decimais}f}}"
        formatted = format_str.format(float(numero))
        return formatted.replace(',', 'X').replace('.', ',').replace('X', '.')

def get_tres_meses_anteriores():
    """Retorna os últimos 3 meses no formato [(YYYYMM, 'nov2025'), ...] ordenados do mais antigo para o mais recente"""
    meses = []
    hoje = datetime.now()
    # Coleta os 3 meses (do mais antigo para o mais recente)
    for i in range(2, -1, -1):  # i = 2, 1, 0 (mais antigo primeiro)
        mes = hoje - relativedelta(months=i)
        # Formato do nome: "nov2025" (abreviação do mês + ano)
        meses_abrev = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                       'jul', 'ago', 'set', 'out', 'nov', 'dez']
        mes_nome = f"{meses_abrev[mes.month - 1]}{mes.year}"
        meses.append((mes.strftime('%Y%m'), mes_nome))
    return meses

def gerar_html(df_produtos, df_compras, df_vendas, estoque_vilog):
    """Gera HTML com análise completa dos produtos e botão de imprimir"""
    try:
        # Processar meses de compra
        meses_ref = get_tres_meses_anteriores()
        meses_ano = [m[0] for m in meses_ref]
        meses_nome = [m[1] for m in meses_ref]
        meses_nome_compra = [f"{m} (compra)" for m in meses_nome]
        
        # Processar compras por mês
        if not df_compras.empty:
            # Converter mes_data para datetime se necessário
            if 'mes_data' in df_compras.columns:
                df_compras['yyyymm'] = pd.to_datetime(df_compras['mes_data'], errors='coerce').dt.strftime('%Y%m')
            else:
                # Se não tiver mes_data, usar mes_ref
                df_compras['yyyymm'] = pd.to_datetime(df_compras['mes_ref'], format='%m/%Y', errors='coerce').dt.strftime('%Y%m')
            compras_grouped = df_compras.groupby(['codprod', 'yyyymm'], as_index=False)['qt_comprada'].sum()
            
            compras_dict = {}
            for _, row in compras_grouped.iterrows():
                codprod = int(row['codprod'])
                mes = row['yyyymm']
                if mes in meses_ano:
                    mes_idx = meses_ano.index(mes)
                    mes_nome = meses_nome[mes_idx]
                    if codprod not in compras_dict:
                        compras_dict[codprod] = {mn: 0 for mn in meses_nome}
                    compras_dict[codprod][mes_nome] = float(row['qt_comprada'])
        else:
            compras_dict = {}
        
        # Processar vendas para resumo e médias dos últimos 12 meses
        vendas_resumo = {}
        vendas_12_meses = {}  # Para calcular médias dos últimos 12 meses
        vendas_9_meses = {}  # Para calcular giro médio dos últimos 9 meses
        
        if not df_vendas.empty:
            # Filtrar últimos 12 meses para médias
            hoje = datetime.now()
            meses_12_ultimos = []
            for i in range(12):
                mes = (hoje - relativedelta(months=i)).strftime('%Y%m')
                meses_12_ultimos.append(mes)
            
            # Filtrar últimos 9 meses para giro médio
            meses_9_ultimos = []
            for i in range(9):
                mes = (hoje - relativedelta(months=i)).strftime('%Y%m')
                meses_9_ultimos.append(mes)
            
            for _, row in df_vendas.iterrows():
                codprod = int(row['codprod'])
                
                # Obter mês da venda
                mes_yyyymm = None
                if 'mes_data' in row and pd.notna(row['mes_data']):
                    mes_data = pd.to_datetime(row['mes_data'], errors='coerce')
                    if pd.notna(mes_data):
                        mes_yyyymm = mes_data.strftime('%Y%m')
                elif 'mes_ref' in row and pd.notna(row['mes_ref']):
                    try:
                        mes_data = pd.to_datetime(row['mes_ref'], format='%m/%Y', errors='coerce')
                        if pd.notna(mes_data):
                            mes_yyyymm = mes_data.strftime('%Y%m')
                    except:
                        pass
                
                # Inicializar estruturas
                if codprod not in vendas_resumo:
                    vendas_resumo[codprod] = {
                        'faturamento_mensal': [],
                        'custo_mensal': [],
                        'qt_vendida_mensal': [],
                        'margem_mensal': [],
                        'meses_prejuizo': 0,
                        'qt_vendida_atual': 0  # Quantidade vendida do mês atual
                    }
                
                # Obter mês atual
                mes_atual = hoje.strftime('%Y%m')
                
                # Se o mês está nos últimos 12 meses, adiciona para médias
                if mes_yyyymm and mes_yyyymm in meses_12_ultimos:
                    fat_liquido = float(row['faturamento_liquido'])
                    custo_liquido = float(row['custo_liquido'])
                    qt_vendida = float(row['qt_vendida'])
                    
                    vendas_resumo[codprod]['faturamento_mensal'].append(fat_liquido)
                    vendas_resumo[codprod]['custo_mensal'].append(custo_liquido)
                    vendas_resumo[codprod]['qt_vendida_mensal'].append(qt_vendida)
                    
                    # Se for o mês atual, armazena a quantidade vendida
                    if mes_yyyymm == mes_atual:
                        vendas_resumo[codprod]['qt_vendida_atual'] += qt_vendida
                    
                    # Calcular margem mensal
                    if fat_liquido > 0:
                        margem_mes = ((fat_liquido - custo_liquido) / fat_liquido) * 100
                        vendas_resumo[codprod]['margem_mensal'].append(margem_mes)
                    
                    if fat_liquido < custo_liquido:
                        vendas_resumo[codprod]['meses_prejuizo'] += 1
                
                # Se o mês está nos últimos 9 meses, adiciona para giro médio
                if mes_yyyymm and mes_yyyymm in meses_9_ultimos:
                    if codprod not in vendas_9_meses:
                        vendas_9_meses[codprod] = []
                    vendas_9_meses[codprod].append(float(row['qt_vendida']))
        
        # Preparar dados para HTML
        rows_data = []
        for _, p in df_produtos.iterrows():
            codprod = int(p['codprod'])
            codprod_str = str(codprod)
            nome = p['descricao']
            embalagem = p['embalagem'] if pd.notna(p['embalagem']) else ''
            qtunitcx = float(p.get('qtunitcx') or 0)
            custo_valor = float(p.get('custo_final') if pd.notna(p.get('custo_final')) else 0)
            custo = format_real(custo_valor)
            
            # Compras dos últimos 3 meses
            compras_meses_units = compras_dict.get(codprod, {mn: 0 for mn in meses_nome})
            compras_meses_caixas = []
            compras_soma = 0
            for mn in meses_nome:
                qtqtd = compras_meses_units.get(mn, 0)
                if qtunitcx > 0:
                    caixas = int(round(qtqtd / qtunitcx))
                else:
                    caixas = 0
                compras_meses_caixas.append(f"{format_numero_br(caixas, 0)} cx")
                compras_soma += caixas
            
            # Estoques
            estoque_cx_oracle = float(p['estoque_cx_oracle']) if pd.notna(p['estoque_cx_oracle']) else 0
            estoque_cx_vilog = float(estoque_vilog.get(codprod_str, 0))
            estoque_total = int(round(estoque_cx_oracle + estoque_cx_vilog))
            estoque_dicon_str = f"{format_numero_br(int(round(estoque_cx_oracle)), 0)} cx"
            estoque_vilog_str = f"{format_numero_br(int(round(estoque_cx_vilog)), 0)} cx"
            estoque_total_str = f"{format_numero_br(estoque_total, 0)} cx"
            
            # Resumo de vendas (médias dos últimos 12 meses)
            resumo = vendas_resumo.get(codprod, {
                'faturamento_mensal': [],
                'custo_mensal': [],
                'qt_vendida_mensal': [],
                'margem_mensal': [],
                'meses_prejuizo': 0,
                'qt_vendida_atual': 0
            })
            
            # Calcular médias dos 12 meses
            if len(resumo['faturamento_mensal']) > 0:
                fat_medio = sum(resumo['faturamento_mensal']) / len(resumo['faturamento_mensal'])
                custo_medio = sum(resumo['custo_mensal']) / len(resumo['custo_mensal'])
                qt_vendida_media = sum(resumo['qt_vendida_mensal']) / len(resumo['qt_vendida_mensal'])
                margem_media = sum(resumo['margem_mensal']) / len(resumo['margem_mensal']) if len(resumo['margem_mensal']) > 0 else 0
                
                # Converter quantidade vendida média para caixas
                if qtunitcx > 0:
                    qt_vendida_cx = qt_vendida_media / qtunitcx
                    qt_vendida_str = f"{format_numero_br(qt_vendida_cx, 2)} cx"
                else:
                    qt_vendida_str = "N/A"
                
                # Converter quantidade vendida do mês atual para caixas
                if qtunitcx > 0:
                    qt_vendida_atual_cx = resumo['qt_vendida_atual'] / qtunitcx
                    qt_vendida_atual_str = f"{format_numero_br(qt_vendida_atual_cx, 2)} cx"
                else:
                    qt_vendida_atual_str = "N/A"
            else:
                fat_medio = 0
                custo_medio = 0
                margem_media = 0
                qt_vendida_str = "N/A"
                qt_vendida_atual_str = "N/A"
            
            lucro_prejuizo = fat_medio - custo_medio
            
            # Calcular giro médio dos últimos 9 meses (em caixas)
            # Converte cada mês para caixas e depois faz a média
            vendas_9m = vendas_9_meses.get(codprod, [])
            if vendas_9m and len(vendas_9m) > 0:
                if qtunitcx > 0:
                    # Converte cada mês para caixas
                    caixas_por_mes = []
                    for unidades_mes in vendas_9m:
                        caixas_mes = unidades_mes / qtunitcx
                        caixas_por_mes.append(caixas_mes)
                    
                    # Calcula a média dos valores mensais
                    venda_media_cx = sum(caixas_por_mes) / len(caixas_por_mes)
                    venda_media_str = f"{format_numero_br(venda_media_cx, 2)} cx"
                else:
                    venda_media_str = "N/A"
            else:
                venda_media_str = "N/A"
            
            rows_data.append({
                'codprod': codprod,
                'nome': nome,
                'embalagem': embalagem,
                'compras_meses': compras_meses_caixas,
                'estoque_dicon': estoque_dicon_str,
                'estoque_vilog': estoque_vilog_str,
                'estoque_total': estoque_total_str,
                'estoque_total_num': estoque_total,  # Para ordenação
                'custo': custo,
                'venda_media_cx': venda_media_str,
                'qt_vendida_media': qt_vendida_str,
                'qt_vendida_atual': qt_vendida_atual_str,
                'faturamento': format_real(fat_medio),
                'custo_total': format_real(custo_medio),
                'lucro_prejuizo': format_real(lucro_prejuizo),
                'margem': f"{margem_media:.1f}%",
                'meses_prejuizo': resumo['meses_prejuizo'],
                'lucro_valor': lucro_prejuizo,
                'compras_soma': compras_soma
            })
        
        # Ordenar por compras e estoque
        rows_data.sort(key=lambda x: (-x['compras_soma'], -x['estoque_total_num'], x['nome'].lower()))
        
        # Gerar HTML
        html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Análise Completa de Produtos</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        
        .header {{
            background: linear-gradient(135deg, #124377 0%, #1a5a9a 100%);
            color: white;
            padding: 30px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .header h1 {{
            font-size: 28px;
            margin-bottom: 10px;
        }}
        
        .header p {{
            font-size: 14px;
            opacity: 0.9;
        }}
        
        .excel-button {{
            background-color: #28a745;
            color: white;
            border: none;
            padding: 12px 30px;
            font-size: 16px;
            border-radius: 5px;
            cursor: pointer;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            transition: background-color 0.3s;
        }}
        
        .excel-button:hover {{
            background-color: #218838;
        }}
        
        .excel-button:active {{
            transform: scale(0.98);
        }}
        
        .controls-panel {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .controls-panel h3 {{
            margin-bottom: 15px;
            color: #124377;
            font-size: 16px;
        }}
        
        .controls-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 10px;
        }}
        
        .control-item {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        
        .control-item input[type="checkbox"] {{
            width: 18px;
            height: 18px;
            cursor: pointer;
        }}
        
        .control-item label {{
            cursor: pointer;
            font-size: 14px;
            user-select: none;
        }}
        
        .table-container {{
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            overflow-x: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }}
        
        thead {{
            background: linear-gradient(135deg, #124377 0%, #1a5a9a 100%);
            color: white;
        }}
        
        th {{
            padding: 12px 8px;
            text-align: center;
            font-weight: 600;
            border-right: 1px solid rgba(255,255,255,0.2);
            white-space: nowrap;
        }}
        
        th:first-child {{
            border-top-left-radius: 8px;
        }}
        
        th:last-child {{
            border-top-right-radius: 8px;
            border-right: none;
        }}
        
        td {{
            padding: 10px 8px;
            text-align: center;
            border-bottom: 1px solid #e0e0e0;
            border-right: 1px solid #e0e0e0;
        }}
        
        td:last-child {{
            border-right: none;
        }}
        
        tbody tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        
        tbody tr:hover {{
            background-color: #e3f2fd;
        }}
        
        .produto-nome {{
            text-align: left !important;
            font-weight: 500;
            max-width: 250px;
        }}
        
        .codigo {{
            font-weight: 600;
            color: #124377;
        }}
        
        .negativo {{
            color: #dc3545;
            font-weight: 600;
        }}
        
        .positivo {{
            color: #28a745;
            font-weight: 600;
        }}
        
        .prejuizo {{
            background-color: #fff3cd !important;
        }}
        
        .col-hidden {{
            display: none !important;
        }}
        
        @media print {{
            body {{
                background-color: white;
                padding: 0;
            }}
            
            .excel-button, .controls-panel {{
                display: none;
            }}
            
            .header {{
                margin-bottom: 20px;
            }}
            
            .table-container {{
                box-shadow: none;
            }}
            
            table {{
                font-size: 10px;
            }}
            
            th, td {{
                padding: 6px 4px;
            }}
            
            @page {{
                margin: 1cm;
                size: landscape;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Análise Completa de Produtos</h1>
        <p>Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
    </div>
    
    <button class="excel-button" onclick="exportarExcel()">📊 Salvar como Excel</button>
    
    <div class="controls-panel">
        <h3>📋 Filtrar Colunas (marque para mostrar):</h3>
        <div class="controls-grid">
            <div class="control-item">
                <input type="checkbox" id="col-codigo" checked onchange="toggleColumn('col-codigo', 0)">
                <label for="col-codigo">Código</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-produto" checked onchange="toggleColumn('col-produto', 1)">
                <label for="col-produto">Produto</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-embalagem" checked onchange="toggleColumn('col-embalagem', 2)">
                <label for="col-embalagem">Embalagem</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-compra1" checked onchange="toggleColumn('col-compra1', 3)">
                <label for="col-compra1">{meses_nome_compra[0]}</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-compra2" checked onchange="toggleColumn('col-compra2', 4)">
                <label for="col-compra2">{meses_nome_compra[1]}</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-compra3" checked onchange="toggleColumn('col-compra3', 5)">
                <label for="col-compra3">{meses_nome_compra[2]}</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-est-dicon" checked onchange="toggleColumn('col-est-dicon', 6)">
                <label for="col-est-dicon">Est. DICON</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-est-vilog" checked onchange="toggleColumn('col-est-vilog', 7)">
                <label for="col-est-vilog">Est. VILOG</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-est-total" checked onchange="toggleColumn('col-est-total', 8)">
                <label for="col-est-total">Est. Total</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-custo" checked onchange="toggleColumn('col-custo', 9)">
                <label for="col-custo">Custo</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-venda-media" checked onchange="toggleColumn('col-venda-media', 10)">
                <label for="col-venda-media">Giro Médio</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-qt-vendida" checked onchange="toggleColumn('col-qt-vendida', 11)">
                <label for="col-qt-vendida">Qtd Vendida (média)</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-qt-vendida-atual" checked onchange="toggleColumn('col-qt-vendida-atual', 12)">
                <label for="col-qt-vendida-atual">Qtd Vendida (atual)</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-fat-12m" checked onchange="toggleColumn('col-fat-12m', 13)">
                <label for="col-fat-12m">Fat. 12m (média)</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-custo-12m" checked onchange="toggleColumn('col-custo-12m', 14)">
                <label for="col-custo-12m">Custo 12m (média)</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-lucro-prej" checked onchange="toggleColumn('col-lucro-prej', 15)">
                <label for="col-lucro-prej">Lucro/Prej</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-margem" checked onchange="toggleColumn('col-margem', 16)">
                <label for="col-margem">Margem (média)</label>
            </div>
            <div class="control-item">
                <input type="checkbox" id="col-meses-prej" checked onchange="toggleColumn('col-meses-prej', 17)">
                <label for="col-meses-prej">Meses Prej</label>
            </div>
        </div>
    </div>
    
    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th class="col-codigo">Código</th>
                    <th class="col-produto">Produto</th>
                    <th class="col-embalagem">Embalagem</th>
                    <th class="col-compra1">{meses_nome_compra[0]}</th>
                    <th class="col-compra2">{meses_nome_compra[1]}</th>
                    <th class="col-compra3">{meses_nome_compra[2]}</th>
                    <th class="col-est-dicon">Est. DICON</th>
                    <th class="col-est-vilog">Est. VILOG</th>
                    <th class="col-est-total">Est. Total</th>
                    <th class="col-custo">Custo</th>
                    <th class="col-venda-media">Giro Médio</th>
                    <th class="col-qt-vendida">Qtd Vendida (média)</th>
                    <th class="col-qt-vendida-atual">Qtd Vendida (atual)</th>
                    <th class="col-fat-12m">Fat. 12m (média)</th>
                    <th class="col-custo-12m">Custo 12m (média)</th>
                    <th class="col-lucro-prej">Lucro/Prej</th>
                    <th class="col-margem">Margem (média)</th>
                    <th class="col-meses-prej">Meses Prej</th>
                </tr>
            </thead>
            <tbody>
"""
        
        for row in rows_data:
            lucro_class = 'negativo' if row['lucro_valor'] < 0 else 'positivo'
            prejuizo_class = 'prejuizo' if row['meses_prejuizo'] > 0 else ''
            
            html_content += f"""
                <tr class="{prejuizo_class}">
                    <td class="codigo col-codigo">{row['codprod']}</td>
                    <td class="produto-nome col-produto">{row['nome']}</td>
                    <td class="col-embalagem">{row['embalagem']}</td>
                    <td class="col-compra1">{row['compras_meses'][0]}</td>
                    <td class="col-compra2">{row['compras_meses'][1]}</td>
                    <td class="col-compra3">{row['compras_meses'][2]}</td>
                    <td class="col-est-dicon">{row['estoque_dicon']}</td>
                    <td class="col-est-vilog">{row['estoque_vilog']}</td>
                    <td class="col-est-total"><strong>{row['estoque_total']}</strong></td>
                    <td class="col-custo">{row['custo']}</td>
                    <td class="col-venda-media">{row['venda_media_cx']}</td>
                    <td class="col-qt-vendida">{row['qt_vendida_media']}</td>
                    <td class="col-qt-vendida-atual">{row['qt_vendida_atual']}</td>
                    <td class="col-fat-12m">{row['faturamento']}</td>
                    <td class="col-custo-12m">{row['custo_total']}</td>
                    <td class="col-lucro-prej {lucro_class}">{row['lucro_prejuizo']}</td>
                    <td class="col-margem {lucro_class}">{row['margem']}</td>
                    <td class="col-meses-prej">{row['meses_prejuizo']}</td>
                </tr>
"""
        
        html_content += """
            </tbody>
        </table>
    </div>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script>
        function toggleColumn(colClass, colIndex) {
            const checkbox = document.getElementById(colClass);
            const isVisible = checkbox.checked;
            
            // Esconder/mostrar cabeçalho
            const headers = document.querySelectorAll(`th.${colClass}`);
            headers.forEach(header => {
                if (isVisible) {
                    header.classList.remove('col-hidden');
                } else {
                    header.classList.add('col-hidden');
                }
            });
            
            // Esconder/mostrar células
            const cells = document.querySelectorAll(`td.${colClass}`);
            cells.forEach(cell => {
                if (isVisible) {
                    cell.classList.remove('col-hidden');
                } else {
                    cell.classList.add('col-hidden');
                }
            });
        }
        
        function exportarExcel() {
            try {
                const table = document.querySelector('table');
                const wb = XLSX.utils.book_new();
                
                // Pegar apenas colunas visíveis
                const headers = [];
                const headerCells = table.querySelectorAll('thead th:not(.col-hidden)');
                headerCells.forEach(th => {
                    headers.push(th.textContent.trim());
                });
                
                // Pegar dados das linhas (apenas colunas visíveis)
                const data = [headers];
                const rows = table.querySelectorAll('tbody tr');
                rows.forEach(row => {
                    const rowData = [];
                    const cells = row.querySelectorAll('td:not(.col-hidden)');
                    cells.forEach(cell => {
                        rowData.push(cell.textContent.trim());
                    });
                    data.push(rowData);
                });
                
                // Criar worksheet
                const ws = XLSX.utils.aoa_to_sheet(data);
                
                // Ajustar largura das colunas automaticamente
                const colWidths = [];
                for (let i = 0; i < headers.length; i++) {
                    let maxLength = headers[i].length;
                    data.forEach(row => {
                        if (row[i] && row[i].toString().length > maxLength) {
                            maxLength = row[i].toString().length;
                        }
                    });
                    // Limitar largura máxima e adicionar padding
                    colWidths.push({ wch: Math.min(Math.max(maxLength + 2, 10), 60) });
                }
                ws['!cols'] = colWidths;
                
                // Congelar primeira linha (cabeçalho)
                ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft' };
                
                // Adicionar worksheet ao workbook
                XLSX.utils.book_append_sheet(wb, ws, "Análise Produtos");
                
                // Gerar nome do arquivo
                const dataAtual = new Date();
                const nomeArquivo = `analise_produtos_${dataAtual.getFullYear()}${String(dataAtual.getMonth() + 1).padStart(2, '0')}${String(dataAtual.getDate()).padStart(2, '0')}_${String(dataAtual.getHours()).padStart(2, '0')}${String(dataAtual.getMinutes()).padStart(2, '0')}.xlsx`;
                
                // Salvar arquivo
                XLSX.writeFile(wb, nomeArquivo);
                
                alert('✅ Excel gerado com sucesso!');
            } catch (error) {
                console.error('Erro ao gerar Excel:', error);
                alert('❌ Erro ao gerar Excel. Verifique o console para mais detalhes.');
            }
        }
    </script>
</body>
</html>
"""
        
        # Salvar HTML
        nome_arquivo = f'analise_produtos_{datetime.now().strftime("%Y%m%d_%H%M")}.html'
        html_path = Path(nome_arquivo)
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Abrir no navegador
        os.startfile(str(html_path.resolve()))
        
        print(f"\n✅ HTML gerado: {html_path.resolve()}")
        return str(html_path.resolve())
        
    except Exception as e:
        print(f"❌ Erro ao gerar HTML: {e}")
        import traceback
        traceback.print_exc()
        return None

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    print("=" * 100)
    print("📊 ANÁLISE COMPLETA DE PRODUTOS")
    print("=" * 100)
    print()
    
    # Solicita códigos dos produtos
    try:
        codigos_input = input("Digite os códigos dos produtos separados por vírgula (ex: 875,879,1234): ").strip()
        
        if not codigos_input:
            print("❌ Nenhum código informado.")
            exit(1)
        
        # Processa códigos
        codprod_list = []
        for cod in codigos_input.split(','):
            cod_limpo = cod.strip()
            if cod_limpo.isdigit():
                codprod_list.append(int(cod_limpo))
        
        if not codprod_list:
            print("❌ Nenhum código válido encontrado.")
            exit(1)
        
        print()
        print(f"🔍 Analisando {len(codprod_list)} produto(s): {codprod_list}")
        print()
        
    except ValueError:
        print("❌ Entrada inválida.")
        exit(1)
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        exit(1)
    
    try:
        # 1. Busca dados dos produtos
        print("📦 Buscando dados dos produtos...")
        df_produtos = buscar_dados_produtos(conn, codprod_list)
        
        if df_produtos.empty:
            print("⚠️  Nenhum produto encontrado com os códigos informados.")
            exit(0)
        
        print(f"✅ {len(df_produtos)} produto(s) encontrado(s).\n")
        
        # 2. Busca estoque da Vilog
        print("📊 Buscando estoque da Vilog...")
        estoque_vilog = get_vilog_stock_data()
        print(f"✅ {len(estoque_vilog)} produtos encontrados na Vilog.\n")
        
        # 3. Busca compras dos últimos 3 meses
        print("🛒 Buscando compras dos últimos 3 meses...")
        df_compras = buscar_compras_ultimos_3_meses(conn, codprod_list)
        print(f"✅ {len(df_compras)} registro(s) de compra encontrado(s).\n")
        
        # 4. Busca vendas dos últimos 12 meses
        print("💰 Buscando vendas dos últimos 12 meses (análise de lucro/prejuízo)...")
        df_vendas = buscar_vendas_ultimos_12_meses(conn, codprod_list)
        print(f"✅ {len(df_vendas)} registro(s) de venda encontrado(s).\n")
        
        # Processa cada produto
        print("=" * 120)
        print("📋 ANÁLISE POR PRODUTO")
        print("=" * 120)
        
        for _, produto in df_produtos.iterrows():
            codprod = produto['codprod']
            codprod_str = str(codprod)
            
            print()
            print("=" * 120)
            print(f"📦 PRODUTO: {codprod} - {produto['descricao']}")
            print("=" * 120)
            
            # Dados básicos
            print(f"\n📋 DADOS BÁSICOS:")
            print(f"   CODFAB: {produto['codfab'] if pd.notna(produto['codfab']) else 'N/A'}")
            print(f"   EAN: {produto['ean'] if pd.notna(produto['ean']) else 'N/A'}")
            print(f"   Departamento: {produto['departamento'] if pd.notna(produto['departamento']) else 'N/A'}")
            print(f"   Embalagem: {produto['embalagem'] if pd.notna(produto['embalagem']) else 'N/A'}")
            print(f"   Custo Final: R$ {produto['custo_final']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(produto['custo_final']) else 'N/A')
            
            # Estoque
            estoque_cx_oracle = float(produto['estoque_cx_oracle']) if pd.notna(produto['estoque_cx_oracle']) else 0.0
            estoque_cx_vilog = float(estoque_vilog.get(codprod_str, 0))
            estoque_cx_total = estoque_cx_oracle + estoque_cx_vilog
            
            print(f"\n📊 ESTOQUE ATUAL:")
            print(f"   Oracle: {estoque_cx_oracle:.2f} caixas")
            print(f"   Vilog: {estoque_cx_vilog:.2f} caixas")
            print(f"   TOTAL: {estoque_cx_total:.2f} caixas")
            
            # Compras dos últimos 3 meses
            compras_produto = df_compras[df_compras['codprod'] == codprod].copy()
            compras_produto = compras_produto.sort_values('mes_data', ascending=False)
            
            print(f"\n🛒 COMPRAS DOS ÚLTIMOS 3 MESES:")
            if compras_produto.empty:
                print("   Nenhuma compra encontrada.")
            else:
                total_compras = 0
                for _, compra in compras_produto.iterrows():
                    mes = compra['mes_ref']
                    qt = compra['qt_comprada']
                    total_compras += qt
                    print(f"   {mes}: {qt:.2f} unidades")
                print(f"   TOTAL: {total_compras:.2f} unidades")
            
            # Análise de lucro/prejuízo dos últimos 12 meses
            vendas_produto = df_vendas[df_vendas['codprod'] == codprod].copy()
            vendas_produto = vendas_produto.sort_values('mes_data', ascending=False)
            
            print(f"\n💰 ANÁLISE DE LUCRO/PREJUÍZO (ÚLTIMOS 12 MESES):")
            if vendas_produto.empty:
                print("   Nenhuma venda encontrada.")
            else:
                # Calcula totais
                faturamento_total = vendas_produto['faturamento_liquido'].sum()
                custo_total = vendas_produto['custo_liquido'].sum()
                lucro_prejuizo = faturamento_total - custo_total
                margem = (lucro_prejuizo / faturamento_total * 100) if faturamento_total > 0 else 0
                
                print(f"   Faturamento Total: R$ {faturamento_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                print(f"   Custo Total: R$ {custo_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                print(f"   {'✅ LUCRO' if lucro_prejuizo >= 0 else '❌ PREJUÍZO'}: R$ {abs(lucro_prejuizo):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                print(f"   Margem: {margem:.2f}%")
                
                # Mostra meses com prejuízo
                meses_prejuizo = vendas_produto[vendas_produto['faturamento_liquido'] < vendas_produto['custo_liquido']]
                if len(meses_prejuizo) > 0:
                    print(f"\n   ⚠️  Meses com PREJUÍZO ({len(meses_prejuizo)}):")
                    for _, mes in meses_prejuizo.iterrows():
                        prejuizo = mes['custo_liquido'] - mes['faturamento_liquido']
                        print(f"      {mes['mes_ref']}: R$ {prejuizo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                
                # Mostra últimos 6 meses detalhados
                print(f"\n   📅 ÚLTIMOS 6 MESES DETALHADOS:")
                for _, venda in vendas_produto.head(6).iterrows():
                    mes = venda['mes_ref']
                    fat = venda['faturamento_liquido']
                    cust = venda['custo_liquido']
                    lucro = fat - cust
                    qt = venda['qt_vendida']
                    status = "✅" if lucro >= 0 else "❌"
                    print(f"      {status} {mes}: Fat=R$ {fat:,.2f}, Custo=R$ {cust:,.2f}, {'Lucro' if lucro >= 0 else 'Prejuízo'}=R$ {abs(lucro):,.2f}, Qtd={qt:.0f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        print()
        print("=" * 120)
        print("✅ ANÁLISE CONCLUÍDA!")
        print("=" * 120)
        
        # Pergunta se deseja gerar HTML
        print()
        gerar_html_input = input("📄 Deseja gerar um relatório HTML com a análise? (s/n): ").strip().lower()
        
        if gerar_html_input in ['s', 'sim', 'y', 'yes']:
            print("\n📄 Gerando relatório HTML...")
            gerar_html(df_produtos, df_compras, df_vendas, estoque_vilog)
        
    except Exception as e:
        print(f"❌ Erro ao executar análise: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# portfolio-commit-ready: analise_produtos_completa
