"""
Script para gerar todas as tabelas Fato e Dimensões para Power BI
Gera CSVs na pasta DADOS_BI_VENDAS com estrutura de modelo estrela
"""

import oracledb
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import os

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

# Pasta de saída
PASTA_SAIDA = Path("DADOS_BI_VENDAS")

# Período de dados (anos para buscar)
ANOS_HISTORICO = 3  # Busca últimos 3 anos

# =========================================================================
# FUNÇÕES DE CONEXÃO
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

# =========================================================================
# FUNÇÕES PARA DIMENSÕES
# =========================================================================
def gerar_dim_calendario(data_inicio, data_fim):
    """
    Gera tabela de calendário (dim_calendario)
    Parâmetros: data_inicio e data_fim (datetime)
    """
    print("📅 Gerando dim_calendario...")
    
    datas = pd.date_range(start=data_inicio, end=data_fim, freq='D')
    
    # Cria DataFrame com todas as colunas
    df_calendario = pd.DataFrame({
        'data': datas,
        'data_key': datas.strftime('%Y%m%d').astype(int),
        'ano': datas.year,
        'mes': datas.month,
        'dia': datas.day,
        'nome_mes': datas.strftime('%B'),
        'nome_mes_abrev': datas.strftime('%b'),
        'trimestre': datas.quarter,
        'semestre': [1 if q <= 2 else 2 for q in datas.quarter],
        'ano_mes': datas.strftime('%Y-%m'),
        'ano_mes_num': datas.strftime('%Y%m').astype(int),
        'mes_ano': datas.strftime('%m/%Y'),
        'dia_semana': datas.dayofweek + 1,  # 1=Segunda, 7=Domingo
        'nome_dia_semana': datas.strftime('%A'),
        'nome_dia_semana_abrev': datas.strftime('%a'),
        'eh_fim_semana': (datas.dayofweek >= 5).astype(int),  # 1 se sábado ou domingo
        'eh_ultimo_dia_mes': (datas == datas.to_period('M').to_timestamp('M')).astype(int),
        'semana_ano': datas.isocalendar().week,
        'dia_ano': datas.dayofyear,
    })
    
    # Ordena por data
    df_calendario = df_calendario.sort_values('data').reset_index(drop=True)
    
    # Garante que data seja datetime
    df_calendario['data'] = pd.to_datetime(df_calendario['data'])
    
    # Garante tipos corretos
    df_calendario['data_key'] = df_calendario['data_key'].astype(int)
    df_calendario['ano'] = df_calendario['ano'].astype(int)
    df_calendario['mes'] = df_calendario['mes'].astype(int)
    df_calendario['dia'] = df_calendario['dia'].astype(int)
    df_calendario['trimestre'] = df_calendario['trimestre'].astype(int)
    df_calendario['semestre'] = df_calendario['semestre'].astype(int)
    df_calendario['ano_mes_num'] = df_calendario['ano_mes_num'].astype(int)
    df_calendario['dia_semana'] = df_calendario['dia_semana'].astype(int)
    df_calendario['eh_fim_semana'] = df_calendario['eh_fim_semana'].astype(int)
    df_calendario['eh_ultimo_dia_mes'] = df_calendario['eh_ultimo_dia_mes'].astype(int)
    df_calendario['semana_ano'] = df_calendario['semana_ano'].astype(int)
    df_calendario['dia_ano'] = df_calendario['dia_ano'].astype(int)
    
    print(f"   ✅ {len(df_calendario)} registros gerados")
    return df_calendario

def gerar_dim_clientes(conn):
    """Gera dimensão de clientes"""
    print("👥 Gerando dim_clientes...")
    
    query = """
    SELECT DISTINCT
        CL.CODCLI AS cliente_id,
        CL.CODCLI AS codcli,
        CL.CLIENTE AS nome_cliente,
        CL.CGCENT AS cnpj
    FROM PCCLIENT CL
    WHERE CL.CODCLI IS NOT NULL
    ORDER BY CL.CODCLI
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'cliente_id' in df.columns:
            df['cliente_id'] = pd.to_numeric(df['cliente_id'], errors='coerce').fillna(0).astype(int)
        if 'codcli' in df.columns:
            df['codcli'] = pd.to_numeric(df['codcli'], errors='coerce').fillna(0).astype(int)
        
        if df.empty:
            print(f"   ⚠️  Nenhum registro encontrado (tabela pode estar vazia)")
        else:
            print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro ao buscar clientes: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_dim_produtos(conn):
    """Gera dimensão de produtos"""
    print("📦 Gerando dim_produtos...")
    
    query = """
    SELECT DISTINCT
        P.CODPROD AS produto_id,
        P.CODPROD AS codprod,
        P.DESCRICAO AS descricao_produto,
        P.UNIDADE AS unidade,
        P.EMBALAGEM AS embalagem,
        NVL(P.QTUNITCX, 1) AS qt_unit_cx,
        P.CODEPTO AS cod_depto,
        D.DESCRICAO AS descricao_depto,
        P.CODSEC AS cod_secao,
        P.CODCATEGORIA AS cod_categoria,
        CAT.CATEGORIA AS nome_categoria,
        P.CODFORNEC AS cod_fornecedor,
        F.FORNECEDOR AS nome_fornecedor,
        P.CODFAB AS cod_fabricante,
        P.OBS AS observacao
    FROM PCPRODUT P
    LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
    LEFT JOIN PCCATEGORIA CAT ON P.CODCATEGORIA = CAT.CODCATEGORIA AND P.CODSEC = CAT.CODSEC
    LEFT JOIN PCFORNEC F ON P.CODFORNEC = F.CODFORNEC
    WHERE P.CODPROD IS NOT NULL
    ORDER BY P.CODPROD
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'produto_id' in df.columns:
            df['produto_id'] = pd.to_numeric(df['produto_id'], errors='coerce').fillna(0).astype(int)
        if 'codprod' in df.columns:
            df['codprod'] = pd.to_numeric(df['codprod'], errors='coerce').fillna(0).astype(int)
        if 'cod_depto' in df.columns:
            df['cod_depto'] = pd.to_numeric(df['cod_depto'], errors='coerce').fillna(0).astype(int)
        if 'qt_unit_cx' in df.columns:
            df['qt_unit_cx'] = pd.to_numeric(df['qt_unit_cx'], errors='coerce').fillna(1.0).astype(float)
        
        if df.empty:
            print(f"   ⚠️  Nenhum registro encontrado (tabela pode estar vazia)")
        else:
            print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro ao buscar produtos: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_dim_vendedores(conn):
    """Gera dimensão de vendedores"""
    print("👤 Gerando dim_vendedores...")
    
    query = """
    SELECT DISTINCT
        U.CODUSUR AS vendedor_id,
        U.CODUSUR AS codusur,
        U.NOME AS nome_vendedor,
        U.CODSUPERVISOR AS cod_supervisor,
        S.NOME AS nome_supervisor
    FROM PCUSUARI U
    LEFT JOIN PCSUPERV S ON U.CODSUPERVISOR = S.CODSUPERVISOR
    WHERE U.CODUSUR IS NOT NULL
    ORDER BY U.CODUSUR
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'vendedor_id' in df.columns:
            df['vendedor_id'] = pd.to_numeric(df['vendedor_id'], errors='coerce').fillna(0).astype(int)
        if 'codusur' in df.columns:
            df['codusur'] = pd.to_numeric(df['codusur'], errors='coerce').fillna(0).astype(int)
        if 'cod_supervisor' in df.columns:
            df['cod_supervisor'] = pd.to_numeric(df['cod_supervisor'], errors='coerce').fillna(0).astype(int)
        
        if df.empty:
            print(f"   ⚠️  Nenhum registro encontrado (tabela pode estar vazia)")
        else:
            print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro ao buscar vendedores: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_dim_departamentos(conn):
    """Gera dimensão de departamentos"""
    print("🏢 Gerando dim_departamentos...")
    
    query = """
    SELECT DISTINCT
        D.CODEPTO AS depto_id,
        D.CODEPTO AS codepto,
        D.DESCRICAO AS descricao_depto
    FROM PCDEPTO D
    WHERE D.CODEPTO IS NOT NULL
    ORDER BY D.CODEPTO
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'depto_id' in df.columns:
            df['depto_id'] = pd.to_numeric(df['depto_id'], errors='coerce').fillna(0).astype(int)
        if 'codepto' in df.columns:
            df['codepto'] = pd.to_numeric(df['codepto'], errors='coerce').fillna(0).astype(int)
        
        if df.empty:
            print(f"   ⚠️  Nenhum registro encontrado (tabela pode estar vazia)")
        else:
            print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro ao buscar departamentos: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_dim_filiais(conn):
    """Gera dimensão de filiais baseada nos valores únicos de CODFILIAL"""
    print("🏪 Gerando dim_filiais...")
    
    # Busca filiais únicas da tabela PCPEDC
    query = """
    SELECT DISTINCT
        C.CODFILIAL AS filial_id,
        C.CODFILIAL AS codfilial,
        'Filial ' || C.CODFILIAL AS nome_filial
    FROM PCPEDC C
    WHERE C.CODFILIAL IS NOT NULL
    ORDER BY C.CODFILIAL
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'filial_id' in df.columns:
            df['filial_id'] = df['filial_id'].astype(str)
        if 'codfilial' in df.columns:
            df['codfilial'] = df['codfilial'].astype(str)
        
        if df.empty:
            print(f"   ⚠️  Nenhum registro encontrado (tabela pode estar vazia)")
        else:
            print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro ao buscar filiais: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# FUNÇÃO PARA FATO VENDAS
# =========================================================================
def gerar_fato_vendas(conn, data_inicio, data_fim):
    """
    Gera fato de vendas com granularidade de item de pedido
    Inclui devoluções e cálculos de faturamento líquido e margem
    Aplica todos os filtros padrão de vendas válidas
    """
    print("💰 Gerando fato_vendas...")
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    WITH VENDAS AS (
        SELECT 
            -- Chaves de relacionamento
            C.CODCLI AS cliente_id,
            I.CODPROD AS produto_id,
            C.CODUSUR AS vendedor_id,
            P.CODEPTO AS depto_id,
            C.CODFILIAL AS filial_id,
            
            -- Dimensão de tempo
            C.DATA AS data_venda,
            TO_NUMBER(TO_CHAR(C.DATA, 'YYYYMMDD')) AS data_key,
            
            -- Identificadores do pedido
            C.NUMPED AS num_pedido,
            C.NUMNOTA AS num_nota,
            
            -- Dados do item
            I.QT AS quantidade,
            I.PVENDA AS preco_unitario,
            ROUND(I.QT * I.PVENDA, 2) AS valor_bruto_item,
            I.VLCUSTOFIN AS custo_unitario,
            ROUND(I.QT * I.VLCUSTOFIN, 2) AS custo_total_item,
            ROUND((I.QT * I.PVENDA) - (I.QT * I.VLCUSTOFIN), 2) AS margem_item,
            
            -- Flags e classificações
            CASE WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 0 ELSE 1 END AS flag_bonificacao,
            C.CONDVENDA AS condicao_venda,
            C.POSICAO AS posicao_pedido,
            
            -- Valores de faturamento (sem DECODE - cálculo simples)
            -- O usuário vai calcular faturamento líquido no Power BI: valor_faturamento - valor_devolucao
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                ELSE 0 
            END AS valor_faturamento,
            
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    ROUND(NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0), 2)
                ELSE 0 
            END AS custo_faturamento,
            
            -- Outros campos úteis
            I.VLSUBTOTITEM AS valor_subtotal_item,
            I.VLOUTRASDESP AS outras_despesas,
            I.VLFRETE AS valor_frete,
            C.DTCANCEL AS data_cancelamento,
            CASE WHEN C.DTCANCEL IS NULL THEN 0 ELSE 1 END AS flag_cancelado
            
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
    ),
    DEVOLUCOES_AGREGADAS AS (
        -- Devoluções agregadas por vendedor e data (mesma lógica do analise_vendedores_gui.py)
        SELECT 
            D.CODUSUR AS vendedor_id,
            TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD')) AS data_key,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS valor_devolucao_total,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS custo_devolucao_total
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
        GROUP BY D.CODUSUR, TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD'))
    ),
    VENDAS_AGREGADAS AS (
        -- Agrega vendas por vendedor e data para calcular proporção
        SELECT 
            V.vendedor_id,
            V.data_key,
            SUM(V.valor_faturamento) AS valor_faturamento_total,
            SUM(V.custo_faturamento) AS custo_faturamento_total
        FROM VENDAS V
        GROUP BY V.vendedor_id, V.data_key
    )
    SELECT 
        V.*,
        -- Devoluções proporcionais ao item (baseado na participação do item no faturamento do vendedor na data)
        CASE 
            WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
            ELSE 0
        END AS valor_devolucao,
        CASE 
            WHEN NVL(VA.custo_faturamento_total, 0) > 0 THEN
                ROUND((V.custo_faturamento / VA.custo_faturamento_total) * NVL(DA.custo_devolucao_total, 0), 2)
            ELSE 0
        END AS custo_devolucao,
        -- Faturamento líquido (FAT - DEV) - calculado proporcionalmente
        CASE 
            WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                V.valor_faturamento - ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
            ELSE V.valor_faturamento
        END AS faturamento_liquido,
        -- Custo líquido
        CASE 
            WHEN NVL(VA.custo_faturamento_total, 0) > 0 THEN
                V.custo_faturamento - ROUND((V.custo_faturamento / VA.custo_faturamento_total) * NVL(DA.custo_devolucao_total, 0), 2)
            ELSE V.custo_faturamento
        END AS custo_liquido,
        -- Margem percentual (formato simples: 15 para 15%, 6.2 para 6.2%)
        CASE 
            WHEN (CASE 
                    WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                        V.valor_faturamento - ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
                    ELSE V.valor_faturamento
                  END) > 0 THEN
                ROUND(
                    ((CASE 
                        WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                            V.valor_faturamento - ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
                        ELSE V.valor_faturamento
                      END) - 
                     (CASE 
                        WHEN NVL(VA.custo_faturamento_total, 0) > 0 THEN
                            V.custo_faturamento - ROUND((V.custo_faturamento / VA.custo_faturamento_total) * NVL(DA.custo_devolucao_total, 0), 2)
                        ELSE V.custo_faturamento
                      END)) / 
                    (CASE 
                        WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                            V.valor_faturamento - ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
                        ELSE V.valor_faturamento
                      END) * 100, 
                    2
                )
            ELSE 0
        END AS margem_percentual,
        -- Supervisor (adicionado no final para não quebrar estrutura existente)
        U.CODSUPERVISOR AS cod_supervisor,
        S.NOME AS nome_supervisor
    FROM VENDAS V
    LEFT JOIN VENDAS_AGREGADAS VA ON V.vendedor_id = VA.vendedor_id AND V.data_key = VA.data_key
    LEFT JOIN DEVOLUCOES_AGREGADAS DA ON V.vendedor_id = DA.vendedor_id AND V.data_key = DA.data_key
    LEFT JOIN PCUSUARI U ON V.vendedor_id = U.CODUSUR
    LEFT JOIN PCSUPERV S ON U.CODSUPERVISOR = S.CODSUPERVISOR
    ORDER BY V.data_venda DESC, V.num_pedido, V.produto_id
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # ============================================================
        # FORMATAÇÃO DE TIPOS DE DADOS
        # ============================================================
        
        # DATAS - CRÍTICO!
        if 'data_venda' in df.columns:
            df['data_venda'] = pd.to_datetime(df['data_venda'], errors='coerce')
        
        if 'data_cancelamento' in df.columns:
            df['data_cancelamento'] = pd.to_datetime(df['data_cancelamento'], errors='coerce')
        
        # INTEIROS
        int_columns = ['data_key', 'cliente_id', 'produto_id', 'vendedor_id', 'cod_supervisor', 'depto_id', 
                      'filial_id', 'num_pedido', 'num_nota', 'quantidade', 'flag_bonificacao',
                      'condicao_venda', 'flag_cancelado']
        for col in int_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        # FLOATS (valores monetários e percentuais)
        float_columns = ['preco_unitario', 'valor_bruto_item', 'custo_unitario', 'custo_total_item',
                        'margem_item', 'valor_faturamento', 'custo_faturamento', 'valor_subtotal_item',
                        'outras_despesas', 'valor_frete', 'valor_devolucao', 'custo_devolucao',
                        'faturamento_liquido', 'custo_liquido', 'margem_percentual']
        for col in float_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
        
        # STRINGS
        if 'posicao_pedido' in df.columns:
            df['posicao_pedido'] = df['posicao_pedido'].astype(str)
        
        if 'nome_supervisor' in df.columns:
            df['nome_supervisor'] = df['nome_supervisor'].astype(str).fillna('')
        
        print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_fato_vendas_produto(conn, data_inicio, data_fim):
    """
    Gera fato de vendas agregado por produto
    Focado em produto para cálculo correto de margem
    """
    print("📦 Gerando fato_vendas_produto...")
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    WITH VENDAS AS (
        SELECT 
            I.CODPROD AS produto_id,
            P.CODEPTO AS depto_id,
            TO_NUMBER(TO_CHAR(C.DATA, 'YYYYMMDD')) AS data_key,
            C.DATA AS data_venda,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                    ELSE 0 
                END
            ) AS valor_faturamento,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        ROUND(NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0), 2)
                    ELSE 0 
                END
            ) AS custo_faturamento,
            SUM(NVL(I.QT, 0)) AS quantidade_total
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
        GROUP BY I.CODPROD, P.CODEPTO, TO_NUMBER(TO_CHAR(C.DATA, 'YYYYMMDD')), C.DATA
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD AS produto_id,
            TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD')) AS data_key,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS valor_devolucao,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS custo_devolucao
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
        GROUP BY D.CODPROD, TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD'))
    )
    SELECT 
        V.produto_id,
        V.depto_id,
        V.data_key,
        V.data_venda,
        V.quantidade_total,
        V.valor_faturamento,
        V.custo_faturamento,
        NVL(D.valor_devolucao, 0) AS valor_devolucao,
        NVL(D.custo_devolucao, 0) AS custo_devolucao,
        (V.valor_faturamento - NVL(D.valor_devolucao, 0)) AS faturamento_liquido,
        (V.custo_faturamento - NVL(D.custo_devolucao, 0)) AS custo_liquido,
        CASE
            WHEN (V.valor_faturamento - NVL(D.valor_devolucao, 0)) > 0 THEN
                ROUND(
                    ((V.valor_faturamento - NVL(D.valor_devolucao, 0)) - 
                     (V.custo_faturamento - NVL(D.custo_devolucao, 0))) /
                    (V.valor_faturamento - NVL(D.valor_devolucao, 0)) * 100,
                    2
                )
            ELSE 0
        END AS margem_percentual
    FROM VENDAS V
    LEFT JOIN DEVOLUCOES D ON V.produto_id = D.produto_id AND V.data_key = D.data_key
    ORDER BY V.data_venda DESC, V.produto_id
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'data_venda' in df.columns:
            df['data_venda'] = pd.to_datetime(df['data_venda'], errors='coerce')
        
        int_columns = ['data_key', 'produto_id', 'depto_id', 'quantidade_total']
        for col in int_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        float_columns = ['valor_faturamento', 'custo_faturamento', 'valor_devolucao', 'custo_devolucao',
                        'faturamento_liquido', 'custo_liquido', 'margem_percentual']
        for col in float_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
        
        print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_fato_devolucoes(conn, data_inicio, data_fim):
    """
    Gera tabela apenas com devoluções
    """
    print("🔄 Gerando fato_devolucoes...")
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    SELECT 
        D.CODCLI AS cliente_id,
        D.CODPROD AS produto_id,
        D.CODUSUR AS vendedor_id,
        U.CODSUPERVISOR AS cod_supervisor,
        P.CODEPTO AS depto_id,
        TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD')) AS data_key,
        D.DTENT AS data_devolucao,
        NVL(D.VLDEVOLUCAO, 0) AS valor_devolucao,
        NVL(D.VLCUSTOFIN, 0) AS custo_devolucao,
        NVL(D.QT, 0) AS quantidade_devolvida
    FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
    LEFT JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
    LEFT JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
    WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
      AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
    ORDER BY D.DTENT DESC, D.CODPROD
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'data_devolucao' in df.columns:
            df['data_devolucao'] = pd.to_datetime(df['data_devolucao'], errors='coerce')
        
        int_columns = ['data_key', 'cliente_id', 'produto_id', 'vendedor_id', 'cod_supervisor', 'depto_id', 'quantidade_devolvida']
        for col in int_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        float_columns = ['valor_devolucao', 'custo_devolucao']
        for col in float_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
        
        print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def gerar_fato_custos(conn, data_inicio, data_fim):
    """
    Gera tabela apenas com custos das vendas
    """
    print("💰 Gerando fato_custos...")
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    SELECT 
        C.CODCLI AS cliente_id,
        I.CODPROD AS produto_id,
        C.CODUSUR AS vendedor_id,
        U.CODSUPERVISOR AS cod_supervisor,
        P.CODEPTO AS depto_id,
        C.CODFILIAL AS filial_id,
        TO_NUMBER(TO_CHAR(C.DATA, 'YYYYMMDD')) AS data_key,
        C.DATA AS data_venda,
        I.QT AS quantidade,
        I.VLCUSTOFIN AS custo_unitario,
        ROUND(I.QT * I.VLCUSTOFIN, 2) AS custo_total,
        I.PVENDA AS preco_unitario,
        ROUND(I.QT * I.PVENDA, 2) AS valor_total
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
      AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
      AND NVL(I.BONIFIC, 'N') = 'N'
    ORDER BY C.DATA DESC, I.CODPROD
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Formatação de tipos
        if 'data_venda' in df.columns:
            df['data_venda'] = pd.to_datetime(df['data_venda'], errors='coerce')
        
        int_columns = ['data_key', 'cliente_id', 'produto_id', 'vendedor_id', 'cod_supervisor', 'depto_id', 'filial_id', 'quantidade']
        for col in int_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        float_columns = ['custo_unitario', 'custo_total', 'preco_unitario', 'valor_total']
        for col in float_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
        
        print(f"   ✅ {len(df)} registros encontrados")
        return df
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# FUNÇÃO PRINCIPAL
# =========================================================================
def main():
    """Função principal que gera todas as tabelas"""
    print("=" * 100)
    print("🚀 GERADOR DE DADOS PARA POWER BI - MODELO ESTRELA")
    print("=" * 100)
    print()
    
    # Cria pasta de saída
    PASTA_SAIDA.mkdir(exist_ok=True)
    print(f"📁 Pasta de saída: {PASTA_SAIDA.absolute()}")
    print()
    
    # Define período de dados
    hoje = datetime.now()
    data_fim = hoje.replace(day=1)  # Primeiro dia do mês atual
    data_inicio = data_fim.replace(year=data_fim.year - ANOS_HISTORICO)  # 3 anos atrás
    
    print(f"📅 Período de dados:")
    print(f"   Início: {data_inicio.strftime('%d/%m/%Y')}")
    print(f"   Fim: {data_fim.strftime('%d/%m/%Y')}")
    print()
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        return
    
    try:
        print("✅ Conectado!")
        print()
        
        # Gera dimensões
        print("=" * 100)
        print("📊 GERANDO DIMENSÕES")
        print("=" * 100)
        print()
        
        # Gera calendário primeiro (usa as mesmas datas do período de vendas)
        df_calendario = gerar_dim_calendario(data_inicio, data_fim)
        df_clientes = gerar_dim_clientes(conn)
        df_produtos = gerar_dim_produtos(conn)
        df_vendedores = gerar_dim_vendedores(conn)
        df_departamentos = gerar_dim_departamentos(conn)
        df_filiais = gerar_dim_filiais(conn)
        
        print()
        
        # Gera fato
        print("=" * 100)
        print("💰 GERANDO FATO")
        print("=" * 100)
        print()
        
        df_vendas = gerar_fato_vendas(conn, data_inicio, data_fim)
        
        print()
        
        # Gera novas tabelas fato
        print("=" * 100)
        print("📊 GERANDO TABELAS FATO ADICIONAIS")
        print("=" * 100)
        print()
        
        df_vendas_produto = gerar_fato_vendas_produto(conn, data_inicio, data_fim)
        df_devolucoes = gerar_fato_devolucoes(conn, data_inicio, data_fim)
        df_custos = gerar_fato_custos(conn, data_inicio, data_fim)
        
        print()
        
        # Salva todos os CSVs
        print("=" * 100)
        print("💾 SALVANDO ARQUIVOS CSV")
        print("=" * 100)
        print()
        
        arquivos_gerados = []
        
        # Salva dimensões
        if not df_calendario.empty:
            arquivo = PASTA_SAIDA / "dim_calendario.csv"
            df_calendario.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',', date_format='%Y-%m-%d')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_calendario)} registros")
        
        if not df_clientes.empty:
            arquivo = PASTA_SAIDA / "dim_clientes.csv"
            df_clientes.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_clientes)} registros")
        
        if not df_produtos.empty:
            arquivo = PASTA_SAIDA / "dim_produtos.csv"
            df_produtos.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_produtos)} registros")
        
        if not df_vendedores.empty:
            arquivo = PASTA_SAIDA / "dim_vendedores.csv"
            df_vendedores.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_vendedores)} registros")
        
        if not df_departamentos.empty:
            arquivo = PASTA_SAIDA / "dim_departamentos.csv"
            df_departamentos.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_departamentos)} registros")
        
        if not df_filiais.empty:
            arquivo = PASTA_SAIDA / "dim_filiais.csv"
            df_filiais.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_filiais)} registros")
        
        # Salva fato
        if not df_vendas.empty:
            arquivo = PASTA_SAIDA / "fato_vendas.csv"
            df_vendas.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',', date_format='%Y-%m-%d')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_vendas)} registros")
        
        # Salva tabelas fato adicionais
        if not df_vendas_produto.empty:
            arquivo = PASTA_SAIDA / "fato_vendas_produto.csv"
            df_vendas_produto.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',', date_format='%Y-%m-%d')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_vendas_produto)} registros")
        
        if not df_devolucoes.empty:
            arquivo = PASTA_SAIDA / "fato_devolucoes.csv"
            df_devolucoes.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',', date_format='%Y-%m-%d')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_devolucoes)} registros")
        
        if not df_custos.empty:
            arquivo = PASTA_SAIDA / "fato_custos.csv"
            df_custos.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',', date_format='%Y-%m-%d')
            arquivos_gerados.append(arquivo)
            print(f"✅ {arquivo.name} - {len(df_custos)} registros")
        
        print()
        print("=" * 100)
        print("✅ PROCESSO CONCLUÍDO!")
        print("=" * 100)
        print()
        print(f"📁 Total de arquivos gerados: {len(arquivos_gerados)}")
        print(f"📂 Pasta: {PASTA_SAIDA.absolute()}")
        print()
        print("📋 Arquivos gerados:")
        for arquivo in arquivos_gerados:
            tamanho_mb = arquivo.stat().st_size / (1024 * 1024)
            print(f"   • {arquivo.name} ({tamanho_mb:.2f} MB)")
        
        # Resumo de tabelas que não foram geradas
        tabelas_esperadas = {
            'dim_calendario': df_calendario,
            'dim_clientes': df_clientes,
            'dim_produtos': df_produtos,
            'dim_vendedores': df_vendedores,
            'dim_departamentos': df_departamentos,
            'dim_filiais': df_filiais,
            'fato_vendas': df_vendas,
            'fato_vendas_produto': df_vendas_produto,
            'fato_devolucoes': df_devolucoes,
            'fato_custos': df_custos
        }
        
        tabelas_faltando = [nome for nome, df in tabelas_esperadas.items() if df.empty]
        if tabelas_faltando:
            print()
            print("⚠️  ATENÇÃO: As seguintes tabelas não foram geradas (vazias ou com erro):")
            for tabela in tabelas_faltando:
                print(f"   • {tabela}.csv")
            print()
            print("💡 Verifique os erros acima para identificar o problema.")
        
    except Exception as e:
        print(f"❌ Erro durante o processo: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print()
        print("🔌 Conexão fechada.")

if __name__ == "__main__":
    main()

# portfolio-commit-ready: gerar_dados_bi_vendas
