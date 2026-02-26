"""
Relatório de produtos por departamento
- Filtra produtos com CODCATEGORIA 100 e 101
- Mostra: codprod, nomeprodut, categoria, estoque total cx, média mensal venda
- Separado por departamento
- HTML com botão de imprimir
"""

import oracledb
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
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

# Categorias a buscar
CODCATEGORIA_LIST = [100, 101]

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

def buscar_produtos_categorias(conn):
    """Busca produtos com CODCATEGORIA 100 (CONGELADO), 101 (RESFRIADO), 102 (SECO) ou sem categoria"""
    codcategoria_str = ', '.join(map(str, CODCATEGORIA_LIST))
    
    query = f"""
    SELECT 
        P.CODPROD,
        P.DESCRICAO AS NOMEPRODUT,
        CASE 
            WHEN P.CODCATEGORIA = 100 THEN 'CONGELADO'
            WHEN P.CODCATEGORIA = 101 THEN 'RESFRIADO'
            WHEN P.CODCATEGORIA IS NULL THEN 'SEM CATEGORIA'
            ELSE NVL(CAT.CATEGORIA, 'OUTRO')
        END AS CATEGORIA,
        P.CODCATEGORIA,
        D.DESCRICAO AS DEPARTAMENTO,
        D.CODEPTO,
        P.QTUNITCX,
        COALESCE((E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA) / NULLIF(P.QTUNITCX, 1), 
                 (E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA), 0) AS ESTOQUE_CX_ORACLE
    FROM PCPRODUT P
    LEFT JOIN PCCATEGORIA CAT ON P.CODCATEGORIA = CAT.CODCATEGORIA AND P.CODSEC = CAT.CODSEC
    LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
    LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
    WHERE (P.CODCATEGORIA IN ({codcategoria_str}) OR P.CODCATEGORIA IS NULL)
      AND P.DTEXCLUSAO IS NULL
      AND (P.OBS2 <> 'FL' OR P.OBS2 IS NULL)
      AND P.REVENDA = 'S'
    ORDER BY D.DESCRICAO, P.DESCRICAO
    """
    
    df = pd.read_sql_query(query, conn)
    df.columns = [x.lower() for x in df.columns]
    return df

def buscar_vendas_ultimos_12_meses(conn, codprod_list):
    """Busca vendas dos últimos 12 meses para calcular média mensal"""
    if not codprod_list:
        return pd.DataFrame()
    
    # Calcula data de 12 meses atrás
    data_inicio = datetime.now() - relativedelta(months=12)
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = datetime.now().strftime('%d/%m/%Y')
    
    # Divide a lista em chunks de 1000 (limite do Oracle)
    chunk_size = 1000
    chunks = [codprod_list[i:i + chunk_size] for i in range(0, len(codprod_list), chunk_size)]
    
    dfs = []
    for idx, chunk in enumerate(chunks):
        codprod_str = ', '.join(map(str, chunk))
        
        query = f"""
        SELECT 
            I.CODPROD,
            TO_CHAR(TRUNC(C.DATA, 'MM'), 'MM/YYYY') AS MES_REF,
            TRUNC(C.DATA, 'MM') AS MES_DATA,
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
        ORDER BY I.CODPROD, TRUNC(C.DATA, 'MM') DESC
        """
        
        df_chunk = pd.read_sql_query(query, conn)
        dfs.append(df_chunk)
        
        if len(chunks) > 1:
            print(f"   Processando lote {idx + 1}/{len(chunks)} ({len(chunk)} produtos)...")
    
    # Concatena todos os DataFrames
    if dfs:
        df = pd.concat(dfs, ignore_index=True)
        df.columns = [x.lower() for x in df.columns]
        return df
    else:
        return pd.DataFrame()

def format_numero_br(numero, decimais=2):
    """Formata número com pontuação brasileira"""
    if pd.isna(numero) or numero == 0:
        return "0" if decimais == 0 else "0," + "0" * decimais
    
    if decimais == 0:
        return f"{int(numero):,}".replace(',', '.')
    else:
        format_str = f"{{:,.{decimais}f}}"
        formatted = format_str.format(float(numero))
        return formatted.replace(',', 'X').replace('.', ',').replace('X', '.')

def gerar_html(df_produtos, df_vendas, estoque_vilog):
    """Gera HTML separado por departamento"""
    try:
        # Processar estoque Vilog (já vem em caixas da coluna E)
        # Mas vamos usar a coluna J (PESO TOTAL) se disponível, ou coluna E (quantidade em cx)
        estoque_vilog_cx = {}
        for codprod_str, valor in estoque_vilog.items():
            # get_vilog_stock_data retorna quantidade em caixas da coluna E
            estoque_vilog_cx[codprod_str] = float(valor) if valor else 0
        
        # Processar vendas para calcular média mensal
        vendas_por_produto = {}
        if not df_vendas.empty:
            for _, row in df_vendas.iterrows():
                codprod = int(row['codprod'])
                qt_vendida = float(row['qt_vendida'])
                
                if codprod not in vendas_por_produto:
                    vendas_por_produto[codprod] = []
                vendas_por_produto[codprod].append(qt_vendida)
        
        # Calcular média mensal de vendas em caixas para cada produto
        media_vendas_por_produto = {}
        for codprod, vendas_mensais in vendas_por_produto.items():
            # Busca qtunitcx do produto
            produto_info = df_produtos[df_produtos['codprod'] == codprod]
            if not produto_info.empty:
                qtunitcx = float(produto_info.iloc[0]['qtunitcx']) if pd.notna(produto_info.iloc[0]['qtunitcx']) and float(produto_info.iloc[0]['qtunitcx']) > 0 else 1.0
                
                # Converte vendas mensais (unidades) para caixas e calcula média
                vendas_cx = [v / qtunitcx for v in vendas_mensais if v > 0]
                if vendas_cx:
                    media_vendas_por_produto[codprod] = sum(vendas_cx) / len(vendas_cx)
                else:
                    media_vendas_por_produto[codprod] = 0
            else:
                media_vendas_por_produto[codprod] = 0
        
        # Agrupar produtos por departamento
        produtos_por_depto = {}
        for _, produto in df_produtos.iterrows():
            depto = produto['departamento'] if pd.notna(produto['departamento']) else 'SEM DEPARTAMENTO'
            codprod = int(produto['codprod'])
            codprod_str = str(codprod)
            
            # Estoque total em caixas
            estoque_cx_oracle = float(produto['estoque_cx_oracle']) if pd.notna(produto['estoque_cx_oracle']) else 0
            estoque_cx_vilog = estoque_vilog_cx.get(codprod_str, 0)
            estoque_total_cx = estoque_cx_oracle + estoque_cx_vilog
            
            # Média mensal de vendas em caixas
            media_mensal_cx = media_vendas_por_produto.get(codprod, 0)
            
            # Filtrar: só incluir se média mensal >= 5 caixas
            if media_mensal_cx < 5:
                continue
            
            if depto not in produtos_por_depto:
                produtos_por_depto[depto] = []
            
            produtos_por_depto[depto].append({
                'codprod': codprod,
                'nomeprodut': produto['nomeprodut'],
                'categoria': produto['categoria'] if pd.notna(produto['categoria']) else 'N/A',
                'estoque_total_cx': estoque_total_cx,
                'media_mensal_cx': media_mensal_cx
            })
        
        # Ordenar produtos dentro de cada departamento por média mensal de venda (maior para menor)
        for depto in produtos_por_depto.keys():
            produtos_por_depto[depto].sort(key=lambda x: x['media_mensal_cx'], reverse=True)
        
        # Calcular giro total de cada departamento (soma das médias mensais)
        depto_giro = {}
        for depto, produtos in produtos_por_depto.items():
            giro_total = sum(p['media_mensal_cx'] for p in produtos)
            depto_giro[depto] = giro_total
        
        # Ordenar departamentos por giro total (maior para menor)
        deptos_ordenados = sorted(produtos_por_depto.keys(), key=lambda d: depto_giro[d], reverse=True)
        
        # Gerar HTML
        html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Produtos - Categorias 100 e 101</title>
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
        
        .print-button {{
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
        
        .print-button:hover {{
            background-color: #218838;
        }}
        
        .print-button:active {{
            transform: scale(0.98);
        }}
        
        .depto-section {{
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 30px;
            overflow: hidden;
        }}
        
        .depto-header {{
            background: linear-gradient(135deg, #1a5a9a 0%, #124377 100%);
            color: white;
            padding: 20px;
            font-size: 22px;
            font-weight: 600;
        }}
        
        .table-container {{
            overflow-x: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        
        thead {{
            background: #f8f9fa;
        }}
        
        th {{
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #dee2e6;
            color: #124377;
        }}
        
        td {{
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
        }}
        
        tbody tr:hover {{
            background-color: #f8f9fa;
        }}
        
        .codigo {{
            font-weight: 600;
            color: #124377;
        }}
        
        .numero {{
            text-align: right;
        }}
        
        @media print {{
            body {{
                background-color: white;
                padding: 10px;
            }}
            
            .print-button {{
                display: none;
            }}
            
            .header {{
                background: linear-gradient(135deg, #124377 0%, #1a5a9a 100%) !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
                color: white !important;
                margin-bottom: 20px;
                padding: 20px;
            }}
            
            .depto-section {{
                page-break-inside: avoid;
                margin-bottom: 25px;
                border: 1px solid #ddd;
            }}
            
            .depto-header {{
                background: linear-gradient(135deg, #1a5a9a 0%, #124377 100%) !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
                color: white !important;
                padding: 15px;
            }}
            
            thead {{
                background: #f8f9fa !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            th {{
                background: #f8f9fa !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
                color: #124377 !important;
                border-bottom: 2px solid #dee2e6 !important;
                padding: 10px 8px;
            }}
            
            td {{
                border-bottom: 1px solid #e0e0e0 !important;
                padding: 8px;
            }}
            
            tbody tr:nth-child(even) {{
                background-color: #f8f9fa !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            .codigo {{
                color: #124377 !important;
                font-weight: 600;
            }}
            
            table {{
                border-collapse: collapse;
                width: 100%;
            }}
            
            @page {{
                margin: 1.5cm;
                size: A4 landscape;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Relatório de Produtos - Categorias 100 e 101</h1>
        <p>Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
        <p>Categorias: 100 (CONGELADO), 101 (RESFRIADO) e SEM CATEGORIA</p>
    </div>
    
    <button class="print-button" onclick="window.print()">🖨️ Imprimir</button>
"""
        
        # Adiciona cada departamento
        for depto in deptos_ordenados:
            produtos = produtos_por_depto[depto]
            
            html_content += f"""
    <div class="depto-section">
        <div class="depto-header">{depto}</div>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Código</th>
                        <th>Nome do Produto</th>
                        <th>Categoria</th>
                        <th class="numero">Estoque Total (cx)</th>
                        <th class="numero">Média Mensal Venda (cx)</th>
                    </tr>
                </thead>
                <tbody>
"""
            
            for produto in produtos:
                html_content += f"""
                    <tr>
                        <td class="codigo">{produto['codprod']}</td>
                        <td>{produto['nomeprodut']}</td>
                        <td>{produto['categoria']}</td>
                        <td class="numero">{format_numero_br(produto['estoque_total_cx'], 2)} cx</td>
                        <td class="numero">{format_numero_br(produto['media_mensal_cx'], 2)} cx</td>
                    </tr>
"""
            
            html_content += """
                </tbody>
            </table>
        </div>
    </div>
"""
        
        html_content += """
</body>
</html>
"""
        
        # Salvar HTML
        nome_arquivo = f'relatorio_categorias_100_101_{datetime.now().strftime("%Y%m%d_%H%M")}.html'
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
    print("📊 RELATÓRIO DE PRODUTOS - CATEGORIAS 100 E 101")
    print("=" * 100)
    print()
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        exit(1)
    
    try:
        # 1. Busca produtos das categorias 100 e 101
        print("📦 Buscando produtos das categorias 100 e 101...")
        df_produtos = buscar_produtos_categorias(conn)
        
        if df_produtos.empty:
            print("⚠️  Nenhum produto encontrado com as categorias especificadas.")
            exit(0)
        
        print(f"✅ {len(df_produtos)} produto(s) encontrado(s).\n")
        
        # 2. Busca estoque da Vilog
        print("📊 Buscando estoque da Vilog...")
        estoque_vilog = get_vilog_stock_data()
        print(f"✅ {len(estoque_vilog)} produtos encontrados na Vilog.\n")
        
        # 3. Busca vendas dos últimos 12 meses
        codprod_list = df_produtos['codprod'].unique().tolist()
        print(f"💰 Buscando vendas dos últimos 12 meses para {len(codprod_list)} produtos...")
        df_vendas = buscar_vendas_ultimos_12_meses(conn, codprod_list)
        print(f"✅ {len(df_vendas)} registro(s) de venda encontrado(s).\n")
        
        # 4. Gera HTML
        print("📄 Gerando relatório HTML...")
        gerar_html(df_produtos, df_vendas, estoque_vilog)
        
        print()
        print("=" * 100)
        print("✅ RELATÓRIO CONCLUÍDO!")
        print("=" * 100)
        
    except Exception as e:
        print(f"❌ Erro ao executar análise: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# portfolio-commit-ready: relatorio_categorias_100_101
