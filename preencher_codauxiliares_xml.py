import pandas as pd
import oracledb
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

# =========================================================================
# CONFIGURAÇÕES DE BANCO
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# Namespace do XML de NF-e
NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}


def get_db_connection():
    """Conecta ao banco Oracle."""
    try:
        try:
            oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
        except Exception:
            pass
        
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        print(f"Erro ao conectar no Oracle: {e}")
        return None


def buscar_produtos_vazios(conn):
    """
    Busca produtos com CODAUXILIAR, CODAUXILIAR2 e CODAUXILIARTRIB zerados
    que tiveram venda nos últimos 6 meses.
    """
    query = """
        WITH VENDAS_6M AS (
            SELECT
                I.CODPROD,
                SUM(NVL(I.QT, 0)) AS QT_VENDIDA_6M
            FROM PCPEDC C
            JOIN PCPEDI I ON C.NUMPED = I.NUMPED
            WHERE C.DATA >= ADD_MONTHS(TRUNC(SYSDATE), -6)
              AND C.POSICAO = 'F'
              AND C.DTCANCEL IS NULL
              AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
            GROUP BY I.CODPROD
        )
        SELECT
            P.CODPROD,
            P.CODFAB,
            P.DESCRICAO,
            P.CODEPTO,
            D.DESCRICAO AS DEPARTAMENTO,
            P.CODFORNEC,
            F.FORNECEDOR,
            P.EMBALAGEM,
            P.UNIDADE,
            P.CODAUXILIAR,
            P.CODAUXILIAR2,
            P.CODAUXILIARTRIB,
            V.QT_VENDIDA_6M
        FROM PCPRODUT P
        JOIN VENDAS_6M V ON V.CODPROD = P.CODPROD
        LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        LEFT JOIN PCFORNEC F ON P.CODFORNEC = F.CODFORNEC
        WHERE P.DTEXCLUSAO IS NULL
          AND NVL(P.CODAUXILIAR, 0) = 0
          AND NVL(P.CODAUXILIAR2, 0) = 0
          AND NVL(P.CODAUXILIARTRIB, 0) = 0
        ORDER BY V.QT_VENDIDA_6M DESC, P.DESCRICAO
    """
    
    print("Buscando produtos com códigos auxiliares zerados no Oracle...")
    df = pd.read_sql_query(query, conn)
    print(f"Encontrados {len(df)} produtos com códigos auxiliares zerados.")
    return df


def extrair_dados_xml(xml_path):
    """
    Extrai dados de produtos de um XML de NF-e.
    Retorna lista de dicionários com: codfab, cEAN, cEANTrib, descricao, xml_file
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Nome do arquivo XML (sem caminho)
        xml_filename = Path(xml_path).name
        
        produtos = []
        
        # Busca todos os itens (det)
        for det in root.findall('.//nfe:det', NS):
            prod = det.find('nfe:prod', NS)
            if prod is not None:
                cprod = prod.find('nfe:cProd', NS)
                cean = prod.find('nfe:cEAN', NS)
                ceantrib = prod.find('nfe:cEANTrib', NS)
                xprod = prod.find('nfe:xProd', NS)
                
                # Só adiciona se tiver cProd (CODFAB)
                if cprod is not None and cprod.text:
                    codfab = cprod.text.strip()
                    ean = cean.text.strip() if cean is not None and cean.text else ''
                    ean_trib = ceantrib.text.strip() if ceantrib is not None and ceantrib.text else ''
                    descricao = xprod.text.strip() if xprod is not None and xprod.text else ''
                    
                    # Remove EANs vazios ou "SEM GTIN"
                    if ean.upper() in ('', 'SEM GTIN'):
                        ean = ''
                    if ean_trib.upper() in ('', 'SEM GTIN'):
                        ean_trib = ''
                    
                    produtos.append({
                        'codfab': codfab,
                        'cEAN': ean,
                        'cEANTrib': ean_trib,
                        'descricao_xml': descricao,
                        'xml_file': xml_filename
                    })
        
        return produtos
    
    except Exception as e:
        print(f"Erro ao ler XML {xml_path}: {e}")
        return []


def processar_xmls(pasta_xmls):
    """
    Lê todos os XMLs da pasta e retorna DataFrame consolidado.
    """
    pasta = Path(pasta_xmls)
    
    if not pasta.exists():
        print(f"ERRO: Pasta não encontrada: {pasta_xmls}")
        return pd.DataFrame()
    
    xml_files = list(pasta.glob('*.xml'))
    print(f"\nEncontrados {len(xml_files)} arquivos XML na pasta.")
    
    if len(xml_files) == 0:
        print("Nenhum arquivo XML encontrado!")
        return pd.DataFrame()
    
    todos_produtos = []
    
    for i, xml_file in enumerate(xml_files, 1):
        if i % 10 == 0:
            print(f"Processando XML {i}/{len(xml_files)}...")
        
        produtos = extrair_dados_xml(xml_file)
        todos_produtos.extend(produtos)
    
    print(f"\nTotal de produtos extraídos dos XMLs: {len(todos_produtos)}")
    
    # Converte para DataFrame
    df_xml = pd.DataFrame(todos_produtos)
    
    # Remove duplicatas (mesmo CODFAB pode aparecer em vários XMLs - mantém o primeiro)
    if not df_xml.empty:
        df_xml = df_xml.drop_duplicates(subset=['codfab'], keep='first')
        print(f"Produtos únicos (por CODFAB): {len(df_xml)}")
    
    return df_xml


def gerar_excel_preenchido(df_produtos_oracle, df_xml, output_path):
    """
    Faz o merge entre produtos do Oracle e dados dos XMLs,
    preenchendo os códigos auxiliares vazios.
    """
    print("\nFazendo match entre produtos do Oracle e XMLs (por CODFAB)...")
    
    # Garante que CODFAB seja string em ambos os DataFrames
    df_produtos_oracle['CODFAB'] = df_produtos_oracle['CODFAB'].astype(str).str.strip()
    df_xml['codfab'] = df_xml['codfab'].astype(str).str.strip()
    
    # Faz o merge
    df_resultado = df_produtos_oracle.merge(
        df_xml,
        left_on='CODFAB',
        right_on='codfab',
        how='left'
    )
    
    # Preenche os códigos auxiliares diretamente (substitui os vazios pelos valores do XML)
    df_resultado['CODAUXILIAR'] = df_resultado['cEAN']
    df_resultado['CODAUXILIAR2'] = df_resultado['cEAN']  # Mesmo valor que CODAUXILIAR
    df_resultado['CODAUXILIARTRIB'] = df_resultado['cEANTrib']
    
    # Reorganiza colunas para ficar mais legível
    colunas_finais = [
        'CODPROD',
        'CODFAB',
        'DESCRICAO',
        'CODEPTO',
        'DEPARTAMENTO',
        'CODFORNEC',
        'FORNECEDOR',
        'EMBALAGEM',
        'UNIDADE',
        'QT_VENDIDA_6M',
        'CODAUXILIAR',
        'CODAUXILIAR2',
        'CODAUXILIARTRIB',
        'xml_file',
        'descricao_xml'
    ]
    
    df_resultado = df_resultado[colunas_finais]
    
    # Produtos não encontrados nos XMLs
    df_nao_encontrados = df_resultado[df_resultado['xml_file'].isna()].copy()
    
    # Segunda aba: Departamentos dos produtos não encontrados
    if not df_nao_encontrados.empty:
        df_departamentos = df_nao_encontrados.groupby(['CODEPTO', 'DEPARTAMENTO']).agg({
            'CODPROD': 'count',
            'QT_VENDIDA_6M': 'sum'
        }).reset_index()
        
        df_departamentos.columns = ['CODEPTO', 'DEPARTAMENTO', 'QTD_PRODUTOS_VAZIOS', 'QT_VENDIDA_TOTAL_6M']
        df_departamentos = df_departamentos.sort_values('QTD_PRODUTOS_VAZIOS', ascending=False)
    else:
        df_departamentos = pd.DataFrame(columns=['CODEPTO', 'DEPARTAMENTO', 'QTD_PRODUTOS_VAZIOS', 'QT_VENDIDA_TOTAL_6M'])
    
    # Salva Excel com duas abas
    print(f"\nGerando Excel: {output_path}")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Produtos_Preenchidos')
        df_departamentos.to_excel(writer, index=False, sheet_name='Departamentos_Vazios')
    
    # Estatísticas
    total_produtos = len(df_resultado)
    produtos_encontrados = df_resultado['xml_file'].notna().sum()
    produtos_nao_encontrados = total_produtos - produtos_encontrados
    
    print(f"\n{'='*60}")
    print(f"RESULTADOS:")
    print(f"{'='*60}")
    print(f"Total de produtos com códigos zerados: {total_produtos}")
    print(f"Produtos encontrados nos XMLs: {produtos_encontrados}")
    print(f"Produtos NÃO encontrados nos XMLs: {produtos_nao_encontrados}")
    print(f"Departamentos com produtos vazios: {len(df_departamentos)}")
    print(f"{'='*60}")
    
    return df_resultado


def main():
    """Função principal."""
    print("="*80)
    print("PREENCHIMENTO DE CÓDIGOS AUXILIARES A PARTIR DE XMLs")
    print("="*80)
    
    # Pasta com os XMLs
    pasta_xmls = r"C:\Users\DESENVOLVIMENTO\Documents\ALL_XMLS"
    
    # Conecta no Oracle
    print("\n1. Conectando ao Oracle...")
    conn = get_db_connection()
    if not conn:
        print("ERRO: Não foi possível conectar ao banco de dados.")
        return
    
    try:
        # Busca produtos com códigos vazios
        print("\n2. Buscando produtos no Oracle...")
        df_oracle = buscar_produtos_vazios(conn)
        
        if df_oracle.empty:
            print("Nenhum produto encontrado com códigos auxiliares zerados.")
            return
        
        # Processa XMLs
        print("\n3. Processando XMLs...")
        df_xml = processar_xmls(pasta_xmls)
        
        if df_xml.empty:
            print("ERRO: Nenhum dado foi extraído dos XMLs.")
            return
        
        # Gera Excel
        print("\n4. Gerando Excel com dados preenchidos...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(__file__).parent / f"produtos_codauxiliares_preenchidos_{timestamp}.xlsx"
        
        df_resultado = gerar_excel_preenchido(df_oracle, df_xml, output_path)
        
        print(f"\n✅ Arquivo gerado com sucesso!")
        print(f"📁 Local: {output_path}")
        
    except Exception as e:
        print(f"\n❌ Erro durante o processamento: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        conn.close()
        print("\n🔌 Conexão com o banco fechada.")


if __name__ == "__main__":
    main()

