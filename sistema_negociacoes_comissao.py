"""
Sistema de Negociações e Comissões
Lê dados da planilha NEGOCIAÇÕES e gera relatório de comissões
"""

import oracledb
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import sys
import os

# Adiciona o diretório raiz ao path para importar google_sheets_api
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from google_sheets_api import _get_worksheet

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

SPREADSHEET_NAME = 'NEGOCIAÇÕES'
ABA_NEGOCIACOES = 'NEGOCIAÇÕES'
ABA_RELATORIO = 'RELATÓRIO NEGOCIAÇÕES'

# =========================================================================
# FUNÇÕES DE BANCO DE DADOS
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

def buscar_dados_pedido(conn, numped, codprod_list=None):
    """
    Busca dados do pedido no banco Oracle
    
    Args:
        conn: Conexão com o banco
        numped: Número do pedido
        codprod_list: Lista de códigos de produtos específicos 
                     (None ou vazio = busca TODOS os produtos do pedido)
    
    Returns:
        Lista de dicionários com os dados do pedido
    """
    # Monta filtro de produtos se fornecido
    # Se codprod_list for None ou vazio, não aplica filtro = busca TODOS os produtos
    filtro_produto = ""
    if codprod_list and len(codprod_list) > 0:
        codprod_str = ', '.join(map(str, codprod_list))
        filtro_produto = f"AND I.CODPROD IN ({codprod_str})"
        print(f"      🔍 Buscando apenas produtos específicos: {codprod_list}")
    else:
        print(f"      🔍 Buscando TODOS os produtos do pedido")
    
    query = f"""
    SELECT 
        C.DATA,
        C.CODCLI,
        CL.CLIENTE,
        C.NUMPED,
        C.CODUSUR,
        U.NOME AS VENDEDOR,
        I.CODPROD,
        ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2) AS VALOR_TOTAL
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.NUMPED = {numped}
        AND C.CODFILIAL IN ('1', '98')
        AND C.POSICAO = 'F'
        AND C.DTCANCEL IS NULL
        AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        AND NVL(I.BONIFIC, 'N') = 'N'
        {filtro_produto}
    ORDER BY I.CODPROD
    """
    
    try:
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]
        results = []
        for row in cursor.fetchall():
            results.append(dict(zip(columns, row)))
        cursor.close()
        return results
    except Exception as e:
        print(f"❌ Erro ao buscar dados do pedido {numped}: {e}")
        return []

# =========================================================================
# FUNÇÕES DO GOOGLE SHEETS
# =========================================================================
def ler_negociacoes():
    """
    Lê os dados da aba NEGOCIAÇÕES
    Retorna lista de dicionários com os dados
    """
    print(f"📖 Lendo aba '{ABA_NEGOCIACOES}' da planilha '{SPREADSHEET_NAME}'...")
    worksheet = _get_worksheet(SPREADSHEET_NAME, ABA_NEGOCIACOES)
    
    if not worksheet:
        print(f"❌ Não foi possível acessar a aba '{ABA_NEGOCIACOES}'")
        return []
    
    data = worksheet.get_all_values()
    
    if not data or len(data) < 2:  # Precisa ter pelo menos cabeçalho + 1 linha
        print("⚠️  A planilha está vazia ou não contém dados.")
        return []
    
    negociacoes = []
    # Pula o cabeçalho (linha 1) e processa a partir da linha 2
    print(f"   Total de linhas na planilha: {len(data)}")
    for idx, row in enumerate(data[1:], start=2):
        try:
            # Coluna A: NUMPED
            numped_str = str(row[0]).strip() if len(row) > 0 and row[0] else ''
            if not numped_str or not numped_str.isdigit():
                if numped_str:  # Só mostra log se tiver algo (não vazio)
                    print(f"   ⚠️  Linha {idx}: NUMPED inválido ou vazio: '{numped_str}'")
                continue  # Pula linhas sem NUMPED válido
            
            numped = int(numped_str)
            
            # Coluna B: Porcentagem de comissão
            comissao_str = str(row[1]).strip() if len(row) > 1 and row[1] else '0'
            try:
                comissao_pct = float(comissao_str)
            except ValueError:
                comissao_pct = 0.0
                print(f"   ⚠️  Linha {idx}: Comissão inválida, usando 0%")
            
            # Colunas C, D, E: ITEM 1, ITEM 2, ITEM 3 (codprod específicos)
            # Se não tiver nada nessas colunas, considera TODOS os produtos do pedido
            codprod_list = []
            for col_idx in [2, 3, 4]:  # Colunas C, D, E (índices 2, 3, 4)
                if len(row) > col_idx and row[col_idx]:
                    codprod_str = str(row[col_idx]).strip()
                    if codprod_str and codprod_str.isdigit():
                        codprod_list.append(int(codprod_str))
            
            # Se codprod_list estiver vazio, seta como None para indicar "todos os produtos"
            neg_info = {
                'linha': idx,
                'numped': numped,
                'comissao_pct': comissao_pct,
                'codprod_list': codprod_list if len(codprod_list) > 0 else None
            }
            negociacoes.append(neg_info)
            
            # Log informativo
            if codprod_list:
                produtos_info = f"Produtos específicos: {codprod_list}"
            else:
                produtos_info = "TODOS os produtos do pedido"
            print(f"   ✅ Linha {idx}: Pedido {numped} - Comissão: {comissao_pct}% - {produtos_info}")
            
        except Exception as e:
            print(f"⚠️  Erro ao processar linha {idx}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    print(f"✅ {len(negociacoes)} negociação(ões) encontrada(s).")
    return negociacoes

def ler_relatorio_existente():
    """
    Lê os dados existentes da aba RELATÓRIO NEGOCIAÇÕES
    Retorna um set com os NUMPEDs já processados
    """
    print(f"📖 Verificando dados existentes na aba '{ABA_RELATORIO}'...")
    worksheet = _get_worksheet(SPREADSHEET_NAME, ABA_RELATORIO)
    
    if not worksheet:
        print(f"⚠️  Aba '{ABA_RELATORIO}' não encontrada ou vazia. Será criada.")
        return set()
    
    data = worksheet.get_all_values()
    
    if not data or len(data) < 2:  # Precisa ter pelo menos cabeçalho + 1 linha
        return set()
    
    numpeds_processados = set()
    # Coluna C (índice 2) = NUMERO DO PEDIDO
    for row in data[1:]:  # Pula cabeçalho
        try:
            if len(row) > 2 and row[2]:
                numped_str = str(row[2]).strip()
                if numped_str and numped_str.isdigit():
                    numpeds_processados.add(int(numped_str))
        except:
            continue
    
    print(f"✅ {len(numpeds_processados)} pedido(s) já processado(s) encontrado(s).")
    return numpeds_processados

def preparar_cabecalho_relatorio(worksheet):
    """
    Prepara o cabeçalho da aba RELATÓRIO NEGOCIAÇÕES se não existir
    """
    data = worksheet.get_all_values()
    
    # Se não tem cabeçalho, adiciona
    if not data or len(data) == 0:
        cabecalho = [
            'DATA DO PEDIDO',
            'CODCLI',
            'NUMERO DO PEDIDO',
            'COMISSAO %',
            'CODUSUR',
            'NOME DO VENDEDOR',
            'VALOR',
            'VALOR DE COMISSAO'
        ]
        worksheet.append_row(cabecalho)
        print("✅ Cabeçalho criado na aba RELATÓRIO NEGOCIAÇÕES.")

def escrever_relatorio(negociacoes_processadas):
    """
    Escreve os dados processados na aba RELATÓRIO NEGOCIAÇÕES
    
    Args:
        negociacoes_processadas: Lista de dicionários com os dados processados
    """
    print(f"📝 Escrevendo dados na aba '{ABA_RELATORIO}'...")
    worksheet = _get_worksheet(SPREADSHEET_NAME, ABA_RELATORIO)
    
    if not worksheet:
        print(f"❌ Não foi possível acessar a aba '{ABA_RELATORIO}'")
        return False
    
    # Prepara cabeçalho se necessário
    preparar_cabecalho_relatorio(worksheet)
    
    # Prepara dados para escrita
    rows_to_append = []
    for neg in negociacoes_processadas:
        # Formata data
        data_str = ''
        if neg.get('data'):
            if isinstance(neg['data'], datetime):
                data_str = neg['data'].strftime('%d/%m/%Y')
            else:
                data_str = str(neg['data'])
        
        # Formata valor (sem R$, tipo 17587,2 com vírgula)
        valor = neg.get('valor', 0)
        valor_str = f"{valor:.2f}".replace('.', ',')
        
        # Formata valor de comissão (sem R$, tipo 175,87 com vírgula)
        valor_comissao = neg.get('valor_comissao', 0)
        valor_comissao_str = f"{valor_comissao:.2f}".replace('.', ',')
        
        # Formata porcentagem de comissão (sem vírgula se for inteiro)
        comissao_pct = neg.get('comissao_pct', 0)
        if comissao_pct == int(comissao_pct):
            comissao_pct_str = str(int(comissao_pct))
        else:
            comissao_pct_str = f"{comissao_pct:.2f}".replace('.', ',')
        
        row = [
            data_str,  # A: DATA DO PEDIDO
            str(neg.get('codcli', '')),  # B: CODCLI (código do cliente)
            str(neg.get('numped', '')),  # C: NUMERO DO PEDIDO
            comissao_pct_str,  # D: COMISSAO %
            str(neg.get('codusur', '')),  # E: CODUSUR
            neg.get('vendedor', ''),  # F: NOME DO VENDEDOR
            valor_str,  # G: VALOR
            valor_comissao_str  # H: VALOR DE COMISSAO
        ]
        rows_to_append.append(row)
        print(f"   📋 Preparado: Pedido {neg.get('numped')} - Cliente: {neg.get('codcli')} - R$ {valor:,.2f}")
    
    # Escreve todas as linhas de uma vez
    if rows_to_append:
        try:
            print(f"\n💾 Tentando escrever {len(rows_to_append)} linha(s)...")
            worksheet.append_rows(rows_to_append)
            print(f"✅ {len(rows_to_append)} linha(s) adicionada(s) ao relatório com sucesso!")
            
            # Verifica se foi escrito (lê a última linha)
            data_verificacao = worksheet.get_all_values()
            if len(data_verificacao) > len(rows_to_append):
                print(f"✅ Verificação: {len(data_verificacao)} linha(s) encontrada(s) na planilha.")
            else:
                print(f"⚠️  Verificação: Apenas {len(data_verificacao)} linha(s) encontrada(s). Pode haver problema.")
            
            return True
        except Exception as e:
            print(f"❌ Erro ao escrever no Google Sheets: {e}")
            print("   Tentando escrever linha por linha...")
            
            # Tenta escrever linha por linha como fallback
            sucesso = 0
            for idx, row in enumerate(rows_to_append, start=1):
                try:
                    worksheet.append_row(row)
                    sucesso += 1
                    print(f"   ✅ Linha {idx} escrita com sucesso.")
                except Exception as e2:
                    print(f"   ❌ Erro ao escrever linha {idx}: {e2}")
            
            if sucesso > 0:
                print(f"✅ {sucesso} de {len(rows_to_append)} linha(s) escrita(s) com sucesso.")
                return True
            else:
                print(f"❌ Nenhuma linha foi escrita.")
                return False
    else:
        print("⚠️  Nenhum dado para adicionar.")
        return False

# =========================================================================
# PROCESSAMENTO
# =========================================================================
def processar_negociacoes():
    """
    Processa as negociações e gera o relatório
    """
    print("=" * 100)
    print("📊 SISTEMA DE NEGOCIAÇÕES E COMISSÕES")
    print("=" * 100)
    print()
    
    # Lê negociações da planilha
    negociacoes = ler_negociacoes()
    
    if not negociacoes:
        print("❌ Nenhuma negociação encontrada para processar.")
        return
    
    # Lê pedidos já processados
    numpeds_processados = ler_relatorio_existente()
    
    # Conecta ao banco
    print("\n🔌 Conectando ao banco Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        return
    
    try:
        negociacoes_processadas = []
        negociacoes_nao_encontradas = []
        
        print(f"\n🔍 Processando {len(negociacoes)} negociação(ões)...")
        print()
        
        for neg in negociacoes:
            numped = neg['numped']
            comissao_pct = neg['comissao_pct']
            codprod_list = neg['codprod_list']
            
            # Verifica se já foi processado
            if numped in numpeds_processados:
                print(f"⏭️  Pedido {numped} já processado. Pulando...")
                continue
            
            print(f"📦 Processando pedido {numped} (Comissão: {comissao_pct}%)...")
            
            # Busca dados do pedido
            # Se codprod_list for None, busca TODOS os produtos do pedido
            dados_pedido = buscar_dados_pedido(conn, numped, codprod_list)
            
            if not dados_pedido:
                print("❌ Pedido não encontrado ou sem itens válidos.")
                negociacoes_nao_encontradas.append(numped)
                continue
            
            # Agrupa dados por pedido (pode ter múltiplos itens)
            # Pega dados do primeiro item (são os mesmos para todos os itens do pedido)
            primeiro_item = dados_pedido[0]
            
            # Calcula valor total (soma de todos os itens)
            valor_total = sum(item['VALOR_TOTAL'] for item in dados_pedido)
            
            # Calcula valor da comissão
            valor_comissao = (valor_total * comissao_pct) / 100
            
            # Prepara dados para o relatório
            neg_processada = {
                'data': primeiro_item['DATA'],
                'codcli': primeiro_item['CODCLI'],  # Usa CODCLI ao invés do nome
                'numped': numped,
                'comissao_pct': comissao_pct,
                'codusur': primeiro_item['CODUSUR'],
                'vendedor': primeiro_item['VENDEDOR'] or '',
                'valor': round(valor_total, 2),
                'valor_comissao': round(valor_comissao, 2)
            }
            
            negociacoes_processadas.append(neg_processada)
            print(f"✅ Valor: R$ {valor_total:,.2f} | Comissão: R$ {valor_comissao:,.2f}")
        
        # Ordena por data (mais antiga para mais nova)
        if negociacoes_processadas:
            print(f"\n📅 Ordenando por data (mais antiga para mais nova)...")
            negociacoes_processadas.sort(key=lambda x: x.get('data') if x.get('data') else datetime.min)
            print(f"✅ Ordenação concluída.")
        
        # Escreve no relatório
        print(f"\n📊 Resumo do processamento:")
        print(f"   - Negociações processadas: {len(negociacoes_processadas)}")
        print(f"   - Pedidos não encontrados: {len(negociacoes_nao_encontradas)}")
        
        if negociacoes_processadas:
            print(f"\n📝 Escrevendo {len(negociacoes_processadas)} registro(s) no relatório...")
            print(f"   Dados a serem escritos (ordenados por data):")
            for idx, neg in enumerate(negociacoes_processadas, start=1):
                data_str = neg.get('data').strftime('%d/%m/%Y') if isinstance(neg.get('data'), datetime) else str(neg.get('data', ''))
                print(f"   {idx}. {data_str} - Pedido {neg.get('numped')} - Cliente: {neg.get('codcli')} - Valor: R$ {neg.get('valor', 0):,.2f}")
            
            resultado = escrever_relatorio(negociacoes_processadas)
            if resultado:
                print("\n✅ Processamento concluído com sucesso!")
            else:
                print("\n⚠️  Processamento concluído, mas houve problemas ao escrever no relatório.")
        else:
            print("\n⚠️  Nenhuma negociação nova para processar.")
        
        if negociacoes_nao_encontradas:
            print(f"\n⚠️  {len(negociacoes_nao_encontradas)} pedido(s) não encontrado(s): {negociacoes_nao_encontradas}")
        
    except Exception as e:
        print(f"\n❌ Erro ao processar negociações: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# =========================================================================
# TESTE DE ESCRITA
# =========================================================================
def testar_escrita():
    """Função de teste para verificar se a escrita está funcionando"""
    print("=" * 100)
    print("🧪 TESTE DE ESCRITA NO GOOGLE SHEETS")
    print("=" * 100)
    print()
    
    worksheet = _get_worksheet(SPREADSHEET_NAME, ABA_RELATORIO)
    
    if not worksheet:
        print(f"❌ Não foi possível acessar a aba '{ABA_RELATORIO}'")
        return
    
    # Prepara cabeçalho se necessário
    preparar_cabecalho_relatorio(worksheet)
    
    # Testa escrita de uma linha
    linha_teste = [
        '01/12/2025',  # DATA
        'CLIENTE TESTE',  # CLIENTE
        '999999',  # NUMPED
        '1',  # COMISSAO %
        '1',  # CODUSUR
        'VENDEDOR TESTE',  # VENDEDOR
        '1000,00',  # VALOR
        '10,00'  # VALOR COMISSAO
    ]
    
    try:
        print("📝 Tentando escrever linha de teste...")
        worksheet.append_row(linha_teste)
        print("✅ Linha de teste escrita com sucesso!")
        
        # Verifica se foi escrito
        data = worksheet.get_all_values()
        print(f"✅ Total de linhas na planilha: {len(data)}")
        if len(data) > 0:
            print(f"   Última linha: {data[-1]}")
        
    except Exception as e:
        print(f"❌ Erro ao escrever linha de teste: {e}")
        import traceback
        traceback.print_exc()

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    import sys
    
    # Se passar --teste como argumento, executa teste de escrita
    if len(sys.argv) > 1 and sys.argv[1] == '--teste':
        testar_escrita()
    else:
        try:
            processar_negociacoes()
        except KeyboardInterrupt:
            print("\n\n⚠️  Processamento interrompido pelo usuário.")
        except Exception as e:
            print(f"\n❌ Erro fatal: {e}")
            import traceback
            traceback.print_exc()

