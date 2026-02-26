"""
Análise de vendas por dia (geração direta em Excel)
- Pergunta códigos de produtos separados por vírgula
- Pergunta uma data específica no formato DD/MM/AAAA
- Gera Excel com todas as vendas do dia para esses produtos
  - Aba Detalhado: NF, data, produto, preço, quantidade, valor total, cliente, vendedor
  - Aba Resumo por NF: valor total por nota/cliente
- Abre o arquivo do Excel automaticamente ao final
"""

import oracledb
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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


# =========================================================================
# FUNÇÕES
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None


def montar_query_vendas(data_str: str, codprod_list):
    """Monta a query SQL para buscar vendas dos produtos na data informada."""
    codprod_str = ', '.join(map(str, codprod_list))

    query = f"""
    SELECT 
        C.NUMNOTA AS NF,
        C.DATA,
        I.CODPROD,
        P.DESCRICAO AS PRODUTO,
        I.PVENDA AS PRECO,
        I.QT AS QUANTIDADE,
        ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL,
        CL.CODCLI,
        CL.CLIENTE,
        C.CODUSUR,
        U.NOME AS VENDEDOR
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA = TO_DATE('{data_str}', 'DD/MM/YYYY')
      AND I.CODPROD IN ({codprod_str})
      AND C.CODFILIAL IN ('1', '98')
      AND C.POSICAO = 'F'
      AND C.DTCANCEL IS NULL
      AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
      AND NVL(I.BONIFIC, 'N') = 'N'
    ORDER BY C.DATA, C.NUMNOTA, I.CODPROD
    """

    return query


def formatar_excel(arquivo_excel: Path, df: pd.DataFrame, resumo_nf: pd.DataFrame):
    """Cria e formata o arquivo Excel com abas Detalhado e Resumo por NF."""
    from openpyxl import load_workbook

    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        # Aba 1: Detalhado
        df.to_excel(writer, sheet_name='Detalhado', index=False)
        ws_detalhado = writer.sheets['Detalhado']

        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )

        # Formata cabeçalho
        for cell in ws_detalhado[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        # Ajusta largura das colunas
        col_widths = {
            'A': 12,  # NF
            'B': 12,  # DATA
            'C': 10,  # CODPROD
            'D': 40,  # PRODUTO
            'E': 12,  # PRECO
            'F': 10,  # QUANTIDADE
            'G': 15,  # VALOR_TOTAL
            'H': 10,  # CODCLI
            'I': 40,  # CLIENTE
            'J': 10,  # CODUSUR
            'K': 25,  # VENDEDOR
        }
        for col, width in col_widths.items():
            ws_detalhado.column_dimensions[col].width = width

        # Formata dados
        for row in ws_detalhado.iter_rows(min_row=2, max_row=ws_detalhado.max_row):
            for cell in row:
                cell.border = border
                if cell.column == 5:  # PRECO
                    cell.number_format = 'R$ #,##0.00'
                elif cell.column == 7:  # VALOR_TOTAL
                    cell.number_format = 'R$ #,##0.00'
                elif cell.column == 2:  # DATA
                    cell.number_format = 'DD/MM/YYYY'

        # Aba 2: Resumo por NF
        resumo_nf_excel = resumo_nf.rename(columns={'QUANTIDADE': 'QTD_TOTAL'})
        resumo_nf_excel.to_excel(writer, sheet_name='Resumo por NF', index=False)
        ws_resumo = writer.sheets['Resumo por NF']

        # Formata cabeçalho do resumo
        for cell in ws_resumo[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        # Ajusta largura das colunas do resumo
        ws_resumo.column_dimensions['A'].width = 12  # NF
        ws_resumo.column_dimensions['B'].width = 12  # DATA
        ws_resumo.column_dimensions['C'].width = 12  # QTD_TOTAL
        ws_resumo.column_dimensions['D'].width = 15  # VALOR_TOTAL
        ws_resumo.column_dimensions['E'].width = 40  # CLIENTE

        # Formata dados do resumo
        for row in ws_resumo.iter_rows(min_row=2, max_row=ws_resumo.max_row):
            for cell in row:
                cell.border = border
                if cell.column == 4:  # VALOR_TOTAL
                    cell.number_format = 'R$ #,##0.00'
                elif cell.column == 2:  # DATA
                    cell.number_format = 'DD/MM/YYYY'

        # Congela primeira linha
        ws_detalhado.freeze_panes = 'A2'
        ws_resumo.freeze_panes = 'A2'


# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    print("=" * 100)
    print("📊 ANÁLISE DE VENDAS POR DIA - EXCEL")
    print("=" * 100)
    print()

    # Lê códigos de produtos
    try:
        codigos_input = input(
            "Digite os códigos dos produtos separados por vírgula (ex: 2608 ou 2608,875,879): "
        ).strip()

        if not codigos_input:
            print("❌ Nenhum código informado.")
            raise SystemExit(1)

        codprod_list = []
        for cod in codigos_input.split(','):
            cod_limpo = cod.strip()
            if cod_limpo.isdigit():
                codprod_list.append(int(cod_limpo))

        if not codprod_list:
            print("❌ Nenhum código válido encontrado.")
            raise SystemExit(1)

        # Lê data específica
        data_input = input("Digite a DATA específica (formato DD/MM/AAAA): ").strip()
        try:
            data_obj = datetime.strptime(data_input, "%d/%m/%Y")
        except ValueError:
            print("❌ Data inválida. Use o formato DD/MM/AAAA.")
            raise SystemExit(1)

        data_str = data_obj.strftime("%d/%m/%Y")

        print()
        print(f"🔍 Buscando vendas na data: {data_str}")
        print(f"   Produtos: {codprod_list}")
        print()

    except KeyboardInterrupt:
        print("\nOperação cancelada pelo usuário.")
        raise SystemExit(1)

    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()

    if not conn:
        print("❌ Falha na conexão. Abortando.")
        raise SystemExit(1)

    try:
        print("📈 Executando query...")
        query = montar_query_vendas(data_str, codprod_list)
        df = pd.read_sql_query(query, conn)

        if df.empty:
            print("⚠️  Nenhuma venda encontrada com os critérios especificados.")
            raise SystemExit(0)

        print(f"✅ {len(df)} registro(s) encontrado(s).\n")

        # Converte colunas de data se necessário
        if 'DATA' in df.columns:
            df['DATA'] = pd.to_datetime(df['DATA'])

        # Resumo por NF
        resumo_nf = df.groupby(['NF', 'DATA', 'CLIENTE'], as_index=False).agg(
            QUANTIDADE=('QUANTIDADE', 'sum'),
            VALOR_TOTAL=('VALOR_TOTAL', 'sum'),
        )
        resumo_nf = resumo_nf.sort_values('VALOR_TOTAL', ascending=False)

        # Nome do arquivo Excel
        cods_nome = "_".join(map(str, codprod_list))
        data_nome = data_obj.strftime("%Y%m%d")
        nome_arquivo = f"vendas_dia_{data_nome}_produtos_{cods_nome}.xlsx"
        caminho_arquivo = Path(nome_arquivo).resolve()

        print("💾 Gerando arquivo Excel...")
        formatar_excel(caminho_arquivo, df, resumo_nf)

        print(f"✅ Excel gerado: {caminho_arquivo}")

        # Abre o arquivo automaticamente
        try:
            os.startfile(str(caminho_arquivo))
            print("📂 Abrindo o arquivo no Excel...")
        except Exception as e:
            print(f"⚠️ Não foi possível abrir o Excel automaticamente: {e}")

    except Exception as e:
        print(f"❌ Erro ao executar análise: {e}")
        import traceback

        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")





# portfolio-commit-ready: analise_vendas_dia_excel
