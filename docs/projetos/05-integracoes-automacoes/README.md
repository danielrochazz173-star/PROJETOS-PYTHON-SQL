# 05 - Integrações e Automações

Ferramentas de produtividade operacional que conectam Oracle ERP, arquivos XML e planilhas, eliminando tarefas manuais repetitivas.  
Respondem perguntas como: *Como atualizar cadastros no ERP sem digitar um por um? Como enviar dados de compras para uma planilha automaticamente?*

## Scripts deste grupo

| Script | O que faz | Saída |
|---|---|---|
| `atualizar_codauxiliares_oracle.py` | Interface desktop para atualizar códigos auxiliares de produtos no Oracle. | Oracle DB atualizado |
| `atualizar_codauxiliares_oracle_v2.py` | Versão evoluida com validações adicionais e log de alterações. | Oracle DB atualizado |
| `preencher_codauxiliares_xml.py` | Preenche códigos auxiliares em XMLs de NF com base em consulta ao Oracle. | XMLs atualizados |
| `enviar_compras_planilha.py` | Exporta dados de compras do Oracle diretamente para Excel ou Google Sheets. | Excel / Google Sheets |
| `exportar_query_excel.py` | Utilitário genérico: executa qualquer query SQL Oracle e exporta para Excel. | Excel |
| `falta_produtos_interface.py` | Interface para registrar e acompanhar faltas de produtos e gerar relatório. | Excel / relatório operacional |
| `pesquisar_codfab.py` | Consulta rápida de produto pelo código de fábrica (CODFAB) no Oracle. | Console |

## Tecnologias utilizadas

- `oracledb`, `pandas`
- `gspread`, `xml.etree.ElementTree`

## Valor para o negócio

- Redução de erros e tempo em tarefas de cadastro que antes eram manuais
- Integração direta entre ERP e planilhas sem copiar/colar dados
- Aumento de velocidade operacional no backoffice comercial e de compras

