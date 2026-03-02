# 01 - Análises de Produtos e Vendas

Ferramentas para diagnosticar a performance comercial por produto, cliente, data e departamento.  
Respondem perguntas como: *Quais produtos vendem mais? Qual a margem real após devoluções? Quais clientes compraram determinado item?*

## Scripts deste grupo

| Script | O que faz | Saída |
|---|---|---|
| `analise_clientes_produtos_875_879.py` | Lista clientes que compraram produtos específicos a um preço definido, por mês/ano. | Excel formatado |
| `analise_faturamento_depto_109_127_129.py` | Calcula faturamento de departamentos específicos em um recorte de datas. | Console / Excel |
| `analise_produto_2608.py` | Análise completa de um produto: notas fiscais, preço praticado e volume. | Console / Excel |
| `analise_produtos_1615_543_3meses.py` | Compara vendas de dois produtos em janela de 3 meses. | Console / Excel |
| `analise_produtos_abaixo_preço.py` | Identifica vendas abaixo do custo (margem negativa) e gera relatório detalhado. | Excel formatado |
| `analise_produtos_completa.py` | Análise completa de linhá de produtos: compras, estoque, vendas e indicadores de decisão. | Excel |
| `analise_produtos_gui.py` | Interface desktop para filtrar produtos por período/departamento e exportar. | Excel + PDF |
| `analise_produtos_novembro_corrigido.py` | Análise mensal com ajuste de lógica de inclusão/exclusão de pedidos especiais. | Excel |
| `analise_vendas_dia_excel.py` | Gera resumo de vendas de um dia específico formatado para envio gerencial. | Excel |
| `analise_vendas_produtos_universal.py` | Filtro de vendas por códigos de produto e faixa de preço. | Excel |
| `analise_vendas_universal_gui.py` | Interface desktop para filtrar e exportar qualquer corte de vendas. | Excel |
| `curva_abc_clientes_pdf.py` | Classifica clientes em A/B/C por faturamento para priorização da equipe comercial. | PDF |
| `produtos_nf.py` | Análise de vendas por produto focada em dados de nota fiscal. | Console / Excel |
| `relatório_vendas_depto_mes.py` | Relatório de vendas por departamento consolidado em múltiplos meses. | Excel |
| `vendas_cliente_pdf.py` | Histórico de compras de um cliente específico com indicadores de período. | PDF |

## Tecnologias utilizadas

- `oracledb`, `pandas`, `openpyxl`
- `tkinter`, `reportlab`, `weasyprint`

## Valor para o negócio

- Visibilidade de margem e rentabilidade por produto em minutos, sem depender do time de TI
- Identificação rápida de vendas abaixo do custo para correção de preço ou bloqueio
- Curva ABC automatizada para o time comercial focar nos clientes de maior retorno

