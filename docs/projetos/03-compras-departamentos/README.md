# 03 - Compras e Departamentos

Ferramentas para comparar compras versus vendas, identificar pontos de risco de margem e avaliar eficiencia por departamento.  
Respondem perguntas como: *Estamos comprando mais do que vendendo? Qual departamento esta com margem negativa? Qual o impacto de excluir uma NF nos totais?*

## Scripts deste grupo

| Script | O que faz | Saída |
|---|---|---|
| `analise_depto_período.py` | Relatório completo de um departamento em período selecionável: vendas, compras, margem. | Excel / PDF |
| `compras_maior_venda_depto.py` | Identifica departamentos onde o volume comprado superou o volume vendido no período. | Console / Excel |
| `compraxvenda_fornecedor.py` | Comparativo mensal de compra x venda por departamento e fornecedor. | Excel |
| `compraxvenda_totais.py` | Visão consolidada de compra x venda por departamento. | Console / Excel |
| `excluir_nf_totais.py` | Simula exclusão de notas fiscais e mede impacto nos totais e na margem. | Console |
| `relatório_categorias_100_101.py` | Relatório de produtos das categorias 100 e 101 por departamento. | Excel |

## Tecnologias utilizadas

- `oracledb`, `pandas`
- `tkinter`, `reportlab`, `weasyprint`

## Valor para o negócio

- Apoio a decisoes de compras com base no sell-out real, evitando sobrestoque
- Identificação proativa de departamentos com risco de margem negativa
- Controle gerencial por departamento e categoria para reuniões de resultado

