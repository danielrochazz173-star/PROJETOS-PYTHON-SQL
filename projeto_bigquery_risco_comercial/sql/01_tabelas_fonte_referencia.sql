-- Projeto ficticio de referencia para portfolio
-- Dataset sugerido: raw_erp

CREATE TABLE IF NOT EXISTS `portfolio-retail-ops.raw_erp.pedidos_cab` (
  id_pedido STRING,
  cod_cliente INT64,
  canal_venda STRING,
  status_pedido STRING,
  data_faturamento TIMESTAMP
)
PARTITION BY DATE(data_faturamento);

CREATE TABLE IF NOT EXISTS `portfolio-retail-ops.raw_erp.pedidos_itens` (
  id_pedido STRING,
  cod_produto INT64,
  cod_departamento INT64,
  qtd NUMERIC,
  valor_total NUMERIC
)
CLUSTER BY cod_departamento, cod_produto;

CREATE TABLE IF NOT EXISTS `portfolio-retail-ops.raw_erp.devolucoes_itens` (
  id_devolucao STRING,
  cod_cliente INT64,
  cod_produto INT64,
  data_devolucao TIMESTAMP,
  valor_devolvido NUMERIC
)
PARTITION BY DATE(data_devolucao);

CREATE TABLE IF NOT EXISTS `portfolio-retail-ops.raw_erp.custos_produtos` (
  cod_produto INT64,
  data_custo TIMESTAMP,
  custo_medio NUMERIC
)
PARTITION BY DATE(data_custo)
CLUSTER BY cod_produto;
