# 06 - Monitoramento e Utilitários

Ferramentas de suporte operacional para alertas em tempo real e validações rápidas.  
Respondem perguntas como: *Como saber quando um pedido fica parado no checkout sem ficar abrindo o sistema?*

## Scripts deste grupo

| Script | O que faz | Saída |
|---|---|---|
| `monitor_pedidos_checkout.py` | Monitora continuamente pedidos parados em checkout e dispara notificação desktop. | Notificação Windows |
| `testar_notificação.py` | Valida o canal de notificação local antes de colocar o monitor em produção. | Notificação Windows |

## Tecnologias utilizadas

- `oracledb`, `threading`
- `plyer`, `win10toast`

## Valor para o negócio

- Resposta mais rápida a pedidos parados, reduzindo o tempo de espera do cliente
- Validacao do canal de alerta antes do deploy em ambiente real

