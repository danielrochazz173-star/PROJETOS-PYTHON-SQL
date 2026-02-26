"""
Sistema de Monitoramento de Pedidos - Checkout
Monitora a tabela PCPEDC e notifica quando DTINICIALCHECKOUT e DTFINALCHECKOUT estão preenchidas
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import oracledb
import pandas as pd
from datetime import datetime
import threading
import time
import json
from pathlib import Path

# Tentar importar bibliotecas de notificação
NOTIFICACAO_PLYER = False
ToastNotifier = None
notification = None

try:
    from win10toast import ToastNotifier
    NOTIFICACAO_DISPONIVEL = True
except ImportError:
    try:
        from plyer import notification
        NOTIFICACAO_PLYER = True
        NOTIFICACAO_DISPONIVEL = True
    except ImportError:
        NOTIFICACAO_DISPONIVEL = False

# Configuração do Oracle Instant Client
try:
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
except Exception:
    pass

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# Arquivo para salvar pedidos já notificados
ARQUIVO_NOTIFICADOS = 'pedidos_notificados.json'

# =========================================================================
# FUNÇÕES DE BANCO
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

def buscar_pedidos_checkout_completo(conn, apenas_hoje=True):
    """
    Busca pedidos onde DTINICIALCHECKOUT e DTFINALCHECKOUT estão preenchidas.
    Se apenas_hoje=True, filtra pelo dia atual (considera qualquer uma das datas de checkout).
    """
    # Filtro de data do dia atual (considera inicial OU final)
    filtro_data = ""
    if apenas_hoje:
        filtro_data = """
          AND (
                TRUNC(C.DTINICIALCHECKOUT) = TRUNC(SYSDATE)
             OR TRUNC(C.DTFINALCHECKOUT)   = TRUNC(SYSDATE)
          )
        """
    
    query = f"""
    SELECT 
        C.NUMPED,
        C.NUMNOTA,
        TO_CHAR(C.DATA, 'DD/MM/YYYY') AS DATA_PEDIDO,
        C.CODCLI,
        CL.CLIENTE,
        C.CODFILIAL,
        TO_CHAR(C.DTINICIALCHECKOUT, 'DD/MM/YYYY HH24:MI:SS') AS DTINICIALCHECKOUT,
        TO_CHAR(C.DTFINALCHECKOUT, 'DD/MM/YYYY HH24:MI:SS') AS DTFINALCHECKOUT,
        I.CODPROD,
        P.DESCRICAO AS PRODUTO,
        I.QT AS QUANTIDADE,
        I.PVENDA AS PRECO_UNIT,
        ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL_ITEM,
        C.CODUSUR,
        U.NOME AS VENDEDOR
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DTINICIALCHECKOUT IS NOT NULL
      AND C.DTFINALCHECKOUT IS NOT NULL
      AND C.DTCANCEL IS NULL
      {filtro_data}
    ORDER BY C.DTFINALCHECKOUT DESC, C.NUMPED, I.CODPROD
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        return df
    except Exception as e:
        print(f"❌ Erro ao buscar pedidos: {e}")
        return pd.DataFrame()

def carregar_pedidos_notificados():
    """Carrega lista de pedidos já notificados"""
    if Path(ARQUIVO_NOTIFICADOS).exists():
        try:
            with open(ARQUIVO_NOTIFICADOS, 'r', encoding='utf-8') as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def salvar_pedido_notificado(numped):
    """Salva pedido na lista de notificados"""
    notificados = carregar_pedidos_notificados()
    notificados.add(int(numped))
    try:
        with open(ARQUIVO_NOTIFICADOS, 'w', encoding='utf-8') as f:
            json.dump(list(notificados), f, indent=2)
    except Exception:
        pass

def formatar_detalhes_pedido(df_pedido):
    """Formata detalhes do pedido para exibição"""
    if df_pedido.empty:
        return "Nenhum detalhe disponível"
    
    # Agrupa por pedido
    pedidos_info = []
    for numped in df_pedido['NUMPED'].unique():
        pedido = df_pedido[df_pedido['NUMPED'] == numped].iloc[0]
        itens = df_pedido[df_pedido['NUMPED'] == numped]
        
        valor_total = itens['VALOR_TOTAL_ITEM'].sum()
        qtd_total = itens['QUANTIDADE'].sum()
        
        info = f"""
📦 PEDIDO: {numped}
👤 Cliente: {pedido['CLIENTE']} (Cód: {pedido['CODCLI']})
📅 Data: {pedido['DATA_PEDIDO']}
🏢 Filial: {pedido['CODFILIAL']}
👨‍💼 Vendedor: {pedido.get('VENDEDOR', 'N/A')}
📋 NF: {pedido.get('NUMNOTA', 'N/A')}
⏰ Checkout Inicial: {pedido.get('DTINICIALCHECKOUT', 'N/A')}
⏰ Checkout Final: {pedido.get('DTFINALCHECKOUT', 'N/A')}
💰 Valor Total: R$ {valor_total:,.2f}
📊 Quantidade Total: {qtd_total:.0f}

📦 PRODUTOS:
"""
        for _, item in itens.iterrows():
            info += f"  • {item['PRODUTO']} (Cód: {item['CODPROD']})\n"
            info += f"    Qtd: {item['QUANTIDADE']:.0f} | Preço: R$ {item['PRECO_UNIT']:,.2f} | Total: R$ {item['VALOR_TOTAL_ITEM']:,.2f}\n"
        
        pedidos_info.append(info)
    
    return "\n" + "="*80 + "\n".join(pedidos_info)

def enviar_notificacao(titulo, mensagem, duracao=10):
    """Envia notificação Windows"""
    if not NOTIFICACAO_DISPONIVEL:
        print(f"🔔 {titulo}: {mensagem}")
        return
    
    try:
        if ToastNotifier is not None:
            toaster = ToastNotifier()
            toaster.show_toast(titulo, mensagem, duration=duracao, threaded=True)
        elif NOTIFICACAO_PLYER and notification is not None:
            notification.notify(
                title=titulo,
                message=mensagem,
                timeout=duracao,
                app_name="Monitor de Pedidos"
            )
        else:
            print(f"🔔 {titulo}: {mensagem}")
    except Exception as e:
        print(f"❌ Erro ao enviar notificação: {e}")

# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class MonitorPedidosCheckout:
    def __init__(self, root):
        self.root = root
        self.root.title("🔔 Monitor de Pedidos - Checkout")
        self.root.geometry("900x700")
        
        self.monitorando = False
        self.thread_monitor = None
        self.conn = None
        self.pedidos_notificados = carregar_pedidos_notificados()
        self.intervalo = 30  # segundos
        
        self.criar_interface()
        self.conectar_banco()
        
        # Verificação imediata ao abrir (após conectar)
        if self.conn:
            self.root.after(1000, self.verificacao_imediata)
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(
            main_frame, 
            text="🔔 Monitor de Pedidos - Checkout", 
            font=("Arial", 16, "bold")
        )
        titulo.pack(pady=(0, 20))
        
        # Frame de controle
        frame_controle = ttk.LabelFrame(main_frame, text="Controle", padding="15")
        frame_controle.pack(fill=tk.X, pady=10)
        
        # Status
        self.label_status = ttk.Label(
            frame_controle, 
            text="Status: Parado", 
            font=("Arial", 11, "bold"),
            foreground="red"
        )
        self.label_status.pack(side=tk.LEFT, padx=10)
        
        # Intervalo
        frame_intervalo = ttk.Frame(frame_controle)
        frame_intervalo.pack(side=tk.LEFT, padx=20)
        ttk.Label(frame_intervalo, text="Intervalo (segundos):").pack(side=tk.LEFT, padx=5)
        self.spin_intervalo = ttk.Spinbox(
            frame_intervalo, 
            from_=10, 
            to=300, 
            width=10,
            value=30
        )
        self.spin_intervalo.pack(side=tk.LEFT, padx=5)
        
        # Botões
        frame_botoes = ttk.Frame(frame_controle)
        frame_botoes.pack(side=tk.RIGHT, padx=10)
        
        self.btn_iniciar = ttk.Button(
            frame_botoes, 
            text="▶ Iniciar Monitoramento", 
            command=self.iniciar_monitoramento,
            style="Accent.TButton"
        )
        self.btn_iniciar.pack(side=tk.LEFT, padx=5)
        
        self.btn_parar = ttk.Button(
            frame_botoes, 
            text="⏹ Parar Monitoramento", 
            command=self.parar_monitoramento,
            state=tk.DISABLED
        )
        self.btn_parar.pack(side=tk.LEFT, padx=5)
        
        self.btn_verificar = ttk.Button(
            frame_botoes, 
            text="🔍 Verificar Agora", 
            command=self.verificacao_imediata
        )
        self.btn_verificar.pack(side=tk.LEFT, padx=5)
        
        self.btn_limpar = ttk.Button(
            frame_botoes, 
            text="🗑 Limpar Histórico", 
            command=self.limpar_historico
        )
        self.btn_limpar.pack(side=tk.LEFT, padx=5)
        
        # Frame de informações
        frame_info = ttk.LabelFrame(main_frame, text="Informações", padding="15")
        frame_info.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Estatísticas
        frame_stats = ttk.Frame(frame_info)
        frame_stats.pack(fill=tk.X, pady=5)
        
        self.label_stats = ttk.Label(
            frame_stats,
            text="Pedidos notificados: 0 | Última verificação: Nunca",
            font=("Arial", 10)
        )
        self.label_stats.pack()
        
        # Log
        ttk.Label(frame_info, text="Log de Atividades:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        self.text_log = scrolledtext.ScrolledText(
            frame_info,
            height=20,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.text_log.pack(fill=tk.BOTH, expand=True)
        
        # Aviso sobre notificações
        if not NOTIFICACAO_DISPONIVEL:
            aviso = ttk.Label(
                frame_info,
                text="⚠️ Bibliotecas de notificação não encontradas. Instale: pip install win10toast ou pip install plyer",
                foreground="orange",
                font=("Arial", 9)
            )
            aviso.pack(pady=5)
    
    def log(self, mensagem):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.text_log.insert(tk.END, f"[{timestamp}] {mensagem}\n")
        self.text_log.see(tk.END)
        self.root.update_idletasks()
    
    def conectar_banco(self):
        """Conecta ao banco de dados"""
        self.log("🔌 Conectando ao banco de dados...")
        self.conn = get_db_connection()
        if self.conn:
            self.log("✅ Conectado ao banco de dados com sucesso!")
        else:
            self.log("❌ Falha ao conectar ao banco de dados!")
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados!")
    
    def iniciar_monitoramento(self):
        """Inicia o monitoramento"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados!")
            return
        
        if self.monitorando:
            return
        
        try:
            self.intervalo = int(self.spin_intervalo.get())
        except ValueError:
            messagebox.showerror("Erro", "Intervalo inválido!")
            return
        
        self.monitorando = True
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_parar.config(state=tk.NORMAL)
        self.label_status.config(text="Status: Monitorando...", foreground="green")
        
        self.log(f"▶ Monitoramento iniciado (intervalo: {self.intervalo}s)")
        
        # Inicia thread de monitoramento
        self.thread_monitor = threading.Thread(target=self.loop_monitoramento, daemon=True)
        self.thread_monitor.start()
    
    def parar_monitoramento(self):
        """Para o monitoramento"""
        self.monitorando = False
        self.btn_iniciar.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED)
        self.label_status.config(text="Status: Parado", foreground="red")
        self.log("⏹ Monitoramento parado")
    
    def limpar_historico(self):
        """Limpa o histórico de pedidos notificados"""
        if messagebox.askyesno("Confirmar", "Deseja limpar o histórico de pedidos notificados?"):
            self.pedidos_notificados = set()
            if Path(ARQUIVO_NOTIFICADOS).exists():
                Path(ARQUIVO_NOTIFICADOS).unlink()
            self.log("🗑 Histórico limpo!")
            self.atualizar_stats()
    
    def atualizar_stats(self):
        """Atualiza estatísticas na interface"""
        qtd_notificados = len(self.pedidos_notificados)
        data_hoje = datetime.now().strftime('%d/%m/%Y')
        self.label_stats.config(
            text=f"Pedidos notificados: {qtd_notificados} | Última verificação: {datetime.now().strftime('%H:%M:%S')} | Data: {data_hoje}"
        )
    
    def verificacao_imediata(self):
        """Faz uma verificação imediata dos pedidos do dia atual"""
        if not self.conn:
            self.log("❌ Não há conexão com o banco de dados!")
            return
        
        self.log("🔍 Verificação imediata do dia atual...")
        # Executa em thread separada para não travar a interface
        thread_verificacao = threading.Thread(target=lambda: self.processar_pedidos(primeira_vez=True), daemon=True)
        thread_verificacao.start()
    
    def processar_pedidos(self, primeira_vez=False):
        """
        Processa pedidos do dia atual e envia notificações
        primeira_vez: True se for a primeira verificação ao abrir o programa
        """
        try:
            # Busca apenas pedidos do dia atual
            df_pedidos = buscar_pedidos_checkout_completo(self.conn, apenas_hoje=True)
            self.log(f"📆 Pedidos encontrados (dia atual): {len(df_pedidos)}")
            
            if not df_pedidos.empty:
                # Agrupa por NUMPED
                pedidos_unicos = df_pedidos['NUMPED'].unique()
                novos_pedidos = []
                
                for numped in pedidos_unicos:
                    numped_int = int(numped)
                    if numped_int not in self.pedidos_notificados:
                        novos_pedidos.append(numped_int)
                
                if novos_pedidos:
                    if primeira_vez:
                        self.log(f"📋 {len(novos_pedidos)} pedido(s) do dia atual encontrado(s) ao abrir o programa!")
                    else:
                        self.log(f"🔔 {len(novos_pedidos)} novo(s) pedido(s) encontrado(s)!")
                    
                    for numped in novos_pedidos:
                        # Busca detalhes do pedido
                        pedido_df = df_pedidos[df_pedidos['NUMPED'] == numped]
                        primeiro_item = pedido_df.iloc[0]
                        
                        # Prepara notificação
                        titulo = f"✅ Pedido {numped} - Checkout Completo!"
                        mensagem = f"Cliente: {primeiro_item['CLIENTE']}\n"
                        mensagem += f"Valor Total: R$ {pedido_df['VALOR_TOTAL_ITEM'].sum():,.2f}\n"
                        mensagem += f"Produtos: {len(pedido_df)} item(ns)"
                        
                        # Envia notificação
                        enviar_notificacao(titulo, mensagem, duracao=15)
                        
                        # Salva como notificado
                        salvar_pedido_notificado(numped)
                        self.pedidos_notificados.add(numped)
                        
                        # Log detalhado
                        detalhes = formatar_detalhes_pedido(pedido_df)
                        self.log(detalhes)
                        
                        self.log(f"✅ Pedido {numped} notificado!")
                else:
                    if primeira_vez:
                        self.log("ℹ️ Nenhum pedido novo do dia atual encontrado.")
                    else:
                        self.log("ℹ️ Nenhum pedido novo encontrado.")
            else:
                if primeira_vez:
                    self.log("ℹ️ Nenhum pedido com checkout completo encontrado no dia atual.")
                else:
                    self.log("ℹ️ Nenhum pedido com checkout completo encontrado no dia atual.")
            
            self.atualizar_stats()
            
        except Exception as e:
            self.log(f"❌ Erro ao processar pedidos: {e}")
            import traceback
            self.log(traceback.format_exc())
    
    def loop_monitoramento(self):
        """Loop principal de monitoramento"""
        while self.monitorando:
            try:
                # Reconecta se necessário
                if not self.conn:
                    self.conn = get_db_connection()
                    if not self.conn:
                        self.log("❌ Erro: Sem conexão com o banco!")
                        time.sleep(self.intervalo)
                        continue
                
                # Busca pedidos
                self.log("🔍 Verificando pedidos do dia atual...")
                self.processar_pedidos()
                
            except Exception as e:
                self.log(f"❌ Erro no monitoramento: {e}")
                import traceback
                self.log(traceback.format_exc())
            
            # Aguarda intervalo
            for _ in range(self.intervalo):
                if not self.monitorando:
                    break
                time.sleep(1)
    
    def __del__(self):
        """Fecha conexão ao destruir"""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = MonitorPedidosCheckout(root)
    root.mainloop()

# portfolio-commit-ready: monitor_pedidos_checkout
