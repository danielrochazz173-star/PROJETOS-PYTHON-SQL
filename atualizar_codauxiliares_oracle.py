import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import oracledb
from pathlib import Path
import os

# =========================================================================
# CONFIGURAÇÕES DE BANCO
# =========================================================================
DB_HOST = os.getenv('DB_HOST', '10.0.0.10')
DB_PORT = int(os.getenv('DB_PORT', '1521'))
DB_SERVICE = os.getenv('DB_SERVICE', 'PROD')

# Credenciais padrão (podem ser alteradas no login)
DEFAULT_USER = 'powerbi'
DEFAULT_PASSWORD = os.getenv('DB_PASSWORD', '')

# Credenciais ativas (serão definidas no login)
ACTIVE_USER = None
ACTIVE_PASSWORD = None


def get_db_connection():
    """Conecta ao banco Oracle usando as credenciais ativas."""
    try:
        try:
            oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
        except Exception:
            pass
        
        return oracledb.connect(
            user=ACTIVE_USER,
            password=ACTIVE_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        return None


class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Login - Atualizar Códigos Auxiliares")
        self.root.geometry("400x250")
        self.root.resizable(False, False)
        
        # Centraliza a janela
        self.centralizar_janela()
        
        self.login_ok = False
        self.criar_interface()
    
    def centralizar_janela(self):
        """Centraliza a janela na tela."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def criar_interface(self):
        """Cria a interface de login."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            main_frame, 
            text="🔐 LOGIN ORACLE", 
            font=('Arial', 14, 'bold')
        ).pack(pady=(0, 20))
        
        # Frame de campos
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill='x', pady=10)
        
        # Usuário
        ttk.Label(fields_frame, text="Usuário:", width=10).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.entry_user = ttk.Entry(fields_frame, width=25)
        self.entry_user.grid(row=0, column=1, padx=5, pady=5)
        self.entry_user.insert(0, DEFAULT_USER)
        
        # Senha
        ttk.Label(fields_frame, text="Senha:", width=10).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.entry_password = ttk.Entry(fields_frame, width=25, show='*')
        self.entry_password.grid(row=1, column=1, padx=5, pady=5)
        self.entry_password.insert(0, DEFAULT_PASSWORD)
        
        # Bind Enter para fazer login
        self.entry_user.bind('<Return>', lambda e: self.entry_password.focus())
        self.entry_password.bind('<Return>', lambda e: self.fazer_login())
        
        # Frame de botões
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(pady=20)
        
        ttk.Button(
            buttons_frame, 
            text="Conectar", 
            command=self.fazer_login,
            width=12
        ).pack(side='left', padx=5)
        
        ttk.Button(
            buttons_frame, 
            text="Cancelar", 
            command=self.root.quit,
            width=12
        ).pack(side='left', padx=5)
        
        # Foco no campo de usuário
        self.entry_user.focus()
    
    def fazer_login(self):
        """Tenta fazer login no Oracle."""
        global ACTIVE_USER, ACTIVE_PASSWORD
        
        usuario = self.entry_user.get().strip()
        senha = self.entry_password.get().strip()
        
        if not usuario or not senha:
            messagebox.showwarning("Aviso", "Preencha usuário e senha!")
            return
        
        # Testa conexão
        try:
            try:
                oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
            except Exception:
                pass
            
            conn = oracledb.connect(
                user=usuario,
                password=senha,
                dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
            )
            
            conn.close()
            
            # Login OK - define credenciais ativas
            ACTIVE_USER = usuario
            ACTIVE_PASSWORD = senha
            self.login_ok = True
            
            messagebox.showinfo("Sucesso", f"Login realizado com sucesso!\n\nUsuário: {usuario}")
            self.root.destroy()
            
        except oracledb.DatabaseError as e:
            messagebox.showerror(
                "Erro de Conexão",
                f"Não foi possível conectar ao Oracle!\n\n"
                f"Erro: {str(e)}\n\n"
                f"Verifique usuário e senha."
            )


class AtualizadorCodAuxiliares:
    def __init__(self, root):
        self.root = root
        self.root.title("Atualizar Códigos Auxiliares no Oracle")
        self.root.geometry("900x600")
        
        self.df = None
        self.arquivo_excel = None
        
        self.criar_interface()
        self.carregar_ultimo_excel()
    
    def criar_interface(self):
        """Cria a interface gráfica."""
        # Frame superior - Info do usuário logado
        frame_usuario = ttk.Frame(self.root, padding=5)
        frame_usuario.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(
            frame_usuario, 
            text=f"👤 Conectado como: {ACTIVE_USER}",
            font=('Arial', 9, 'bold'),
            foreground='blue'
        ).pack(side='left')
        
        # Frame - Seleção de arquivo
        frame_arquivo = ttk.LabelFrame(self.root, text="Arquivo Excel", padding=10)
        frame_arquivo.pack(fill='x', padx=10, pady=5)
        
        self.lbl_arquivo = ttk.Label(frame_arquivo, text="Nenhum arquivo carregado", foreground="red")
        self.lbl_arquivo.pack(side='left', padx=5)
        
        ttk.Button(frame_arquivo, text="Selecionar Excel", command=self.selecionar_excel).pack(side='right', padx=5)
        
        # Frame do meio - Busca por CODPROD
        frame_busca = ttk.LabelFrame(self.root, text="Buscar Produto Específico", padding=10)
        frame_busca.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_busca, text="CODPROD:").pack(side='left', padx=5)
        
        self.entry_codprod = ttk.Entry(frame_busca, width=15)
        self.entry_codprod.pack(side='left', padx=5)
        self.entry_codprod.bind('<Return>', lambda e: self.buscar_produto())
        
        ttk.Button(frame_busca, text="Buscar", command=self.buscar_produto).pack(side='left', padx=5)
        ttk.Button(frame_busca, text="Limpar", command=self.limpar_busca).pack(side='left', padx=5)
        
        # Frame de informações
        frame_info = ttk.LabelFrame(self.root, text="Informações", padding=10)
        frame_info.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.text_info = tk.Text(frame_info, height=20, width=80, wrap='word')
        self.text_info.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(frame_info, command=self.text_info.yview)
        scrollbar.pack(side='right', fill='y')
        self.text_info.config(yscrollcommand=scrollbar.set)
        
        # Frame de ações
        frame_acoes = ttk.Frame(self.root, padding=10)
        frame_acoes.pack(fill='x', padx=10, pady=5)
        
        self.btn_atualizar = ttk.Button(
            frame_acoes, 
            text="Atualizar Produto Selecionado", 
            command=self.atualizar_produto,
            state='disabled'
        )
        self.btn_atualizar.pack(side='left', padx=5)
        
        self.btn_atualizar_todos = ttk.Button(
            frame_acoes, 
            text="Atualizar TODOS os Produtos", 
            command=self.atualizar_todos,
            state='disabled'
        )
        self.btn_atualizar_todos.pack(side='left', padx=5)
        
        ttk.Button(frame_acoes, text="Sair", command=self.root.quit).pack(side='right', padx=5)
    
    def carregar_ultimo_excel(self):
        """Carrega automaticamente o último Excel gerado."""
        pasta = Path(__file__).parent
        arquivos = list(pasta.glob("produtos_codauxiliares_preenchidos_*.xlsx"))
        
        if arquivos:
            # Pega o mais recente
            arquivo_mais_recente = max(arquivos, key=lambda p: p.stat().st_mtime)
            self.carregar_excel(arquivo_mais_recente)
    
    def selecionar_excel(self):
        """Permite selecionar um arquivo Excel manualmente."""
        arquivo = filedialog.askopenfilename(
            title="Selecione o Excel com os produtos preenchidos",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if arquivo:
            self.carregar_excel(Path(arquivo))
    
    def carregar_excel(self, caminho):
        """Carrega o Excel para memória."""
        try:
            self.df = pd.read_excel(caminho, sheet_name='Produtos_Preenchidos')
            
            # Filtra apenas produtos que foram preenchidos (que têm xml_file)
            self.df = self.df[self.df['xml_file'].notna()].copy()
            
            self.arquivo_excel = caminho
            self.lbl_arquivo.config(
                text=f"✓ {caminho.name} ({len(self.df)} produtos preenchidos)",
                foreground="green"
            )
            
            self.btn_atualizar_todos.config(state='normal')
            
            self.text_info.delete(1.0, tk.END)
            self.text_info.insert(1.0, f"Excel carregado com sucesso!\n\n")
            self.text_info.insert(tk.END, f"Total de produtos preenchidos: {len(self.df)}\n\n")
            self.text_info.insert(tk.END, "Digite um CODPROD para buscar especificamente\n")
            self.text_info.insert(tk.END, "ou clique em 'Atualizar TODOS' para atualizar todos de uma vez.")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar Excel:\n{e}")
    
    def buscar_produto(self):
        """Busca um produto específico pelo CODPROD."""
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue um arquivo Excel primeiro!")
            return
        
        codprod = self.entry_codprod.get().strip()
        
        if not codprod:
            messagebox.showwarning("Aviso", "Digite um CODPROD!")
            return
        
        try:
            codprod = int(codprod)
        except ValueError:
            messagebox.showerror("Erro", "CODPROD deve ser um número!")
            return
        
        # Busca no DataFrame
        produto = self.df[self.df['CODPROD'] == codprod]
        
        if produto.empty:
            messagebox.showinfo(
                "Não encontrado", 
                f"Produto {codprod} não encontrado no Excel.\n\n"
                "Isso pode significar que:\n"
                "- O produto não tem códigos auxiliares zerados\n"
                "- O produto não foi encontrado nos XMLs\n"
                "- O CODPROD não existe"
            )
            return
        
        # Mostra informações do produto
        self.mostrar_produto(produto.iloc[0])
        self.btn_atualizar.config(state='normal')
    
    def mostrar_produto(self, produto):
        """Mostra as informações de um produto."""
        self.text_info.delete(1.0, tk.END)
        
        self.text_info.insert(1.0, "="*80 + "\n")
        self.text_info.insert(tk.END, "PRODUTO ENCONTRADO\n")
        self.text_info.insert(tk.END, "="*80 + "\n\n")
        
        self.text_info.insert(tk.END, f"CODPROD: {produto['CODPROD']}\n")
        self.text_info.insert(tk.END, f"CODFAB: {produto['CODFAB']}\n")
        self.text_info.insert(tk.END, f"DESCRIÇÃO: {produto['DESCRICAO']}\n")
        self.text_info.insert(tk.END, f"DEPARTAMENTO: {produto['DEPARTAMENTO']}\n")
        self.text_info.insert(tk.END, f"FORNECEDOR: {produto['FORNECEDOR']}\n\n")
        
        self.text_info.insert(tk.END, "-"*80 + "\n")
        self.text_info.insert(tk.END, "CÓDIGOS AUXILIARES QUE SERÃO ATUALIZADOS:\n")
        self.text_info.insert(tk.END, "-"*80 + "\n\n")
        
        codaux = produto['CODAUXILIAR'] if pd.notna(produto['CODAUXILIAR']) else '(vazio)'
        codaux2 = produto['CODAUXILIAR2'] if pd.notna(produto['CODAUXILIAR2']) else '(vazio)'
        codauxtrib = produto['CODAUXILIARTRIB'] if pd.notna(produto['CODAUXILIARTRIB']) else '(vazio)'
        
        self.text_info.insert(tk.END, f"CODAUXILIAR: {codaux}\n")
        self.text_info.insert(tk.END, f"CODAUXILIAR2: {codaux2}\n")
        self.text_info.insert(tk.END, f"CODAUXILIARTRIB: {codauxtrib}\n\n")
        
        self.text_info.insert(tk.END, f"Origem: {produto['xml_file']}\n\n")
        self.text_info.insert(tk.END, "="*80 + "\n")
        self.text_info.insert(tk.END, "✓ Clique em 'Atualizar Produto Selecionado' para confirmar\n")
        self.text_info.insert(tk.END, "="*80 + "\n")
    
    def limpar_busca(self):
        """Limpa a busca e volta ao estado inicial."""
        self.entry_codprod.delete(0, tk.END)
        self.btn_atualizar.config(state='disabled')
        
        if self.df is not None:
            self.text_info.delete(1.0, tk.END)
            self.text_info.insert(1.0, f"Total de produtos preenchidos: {len(self.df)}\n\n")
            self.text_info.insert(tk.END, "Digite um CODPROD para buscar especificamente\n")
            self.text_info.insert(tk.END, "ou clique em 'Atualizar TODOS' para atualizar todos de uma vez.")
    
    def atualizar_produto(self):
        """Atualiza um produto específico no Oracle."""
        if self.df is None:
            return
        
        codprod = self.entry_codprod.get().strip()
        if not codprod:
            return
        
        try:
            codprod = int(codprod)
        except ValueError:
            return
        
        produto = self.df[self.df['CODPROD'] == codprod]
        if produto.empty:
            return
        
        # Confirmação
        confirmacao = messagebox.askyesno(
            "Confirmar Atualização",
            f"Confirma a atualização do produto {codprod}?\n\n"
            f"{produto.iloc[0]['DESCRICAO']}\n\n"
            "Os códigos auxiliares serão preenchidos no Oracle."
        )
        
        if not confirmacao:
            return
        
        # Executa UPDATE
        self.executar_update([produto.iloc[0]])
    
    def atualizar_todos(self):
        """Atualiza todos os produtos de uma vez."""
        if self.df is None or len(self.df) == 0:
            return
        
        # Confirmação com aviso
        confirmacao = messagebox.askyesno(
            "⚠️ CONFIRMAR ATUALIZAÇÃO EM MASSA",
            f"ATENÇÃO!\n\n"
            f"Você está prestes a atualizar {len(self.df)} produtos de uma só vez!\n\n"
            f"Os códigos auxiliares de TODOS esses produtos serão preenchidos no Oracle.\n\n"
            f"TEM CERTEZA QUE DESEJA CONTINUAR?",
            icon='warning'
        )
        
        if not confirmacao:
            return
        
        # Segunda confirmação
        confirmacao2 = messagebox.askyesno(
            "⚠️ ÚLTIMA CONFIRMAÇÃO",
            f"Esta é sua ÚLTIMA CHANCE!\n\n"
            f"Atualizar {len(self.df)} produtos?\n\n"
            f"Continuar?",
            icon='warning'
        )
        
        if not confirmacao2:
            return
        
        # Executa UPDATEs
        self.executar_update(self.df.to_dict('records'))
    
    def executar_update(self, produtos):
        """Executa os UPDATEs no Oracle."""
        conn = get_db_connection()
        if not conn:
            messagebox.showerror("Erro", "Não foi possível conectar ao banco Oracle!")
            return
        
        try:
            cursor = conn.cursor()
            
            sucesso = 0
            erros = 0
            
            for produto in produtos:
                try:
                    codprod = produto['CODPROD']
                    codaux = produto['CODAUXILIAR'] if pd.notna(produto['CODAUXILIAR']) else None
                    codaux2 = produto['CODAUXILIAR2'] if pd.notna(produto['CODAUXILIAR2']) else None
                    codauxtrib = produto['CODAUXILIARTRIB'] if pd.notna(produto['CODAUXILIARTRIB']) else None
                    
                    # UPDATE
                    sql = """
                        UPDATE PCPRODUT
                        SET CODAUXILIAR = :codaux,
                            CODAUXILIAR2 = :codaux2,
                            CODAUXILIARTRIB = :codauxtrib
                        WHERE CODPROD = :codprod
                    """
                    
                    cursor.execute(sql, {
                        'codaux': codaux,
                        'codaux2': codaux2,
                        'codauxtrib': codauxtrib,
                        'codprod': codprod
                    })
                    
                    sucesso += 1
                    
                except Exception as e:
                    erros += 1
                    print(f"Erro ao atualizar produto {produto['CODPROD']}: {e}")
            
            # Commit
            conn.commit()
            
            # Mensagem de sucesso
            messagebox.showinfo(
                "Atualização Concluída",
                f"✓ Atualização concluída!\n\n"
                f"Produtos atualizados: {sucesso}\n"
                f"Erros: {erros}"
            )
            
            # Atualiza texto
            self.text_info.delete(1.0, tk.END)
            self.text_info.insert(1.0, f"✓ ATUALIZAÇÃO CONCLUÍDA\n\n")
            self.text_info.insert(tk.END, f"Produtos atualizados: {sucesso}\n")
            self.text_info.insert(tk.END, f"Erros: {erros}\n")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar atualização:\n{e}")
            conn.rollback()
        
        finally:
            cursor.close()
            conn.close()


def main():
    # Primeira janela: Login
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()
    
    # Se login OK, abre janela principal
    if login_app.login_ok:
        root = tk.Tk()
        app = AtualizadorCodAuxiliares(root)
        root.mainloop()


if __name__ == "__main__":
    main()

# portfolio-commit-ready: atualizar_codauxiliares_oracle
