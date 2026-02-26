import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import xml.etree.ElementTree as ET
from pathlib import Path
import threading

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'

DEFAULT_USER = 'powerbi'
DEFAULT_PASSWORD = 'cbjc4xp3nlq6'

PASTA_XMLS = r"C:\Users\DESENVOLVIMENTO\Documents\ALL_XMLS"

# Namespace do XML de NF-e
NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

# Credenciais ativas
ACTIVE_USER = None
ACTIVE_PASSWORD = None


def get_db_connection():
    """Conecta ao banco Oracle."""
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
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(
            main_frame, 
            text="🔐 LOGIN ORACLE", 
            font=('Arial', 14, 'bold')
        ).pack(pady=(0, 20))
        
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill='x', pady=10)
        
        ttk.Label(fields_frame, text="Usuário:", width=10).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.entry_user = ttk.Entry(fields_frame, width=25)
        self.entry_user.grid(row=0, column=1, padx=5, pady=5)
        self.entry_user.insert(0, DEFAULT_USER)
        
        ttk.Label(fields_frame, text="Senha:", width=10).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.entry_password = ttk.Entry(fields_frame, width=25, show='*')
        self.entry_password.grid(row=1, column=1, padx=5, pady=5)
        self.entry_password.insert(0, DEFAULT_PASSWORD)
        
        self.entry_user.bind('<Return>', lambda e: self.entry_password.focus())
        self.entry_password.bind('<Return>', lambda e: self.fazer_login())
        
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
        
        self.entry_user.focus()
    
    def fazer_login(self):
        """Tenta fazer login no Oracle."""
        global ACTIVE_USER, ACTIVE_PASSWORD
        
        usuario = self.entry_user.get().strip()
        senha = self.entry_password.get().strip()
        
        if not usuario or not senha:
            messagebox.showwarning("Aviso", "Preencha usuário e senha!")
            return
        
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
            
            ACTIVE_USER = usuario
            ACTIVE_PASSWORD = senha
            self.login_ok = True
            
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
        self.root.geometry("900x650")
        
        self.produtos_vazios = {}  # {codprod: {...dados...}}
        self.dados_xml = {}  # {codfab: {...dados...}}
        self.produtos_preenchidos = {}  # {codprod: {...dados completos...}}
        
        self.criar_interface()
        self.carregar_dados()
    
    def criar_interface(self):
        """Cria a interface gráfica."""
        # Frame superior - Info do usuário
        frame_usuario = ttk.Frame(self.root, padding=5)
        frame_usuario.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(
            frame_usuario, 
            text=f"👤 Conectado como: {ACTIVE_USER}",
            font=('Arial', 9, 'bold'),
            foreground='blue'
        ).pack(side='left')
        
        # Frame de status
        frame_status = ttk.LabelFrame(self.root, text="Status", padding=10)
        frame_status.pack(fill='x', padx=10, pady=5)
        
        self.lbl_status = ttk.Label(frame_status, text="Carregando dados...", foreground="orange")
        self.lbl_status.pack()
        
        # Progress bar
        self.progress = ttk.Progressbar(frame_status, mode='indeterminate')
        self.progress.pack(fill='x', pady=5)
        
        # Frame de busca
        frame_busca = ttk.LabelFrame(self.root, text="Buscar Produto Específico", padding=10)
        frame_busca.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(frame_busca, text="CODPROD:").pack(side='left', padx=5)
        
        self.entry_codprod = ttk.Entry(frame_busca, width=15, state='disabled')
        self.entry_codprod.pack(side='left', padx=5)
        self.entry_codprod.bind('<Return>', lambda e: self.buscar_produto())
        
        self.btn_buscar = ttk.Button(frame_busca, text="Buscar", command=self.buscar_produto, state='disabled')
        self.btn_buscar.pack(side='left', padx=5)
        
        self.btn_limpar = ttk.Button(frame_busca, text="Limpar", command=self.limpar_busca, state='disabled')
        self.btn_limpar.pack(side='left', padx=5)
        
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
    
    def carregar_dados(self):
        """Carrega dados do Oracle e XMLs em background."""
        self.progress.start()
        
        thread = threading.Thread(target=self._carregar_dados_thread, daemon=True)
        thread.start()
    
    def _carregar_dados_thread(self):
        """Thread para carregar dados."""
        try:
            # 1. Busca produtos vazios no Oracle
            self.root.after(0, lambda: self.lbl_status.config(text="Conectando ao Oracle..."))
            
            conn = get_db_connection()
            if not conn:
                self.root.after(0, lambda: self.mostrar_erro("Erro ao conectar no Oracle!"))
                return
            
            self.root.after(0, lambda: self.lbl_status.config(text="Buscando produtos com códigos vazios..."))
            
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
            
            cursor = conn.cursor()
            cursor.execute(query)
            
            for row in cursor:
                codprod = row[0]
                self.produtos_vazios[codprod] = {
                    'CODPROD': row[0],
                    'CODFAB': str(row[1]).strip() if row[1] else '',
                    'DESCRICAO': row[2],
                    'CODEPTO': row[3],
                    'DEPARTAMENTO': row[4],
                    'CODFORNEC': row[5],
                    'FORNECEDOR': row[6],
                    'EMBALAGEM': row[7],
                    'UNIDADE': row[8],
                    'QT_VENDIDA_6M': row[9]
                }
            
            cursor.close()
            conn.close()
            
            total_vazios = len(self.produtos_vazios)
            self.root.after(0, lambda: self.lbl_status.config(
                text=f"Encontrados {total_vazios} produtos vazios. Processando XMLs..."
            ))
            
            # 2. Processa XMLs
            pasta = Path(PASTA_XMLS)
            if not pasta.exists():
                self.root.after(0, lambda: self.mostrar_erro(f"Pasta de XMLs não encontrada: {PASTA_XMLS}"))
                return
            
            xml_files = list(pasta.glob('*.xml'))
            total_xmls = len(xml_files)
            
            for i, xml_file in enumerate(xml_files, 1):
                if i % 10 == 0:
                    self.root.after(0, lambda i=i, t=total_xmls: self.lbl_status.config(
                        text=f"Processando XMLs: {i}/{t}..."
                    ))
                
                produtos = self.extrair_dados_xml(xml_file)
                for p in produtos:
                    codfab = p['codfab']
                    if codfab not in self.dados_xml:
                        self.dados_xml[codfab] = p
            
            # 3. Faz o match
            self.root.after(0, lambda: self.lbl_status.config(text="Fazendo match entre produtos e XMLs..."))
            
            for codprod, dados in self.produtos_vazios.items():
                codfab = dados['CODFAB']
                
                if codfab and codfab in self.dados_xml:
                    # Encontrou no XML!
                    xml_data = self.dados_xml[codfab]
                    
                    self.produtos_preenchidos[codprod] = {
                        **dados,
                        'CODAUXILIAR': xml_data['cEAN'],
                        'CODAUXILIAR2': xml_data['cEAN'],
                        'CODAUXILIARTRIB': xml_data['cEANTrib'],
                        'xml_file': xml_data['xml_file'],
                        'descricao_xml': xml_data['descricao_xml']
                    }
            
            # 4. Finaliza
            total_preenchidos = len(self.produtos_preenchidos)
            total_nao_encontrados = total_vazios - total_preenchidos
            
            self.root.after(0, lambda: self.finalizar_carregamento(
                total_vazios, total_preenchidos, total_nao_encontrados
            ))
            
        except Exception as e:
            self.root.after(0, lambda: self.mostrar_erro(f"Erro ao carregar dados:\n{e}"))
    
    def extrair_dados_xml(self, xml_path):
        """Extrai dados de produtos de um XML."""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            xml_filename = Path(xml_path).name
            
            produtos = []
            
            for det in root.findall('.//nfe:det', NS):
                prod = det.find('nfe:prod', NS)
                if prod is not None:
                    cprod = prod.find('nfe:cProd', NS)
                    cean = prod.find('nfe:cEAN', NS)
                    ceantrib = prod.find('nfe:cEANTrib', NS)
                    xprod = prod.find('nfe:xProd', NS)
                    
                    if cprod is not None and cprod.text:
                        codfab = cprod.text.strip()
                        ean = cean.text.strip() if cean is not None and cean.text else ''
                        ean_trib = ceantrib.text.strip() if ceantrib is not None and ceantrib.text else ''
                        descricao = xprod.text.strip() if xprod is not None and xprod.text else ''
                        
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
        
        except Exception:
            return []
    
    def finalizar_carregamento(self, total_vazios, total_preenchidos, total_nao_encontrados):
        """Finaliza o carregamento e habilita interface."""
        self.progress.stop()
        self.progress.pack_forget()
        
        self.lbl_status.config(
            text=f"✓ Pronto! {total_preenchidos} produtos podem ser atualizados | "
                 f"{total_nao_encontrados} não encontrados nos XMLs",
            foreground="green"
        )
        
        self.entry_codprod.config(state='normal')
        self.btn_buscar.config(state='normal')
        self.btn_limpar.config(state='normal')
        
        if total_preenchidos > 0:
            self.btn_atualizar_todos.config(state='normal')
        
        self.text_info.delete(1.0, tk.END)
        self.text_info.insert(1.0, f"✓ DADOS CARREGADOS COM SUCESSO!\n\n")
        self.text_info.insert(tk.END, f"Produtos com códigos vazios: {total_vazios}\n")
        self.text_info.insert(tk.END, f"Produtos encontrados nos XMLs: {total_preenchidos}\n")
        self.text_info.insert(tk.END, f"Produtos NÃO encontrados: {total_nao_encontrados}\n\n")
        self.text_info.insert(tk.END, "="*80 + "\n\n")
        self.text_info.insert(tk.END, "Digite um CODPROD para buscar especificamente\n")
        self.text_info.insert(tk.END, "ou clique em 'Atualizar TODOS' para atualizar todos de uma vez.")
    
    def mostrar_erro(self, mensagem):
        """Mostra erro e fecha aplicação."""
        self.progress.stop()
        messagebox.showerror("Erro", mensagem)
        self.root.quit()
    
    def buscar_produto(self):
        """Busca um produto específico."""
        codprod = self.entry_codprod.get().strip()
        
        if not codprod:
            messagebox.showwarning("Aviso", "Digite um CODPROD!")
            return
        
        try:
            codprod = int(codprod)
        except ValueError:
            messagebox.showerror("Erro", "CODPROD deve ser um número!")
            return
        
        if codprod not in self.produtos_preenchidos:
            messagebox.showinfo(
                "Não encontrado", 
                f"Produto {codprod} não está na lista de produtos que podem ser atualizados.\n\n"
                "Isso significa que:\n"
                "- O produto não tem códigos auxiliares zerados, ou\n"
                "- O produto não foi encontrado nos XMLs"
            )
            return
        
        produto = self.produtos_preenchidos[codprod]
        self.mostrar_produto(produto)
        self.btn_atualizar.config(state='normal')
    
    def mostrar_produto(self, produto):
        """Mostra informações do produto."""
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
        
        self.text_info.insert(tk.END, f"CODAUXILIAR: {produto['CODAUXILIAR']}\n")
        self.text_info.insert(tk.END, f"CODAUXILIAR2: {produto['CODAUXILIAR2']}\n")
        self.text_info.insert(tk.END, f"CODAUXILIARTRIB: {produto['CODAUXILIARTRIB']}\n\n")
        
        self.text_info.insert(tk.END, f"Origem: {produto['xml_file']}\n\n")
        self.text_info.insert(tk.END, "="*80 + "\n")
        self.text_info.insert(tk.END, "✓ Clique em 'Atualizar Produto Selecionado' para confirmar\n")
        self.text_info.insert(tk.END, "="*80 + "\n")
    
    def limpar_busca(self):
        """Limpa a busca."""
        self.entry_codprod.delete(0, tk.END)
        self.btn_atualizar.config(state='disabled')
        
        total = len(self.produtos_preenchidos)
        self.text_info.delete(1.0, tk.END)
        self.text_info.insert(1.0, f"Total de produtos que podem ser atualizados: {total}\n\n")
        self.text_info.insert(tk.END, "Digite um CODPROD para buscar especificamente\n")
        self.text_info.insert(tk.END, "ou clique em 'Atualizar TODOS' para atualizar todos de uma vez.")
    
    def atualizar_produto(self):
        """Atualiza um produto específico."""
        codprod = self.entry_codprod.get().strip()
        if not codprod:
            return
        
        try:
            codprod = int(codprod)
        except ValueError:
            return
        
        if codprod not in self.produtos_preenchidos:
            return
        
        produto = self.produtos_preenchidos[codprod]
        
        confirmacao = messagebox.askyesno(
            "Confirmar Atualização",
            f"Confirma a atualização do produto {codprod}?\n\n"
            f"{produto['DESCRICAO']}\n\n"
            "Os códigos auxiliares serão preenchidos no Oracle."
        )
        
        if not confirmacao:
            return
        
        self.executar_update([produto])
    
    def atualizar_todos(self):
        """Atualiza todos os produtos."""
        total = len(self.produtos_preenchidos)
        
        if total == 0:
            return
        
        confirmacao = messagebox.askyesno(
            "⚠️ CONFIRMAR ATUALIZAÇÃO EM MASSA",
            f"ATENÇÃO!\n\n"
            f"Você está prestes a atualizar {total} produtos de uma só vez!\n\n"
            f"Os códigos auxiliares de TODOS esses produtos serão preenchidos no Oracle.\n\n"
            f"TEM CERTEZA QUE DESEJA CONTINUAR?",
            icon='warning'
        )
        
        if not confirmacao:
            return
        
        confirmacao2 = messagebox.askyesno(
            "⚠️ ÚLTIMA CONFIRMAÇÃO",
            f"Esta é sua ÚLTIMA CHANCE!\n\n"
            f"Atualizar {total} produtos?\n\n"
            f"Continuar?",
            icon='warning'
        )
        
        if not confirmacao2:
            return
        
        self.executar_update(list(self.produtos_preenchidos.values()))
    
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
                    codaux = produto['CODAUXILIAR'] if produto['CODAUXILIAR'] else None
                    codaux2 = produto['CODAUXILIAR2'] if produto['CODAUXILIAR2'] else None
                    codauxtrib = produto['CODAUXILIARTRIB'] if produto['CODAUXILIARTRIB'] else None
                    
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
            
            # COMMIT
            conn.commit()
            
            messagebox.showinfo(
                "Atualização Concluída",
                f"✓ Atualização concluída com sucesso!\n\n"
                f"Produtos atualizados: {sucesso}\n"
                f"Erros: {erros}\n\n"
                f"As alterações foram confirmadas no banco (COMMIT)."
            )
            
            self.text_info.delete(1.0, tk.END)
            self.text_info.insert(1.0, f"✓ ATUALIZAÇÃO CONCLUÍDA\n\n")
            self.text_info.insert(tk.END, f"Produtos atualizados: {sucesso}\n")
            self.text_info.insert(tk.END, f"Erros: {erros}\n\n")
            self.text_info.insert(tk.END, f"COMMIT realizado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar atualização:\n{e}")
            conn.rollback()
        
        finally:
            cursor.close()
            conn.close()


def main():
    # Tela de login
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











# portfolio-commit-ready: atualizar_codauxiliares_oracle_v2
