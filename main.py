import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import unicodedata
from decimal import Decimal
import csv
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def formatar_moeda(valor):
    """
    Formata número/int/str/None para padrão brasileiro: R$ 1.234,56
    Aceita:
      - None -> R$ 0,00
      - '1.234,56' ou '1234.56' ou 'R$ 1.234,56'
      - int/float
    """
    try:
        if valor is None:
            v = 0.0
        elif isinstance(valor, str):
            s = valor.strip().replace('R$', '').replace(' ', '')
            if s == '':
                v = 0.0
            else:
                # Normaliza milhares e decimais
                if ',' in s and '.' in s:
                    s = s.replace('.', '').replace(',', '.')
                else:
                    s = s.replace('.', '').replace(',', '.')
                v = float(s)
        else:
            v = float(valor)
    except Exception:
        v = 0.0

    texto = f"{v:,.2f}"            # ex: "1234.56" -> "1,234.56"
    texto = texto.replace(",", "V").replace(".", ",").replace("V", ".")  # -> "1.234,56"
    return f"R$ {texto}"




def normalizar_chave(texto):
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                    if unicodedata.category(c) != 'Mn')
    texto = texto.lower().replace(':', '').replace('%', '').replace('(', '').replace(')', '').strip()
    texto = texto.replace(' ', '_')
    return texto

class SistemaPedidos:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Orçamentos")
        self.root.geometry("1200x700")
        
        # Inicializar banco de dados
        self.init_db()
        
        # Criar notebook (abas)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # criar estruturas temporárias
        self.itens_pedido_temp = []
        
        # Criar abas
        self.criar_aba_clientes()
        self.criar_aba_produtos()
        self.criar_aba_pedidos()
        self.criar_aba_consulta_orcamentos()

    
    def init_db(self):
        self.conn = sqlite3.connect('pedidos.db')
        self.cursor = self.conn.cursor()
        # clientes
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                razao_social TEXT NOT NULL,
                cnpj TEXT UNIQUE NOT NULL,
                ie TEXT,
                endereco TEXT,
                cidade TEXT,
                estado TEXT,
                cep TEXT,
                telefone TEXT,
                email TEXT
            )
        ''')
        # produtos
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE NOT NULL,
                descricao TEXT NOT NULL,
                voltagem TEXT,
                valor_unitario REAL NOT NULL,
                aliq_icms REAL DEFAULT 0,
                aliq_ipi REAL DEFAULT 0,
                aliq_pis REAL DEFAULT 0,
                aliq_cofins REAL DEFAULT 0
            )
        ''')
        # pedidos
        # pedidos
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS pedidos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero_pedido TEXT UNIQUE NOT NULL,
                data_pedido TEXT NOT NULL,
                cliente_id INTEGER NOT NULL,
                valor_produtos REAL,
                valor_icms REAL,
                valor_ipi REAL,
                valor_pis REAL,
                valor_cofins REAL,
                valor_total REAL,
                representante TEXT,
                condicoes_pagamento TEXT,
                desconto REAL DEFAULT 0,
                status TEXT DEFAULT 'Em Aberto',
                observacoes TEXT,
                validade TEXT,
                FOREIGN KEY (cliente_id) REFERENCES clientes (id)
            )
        ''')

        # itens_pedido
        # Itens do pedido (ligados ao número do orçamento)
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS pedido_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero_pedido TEXT NOT NULL,
                produto_id INTEGER NOT NULL,
                qtd REAL NOT NULL,
                valor_unitario REAL NOT NULL,
                FOREIGN KEY (numero_pedido) REFERENCES pedidos (numero_pedido),
                FOREIGN KEY (produto_id) REFERENCES produtos (id)
            )
        ''')

        self.conn.commit()
    
    def calcular_totais(self, itens):
        subtotal = sum(i['qtd'] * i['valor'] for i in itens)
        total_icms = total_ipi = total_pis = total_cofins = 0
        for item in itens:
            self.cursor.execute('SELECT aliq_icms, aliq_ipi, aliq_pis, aliq_cofins FROM produtos WHERE id=?', (item['produto_id'],))
            icms, ipi, pis, cofins = self.cursor.fetchone()
            base = item['qtd'] * item['valor']
            total_icms += base * (icms or 0) / 100
            total_ipi += base * (ipi or 0) / 100
            total_pis += base * (pis or 0) / 100
            total_cofins += base * (cofins or 0) / 100
        total = subtotal + total_icms + total_ipi + total_pis + total_cofins
        return subtotal, total_icms, total_ipi, total_pis, total_cofins, total

    # ------------------- Clientes -------------------
    
    
    def criar_aba_clientes(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Clientes")
        form_frame = ttk.LabelFrame(frame, text="Cadastro de Cliente", padding=10)
        form_frame.pack(fill='x', padx=10, pady=10)
        
        labels = ['Razão Social:', 'CNPJ:', 'IE:', 'Endereço:', 'Cidade:', 'Estado:', 'CEP:', 'Telefone:', 'Email:']
        self.cliente_entries = {}
        for i, label in enumerate(labels):
            ttk.Label(form_frame, text=label).grid(row=i//3, column=(i%3)*2, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(form_frame, width=28)
            entry.grid(row=i//3, column=(i%3)*2+1, padx=5, pady=5)
            chave = normalizar_chave(label)
            self.cliente_entries[chave] = entry

        search_frame = ttk.Frame(frame)
        search_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(search_frame, text="Buscar Cliente:").pack(side='left', padx=5)
        self.entry_busca_cliente = ttk.Entry(search_frame, width=40)
        self.entry_busca_cliente.pack(side='left', padx=5)
        self.entry_busca_cliente.bind("<KeyRelease>", lambda e: self.carregar_clientes(filtro=self.entry_busca_cliente.get()))

        # Permitir duplo clique para editar
       
        
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=3, column=0, columnspan=6, pady=10)
        ttk.Button(btn_frame, text="Salvar Cliente", command=self.salvar_cliente).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Limpar", command=self.limpar_cliente).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Importar Clientes", command=lambda: self.importar_dados("clientes")).pack(side='left', padx=5)

        
        list_frame = ttk.LabelFrame(frame, text="Clientes Cadastrados", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        cols = ('ID', 'Razão Social', 'CNPJ', 'Cidade', 'Telefone')
        self.tree_clientes = ttk.Treeview(list_frame, columns=cols, show='headings', height=8)
        for col in cols:
            self.tree_clientes.heading(col, text=col)
            self.tree_clientes.column(col, width=140)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree_clientes.yview)
        self.tree_clientes.configure(yscrollcommand=scrollbar.set)
        self.tree_clientes.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.tree_clientes.bind("<Double-1>", self.visualizar_cliente)


        
        self.carregar_clientes()
    def criar_aba_consulta_orcamentos(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Consultar Orçamentos")

        search_frame = ttk.LabelFrame(frame, text="Filtros de Busca", padding=10)
        search_frame.pack(fill="x", padx=10, pady=10)  # mantém pack só para o frame em si

        # Número
        ttk.Label(search_frame, text="Número:").pack(side="left", padx=5)
        self.entry_busca_orc = ttk.Entry(search_frame, width=15)
        self.entry_busca_orc.pack(side="left", padx=5)

        # Cliente
        ttk.Label(search_frame, text="Cliente:").pack(side="left", padx=5)
        self.entry_busca_cliente = ttk.Entry(search_frame, width=25)
        self.entry_busca_cliente.pack(side="left", padx=5)

        # Representante
        ttk.Label(search_frame, text="Representante:").pack(side="left", padx=5)
        self.entry_busca_repr = ttk.Entry(search_frame, width=20)
        self.entry_busca_repr.pack(side="left", padx=5)

        # Status
        ttk.Label(search_frame, text="Status:").pack(side="left", padx=5)
        self.combo_status = ttk.Combobox(search_frame, values=["", "Em Aberto", "Aprovado", "Cancelado"], width=15)
        self.combo_status.pack(side="left", padx=5)

        # Datas
        ttk.Label(search_frame, text="Data Inicial (dd/mm/aaaa):").pack(side="left", padx=5)
        self.entry_data_ini = ttk.Entry(search_frame, width=12)
        self.entry_data_ini.pack(side="left", padx=5)

        ttk.Label(search_frame, text="Data Final (dd/mm/aaaa):").pack(side="left", padx=5)
        self.entry_data_fim = ttk.Entry(search_frame, width=12)
        self.entry_data_fim.pack(side="left", padx=5)

        # Botão buscar
        ttk.Button(search_frame, text="Buscar", command=self.buscar_orcamento).pack(side="left", padx=5)

        # Treeview de resultados
        list_frame = ttk.LabelFrame(frame, text="Resultados", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("Número", "Data", "Cliente", "Total", "Representante", "Status")
        self.tree_orcamentos = ttk.Treeview(list_frame, columns=cols, show="headings", height=12)
        for col in cols:
            self.tree_orcamentos.heading(col, text=col)
            self.tree_orcamentos.column(col, width=140)
        self.tree_orcamentos.pack(fill="both", expand=True)

        self.tree_orcamentos.bind("<Double-1>", self.visualizar_orcamento)

        # 👉 já carrega todos os orçamentos ao abrir a aba
        self.buscar_orcamento()

    def buscar_orcamento(self):
        numero = self.entry_busca_orc.get().strip()
        cliente = self.entry_busca_cliente.get().strip()
        representante = self.entry_busca_repr.get().strip()
        status = self.combo_status.get().strip()
        data_ini = self.entry_data_ini.get().strip()
        data_fim = self.entry_data_fim.get().strip()

        # Limpa resultados anteriores
        for item in self.tree_orcamentos.get_children():
            self.tree_orcamentos.delete(item)

        # Monta query dinâmica
        query = '''
            SELECT p.numero_pedido, p.data_pedido, c.razao_social, p.valor_total, p.representante, p.status
            FROM pedidos p
            JOIN clientes c ON p.cliente_id = c.id
            WHERE 1=1
        '''
        params = []

        if numero:
            query += " AND p.numero_pedido LIKE ?"
            params.append(f"%{numero}%")
        if cliente:
            query += " AND c.razao_social LIKE ?"
            params.append(f"%{cliente}%")
        if representante:
            query += " AND p.representante LIKE ?"
            params.append(f"%{representante}%")
        if status:
            query += " AND p.status = ?"
            params.append(status)
        if data_ini:
            try:
                dt_ini = datetime.strptime(data_ini, "%d/%m/%Y").strftime("%Y-%m-%d 00:00:00")
                query += " AND p.data_pedido >= ?"
                params.append(dt_ini)
            except:
                messagebox.showwarning("Atenção", "Data inicial inválida! Use dd/mm/aaaa.")
        if data_fim:
            try:
                dt_fim = datetime.strptime(data_fim, "%d/%m/%Y").strftime("%Y-%m-%d 23:59:59")
                query += " AND p.data_pedido <= ?"
                params.append(dt_fim)
            except:
                messagebox.showwarning("Atenção", "Data final inválida! Use dd/mm/aaaa.")

        query += " ORDER BY p.id DESC"

        self.cursor.execute(query, tuple(params))
        rows = self.cursor.fetchall()

        if not rows:
            messagebox.showinfo("Resultado", "Nenhum orçamento encontrado!")
            return

        for row in rows:
            valores = list(row)
            valores[3] = formatar_moeda(valores[3])  # formatar total
            self.tree_orcamentos.insert("", "end", values=valores)




    def salvar_cliente(self):
            try:
                dados = {k: v.get() for k, v in self.cliente_entries.items()}
                if not dados['razao_social'] or not dados['cnpj']:
                    messagebox.showwarning("Atenção", "Razão Social e CNPJ são obrigatórios!")
                    return

                if hasattr(self, 'cliente_edicao_id') and self.cliente_edicao_id:
                    # Atualizar cliente existente
                    self.cursor.execute('''
                        UPDATE clientes
                        SET razao_social=?, cnpj=?, ie=?, endereco=?, cidade=?, estado=?, cep=?, telefone=?, email=?
                        WHERE id=?
                    ''', (dados['razao_social'], dados['cnpj'], dados.get('ie'), dados.get('endereco'),
                        dados.get('cidade'), dados.get('estado'), dados.get('cep'),
                        dados.get('telefone'), dados.get('email'), self.cliente_edicao_id))
                    self.conn.commit()
                    messagebox.showinfo("Sucesso", "Cliente atualizado com sucesso!")
                    self.cliente_edicao_id = None
                else:
                    # Inserir novo cliente
                    self.cursor.execute('''
                        INSERT INTO clientes (razao_social, cnpj, ie, endereco, cidade, estado, cep, telefone, email)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (dados['razao_social'], dados['cnpj'], dados.get('ie'), dados.get('endereco'),
                        dados.get('cidade'), dados.get('estado'), dados.get('cep'),
                        dados.get('telefone'), dados.get('email')))
                    self.conn.commit()
                    messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")

                self.limpar_cliente()
                self.carregar_clientes()
                self.carregar_combos_pedido()
            except sqlite3.IntegrityError:
                messagebox.showerror("Erro", "CNPJ já cadastrado!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar cliente: {e}")
    
    def limpar_cliente(self):
        for entry in self.cliente_entries.values():
            entry.delete(0, tk.END)
    
    def carregar_clientes(self, filtro=""):
        for item in self.tree_clientes.get_children():
            self.tree_clientes.delete(item)
        if filtro:
            self.cursor.execute('''
                SELECT id, razao_social, cnpj, cidade, telefone 
                FROM clientes 
                WHERE razao_social LIKE ? OR cnpj LIKE ?
            ''', (f'%{filtro}%', f'%{filtro}%'))
        else:
            self.cursor.execute('SELECT id, razao_social, cnpj, cidade, telefone FROM clientes')
        for row in self.cursor.fetchall():
            self.tree_clientes.insert('', 'end', values=row)

    def editar_cliente(self, event):
            item = self.tree_clientes.selection()
            if not item: return
            valores = self.tree_clientes.item(item[0], "values")
            cliente_id = valores[0]
            self.cursor.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
            cliente = self.cursor.fetchone()
            if cliente:
                keys = ['id','razao_social','cnpj','ie','endereco','cidade','estado','cep','telefone','email']
                for k, v in zip(keys, cliente):
                    if k in self.cliente_entries:
                        self.cliente_entries[k].delete(0, tk.END)
                        self.cliente_entries[k].insert(0, v or "")
                self.cliente_edicao_id = cliente_id

    # ------------------- Produtos -------------------
    def criar_aba_produtos(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Produtos")
        form_frame = ttk.LabelFrame(frame, text="Cadastro de Produto", padding=10)
        form_frame.pack(fill='x', padx=10, pady=10)
        
        labels = ['Código:', 'Descrição:', 'Voltagem:', 'Valor Unitário:', 'ICMS (%):', 'IPI (%):', 'PIS (%):', 'COFINS (%):']
        self.produto_entries = {}
        for i, label in enumerate(labels):
            ttk.Label(form_frame, text=label).grid(row=i//4, column=(i%4)*2, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(form_frame, width=20)
            entry.grid(row=i//4, column=(i%4)*2+1, padx=5, pady=5)
            chave = normalizar_chave(label)  # ex: codigo, descricao, valor_unitario, icms, ...
            self.produto_entries[chave] = entry
        
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=2, column=0, columnspan=8, pady=10)
        ttk.Button(btn_frame, text="Salvar Produto", command=self.salvar_produto).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Limpar", command=self.limpar_produto).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Importar Produtos", command=lambda: self.importar_dados("produtos")).pack(side='left', padx=5)
            
        
        list_frame = ttk.LabelFrame(frame, text="Produtos Cadastrados", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        cols = ('ID', 'Código', 'Descrição', 'Voltagem', 'Valor', 'ICMS%', 'IPI%')
        self.tree_produtos = ttk.Treeview(list_frame, columns=cols, show='headings', height=8)
        for col in cols:
            self.tree_produtos.heading(col, text=col)
            self.tree_produtos.column(col, width=140)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree_produtos.yview)
        self.tree_produtos.configure(yscrollcommand=scrollbar.set)
        self.tree_produtos.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.carregar_produtos()
    
    def salvar_produto(self):
        try:
            dados = {k: v.get() for k, v in self.produto_entries.items()}
            # validar chaves esperadas (usamos normalizar_chave)
            if not dados.get('codigo') or not dados.get('descricao') or not dados.get('valor_unitario'):
                messagebox.showwarning("Atenção", "Código, Descrição e Valor Unitário são obrigatórios!")
                return
            # converter valores
            valor_unit = float(dados.get('valor_unitario') or 0)
            aliq_icms = float(dados.get('icms') or 0)
            aliq_ipi = float(dados.get('ipi') or 0)
            aliq_pis = float(dados.get('pis') or 0)
            aliq_cofins = float(dados.get('cofins') or 0)
            self.cursor.execute('''
                INSERT INTO produtos (codigo, descricao, voltagem, valor_unitario, aliq_icms, aliq_ipi, aliq_pis, aliq_cofins)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (dados.get('codigo'), dados.get('descricao'), dados.get('voltagem'), valor_unit,
                  aliq_icms, aliq_ipi, aliq_pis, aliq_cofins))
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!")
            self.limpar_produto()
            self.carregar_produtos()
            self.carregar_combos_pedido()
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Código já cadastrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar produto: {e}")
    
    def limpar_produto(self):
        for entry in self.produto_entries.values():
            entry.delete(0, tk.END)
    
    def carregar_produtos(self):
        for item in self.tree_produtos.get_children():
            self.tree_produtos.delete(item)
        self.cursor.execute('SELECT id, codigo, descricao, voltagem, valor_unitario, aliq_icms, aliq_ipi FROM produtos')
        for row in self.cursor.fetchall():
            self.tree_produtos.insert('', 'end', values=row)
    def editar_produto(self, event):
        item = self.tree_produtos.selection()
        if not item:
            return
        valores = self.tree_produtos.item(item[0], "values")
        produto_id = valores[0]
        self.cursor.execute("SELECT * FROM produtos WHERE id=?", (produto_id,))
        produto = self.cursor.fetchone()
        if produto:
            keys = ['id','codigo','descricao','voltagem','valor_unitario','aliq_icms','aliq_ipi','aliq_pis','aliq_cofins']
            for k, v in zip(keys, produto):
                if k in self.produto_entries:
                    self.produto_entries[k].delete(0, tk.END)
                    self.produto_entries[k].insert(0, v or "")
            self.produto_edicao_id = produto_id
            self.btn_salvar_produto.config(text="Atualizar Produto")
        
    # ------------------- Pedidos/Orçamentos -------------------
    def criar_aba_pedidos(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Orçamentos")
        
        # Cabeçalho (Data, Número, Representante)
        header_frame = ttk.LabelFrame(frame, text="Cabeçalho do Orçamento", padding=6)
        header_frame.pack(fill='x', padx=10, pady=6)
        ttk.Label(header_frame, text="Data do Orçamento:").grid(row=0, column=0, sticky='w', padx=5)
        self.entry_data_orc = ttk.Entry(header_frame, width=15)
        self.entry_data_orc.grid(row=0, column=1, padx=5)
        self.entry_data_orc.insert(0, datetime.now().strftime('%d/%m/%Y'))
        ttk.Label(header_frame, text="Nº do Orçamento:").grid(row=0, column=2, sticky='w', padx=5)
        self.label_numero_orc = ttk.Label(header_frame, text="Será gerado ao salvar")
        self.label_numero_orc.grid(row=0, column=3, padx=5)
        ttk.Label(header_frame, text="Representante:").grid(row=0, column=4, sticky='w', padx=5)
        self.entry_representante = ttk.Entry(header_frame, width=25)
        self.entry_representante.grid(row=0, column=5, padx=5)
        
        # Cartao BNDES
        card_frame = ttk.LabelFrame(frame, text="Cartão BNDES (opcional)", padding=6)
        card_frame.pack(fill='x', padx=10, pady=6)
        labels_card = ['Nº DO CARTÃO:', 'VALIDADE:', 'BANDEIRA:', 'BANCO:', 'PARCELAS:', 'NOME DO CARTÃO:']
        self.card_entries = {}
        for i, l in enumerate(labels_card):
            ttk.Label(card_frame, text=l).grid(row=i//3, column=(i%3)*2, sticky='w', padx=5, pady=3)
            e = ttk.Entry(card_frame, width=30)
            e.grid(row=i//3, column=(i%3)*2+1, padx=5, pady=3)
            self.card_entries[normalizar_chave(l)] = e
                # Campos adicionais do orçamento
        extra_frame = ttk.LabelFrame(frame, text="Informações Comerciais", padding=6)
        extra_frame.pack(fill='x', padx=10, pady=6)

        ttk.Label(extra_frame, text="Condições de Pagamento:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.entry_cond_pag = ttk.Entry(extra_frame, width=30)
        self.entry_cond_pag.grid(row=0, column=1, padx=5)

        ttk.Label(extra_frame, text="Validade (dias):").grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.entry_validade = ttk.Entry(extra_frame, width=10)
        self.entry_validade.grid(row=0, column=3, padx=5)

        ttk.Label(extra_frame, text="Desconto (R$):").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.entry_desconto = ttk.Entry(extra_frame, width=15)
        self.entry_desconto.grid(row=1, column=1, padx=5)

        ttk.Label(extra_frame, text="Observações:").grid(row=2, column=0, sticky='nw', padx=5, pady=2)
        self.text_obs = tk.Text(extra_frame, width=80, height=3)
        self.text_obs.grid(row=2, column=1, columnspan=3, padx=5, pady=2)




        # seleção cliente / produto
        top_frame = ttk.LabelFrame(frame, text="Novo Item", padding=6)
        top_frame.pack(fill='x', padx=10, pady=6)
        ttk.Label(top_frame, text="Cliente:").grid(row=0, column=0, sticky='w', padx=5)
        self.combo_cliente = ttk.Combobox(top_frame, width=60)
        self.combo_cliente.grid(row=0, column=1, padx=5, columnspan=4)
        self.combo_cliente.bind("<Return>", self.filtrar_clientes)
        ttk.Label(top_frame, text="Produto:").grid(row=1, column=0, sticky='w', padx=5)
        self.combo_produto = ttk.Combobox(top_frame, width=60, state='normal')
        self.combo_produto.grid(row=1, column=1, padx=5, columnspan=3)
        ttk.Label(top_frame, text="Quantidade:").grid(row=1, column=4, sticky='w', padx=5)
        self.entry_qtd = ttk.Entry(top_frame, width=8)
        self.entry_qtd.grid(row=1, column=5, padx=5)
        ttk.Button(top_frame, text="Adicionar Item", command=self.adicionar_item_pedido).grid(row=1, column=6, padx=5)
        
        # itens
        items_frame = ttk.LabelFrame(frame, text="Itens do Orçamento", padding=6)
        items_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('Produto', 'Qtd', 'Valor Unit.', 'Total')
        self.tree_pedido_items = ttk.Treeview(items_frame, columns=cols, show='headings', height=9)
        for col in cols:
            self.tree_pedido_items.heading(col, text=col)
            self.tree_pedido_items.column(col, width=180)
        self.tree_pedido_items.pack(fill='both', expand=True)
        
        # totais + ações
        totals_frame = ttk.LabelFrame(frame, text="Totais & Ações", padding=6)
        totals_frame.pack(fill='x', padx=10, pady=6)
        self.label_subtotal = ttk.Label(totals_frame, text="Subtotal: R$ 0,00")
        self.label_subtotal.grid(row=0, column=0, padx=8)
        self.label_impostos = ttk.Label(totals_frame, text="Impostos: R$ 0,00")
        self.label_impostos.grid(row=0, column=1, padx=8)
        self.label_total = ttk.Label(totals_frame, text="TOTAL: R$ 0,00", font=('Arial', 11, 'bold'))
        self.label_total.grid(row=0, column=2, padx=8)
        
        ttk.Button(totals_frame, text="Finalizar (Salvar) Orçamento", command=self.finalizar_pedido).grid(row=0, column=3, padx=8)
        ttk.Button(totals_frame, text="Limpar Orçamento", command=self.limpar_pedido).grid(row=0, column=4, padx=8)
        ttk.Button(totals_frame, text="Gerar PDF", command=self.gerar_pdf_orcamento).grid(row=0, column=5, padx=8)
        ttk.Button(totals_frame, text="Exportar Excel", command=self.exportar_excel_orcamento).grid(row=0, column=6, padx=8)
        
        self.carregar_combos_pedido()
    
    def carregar_combos_pedido(self):
        self.cursor.execute('SELECT id, razao_social FROM clientes')
        clientes = [f"{row[0]} - {row[1]}" for row in self.cursor.fetchall()]
        self.combo_cliente['values'] = clientes     
        self.cursor.execute('SELECT id, codigo, descricao FROM produtos')
        produtos = [f"{row[0]} - {row[1]} - {row[2]}" for row in self.cursor.fetchall()]
        self.combo_produto['values'] = produtos
    def filtrar_clientes(self, event=None):
        """Filtra clientes conforme o usuário digita no combobox do orçamento."""
        texto = self.combo_cliente.get().lower()
        self.cursor.execute('SELECT id, razao_social FROM clientes')
        todos = [f"{row[0]} - {row[1]}" for row in self.cursor.fetchall()]

        if texto:
            filtrados = [c for c in todos if texto in c.lower()]
        else:
            filtrados = todos

        self.combo_cliente['values'] = filtrados
        # mantém o texto digitado
        self.combo_cliente.event_generate('<Down>')  # abre a lista automaticamente

    def adicionar_item_pedido(self):
        if not self.combo_produto.get() or not self.entry_qtd.get():
            messagebox.showwarning("Atenção", "Selecione um produto e informe a quantidade!")
            return
        try:
            produto_id = int(self.combo_produto.get().split(' - ')[0])
            qtd = int(self.entry_qtd.get())
            self.cursor.execute('SELECT codigo, descricao, valor_unitario FROM produtos WHERE id = ?', (produto_id,))
            produto = self.cursor.fetchone()
            if produto:
                codigo, desc, valor = produto
                total = valor * qtd
                self.tree_pedido_items.insert('', 'end', values=(f"{codigo} - {desc}", qtd, formatar_moeda(valor), formatar_moeda(total)))
                self.itens_pedido_temp.append({'produto_id': produto_id, 'codigo': codigo, 'descricao': desc, 'qtd': qtd, 'valor': float(valor)})
                self.atualizar_totais()
                self.entry_qtd.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar item: {e}")
    
    def atualizar_totais(self):
        subtotal = sum(item['qtd'] * item['valor'] for item in self.itens_pedido_temp)
        
        # calcula impostos de cada item
        total_icms = 0
        total_ipi = 0
        total_pis = 0
        total_cofins = 0

        for item in self.itens_pedido_temp:
            self.cursor.execute('''
                SELECT aliq_icms, aliq_ipi, aliq_pis, aliq_cofins
                FROM produtos WHERE id = ?
            ''', (item['produto_id'],))
            aliq_icms, aliq_ipi, aliq_pis, aliq_cofins = self.cursor.fetchone()
            
            base = item['qtd'] * item['valor']
            total_icms   += base * (aliq_icms or 0) / 100
            total_ipi    += base * (aliq_ipi or 0) / 100
            total_pis    += base * (aliq_pis or 0) / 100
            total_cofins += base * (aliq_cofins or 0) / 100

        total_impostos = total_icms + total_ipi + total_pis + total_cofins
        desconto = float(self.entry_desconto.get() or 0)
        total = subtotal + total_impostos - desconto

        self.label_subtotal.config(text=f"Subtotal: {formatar_moeda(subtotal)}")
        self.label_impostos.config(text=f"Impostos: {formatar_moeda(total_impostos)}")
        self.label_total.config(text=f"TOTAL: {formatar_moeda(total)}")
    
    def visualizar_cliente(self, event):
        item = self.tree_clientes.selection()
        if not item:
            return
        
        valores = self.tree_clientes.item(item[0], "values")
        cliente_id = valores[0]  # primeira coluna do Treeview é o ID

        # Buscar dados completos no banco
        self.cursor.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
        cliente = self.cursor.fetchone()

        if not cliente:
            messagebox.showerror("Erro", "Cliente não encontrado.")
            return

        # Campos da tabela clientes
        keys = ["ID", "Razão Social", "CNPJ", "IE", "Endereço", "Cidade", "Estado", "CEP", "Telefone", "Email"]

        # Criar nova janela
        top = tk.Toplevel(self.root)
        top.title(f"Cliente - {cliente[1]}")
        top.geometry("600x400")

        frame_info = ttk.LabelFrame(top, text="Dados do Cliente", padding=10)
        frame_info.pack(fill="both", expand=True, padx=10, pady=10)

        # Mostrar cada campo em label
        for i, (k, v) in enumerate(zip(keys, cliente)):
            tk.Label(frame_info, text=f"{k}:", font=("Arial", 10, "bold")).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            tk.Label(frame_info, text=v if v else "Não informado").grid(row=i, column=1, sticky="w", padx=5, pady=2)

        # Botões de ação
        frame_botoes = ttk.Frame(top)
        frame_botoes.pack(pady=10)

        def acao_editar():
            """Chama o método editar_cliente para preencher o formulário da aba."""
            self.tree_clientes.selection_set(item)   # garante que o cliente está selecionado
            self.editar_cliente(None)                # usa o método existente
            top.destroy()                            # fecha a janelinha

        ttk.Button(frame_botoes, text="Editar Cliente", command=acao_editar).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Fechar", command=top.destroy).pack(side="left", padx=5)


    
    
    
    def visualizar_orcamento(self, event):
        item = self.tree_orcamentos.selection()
        if not item:
            return
        
        valores = self.tree_orcamentos.item(item[0], "values")
        numero_pedido = valores[0]  # primeira coluna = Número do Orçamento
        
        # Nova janela
        top = tk.Toplevel(self.root)
        top.title(f"Orçamento {numero_pedido}")
        top.geometry("850x600")

        # Buscar dados principais do pedido
        self.cursor.execute('''
            SELECT p.data_pedido, c.razao_social, c.cnpj, c.endereco, c.cidade, c.estado,
                p.valor_total, p.representante, p.condicoes_pagamento, p.observacoes, p.validade
            FROM pedidos p
            JOIN clientes c ON p.cliente_id = c.id
            WHERE p.numero_pedido = ?
        ''', (numero_pedido,))
        pedido = self.cursor.fetchone()

        if not pedido:
            messagebox.showerror("Erro", "Orçamento não encontrado.")
            top.destroy()
            return

        data, cliente, cnpj, endereco, cidade, estado, total, representante, cond_pag, obs, validade = pedido

        # Cabeçalho
        frame_info = ttk.LabelFrame(top, text="Dados do Orçamento", padding=10)
        frame_info.pack(fill="x", padx=10, pady=10)
        partes_endereco = []
        if endereco:
            partes_endereco.append(endereco)
        if cidade:
            partes_endereco.append(cidade)
        if estado:
            partes_endereco.append(estado)

        texto_endereco = " - ".join(partes_endereco) if partes_endereco else "Não informado"
        tk.Label(frame_info, text=f"Data: {data}").grid(row=0, column=0, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Cliente: {cliente}").grid(row=1, column=0, sticky="w", padx=5)
        tk.Label(frame_info, text=f"CNPJ: {cnpj}").grid(row=1, column=1, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Endereço: {texto_endereco}").grid(row=2, column=0, columnspan=2, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Representante: {representante}").grid(row=3, column=0, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Condições de Pagamento: {cond_pag}").grid(row=4, column=0, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Validade: {validade} dias").grid(row=4, column=1, sticky="w", padx=5)
        tk.Label(frame_info, text=f"Observações: {obs}").grid(row=5, column=0, columnspan=2, sticky="w", padx=5)
        tk.Label(frame_info, text=f"TOTAL: {formatar_moeda(total)}", font=("Arial", 11, "bold")).grid(row=6, column=0, sticky="w", padx=5, pady=5)

        # Itens do orçamento
        frame_itens = ttk.LabelFrame(top, text="Itens do Orçamento", padding=10)
        frame_itens.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("Código", "Descrição", "Qtd", "Valor Unit.", "Total")
        tree_itens = ttk.Treeview(frame_itens, columns=cols, show="headings", height=10)
        for col in cols:
            tree_itens.heading(col, text=col)
            tree_itens.column(col, width=130)
        tree_itens.pack(fill="both", expand=True)

        # Carregar itens do pedido
        self.cursor.execute('''
            SELECT pr.codigo, pr.descricao, pi.qtd, pi.valor_unitario
            FROM pedido_itens pi
            JOIN produtos pr ON pi.produto_id = pr.id
            WHERE pi.numero_pedido = ?
        ''', (numero_pedido,))
        itens = self.cursor.fetchall()

        for cod, desc, qtd, valor_unit in itens:
            tree_itens.insert("", "end", values=(
                cod, desc, qtd, formatar_moeda(valor_unit), formatar_moeda(qtd * valor_unit)
            ))

        # Rodapé com botões
        frame_botoes = ttk.Frame(top)
        frame_botoes.pack(fill="x", pady=10)

        ttk.Button(frame_botoes, text="Gerar PDF", 
                command=lambda: self.gerar_pdf_orcamento(numero_pedido)).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Exportar Excel", 
                command=self.exportar_excel_orcamento).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Fechar", 
                command=top.destroy).pack(side="right", padx=5)

            

        
    def limpar_pedido(self):
        for item in self.tree_pedido_items.get_children():
            self.tree_pedido_items.delete(item)
            self.itens_pedido_temp = []
            self.atualizar_totais()
            self.combo_cliente.set('')
            self.combo_produto.set('')
            self.entry_qtd.delete(0, tk.END)
            # limpar header e cartão se desejar:
            # self.entry_numero.delete(0, tk.END)
            # self.entry_representante.delete(0, tk.END)
            # for e in self.card_entries.values(): e.delete(0, tk.END)
    
    def finalizar_pedido(self):
        if not self.combo_cliente.get() or not self.itens_pedido_temp:
            messagebox.showwarning("Atenção", "Selecione um cliente e adicione itens!")
            return
        try:
            cliente_id = int(self.combo_cliente.get().split(' - ')[0])
            self.cursor.execute("SELECT COUNT(*) FROM pedidos")
            total_registros = self.cursor.fetchone()[0] or 0
            numero_pedido = f"ORC-{total_registros+1:04d}"
            self.label_numero_orc.config(text=numero_pedido)

            data_pedido = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 🔹 Calcula todos os totais corretamente
            subtotal, total_icms, total_ipi, total_pis, total_cofins, total = self.calcular_totais(self.itens_pedido_temp)

            representante = self.entry_representante.get()
            cond_pag = self.entry_cond_pag.get()
            validade = self.entry_validade.get()
            desconto = float(self.entry_desconto.get() or 0)
            observacoes = self.text_obs.get("1.0", tk.END).strip()

            # aplica desconto ao total
            total -= desconto

            # 🔹 Agora grava todos os impostos no banco
            self.cursor.execute('''
                INSERT INTO pedidos 
                (numero_pedido, data_pedido, cliente_id, valor_produtos, valor_icms, valor_ipi, valor_pis, valor_cofins, valor_total, representante,
                condicoes_pagamento, desconto, status, observacoes, validade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (numero_pedido, data_pedido, cliente_id, subtotal, total_icms, total_ipi, total_pis, total_cofins, total,
                representante, cond_pag, desconto, "Em Aberto", observacoes, validade))

            # Inserir os itens do orçamento
            for item in self.itens_pedido_temp:
                self.cursor.execute('''
                    INSERT INTO pedido_itens (numero_pedido, produto_id, qtd, valor_unitario)
                    VALUES (?, ?, ?, ?)
                ''', (numero_pedido, item['produto_id'], item['qtd'], item['valor']))

            self.conn.commit()
            messagebox.showinfo("Sucesso", f"Orçamento {numero_pedido} salvo com sucesso!")

            # Limpar itens temporários
            self.itens_pedido_temp.clear()
            self.tree_pedido_items.delete(*self.tree_pedido_items.get_children())

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar orçamento: {e}")


    
    # ------------------- Export PDF -------------------
    def gerar_pdf_orcamento(self, numero_pedido=None):
        try:
            # Se o botão chamar sem argumento, pega do label
            if numero_pedido is None:
                numero_pedido = self.label_numero_orc.cget("text")



            # -------- Padronização de tabelas --------
            COR_CABECALHO = colors.HexColor("#1E3A8A")  # Azul escuro
            COR_TEXTO_CAB = colors.white
            COR_GRID = colors.HexColor("#D1D5DB")

            def estilo_tabela(tabela, header=True):
                estilos = [
                    ("GRID", (0, 0), (-1, -1), 0.5, COR_GRID),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),         
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                ]
                if header:
                    estilos += [
                        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO),
                        ("TEXTCOLOR", (0, 0), (-1, 0), COR_TEXTO_CAB),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ]
                tabela.setStyle(TableStyle(estilos))
                return tabela

            # Buscar dados do pedido principal
            self.cursor.execute('''
                SELECT p.data_pedido, c.razao_social, c.cnpj, c.endereco, c.cidade, c.estado,
                    p.valor_produtos, p.valor_icms, p.valor_ipi, p.valor_pis, p.valor_cofins, p.valor_total,
                    p.representante, p.condicoes_pagamento, p.desconto, p.status, p.observacoes, p.validade
                FROM pedidos p
                JOIN clientes c ON p.cliente_id = c.id
                WHERE p.numero_pedido = ?
            ''', (numero_pedido,))
            pedido = self.cursor.fetchone()
            if not pedido:
                messagebox.showerror("Erro", "Pedido não encontrado no banco.")
                return

            (data_pedido, razao_social, cnpj, endereco, cidade, estado,
            subtotal, total_icms, total_ipi, total_pis, total_cofins, total,
            representante, cond_pag, desconto, status, observacoes, validade) = pedido

            # Garantir que não existam valores None
            subtotal = subtotal or 0.0
            total_icms = total_icms or 0.0
            total_ipi = total_ipi or 0.0
            total_pis = total_pis or 0.0
            total_cofins = total_cofins or 0.0
            total = total or 0.0
            desconto = desconto or 0.0
            cond_pag = cond_pag or ""
            observacoes = observacoes or ""
            validade = validade or ""
            representante = representante or ""
            status = status or "Em Aberto"

            # Buscar itens do pedido
            self.cursor.execute('''
                SELECT pr.codigo, pr.descricao, pi.qtd, pi.valor_unitario
                FROM pedido_itens pi
                JOIN produtos pr ON pi.produto_id = pr.id
                WHERE pi.numero_pedido = ?
            ''', (numero_pedido,))
            produtos = self.cursor.fetchall()

            # Montar PDF
            nome_arquivo = f"orcamento_{numero_pedido}.pdf"
            doc = SimpleDocTemplate(nome_arquivo, pagesize=A4,
                                    leftMargin=30, rightMargin=30, topMargin=40, bottomMargin=40)
            estilos = getSampleStyleSheet()
            elementos = []

            # Logo da empresa
            try:
                from reportlab.platypus import Image
                logo = Image("logo.png", width=120, height=30)  # ajuste conforme a sua logo
                logo.hAlign = 'LEFT'
                elementos.append(logo)
            except:
                pass

            # Título e número do orçamento
            titulo = Paragraph("<para align='center'><font size=18><b>ORÇAMENTO</b></font></para>", estilos['Normal'])
            elementos.append(titulo)
            elementos.append(Spacer(1, 6))

            num_orc = Paragraph(f"<para align='center'><font size=12>Orçamento nº <b>{numero_pedido}</b></font></para>",
                                estilos['Normal'])
            elementos.append(num_orc)
            elementos.append(Spacer(1, 20))

            partes_endereco = []
            if endereco:
                partes_endereco.append(endereco)
            if cidade:
                partes_endereco.append(cidade)
            if estado:
                partes_endereco.append(estado)
            
            
            endereco_formatado = " - ".join(partes_endereco) if partes_endereco else "Não informado"    

            
            try:
                data_formatada = datetime.strptime(pedido[0], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%y %H:%M")
            except:
                data_formatada = pedido[0] 


            # Informações do cliente (tabela)
            info_cliente = [
                ["Data:", data_formatada],
                ["Cliente:", razao_social],
                ["CNPJ:", cnpj],
                ["Endereço:", endereco_formatado],
                ["Representante:", representante],
                ["Status:", status],
            ]
            t_info = Table(info_cliente, colWidths=[100, 380])
            estilo_tabela(t_info, header=False)
            elementos.append(t_info)
            elementos.append(Spacer(1, 20))

            # Tabela de produtos
            tabela = [["Código", "Descrição", "Qtd", "Valor Unit.", "Total"]]
            for codigo, descricao, qtd, valor_unit in produtos:
                total_item = (qtd or 0) * (valor_unit or 0)
                tabela.append([
                    codigo,
                    descricao,
                    str(qtd),
                    formatar_moeda(valor_unit or 0),
                    formatar_moeda(total_item)
                ])

            t = Table(tabela, colWidths=[70, 230, 50, 90, 90], repeatRows=1, hAlign="CENTER")
            estilo_tabela(t, header=True)
            elementos.append(t)
            elementos.append(Spacer(1, 20))

            # Totais
            totais = [
                ['Subtotal', formatar_moeda(subtotal)],
                ['ICMS', formatar_moeda(total_icms)],
                ['IPI', formatar_moeda(total_ipi)],
                ['PIS', formatar_moeda(total_pis)],
                ['COFINS', formatar_moeda(total_cofins)],
            ]
            if desconto:
                totais.append(['Desconto', f"- {formatar_moeda(desconto)}"])
            totais.append(['TOTAL GERAL', formatar_moeda(total - desconto)])
            GRID_COLOR = colors.HexColor("#D1D5DB")
                
            t_tot = Table(totais, colWidths=[400, 100], hAlign="RIGHT")
            estilo_tabela(t_tot, header=True)   

            elementos.append(t_tot)
            elementos.append(Spacer(1, 20))

            # Condições e observações
            if cond_pag:
                elementos.append(Paragraph(f"<b>Condições de Pagamento:</b> {cond_pag}", estilos['Normal']))
            if validade:
                elementos.append(Paragraph(f"<b>Validade:</b> {validade} dias", estilos['Normal']))
            if observacoes:
                elementos.append(Paragraph(f"<b>Observações:</b> {observacoes}", estilos['Normal']))
            elementos.append(Spacer(1, 20))

            # Rodapé fixo
            elementos.append(Paragraph("<font size=9>Orçamento válido conforme condições acima.</font>", estilos['Italic']))
            elementos.append(Paragraph("<font size=9>Valores sujeitos a alterações sem aviso prévio.</font>",
                                    estilos['Italic']))

            # Gerar arquivo
            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"PDF gerado: {nome_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF: {e}")


  
    def importar_dados(self, tipo="clientes"):
    
                caminho = filedialog.askopenfilename(
                    filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos", "*.*")]
                )
                if not caminho:
                    return

                try:
                    # --- ler arquivo (xlsx ou csv) ---
                    if caminho.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
                        wb = openpyxl.load_workbook(caminho, data_only=True)
                        ws = wb.active
                        rows = list(ws.iter_rows(values_only=True))
                        if not rows:
                            messagebox.showerror("Erro", "Arquivo vazio")
                            return
                        header = [ (str(c) if c is not None else "").strip() for c in rows[0] ]
                        data_rows = [ list(row) for row in rows[1:] ]
                    else:
                        # CSV: tenta descobrir delimitador e lê
                        with open(caminho, newline='', encoding='utf-8') as f:
                            sample = f.read(2048)
                            f.seek(0)
                            try:    
                                dialect = csv.Sniffer().sniff(sample, delimiters=";,")
                                delim = dialect.delimiter
                            except Exception:
                                delim = ';'
                            reader = csv.reader(f, delimiter=delim)
                            header = [ (str(c) if c is not None else "").strip() for c in next(reader) ]
                            data_rows = [row for row in reader]

                    # --- helpers ---
                    def norm(s):
                        return ''.join(c for c in unicodedata.normalize('NFD', s or '') if unicodedata.category(c) != 'Mn').lower().strip()

                    header_norm = [norm(h) for h in header]

                    def find_index(*patterns):
                        for pat in patterns:
                            for i, h in enumerate(header_norm):
                                if pat in h:
                                    return i
                        return None

                    # ---------- importar clientes ----------
                    if tipo == "clientes":
                        idx_razao = find_index('razao', 'cliente', 'empresa', 'nome')
                        idx_cnpj  = find_index('cnpj')
                        idx_tel   = find_index('telefone', 'contato', 'tel')
                        idx_cidade= find_index('cidade')
                        idx_estado= find_index('estado', 'regiao', 'região')

                        if idx_razao is None or idx_cnpj is None:
                            messagebox.showerror("Erro de Coluna", "Arquivo de clientes precisa conter pelo menos as colunas 'Cliente/Razão' e 'CNPJ'.")
                            return

                        importados = 0
                        for row in data_rows:
                            try:
                                razao = str(row[idx_razao]).strip() if idx_razao < len(row) and row[idx_razao] is not None else ''
                                cnpj  = str(row[idx_cnpj]).strip()  if idx_cnpj < len(row) and row[idx_cnpj] is not None else ''
                                telefone = str(row[idx_tel]).strip() if idx_tel is not None and idx_tel < len(row) and row[idx_tel] is not None else ''
                                cidade = str(row[idx_cidade]).strip() if idx_cidade is not None and idx_cidade < len(row) and row[idx_cidade] is not None else ''
                                estado = str(row[idx_estado]).strip() if idx_estado is not None and idx_estado < len(row) and row[idx_estado] is not None else ''

                                if not razao or not cnpj:
                                    continue

                                self.cursor.execute('''
                                    INSERT INTO clientes (razao_social, cnpj, ie, endereco, cidade, estado, cep, telefone, email)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (razao, cnpj, '', '', cidade, estado, '', telefone, ''))
                                importados += 1
                            except sqlite3.IntegrityError:
                                # duplicado: ignora
                                continue

                        self.conn.commit()
                        messagebox.showinfo("Sucesso", f"{importados} clientes importados com sucesso!")
                        self.carregar_clientes()
                        self.carregar_combos_pedido()
                        return

                    # ---------- importar produtos ----------
                    if tipo == "produtos":
                        idx_codigo = find_index('codigo', 'cod', 'produto_id', 'id')
                        idx_desc   = find_index('descricao', 'descri', 'desc', 'produto', 'nome')
                        idx_volt   = find_index('voltagem', 'voltage', 'voltag')
                        idx_valor  = find_index('valor_unitario', 'valor unitario', 'valor', 'preco', 'price', 'unitario')
                        idx_icms   = find_index('icms')
                        idx_ipi    = find_index('ipi')
                        idx_pis    = find_index('pis')
                        idx_cofins = find_index('cofins', 'cofins%')

                        if idx_codigo is None or idx_desc is None or idx_valor is None:
                            messagebox.showerror("Erro de Coluna", "Arquivo de produtos precisa conter ao menos 'codigo', 'descricao' e 'preco/valor'.")
                            return

                        def to_float_cell(r, idx):
                            if idx is None or idx >= len(r): return 0.0
                            v = r[idx]
                            if v is None: return 0.0
                            s = str(v).replace('R$', '').replace('%', '').replace(',', '.').strip()
                            try:
                                return float(s)
                            except:
                                return 0.0

                        importados = 0
                        for row in data_rows:
                            try:
                                codigo = str(row[idx_codigo]).strip() if idx_codigo < len(row) and row[idx_codigo] is not None else ''
                                descricao = str(row[idx_desc]).strip() if idx_desc < len(row) and row[idx_desc] is not None else ''
                                voltagem = str(row[idx_volt]).strip() if idx_volt is not None and idx_volt < len(row) and row[idx_volt] is not None else ''
                                valor_unitario = to_float_cell(row, idx_valor)
                                aliq_icms = to_float_cell(row, idx_icms)
                                aliq_ipi  = to_float_cell(row, idx_ipi)
                                aliq_pis  = to_float_cell(row, idx_pis)
                                aliq_cofins = to_float_cell(row, idx_cofins)

                                if not codigo or not descricao:
                                    continue

                                self.cursor.execute('''
                                    INSERT INTO produtos (codigo, descricao, voltagem, valor_unitario, aliq_icms, aliq_ipi, aliq_pis, aliq_cofins)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (codigo, descricao, voltagem, valor_unitario, aliq_icms, aliq_ipi, aliq_pis, aliq_cofins))
                                importados += 1
                            except sqlite3.IntegrityError:
                                # ignora duplicados
                                continue
                            except Exception:
                                # pula linha problemática
                                continue

                        self.conn.commit()
                        messagebox.showinfo("Sucesso", f"{importados} produtos importados com sucesso!")
                        self.carregar_produtos()
                        self.carregar_combos_pedido()
                        return

                    # tipo inválido
                    messagebox.showerror("Tipo inválido", f"Tipo de importação '{tipo}' não suportado.")
                    return

                except Exception as e:
                    messagebox.showerror("Erro", f"Falha ao importar: {e}")
    
    
    def exportar_excel_orcamento(self):
        if not self.combo_cliente.get() or not self.itens_pedido_temp:
            messagebox.showwarning("Atenção", "Selecione um cliente e adicione itens para exportar Excel!")
            return
        

        cliente_id = int(self.combo_cliente.get().split(' - ')[0])
        self.cursor.execute('SELECT razao_social, cnpj, cidade, endereco, telefone, email FROM clientes WHERE id = ?', (cliente_id,))
        cliente = self.cursor.fetchone()

        subtotal, total_icms, total_ipi, total_pis, total_cofins, total = self.calcular_totais(self.itens_pedido_temp)

        dados_orc = {
            'data': self.entry_data_orc.get(),
            'numero': self.label_numero_orc.cget("text") or f"ORC-{datetime.now().strftime('%Y%m%d%H%M%S')}",
            'representante': self.entry_representante.get(),
            'cliente': {
                'razao_social': cliente[0],
                'cnpj': cliente[1],
                'cidade': cliente[2],
                'endereco': cliente[3] or '',
                'telefone': cliente[4] or '',
                'email': cliente[5] or ''
            },
            'produtos': self.itens_pedido_temp
        }

        cliente_nome = unicodedata.normalize('NFKD', dados_orc['cliente']['razao_social']).encode('ASCII', 'ignore').decode('ASCII')
        cliente_nome = cliente_nome.replace(" ", "_")
        nome_padrao = f"Orcamento_{cliente_nome}_{datetime.now().strftime('%Y%m%d')}.xlsx"

        caminho = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')], title='Salvar Orçamento', initialfile=nome_padrao)
        if not caminho: return

        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from openpyxl.utils import get_column_letter
            from openpyxl.drawing.image import Image as XLImage

            wb = openpyxl.Workbook()

            # Paleta de cores profissional
            COR_PRIMARIA = "1E3A8A"      # Azul escuro corporativo
            COR_SECUNDARIA = "3B82F6"    # Azul médio
            COR_DESTAQUE = "10B981"      # Verde para totais
            COR_HEADER = "F3F4F6"        # Cinza claro
            COR_TEXTO_ESCURO = "1F2937"  # Cinza escuro

            # Estilos refinados
            thin = Side(border_style="thin", color="D1D5DB")
            border_light = Border(top=thin, left=thin, right=thin, bottom=thin)
            
            medium = Side(border_style="medium", color=COR_PRIMARIA)
            border_strong = Border(top=medium, left=medium, right=medium, bottom=medium)

            # ------------------- ABA PRINCIPAL - ORÇAMENTO -------------------
            ws = wb.active
            ws.title = "Orçamento"

            # Definir larguras das colunas
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 15

            # CABEÇALHO EMPRESA
            linha_atual = 1
            
            # Logo (se existir)
            try:
                img = XLImage("logo .png")
                img.width, img.height = 120, 50
                ws.add_image(img, "A1")
            except:
                ws.merge_cells(f"A{linha_atual}:B{linha_atual}")
                ws[f"A{linha_atual}"].value = "SUA EMPRESA"
                ws[f"A{linha_atual}"].font = Font(size=18, bold=True, color=COR_PRIMARIA)

            # Informações no cabeçalho
            ws.merge_cells(f"D{linha_atual}:F{linha_atual}")
            ws[f"D{linha_atual}"].value = "ORÇAMENTO"
            ws[f"D{linha_atual}"].font = Font(size=20, bold=True, color=COR_PRIMARIA)
            ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right", vertical="center")

            linha_atual += 1
            ws.merge_cells(f"D{linha_atual}:F{linha_atual}")
            ws[f"D{linha_atual}"].value = f"Nº {dados_orc['numero']}"
            ws[f"D{linha_atual}"].font = Font(size=11, color=COR_TEXTO_ESCURO)
            ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right")

            linha_atual += 1
            ws.merge_cells(f"D{linha_atual}:F{linha_atual}")
            ws[f"D{linha_atual}"].value = f"Data: {dados_orc['data']}"
            ws[f"D{linha_atual}"].font = Font(size=11, color=COR_TEXTO_ESCURO)
            ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right")

            linha_atual += 2

            # SEÇÃO CLIENTE
            ws.merge_cells(f"A{linha_atual}:F{linha_atual}")
            ws[f"A{linha_atual}"].value = "DADOS DO CLIENTE"
            ws[f"A{linha_atual}"].font = Font(size=12, bold=True, color="FFFFFF")
            ws[f"A{linha_atual}"].fill = PatternFill("solid", fgColor=COR_PRIMARIA)
            ws[f"A{linha_atual}"].alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[linha_atual].height = 25

            linha_atual += 1
            
            # Grid de informações do cliente
            info_cliente = [
                ["Razão Social:", dados_orc['cliente']['razao_social'], "CNPJ:", dados_orc['cliente']['cnpj']],
                ["Endereço:", dados_orc['cliente']['endereco'], "Cidade:", dados_orc['cliente']['cidade']],
                ["Telefone:", dados_orc['cliente']['telefone'], "E-mail:", dados_orc['cliente']['email']],
                ["Representante:", dados_orc['representante'], "", ""]
            ]

            for info_row in info_cliente:
                ws[f"A{linha_atual}"].value = info_row[0]
                ws[f"A{linha_atual}"].font = Font(bold=True, size=10, color=COR_TEXTO_ESCURO)
                ws[f"A{linha_atual}"].fill = PatternFill("solid", fgColor=COR_HEADER)
                
                ws.merge_cells(f"B{linha_atual}:C{linha_atual}")
                ws[f"B{linha_atual}"].value = info_row[1]
                ws[f"B{linha_atual}"].font = Font(size=10)
                
                if info_row[2]:
                    ws[f"D{linha_atual}"].value = info_row[2]
                    ws[f"D{linha_atual}"].font = Font(bold=True, size=10, color=COR_TEXTO_ESCURO)
                    ws[f"D{linha_atual}"].fill = PatternFill("solid", fgColor=COR_HEADER)
                    
                    ws.merge_cells(f"E{linha_atual}:F{linha_atual}")
                    ws[f"E{linha_atual}"].value = info_row[3]
                    ws[f"E{linha_atual}"].font = Font(size=10)
                
                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws[f"{col}{linha_atual}"].border = border_light
                
                linha_atual += 1

            linha_atual += 1

            # SEÇÃO PRODUTOS
            ws.merge_cells(f"A{linha_atual}:F{linha_atual}")
            ws[f"A{linha_atual}"].value = "ITENS DO ORÇAMENTO"
            ws[f"A{linha_atual}"].font = Font(size=12, bold=True, color="FFFFFF")
            ws[f"A{linha_atual}"].fill = PatternFill("solid", fgColor=COR_PRIMARIA)
            ws[f"A{linha_atual}"].alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[linha_atual].height = 25

            linha_atual += 1
            linha_header = linha_atual

            # Cabeçalho da tabela de produtos
            headers = ['Código', 'Descrição', 'Quantidade', 'Valor Unit.', 'Subtotal', 'Total Item']
            for idx, header in enumerate(headers, start=1):
                col_letter = get_column_letter(idx)
                ws[f"{col_letter}{linha_atual}"].value = header
                ws[f"{col_letter}{linha_atual}"].font = Font(bold=True, size=10, color=COR_TEXTO_ESCURO)
                ws[f"{col_letter}{linha_atual}"].fill = PatternFill("solid", fgColor=COR_HEADER)
                ws[f"{col_letter}{linha_atual}"].alignment = Alignment(horizontal="center", vertical="center")
                ws[f"{col_letter}{linha_atual}"].border = border_light
            
            ws.row_dimensions[linha_atual].height = 22

            linha_atual += 1
            linha_inicio_dados = linha_atual

            # Dados dos produtos
            for idx, p in enumerate(dados_orc['produtos'], start=1):
                total_item = p['qtd'] * p['valor']
                
                ws[f"A{linha_atual}"].value = p['codigo']
                ws[f"A{linha_atual}"].alignment = Alignment(horizontal="center")
                
                ws[f"B{linha_atual}"].value = p['descricao']
                ws[f"B{linha_atual}"].alignment = Alignment(horizontal="left")
                
                ws[f"C{linha_atual}"].value = p['qtd']
                ws[f"C{linha_atual}"].alignment = Alignment(horizontal="center")
                ws[f"C{linha_atual}"].number_format = '#,##0'
                
                ws[f"D{linha_atual}"].value = p['valor']
                ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right")
                ws[f"D{linha_atual}"].number_format = 'R$ #,##0.00'
                
                ws[f"E{linha_atual}"].value = total_item
                ws[f"E{linha_atual}"].alignment = Alignment(horizontal="right")
                ws[f"E{linha_atual}"].number_format = 'R$ #,##0.00'
                
                ws[f"F{linha_atual}"].value = total_item
                ws[f"F{linha_atual}"].alignment = Alignment(horizontal="right")
                ws[f"F{linha_atual}"].number_format = 'R$ #,##0.00'
                ws[f"F{linha_atual}"].font = Font(bold=True)

                # Bordas e zebrado
                fill_color = "FFFFFF" if idx % 2 == 0 else "F9FAFB"
                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws[f"{col}{linha_atual}"].border = border_light
                    ws[f"{col}{linha_atual}"].fill = PatternFill("solid", fgColor=fill_color)

                linha_atual += 1

            linha_atual += 1

            # RESUMO FINANCEIRO
            ws.merge_cells(f"A{linha_atual}:C{linha_atual}")
            ws[f"A{linha_atual}"].value = "RESUMO FINANCEIRO"
            ws[f"A{linha_atual}"].font = Font(size=11, bold=True, color=COR_PRIMARIA)
            
            linha_atual += 1

            resumo_itens = [
                ("Subtotal:", subtotal, False),
                ("ICMS:", total_icms, False),
                ("IPI:", total_ipi, False),
                ("PIS:", total_pis, False),
                ("COFINS:", total_cofins, False),
            ]

            for label, valor, _ in resumo_itens:
                ws.merge_cells(f"D{linha_atual}:E{linha_atual}")
                ws[f"D{linha_atual}"].value = label
                ws[f"D{linha_atual}"].font = Font(size=10, color=COR_TEXTO_ESCURO)
                ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right")
                
                ws[f"F{linha_atual}"].value = valor
                ws[f"F{linha_atual}"].number_format = 'R$ #,##0.00'
                ws[f"F{linha_atual}"].alignment = Alignment(horizontal="right")
                ws[f"F{linha_atual}"].border = border_light
                
                linha_atual += 1

            # TOTAL GERAL - Destaque especial
            ws.merge_cells(f"D{linha_atual}:E{linha_atual}")
            ws[f"D{linha_atual}"].value = "TOTAL GERAL"
            ws[f"D{linha_atual}"].font = Font(size=13, bold=True, color="FFFFFF")
            ws[f"D{linha_atual}"].fill = PatternFill("solid", fgColor=COR_DESTAQUE)
            ws[f"D{linha_atual}"].alignment = Alignment(horizontal="right", vertical="center")
            ws[f"D{linha_atual}"].border = border_strong
            
            ws[f"F{linha_atual}"].value = total
            ws[f"F{linha_atual}"].font = Font(size=13, bold=True, color="FFFFFF")
            ws[f"F{linha_atual}"].fill = PatternFill("solid", fgColor=COR_DESTAQUE)
            ws[f"F{linha_atual}"].number_format = 'R$ #,##0.00'
            ws[f"F{linha_atual}"].alignment = Alignment(horizontal="right", vertical="center")
            ws[f"F{linha_atual}"].border = border_strong
            ws.row_dimensions[linha_atual].height = 28

            linha_atual += 3

            # CONDIÇÕES COMERCIAIS
            ws.merge_cells(f"A{linha_atual}:F{linha_atual}")
            ws[f"A{linha_atual}"].value = "CONDIÇÕES COMERCIAIS"
            ws[f"A{linha_atual}"].font = Font(size=11, bold=True, color=COR_PRIMARIA)
            
            linha_atual += 1
            
            condicoes = [
                "• Validade do orçamento: 10 dias corridos",
                "• Forma de pagamento: A combinar",
                "• Prazo de entrega: Conforme disponibilidade",
                "• Valores sujeitos a alteração sem aviso prévio",
                "• Frete não incluso"
            ]

            for condicao in condicoes:
                ws.merge_cells(f"A{linha_atual}:F{linha_atual}")
                ws[f"A{linha_atual}"].value = condicao
                ws[f"A{linha_atual}"].font = Font(size=9, color=COR_TEXTO_ESCURO)
                ws[f"A{linha_atual}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                linha_atual += 1

            linha_atual += 2

            # RODAPÉ
            ws.merge_cells(f"A{linha_atual}:F{linha_atual}")
            ws[f"A{linha_atual}"].value = "Orçamento gerado automaticamente pelo sistema"
            ws[f"A{linha_atual}"].font = Font(size=8, italic=True, color="9CA3AF")
            ws[f"A{linha_atual}"].alignment = Alignment(horizontal="center")

            # Filtros automáticos na tabela de produtos
            ws.auto_filter.ref = f"A{linha_header}:F{linha_inicio_dados + len(dados_orc['produtos']) - 1}"

            # ------------------- ABA DETALHAMENTO -------------------
            ws_det = wb.create_sheet("Detalhamento")
            ws_det.column_dimensions['A'].width = 20
            ws_det.column_dimensions['B'].width = 50

            linha = 1
            ws_det[f"A{linha}"].value = "DETALHAMENTO DO ORÇAMENTO"
            ws_det[f"A{linha}"].font = Font(size=14, bold=True, color=COR_PRIMARIA)
            ws_det.merge_cells(f"A{linha}:B{linha}")
            
            linha += 2

            detalhes = [
                ["Número do Orçamento", dados_orc['numero']],
                ["Data de Emissão", dados_orc['data']],
                ["Cliente", dados_orc['cliente']['razao_social']],
                ["CNPJ", dados_orc['cliente']['cnpj']],
                ["Representante", dados_orc['representante']],
                ["Quantidade de Itens", len(dados_orc['produtos'])],
                ["Subtotal", f"R$ {subtotal:,.2f}"],
                ["Impostos (ICMS+IPI+PIS+COFINS)", f"R$ {(total_icms + total_ipi + total_pis + total_cofins):,.2f}"],
                ["Valor Total", f"R$ {total:,.2f}"]
            ]

            for detalhe in detalhes:
                ws_det[f"A{linha}"].value = detalhe[0]
                ws_det[f"A{linha}"].font = Font(bold=True, color=COR_TEXTO_ESCURO)
                ws_det[f"A{linha}"].fill = PatternFill("solid", fgColor=COR_HEADER)
                ws_det[f"A{linha}"].border = border_light
                
                ws_det[f"B{linha}"].value = detalhe[1]
                ws_det[f"B{linha}"].border = border_light
                
                linha += 1

            # Salvar
            wb.save(caminho)
            messagebox.showinfo("Sucesso", f"Orçamento exportado com sucesso!\n\n{caminho}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar orçamento:\n{str(e)}")


    
    def __del__(self):
            if hasattr(self, 'conn'):
                self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaPedidos(root)
    root.mainloop()
