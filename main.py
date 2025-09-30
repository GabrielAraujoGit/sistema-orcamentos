import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import unicodedata
from decimal import Decimal
import csv
import openpyxl
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

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
                FOREIGN KEY (cliente_id) REFERENCES clientes (id)
            )
        ''')
        # itens_pedido
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS itens_pedido (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id INTEGER NOT NULL,
                produto_id INTEGER NOT NULL,
                quantidade INTEGER NOT NULL,
                valor_unitario REAL NOT NULL,
                valor_total REAL NOT NULL,
                FOREIGN KEY (pedido_id) REFERENCES pedidos (id),
                FOREIGN KEY (produto_id) REFERENCES produtos (id)
            )
        ''')
        self.conn.commit()
    
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

        self.tree_clientes.bind("<Double-1>", self.editar_cliente)
        
        self.carregar_clientes()
    
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
        ttk.Label(header_frame, text="Nº do Pedido:").grid(row=0, column=2, sticky='w', padx=5)
        self.entry_numero = ttk.Entry(header_frame, width=20)
        self.entry_numero.grid(row=0, column=3, padx=5)
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
                self.tree_pedido_items.insert('', 'end', values=(f"{codigo} - {desc}", qtd, f"R$ {valor:.2f}", f"R$ {total:.2f}"))
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
        total = subtotal + total_impostos

        self.label_subtotal.config(text=f"Subtotal: R$ {subtotal:.2f}")
        self.label_impostos.config(text=f"Impostos: R$ {total_impostos:.2f}")
        self.label_total.config(text=f"TOTAL: R$ {total:.2f}")

    
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
            numero_pedido = self.entry_numero.get() or f"ORC-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            data_pedido = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            subtotal = sum(item['qtd'] * item['valor'] for item in self.itens_pedido_temp)
            impostos = subtotal * 0.20
            total = subtotal + impostos
            representante = self.entry_representante.get()
            self.cursor.execute('''
                INSERT INTO pedidos (numero_pedido, data_pedido, cliente_id, valor_produtos, valor_icms, valor_total, representante)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (numero_pedido, data_pedido, cliente_id, subtotal, impostos, total, representante))
            pedido_id = self.cursor.lastrowid
            for item in self.itens_pedido_temp:
                self.cursor.execute('''
                    INSERT INTO itens_pedido (pedido_id, produto_id, quantidade, valor_unitario, valor_total)
                    VALUES (?, ?, ?, ?, ?)
                ''', (pedido_id, item['produto_id'], item['qtd'], item['valor'], item['qtd'] * item['valor']))
            self.conn.commit()
            messagebox.showinfo("Sucesso", f"Orçamento {numero_pedido} salvo com sucesso!")
            self.limpar_pedido()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao finalizar orçamento: {e}")
    
    # ------------------- Export PDF -------------------
    def gerar_pdf_orcamento(self):
        if not self.combo_cliente.get() or not self.itens_pedido_temp:
            messagebox.showwarning("Atenção", "Selecione um cliente e adicione itens para gerar PDF!")
            return
        # cliente
        cliente_id = int(self.combo_cliente.get().split(' - ')[0])
        self.cursor.execute('SELECT razao_social, cnpj, cidade, endereco, telefone, email FROM clientes WHERE id = ?', (cliente_id,))
        cliente = self.cursor.fetchone()
        
        dados_orc = {
            'data': self.entry_data_orc.get(),
            'numero': self.entry_numero.get() or f"ORC-{datetime.now().strftime('%Y%m%d%H%M%S')}",
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

        # dados do cartão
        card_dados = {k: v.get() for k, v in self.card_entries.items()}
        tem_cartao = any(card_dados.values())

        # ask save
        cliente_nome = unicodedata.normalize('NFKD', dados_orc['cliente']['razao_social']).encode('ASCII', 'ignore').decode('ASCII')
        cliente_nome = cliente_nome.replace(" ", "_")
        data_hoje = datetime.now().strftime('%Y-%m-%d')
        nome_padrao = f"{cliente_nome}_{data_hoje}.pdf"

        caminho = filedialog.asksaveasfilename(
    defaultextension='.pdf',
    filetypes=[('PDF','*.pdf')],
    title='Salvar PDF',
    initialfile=nome_padrao
)
        if not caminho:
            return

        try:
            doc = SimpleDocTemplate(caminho, pagesize=A4)
            elementos = []
            estilos = getSampleStyleSheet()

            # título
            elementos.append(Paragraph("ORÇAMENTO DE PRODUTOS E SERVIÇOS", estilos['Title']))
            elementos.append(Spacer(1, 12))

            # cabeçalho: cliente + data
            cabecalho = [
                ['Cliente:', dados_orc['cliente']['razao_social'], 'Data:', dados_orc['data']],
                ['Endereço:', dados_orc['cliente']['endereco'], 'Cidade:', dados_orc['cliente']['cidade']],
                ['Telefone:', dados_orc['cliente']['telefone'], 'Email:', dados_orc['cliente']['email']]
            ]
            t_cab = Table(cabecalho, colWidths=[70, 200, 70, 200])
            t_cab.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica')
            ]))
            elementos.append(t_cab)
            elementos.append(Spacer(1, 12))

            # tabela de produtos
            tabela = [['Quantidade', 'Descrição', 'Valor Unitário', 'Valor Total']]
            subtotal = 0
            for p in dados_orc['produtos']:
                qtd = p['qtd']
                desc = p['descricao']
                unit = float(p['valor'])
                total = qtd * unit
                subtotal += total
                tabela.append([str(qtd), desc, f"R$ {unit:.2f}", f"R$ {total:.2f}"])

            t_prod = Table(tabela, colWidths=[80, 250, 100, 100])
            t_prod.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')
            ]))
            elementos.append(t_prod)
            elementos.append(Spacer(1, 12))

            # totais
            impostos = subtotal * 0.20  # pode trocar depois para cálculo detalhado
            total = subtotal + impostos
            totais = [['Total', f"R$ {total:.2f}"]]
            t_tot = Table(totais, colWidths=[430, 100])
            t_tot.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.7, colors.black),
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#BDD7EE")),
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold')
            ]))
            elementos.append(t_tot)
            elementos.append(Spacer(1, 12))

            # formas de pagamento (somente se houver cartão preenchido)
            if tem_cartao:
                elementos.append(Paragraph("Formas de Pagamento e Condições", estilos['Heading3']))
                formas = [[""]]
                t_cond = Table(formas, colWidths=[530])
                t_cond.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('FONTNAME', (0,0), (-1,-1), 'Helvetica')
                ]))
                elementos.append(t_cond)

            # gerar
            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"PDF gerado em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF: {e}")
        


    def importar_dados(self, tipo="clientes"):
            caminho = filedialog.askopenfilename(
                filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos", "*.*")]
            )
            if not caminho:
                return

            try:
                # Detecta extensão e lê dados
                if caminho.endswith(".xlsx"):
                    wb = openpyxl.load_workbook(caminho)
                    ws = wb.active
                    cabecalho = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
                    linhas = [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]
                else:  # CSV
                    with open(caminho, mode='r', encoding='utf-8') as arquivo_csv:
                        leitor = csv.reader(arquivo_csv)
                        cabecalho = next(leitor)
                        linhas = list(leitor)

                # Colunas esperadas no arquivo
                colunas_para_extrair = [
                    'Cliente', 
                    'CNPJ', 
                    'Contato /',   # mapeado para telefone
                    'Cidade', 
                    'Região'       # mapeado para estado
                ]

                # Mapeia colunas do arquivo -> banco de dados
                mapeamento = {}
                for db_col, csv_col in zip(
                    ['razao_social', 'cnpj', 'telefone', 'cidade', 'estado'],
                    colunas_para_extrair
                ):
                    try:
                        mapeamento[db_col] = cabecalho.index(csv_col)
                    except ValueError:
                        messagebox.showerror("Erro de Coluna", f"A coluna '{csv_col}' não foi encontrada no arquivo.")
                        return

                # Insere dados
                clientes_importados = 0
                for linha in linhas:
                    try:
                        razao_social = linha[mapeamento['razao_social']] or ''
                        cnpj = linha[mapeamento['cnpj']] or ''
                        telefone = linha[mapeamento['telefone']] or ''
                        cidade = linha[mapeamento['cidade']] or ''
                        estado = linha[mapeamento['estado']] or ''
                        
                        if not razao_social or not cnpj:
                            continue  # pula registros inválidos
                        
                        self.cursor.execute('''
                            INSERT INTO clientes (razao_social, cnpj, ie, endereco, cidade, estado, cep, telefone, email)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (razao_social, cnpj, '', '', cidade, estado, '', telefone, ''))
                        clientes_importados += 1
                    except sqlite3.IntegrityError:
                        # ignora duplicados de CNPJ
                        continue

                self.conn.commit()
                messagebox.showinfo("Sucesso", f"{clientes_importados} clientes importados com sucesso!")
                self.carregar_clientes()
                self.carregar_combos_pedido()

            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao importar clientes: {e}")


    
    # ------------------- Export Excel ----------------    -
    def exportar_excel_orcamento(self):
        if not self.combo_cliente.get() or not self.itens_pedido_temp:
            messagebox.showwarning("Atenção", "Selecione um cliente e adicione itens para exportar Excel!")
            return
        
        cliente_id = int(self.combo_cliente.get().split(' - ')[0])
        self.cursor.execute('SELECT razao_social, cnpj, cidade, endereco, telefone, email FROM clientes WHERE id = ?', (cliente_id,))
        cliente = self.cursor.fetchone()
        
        dados_orc = {
            'data': self.entry_data_orc.get(),
            'numero': self.entry_numero.get() or f"ORC-{datetime.now().strftime('%Y%m%d%H%M%S')}",
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

        # sanitizar nome do cliente
        cliente_nome = unicodedata.normalize('NFKD', dados_orc['cliente']['razao_social']).encode('ASCII', 'ignore').decode('ASCII')
        cliente_nome = cliente_nome.replace(" ", "_")
        data_hoje = datetime.now().strftime('%Y-%m-%d')
        nome_padrao = f"{cliente_nome}_{data_hoje}.xlsx"

        caminho = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel','*.xlsx')],
            title='Salvar Excel',
            initialfile=nome_padrao
        )
        if not caminho:
            return

        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Orçamento"

            # estilos
            bold = Font(bold=True)
            center = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                top=Side(style="thin"), bottom=Side(style="thin"))
            header_fill = PatternFill("solid", fgColor="DDDDDD")
            blue_fill = PatternFill("solid", fgColor="BDD7EE")

            # título
            ws.merge_cells('A1:D1')
            ws['A1'] = "ORÇAMENTO DE PRODUTOS E SERVIÇOS"
            ws['A1'].font = Font(bold=True, size=14)
            ws['A1'].alignment = center

            # Data
            ws['E1'] = f"Data: {dados_orc['data']}"
            ws['E1'].alignment = Alignment(horizontal="right")

            ws.append([])

            # dados do cliente
            ws.append(['Cliente:', dados_orc['cliente']['razao_social']])
            ws.append(['Endereço:', dados_orc['cliente']['endereco']])
            ws.append(['Cidade:', dados_orc['cliente']['cidade']])
            ws.append(['Telefone:', dados_orc['cliente']['telefone']])
            ws.append(['Email:', dados_orc['cliente']['email']])
            ws.append([])

            # cabeçalho da tabela
            headers = ['Quantidade', 'Descrição', 'Valor Unitário', 'Valor Total']
            ws.append(headers)
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=ws.max_row, column=col)
                c.font = bold
                c.alignment = center
                c.fill = header_fill
                c.border = thin_border

            # dados dos produtos
            for p in dados_orc['produtos']:
                qtd = p['qtd']
                desc = p['descricao']
                unit = float(p['valor'])
                total = qtd * unit
                ws.append([qtd, desc, unit, total])
                for col in range(1, 5):
                    ws.cell(row=ws.max_row, column=col).border = thin_border

            # total
            subtotal = sum(p['qtd']*p['valor'] for p in dados_orc['produtos'])
            impostos = subtotal * 0.20  # 👉 pode ser substituído pelo cálculo detalhado
            total = subtotal + impostos

            ws.append([])
            ws.append(['', '', 'Total', total])
            for col in range(1, 5):
                c = ws.cell(row=ws.max_row, column=col)
                c.font = bold
                c.fill = blue_fill
                c.border = thin_border

            ws.append([])
            ws.append(['Formas de Pagamento e Condições'])
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=4)
            ws.cell(row=ws.max_row, column=1).font = bold

            ws.append(['À vista ou em até 3x sem juros no cartão - Orçamento válido por 10 dias'])
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=4)

            # autoajuste das colunas
            for i, col in enumerate(ws.columns, 1):
                max_length = 0
                col_letter = get_column_letter(i)
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max_length + 2

            wb.save(caminho)
            messagebox.showinfo("Sucesso", f"Excel salvo em:\n{caminho}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar Excel: {e}")
    
        def __del__(self):
            if hasattr(self, 'conn'):
                self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaPedidos(root)
    root.mainloop()
