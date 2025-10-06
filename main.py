import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import unicodedata
import re
from decimal import Decimal
import csv
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import ttkbootstrap as tb
from ttkbootstrap.constants import * 
import os

def formatar_moeda(valor):
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
        self.edicao_numero_pedido = None

         # Atalhos de teclado
        self.root.bind("<Control-n>", lambda e: self.limpar_pedido())      # Novo Orçamento
        self.root.bind("<Control-s>", lambda e: self.finalizar_pedido())   # Salvar Orçamento
         
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
        # produtos
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE NOT NULL,
                descricao TEXT NOT NULL,
                valor_unitario REAL NOT NULL,
                tipo TEXT,
                origem_tributacao TEXT,
                voltagem TEXT,
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

    def criar_aba_clientes(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Clientes")

        # --- Barra de pesquisa ---
        search_frame = ttk.Frame(frame)
        search_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(search_frame, text="Pesquisar:").pack(side='left', padx=5)
        self.entry_pesquisa_cliente = ttk.Entry(search_frame, width=40)
        self.entry_pesquisa_cliente.pack(side='left', padx=5)
        
        ttk.Button(search_frame, text="Buscar", command=lambda: self.carregar_clientes(self.entry_pesquisa_cliente.get())).pack(side='left', padx=5)
        ttk.Button(search_frame, text="Limpar", command=lambda: self.carregar_clientes()).pack(side='left', padx=5)

        # --- Barra de botões ---
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(btn_frame, text="Adicionar", bootstyle=SUCCESS, command=self.adicionar_cliente).pack(side='left', padx=5 )
        ttk.Button(btn_frame, text="Editar", bootstyle=INFO, command=lambda: self.editar_cliente(None)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Excluir", bootstyle=DANGER, command=self.excluir_cliente).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Importar Arquivo", bootstyle=WARNING, command=lambda: self.importar_dados("clientes")).pack(side='left', padx=5)

        # --- Lista de clientes ---
        list_frame = ttk.LabelFrame(frame, text="Clientes Cadastrados", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)

        cols = ('ID', 'Razão Social', 'CNPJ', 'Cidade', 'Telefone')
        self.tree_clientes = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        for col in cols:
            self.tree_clientes.heading(col, text=col)
            self.tree_clientes.column(col, width=140)

        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree_clientes.yview)
        self.tree_clientes.configure(yscrollcommand=scrollbar.set)
        self.tree_clientes.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Duplo clique abre visualização
        self.tree_clientes.bind("<Double-1>", self.visualizar_cliente)

        self.carregar_clientes()

    def abrir_formulario_cliente(self, cliente=None):
        """Abre popup para adicionar/editar cliente."""
        top = tk.Toplevel(self.root)
        top.title("Cadastro de Cliente")
        top.geometry("600x400")

        # deixa responsivo
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(3, weight=1)

        labels = ['Razão Social:', 'CNPJ:', 'IE:', 'Endereço:', 'Cidade:', 'Estado:', 'CEP:', 'Telefone:', 'Email:']
        entries = {}
        for i, label in enumerate(labels):
            ttk.Label(top, text=label).grid(row=i//2, column=(i%2)*2, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(top, width=28)
            entry.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=5, sticky="ew")
            chave = normalizar_chave(label)
            entries[chave] = entry

        # Se for edição, preencher os campos
        if cliente:
            keys = ['id','razao_social','cnpj','ie','endereco','cidade','estado','cep','telefone','email']
            for k, v in zip(keys, cliente):
                if k in entries:
                    entries[k].insert(0, v or "")
        
        def salvar():
            dados = {k: v.get().strip() for k, v in entries.items()}
            campos_obrigatorios = {"razao_social": "Razão Social", "cnpj": "CNPJ"}
            faltando = [nome for chave, nome in campos_obrigatorios.items() if not dados.get(chave)]

            if faltando:
                messagebox.showwarning("Atenção", f"Preencha os campos obrigatórios: {', '.join(faltando)}")
                return

            try:
                if cliente:  # editar
                    self.cursor.execute('''
                        UPDATE clientes
                        SET razao_social=?, cnpj=?, ie=?, endereco=?, cidade=?, estado=?, cep=?, telefone=?, email=?
                        WHERE id=?
                    ''', (dados['razao_social'], dados['cnpj'], dados.get('ie'), dados.get('endereco'),
                        dados.get('cidade'), dados.get('estado'), dados.get('cep'),
                        dados.get('telefone'), dados.get('email'), cliente[0]))
                else:  # novo
                    self.cursor.execute('''
                        INSERT INTO clientes (razao_social, cnpj, ie, endereco, cidade, estado, cep, telefone, email)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (dados['razao_social'], dados['cnpj'], dados.get('ie'), dados.get('endereco'),
                        dados.get('cidade'), dados.get('estado'), dados.get('cep'),
                        dados.get('telefone'), dados.get('email')))
                self.conn.commit()
                self.carregar_clientes()
                top.destroy()
            except sqlite3.IntegrityError:
                messagebox.showerror("Erro", "CNPJ já cadastrado!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar cliente: {e}")

        ttk.Button(top, text="Salvar", command=salvar).grid(row=6, column=0, columnspan=2, pady=10)

    def adicionar_cliente(self):
        self.abrir_formulario_cliente()

    def editar_cliente(self, event=None):
        item = self.tree_clientes.selection()
        if not item:
            messagebox.showwarning("Atenção", "Selecione um cliente para editar.")
            return
        valores = self.tree_clientes.item(item[0], "values")
        cliente_id = valores[0]
        self.cursor.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
        cliente = self.cursor.fetchone()
        if cliente:
            self.abrir_formulario_cliente(cliente)


    def excluir_cliente(self):
        selecionados = self.tree_clientes.selection()
        if not selecionados:
            messagebox.showwarning("Atenção", "Selecione um ou mais clientes para excluir.")
            return

        if not messagebox.askyesno("Confirmação", f"Deseja realmente excluir {len(selecionados)} cliente(s)?"):
            return

        try:
            for item in selecionados:
                valores = self.tree_clientes.item(item, "values")
                cliente_id = valores[0]
                self.cursor.execute("DELETE FROM clientes WHERE id=?", (cliente_id,))
            
            self.conn.commit()
            self.carregar_clientes()
            messagebox.showinfo("Sucesso", f"{len(selecionados)} cliente(s) excluído(s) com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir: {e}")

   
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
        self.combo_status = ttk.Combobox(search_frame, values=["", "Em Aberto", "Aprovado", "Cancelado", "Rejeitado"], width=15)
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

        # Configuração de cores só no texto do Status
        self.tree_orcamentos.tag_configure("Em Aberto", foreground="#1d4ed8")   # azul
        self.tree_orcamentos.tag_configure("Aprovado", foreground="#15803d")    # verde
        self.tree_orcamentos.tag_configure("Rejeitado", foreground="#dc2626")   # vermelho
        self.tree_orcamentos.tag_configure("Cancelado", foreground="#ea580c")   # laranja

        self.tree_orcamentos.bind("<Double-1>", self.visualizar_orcamento)

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
            status = valores[5]  # coluna Status

            # insere normalmente
            item_id = self.tree_orcamentos.insert("", "end", values=valores)

            # aplica a cor apenas no campo Status
            self.tree_orcamentos.item(item_id, tags=(status,))

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

    def editar_cliente(self, event=None):
        item = self.tree_clientes.selection()
        if not item:
            messagebox.showwarning("Atenção", "Selecione um cliente para editar.")
            return
        valores = self.tree_clientes.item(item[0], "values")
        cliente_id = valores[0]
        self.cursor.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
        cliente = self.cursor.fetchone()
        if cliente:
            self.abrir_formulario_cliente(cliente)

    # ------------------- Produtos -------------------
    def criar_aba_produtos(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Produtos")

        # --- Barra de botões ---
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(btn_frame, text="Adicionar", bootstyle=SUCCESS,
                command=lambda: self.abrir_formulario_produto()).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Editar", bootstyle=INFO,
                command=self.editar_produto).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Excluir", bootstyle=DANGER,
                command=self.excluir_produto).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Importar Arquivo", bootstyle=WARNING,
                command=lambda: self.importar_dados("produtos")).pack(side='left', padx=5)

        # --- Filtro por Tipo ---
        filtro_frame = ttk.Frame(frame)
        filtro_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(filtro_frame, text="Filtrar por Tipo:").pack(side='left', padx=5)
        self.combo_filtro_tipo = ttk.Combobox(filtro_frame, width=25, state="readonly")
        self.combo_filtro_tipo.pack(side='left', padx=5)
        ttk.Button(filtro_frame, text="Aplicar", command=self.filtrar_produtos_tipo).pack(side='left', padx=5)
        ttk.Button(filtro_frame, text="Limpar", command=lambda: self.carregar_produtos()).pack(side='left', padx=5)

        # --- Lista de produtos ---
        list_frame = ttk.LabelFrame(frame, text="Produtos Cadastrados", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)

        cols = ('ID', 'Código', 'Descrição', 'Tipo', 'Origem', 'Valor', 'ICMS%', 'IPI%', 'PIS/COFINS%')
        self.tree_produtos = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)

        for col in cols:
            self.tree_produtos.heading(col, text=col)
            if col in ('ID', 'ICMS%', 'IPI%', 'PIS/COFINS%'):
                self.tree_produtos.column(col, width=80)
            elif col == 'Valor':
                self.tree_produtos.column(col, width=100)
            else:
                self.tree_produtos.column(col, width=160)

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

        self.cursor.execute("SELECT DISTINCT tipo FROM produtos WHERE tipo IS NOT NULL AND tipo <> ''")
        tipos = [row[0] for row in self.cursor.fetchall()]
        self.combo_filtro_tipo['values'] = [""] + tipos  # vazio = mostrar todos
        
        # Agora trazemos todas as colunas que o Treeview espera
        self.cursor.execute('''
            SELECT id, codigo, descricao, tipo, origem_tributacao,
                valor_unitario, aliq_icms, aliq_ipi, aliq_pis
            FROM produtos
        ''')
        # Atualizar combobox com os tipos disponíveis
        

        for row in self.cursor.fetchall():
            # formatar valor unitário como moeda
            row = list(row)
            row[5] = formatar_moeda(row[5])  # valor_unitario
            self.tree_produtos.insert('', 'end', values=row)
    
    def filtrar_produtos_tipo(self):
        tipo = self.combo_filtro_tipo.get().strip()

        for item in self.tree_produtos.get_children():
            self.tree_produtos.delete(item)

        if not tipo:  # se não selecionar nada, mostra todos
            self.carregar_produtos()
            return

        self.cursor.execute('''
            SELECT id, codigo, descricao, tipo, origem_tributacao,
                valor_unitario, aliq_icms, aliq_ipi, aliq_pis
            FROM produtos
            WHERE tipo = ?
        ''', (tipo,))
        
        for row in self.cursor.fetchall():
            row = list(row)
            row[5] = formatar_moeda(row[5])  # formatar valor unitário
            self.tree_produtos.insert('', 'end', values=row)
    
    def editar_produto(self, event=None):
        item = self.tree_produtos.selection()
        if not item:
            messagebox.showwarning("Atenção", "Selecione um produto para editar.")
            return
        valores = self.tree_produtos.item(item[0], "values")
        produto_id = valores[0]
        self.cursor.execute("SELECT * FROM produtos WHERE id=?", (produto_id,))
        produto = self.cursor.fetchone()
        if produto:
            self.abrir_formulario_produto(produto)
    def excluir_produto(self):
        selecionados = self.tree_produtos.selection()
        if not selecionados:
            messagebox.showwarning("Atenção", "Selecione um ou mais produtos para excluir.")
            return

        if not messagebox.askyesno("Confirmação", f"Deseja realmente excluir {len(selecionados)} produto(s)?"):
            return

        try:
            for item in selecionados:
                valores = self.tree_produtos.item(item, "values")
                produto_id = valores[0]
                self.cursor.execute("DELETE FROM produtos WHERE id=?", (produto_id,))
            
            self.conn.commit()
            self.carregar_produtos()
            messagebox.showinfo("Sucesso", f"{len(selecionados)} produto(s) excluído(s) com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir produto(s): {e}")


    # ------------------- Pedidos/Orçamentos -------------------

    def criar_aba_pedidos(self):
        frame = tb.Frame(self.notebook)
        self.notebook.add(frame, text="Orçamentos")

        # Cabeçalho
        header_frame = tb.Labelframe(frame, text="Cabeçalho do Orçamento", padding=6, bootstyle="primary")
        header_frame.pack(fill='x', padx=10, pady=6)

        tb.Label(header_frame, text="Data do Orçamento:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5)
        self.entry_data_orc = tb.Entry(header_frame, width=15)
        self.entry_data_orc.grid(row=0, column=1, padx=5)
        self.entry_data_orc.insert(0, datetime.now().strftime('%d/%m/%Y'))

        # campo do número do orçamento (só aparece em edição)
        self.label_numero_orc_lbl = tb.Label(header_frame, text="Nº do Orçamento:", bootstyle="inverse-primary")
        self.label_numero_orc = tb.Label(header_frame, text="", bootstyle="secondary")


        tb.Label(header_frame, text="Representante:", bootstyle="inverse-primary").grid(row=0, column=4, sticky=W, padx=5)
        self.entry_representante = tb.Entry(header_frame, width=25)
        self.entry_representante.grid(row=0, column=5, padx=5)

        # 👉 Campo de Status
        # criar mas não mostrar ainda
        self.label_status_orc = tb.Label(header_frame, text="Status:", bootstyle="inverse-primary")
        self.combo_status_orc = tb.Combobox(
            header_frame,
            values=["Em Aberto", "Aprovado", "Cancelado", "Rejeitado"],
            width=15,
            state="readonly"
        )

        # Informações Comerciais
        extra_frame = tb.Labelframe(frame, text="Informações Comerciais", padding=6, bootstyle="info")
        extra_frame.pack(fill='x', padx=10, pady=6)

        tb.Label(extra_frame, text="Condições de Pagamento:").grid(row=0, column=0, sticky=W, padx=5, pady=2)
        self.entry_cond_pag = tb.Entry(extra_frame, width=30)
        self.entry_cond_pag.grid(row=0, column=1, padx=5)

        tb.Label(extra_frame, text="Validade (dias):").grid(row=0, column=2, sticky=W, padx=5, pady=2)
        self.entry_validade = tb.Entry(extra_frame, width=10)
        self.entry_validade.grid(row=0, column=3, padx=5)

        tb.Label(extra_frame, text="Desconto (R$):").grid(row=1, column=0, sticky=W, padx=5, pady=2)
        self.entry_desconto = tb.Entry(extra_frame, width=15)
        self.entry_desconto.grid(row=1, column=1, padx=5)

        tb.Label(extra_frame, text="Observações:").grid(row=2, column=0, sticky=NW, padx=5, pady=2)
        self.text_obs = tk.Text(extra_frame, width=80, height=3)
        self.text_obs.grid(row=2, column=1, columnspan=3, padx=5, pady=2)

        # Seleção cliente / produto
        top_frame = tb.Labelframe(frame, text="Novo Item", padding=6, bootstyle="secondary")
        top_frame.pack(fill='x', padx=10, pady=6)

        tb.Label(top_frame, text="Cliente:").grid(row=0, column=0, sticky=W, padx=5)
        self.combo_cliente = tb.Combobox(top_frame, width=60)
        self.combo_cliente.grid(row=0, column=1, padx=5, columnspan=4)

        tb.Label(top_frame, text="Produto:").grid(row=1, column=0, sticky=W, padx=5)
        self.combo_produto = tb.Combobox(top_frame, width=60, state='normal')
        self.combo_produto.grid(row=1, column=1, padx=5, columnspan=3)

        tb.Label(top_frame, text="Quantidade:").grid(row=1, column=4, sticky=W, padx=5)
        self.entry_qtd = tb.Entry(top_frame, width=8)
        self.entry_qtd.grid(row=1, column=5, padx=5)

        tb.Button(top_frame, text="Adicionar Item", command=self.adicionar_item_pedido, bootstyle="success").grid(row=1, column=6, padx=5)
        tb.Button(top_frame, text="Remover Item", command=self.remover_item, bootstyle="danger").grid(row=1, column=7, padx=5)

        # Itens
        items_frame = tb.Labelframe(frame, text="Itens do Orçamento", padding=6, bootstyle="info")
        items_frame.pack(fill='both', expand=True, padx=10, pady=6)

        cols = ('Produto', 'Qtd', 'Valor Unit.', 'Total')
        self.tree_pedido_items = tb.Treeview(items_frame, columns=cols, show='headings', height=9, bootstyle="info")
        for col in cols:
            self.tree_pedido_items.heading(col, text=col)
            self.tree_pedido_items.column(col, width=180)
            self.tree_pedido_items.pack(fill='both', expand=True)

        # Totais + ações
        totals_frame = tb.Labelframe(frame, text="Totais & Ações", padding=6, bootstyle="warning")
        totals_frame.pack(fill='x', padx=10, pady=6)

        self.label_subtotal = tb.Label(totals_frame, text="Subtotal: R$ 0,00", bootstyle="secondary")
        self.label_subtotal.grid(row=0, column=0, padx=8)

        self.label_impostos = tb.Label(totals_frame, text="Impostos: R$ 0,00", bootstyle="secondary")
        self.label_impostos.grid(row=0, column=1, padx=8)

        self.label_total = tb.Label(totals_frame, text="TOTAL: R$ 0,00",
                                    font=('Segoe UI', 11, 'bold'), bootstyle="success")
        self.label_total.grid(row=0, column=2, padx=8)

        self.btn_finalizar_pedido = tb.Button(
            totals_frame, text="Salvar Orçamento",
            command=self.finalizar_pedido, bootstyle="success"
        )
        self.btn_finalizar_pedido.grid(row=0, column=3, padx=8)

        tb.Button(
            totals_frame, text="Novo Orçamento",
            command=self.limpar_pedido, bootstyle="secondary"
        ).grid(row=0, column=4, padx=8)

        tb.Button(
            totals_frame, text="Exportar p/ PDF",
            command=lambda: self.gerar_pdf_orcamento(self.label_numero_orc.cget("text")),
            bootstyle="danger-outline"
        ).grid(row=0, column=6, padx=8)

        # carregar clientes e produtos nos combos
        self.carregar_combos_pedido()


    def carregar_orcamento_para_edicao(self, numero_pedido):
        """
        Carrega um orçamento salvo para edição na aba 'Orçamentos'.
        """
        # Buscar dados principais do pedido (inclui cliente_id e status)
        self.cursor.execute('''
            SELECT p.data_pedido, p.cliente_id, p.valor_produtos, p.valor_icms, p.valor_ipi, 
                p.valor_pis, p.valor_cofins, p.valor_total, p.representante, 
                p.condicoes_pagamento, p.desconto, p.observacoes, p.validade, p.status
            FROM pedidos p
            WHERE p.numero_pedido = ?
        ''', (numero_pedido,))
        pedido = self.cursor.fetchone()
        if not pedido:
            messagebox.showerror("Erro", "Orçamento não encontrado para edição.")
            return

        (data_pedido, cliente_id, subtotal, icms, ipi, pis, cofins, total,
        representante, cond_pag, desconto, observacoes, validade, status) = pedido

        # preencher cabeçalho
        try:
            data_str = datetime.strptime(data_pedido, "%Y-%m-%d %H:%M:%S").strftime('%d/%m/%Y')
        except:
            data_str = data_pedido

        self.entry_data_orc.delete(0, tk.END)
        self.entry_data_orc.insert(0, data_str)
        self.label_numero_orc.config(text=numero_pedido)
        self.entry_representante.delete(0, tk.END)
        self.entry_representante.insert(0, representante or "")
        self.entry_cond_pag.delete(0, tk.END)
        self.entry_cond_pag.insert(0, cond_pag or "")
        self.entry_validade.delete(0, tk.END)
        self.entry_validade.insert(0, validade or "")
        self.entry_desconto.delete(0, tk.END)
        self.entry_desconto.insert(0, str(desconto or 0))
        self.text_obs.delete("1.0", tk.END)
        self.text_obs.insert("1.0", observacoes or "")

       
        
        self.label_status_orc.grid(row=0, column=6, sticky=W, padx=5)
        self.combo_status_orc.grid(row=0, column=7, padx=5)
        self.combo_status_orc.set(status or "Em Aberto")
        # 👉 mostrar número do orçamento quando editar
        self.label_numero_orc_lbl.grid(row=0, column=2, sticky=W, padx=5)
        self.label_numero_orc.grid(row=0, column=3, padx=5)
        self.label_numero_orc.config(text=numero_pedido)

        # selecionar cliente no combo (formato: "id - nome")
        self.cursor.execute("SELECT razao_social FROM clientes WHERE id=?", (cliente_id,))
        cliente_nome = self.cursor.fetchone()
        if cliente_nome:
            self.combo_cliente.set(f"{cliente_id} - {cliente_nome[0]}")
        else:
            self.combo_cliente.set("")

        # carregar itens do pedido para a lista temporária e treeview
        self.itens_pedido_temp = []
        for item in self.tree_pedido_items.get_children():
            self.tree_pedido_items.delete(item)

        self.cursor.execute('''
            SELECT pr.id, pr.codigo, pr.descricao, pi.qtd, pi.valor_unitario
            FROM pedido_itens pi
            JOIN produtos pr ON pi.produto_id = pr.id
            WHERE pi.numero_pedido = ?
        ''', (numero_pedido,))
        itens = self.cursor.fetchall()
        for prod_id, codigo, descricao, qtd, valor_unit in itens:
            total_item = (qtd or 0) * (valor_unit or 0)
            self.tree_pedido_items.insert('', 'end',
                values=(f"{codigo} - {descricao}", qtd,
                        formatar_moeda(valor_unit), formatar_moeda(total_item)))
            self.itens_pedido_temp.append({
                'produto_id': prod_id,
                'codigo': codigo,
                'descricao': descricao,
                'qtd': qtd,
                'valor': float(valor_unit)
            })

        # atualizar totais e marcar modo edição
        self.atualizar_totais()
        self.edicao_numero_pedido = numero_pedido
        # trocar texto do botão
        try:
            self.btn_finalizar_pedido.config(text="Atualizar Orçamento")
        except:
            pass

        # muda para a aba de orçamentos
        for i in range(len(self.notebook.tabs())):
            if self.notebook.tab(i, "text") == "Orçamentos":
                self.notebook.select(i)
                break

    def remover_item(self):
        selecionado = self.tree_pedido_items.selection()
        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione um item para remover.")
            return

        for item_id in selecionado:
            valores = self.tree_pedido_items.item(item_id, "values")
            if valores:
                descricao = valores[0]  
                for i, item in enumerate(self.itens_pedido_temp):
                    if f"{item['codigo']} - {item['descricao']}" == descricao:
                        del self.itens_pedido_temp[i]
                        break
           
            self.tree_pedido_items.delete(item_id)

        self.atualizar_totais() 

    def carregar_combos_pedido(self):
        self.cursor.execute('SELECT id, razao_social FROM clientes')
        clientes = [f"{row[0]} - {row[1]}" for row in self.cursor.fetchall()]
        self.combo_cliente['values'] = clientes     
        self.cursor.execute('SELECT id, codigo, descricao FROM produtos')
        produtos = [f"{row[0]} - {row[1]} - {row[2]}" for row in self.cursor.fetchall()]
        self.combo_produto['values'] = produtos
    def filtrar_clientes(self, event=None):
        texto = self.combo_cliente.get().lower()
        self.cursor.execute('SELECT id, razao_social FROM clientes')
        todos = [f"{row[0]} - {row[1]}" for row in self.cursor.fetchall()]

        if texto:
            filtrados = [c for c in todos if texto in c.lower()]
        else:
            filtrados = todos

        self.combo_cliente['values'] = filtrados
        self.combo_cliente.event_generate('<Down>')  # abre a lista automaticamente

    def adicionar_item_pedido(self):
        try:
            produto_str = self.combo_produto.get()
            if not produto_str:
                messagebox.showwarning("Atenção", "Selecione um produto!")
                return

            produto_id = int(produto_str.split(" - ")[0])
            qtd = float(self.entry_qtd.get() or 1)

            self.cursor.execute("SELECT codigo, descricao, valor_unitario FROM produtos WHERE id = ?", (produto_id,))
            produto = self.cursor.fetchone()
            if not produto:
                messagebox.showerror("Erro", "Produto não encontrado.")
                return

            codigo, descricao, preco_venda = produto

            # Verifica se já existe o mesmo produto na lista
            encontrado = False
            for item in self.itens_pedido_temp:
                if item["produto_id"] == produto_id:
                    # Atualiza quantidade e recalcula
                    item["qtd"] += qtd
                    encontrado = True
                    break

            if not encontrado:
                # adiciona novo item
                self.itens_pedido_temp.append({
                    "produto_id": produto_id,
                    "codigo": codigo,
                    "descricao": descricao,
                    "qtd": qtd,
                    "valor": float(preco_venda or 0)
                })

            # Atualiza visualização na treeview
            self.tree_pedido_items.delete(*self.tree_pedido_items.get_children())
            for item in self.itens_pedido_temp:
                total_item = item["qtd"] * item["valor"]
                self.tree_pedido_items.insert(
                    "", "end",
                    values=(f"{item['codigo']} - {item['descricao']}",
                            item["qtd"],
                            formatar_moeda(item["valor"]),
                            formatar_moeda(total_item))
                )

            self.atualizar_totais()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao adicionar item: {e}")

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
        cliente_id = valores[0]  
        self.cursor.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
        cliente = self.cursor.fetchone()

        if not cliente:
            messagebox.showerror("Erro", "Cliente não encontrado.")
            return

        keys = ["ID", "Razão Social", "CNPJ", "IE", "Endereço", "Cidade", "Estado", "CEP", "Telefone", "Email"]

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

        
        ttk.Button(frame_botoes, text="Exportar p/ Excel", bootstyle=DANGER,command=self.exportar_excel_orcamento).pack(side="left", padx=5)
        ttk.Button(frame_botoes, bootstyle=INFO, text="Abrir Orçamento", 
                command=lambda: (top.destroy(), self.carregar_orcamento_para_edicao(numero_pedido))).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Fechar", command=top.destroy).pack(side="right", padx=5)
        
    def limpar_pedido(self):
        for item in self.tree_pedido_items.get_children():
            self.tree_pedido_items.delete(item)
            self.itens_pedido_temp = []
            self.atualizar_totais()
            self.combo_cliente.set('')
            self.combo_produto.set('')
            self.entry_qtd.delete(0, tk.END)
        # esconder campo de status quando for novo orçamento
        try:
            self.label_status_orc.grid_forget()
            self.combo_status_orc.grid_forget()
        except:
            pass

        # esconder número do orçamento quando for novo
        try:
            self.label_numero_orc_lbl.grid_forget()
            self.label_numero_orc.grid_forget()
        except:
            pass

    
    def finalizar_pedido(self):
        if not self.combo_cliente.get() or not self.itens_pedido_temp:
            messagebox.showwarning("Atenção", "Selecione um cliente e adicione itens!")
            return
        try:
            cliente_id = int(self.combo_cliente.get().split(' - ')[0])
            # se estiver em edição, usamos o numero existente
            if hasattr(self, 'edicao_numero_pedido') and self.edicao_numero_pedido:
                numero_pedido = self.edicao_numero_pedido
            else:
                self.cursor.execute("SELECT COUNT(*) FROM pedidos")
                total_registros = self.cursor.fetchone()[0] or 0
                numero_pedido = f"ORC-{total_registros+1:04d}"
                self.label_numero_orc.config(text=numero_pedido)

            data_pedido = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # calcular totais
            subtotal, total_icms, total_ipi, total_pis, total_cofins, total = self.calcular_totais(self.itens_pedido_temp)

            representante = self.entry_representante.get()
            cond_pag = self.entry_cond_pag.get()
            validade = self.entry_validade.get()
            desconto = float(self.entry_desconto.get() or 0)
            observacoes = self.text_obs.get("1.0", tk.END).strip()

            total_final = total - desconto

            if hasattr(self, 'edicao_numero_pedido') and self.edicao_numero_pedido:
                # === Atualizar pedido existente ===
                status = self.combo_status_orc.get() or "Em Aberto"  # pega do combobox

                self.cursor.execute('''
                    UPDATE pedidos
                    SET data_pedido=?, cliente_id=?, valor_produtos=?, valor_icms=?, valor_ipi=?, valor_pis=?, valor_cofins=?, valor_total=?,
                        representante=?, condicoes_pagamento=?, desconto=?, status=?, observacoes=?, validade=?
                    WHERE numero_pedido=?
                ''', (data_pedido, cliente_id, subtotal, total_icms, total_ipi, total_pis, total_cofins, total_final,
                    representante, cond_pag, desconto, status, observacoes, validade, numero_pedido))

                # remover itens antigos e inserir os novos
                self.cursor.execute('DELETE FROM pedido_itens WHERE numero_pedido = ?', (numero_pedido,))
                for item in self.itens_pedido_temp:
                    self.cursor.execute('''
                        INSERT INTO pedido_itens (numero_pedido, produto_id, qtd, valor_unitario)
                        VALUES (?, ?, ?, ?)
                    ''', (numero_pedido, item['produto_id'], item['qtd'], item['valor']))

                self.conn.commit()
                messagebox.showinfo("Sucesso", f"Orçamento {numero_pedido} atualizado com sucesso!")
                # sair do modo edição
                self.edicao_numero_pedido = None
                try:
                    self.btn_finalizar_pedido.config(text="Salvar Orçamento")
                except:
                    pass
            else:
                # === Inserir novo pedido ===
                self.cursor.execute('''
                    INSERT INTO pedidos 
                    (numero_pedido, data_pedido, cliente_id, valor_produtos, valor_icms, valor_ipi, valor_pis, valor_cofins, valor_total, representante,
                    condicoes_pagamento, desconto, status, observacoes, validade)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (numero_pedido, data_pedido, cliente_id, subtotal, total_icms, total_ipi, total_pis, total_cofins, total_final,
                    representante, cond_pag, desconto, "Em Aberto", observacoes, validade))

                for item in self.itens_pedido_temp:
                    self.cursor.execute('''
                        INSERT INTO pedido_itens (numero_pedido, produto_id, qtd, valor_unitario)
                        VALUES (?, ?, ?, ?)
                    ''', (numero_pedido, item['produto_id'], item['qtd'], item['valor']))

                self.conn.commit()
                messagebox.showinfo("Sucesso", f"Orçamento {numero_pedido} salvo com sucesso!")

            # limpar itens temporários e UI
            self.itens_pedido_temp.clear()
            self.tree_pedido_items.delete(*self.tree_pedido_items.get_children())
            self.carregar_combos_pedido()
            try:
                # recarregar a lista de orçamentos (se existir)
                self.buscar_orcamento()
            except:
                pass

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar/atualizar orçamento: {e}")

    def abrir_formulario_produto(self, produto=None):
        top = tk.Toplevel(self.root)
        top.title("Cadastro de Produto")
        top.geometry("700x400")

        labels = [
            'Código:', 'Descrição:', 'Valor Unitário:',
            'Tipo:', 'Origem Tributação:', 'Voltagem:',
            'ICMS (%):', 'IPI (%):', 'PIS (%):', 'COFINS (%):'
        ]

        entries = {}
        for i, label in enumerate(labels):
            ttk.Label(top, text=label).grid(row=i//2, column=(i%2)*2, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(top, width=28)
            entry.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=5)
            chave = normalizar_chave(label)
            entries[chave] = entry

        # se for edição, preencher
        if produto:
            keys = ['id','codigo','descricao','valor_unitario','tipo','origem_tributacao','voltagem',
                    'aliq_icms','aliq_ipi','aliq_pis','aliq_cofins']
            for k, v in zip(keys, produto):
                if k in entries:
                    entries[k].insert(0, v or "")

            def salvar():
                dados = {k: v.get().strip() for k,v in entries.items()}
                if not dados['codigo']:
                    messagebox.showerror("Erro", "O campo CÓDIGO do produto está vazio.")
                    return
                if not dados['descricao']:
                    messagebox.showerror("Erro", "O campo DESCRIÇÃO do produto está vazio.")
                    return
                if not dados['valor_unitario']:
                    messagebox.showerror("Erro", "O campo VALOR UNITÁRIO está vazio.")
                    return

                try:
                    if produto:  # atualização
                        self.cursor.execute('''
                            UPDATE produtos
                            SET codigo=?, descricao=?, valor_unitario=?, tipo=?, origem_tributacao=?, voltagem=?,
                                aliq_icms=?, aliq_ipi=?, aliq_pis=?, aliq_cofins=?
                            WHERE id=?
                        ''', (dados['codigo'], dados['descricao'], float(dados['valor_unitario'] or 0),
                            dados.get('tipo'), dados.get('origem_tributacao'), dados.get('voltagem'),
                            float(dados.get('icms') or 0), float(dados.get('ipi') or 0),
                            float(dados.get('pis') or 0), float(dados.get('cofins') or 0),
                            produto[0]))
                    else:  # novo
                        self.cursor.execute('''
                            INSERT INTO produtos (codigo, descricao, valor_unitario, tipo, origem_tributacao, voltagem,
                                                aliq_icms, aliq_ipi, aliq_pis, aliq_cofins)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (dados['codigo'], dados['descricao'], float(dados['valor_unitario'] or 0),
                            dados.get('tipo'), dados.get('origem_tributacao'), dados.get('voltagem'),
                            float(dados.get('icms') or 0), float(dados.get('ipi') or 0),
                            float(dados.get('pis') or 0), float(dados.get('cofins') or 0)))
                    self.conn.commit()
                    self.carregar_produtos()
                    top.destroy()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Erro", "Código já cadastrado!")   
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar produto: {e}")

            ttk.Button(top, text="Salvar", command=salvar).grid(row=6, column=0, columnspan=2, pady=10)

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

           
            # --- pegar nome do cliente para o nome do arquivo ---
            cliente_nome = None

            # tenta pegar do combo
            if self.combo_cliente.get():
                try:
                    cliente_nome = self.combo_cliente.get().split(" - ", 1)[1]
                except:
                    cliente_nome = self.combo_cliente.get()

            # se ainda não tiver nome, busca pelo numero do pedido
            if not cliente_nome:
                numero_pedido = self.label_numero_orc.cget("text") if self.label_numero_orc.cget("text") else self.edicao_numero_pedido
                if numero_pedido:
                    self.cursor.execute("""
                        SELECT c.razao_social 
                        FROM pedidos p
                        JOIN clientes c ON p.cliente_id = c.id
                        WHERE p.numero_pedido = ?
                    """, (numero_pedido,))
                    row = self.cursor.fetchone()
                    if row:
                        cliente_nome = row[0]

            # fallback se ainda estiver vazio
            if not cliente_nome:
                cliente_nome = "cliente"

            # sanitizar nome (sem acentos/espacos estranhos)
            cliente_nome = unicodedata.normalize('NFKD', cliente_nome).encode('ASCII', 'ignore').decode('utf-8')
            cliente_nome = re.sub(r'[^a-zA-Z0-9_-]', '_', cliente_nome)

            # data no formato dd-mm-yy
            from datetime import datetime
            data_str = datetime.now().strftime("%d-%m-%y")

            nome_sugerido = f"orcamento-{cliente_nome}-{data_str}.pdf"

            data_str = datetime.now().strftime("%d-%m-%y")

            nome_sugerido = f"orcamento-{cliente_nome}-{data_str}.pdf"

            # Montar PDF
            arquivo = filedialog.asksaveasfilename(
                initialfile=nome_sugerido,
                defaultextension=".pdf",
                filetypes=[("Arquivos PDF", "*.pdf")],
                title="Salvar orçamento em PDF"
            )
            if not arquivo:
                return  # se cancelar, não gera
            # cria o doc no caminho escolhido
            doc = SimpleDocTemplate(arquivo, pagesize=A4,
                                    leftMargin=30, rightMargin=30,
                                    topMargin=40, bottomMargin=40)
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
                ["Validade do Orçamento:", f"{validade} dias"],
            ]
            t_info = Table(info_cliente, colWidths=[100, 380])
            estilo_tabela(t_info, header=False)
            elementos.append(t_info)
            elementos.append(Spacer(1, 20))

            desc_style = ParagraphStyle(
                name="Descricao",
                fontSize=8,          # um pouco menor ajuda caber
                leading=10,          # altura da linha
                wordWrap='CJK'       # força quebra automática
            )

            # Tabela de produtos
            tabela = [["Código", "Descrição", "Qtd", "Valor Unit.", "Total"]]
            for codigo, descricao, qtd, valor_unit in produtos:
                total_item = (qtd or 0) * (valor_unit or 0)
                tabela.append([
                    codigo,
                    Paragraph(descricao, desc_style),
                    str(qtd),
                    formatar_moeda(valor_unit or 0),
                    formatar_moeda(total_item)
                ])

            t = Table(tabela, colWidths=[60, 280, 50, 80, 80], repeatRows=1, hAlign="CENTER")
            estilo_tabela(t, header=True)
            elementos.append(t)

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
            messagebox.showinfo("Sucesso", f"PDF gerado: {arquivo}")

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
                        idx_codigo = find_index('cod', 'codigo', 'produto_id', 'id')
                        idx_desc   = find_index('descricao', 'descri', 'produto', 'nome')
                        idx_valor  = find_index('valor', 'preco', 'unitario')
                        idx_tipo   = find_index('tipo')
                        idx_origem = find_index('origem', 'tributacao')
                        idx_ipi    = find_index('ipi')
                        idx_pis    = find_index('pis', 'pis/cofins', 'cofins')
                        idx_icms   = find_index('icms')


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
                                codigo = str(row[idx_codigo]).strip() if idx_codigo is not None and row[idx_codigo] else ''
                                descricao = str(row[idx_desc]).strip() if idx_desc is not None and row[idx_desc] else ''
                                valor_unitario = to_float_cell(row, idx_valor)

                                tipo = str(row[idx_tipo]).strip() if idx_tipo is not None and row[idx_tipo] else ''
                                origem = str(row[idx_origem]).strip() if idx_origem is not None and row[idx_origem] else ''

                                aliq_icms = to_float_cell(row, idx_icms)
                                aliq_ipi  = to_float_cell(row, idx_ipi)
                                aliq_pis  = to_float_cell(row, idx_pis)
                                aliq_cofins = 0.0  # opcional (se quiser separar do PIS/COFINS)

                                if not codigo or not descricao:
                                    continue

                                self.cursor.execute('''
                                    INSERT INTO produtos (codigo, descricao, valor_unitario, tipo, origem_tributacao,
                                                        aliq_icms, aliq_ipi, aliq_pis, aliq_cofins)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (codigo, descricao, valor_unitario, tipo, origem,
                                    aliq_icms, aliq_ipi, aliq_pis, aliq_cofins))
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
           
    def exportar_excel_orcamento(self, numero_pedido=None):
        if not numero_pedido:
            item = self.tree_orcamentos.selection()
            if item:
                valores = self.tree_orcamentos.item(item[0], "values")
                numero_pedido = valores[0]
            else:
                messagebox.showerror("Erro", "Nenhum orçamento selecionado.")
                return

        # Buscar dados do orçamento
        self.cursor.execute('''
            SELECT p.data_pedido, c.razao_social, c.cnpj, c.endereco, c.cidade, c.estado,
                p.valor_produtos, p.valor_icms, p.valor_ipi, p.valor_pis, p.valor_cofins,
                p.valor_total, p.representante, p.condicoes_pagamento, p.observacoes, p.validade
            FROM pedidos p
            JOIN clientes c ON p.cliente_id = c.id
            WHERE p.numero_pedido = ?
        ''', (numero_pedido,))
        pedido = self.cursor.fetchone()

        if not pedido:
            messagebox.showerror("Erro", "Orçamento não encontrado.")
            return

        # Preparar dados
        try:
            data_formatada = datetime.strptime(pedido[0], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%y %H:%M")
        except:
            data_formatada = pedido[0]

        cliente, cnpj, endereco, cidade, estado, subtotal, icms, ipi, pis, cofins, total, representante, cond_pag, obs, validade = (
            pedido[1], pedido[2], pedido[3], pedido[4], pedido[5],
            pedido[6], pedido[7], pedido[8], pedido[9], pedido[10],
            pedido[11], pedido[12], pedido[13], pedido[14], pedido[15]
        )

        endereco_formatado = " - ".join(filter(None, [endereco, cidade, estado])) or "Não informado"

        # Criar planilha
        wb = Workbook()
        ws = wb.active
        ws.title = f"Orçamento {numero_pedido}"

        # Estilos
        header_fill = PatternFill("solid", fgColor="1E3A8A")  # azul escuro
        header_font = Font(bold=True, color="FFFFFF")
        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))

        # Logo (se existir)
        try:
            logo = Image("logo.png")
            logo.width, logo.height = 150, 40
            ws.add_image(logo, "A1")
        except:
            pass

        # Cabeçalho
        ws.merge_cells("A2:E2")
        ws["A2"] = f"ORÇAMENTO Nº {numero_pedido}"
        ws["A2"].font = Font(size=14, bold=True)
        ws["A2"].alignment = Alignment(horizontal="center")

        ws["A4"], ws["B4"] = "Data:", data_formatada
        ws["A5"], ws["B5"] = "Cliente:", cliente
        ws["A6"], ws["B6"] = "CNPJ:", cnpj
        ws["A7"], ws["B7"] = "Endereço:", endereco_formatado
        ws["A8"], ws["B8"] = "Representante:", representante
        ws["A9"], ws["B9"] = "Condições de Pagamento:", cond_pag
        ws["A10"], ws["B10"] = "Validade:", f"{validade} dias"

        # Itens do orçamento
        ws.append([])
        linha_inicio = 12
        colunas = ["Código", "Descrição", "Qtd", "Valor Unitário", "Total Item"]

        for col, nome in enumerate(colunas, start=1):
            cel = ws.cell(row=linha_inicio, column=col, value=nome)
            cel.font = header_font
            cel.fill = header_fill
            cel.alignment = Alignment(horizontal="center")
            cel.border = border

        # Buscar itens
        self.cursor.execute('''
            SELECT pr.codigo, pr.descricao, pi.qtd, pi.valor_unitario
            FROM pedido_itens pi
            JOIN produtos pr ON pi.produto_id = pr.id
            WHERE pi.numero_pedido = ?
        ''', (numero_pedido,))
        itens = self.cursor.fetchall()

        linha = linha_inicio + 1
        subtotal_calc = 0

        for cod, desc, qtd, v_unit in itens:
            v_total_item = qtd * v_unit
            subtotal_calc += v_total_item
            dados = [cod, desc, qtd, v_unit, v_total_item]

            for col, valor in enumerate(dados, start=1):
                cel = ws.cell(row=linha, column=col, value=valor)
                if col >= 3:  # qtd, unitário e total
                    cel.number_format = 'R$ #,##0.00'
                    cel.alignment = Alignment(horizontal="right")
                cel.border = border

            # Linhas alternadas (efeito zebra)
            if (linha % 2) == 0:
                for col in range(1, 6):
                    ws.cell(row=linha, column=col).fill = PatternFill("solid", fgColor="F9F9F9")

            linha += 1

        # Resumo de totais
        linha += 1
        totais = [
            ("Subtotal:", subtotal_calc),
            ("ICMS:", icms),
            ("IPI:", ipi),
            ("PIS:", pis),
            ("COFINS:", cofins)
        ]

        desconto = float(self.entry_desconto.get() or 0)
        if desconto > 0:
            totais.append(("Desconto:", -desconto))

        total_geral = subtotal_calc + icms + ipi + pis + cofins - desconto
        totais.append(("TOTAL GERAL", total_geral))

        for label, valor in totais:
            ws.merge_cells(start_row=linha, start_column=4, end_row=linha, end_column=4)
            cel_label = ws.cell(row=linha, column=4, value=label)
            cel_label.alignment = Alignment(horizontal="right")
            cel_label.font = Font(bold=True)
            cel_val = ws.cell(row=linha, column=5, value=valor)
            cel_val.number_format = 'R$ #,##0.00'
            if "TOTAL" in label:
                cel_label.fill = PatternFill("solid", fgColor="228B22")
                cel_label.font = Font(bold=True, size=12, color="FFFFFF")
                cel_val.font = Font(bold=True, size=12, color="FFFFFF")
                cel_val.fill = PatternFill("solid", fgColor="228B22")
            linha += 1

        # Observações
        if obs:
            linha += 1
            ws.merge_cells(f"A{linha}:E{linha+2}")
            cel = ws[f"A{linha}"]
            cel.value = f"Observações: {obs}"
            cel.alignment = Alignment(wrap_text=True, vertical="top")
            cel.fill = PatternFill("solid", fgColor="EFEFEF")

        # Espaço para assinatura
        linha += 4
        ws.merge_cells(f"A{linha}:C{linha}")
        ws[f"A{linha}"] = "Assinatura do Cliente: __________________________"

        # Rodapé
        linha += 2
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=5)
        ws.cell(row=linha, column=1,
                value=f"Gerado em {datetime.now().strftime('%d/%m/%y %H:%M')} - Sistema Interno de Orçamentos"
            ).alignment = Alignment(horizontal="center")

        # Autoajuste das colunas
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        # Salvar
        nome_arquivo = f"Orcamento_{numero_pedido}.xlsx"
        wb.save(nome_arquivo)
        os.startfile(nome_arquivo)

    def __del__(self):
            if hasattr(self, 'conn'):
                self.conn.close()

if __name__ == "__main__":
    root = tb.Window(themename="superhero")
    app = SistemaPedidos(root)
    root.mainloop()
    
