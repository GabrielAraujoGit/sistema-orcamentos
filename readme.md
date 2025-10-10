# 💼 Sistema de Orçamentos

Aplicação em **Python (Tkinter + SQLite)** para cadastro de clientes, produtos e geração de orçamentos em **PDF** e **Excel**.

---

## 🚀 Funcionalidades

- Cadastro, edição, exclusão e importação de **clientes** e **produtos**.  
- Criação e edição de **orçamentos** com controle de status:  
  *Em Aberto*, *Aprovado*, *Cancelado*, *Rejeitado*.  
- Consulta avançada de orçamentos por filtros (**número**, **cliente**, **representante**, **status**, **período**).  
- Exportação de orçamentos em **PDF padronizado** e **Excel (.xlsx)**.  
- Banco de dados **SQLite (`pedidos.db`)** gerado automaticamente.  

---

## 🧩 Requisitos

- Python **3.10+**
- Dependências listadas em `requirements.txt`:
  - `openpyxl`
  - `reportlab`
  - `ttkbootstrap`

---

## ⚙️ Instalação

Clone o repositório:

```bash
git clone https://github.com/SEU_USUARIO/sistema-orcamentos.git
cd sistema-orcamentos
```

Crie um ambiente virtual (opcional, recomendado):

```bash
python -m venv venv
venv\Scripts\activate     # Windows
source venv/bin/activate  # Linux/Mac
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## ▶️ Uso

```bash
python main.py
```

A interface gráfica será aberta com abas para **Clientes**, **Produtos**, **Orçamentos** e **Consulta de Orçamentos**.

---

## 🗂️ Estrutura do Projeto

```
sistema-orcamentos/
│
├── main.py              # Código principal
├── requirements.txt     # Dependências do projeto
├── pedidos.db           # Banco SQLite (gerado automaticamente)
├── docs/
│   └── images/          # Prints de tela (opcional)
└── arquivos_exportados/ # PDFs e planilhas .xlsx geradas
```

---

## 📸 Interface

### Tela de Clientes
<img width="1677" height="965" alt="image" src="https://github.com/user-attachments/assets/69945def-8a5c-46a0-bf45-ba9377eea2da" />


### Tela de Produtos
<img width="1679" height="972" alt="image" src="https://github.com/user-attachments/assets/64821873-58b2-491f-b243-265c946b724f" />


### Tela de Orçamentos
<img width="1676" height="969" alt="image" src="https://github.com/user-attachments/assets/40edfe8e-5ab7-4a93-af85-6f62e011f55f" />


---


