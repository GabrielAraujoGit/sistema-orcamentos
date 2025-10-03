# Sistema de Orçamentos

Aplicação em Python com Tkinter e SQLite para cadastro de clientes, produtos e geração de orçamentos em PDF e Excel.

## Funcionalidades

- Cadastro, edição, exclusão e importação de **clientes** e **produtos**.
- Criação e edição de **orçamentos** com controle de status: *Em Aberto*, *Aprovado*, *Cancelado*, *Rejeitado*.
- Consulta avançada de orçamentos por filtros (número, cliente, representante, status e período).
- Exportação de orçamentos em **PDF padronizado** e **Excel (.xlsx)**.
- Banco de dados SQLite (`pedidos.db`) gerado automaticamente.

## Requisitos

- Python 3.10+
- Dependências listadas em `requirements.txt`:
  - `openpyxl`
  - `reportlab`
  - `ttkbootstrap`

## Instalação

Clone o repositório:

```bash
git clone https://github.com/SEU_USUARIO/sistema-orcamentos.git
cd sistema-orcamentos
```

Crie um ambiente virtual (opcional, recomendado):

```bash
python -m venv venv
venv\Scripts\activate     # Windows
source venv/bin/activate    # Linux/Mac
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

A interface gráfica será aberta com abas para **Clientes**, **Produtos**, **Orçamentos** e **Consulta de Orçamentos**.

## Estrutura

- `main.py` → código principal  
- `requirements.txt` → dependências  
- `pedidos.db` → banco SQLite (gerado automaticamente)  
- `*.xlsx` / `*.pdf` → arquivos exportados  
