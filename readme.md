# Sistema de Orçamentos

Aplicação em Python com Tkinter e SQLite para cadastro de clientes, produtos e geração de orçamentos em PDF e Excel.

## Requisitos

- Python 3.10+
- Bibliotecas listadas em `requirements.txt`

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
source venv/bin/activate  # Linux/Mac
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

## Estrutura

- `main.py` → código principal  
- `requirements.txt` → dependências  
- `pedidos.db` → banco SQLite (gerado automaticamente)  
- `*.xlsx` / `*.pdf` → arquivos exportados  
