# EletroFlow — Sistema Interno de Orçamentos

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Version](https://img.shields.io/badge/Version-v1.0.0-brightgreen?style=for-the-badge)](#)
[![License](https://img.shields.io/badge/License-Internal-yellow?style=for-the-badge)](#)
[![Status](https://img.shields.io/badge/Status-MAINTENANCE-orange?style=for-the-badge)](#)
[![Last Update](https://img.shields.io/badge/Updated-Oct_2025-lightgrey?style=for-the-badge)](#)
[![Support](https://img.shields.io/badge/Support-TI_Eletrofrio-8A2BE2?style=for-the-badge)](#)



## 📋 Sumário  
1. [Visão Geral](#visão-geral)  
2. [Funcionalidades](#funcionalidades)  
3. [Fluxo de Uso (Exemplo)](#fluxo-de-uso-exemplo)  
4. [Instalação & Execução](#instalação--execução)  
5. [Estrutura do Projeto](#estrutura-do-projeto)  
6. [Exemplos de Saída](#exemplos-de-saída)  
7. [Contribuição & Padrões](#contribuição--padrões)  
8. [Backup & Migração de Dados](#backup--migração-de-dados)  
9. [Contatos / Suporte Interno](#contatos--suporte-interno)

---

## Visão Geral  

Aplicação interna desenvolvida em **Python (Tkinter + SQLite)** para **gestão de orçamentos comerciais** da Eletrofrio.  
Centraliza cadastros de clientes e produtos, gera documentos padronizados e mantém histórico local de orçamentos.  

Principais objetivos:
- Reduzir retrabalho e erros manuais;  
- Padronizar a emissão de orçamentos;  
- Facilitar consultas e controle de status;  
- Permitir exportação em formatos oficiais (PDF/Excel).  

---

## Funcionalidades  

- Cadastro, edição e exclusão de **clientes** e **produtos**;  
- Emissão e controle de **orçamentos comerciais**;  
- Status configuráveis: *Aberto*, *Aprovado*, *Cancelado*, *Rejeitado*;  
- Filtros avançados por **cliente**, **representante**, **status**, **período**;  
- Exportação para **PDF padronizado** e **Excel (.xlsx)**;  
- Banco de dados **SQLite (`pedidos.db`)** criado automaticamente.  

---

## Fluxo de Uso (Exemplo)  

1. Abrir o sistema (`python main.py`);  
2. Cadastrar ou importar clientes e produtos;  
3. Criar um novo orçamento e adicionar itens;  
4. Exportar o documento em PDF ou Excel;  
5. Atualizar o status conforme aprovação ou cancelamento.  

---

## Instalação & Execução  

### Requisitos  

- **Python 3.10+**  
- Dependências listadas em `requirements.txt`:  
  - `openpyxl`  
  - `reportlab`  
  - `ttkbootstrap`  

### Instalação  

```bash
git clone https://github.com/eletrofrio/sistema-orcamentos.git
cd sistema-orcamentos
python -m venv venv
venv\Scripts\activate     # Windows
source venv/bin/activate   # Linux/Mac
pip install -r requirements.txt
```

### Execução  

```bash
python main.py
```

A interface gráfica será aberta com abas para **Clientes**, **Produtos**, **Orçamentos** e **Consultas**.

---

## Estrutura do Projeto  

```
sistema-orcamentos/
│
├── main.py               # Ponto de entrada da aplicação
├── requirements.txt      # Dependências do projeto
├── pedidos.db            # Banco SQLite (gerado automaticamente)
├── utils/                # Funções auxiliares
├── assets/               # Imagens e logotipos internos
├── arquivos_exportados/  # PDFs e planilhas .xlsx geradas
└── docs/
    └── images/           # Capturas de tela e documentação técnica
```

---

## Exemplos de Saída  

### Tela de Clientes  
![Tela de Clientes](https://github.com/user-attachments/assets/69945def-8a5c-46a0-bf45-ba9377eea2da)

### Tela de Produtos  
![Tela de Produtos](https://github.com/user-attachments/assets/64821873-58b2-491f-b243-265c946b724f)

### Tela de Orçamentos  
![Tela de Orçamentos](https://github.com/user-attachments/assets/40edfe8e-5ab7-4a93-af85-6f62e011f55f)

---

## Contribuição & Padrões  

- Seguir convenção **PEP8**;  
- Nomear commits conforme padrão: `feat/`, `fix/`, `docs/`, `refactor/`;  
- Alterações relevantes devem ser registradas no changelog;  
- Atualizar `version.json` antes de cada release interna.  

---

## Backup & Migração de Dados  

- O banco local `pedidos.db` deve ser incluído nos backups periódicos da estação;  
- Antes de atualizar versões, recomenda-se exportar os dados para Excel;  
- As migrações de estrutura (schema) devem ser documentadas no diretório `/docs/migrations/`.  

---

## Contatos / Suporte Interno  

**Responsável técnico:** Gabriel Araújo  
**Departamento:** TI – Eletrofrio  
**Status do projeto:** Em uso interno / manutenção contínua  
**Última atualização:** Outubro de 2025  
