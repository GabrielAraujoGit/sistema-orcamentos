# Changelog

Todas as mudanças notáveis neste projeto serão documentadas aqui.

O formato segue o padrão [Keep a
Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

## \[Unreleased\]

## [1.4.1] - 09/Out/25

### Changed

* **Removida a exibição dos IDs** nas abas **Empresas** e **Produtos**, mantendo-os apenas para uso interno.
* **`carregar_empresas()`** e **`editar_empresa()`** ajustados para uso de `tags` (armazenando o ID de forma oculta).
* **`carregar_produtos()`** atualizado para alinhar corretamente as colunas e remover o campo `id` do `SELECT`.
* Layout mais limpo e consistente entre todas as abas do sistema.

### Fixed

* Corrigido desalinhamento de colunas causado pela presença do ID oculto no `SELECT`.
* Corrigido bug menor ao abrir o formulário de edição de empresa quando a coluna ID era removida da tabela.

---

## [1.4.0] - 07/Out/25

### Added
- **Cadastro de Empresas Emissoras**:
  - Nova aba para gerenciar empresas vinculadas aos orçamentos.
  - Campos de nome, CNPJ, endereço, e-mail e telefone.
  - Upload de **logos em PNG** com pré-visualização na interface.
  - Armazenamento automático da logo na pasta `/logos`.

- **Integração de Empresas com Orçamentos**:
  - Cada orçamento agora pode estar vinculado a uma empresa emissora.
  - Dados da empresa (nome, CNPJ, logo, etc.) são exibidos no PDF gerado.

- **Diferenciação visual entre “Novo” e “Edição de Orçamento”**:
  - Faixa colorida no topo da aba indicando o modo atual:
    - 🟢 Verde para “Novo Orçamento”
    - 🟠 Laranja para “Edição de Orçamento”
  - Botão principal muda texto e cor conforme o modo.
  - Campos de **Cliente** e **Empresa** são bloqueados durante edição.

- **Layout PDF Padronizado (multiempresa)**:
  - Cabeçalho fixo com logo e informações da empresa emissora.
  - Estrutura universal para todas as empresas do grupo.
  - Melhor espaçamento, margens consistentes e total destacado.
  - Nome de arquivo padronizado com data e hora (`orcamento_<num>_<data>.pdf`).

### Changed
- Função `gerar_pdf_orcamento` totalmente revisada para suportar múltiplas empresas.
- Imports reorganizados (evitando conflitos entre `Image` do Pillow, ReportLab e OpenPyXL).
- Melhor tratamento para campos ausentes (exibe “–” quando não há dados).
- Botão de salvar/atualizar orçamento agora reflete automaticamente o modo ativo.
- Títulos e rótulos atualizados para maior clareza visual.

### Fixed
- Corrigido erro ao gerar PDF com empresa sem logo.
- Corrigido bug em `finalizar_pedido` que não atualizava interface após salvar.
- Ajustada proporção de logos no PDF (largura fixa, altura proporcional).

---



## [1.3.1] - 06/Out/25

### Added
- Centralização visual aplicada a todas as tabelas (`Treeview`) do sistema.

### Changed
- Colunas das abas **Clientes**, **Produtos**, **Consultar Orçamentos**, **Itens do Orçamento** e **Visualização de Orçamento (popup)** agora exibem o conteúdo centralizado.
- Melhoria geral na legibilidade e alinhamento das informações nas tabelas.

---

## \[1.3.0\] - 02/Out/25

### Added

-   Nova janela **popup** para cadastro/edição de clientes (mais
    intuitiva).

### Changed

-   Aba **Clientes** simplificada:
    -   Removido formulário fixo acima da tabela.
    -   Mantidos apenas os botões de ação no topo (**Adicionar**,
        **Editar**, **Excluir**, **Importar Arquivo**) e a lista de
        clientes abaixo.
    -   Estilização dos botões com `ttkbootstrap` (`success`, `info`,
        `danger`, `warning`).
-   `editar_cliente` atualizado para abrir o formulário popup, sem
    depender de `self.cliente_entries`.

### Removed

-   Formulário embutido de clientes na aba principal (substituído por
    popup).

------------------------------------------------------------------------

## \[1.2.0\] - 02/Out/25

### Added

-   Botão de **Exportar PDF** na aba de Orçamentos.

### Changed

-   Interface da aba **Orçamentos** modernizada com `ttkbootstrap`
    (layout mais moderno e intuitivo).
-   Textos dos botões padronizados para ficarem mais claros para o
    usuário.
-   Melhorias no fluxo de edição de orçamentos.

### Removed

-   Geração automática de PDF ao salvar/atualizar orçamentos (removida
    para evitar criação de múltiplos arquivos desnecessários a cada
    alteração).

------------------------------------------------------------------------

## \[1.1.0\] - 01/Out/25

### Added

-   Novos filtros para consulta de orçamentos.
-   Ajustes de tela e alterações no banco de dados.

### Changed

-   Refinamentos na exportação e relatórios.
-   Ajustes finais em telas e lógicas.

------------------------------------------------------------------------

## \[1.0.1\] - 30/Set/25

### Added

-   Função de importar produtos a partir de arquivo.

### Changed

-   Customização inicial dos PDFs.
-   Correções gerais no sistema.

------------------------------------------------------------------------

## \[1.0.0\] - 29/Set/25

### Added

-   Versão inicial do **Sistema de Orçamentos** com:
    -   Cadastro de clientes e produtos.
    -   Cadastro e edição de orçamentos.
    -   Exportação de orçamentos para Excel.
    -   Geração de PDF simples.
