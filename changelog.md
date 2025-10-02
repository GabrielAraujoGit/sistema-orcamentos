# Changelog

Todas as mudanças notáveis neste projeto serão documentadas aqui.

O formato segue o padrão [Keep a
Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

## \[Unreleased\]

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
