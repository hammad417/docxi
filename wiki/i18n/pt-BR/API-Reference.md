<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Referência da API do docxi: classes e métodos

O docxi é uma biblioteca Ruby, então esta página lista a API Ruby dele. Não existe API HTTP. As opções de cada método estão descritas em [Recursos](Features), e as classes são explicadas em [Arquitetura](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Método | Retorna | Observações |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` são as margens da página. Veja [Configuração](Configuration) |
| `#document` | `Docxi::Word::Document` | O corpo ao qual você adiciona conteúdo |
| `#render` | `StringIO` | O pacote `.docx` completo. `.string` retorna os bytes |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Método | Retorna | Adiciona |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Um parágrafo |
| `text(string, options = {})` | Paragraph | Um parágrafo com um único trecho de texto |
| `br(page: true)` | Break | Uma quebra de página |
| `tab(page: true)` | Tab | Um parágrafo com uma tabulação |
| `table(options = {}) { \|table\| }` | Table | Uma tabela |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Entradas do sumário |
| `frame(text, options = {})` | Paragraph | Um parágrafo com uma caixa de texto |
| `add_media(io_path_or_url, options = {})` | Media | Uma imagem para posicionar com `image` |
| `add_hyperlink(url, options = {})` | String | Um id de relacionamento para `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Cabeçalho da página 2 em diante e da página 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Rodapé da página 2 em diante e da página 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Parágrafo (`Docxi::Word::Contents::Paragraph`)

| Método | Adiciona |
|---|---|
| `text(string, options = {})` | Um trecho de texto. As opções do parágrafo são o padrão dele |
| `br(times: 1)` | Quebras de linha |
| `tab(times: 1)` | Caracteres de tabulação |
| `image(media, options = {})` | Uma imagem embutida ou posicionada |
| `hyperlink(text, url, id)` | Um link que usa o `id` de relacionamento retornado por `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Uma caixa de texto |

## <a id="tables"></a>Tabelas

| Método | Adiciona |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Uma linha. As opções da tabela são o padrão dela |
| `TableRow#tc(options = {}) { \|tc\| }` | Uma célula. As opções da linha são o padrão dela |
| `TableCell#text(string, options = {})` | Um parágrafo. As opções da célula são o padrão dele |
| `TableCell#p(options = {}) { \|p\| }` | Um parágrafo com as próprias opções |
| `TableCell#image(media, width:, height:)` | Uma imagem |
| `TableCell#page_numbers` | Um campo de número de página |

## <a id="table-of-contents"></a>Sumário

| Método | Adiciona |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Uma entrada do sumário |

## <a id="headers-and-footers"></a>Cabeçalhos e rodapés

| Classe | Métodos |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Quadros

`Docxi::Word::Contents::Frame` aceita `text`, `br`, `tab`, `image` e `frame` dentro do bloco.

## <a id="constants-and-helpers"></a>Constantes e métodos auxiliares

| Nome | Significado |
|---|---|
| `Docxi::VERSION` | A versão instalada, por exemplo `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Grava uma parte no arquivo compactado com o tamanho definido antecipadamente. Usado internamente por todas as partes |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Métodos auxiliares internos com as regras de `blank?` do ActiveSupport |

## <a id="related-pages"></a>Páginas relacionadas

- [Recursos](Features)
- [Exemplos de uso](Usage-Examples)
- [Arquitetura](Architecture)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
