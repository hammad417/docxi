<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Recursos do docxi: parágrafos, texto, tabelas, imagens, cabeçalhos e rodapés

Esta página lista todos os elementos que o docxi pode adicionar a um documento do Word e as opções que cada um aceita. Os exemplos partem desta configuração:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

As opções de que você não precisa podem ser omitidas. As cores são valores hexadecimais sem `#`, por exemplo `"8B0000"`. Para unidades e configuração de página, veja [Configuração](Configuration).

## <a id="paragraphs"></a>Parágrafos

`document.p(options) { |p| ... }` adiciona um parágrafo.

| Opção | Valores | Efeito |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Alinhamento. `"both"` justifica o texto |
| `fill` | Cor hexadecimal | Sombreamento atrás do parágrafo e do texto |
| `left`, `right` | Vigésimos de ponto | Recuo. Defina os dois, porque um sozinho é ignorado |
| `ul` | `true` | Transforma o parágrafo em um item de lista com marcadores |
| `bottom` | `true` | Desenha uma borda abaixo do parágrafo |

Opções de texto como `size:`, `bold:` ou `color:` também podem ser passadas para `p`. Elas se tornam o padrão de cada trecho de texto do parágrafo.

Dentro do bloco, você pode chamar `text`, `br`, `tab`, `image`, `hyperlink` e `frame`. `document.text(string, options)` é um atalho para um parágrafo com um único trecho.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Texto

`p.text(string, options)` adiciona um trecho de texto (run).

| Opção | Valores | Efeito |
|---|---|---|
| `bold` | `true` | Negrito |
| `italic` | `true` | Itálico |
| `underline` | `true` | Sublinhado simples |
| `color` | Cor hexadecimal | Cor do texto |
| `size` | Pontos | Tamanho da fonte |
| `font` | Nome da fonte | Fonte, por exemplo `"Georgia"` |
| `fill` | Cor hexadecimal | Sombreamento atrás do texto |
| `space` | `true` | Mantém os espaços no início e no fim do texto |
| `br` | `true` | Adiciona uma quebra de linha depois do texto |

Uma string `nil` é ignorada.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Quebras de linha, quebras de página e tabulações

- `p.br`, ou `p.br(times: 3)`, adiciona quebras de linha dentro de um parágrafo.
- `document.br(page: true)` inicia uma nova página.
- `p.tab`, ou `p.tab(times: 2)`, adiciona caracteres de tabulação dentro de um parágrafo.
- `document.tab(page: true)` adiciona um parágrafo com uma única tabulação.

Não chame `document.br` nem `document.tab` sem `page: true`. Eles colocam um trecho de texto diretamente no corpo do documento, o que não é WordprocessingML válido.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tabelas

`document.table(options)` recebe um bloco. Dentro dele, `table.tr(options)` adiciona linhas e `tr.tc(options)` adiciona células.

**Opções da tabela**

| Opção | Valores | Efeito |
|---|---|---|
| `columns_width` | Array de números | Larguras das colunas, multiplicadas por 14.92 para vigésimos de ponto. Também são as larguras padrão das células |
| `height` | Vigésimos de ponto | Altura mínima de cada linha |
| `iwidth` | Vigésimos de ponto | Recuo da tabela a partir da margem esquerda |
| `style` | Id do estilo | Estilo da tabela. O docxi define `TableNormal` e `TableCustom` |
| `width` | Número | Largura preferencial da tabela, gravada com o tipo `auto`. Sem `columns_width`, cada célula também recebe essa largura, multiplicada por 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Alinhamento da tabela. O texto adicionado com `tc.text` é alinhado da mesma forma |
| `padding` | `true` | Adiciona um espaçamento interno fixo nas células |
| `borders` | Hash | Problema conhecido: é aplicada a cada célula, e não ao contorno da tabela |

**Opções da linha**: `height:` define a altura mínima da linha. Qualquer opção de célula passada para `tr` vale para todas as células da linha. As opções da tabela são repassadas da mesma forma para cada linha e célula.

**Opções da célula**

| Opção | Valores | Efeito |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Espessura da borda em oitavos de ponto. `nil` significa sem borda naquele lado |
| `fill` | Cor hexadecimal | Sombreamento da célula. Células sem preenchimento são brancas (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Alinhamento vertical |
| `colspan` | Inteiro | Número de colunas que a célula ocupa |
| `rowspan` | `true` | Inicia uma mesclagem vertical |
| `merged` | `true` | Continua a mesclagem vertical da célula de cima. O conteúdo dela é ignorado |
| `width` | Número | Largura da célula, multiplicada por 14.92 |

**Conteúdo da célula**

- `tc.text(string, options)`: as opções da célula e da linha, como `bold:` ou `align:`, também valem para o texto.
- `tc.p(options) { |p| ... }`: um parágrafo apenas com as próprias opções.
- `tc.image(media, width:, height:)`: uma imagem. Veja o problema de tamanho em [Imagens](#images).
- `tc.page_numbers`: um campo de número de página. Problema conhecido: ele vem precedido do texto "GlobalOptions".

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Imagens

`media = document.add_media(source)` registra uma imagem. `source` pode ser:

- um objeto IO, como um `File` ou um `StringIO`
- um caminho de arquivo
- uma URL `http`, `https` ou `ftp`

A imagem é lida quando você chama `render` e é sempre armazenada como uma parte PNG, então arquivos PNG são a escolha segura.

**Imagens embutidas**: `p.image(media, options)` coloca a imagem na linha de texto.

| Opção | Valores | Efeito |
|---|---|---|
| `style` | Hash de propriedades CSS | Tamanho, por exemplo `{ width: "120pt", height: "60pt" }` |
| `title` | String | Título da imagem |
| `fill_color` | Cor VML como `"#00FF00"` | Preenchimento atrás da imagem |
| `opacity` | String como `"50%"` | Opacidade do preenchimento |
| `wrap` | `true` | Desativa a quebra automática de texto em volta da imagem |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Imagens posicionadas**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` ancora a imagem com um deslocamento em relação à coluna (`pH`) e ao parágrafo (`pV`), em EMU. `pH`, `pV` e `wrap` precisam ser informados juntos; caso contrário, a imagem fica embutida.

**Imagens de cabeçalho**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` desativa a quebra automática. Qualquer outro valor faz o texto contornar a imagem em formato quadrado.

**Problema conhecido na 0.1.0**: imagens posicionadas, de cabeçalho e de células de tabela calculam a largura e a altura como `width * height * 14.92` EMU. Elas sempre saem quadradas e geralmente minúsculas. Imagens embutidas dimensionadas com `style:` não são afetadas. Veja [Boas práticas](Best-Practices#images) para uma solução alternativa.

## <a id="headers-and-footers"></a>Cabeçalhos e rodapés

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Um cabeçalho pode conter `text`, `br`, `table`, `image` e `page_numbers`. Um rodapé pode conter `text`, `br`, `table` e `page_numbers`.
- As opções passadas para `Header.new` ou `Footer.new` valem para o texto deles.
- `add_header` e `add_footer` valem da página 2 em diante. A página 1 sempre usa `add_first_header` e `add_first_footer`. Veja [Configuração](Configuration#first-page).

## <a id="page-numbers"></a>Números de página

`page_numbers(name: "Page", align: "center")` grava o texto de `name` seguido de um campo PAGE. O número sai em Arial 10 pt, cinza-escuro. `align` tem `"right"` como padrão.

Problema conhecido: o código de campo é gravado sem a barra invertida de `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Sumário

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Com um número de página, a entrada recebe uma parada de tabulação alinhada à direita em `pos` (vigésimos de ponto; 9360 equivale a 6,5 polegadas), preenchida com o caractere de `leader`. Os valores de preenchimento vêm do WordprocessingML, como `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` e `"none"`.
- Sem número de página, a entrada é o texto depois de uma tabulação.
- Os números de página são os que você informa. O docxi não cria um campo TOC do Word nem calcula páginas.

## <a id="hyperlinks"></a>Hiperlinks

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` retorna um id de relacionamento, e `hyperlink(text, url, id)` insere o link. O destino vem do relacionamento, então passe o id retornado para a URL desejada. Somente links externos são suportados.

## <a id="frames-text-boxes"></a>Quadros (caixas de texto)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Os quadros são caixas de texto VML com preenchimento fixo cinza-claro semitransparente e contorno preto. O texto fica acima e abaixo deles.
- Dentro de um quadro, você pode chamar `text`, `br`, `tab`, `image` e `frame`.
- `spacing: true` adiciona espaço antes do texto.

## <a id="output"></a>Saída

`word.render` retorna um `StringIO` com o pacote `.docx` completo. O arquivo compactado é gravado sem registros Zip64, então o Google Docs e versões mais antigas do LibreOffice conseguem abri-lo. A página [Arquitetura](Architecture) lista cada parte do pacote.

## <a id="related-pages"></a>Páginas relacionadas

- [Exemplos de uso](Usage-Examples)
- [Referência da API](API-Reference)
- [Roteiro](Roadmap), com os problemas conhecidos previstos para a 0.2.0

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
