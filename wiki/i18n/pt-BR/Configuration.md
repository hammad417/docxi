<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Configurar o docxi: margens, configuração de página e unidades

O docxi não tem arquivo de configuração, variáveis de ambiente nem configurações globais. Tudo é definido por documento ou por elemento. Esta página cobre as configurações no nível do documento e as unidades usadas pelas opções. As opções de cada elemento estão em [Recursos](Features).

## <a id="page-margins"></a>Margens da página

As margens são passadas na criação do documento:

```ruby
require "docxi"

# Todos os valores estão em vigésimos de ponto (1440 = 1 polegada).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Opção | Padrão | Significado |
|---|---|---|
| `top` | `720` (0,5 pol.) | Margem superior |
| `right` | `1440` (1 pol.) | Margem direita |
| `bottom` | `720` (0,5 pol.) | Margem inferior |
| `left` | `1440` (1 pol.) | Margem esquerda |
| `header` | `190` | Distância da borda superior da página até o cabeçalho |
| `footer` | `200` | Distância da borda inferior da página até o rodapé |

O hash que você passa substitui os valores padrão, então defina `top`, `right`, `bottom` e `left` juntos. Qualquer uma dessas quatro margens que você omitir é gravada com valor vazio. `header` e `footer` voltam para 190 e 200 quando estão ausentes ou vazios.

## <a id="page-size-and-orientation"></a>Tamanho e orientação da página

As páginas são US Letter em retrato: 12240 × 15840 vigésimos de ponto, ou seja, 8,5 × 11 polegadas. O tamanho e a orientação não podem ser alterados na 0.1.0.

## <a id="first-page"></a>Primeira página

O docxi sempre ativa a opção "Primeira página diferente" do Word. Os cabeçalhos e rodapés adicionados com `add_header` e `add_footer` são usados a partir da página 2. A página 1 usa `add_first_header` e `add_first_footer`, então defina esses também se a primeira página precisar de cabeçalho ou rodapé. Veja [Recursos](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>O que todo documento contém

- A fonte padrão é Arial 12 pt.
- Estilos definidos: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` e `NoList`.
- Os itens de lista com marcadores (`ul: true`) usam um traço como marcador.
- As propriedades do documento informam "Docxi Ruby Gem" como nome do aplicativo e registram a hora em que o documento foi gerado.

## <a id="units"></a>Unidades

| Opção | Unidade |
|---|---|
| `size:` do texto | Pontos (o docxi dobra o valor para os meios pontos do Word) |
| Margens, `left:` e `right:` do parágrafo, `iwidth:` da tabela, `height:` da linha | Vigésimos de ponto (1440 por polegada) |
| Espessuras de `borders:` da célula | Oitavos de ponto |
| `columns_width:` e `width:` da célula | Multiplicados por 14.92 para obter vigésimos de ponto, próximo de pixels a 96 dpi |
| `style:` da imagem embutida | Medidas CSS como `"120pt"` |
| `pH:` e `pV:` da imagem posicionada | EMU (914400 por polegada) |
| Cores de texto e de preenchimento | Hexadecimal sem `#`, por exemplo `"8B0000"` |
| `fill_color:` da imagem | Cor VML com `#`, por exemplo `"#00FF00"` |

## <a id="related-pages"></a>Páginas relacionadas

- [Recursos](Features)
- [Referência da API](API-Reference)
- [Boas práticas](Best-Practices)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
