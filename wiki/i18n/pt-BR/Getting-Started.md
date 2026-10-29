<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Primeiros passos com o docxi: seu primeiro documento do Word em Ruby

Este guia leva você de uma pasta vazia até um arquivo `.docx` formatado. Ele parte do princípio de que você conhece o básico de Ruby. Você não precisa de Rails nem do Microsoft Word.

## <a id="1-install-the-gem"></a>1. Instale a gem

```sh
gem install docxi
```

Em um projeto com Gemfile, adicione `gem "docxi", "~> 0.1"` e execute `bundle install` em vez disso. A página [Instalação](Installation) traz os detalhes.

## <a id="2-create-a-document"></a>2. Crie um documento

Salve isto como `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Execute `ruby hello.rb` (ou `bundle exec ruby hello.rb` em um projeto com Bundler) e abra `hello.docx`.

## <a id="3-what-each-line-does"></a>3. O que cada linha faz

- `Docxi::Document.new` cria um documento vazio com as margens de página padrão.
- `word.document` é o corpo do documento. Tudo o que você adiciona passa por ele.
- `document.p` adiciona um parágrafo. As opções passadas para `p`, como `align:`, valem para o parágrafo e para o texto dentro dele.
- `p.text` adiciona um trecho de texto (um run). `size:` é dado em pontos.
- `word.render` monta o pacote `.docx` e retorna um `StringIO`. O `.string` dele contém os bytes do arquivo.

## <a id="4-add-a-list-and-a-table"></a>4. Adicione uma lista e uma tabela

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

Dois pontos de atenção:

- `space: true` mantém o espaço no início ou no fim de um trecho. Sem essa opção, o Word descarta o espaço e as palavras ficam grudadas.
- As opções de `tr` valem para todas as células daquela linha, por isso a linha de cabeçalho inteira fica sombreada e em negrito.

## <a id="5-where-to-go-next"></a>5. Próximos passos

- [Exemplos de uso](Usage-Examples): receitas completas para relatórios, faturas e cartas
- [Recursos](Features): todas as opções de parágrafos, texto, tabelas, imagens, cabeçalhos e rodapés
- [Integrações](Integrations): enviar o documento a partir de um controller do Rails
- [Boas práticas](Best-Practices): evite as surpresas mais comuns antes que aconteçam

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
