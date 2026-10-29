<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Ambiente de desenvolvimento: rodar os testes e as ferramentas do docxi

Esta página ajuda você a colocar uma cópia de trabalho do docxi para rodar e explica os testes, os documentos de exemplo e a configuração da CI.

## <a id="get-the-code"></a>Obter o código

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Qualquer Ruby a partir da 2.5 serve para o desenvolvimento. As dependências de desenvolvimento são minitest e rake.

## <a id="rake-tasks"></a>Tarefas do Rake

| Comando | O que faz |
|---|---|
| `bundle exec rake test` | Roda a suíte de testes (a tarefa padrão) |
| `bundle exec rake samples` | Grava os documentos de exemplo em `tmp/samples/*.docx`, para abrir no Word, no LibreOffice ou no Google Docs |
| `bundle exec rake golden` | Regrava `test/golden` a partir do código atual. Use somente depois de uma mudança intencional na saída |
| `bundle exec rake build` | Gera a gem em `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Rodar os testes com o Rails carregado

A pasta `gemfiles/` tem um Gemfile por versão do Rails, além de alguns para versões mais antigas do rubyzip e do nokogiri:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` carrega o framework inteiro antes dos testes. Ele também compara as regras de `blank?` do docxi com as do ActiveSupport.

## <a id="what-the-tests-cover"></a>O que os testes cobrem

| Arquivo | Verifica |
|---|---|
| `test/test_golden_output.rb` | Cada parte de dois documentos de exemplo é igual ao que a 0.0.6 gerava |
| `test/test_package.rb` | O XML é bem formado, os relacionamentos apontam para partes existentes, cada parte tem um tipo de conteúdo e não há registros Zip64 |
| `test/test_media.rb` | Imagens a partir de caminhos, objetos IO e URLs, e que uma string com `\|` nunca é executada como comando |
| `test/test_fixed_features.rb` | `tab`, `frame` e os números de página nos cabeçalhos |
| `test/test_util.rb` | `Docxi::Util.blank?` segue as regras do ActiveSupport |
| `test/test_rails_independence.rb` | O docxi funciona sem o ActiveSupport |
| `test/test_warnings.rb` | A geração não emite avisos de descontinuação a partir de `lib/` |

Os documentos de exemplo ficam em `test/support/sample_documents.rb`, e `test/support/package.rb` lê um `.docx` direto das estruturas do zip.

## <a id="continuous-integration"></a>Integração contínua

`.github/workflows/ci.yml` roda em cada push para `master` e em cada pull request. Ele cobre:

- Ruby 2.5, 2.6 e 4.0 sem Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) e 8.1 (4.0)
- Rails `main`, que pode falhar
- rubyzip 1.3 e 2.4, e nokogiri 1.11
- um job que converte os documentos de exemplo para PDF com o LibreOffice

A publicação de versões está descrita em [Implantação](Deployment).

## <a id="related-pages"></a>Páginas relacionadas

- [Como contribuir](Contributing)
- [Arquitetura](Architecture)
- [Implantação](Deployment)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
