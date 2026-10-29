<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Instalar o docxi: requisitos, configuração do Bundler e atualização

O docxi é uma gem feita só em Ruby. A instalação traz apenas duas dependências, nokogiri e rubyzip, e nada mais. Nenhum pacote de escritório é necessário.

## <a id="requirements"></a>Requisitos

| Requisito | Versão |
|---|---|
| Ruby | 2.5 ou mais recente |
| nokogiri | `>= 1.8, < 2`, instalada automaticamente |
| rubyzip | `>= 1.3, < 4`, instalada automaticamente |
| Rails | Não é necessário |

## <a id="install-with-bundler"></a>Instalar com o Bundler

Adicione a gem ao seu Gemfile:

```ruby
gem "docxi", "~> 0.1"
```

Depois instale:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Instalar sem o Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Verificar a instalação

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Em um projeto com Bundler, execute como `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. O comando mostra a versão instalada, por exemplo `0.1.0`.

## <a id="upgrading-from-00x"></a>Atualizar a partir da 0.0.x

Se o seu Gemfile fixa o docxi em uma 0.0.x, por exemplo com `"~> 0.0.6"`, troque a linha por `gem "docxi", "~> 0.1"`. Depois execute `bundle update docxi`. Não é preciso mudar o código, e a 0.1.0 gera os mesmos documentos que a 0.0.6: a suíte de testes compara a saída com arquivos gerados pela 0.0.6.

O que mudou por baixo dos panos:

- nokogiri e rubyzip agora podem estar em qualquer versão dentro das faixas acima, então o docxi pode ser instalado junto com as versões atuais do Rails.
- O docxi não precisa mais que o ActiveSupport esteja carregado.
- Os arquivos compactados são gravados sem registros Zip64.
- As imagens são carregadas sem `Kernel#open`, então URLs funcionam no Ruby 3.
- `tab`, `frame` e os números de página nos cabeçalhos funcionam.

Aplicações com uma versão do Ruby anterior à 2.5 continuam recebendo a 0.0.6, porque a 0.1.0 declara o Ruby 2.5 como mínimo. O [registro de alterações](Changelog) lista todas as mudanças.

## <a id="versions-tested-in-ci"></a>Versões testadas na CI

Cada versão do Rails roda com o framework completo carregado junto com o docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

A CI também executa:

- os testes sem Rails no Ruby 2.5, 2.6 e 4.0
- os testes com rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) e nokogiri 1.11 (Ruby 2.5)
- uma verificação de que o LibreOffice consegue converter os documentos de exemplo

## <a id="platform-notes"></a>Observações sobre plataformas

- O nokogiri distribui gems pré-compiladas para as plataformas Linux, macOS e Windows mais comuns. Quando nenhuma corresponde, ele é compilado a partir do código-fonte e precisa de um compilador C. Veja o [guia de instalação do nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- A CI roda no Ubuntu 24.04, e os testes também já foram executados no macOS. O Windows não é testado.

## <a id="if-installation-fails"></a>Se a instalação falhar

Conflitos do Bundler envolvendo nokogiri ou rubyzip quase sempre significam que um docxi 0.0.6 antigo está sendo resolvido. Veja [Solução de problemas](Troubleshooting#installation).

## <a id="related-pages"></a>Páginas relacionadas

- [Primeiros passos](Getting-Started)
- [Configuração](Configuration)
- [Implantação](Deployment)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
