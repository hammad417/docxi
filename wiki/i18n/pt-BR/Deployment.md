<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# O docxi em produção e a publicação de versões

O docxi é uma biblioteca, não um serviço, então não há nada a implantar por conta própria. Esta página trata do que importa quando você coloca no ar uma aplicação que usa o docxi e de como as novas versões do docxi são publicadas.

## <a id="shipping-an-app-that-uses-docxi"></a>Publicar uma aplicação que usa o docxi

- **Dependências.** `bundle install` instala o nokogiri e o rubyzip. O nokogiri distribui gems pré-compiladas para as plataformas Linux, macOS e Windows mais comuns. Em outras plataformas, como algumas imagens de contêiner mínimas, ele é compilado a partir do código-fonte e precisa de um compilador C. Veja o [guia de instalação do nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Sem pacote de escritório.** Servidores e contêineres não precisam do Microsoft Word nem do LibreOffice.
- **Memória.** `render` monta o documento inteiro na memória e o retorna como um `StringIO`. O uso de memória cresce com o número de páginas e o tamanho das imagens.
- **Rede e arquivos.** O docxi só lê o sistema de arquivos para os caminhos de imagem que você passa para `add_media`, e só usa a rede para as URLs de imagem que você passa. Veja [Segurança](Security).
- **Documentos longos.** Gere-os em um job em segundo plano, e não durante uma requisição web. Veja [Integrações](Integrations).

## <a id="how-docxi-releases-are-published"></a>Como as versões do docxi são publicadas

Estes passos são para quem mantém o projeto. As versões vão para o RubyGems pelo GitHub Actions com publicação confiável (trusted publishing), então nenhuma chave de API fica guardada no repositório.

1. Atualize `lib/docxi/version.rb` e `CHANGELOG.md` e faça o merge em `master`.
2. Crie uma tag no commit com a versão prefixada por `v`, por exemplo `v0.1.1`, e envie a tag com push.
3. `.github/workflows/release.yml` gera a gem e a publica no RubyGems.

Versões prévias seguem os mesmos passos, com uma versão como `0.2.0.rc1`. O Bundler só instala versões prévias quando elas são pedidas explicitamente. A CI (`.github/workflows/ci.yml`) roda em cada push para `master` e em cada pull request.

## <a id="related-pages"></a>Páginas relacionadas

- [Instalação](Installation)
- [Ambiente de desenvolvimento](Development-Setup)
- [Registro de alterações](Changelog)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
