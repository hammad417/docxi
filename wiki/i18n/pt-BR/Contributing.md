<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Como contribuir com o docxi

Relatos de bugs, correções, documentação e traduções são todos bem-vindos. Esta página explica como relatar um problema e o que um pull request precisa ter. Para preparar uma cópia de trabalho, veja [Ambiente de desenvolvimento](Development-Setup).

## <a id="reporting-a-bug"></a>Relatar um bug

Abra uma issue em [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) com:

- um script Ruby curto que reproduza o problema
- as versões do Ruby, do docxi, do nokogiri e do rubyzip (e do Rails, se você o usa)
- o que você esperava e o que obteve. Para documentos com defeito, diga em qual programa você os abriu

Relate problemas de segurança em particular. Veja [Segurança](Security).

## <a id="pull-requests"></a>Pull requests

1. Faça um fork do repositório e crie um branch.
2. Faça a alteração seguindo o estilo do código ao redor.
3. Adicione ou atualize os testes e execute `bundle exec rake test`.
4. Se o XML gerado mudar de propósito, execute `bundle exec rake golden` e explique a mudança no pull request. Os arquivos de referência protegem a saída da qual os usuários atuais dependem.
5. Atualize o `CHANGELOG.md` e a documentação se o comportamento mudar.
6. Abra o pull request. A CI roda os testes no Ruby 2.5 a 4.0 e no Rails 5.2 a 8.1, e verifica se o LibreOffice abre os documentos de exemplo.

Mudanças que alteram documentos existentes, como as correções previstas para a 0.2.0, devem ir em um pull request próprio, para que a diferença nos arquivos de referência seja fácil de revisar.

## <a id="documentation-and-translations"></a>Documentação e traduções

- O README em inglês é a fonte de 15 traduções em `docs/i18n/`. Siga o `docs/i18n/TRANSLATION_GUIDE.md` do repositório ao alterá-lo.
- Esta wiki é escrita na pasta `wiki/` do repositório e publicada a partir dela. O `wiki/README.md` explica como.

## <a id="related-pages"></a>Páginas relacionadas

- [Ambiente de desenvolvimento](Development-Setup)
- [Arquitetura](Architecture)
- [Roteiro](Roadmap)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
