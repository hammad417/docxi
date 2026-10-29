<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Registro de alterações e histórico de versões do docxi

Esta página resume cada versão. As notas completas estão em [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md), no repositório, que é a referência caso os dois sejam diferentes.

## <a id="releases"></a>Versões

| Versão | Data | Notas |
|---|---|---|
| 0.1.1 | 2026-10-29 | Resumo e descrição da gem mais precisos, e link para a documentação. Sem mudanças no código |
| 0.1.0 | 2026-10-28 | Funciona com as versões atuais do Ruby, do Rails e do rubyzip. Gera os mesmos documentos que a 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Versão prévia da 0.1.0 |
| 0.0.6 | 2016-10-18 | Última versão antes da atualização 0.1.0 |
| 0.0.2 | 2016-10-14 | Nenhuma nota de versão foi registrada |
| 0.0.1 | 2016-10-14 | Primeira versão. Nenhuma nota de versão foi registrada |

As versões 0.0.3 a 0.0.5 nunca foram publicadas no RubyGems.

## <a id="011"></a>0.1.1

Sem mudanças no código, e os documentos gerados são os mesmos da 0.1.0.

- O resumo e a descrição da gem agora dizem o que o docxi faz. Antes o apresentavam como gerador de templates.
- O RubyGems aponta para esta wiki como documentação da gem.

## <a id="010"></a>0.1.0

**Compatibilidade**

- Exige Ruby 2.5 ou mais recente. Aplicações em versões mais antigas do Ruby continuam na 0.0.6.
- Funciona com ou sem Rails e é testada junto com o Rails 5.2 a 8.1.
- nokogiri `>= 1.8, < 2` em vez de `>= 1.6.3, < 1.7`, que entrava em conflito com todas as versões do Rails desde a 6.0.
- rubyzip `>= 1.3, < 4` em vez de `~> 0.9.9`.
- Não chama mais `blank?` e `present?` do ActiveSupport, que só funcionavam quando o Rails os tinha carregado.

**Correções**

- Documentos gravados com rubyzip 3 não contêm mais registros Zip64, que o Google Docs e versões mais antigas do LibreOffice rejeitam.
- `add_media` com uma URL de imagem volta a funcionar no Ruby 3.
- `add_media` não executa mais como comando de shell uma string que começa com `|`.
- `tab` e `frame` não lançam mais `NameError`.
- `document.frame("text")` cria uma caixa de texto com o texto.
- `page_numbers` em um cabeçalho não lança mais `NameError`.

**Gem**

- Mudança na manutenção e novos links para a página inicial e o registro de alterações.
- As publicações exigem MFA.

## <a id="related-pages"></a>Páginas relacionadas

- [Roteiro](Roadmap)
- [Instalação](Installation#upgrading-from-00x)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
