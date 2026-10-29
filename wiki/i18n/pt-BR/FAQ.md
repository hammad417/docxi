<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# Perguntas frequentes sobre o docxi: documentos do Word em Ruby e Rails

Respostas curtas para as perguntas mais comuns. Cada resposta tem um link para uma explicação mais longa.

**Preciso ter o Microsoft Word ou o LibreOffice instalado?**
Não. O docxi grava o XML e o arquivo zip por conta própria.

**O docxi exige Rails?**
Não. Ele funciona em qualquer programa Ruby. Os testes também rodam com o Rails 5.2 a 8.1 carregado. Veja [Integrações](Integrations).

**Como gerar um documento do Word no Ruby on Rails?**
Monte o documento em um controller ou em um job em segundo plano e envie `word.render.string` com `send_data`. Veja [Integrações](Integrations).

**Quais versões do Ruby e do Rails são suportadas?**
Ruby 2.5 e mais recentes. O Rails é opcional, e a CI testa o Rails 5.2 a 8.1. Veja [Instalação](Installation).

**O docxi consegue editar um .docx existente ou preencher um modelo do Word?**
Não, ele só cria documentos novos. Para modelos com campos de mesclagem, use uma gem baseada em modelos, como a sablon.

**O sumário é gerado a partir dos títulos?**
Não. Você adiciona cada entrada e o número de página dela. Veja [Recursos](Features#table-of-contents).

**Posso usar listas numeradas ou estilos de título?**
Não na 0.1.0. Os parágrafos podem ser itens com marcadores (`ul: true`), e você pode formatar títulos com opções de texto como `size:` e `bold:`.

**Posso mudar o tamanho ou a orientação da página?**
Ainda não. As páginas são US Letter em retrato, e só as margens podem ser alteradas. Veja [Configuração](Configuration).

**Quais formatos de imagem funcionam?**
PNG. Toda imagem é armazenada como uma parte PNG.

**Posso adicionar uma imagem a partir de uma URL?**
Sim. `add_media` aceita URLs `http`, `https` e `ftp` e as baixa quando você chama `render`. Não passe URLs vindas de usuários sem verificá-las. Veja [Segurança](Security).

**Posso gravar o documento direto em um arquivo?**
`render` retorna um `StringIO`. Grave o `.string` dele com `File.binwrite("out.docx", word.render.string)`.

**O docxi consegue gerar documentos muito grandes por streaming?**
Não. O documento inteiro é montado na memória quando você chama `render`. Veja [Implantação](Deployment).

**Posso definir parágrafos da direita para a esquerda?**
A 0.1.0 não tem opções de direita para a esquerda. O texto é gravado do jeito que você o passa.

**Onde estão listados os problemas conhecidos?**
No [roteiro](Roadmap) e na seção "Limitations and known issues" (limitações e problemas conhecidos) do README.

## <a id="related-pages"></a>Páginas relacionadas

- [Solução de problemas](Troubleshooting)
- [Recursos](Features)
- [Primeiros passos](Getting-Started)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
