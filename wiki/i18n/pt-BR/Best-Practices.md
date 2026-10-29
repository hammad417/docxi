<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Boas práticas com o docxi: evite as surpresas mais comuns

Estes hábitos vêm direto da forma como o docxi funciona, e a maioria evita algum problema listado em [Solução de problemas](Troubleshooting). Eles estão agrupados por tema. As dicas de configuração de página valem para todos, e as de segurança e desempenho importam mais para aplicações em produção.

## <a id="page-setup"></a>Configuração de página

- **Passe as quatro margens juntas.** O hash que você passa para `Docxi::Document.new` substitui os valores padrão, e qualquer um de `top`, `right`, `bottom` e `left` que você omitir é gravado vazio.
- **Defina o cabeçalho e o rodapé da primeira página.** O docxi sempre ativa uma primeira página separada. Se a página 1 precisar de cabeçalho ou rodapé, chame também `add_first_header` e `add_first_footer`.

## <a id="text"></a>Texto

- **Use `space: true` para espaços no início ou no fim.** Caso contrário, o Word os descarta e as palavras de dois trechos se juntam.
- **Coloque a formatação compartilhada no parágrafo.** As opções passadas para `p`, como `size:` ou `font:`, valem para todos os trechos dentro dele, então há menos coisa para repetir.
- **Use `p.br` para quebras de linha e `document.br(page: true)` para quebras de página.** `document.br` sem `page: true` produz XML inválido.

## <a id="tables"></a>Tabelas

- **Defina as bordas nas linhas ou nas células, não na tabela.** Na 0.1.0, `borders:` no nível da tabela é aplicada a cada célula, e não ao contorno da tabela.
- **Informe `columns_width` para todas as colunas.** As células tiram a largura dali, e as tabelas mantêm o layout esperado.

## <a id="images"></a>Imagens

- **Use arquivos PNG.** Toda imagem é armazenada como uma parte PNG.
- **Dimensione imagens embutidas com `style:`**, por exemplo `style: { width: "120pt", height: "60pt" }`. É o único tamanho de imagem que funciona como esperado na 0.1.0.
- **Passe um caminho em vez de um `File` aberto se o documento puder ser gerado mais de uma vez.** Imagens IO são lidas durante `render`, e uma segunda geração recebe uma imagem vazia.
- **Imagens de cabeçalho, de células de tabela e posicionadas até a 0.2.0.** A largura e a altura delas são calculadas como `width * height * 14.92` EMU, então elas saem quadradas. Para obter uma imagem quadrada de `S` polegadas, passe `width` e `height` iguais a `Math.sqrt(S * 914400 / 14.92).round`, que dá cerca de 248 para uma polegada.

## <a id="security"></a>Segurança

- **Não passe strings não confiáveis para `add_media`.** Um caminho faz o docxi ler esse arquivo do seu servidor, e uma URL faz o seu servidor baixá-la. Passe um IO aberto por você ou verifique o valor antes. Veja [Segurança](Security).
- **O texto é escapado para você.** As strings passadas para `text` são gravadas como texto XML, então a entrada de usuários não consegue mudar a estrutura do documento.

## <a id="performance-and-reliability"></a>Desempenho e confiabilidade

- **Gere documentos grandes em um job em segundo plano.** O documento inteiro é montado na memória quando você chama `render`. Veja [Integrações](Integrations).
- **Gere uma vez e reaproveite os bytes.** Guarde `word.render.string` se precisar enviar ou armazenar o documento mais de uma vez.
- **Fixe uma versão compatível** com `gem "docxi", "~> 0.1"`.
- **Teste o que você gera.** Um `.docx` é um zip de arquivos XML, então os testes podem descompactá-lo e verificar o XML. Os próprios testes do docxi funcionam assim. Veja [Arquitetura](Architecture).

## <a id="related-pages"></a>Páginas relacionadas

- [Recursos](Features)
- [Configuração](Configuration)
- [Solução de problemas](Troubleshooting)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
