<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Solução de problemas do docxi: erros de instalação, erros em tempo de execução e documentos com defeito

Encontre o seu sintoma abaixo. A maioria dos problemas em aplicações antigas vem do docxi 0.0.x, e atualizar para a 0.1.0 resolve sem mudanças no código.

## <a id="installation"></a>Instalação

**O Bundler não encontra versões compatíveis de `nokogiri` em uma aplicação Rails 6+.**
O docxi 0.0.6 fixa `nokogiri < 1.7`, e todas as versões do Rails desde a 6.0 exigem nokogiri 1.8.5 ou mais recente. Use `gem "docxi", "~> 0.1"`.

**O Bundler não encontra versões compatíveis de `rubyzip`.**
O docxi 0.0.6 fixa `rubyzip ~> 0.9.9`, o que entra em conflito com qualquer gem que precise de um rubyzip mais novo. O docxi 0.1.0 aceita rubyzip 1.3 a 3.x.

**O Bundler continua instalando a 0.0.6.**
Ou o Gemfile fixa o docxi em uma 0.0.x, por exemplo com `"~> 0.0.6"`, ou o Ruby é anterior à versão 2.5, que a 0.1.0 exige. Troque a linha do Gemfile por `gem "docxi", "~> 0.1"` e confira `ruby -v`.

**O nokogiri não compila.**
Nenhuma gem pré-compilada do nokogiri corresponde à sua plataforma, então ele está sendo compilado a partir do código-fonte. Instale um compilador C e siga o [guia de instalação do nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Erros ao montar um documento

**`NoMethodError: undefined method 'present?'` fora do Rails.**
O docxi 0.0.6 presumia que o ActiveSupport estava carregado. A 0.1.0 não precisa dele.

**`NameError` ao chamar `tab`, `frame` ou o `page_numbers` de um cabeçalho.**
Esses métodos não funcionavam na 0.0.6 e funcionam na 0.1.0, inclusive `document.frame("text")`.

**`Errno::ENOENT` ao adicionar uma imagem.**
O caminho não existe em relação ao diretório em que o seu programa roda. No Ruby 3 com docxi 0.0.6, URLs de imagem também falham dessa forma, porque a 0.0.6 usava `Kernel#open`. A 0.1.0 baixa URLs com `URI#open`.

## <a id="documents-that-dont-look-right"></a>Documentos que não aparecem como deveriam

**O Google Docs ou o LibreOffice não abre o arquivo.**
O docxi 0.0.6 com rubyzip 3 grava registros Zip64 em cada entrada, e esses programas os rejeitam. A 0.1.0 grava os arquivos compactados sem eles.

**A primeira página não tem cabeçalho nem rodapé.**
O docxi dá à primeira página um cabeçalho e um rodapé próprios. Defina-os com `add_first_header` e `add_first_footer`. Veja [Configuração](Configuration#first-page).

**Falta uma margem ou o layout da página parece errado.**
Ao definir as margens, passe `top`, `right`, `bottom` e `left` juntos. As que você omitir são gravadas vazias.

**Os espaços entre trechos de texto desaparecem.**
O Word descarta espaços no início ou no fim de um trecho. Adicione `space: true` a essa chamada de `text`.

**Imagens em cabeçalhos, em células de tabela ou com `pH`/`pV` ficam minúsculas ou quadradas.**
Problema conhecido na 0.1.0: o tamanho delas é calculado como `width * height * 14.92` EMU nas duas dimensões. Use imagens embutidas dimensionadas com `style:` ou veja a solução alternativa em [Boas práticas](Best-Practices#images).

**Uma imagem some na segunda vez que o mesmo documento é gerado.**
Imagens adicionadas a partir de um `File` aberto ou de outro IO são lidas durante `render`, e uma segunda geração lê um fluxo vazio. Passe um caminho em vez disso, ou gere o documento uma única vez.

**As bordas definidas em `table` aparecem em volta de cada célula.**
Problema conhecido: `borders:` no nível da tabela é aplicada a cada célula, e não ao contorno da tabela. Para desenhar só o contorno, defina `borders:` nas células das extremidades, apenas para os lados externos.

**Os números de página em células de tabela começam com "GlobalOptions".**
Problema conhecido em `TableCell#page_numbers`. Coloque os números de página em um rodapé ou cabeçalho.

**Uma imagem JPEG ou GIF aparece errada ou não aparece.**
Toda imagem é armazenada como uma parte PNG com tipo de conteúdo PNG, seja qual for o formato real, então alguns programas podem não exibir outros formatos. Converta as imagens para PNG antes de adicioná-las.

## <a id="inspecting-a-generated-file"></a>Inspecionar um arquivo gerado

Um `.docx` é um arquivo zip, então você pode ver o que tem dentro:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Se o problema continuar, [abra uma issue](https://github.com/hammad417/docxi/issues) com um script curto que o reproduza e as versões do Ruby, do docxi, do nokogiri e do rubyzip.

## <a id="related-pages"></a>Páginas relacionadas

- [Perguntas frequentes](FAQ)
- [Boas práticas](Best-Practices)
- [Instalação](Installation)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
