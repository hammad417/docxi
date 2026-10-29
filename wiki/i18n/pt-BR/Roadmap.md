<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Roteiro e problemas conhecidos do docxi

O docxi não tem calendário de versões nem um roteiro com compromissos. O único trabalho planejado registrado é corrigir os problemas conhecidos abaixo. Eles estão previstos para a 0.2.0, sem data. Ficaram fora da 0.1.0 porque corrigi-los muda os documentos que os usuários atuais já geram.

## <a id="known-issues-planned-for-020"></a>Problemas conhecidos previstos para a 0.2.0

| Problema | Solução alternativa na 0.1.0 |
|---|---|
| Imagens posicionadas, de cabeçalho e de células de tabela calculam o tamanho como `width * height * 14.92` EMU nas duas dimensões | Use imagens embutidas dimensionadas com `style:` ou veja [Boas práticas](Best-Practices#images) |
| `borders:` no nível da tabela é aplicada a cada célula, e não ao contorno da tabela | Para ter só o contorno, defina `borders:` nos lados externos das células das extremidades |
| O código do campo de número de página é gravado sem a barra invertida de `\* MERGEFORMAT` | Nenhuma |
| Os números de página em células de tabela vêm precedidos de "GlobalOptions" | Coloque os números de página em um cabeçalho ou rodapé |
| As imagens são sempre armazenadas como partes PNG, e `[Content_Types].xml` só declara os dois primeiros cabeçalhos e rodapés | Use imagens PNG e no máximo dois cabeçalhos e dois rodapés |
| Passar só algumas margens grava as outras vazias | Passe `top`, `right`, `bottom` e `left` juntos |
| `docProps/core.xml` registra a hora local com um marcador UTC e nunca grava o autor | Nenhuma |
| `Docxi.test` lança `NoMethodError` | Não o chame. É um método auxiliar que sobrou de versões antigas |
| `document.br` e `document.tab` sem `page: true` produzem XML inválido | Use `p.br` e `p.tab` dentro de um parágrafo |

## <a id="not-planned"></a>Não planejado

Estes itens não estão na lista acima, e nada foi decidido sobre eles:

- editar documentos existentes ou preencher modelos
- listas numeradas, estilos de título ou sumário automático
- outros tamanhos de página ou páginas em paisagem

## <a id="related-pages"></a>Páginas relacionadas

- [Registro de alterações](Changelog)
- [Como contribuir](Contributing)
- [Perguntas frequentes](FAQ)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
