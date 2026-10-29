<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki do docxi: gere documentos do Word (.docx) com Ruby

O docxi é uma gem Ruby que gera arquivos `.docx` do Microsoft Word a partir de código. Ele oferece parágrafos, texto formatado, tabelas, imagens, cabeçalhos e rodapés, números de página, hiperlinks e um sumário simples. Não precisa do Microsoft Word, do LibreOffice nem de um modelo, e funciona tanto em Ruby puro quanto em aplicações Ruby on Rails.

O [README](https://github.com/hammad417/docxi#readme) é a introdução curta. Esta wiki cobre todas as opções, traz receitas prontas para copiar, explica como o docxi funciona por dentro e ajuda quando algo dá errado.

Versão atual: **0.1.1**, para Ruby 2.5 ou mais recente. Consulte o [registro de alterações](Changelog).

## <a id="start-here"></a>Comece por aqui

**Primeiro contato com o docxi**

- [Primeiros passos](Getting-Started): instale a gem e crie seu primeiro documento
- [Instalação](Installation): requisitos, Bundler e atualização a partir da 0.0.x
- [Exemplos de uso](Usage-Examples): relatórios, faturas, cartas e tabelas geradas a partir de dados, prontos para copiar

**Criando documentos**

- [Recursos](Features): cada elemento e as opções que ele aceita
- [Configuração](Configuration): margens, configuração de página e unidades
- [Referência da API](API-Reference): classes e métodos em resumo
- [Boas práticas](Best-Practices): hábitos que evitam as surpresas mais comuns

**Quando algo dá errado**

- [Solução de problemas](Troubleshooting): mensagens de erro e documentos que não aparecem como deveriam
- [Perguntas frequentes](FAQ): respostas curtas para dúvidas comuns
- [Segurança](Security): como relatar problemas e lidar com entradas não confiáveis

**Para ir além**

- [Arquitetura](Architecture): como chamadas Ruby se tornam um pacote `.docx`
- [Integrações](Integrations): controllers do Rails, aplicações Rack e jobs em segundo plano
- [Implantação](Deployment): o docxi em produção e como as versões são publicadas
- [Como contribuir](Contributing) e [Ambiente de desenvolvimento](Development-Setup)
- [Roteiro](Roadmap), [Registro de alterações](Changelog) e [Créditos](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>O que o docxi faz e o que ele não faz

O docxi cria documentos novos. Ele não abre nem edita um `.docx` existente, não preenche um modelo do Word e não faz mala direta. Não há listas numeradas, estilos de título nem sumário automático, e as páginas são sempre US Letter em orientação retrato. Os problemas conhecidos previstos para a 0.2.0 estão no [roteiro](Roadmap).

## <a id="credits"></a>Créditos

A manutenção do docxi é feita por [Hammad Maqbool](https://github.com/hammad417), na [Phaedra Solutions](https://www.phaedrasolutions.com/). O projeto começou em 2016 a partir do [docxer](https://github.com/gtimti/docxer), de Timur Gabdrakipov. Veja os [Créditos](Credits) para mais detalhes.

## <a id="languages"></a>Idiomas

A versão em inglês é a versão de referência desta wiki. As traduções têm links na barra lateral e no topo de cada página.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Relatar um problema](https://github.com/hammad417/docxi/issues)
