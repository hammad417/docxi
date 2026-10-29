<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Segurança no docxi: como relatar problemas e lidar com entradas não confiáveis

Esta página explica como relatar um problema de segurança e o que é importante saber quando a sua aplicação repassa entradas de usuários para o docxi.

## <a id="reporting-a-vulnerability"></a>Relatar uma vulnerabilidade

Não relate problemas de segurança em issues públicas. Envie um e-mail para os endereços de manutenção que estão em [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Informe a versão do docxi e uma forma curta de reproduzir o problema.

## <a id="upgrade-from-00x"></a>Atualize a partir da 0.0.x

O docxi 0.0.x passa as strings de caminho de imagem para `Kernel#open`. `Kernel#open` executa como comando de shell uma string que começa com `|`, então quem controla um caminho de imagem poderia executar comandos no seu servidor. A 0.1.0 abre arquivos com `File.open` e baixa URLs com `URI#open`. Nenhum dos dois executa comandos.

## <a id="untrusted-input"></a>Entradas não confiáveis

**Passar texto é seguro.** Texto, conteúdo de tabelas e outras strings são gravados como texto XML ou valores de atributo, que o nokogiri escapa. A entrada de usuários não consegue adicionar elementos ao documento.

**Fontes de imagem exigem cuidado.** `add_media` aceita:

| Fonte | O que o docxi faz |
|---|---|
| Um IO aberto por você | Lê o IO |
| Um caminho de arquivo | Lê esse arquivo do disco do seu servidor |
| Uma URL `http`, `https` ou `ftp` | Baixa a URL a partir do seu servidor |

Se um usuário puder escolher o caminho, poderá fazer com que qualquer arquivo legível do seu servidor, como configurações ou chaves, seja embutido no documento. Se puder escolher a URL, o seu servidor vai requisitá-la, o que pode alcançar endereços internos (falsificação de requisição do lado do servidor, SSRF). Passe um IO aberto por você ou permita apenas caminhos e hosts conhecidos.

**Hiperlinks** apontam para onde a URL indicar. Verifique as URLs fornecidas por usuários antes de passá-las para `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Como as versões são protegidas

- As versões são publicadas pelo GitHub Actions com a publicação confiável (trusted publishing) do RubyGems, então nenhuma chave de API fica guardada no repositório.
- Os metadados da gem definem `rubygems_mfa_required`, então publicações feitas com uma chave de API pessoal exigem autenticação multifator.

## <a id="related-pages"></a>Páginas relacionadas

- [Boas práticas](Best-Practices)
- [Implantação](Deployment)
- [Como contribuir](Contributing)

---

[← Início da wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
