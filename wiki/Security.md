# docxi security: reporting issues and handling untrusted input

This page explains how to report a security problem, and the security points to know when your app passes user input to docxi.

## Reporting a vulnerability

Please don't report security problems in public issues. Email the maintainer instead. The addresses are in [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Include the docxi version and a short way to reproduce the problem.

## Upgrade from 0.0.x

docxi 0.0.x passes image path strings to `Kernel#open`. `Kernel#open` runs a string that starts with `|` as a shell command, so an attacker who controls an image path could run commands on your server. 0.1.0 opens files with `File.open` and downloads URLs with `URI#open`. Neither runs commands.

## Untrusted input

**Text is safe to pass.** Text, table contents and other strings are written as XML text or attribute values, which nokogiri escapes. User input can't add elements to the document.

**Image sources need care.** `add_media` accepts:

| Source | What docxi does |
|---|---|
| An IO you opened | Reads it |
| A file path | Reads that file from your server's disk |
| An `http`, `https` or `ftp` URL | Downloads it from your server |

If a user can choose the path, they could get any readable file on your server, such as configuration or keys, embedded in the document. If a user can choose the URL, your server will request it, which can reach internal addresses (server-side request forgery). Pass an IO you opened yourself, or allow only known paths and hosts.

**Hyperlinks** point wherever the URL says. Check user-provided URLs before passing them to `add_hyperlink`.

## How releases are protected

- Releases are published from GitHub Actions with RubyGems trusted publishing, so no API key is stored in the repository.
- The gem metadata sets `rubygems_mfa_required`, so pushes made with a personal API key need multi-factor authentication.

## Related pages

- [Best Practices](Best-Practices)
- [Deployment](Deployment)
- [Contributing](Contributing)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
