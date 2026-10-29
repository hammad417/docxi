# Package listing

docxi is a Ruby library with no executables, so RubyGems.org is the only registry it's published to. These notes cover the gem's listing and the Ruby sites that read from it.

## RubyGems.org

From 0.1.1 the gemspec has:

- **Summary:** "Generate Microsoft Word (.docx) documents with Ruby". RubyGems shows the summary in `gem list -d`. It uses the same wording as the README title and the GitHub description.
- **Description:** what docxi builds, and that it needs no Word, LibreOffice or template file and works with or without Rails. RubyGems asks for a description that is "more detailed than the summary but not excessively long", with no examples or formatting.
- **Metadata links:** source code, changelog, bug tracker and documentation (the wiki), plus `rubygems_mfa_required`. RubyGems allows keys up to 128 bytes and values up to 1024 bytes.

Before 0.1.1 the summary was "Docx Template Generator" and the description "Create docx template using Ruby". A published version's gemspec can't be edited, so 0.1.0 and older keep those.

There's no keywords field on RubyGems: the name, summary and description are all the text a gem has. `funding_uri` and `mailing_list_uri` are left out because there's no funding page or mailing list. `homepage` stays on the GitHub repository rather than the company site.

## Ruby Toolbox

Ruby Toolbox builds its project pages from RubyGems and GitHub, so the docxi page updates on its own after a release. Its categories come from the [rubytoolbox/catalog](https://github.com/rubytoolbox/catalog) repository. docxi isn't in a category yet, and neither are caracal, sablon or docx; a pull request to the catalog could add it, at the catalog maintainers' discretion.

## RubyDoc.info

RubyDoc.info builds API docs from published gems, but its docxi page wasn't responding in October 2026, and docxi has no YARD comments, so it would only list method names. The wiki's API Reference page is the API documentation.

## Release notes

`CHANGELOG.md` is the source for every release. Write the entry before tagging. The same text works for a GitHub Release if you want one.
