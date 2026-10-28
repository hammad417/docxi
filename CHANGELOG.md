# Changelog

## 0.1.0 (2026-10-28)

docxi installs and runs again in current apps, with or without Rails. No code changes are needed, and the generated documents are the same as 0.0.6's (the tests compare them against 0.0.6's output).

### Compatibility

- Requires Ruby 2.5 or newer. Apps on older Rubies stay on 0.0.6.
- Works with or without Rails, and is tested alongside Rails 5.2 through 8.1.
- nokogiri `>= 1.8, < 2` instead of `>= 1.6.3, < 1.7`, which conflicted with every Rails version since 6.0.
- rubyzip `>= 1.3, < 4` instead of `~> 0.9.9`.
- No longer calls ActiveSupport's `blank?` and `present?`, which only worked when Rails had loaded them.

### Fixes

- Documents written with rubyzip 3 no longer contain Zip64 records, which Google Docs and older LibreOffice versions reject.
- `add_media` with an image URL works on Ruby 3 again.
- `add_media` no longer runs a string that starts with `|` as a shell command.
- `tab` and `frame` no longer raise `NameError`.
- `document.frame("text")` no longer raises `TypeError`.
- `page_numbers` in a header no longer raises `NameError`.

### Gem

- New maintainer, homepage and changelog links.
- Releases require MFA.

## 0.0.6 (2016-10-18)

Last release before the 0.1.0 update.
