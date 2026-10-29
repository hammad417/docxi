# docxi changelog and release history

This page summarizes each release. The full notes are in [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) in the repository, which is the reference if the two ever differ.

## Releases

| Version | Date | Notes |
|---|---|---|
| 0.1.1 | 2026-10-29 | Accurate gem summary and description, and a documentation link. No code changes |
| 0.1.0 | 2026-10-28 | Works with current Ruby, Rails and rubyzip. Generates the same documents as 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Pre-release of 0.1.0 |
| 0.0.6 | 2016-10-18 | Last release before the 0.1.0 update |
| 0.0.2 | 2016-10-14 | No release notes were recorded |
| 0.0.1 | 2016-10-14 | First release. No release notes were recorded |

Versions 0.0.3 to 0.0.5 were never published to RubyGems.

## 0.1.1

No code changes, and the generated documents are the same as 0.1.0's.

- The gem's summary and description now say what docxi does. They used to call it a template generator.
- RubyGems links to this wiki as the gem's documentation.

## 0.1.0

**Compatibility**

- Requires Ruby 2.5 or newer. Apps on older Rubies stay on 0.0.6.
- Works with or without Rails, and is tested alongside Rails 5.2 through 8.1.
- nokogiri `>= 1.8, < 2` instead of `>= 1.6.3, < 1.7`, which conflicted with every Rails version since 6.0.
- rubyzip `>= 1.3, < 4` instead of `~> 0.9.9`.
- No longer calls ActiveSupport's `blank?` and `present?`, which only worked when Rails had loaded them.

**Fixes**

- Documents written with rubyzip 3 no longer contain Zip64 records, which Google Docs and older LibreOffice versions reject.
- `add_media` with an image URL works on Ruby 3 again.
- `add_media` no longer runs a string that starts with `|` as a shell command.
- `tab` and `frame` no longer raise `NameError`.
- `document.frame("text")` builds a text box containing the text.
- `page_numbers` in a header no longer raises `NameError`.

**Gem**

- New maintainer, homepage and changelog links.
- Releases require MFA.

## Related pages

- [Roadmap](Roadmap)
- [Installation](Installation#upgrading-from-00x)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
