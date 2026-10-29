# Contributing to docxi

Bug reports, fixes, documentation and translations are all welcome. This page explains how to report a problem and what a pull request needs. To set up a working copy, see [Development Setup](Development-Setup).

## Reporting a bug

Open an issue at [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) with:

- a short Ruby script that reproduces the problem
- the versions of Ruby, docxi, nokogiri and rubyzip (and Rails, if you use it)
- what you expected, and what you got. For broken documents, say which program you opened them in

Report security problems privately instead. See [Security](Security).

## Pull requests

1. Fork the repository and create a branch.
2. Make the change, and match the style of the surrounding code.
3. Add or update tests, and run `bundle exec rake test`.
4. If the generated XML changes on purpose, run `bundle exec rake golden` and explain the change in the pull request. The golden files guard the output that existing users rely on.
5. Update `CHANGELOG.md` and the documentation if behaviour changes.
6. Open the pull request. CI runs the tests on Ruby 2.5 to 4.0 and Rails 5.2 to 8.1, and checks that LibreOffice opens the sample documents.

Changes that alter existing documents, such as the fixes planned for 0.2.0, belong in their own pull request so the golden file diff is easy to review.

## Documentation and translations

- The English README is the source for 15 translations in `docs/i18n/`. Follow `docs/i18n/TRANSLATION_GUIDE.md` in the repository when you change it.
- This wiki is written in the `wiki/` folder of the repository and published from there. `wiki/README.md` explains how.

## Related pages

- [Development Setup](Development-Setup)
- [Architecture](Architecture)
- [Roadmap](Roadmap)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
