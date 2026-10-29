# Development setup: running docxi's tests and tools

This page gets a working copy of docxi running, and explains the tests, the sample documents and the CI setup.

## Get the code

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Any Ruby from 2.5 up works for development. The development dependencies are minitest and rake.

## Rake tasks

| Command | What it does |
|---|---|
| `bundle exec rake test` | Runs the test suite (the default task) |
| `bundle exec rake samples` | Writes the sample documents to `tmp/samples/*.docx`, for opening in Word, LibreOffice or Google Docs |
| `bundle exec rake golden` | Rewrites `test/golden` from the current code. Only use it after an intentional output change |
| `bundle exec rake build` | Builds the gem into `pkg/` |

## Running the tests with Rails loaded

The `gemfiles/` folder has one Gemfile per Rails version, plus some for older rubyzip and nokogiri versions:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` loads the whole framework before the tests. It also compares docxi's `blank?` rules with ActiveSupport's.

## What the tests cover

| File | Checks |
|---|---|
| `test/test_golden_output.rb` | Every part of two sample documents matches what 0.0.6 generated |
| `test/test_package.rb` | XML is well formed, relationships point to existing parts, every part has a content type, and there are no Zip64 records |
| `test/test_media.rb` | Images from paths, IO objects and URLs, and that a `\|` string is never run as a command |
| `test/test_fixed_features.rb` | `tab`, `frame` and header page numbers |
| `test/test_util.rb` | `Docxi::Util.blank?` follows ActiveSupport's rules |
| `test/test_rails_independence.rb` | docxi runs without ActiveSupport |
| `test/test_warnings.rb` | Rendering emits no deprecation warnings from `lib/` |

The sample documents live in `test/support/sample_documents.rb`, and `test/support/package.rb` reads a `.docx` straight from its zip structures.

## Continuous integration

`.github/workflows/ci.yml` runs on every push to `master` and on every pull request. It covers:

- Ruby 2.5, 2.6 and 4.0 without Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) and 8.1 (4.0)
- Rails `main`, which is allowed to fail
- rubyzip 1.3 and 2.4, and nokogiri 1.11
- a job that converts the sample documents to PDF with LibreOffice

Releases are covered in [Deployment](Deployment).

## Related pages

- [Contributing](Contributing)
- [Architecture](Architecture)
- [Deployment](Deployment)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
