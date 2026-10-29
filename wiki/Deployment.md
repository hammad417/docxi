# Running docxi in production and publishing releases

docxi is a library, not a service, so there's nothing to deploy on its own. This page covers what matters when you ship an app that uses docxi, and how new docxi versions are released.

## Shipping an app that uses docxi

- **Dependencies.** `bundle install` installs nokogiri and rubyzip. nokogiri ships precompiled gems for common Linux, macOS and Windows platforms. On other platforms, for example some minimal container images, it compiles from source and needs a C compiler. See the [nokogiri installation guide](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **No office software.** Servers and containers don't need Microsoft Word or LibreOffice.
- **Memory.** `render` builds the whole document in memory and returns it as a `StringIO`. Memory use grows with the number of pages and the size of the images.
- **Network and files.** docxi reads the file system only for image paths you pass to `add_media`, and uses the network only for image URLs you pass. See [Security](Security).
- **Long documents.** Build them in a background job instead of during a web request. See [Integrations](Integrations).

## How docxi releases are published

These steps are for maintainers. Releases go to RubyGems through GitHub Actions with trusted publishing, so no API key is stored in the repository.

1. Update `lib/docxi/version.rb` and `CHANGELOG.md`, and merge to `master`.
2. Tag the commit with the version prefixed by `v`, for example `v0.1.1`, and push the tag.
3. `.github/workflows/release.yml` builds the gem and pushes it to RubyGems.

Pre-releases follow the same steps with a version such as `0.2.0.rc1`. Bundler only installs pre-releases when they're requested explicitly. CI (`.github/workflows/ci.yml`) runs on every push to `master` and on every pull request.

## Related pages

- [Installation](Installation)
- [Development Setup](Development-Setup)
- [Changelog](Changelog)

---

[← Wiki home](Home) · [README](https://github.com/hammad417/docxi#readme)
