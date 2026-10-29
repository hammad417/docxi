<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# 프로덕션에서 docxi 운영하기와 릴리스 게시

docxi는 서비스가 아니라 라이브러리이므로 그 자체로 배포할 것은 없습니다. 이 페이지에서는 docxi를 사용하는 앱을 출시할 때 중요한 점과, docxi의 새 버전을 릴리스하는 방법을 설명합니다.

## <a id="shipping-an-app-that-uses-docxi"></a>docxi를 사용하는 앱 출시하기

- **의존성.** `bundle install`이 nokogiri와 rubyzip을 설치합니다. nokogiri는 일반적인 Linux, macOS, Windows 플랫폼용으로 미리 컴파일된 gem을 제공합니다. 일부 최소 구성 컨테이너 이미지 같은 다른 플랫폼에서는 소스에서 컴파일되며 C 컴파일러가 필요합니다. [nokogiri 설치 가이드](https://nokogiri.org/tutorials/installing_nokogiri.html)를 참고하세요.
- **오피스 소프트웨어 불필요.** 서버와 컨테이너에 Microsoft Word나 LibreOffice가 필요하지 않습니다.
- **메모리.** `render`는 문서 전체를 메모리에서 만들고 `StringIO`로 반환합니다. 메모리 사용량은 페이지 수와 이미지 크기에 따라 늘어납니다.
- **네트워크와 파일.** docxi는 `add_media`에 전달한 이미지 경로를 읽을 때만 파일 시스템을 사용하고, 전달한 이미지 URL을 받을 때만 네트워크를 사용합니다. [보안](Security)을 참고하세요.
- **긴 문서.** 웹 요청 중이 아니라 백그라운드 작업에서 만드세요. [통합](Integrations)을 참고하세요.

## <a id="how-docxi-releases-are-published"></a>docxi 릴리스 게시 방법

다음 단계는 메인테이너를 위한 것입니다. 릴리스는 GitHub Actions에서 신뢰할 수 있는 게시(trusted publishing)로 RubyGems에 올라가므로, 저장소에는 API 키를 보관하지 않습니다.

1. `lib/docxi/version.rb`와 `CHANGELOG.md`를 업데이트하고 `master`에 병합합니다.
2. 커밋에 `v`를 붙인 버전 태그(예: `v0.1.1`)를 달고 태그를 푸시합니다.
3. `.github/workflows/release.yml`이 gem을 빌드해 RubyGems에 푸시합니다.

프리릴리스도 같은 단계를 따르며, `0.2.0.rc1` 같은 버전을 사용합니다. Bundler는 명시적으로 요청한 경우에만 프리릴리스를 설치합니다. CI(`.github/workflows/ci.yml`)는 `master`에 대한 모든 푸시와 모든 풀 리퀘스트에서 실행됩니다.

## <a id="related-pages"></a>관련 페이지

- [설치](Installation)
- [개발 환경 설정](Development-Setup)
- [변경 이력](Changelog)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
