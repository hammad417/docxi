<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 설치: 요구 사항, Bundler 설정, 업그레이드

docxi는 순수 Ruby로 작성된 gem입니다. 설치하면 nokogiri와 rubyzip, 두 가지 의존성만 함께 설치됩니다. 오피스 소프트웨어는 필요하지 않습니다.

## <a id="requirements"></a>요구 사항

| 요구 사항 | 버전 |
|---|---|
| Ruby | 2.5 이상 |
| nokogiri | `>= 1.8, < 2`, 자동으로 설치됨 |
| rubyzip | `>= 1.3, < 4`, 자동으로 설치됨 |
| Rails | 필요 없음 |

## <a id="install-with-bundler"></a>Bundler로 설치하기

Gemfile에 gem을 추가합니다.

```ruby
gem "docxi", "~> 0.1"
```

그런 다음 설치합니다.

```sh
bundle install
```

## <a id="install-without-bundler"></a>Bundler 없이 설치하기

```sh
gem install docxi
```

## <a id="check-the-installation"></a>설치 확인하기

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Bundler 프로젝트에서는 `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`으로 실행하세요. 설치된 버전(예: `0.1.0`)이 출력됩니다.

## <a id="upgrading-from-00x"></a>0.0.x에서 업그레이드하기

Gemfile이 docxi를 0.0.x로 고정하고 있다면(예: `"~> 0.0.6"`) 해당 줄을 `gem "docxi", "~> 0.1"`로 바꾸세요. 그런 다음 `bundle update docxi`를 실행하세요. 코드를 바꿀 필요가 없으며, 0.1.0은 0.0.6과 같은 문서를 생성합니다. 테스트 스위트가 출력을 0.0.6이 생성한 파일과 비교합니다.

내부적으로 바뀐 점:

- 이제 nokogiri와 rubyzip이 위 범위 안의 어떤 버전이어도 되므로, docxi를 최신 Rails 버전과 함께 설치할 수 있습니다.
- docxi는 더 이상 ActiveSupport가 로드되어 있어야 하지 않습니다.
- 아카이브를 Zip64 레코드 없이 씁니다.
- 이미지를 `Kernel#open` 없이 불러오므로 Ruby 3에서도 URL이 동작합니다.
- `tab`, `frame`, 머리글의 페이지 번호가 동작합니다.

Ruby 2.5보다 오래된 Ruby에서 실행되는 애플리케이션은 계속 0.0.6을 받습니다. 0.1.0이 최소 버전으로 Ruby 2.5를 선언하기 때문입니다. 모든 변경 사항은 [변경 이력](Changelog)에 있습니다.

## <a id="versions-tested-in-ci"></a>CI에서 테스트하는 버전

각 Rails 버전은 전체 프레임워크를 docxi와 함께 로드한 상태로 실행합니다.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI는 다음도 실행합니다.

- Ruby 2.5, 2.6, 4.0에서 Rails 없이 실행하는 테스트
- rubyzip 1.3(Ruby 2.5), rubyzip 2.4(Ruby 3.3), nokogiri 1.11(Ruby 2.5)을 사용하는 테스트
- LibreOffice가 샘플 문서를 변환할 수 있는지 확인하는 검사

## <a id="platform-notes"></a>플랫폼 참고 사항

- nokogiri는 일반적인 Linux, macOS, Windows 플랫폼용으로 미리 컴파일된 gem을 제공합니다. 맞는 것이 없으면 소스에서 컴파일되며 C 컴파일러가 필요합니다. [nokogiri 설치 가이드](https://nokogiri.org/tutorials/installing_nokogiri.html)를 참고하세요.
- CI는 Ubuntu 24.04에서 실행되며, 테스트는 macOS에서도 실행해 보았습니다. Windows는 테스트하지 않습니다.

## <a id="if-installation-fails"></a>설치가 실패할 때

nokogiri나 rubyzip에서 Bundler 충돌이 난다면 거의 항상 오래된 docxi 0.0.6이 선택되고 있는 것입니다. [문제 해결](Troubleshooting#installation)을 참고하세요.

## <a id="related-pages"></a>관련 페이지

- [시작하기](Getting-Started)
- [구성](Configuration)
- [배포](Deployment)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
