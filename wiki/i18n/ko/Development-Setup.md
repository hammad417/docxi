<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# 개발 환경 설정: docxi 테스트와 도구 실행하기

이 페이지에서는 docxi 작업용 복사본을 실행하는 방법과, 테스트, 샘플 문서, CI 구성을 설명합니다.

## <a id="get-the-code"></a>코드 받기

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

개발에는 Ruby 2.5 이상이면 어떤 버전이든 쓸 수 있습니다. 개발 의존성은 minitest와 rake입니다.

## <a id="rake-tasks"></a>Rake 작업

| 명령 | 하는 일 |
|---|---|
| `bundle exec rake test` | 테스트 스위트를 실행합니다(기본 작업) |
| `bundle exec rake samples` | 샘플 문서를 `tmp/samples/*.docx`에 써서 Word, LibreOffice, Google Docs에서 열어 볼 수 있게 합니다 |
| `bundle exec rake golden` | 현재 코드로 `test/golden`을 다시 씁니다. 출력을 의도적으로 바꾼 뒤에만 사용하세요 |
| `bundle exec rake build` | gem을 `pkg/`에 빌드합니다 |

## <a id="running-the-tests-with-rails-loaded"></a>Rails를 로드한 상태로 테스트 실행하기

`gemfiles/` 폴더에는 Rails 버전마다 Gemfile이 하나씩 있고, 오래된 rubyzip과 nokogiri 버전용 Gemfile도 몇 개 있습니다.

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1`은 테스트 전에 프레임워크 전체를 로드합니다. 또한 docxi의 `blank?` 규칙을 ActiveSupport의 규칙과 비교합니다.

## <a id="what-the-tests-cover"></a>테스트가 다루는 내용

| 파일 | 확인하는 것 |
|---|---|
| `test/test_golden_output.rb` | 샘플 문서 두 개의 모든 파트가 0.0.6이 생성한 것과 같은지 |
| `test/test_package.rb` | XML이 올바른 형식인지, 관계가 존재하는 파트를 가리키는지, 모든 파트에 콘텐츠 형식이 있는지, Zip64 레코드가 없는지 |
| `test/test_media.rb` | 경로, IO 객체, URL에서 가져온 이미지, 그리고 `\|` 문자열이 절대 명령으로 실행되지 않는지 |
| `test/test_fixed_features.rb` | `tab`, `frame`, 머리글의 페이지 번호 |
| `test/test_util.rb` | `Docxi::Util.blank?`가 ActiveSupport의 규칙을 따르는지 |
| `test/test_rails_independence.rb` | docxi가 ActiveSupport 없이 실행되는지 |
| `test/test_warnings.rb` | 렌더링할 때 `lib/`에서 사용 중단 경고가 나오지 않는지 |

샘플 문서는 `test/support/sample_documents.rb`에 있으며, `test/support/package.rb`는 `.docx`를 zip 구조에서 바로 읽습니다.

## <a id="continuous-integration"></a>지속적 통합

`.github/workflows/ci.yml`은 `master`에 대한 모든 푸시와 모든 풀 리퀘스트에서 실행됩니다. 다루는 범위는 다음과 같습니다.

- Rails 없이 Ruby 2.5, 2.6, 4.0
- Rails 5.2(Ruby 2.5), 6.0(2.7), 6.1(3.0), 7.0(3.1), 7.1(3.2), 7.2(3.3), 8.0(3.4), 8.1(4.0)
- 실패해도 되는 Rails `main`
- rubyzip 1.3과 2.4, nokogiri 1.11
- LibreOffice로 샘플 문서를 PDF로 변환하는 작업

릴리스는 [배포](Deployment)에서 설명합니다.

## <a id="related-pages"></a>관련 페이지

- [기여하기](Contributing)
- [아키텍처](Architecture)
- [배포](Deployment)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
