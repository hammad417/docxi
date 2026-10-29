<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# docxi에 기여하기

버그 신고, 수정, 문서, 번역 모두 환영합니다. 이 페이지에서는 문제를 신고하는 방법과 풀 리퀘스트에 필요한 것을 설명합니다. 작업용 복사본을 준비하려면 [개발 환경 설정](Development-Setup)을 참고하세요.

## <a id="reporting-a-bug"></a>버그 신고

[github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues)에서 이슈를 열고 다음을 적어 주세요.

- 문제를 재현하는 짧은 Ruby 스크립트
- Ruby, docxi, nokogiri, rubyzip 버전(Rails를 쓴다면 Rails 버전도)
- 기대한 결과와 실제 결과. 문서가 깨졌다면 어떤 프로그램으로 열었는지도

보안 문제는 비공개로 신고해 주세요. [보안](Security)을 참고하세요.

## <a id="pull-requests"></a>풀 리퀘스트

1. 저장소를 포크하고 브랜치를 만듭니다.
2. 변경을 하되, 주변 코드의 스타일에 맞춥니다.
3. 테스트를 추가하거나 업데이트하고 `bundle exec rake test`를 실행합니다.
4. 생성되는 XML이 의도적으로 바뀐다면 `bundle exec rake golden`을 실행하고 풀 리퀘스트에서 변경 내용을 설명하세요. 골든 파일은 기존 사용자가 의존하는 출력을 지켜 줍니다.
5. 동작이 바뀐다면 `CHANGELOG.md`와 문서를 업데이트합니다.
6. 풀 리퀘스트를 엽니다. CI가 Ruby 2.5부터 4.0, Rails 5.2부터 8.1에서 테스트를 실행하고, LibreOffice가 샘플 문서를 여는지 확인합니다.

0.2.0에 예정된 수정처럼 기존 문서를 바꾸는 변경은 골든 파일의 차이를 검토하기 쉽도록 별도의 풀 리퀘스트로 보내 주세요.

## <a id="documentation-and-translations"></a>문서와 번역

- 영어 README는 `docs/i18n/`에 있는 번역 15개의 원본입니다. README를 수정할 때는 저장소의 `docs/i18n/TRANSLATION_GUIDE.md`를 따르세요.
- 이 위키는 저장소의 `wiki/` 폴더에서 작성되고 그곳에서 게시됩니다. 방법은 `wiki/README.md`에 있습니다.

## <a id="related-pages"></a>관련 페이지

- [개발 환경 설정](Development-Setup)
- [아키텍처](Architecture)
- [로드맵](Roadmap)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
