<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 모범 사례: 흔한 실수 피하기

이 습관들은 docxi가 동작하는 방식에서 바로 나온 것이며, 대부분 [문제 해결](Troubleshooting)에 나오는 문제를 예방합니다. 주제별로 묶었습니다. 페이지 설정에 관한 조언은 모두에게 해당하고, 보안과 성능에 관한 조언은 특히 프로덕션 앱에서 중요합니다.

## <a id="page-setup"></a>페이지 설정

- **네 여백을 함께 전달하세요.** `Docxi::Document.new`에 전달한 해시가 기본값을 대체하므로, `top`, `right`, `bottom`, `left` 중 빠진 값은 빈 값으로 기록됩니다.
- **첫 페이지의 머리글과 바닥글을 설정하세요.** docxi는 항상 첫 페이지를 따로 둡니다. 1페이지에 머리글이나 바닥글이 필요하면 `add_first_header`와 `add_first_footer`도 호출하세요.

## <a id="text"></a>텍스트

- **앞뒤 공백에는 `space: true`를 사용하세요.** 그렇지 않으면 Word가 공백을 지워 두 런의 단어가 붙어 버립니다.
- **공통 서식은 단락에 두세요.** `size:`나 `font:`처럼 `p`에 전달한 옵션은 그 안의 모든 런에 적용되므로 반복할 것이 줄어듭니다.
- **줄 바꿈에는 `p.br`, 페이지 나누기에는 `document.br(page: true)`를 사용하세요.** `page: true` 없는 `document.br`은 올바르지 않은 XML을 만듭니다.

## <a id="tables"></a>표

- **테두리는 표가 아니라 행이나 셀에 설정하세요.** 0.1.0에서는 표 수준의 `borders:`가 표의 바깥 테두리가 아니라 모든 셀에 적용됩니다.
- **모든 열에 `columns_width`를 지정하세요.** 셀은 여기서 너비를 가져오므로 표가 기대한 레이아웃을 유지합니다.

## <a id="images"></a>이미지

- **PNG 파일을 사용하세요.** 모든 이미지는 PNG 파트로 저장됩니다.
- **인라인 이미지 크기는 스타일 옵션으로 지정하세요**. 예: `style:`에 `style: { width: "120pt", height: "60pt" }`처럼 지정합니다. 0.1.0에서 기대한 대로 동작하는 이미지 크기 지정 방법은 이것뿐입니다.
- **문서를 여러 번 렌더링할 수 있다면 열린 파일 객체 대신 경로를 전달하세요.** 열린 `File` 같은 IO 이미지는 `render` 중에 읽히므로, 두 번째 렌더링에서는 빈 이미지가 됩니다.
- **0.2.0 전까지의 머리글, 표 셀, 위치를 지정한 이미지.** 너비와 높이가 모두 `width * height * 14.92` EMU로 계산되어 정사각형이 됩니다. 한 변이 `S`인치인 정사각형 이미지를 만들려면 `width`와 `height`를 모두 `Math.sqrt(S * 914400 / 14.92).round`로 지정하세요. 1인치라면 약 248입니다.

## <a id="security"></a>보안

- **신뢰할 수 없는 문자열을 이미지 추가 메서드에 전달하지 마세요.** `add_media`에 경로를 주면 docxi가 서버의 그 파일을 읽고, URL을 주면 서버가 그 URL을 내려받습니다. 직접 연 IO를 전달하거나 값을 먼저 확인하세요. [보안](Security)을 참고하세요.
- **텍스트는 자동으로 이스케이프됩니다.** `text`에 전달한 문자열은 XML 텍스트로 기록되므로, 사용자 입력이 문서 구조를 바꿀 수 없습니다.

## <a id="performance-and-reliability"></a>성능과 안정성

- **큰 문서는 백그라운드 작업에서 만드세요.** `render`를 호출하면 문서 전체가 메모리에서 만들어집니다. [통합](Integrations)을 참고하세요.
- **한 번 렌더링하고 바이트를 재사용하세요.** 문서를 여러 번 보내거나 저장해야 한다면 `word.render.string`을 보관해 두세요.
- **호환되는 버전으로 고정하세요**: `gem "docxi", "~> 0.1"`.
- **생성한 결과를 테스트하세요.** `.docx`는 XML 파일을 묶은 zip이므로 테스트에서 압축을 풀어 XML을 확인할 수 있습니다. docxi 자체의 테스트도 이렇게 동작합니다. [아키텍처](Architecture)를 참고하세요.

## <a id="related-pages"></a>관련 페이지

- [기능](Features)
- [구성](Configuration)
- [문제 해결](Troubleshooting)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
