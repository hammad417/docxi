<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 변경 이력과 릴리스 기록

이 페이지는 각 릴리스를 요약합니다. 전체 릴리스 노트는 저장소의 [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md)에 있으며, 두 내용이 다를 때는 그쪽이 기준입니다.

## <a id="releases"></a>릴리스

| 버전 | 날짜 | 참고 |
|---|---|---|
| 0.1.1 | 2026-10-29 | gem 요약과 설명을 정확하게 고치고 문서 링크를 추가. 코드 변경 없음 |
| 0.1.0 | 2026-10-28 | 최신 Ruby, Rails, rubyzip에서 동작합니다. 0.0.6과 같은 문서를 생성합니다 |
| 0.1.0.rc1 | 2026-10-28 | 0.1.0의 프리릴리스 |
| 0.0.6 | 2016-10-18 | 0.1.0 업데이트 전의 마지막 릴리스 |
| 0.0.2 | 2016-10-14 | 기록된 릴리스 노트가 없습니다 |
| 0.0.1 | 2016-10-14 | 첫 릴리스. 기록된 릴리스 노트가 없습니다 |

0.0.3부터 0.0.5까지의 버전은 RubyGems에 게시된 적이 없습니다.

## <a id="011"></a>0.1.1

코드 변경은 없으며, 생성되는 문서는 0.1.0과 같습니다.

- gem 요약과 설명이 이제 docxi가 하는 일을 설명합니다. 이전에는 템플릿 생성기로 소개했습니다.
- RubyGems에서 gem 문서로 이 위키를 링크합니다.

## <a id="010"></a>0.1.0

**호환성**

- Ruby 2.5 이상이 필요합니다. 더 오래된 Ruby를 쓰는 앱은 0.0.6에 머뭅니다.
- Rails가 있든 없든 동작하며, Rails 5.2부터 8.1까지와 함께 테스트합니다.
- nokogiri 요구 사항을 `>= 1.6.3, < 1.7`에서 `>= 1.8, < 2`로 바꿨습니다. 이전 범위는 Rails 6.0 이후의 모든 버전과 충돌했습니다.
- rubyzip 요구 사항을 `~> 0.9.9`에서 `>= 1.3, < 4`로 바꿨습니다.
- 더 이상 ActiveSupport의 `blank?`와 `present?`를 호출하지 않습니다. 이 메서드들은 Rails가 로드한 경우에만 동작했습니다.

**수정 사항**

- rubyzip 3으로 쓴 문서에 더 이상 Zip64 레코드가 들어가지 않습니다. Google Docs와 이전 버전의 LibreOffice는 이 레코드를 거부합니다.
- 이미지 URL을 전달한 `add_media`가 Ruby 3에서 다시 동작합니다.
- `add_media`는 더 이상 `|`로 시작하는 문자열을 셸 명령으로 실행하지 않습니다.
- `tab`과 `frame`이 더 이상 `NameError`를 일으키지 않습니다.
- `document.frame("text")`는 그 텍스트가 든 텍스트 상자를 만듭니다.
- 머리글의 `page_numbers`가 더 이상 `NameError`를 일으키지 않습니다.

**gem**

- 메인테이너, 홈페이지, 변경 이력 링크가 새로 바뀌었습니다.
- 릴리스에는 MFA가 필요합니다.

## <a id="related-pages"></a>관련 페이지

- [로드맵](Roadmap)
- [설치](Installation#upgrading-from-00x)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
