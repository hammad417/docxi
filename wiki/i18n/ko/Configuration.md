<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 구성: 페이지 여백, 페이지 설정, 단위

docxi에는 설정 파일, 환경 변수, 전역 설정이 없습니다. 모든 것은 문서별 또는 요소별로 지정합니다. 이 페이지에서는 문서 수준의 설정과 옵션에서 쓰는 단위를 설명합니다. 각 요소의 옵션은 [기능](Features)에 있습니다.

## <a id="page-margins"></a>페이지 여백

여백은 문서를 만들 때 전달합니다.

```ruby
require "docxi"

# 모든 값의 단위는 1/20포인트입니다(1440 = 1인치).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| 옵션 | 기본값 | 의미 |
|---|---|---|
| `top` | `720`(0.5인치) | 위쪽 여백 |
| `right` | `1440`(1인치) | 오른쪽 여백 |
| `bottom` | `720`(0.5인치) | 아래쪽 여백 |
| `left` | `1440`(1인치) | 왼쪽 여백 |
| `header` | `190` | 페이지 위쪽 가장자리에서 머리글까지의 거리 |
| `footer` | `200` | 페이지 아래쪽 가장자리에서 바닥글까지의 거리 |

전달한 해시가 기본값을 통째로 대체하므로 `top`, `right`, `bottom`, `left`는 함께 설정하세요. 이 네 가지 중 빠진 값은 빈 값으로 기록됩니다. `header`와 `footer`는 없거나 비어 있으면 190과 200이 사용됩니다.

## <a id="page-size-and-orientation"></a>페이지 크기와 방향

페이지는 세로 방향의 US Letter이며, 1/20포인트 단위로 12240 × 15840, 즉 8.5 × 11인치입니다. 0.1.0에서는 크기와 방향을 바꿀 수 없습니다.

## <a id="first-page"></a>첫 페이지

docxi는 Word의 "첫 페이지를 다르게 지정" 설정을 항상 켭니다. `add_header`와 `add_footer`로 추가한 머리글과 바닥글은 2페이지부터 사용됩니다. 1페이지에는 `add_first_header`와 `add_first_footer`가 사용되므로, 첫 페이지에 머리글이나 바닥글이 필요하면 이것도 설정하세요. [기능](Features#headers-and-footers)을 참고하세요.

## <a id="what-every-document-contains"></a>모든 문서에 들어 있는 것

- 기본 글꼴은 Arial 12pt입니다.
- 정의된 스타일: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom`, `NoList`.
- 글머리 기호 항목(`ul: true`)은 대시를 글머리 기호로 사용합니다.
- 문서 속성에는 애플리케이션 이름이 "Docxi Ruby Gem"으로 기록되고, 문서를 렌더링한 시각도 기록됩니다.

## <a id="units"></a>단위

| 옵션 | 단위 |
|---|---|
| 텍스트의 `size:` | 포인트(docxi가 2배로 바꿔 Word의 반 포인트 단위로 저장) |
| 여백, 단락의 `left:`와 `right:`, 표의 `iwidth:`, 행의 `height:` | 1/20포인트(1인치당 1440) |
| 셀 `borders:`의 두께 | 1/8포인트 |
| `columns_width:`와 셀의 `width:` | 14.92를 곱해 1/20포인트로 변환하며, 96dpi의 픽셀과 비슷합니다 |
| 인라인 이미지의 `style:` | `"120pt"` 같은 CSS 길이 |
| 위치를 지정한 이미지의 `pH:`와 `pV:` | EMU(1인치당 914400) |
| 텍스트와 채우기 색 | `#` 없는 16진수, 예: `"8B0000"` |
| 이미지의 `fill_color:` | `#`이 붙은 VML 색, 예: `"#00FF00"` |

## <a id="related-pages"></a>관련 페이지

- [기능](Features)
- [API 참조](API-Reference)
- [모범 사례](Best-Practices)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
