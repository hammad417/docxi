<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 기능: 단락, 텍스트, 표, 이미지, 머리글과 바닥글

이 페이지에서는 docxi가 Word 문서에 추가할 수 있는 모든 요소와 각 요소가 받는 옵션을 설명합니다. 예제는 모두 다음 설정을 바탕으로 합니다.

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

필요 없는 옵션은 생략할 수 있습니다. 색은 `#` 없는 16진수 값입니다(예: `"8B0000"`). 단위와 페이지 설정은 [구성](Configuration)을 참고하세요.

## <a id="paragraphs"></a>단락

`document.p(options) { |p| ... }`는 단락을 추가합니다.

| 옵션 | 값 | 효과 |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | 맞춤. `"both"`는 양쪽 맞춤입니다 |
| `fill` | 16진수 색 | 단락과 그 텍스트 뒤의 음영 |
| `left`, `right` | 1/20포인트 | 들여쓰기. 하나만 설정하면 무시되므로 둘 다 설정하세요 |
| `ul` | `true` | 단락을 글머리 기호 목록 항목으로 만듭니다 |
| `bottom` | `true` | 단락 아래에 테두리를 그립니다 |

`size:`, `bold:`, `color:` 같은 텍스트 옵션도 `p`에 전달할 수 있습니다. 이 값은 단락 안 모든 텍스트 런의 기본값이 됩니다.

블록 안에서는 `text`, `br`, `tab`, `image`, `hyperlink`, `frame`을 호출할 수 있습니다. `document.text(string, options)`는 런이 하나뿐인 단락을 만드는 단축 메서드입니다.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>텍스트

`p.text(string, options)`는 텍스트 런(run)을 하나 추가합니다.

| 옵션 | 값 | 효과 |
|---|---|---|
| `bold` | `true` | 굵게 |
| `italic` | `true` | 기울임꼴 |
| `underline` | `true` | 한 줄 밑줄 |
| `color` | 16진수 색 | 텍스트 색 |
| `size` | 포인트 | 글꼴 크기 |
| `font` | 글꼴 이름 | 글꼴, 예: `"Georgia"` |
| `fill` | 16진수 색 | 텍스트 뒤의 음영 |
| `space` | `true` | 텍스트 앞뒤의 공백을 유지합니다 |
| `br` | `true` | 텍스트 뒤에 줄 바꿈을 추가합니다 |

`nil` 문자열은 건너뜁니다.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>줄 바꿈, 페이지 나누기, 탭

- `p.br` 또는 `p.br(times: 3)`은 단락 안에 줄 바꿈을 추가합니다.
- `document.br(page: true)`는 새 페이지를 시작합니다.
- `p.tab` 또는 `p.tab(times: 2)`는 단락 안에 탭 문자를 추가합니다.
- `document.tab(page: true)`는 탭 하나만 들어 있는 단락을 추가합니다.

`document.br`나 `document.tab`을 `page: true` 없이 호출하지 마세요. 텍스트 런이 문서 본문에 바로 들어가는데, 이는 올바른 WordprocessingML이 아닙니다.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>표

`document.table(options)`는 블록을 받으며, 그 안에서 `table.tr(options)`가 행을, `tr.tc(options)`가 셀을 추가합니다.

**표 옵션**

| 옵션 | 값 | 효과 |
|---|---|---|
| `columns_width` | 숫자 배열 | 열 너비. 14.92를 곱해 1/20포인트로 변환합니다. 셀의 기본 너비이기도 합니다 |
| `height` | 1/20포인트 | 모든 행의 최소 높이 |
| `iwidth` | 1/20포인트 | 왼쪽 여백에서 표까지의 들여쓰기 |
| `style` | 스타일 ID | 표 스타일. docxi는 `TableNormal`과 `TableCustom`을 정의합니다 |
| `width` | 숫자 | 표의 기본 설정 너비로, 형식 `auto`와 함께 기록됩니다. `columns_width`가 없으면 모든 셀도 이 너비에 14.92를 곱한 값을 받습니다 |
| `align` | `"left"`, `"center"`, `"right"` | 표 맞춤. `tc.text`로 추가한 텍스트도 같은 방식으로 맞춰집니다 |
| `padding` | `true` | 셀 안에 고정된 안쪽 여백을 추가합니다 |
| `borders` | 해시 | 알려진 문제: 표의 바깥 테두리가 아니라 모든 셀에 적용됩니다 |

**행 옵션**: `height:`는 행의 최소 높이를 설정합니다. `tr`에 전달한 셀 옵션은 그 행의 모든 셀에 적용됩니다. 표 옵션도 같은 방식으로 모든 행과 셀에 전달됩니다.

**셀 옵션**

| 옵션 | 값 | 효과 |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | 1/8포인트 단위의 테두리 두께. `nil`은 그쪽에 테두리가 없다는 뜻입니다 |
| `fill` | 16진수 색 | 셀 음영. 채우기가 없는 셀은 흰색(`"FFFFFF"`)입니다 |
| `valign` | `"top"`, `"center"`, `"bottom"` | 세로 맞춤 |
| `colspan` | 정수 | 셀이 차지하는 열 수 |
| `rowspan` | `true` | 세로 병합을 시작합니다 |
| `merged` | `true` | 위쪽 셀의 세로 병합을 이어 갑니다. 이 셀의 내용은 무시됩니다 |
| `width` | 숫자 | 셀 너비. 14.92를 곱합니다 |

**셀 내용**

- `tc.text(string, options)`: `bold:`나 `align:` 같은 셀과 행의 옵션이 텍스트에도 적용됩니다.
- `tc.p(options) { |p| ... }`: 자기 옵션만 사용하는 단락.
- `tc.image(media, width:, height:)`: 이미지. 크기 문제는 [이미지](#images)를 참고하세요.
- `tc.page_numbers`: 페이지 번호 필드. 알려진 문제: 앞에 "GlobalOptions"라는 텍스트가 붙습니다.

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>이미지

`media = document.add_media(source)`는 이미지를 등록합니다. `source`로 쓸 수 있는 것은 다음과 같습니다.

- `File`이나 `StringIO` 같은 IO 객체
- 파일 경로
- `http`, `https`, `ftp` URL

이미지는 `render`를 호출할 때 읽히고 항상 PNG 파트로 저장되므로, PNG 파일을 쓰는 것이 안전합니다.

**인라인 이미지**: `p.image(media, options)`는 이미지를 텍스트 줄 안에 넣습니다.

| 옵션 | 값 | 효과 |
|---|---|---|
| `style` | CSS 속성 해시 | 크기, 예: `{ width: "120pt", height: "60pt" }` |
| `title` | 문자열 | 이미지 제목 |
| `fill_color` | `"#00FF00"` 같은 VML 색 | 이미지 뒤의 채우기 |
| `opacity` | `"50%"` 같은 문자열 | 채우기의 불투명도 |
| `wrap` | `true` | 이미지 주위의 텍스트 줄 바꿈을 끕니다 |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**위치를 지정한 이미지**: `p.image(media, pH:, pV:, wrap: true, width:, height:)`는 열(`pH`)과 단락(`pV`)으로부터의 오프셋을 EMU로 지정해 이미지를 고정합니다. `pH`, `pV`, `wrap`을 모두 지정해야 하며, 그렇지 않으면 이미지가 인라인으로 배치됩니다.

**머리글 이미지**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"`은 줄 바꿈을 끕니다. 다른 값을 주면 텍스트가 이미지 주위를 사각형으로 감쌉니다.

**0.1.0의 알려진 문제**: 위치를 지정한 이미지, 머리글 이미지, 표 셀 이미지는 너비와 높이를 모두 `width * height * 14.92` EMU로 계산합니다. 그래서 항상 정사각형이 되고 대개 아주 작게 나옵니다. `style:`로 크기를 지정한 인라인 이미지는 영향을 받지 않습니다. 해결 방법은 [모범 사례](Best-Practices#images)를 참고하세요.

## <a id="headers-and-footers"></a>머리글과 바닥글

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- 머리글에는 `text`, `br`, `table`, `image`, `page_numbers`를 넣을 수 있습니다. 바닥글에는 `text`, `br`, `table`, `page_numbers`를 넣을 수 있습니다.
- `Header.new`나 `Footer.new`에 전달한 옵션은 그 안의 텍스트에 적용됩니다.
- `add_header`와 `add_footer`는 2페이지부터 적용됩니다. 1페이지에는 항상 `add_first_header`와 `add_first_footer`가 사용됩니다. [구성](Configuration#first-page)을 참고하세요.

## <a id="page-numbers"></a>페이지 번호

`page_numbers(name: "Page", align: "center")`는 `name` 텍스트 뒤에 PAGE 필드를 씁니다. 번호는 Arial 10pt의 짙은 회색으로 표시됩니다. `align`의 기본값은 `"right"`입니다.

알려진 문제: 필드 코드에서 `\* MERGEFORMAT`의 백슬래시가 빠져 있습니다.

## <a id="table-of-contents"></a>목차

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- 페이지 번호를 지정하면 해당 항목에 `pos`(1/20포인트 단위, 9360은 6.5인치)에 오른쪽 맞춤 탭 정지가 생기고, `leader` 문자로 채워집니다. 채움선 값은 WordprocessingML의 값으로, `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"`, `"none"` 등이 있습니다.
- 페이지 번호가 없으면 해당 항목은 탭 뒤의 텍스트가 됩니다.
- 페이지 번호는 전달한 값이 그대로 쓰입니다. docxi는 Word의 TOC 필드를 만들지 않으며 페이지를 계산하지 않습니다.

## <a id="hyperlinks"></a>하이퍼링크

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)`는 관계 ID를 반환하고, `hyperlink(text, url, id)`가 링크를 넣습니다. 링크 대상은 관계에서 정해지므로, 원하는 URL에 대해 반환된 ID를 전달하세요. 외부 링크만 지원합니다.

## <a id="frames-text-boxes"></a>프레임(텍스트 상자)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- 프레임은 연한 회색의 반투명 채우기와 검은 윤곽선이 고정된 VML 텍스트 상자입니다. 텍스트는 프레임의 위아래로 배치됩니다.
- 프레임 안에서는 `text`, `br`, `tab`, `image`, `frame`을 호출할 수 있습니다.
- `spacing: true`는 텍스트 앞에 간격을 추가합니다.

## <a id="output"></a>출력

`word.render`는 완전한 `.docx` 패키지가 담긴 `StringIO`를 반환합니다. 아카이브는 Zip64 레코드 없이 기록되므로 Google Docs와 이전 버전의 LibreOffice에서도 열 수 있습니다. 패키지의 모든 파트는 [아키텍처](Architecture)에 정리되어 있습니다.

## <a id="related-pages"></a>관련 페이지

- [사용 예제](Usage-Examples)
- [API 참조](API-Reference)
- [로드맵](Roadmap): 0.2.0에서 고칠 예정인 알려진 문제

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
