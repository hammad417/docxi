<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# docxi API 참조: 클래스와 메서드

docxi는 Ruby 라이브러리이므로 이 페이지에서는 Ruby API를 정리합니다. HTTP API는 없습니다. 각 메서드의 옵션은 [기능](Features)에서, 클래스는 [아키텍처](Architecture)에서 설명합니다.

## <a id="docxidocument"></a>Docxi::Document

| 메서드 | 반환값 | 참고 |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options`는 페이지 여백입니다. [구성](Configuration)을 참고하세요 |
| `#document` | `Docxi::Word::Document` | 내용을 추가할 본문 |
| `#render` | `StringIO` | 완전한 `.docx` 패키지. `.string`으로 바이트를 얻습니다 |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document(`word.document`)

| 메서드 | 반환값 | 추가하는 것 |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | 단락 |
| `text(string, options = {})` | Paragraph | 텍스트 런이 하나인 단락 |
| `br(page: true)` | Break | 페이지 나누기 |
| `tab(page: true)` | Tab | 탭 하나가 들어 있는 단락 |
| `table(options = {}) { \|table\| }` | Table | 표 |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | 목차 항목 |
| `frame(text, options = {})` | Paragraph | 텍스트 상자가 들어 있는 단락 |
| `add_media(io_path_or_url, options = {})` | Media | `image`로 배치할 이미지 |
| `add_hyperlink(url, options = {})` | String | `hyperlink`에 쓸 관계 ID |
| `add_header(header)`, `add_first_header(header)` | Header | 2페이지부터의 머리글과 1페이지 머리글 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | 2페이지부터의 바닥글과 1페이지 바닥글 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>단락(`Docxi::Word::Contents::Paragraph`)

| 메서드 | 추가하는 것 |
|---|---|
| `text(string, options = {})` | 텍스트 런. 단락의 옵션이 기본값이 됩니다 |
| `br(times: 1)` | 줄 바꿈 |
| `tab(times: 1)` | 탭 문자 |
| `image(media, options = {})` | 인라인 이미지 또는 위치를 지정한 이미지 |
| `hyperlink(text, url, id)` | `add_hyperlink`가 반환한 관계 `id`를 쓰는 링크 |
| `frame(options = {}) { \|frame\| }` | 텍스트 상자 |

## <a id="tables"></a>표

| 메서드 | 추가하는 것 |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | 행. 표의 옵션이 기본값이 됩니다 |
| `TableRow#tc(options = {}) { \|tc\| }` | 셀. 행의 옵션이 기본값이 됩니다 |
| `TableCell#text(string, options = {})` | 단락. 셀의 옵션이 기본값이 됩니다 |
| `TableCell#p(options = {}) { \|p\| }` | 자기 옵션을 가진 단락 |
| `TableCell#image(media, width:, height:)` | 이미지 |
| `TableCell#page_numbers` | 페이지 번호 필드 |

## <a id="table-of-contents"></a>목차

| 메서드 | 추가하는 것 |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | 목차 항목 하나 |

## <a id="headers-and-footers"></a>머리글과 바닥글

| 클래스 | 메서드 |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>프레임

`Docxi::Word::Contents::Frame`의 블록 안에서는 `text`, `br`, `tab`, `image`, `frame`을 쓸 수 있습니다.

## <a id="constants-and-helpers"></a>상수와 헬퍼

| 이름 | 의미 |
|---|---|
| `Docxi::VERSION` | 설치된 버전, 예: `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | 크기를 미리 설정한 뒤 파트 하나를 아카이브에 씁니다. 모든 파트가 내부적으로 사용합니다 |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | ActiveSupport의 `blank?` 규칙을 따르는 내부 헬퍼 |

## <a id="related-pages"></a>관련 페이지

- [기능](Features)
- [사용 예제](Usage-Examples)
- [아키텍처](Architecture)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
