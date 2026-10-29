<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 시작하기: Ruby로 만드는 첫 Word 문서

이 가이드는 빈 폴더에서 시작해 서식이 적용된 `.docx` 파일을 만드는 과정을 안내합니다. Ruby 기초를 알고 있다고 가정합니다. Rails나 Microsoft Word는 필요하지 않습니다.

## <a id="1-install-the-gem"></a>1. gem 설치하기

```sh
gem install docxi
```

Gemfile이 있는 프로젝트에서는 대신 `gem "docxi", "~> 0.1"`을 추가하고 `bundle install`을 실행하세요. 자세한 내용은 [설치](Installation)에 있습니다.

## <a id="2-create-a-document"></a>2. 문서 만들기

다음 내용을 `hello.rb`로 저장합니다.

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

`ruby hello.rb`(Bundler 프로젝트에서는 `bundle exec ruby hello.rb`)를 실행하고 `hello.docx`를 엽니다.

## <a id="3-what-each-line-does"></a>3. 각 줄이 하는 일

- `Docxi::Document.new`는 기본 여백을 가진 빈 문서를 만듭니다.
- `word.document`는 문서의 본문입니다. 추가하는 모든 것은 이것을 거칩니다.
- `document.p`는 단락을 추가합니다. `align:`처럼 `p`에 전달한 옵션은 단락과 그 안의 텍스트에 적용됩니다.
- `p.text`는 텍스트 런(run)을 하나 추가합니다. `size:`의 단위는 포인트입니다.
- `word.render`는 `.docx` 패키지를 만들고 `StringIO`를 반환합니다. 그 `.string`에 파일의 바이트가 들어 있습니다.

## <a id="4-add-a-list-and-a-table"></a>4. 목록과 표 추가하기

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

주의할 점이 두 가지 있습니다.

- `space: true`는 런의 시작이나 끝에 있는 공백을 유지합니다. 이 옵션이 없으면 Word가 공백을 지워 단어가 서로 붙어 버립니다.
- `tr`의 옵션은 그 행의 모든 셀에 적용되므로, 머리글 행 전체에 음영과 굵게가 적용됩니다.

## <a id="5-where-to-go-next"></a>5. 다음 단계

- [사용 예제](Usage-Examples): 보고서, 송장, 편지의 완전한 예제
- [기능](Features): 단락, 텍스트, 표, 이미지, 머리글과 바닥글의 모든 옵션
- [통합](Integrations): Rails 컨트롤러에서 문서 보내기
- [모범 사례](Best-Practices): 흔한 실수를 미리 피하기

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
