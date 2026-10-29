<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 문제 해결: 설치 오류, 런타임 오류, 잘못된 출력

아래에서 증상을 찾아보세요. 오래된 앱에서 생기는 문제는 대부분 docxi 0.0.x 때문이며, 0.1.0으로 업그레이드하면 코드를 바꾸지 않고 해결됩니다.

## <a id="installation"></a>설치

**Rails 6 이상 앱에서 Bundler가 호환되는 `nokogiri` 버전을 찾지 못함**
docxi 0.0.6은 `nokogiri < 1.7`로 고정되어 있는데, Rails 6.0 이후 모든 버전은 nokogiri 1.8.5 이상이 필요합니다. `gem "docxi", "~> 0.1"`을 사용하세요.

**Bundler가 호환되는 `rubyzip` 버전을 찾지 못함**
docxi 0.0.6은 `rubyzip ~> 0.9.9`로 고정되어 있어, 더 새로운 rubyzip이 필요한 gem과 충돌합니다. docxi 0.1.0은 rubyzip 1.3부터 3.x까지 받아들입니다.

**Bundler가 계속 0.0.6을 설치함**
Gemfile이 docxi를 0.0.x로 고정하고 있거나(예: `"~> 0.0.6"`), Ruby가 0.1.0에 필요한 2.5보다 오래된 경우입니다. Gemfile의 해당 줄을 `gem "docxi", "~> 0.1"`로 바꾸고 `ruby -v`를 확인하세요.

**nokogiri 컴파일이 실패함**
플랫폼에 맞는 미리 컴파일된 nokogiri gem이 없어서 소스에서 빌드하는 중입니다. C 컴파일러를 설치하고 [nokogiri 설치 가이드](https://nokogiri.org/tutorials/installing_nokogiri.html)를 따르세요.

## <a id="errors-while-building-a-document"></a>문서를 만드는 중의 오류

**Rails 밖에서 `NoMethodError: undefined method 'present?'` 발생**
docxi 0.0.6은 ActiveSupport가 로드되어 있다고 가정했습니다. 0.1.0에는 필요하지 않습니다.

**`tab`, `frame`, 머리글의 `page_numbers`를 호출할 때 `NameError` 발생**
이 기능들은 0.0.6에서 망가져 있었고 0.1.0에서는 동작합니다. `document.frame("text")`도 마찬가지입니다.

**이미지를 추가할 때 `Errno::ENOENT` 발생**
프로그램이 실행되는 디렉터리를 기준으로 그 경로가 존재하지 않습니다. Ruby 3에서 docxi 0.0.6을 쓰면 이미지 URL도 같은 방식으로 실패하는데, 0.0.6이 `Kernel#open`을 사용했기 때문입니다. 0.1.0은 URL을 `URI#open`으로 내려받습니다.

## <a id="documents-that-dont-look-right"></a>제대로 보이지 않는 문서

**Google Docs나 LibreOffice에서 파일이 열리지 않음**
rubyzip 3과 함께 쓰는 docxi 0.0.6은 모든 항목에 Zip64 레코드를 쓰는데, 이 프로그램들은 이를 거부합니다. 0.1.0은 이 레코드 없이 아카이브를 씁니다.

**첫 페이지에 머리글이나 바닥글이 없음**
docxi는 첫 페이지에 별도의 머리글과 바닥글을 둡니다. `add_first_header`와 `add_first_footer`로 설정하세요. [구성](Configuration#first-page)을 참고하세요.

**여백이 빠졌거나 페이지 레이아웃이 이상함**
여백을 설정할 때는 `top`, `right`, `bottom`, `left`를 함께 전달하세요. 빠진 값은 빈 값으로 기록됩니다.

**텍스트 런 사이의 공백이 사라짐**
Word는 런의 시작이나 끝에 있는 공백을 지웁니다. 해당 `text` 호출에 `space: true`를 추가하세요.

**머리글, 표 셀, 또는 `pH`/`pV`를 쓴 이미지가 아주 작거나 정사각형이 됨**
0.1.0의 알려진 문제입니다. 두 방향의 크기를 모두 `width * height * 14.92` EMU로 계산합니다. `style:`로 크기를 지정한 인라인 이미지를 쓰거나, [모범 사례](Best-Practices#images)의 해결 방법을 참고하세요.

**같은 문서를 두 번째로 렌더링하면 이미지가 사라짐**
열린 `File`이나 다른 IO에서 추가한 이미지는 `render` 중에 읽히므로, 두 번째 렌더링은 빈 스트림을 읽습니다. 대신 경로를 전달하거나 한 번만 렌더링하세요.

**`table`에 설정한 테두리가 모든 셀 주위에 나타남**
알려진 문제입니다. 표 수준의 `borders:`는 표의 바깥 테두리가 아니라 모든 셀에 적용됩니다. 바깥 테두리만 그리려면 가장자리 셀에 바깥쪽 면에 대해서만 `borders:`를 설정하세요.

**표 셀의 페이지 번호가 "GlobalOptions"로 시작함**
`TableCell#page_numbers`의 알려진 문제입니다. 대신 페이지 번호를 바닥글이나 머리글에 넣으세요.

**JPEG나 GIF 이미지가 이상하게 보이거나 보이지 않음**
모든 이미지는 실제 형식과 관계없이 PNG 콘텐츠 형식의 PNG 파트로 저장되므로, 프로그램에 따라 다른 형식을 표시하지 못할 수 있습니다. 추가하기 전에 이미지를 PNG로 변환하세요.

## <a id="inspecting-a-generated-file"></a>생성된 파일 살펴보기

`.docx`는 zip 아카이브이므로 안을 들여다볼 수 있습니다.

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

문제가 계속되면 문제를 재현하는 짧은 스크립트와 Ruby, docxi, nokogiri, rubyzip의 버전을 함께 적어 [이슈를 열어 주세요](https://github.com/hammad417/docxi/issues).

## <a id="related-pages"></a>관련 페이지

- [자주 묻는 질문](FAQ)
- [모범 사례](Best-Practices)
- [설치](Installation)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
