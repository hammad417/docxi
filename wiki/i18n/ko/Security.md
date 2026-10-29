<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 보안: 문제 신고와 신뢰할 수 없는 입력 처리

이 페이지에서는 보안 문제를 신고하는 방법과, 앱이 사용자 입력을 docxi에 전달할 때 알아 두어야 할 보안 사항을 설명합니다.

## <a id="reporting-a-vulnerability"></a>취약점 신고

보안 문제는 공개 이슈로 신고하지 마세요. 대신 메인테이너에게 이메일을 보내 주세요. 주소는 [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec)에 있습니다. docxi 버전과 문제를 재현하는 간단한 방법을 함께 적어 주세요.

## <a id="upgrade-from-00x"></a>0.0.x에서 업그레이드하세요

docxi 0.0.x는 이미지 경로 문자열을 `Kernel#open`에 전달합니다. `Kernel#open`은 `|`로 시작하는 문자열을 셸 명령으로 실행하므로, 이미지 경로를 조작할 수 있는 공격자는 서버에서 명령을 실행할 수 있습니다. 0.1.0은 파일을 `File.open`으로 열고 URL을 `URI#open`으로 내려받습니다. 둘 다 명령을 실행하지 않습니다.

## <a id="untrusted-input"></a>신뢰할 수 없는 입력

**텍스트는 안전하게 전달할 수 있습니다.** 텍스트, 표 내용, 그 밖의 문자열은 XML 텍스트나 속성 값으로 기록되며 nokogiri가 이스케이프합니다. 사용자 입력으로 문서에 요소를 추가할 수 없습니다.

**이미지 소스는 주의가 필요합니다.** `add_media`가 받는 것은 다음과 같습니다.

| 소스 | docxi의 동작 |
|---|---|
| 직접 연 IO | 그것을 읽습니다 |
| 파일 경로 | 서버 디스크에서 그 파일을 읽습니다 |
| `http`, `https`, `ftp` URL | 서버에서 내려받습니다 |

사용자가 경로를 고를 수 있다면, 설정이나 키처럼 서버에서 읽을 수 있는 아무 파일이나 문서에 포함시킬 수 있습니다. 사용자가 URL을 고를 수 있다면 서버가 그 URL로 요청을 보내므로, 내부 주소에도 닿을 수 있습니다(서버 측 요청 위조, SSRF). 직접 연 IO를 전달하거나, 알려진 경로와 호스트만 허용하세요.

**하이퍼링크**는 URL이 가리키는 곳으로 연결됩니다. 사용자가 준 URL은 `add_hyperlink`에 전달하기 전에 확인하세요.

## <a id="how-releases-are-protected"></a>릴리스 보호 방법

- 릴리스는 GitHub Actions에서 RubyGems의 신뢰할 수 있는 게시(trusted publishing)로 게시되므로, 저장소에는 API 키를 보관하지 않습니다.
- gem 메타데이터에 `rubygems_mfa_required`가 설정되어 있으므로, 개인 API 키로 푸시하려면 다단계 인증이 필요합니다.

## <a id="related-pages"></a>관련 페이지

- [모범 사례](Best-Practices)
- [배포](Deployment)
- [기여하기](Contributing)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
