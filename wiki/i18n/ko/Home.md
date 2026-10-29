<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 위키: Ruby로 Word(.docx) 문서 만들기

docxi는 코드로 Microsoft Word `.docx` 파일을 만드는 Ruby gem입니다. 단락, 서식 있는 텍스트, 표, 이미지, 머리글과 바닥글, 페이지 번호, 하이퍼링크, 간단한 목차를 지원합니다. Microsoft Word나 LibreOffice, 템플릿이 필요 없으며, 순수 Ruby와 Ruby on Rails 애플리케이션 모두에서 동작합니다.

[README](https://github.com/hammad417/docxi#readme)는 짧은 소개입니다. 이 위키는 모든 옵션을 설명하고, 바로 복사해 쓸 수 있는 예제를 제공하며, docxi의 내부 동작을 설명하고, 문제가 생겼을 때 도움을 줍니다.

현재 릴리스: **0.1.1**, Ruby 2.5 이상 지원. [변경 이력](Changelog)을 참고하세요.

## <a id="start-here"></a>여기서 시작하세요

**docxi를 처음 사용한다면**

- [시작하기](Getting-Started): gem을 설치하고 첫 문서를 만듭니다
- [설치](Installation): 요구 사항, Bundler, 0.0.x에서 업그레이드
- [사용 예제](Usage-Examples): 데이터로 만드는 보고서, 송장, 편지, 표. 바로 복사해 쓸 수 있습니다

**문서 만들기**

- [기능](Features): 모든 요소와 각 요소가 받는 옵션
- [구성](Configuration): 여백, 페이지 설정, 단위
- [API 참조](API-Reference): 클래스와 메서드 한눈에 보기
- [모범 사례](Best-Practices): 흔한 실수를 피하는 습관

**문제가 생겼을 때**

- [문제 해결](Troubleshooting): 오류 메시지와 제대로 보이지 않는 문서
- [자주 묻는 질문](FAQ): 자주 묻는 질문에 대한 짧은 답변
- [보안](Security): 문제 신고와 신뢰할 수 없는 입력 처리

**더 알아보기**

- [아키텍처](Architecture): Ruby 호출이 `.docx` 패키지가 되는 과정
- [통합](Integrations): Rails 컨트롤러, Rack 앱, 백그라운드 작업
- [배포](Deployment): 프로덕션에서 docxi 운영하기와 릴리스 게시 방법
- [기여하기](Contributing)와 [개발 환경 설정](Development-Setup)
- [로드맵](Roadmap), [변경 이력](Changelog), [크레딧](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi가 하는 일과 하지 않는 일

docxi는 새 문서를 만듭니다. 기존 `.docx`를 열거나 편집할 수 없고, Word 템플릿을 채우거나 편지 병합을 실행할 수도 없습니다. 번호 매기기 목록, 제목 스타일, 자동 목차가 없으며, 페이지는 항상 세로 방향의 US Letter입니다. 0.2.0에서 고칠 예정인 알려진 문제는 [로드맵](Roadmap)에 정리되어 있습니다.

## <a id="credits"></a>크레딧

docxi는 [Phaedra Solutions](https://www.phaedrasolutions.com/)의 [Hammad Maqbool](https://github.com/hammad417)이 유지 관리합니다. 2016년에 Timur Gabdrakipov의 [docxer](https://github.com/gtimti/docxer)를 바탕으로 시작되었습니다. 자세한 내용은 [크레딧](Credits)을 참고하세요.

## <a id="languages"></a>언어

이 위키는 영어판이 기준입니다. 번역본 링크는 사이드바와 각 페이지 상단에 있습니다.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [문제 신고](https://github.com/hammad417/docxi/issues)
