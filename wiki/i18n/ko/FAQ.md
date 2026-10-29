<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 자주 묻는 질문: Ruby와 Rails의 Word 문서

가장 많이 받는 질문에 대한 짧은 답변입니다. 각 답변에는 더 자세한 설명으로 가는 링크가 있습니다.

**Microsoft Word나 LibreOffice를 설치해야 하나요?**
아니요. docxi가 XML과 zip 아카이브를 직접 씁니다.

**docxi에 Rails가 필요한가요?**
아니요. 어떤 Ruby 프로그램에서도 동작합니다. 테스트는 Rails 5.2부터 8.1까지 로드한 상태에서도 실행합니다. [통합](Integrations)을 참고하세요.

**Ruby on Rails에서 Word 문서를 만들려면 어떻게 하나요?**
컨트롤러나 백그라운드 작업에서 문서를 만든 다음 `send_data`로 `word.render.string`을 보내세요. [통합](Integrations)을 참고하세요.

**어떤 Ruby와 Rails 버전을 지원하나요?**
Ruby 2.5 이상입니다. Rails는 선택 사항이며, CI에서 Rails 5.2부터 8.1까지 테스트합니다. [설치](Installation)를 참고하세요.

**docxi로 기존 .docx를 편집하거나 Word 템플릿을 채울 수 있나요?**
아니요, 새 문서를 만들기만 합니다. 병합 필드가 있는 템플릿이 필요하다면 sablon 같은 템플릿 기반 gem을 사용하세요.

**목차가 제목에서 자동으로 만들어지나요?**
아니요. 각 항목과 그 페이지 번호는 직접 추가합니다. [기능](Features#table-of-contents)을 참고하세요.

**번호 매기기 목록이나 제목 스타일을 쓸 수 있나요?**
0.1.0에서는 쓸 수 없습니다. 단락을 글머리 기호 항목(`ul: true`)으로 만들 수 있고, 제목은 `size:`와 `bold:` 같은 텍스트 옵션으로 서식을 지정할 수 있습니다.

**페이지 크기나 방향을 바꿀 수 있나요?**
아직은 안 됩니다. 페이지는 세로 방향의 US Letter이며, 바꿀 수 있는 것은 여백뿐입니다. [구성](Configuration)을 참고하세요.

**어떤 이미지 형식이 동작하나요?**
PNG입니다. 모든 이미지는 PNG 파트로 저장됩니다.

**URL에서 이미지를 추가할 수 있나요?**
네. `add_media`는 `http`, `https`, `ftp` URL을 받으며, `render`를 호출할 때 내려받습니다. 사용자에게서 받은 URL은 확인하지 않고 전달하지 마세요. [보안](Security)을 참고하세요.

**문서를 파일에 바로 쓸 수 있나요?**
`render`는 `StringIO`를 반환합니다. 그 `.string`을 `File.binwrite("out.docx", word.render.string)`로 쓰세요.

**docxi로 아주 큰 문서를 스트리밍할 수 있나요?**
아니요. `render`를 호출하면 문서 전체가 메모리에서 만들어집니다. [배포](Deployment)를 참고하세요.

**오른쪽에서 왼쪽으로 쓰는 단락을 설정할 수 있나요?**
0.1.0에는 오른쪽에서 왼쪽으로 쓰는 옵션이 없습니다. 텍스트는 전달한 그대로 기록됩니다.

**알려진 문제는 어디에 정리되어 있나요?**
[로드맵](Roadmap)과 README의 "Limitations and known issues"(제한 사항과 알려진 문제) 섹션에 있습니다.

## <a id="related-pages"></a>관련 페이지

- [문제 해결](Troubleshooting)
- [기능](Features)
- [시작하기](Getting-Started)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
