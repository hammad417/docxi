<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 로드맵과 알려진 문제

docxi에는 릴리스 일정도, 확정된 로드맵도 없습니다. 기록된 예정 작업은 아래의 알려진 문제를 고치는 것뿐입니다. 이 수정은 0.2.0에 예정되어 있으며 날짜는 정해지지 않았습니다. 고치면 기존 사용자가 이미 생성하고 있는 문서가 바뀌기 때문에 0.1.0에는 넣지 않았습니다.

## <a id="known-issues-planned-for-020"></a>0.2.0에서 고칠 예정인 알려진 문제

| 문제 | 0.1.0에서의 해결 방법 |
|---|---|
| 위치를 지정한 이미지, 머리글 이미지, 표 셀 이미지가 두 방향의 크기를 모두 `width * height * 14.92` EMU로 계산함 | `style:`로 크기를 지정한 인라인 이미지를 사용하거나 [모범 사례](Best-Practices#images)를 참고 |
| 표 수준의 `borders:`가 표의 바깥 테두리가 아니라 모든 셀에 적용됨 | 바깥 테두리만 필요하면 가장자리 셀의 바깥쪽 면에 `borders:`를 설정 |
| 페이지 번호 필드 코드에서 `\* MERGEFORMAT`의 백슬래시가 빠짐 | 없음 |
| 표 셀의 페이지 번호 앞에 "GlobalOptions"가 붙음 | 페이지 번호를 머리글이나 바닥글에 넣기 |
| 이미지가 항상 PNG 파트로 저장되고, `[Content_Types].xml`이 처음 두 개의 머리글과 바닥글만 선언함 | PNG 이미지를 쓰고, 머리글과 바닥글은 각각 두 개까지만 사용 |
| 여백을 일부만 전달하면 나머지가 빈 값으로 기록됨 | `top`, `right`, `bottom`, `left`를 함께 전달 |
| `docProps/core.xml`이 현지 시각에 UTC 표시를 붙여 기록하고 작성자를 기록하지 않음 | 없음 |
| `Docxi.test`가 `NoMethodError`를 일으킴 | 호출하지 마세요. 예전 버전에서 남은 헬퍼입니다 |
| `page: true` 없는 `document.br`과 `document.tab`이 올바르지 않은 XML을 만듦 | 단락 안에서 `p.br`과 `p.tab` 사용 |

## <a id="not-planned"></a>계획에 없는 것

다음 항목은 위 목록에 없으며, 아무것도 결정되지 않았습니다.

- 기존 문서 편집이나 템플릿 채우기
- 번호 매기기 목록, 제목 스타일, 자동 목차
- 다른 페이지 크기나 가로 방향 페이지

## <a id="related-pages"></a>관련 페이지

- [변경 이력](Changelog)
- [기여하기](Contributing)
- [자주 묻는 질문](FAQ)

---

[← 위키 홈](Home) · [README](https://github.com/hammad417/docxi#readme)
