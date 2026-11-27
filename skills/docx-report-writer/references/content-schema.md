# content.json 형식

```jsonc
{
  "meta": {
    "doc_type": "전공 탐구 보고서",          // 표지 상단 작은 글씨
    "title": "보고서 제목",
    "subtitle": "부제",
    "short_title": "머리글에 넣을 짧은 제목",
    "file_topic": "파일명용주제",           // → 파일명용주제_보고서_YYYYMMDD.docx
    "department": "소속",
    "author": "작성자",
    "date": "2026. 11. 27.",
    "keywords": ["키워드1", "키워드2"]
  },
  "summary": ["핵심 요약 문장 [@ref1]", "..."],   // '요약' 장의 강조 박스 (3~5개)
  "sections": [                                   // 장 번호는 자동 (1., 2., ...)
    {
      "heading": "서론",
      "blocks": [ /* 블록 목록 */ ],
      "subsections": [                            // 절 번호는 자동 (1.1, 1.2, ...)
        { "heading": "절 제목", "blocks": [ /* ... */ ] }
      ]
    }
  ],
  "references": [
    { "id": "ref1", "authors": "J. Kim et al.", "title": "논문 제목",
      "source": "Journal, vol. 1, no. 2, pp. 3–4", "year": 2024,
      "url": "https://doi.org/10.xxxx/xxxx" }      // 필수: 클릭 가능한 링크
  ]
}
```

## 인라인 문법 (모든 텍스트에서 사용 가능)
| 문법 | 결과 |
|---|---|
| `[@ref1]`, `[@ref1; @ref2]` | `[1]`, `[1, 2]`: 참고문헌으로 이동하는 링크. 번호는 첫 인용 순서대로 자동 부여 |
| `{tbl:id}`, `{fig:id}` | `표 2`, `그림 1`: 번호 자동 |
| `**텍스트**` | 굵게 |

## 블록 종류
| type | 필드 | 설명 |
|---|---|---|
| `paragraph` | `text` | 양쪽 정렬 본문 문단 |
| `bullets` | `items[]` | 글머리 기호 목록 |
| `callout` | `title`, `items[]` | 강조 박스 |
| `table` | `id`, `caption`, `columns[]`, `rows[][]`, `widths[]`(비율), `align[]`(left/center/right), `note` | 표 (위 캡션, 머리행 반복) |
| `chart` | `id`, `caption`, `categories[]`, `series[{name, values[], labels[]}]`, `unit`, `ylabel`, `ymax`, `note` | 막대 차트 (matplotlib로 그림) |
| `flow` | `id`, `caption`, `steps[{title, desc}]` | 단계 흐름도 (최대 6단계 권장, `desc`의 `\n`은 줄바꿈) |
| `image` | `id`, `caption`, `path`(content.json 기준 상대경로), `width_cm` | 그림 파일 삽입 |

## 생성 시 자동 검증 (오류로 중단)
- 참고문헌에 없는 인용 키를 썼을 때
- 인용되지 않은 참고문헌이 있을 때 (R12)
- 참고문헌에 http(s) URL이 없을 때 (R12)
- 존재하지 않는 `{tbl:}` / `{fig:}`를 참조했을 때
