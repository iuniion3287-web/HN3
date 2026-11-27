---
name: docx-report-writer
description: 정해진 작성·디자인 규칙(R1~R16)에 따라 한국어 학술/전공 보고서를 Word(.docx)로 작성하고 자동 검증한다. 표지·자동 목차·핵심 요약 박스·번호 붙은 장·표/그림 캡션·클릭 가능한 인용과 참고문헌 링크·"현재 쪽 / 전체 쪽" 쪽 번호를 갖춘 문서를 만든다. 사용자가 "보고서 써줘", "DOCX/워드 보고서", "전공 탐구 보고서", "레포트 작성" 등을 요청할 때 사용한다.
---

# DOCX 보고서 작성 스킬

이 스킬은 **내용(content.json)** 과 **형식(스크립트)** 을 분리한다.
Claude는 조사와 글쓰기에 집중하고, 서식·번호·링크·쪽 번호는 스크립트가 규칙대로 일관되게 만든다.
결과물은 `scripts/check_report.py`로 규칙 준수 여부를 기계적으로 검사한다.

- 규칙 전문: [references/report-rules.md](references/report-rules.md) ← **작성 전에 반드시 읽는다**
- 입력 형식: [references/content-schema.md](references/content-schema.md)
- 빈 템플릿: [templates/content.template.json](templates/content.template.json)

## 준비

```bash
pip install -r requirements.txt     # python-docx, matplotlib, (선택) pywin32, pymupdf
```
- Windows + Microsoft Word가 있으면 `finalize_word.py`가 목차·쪽 번호를 채우고 PDF를 만든다.
- Word가 없으면 이 단계를 건너뛴다. DOCX를 Word에서 처음 열 때 "필드 업데이트"에 [예]를 누르면 목차가 채워진다.

## 작업 순서 (반드시 이 순서로)

### 1단계 · 요구사항 정리
주제, 독자(예: 전공 교수), 분량(기본 8~10쪽), 작성자/소속, 작성일을 정한다.
사용자가 지정하지 않은 값은 기본값을 쓰고, 마지막 보고 때 무엇을 기본값으로 정했는지 밝힌다.

### 2단계 · 자료 조사와 출처 검증 (R12, R14)
- 모든 수치와 사실 주장에는 출처가 있어야 한다. **출처를 확인하지 못한 수치는 쓰지 않는다.**
- 학술 논문은 DOI로 확인한다. Crossref(`https://api.crossref.org/works/{DOI}`)에서 저자, 제목, 학술지, 권(호), 쪽, 연도를 그대로 가져온다.
- 본문에 쓰는 숫자는 초록이나 원문에서 직접 확인한다(PubMed E-utilities 등).
- 웹 문서는 URL이 실제로 열리는지(HTTP 200) 확인한다. 열리지 않는 출처는 다른 공신력 있는 출처로 바꾼다.

### 3단계 · content.json 작성
- 형식은 `references/content-schema.md`를 따른다. 장 번호, 표·그림 번호, 인용 번호는 **직접 쓰지 않는다**(스크립트가 자동으로 붙인다).
- 인용은 `[@키]` 또는 `[@키1; @키2]`, 표·그림 참조는 `{tbl:id}` / `{fig:id}`, 강조는 `**굵게**`로 쓴다.
- 구성(R1): 요약(summary, 3~5개 핵심 문장) → 서론 → 본론(2~4장) → 결론 및 시사점 → 참고문헌(자동).
- 문체(R14): '~다' 평서문, 전문용어는 처음 나올 때 영문 병기 `표적 단백질(target protein)`, 약어는 처음 나올 때 풀어 쓴다.

### 4단계 · DOCX 생성
```bash
python scripts/build_report.py content.json            # → {file_topic}_보고서_{YYYYMMDD}.docx (R15)
```
다음 경우 생성이 오류로 멈춘다. 오류 메시지대로 content.json을 고친다.
- 참고문헌에 없는 인용 키를 썼을 때
- 인용되지 않은 참고문헌이 있을 때
- 참고문헌에 URL이 없을 때
- 존재하지 않는 표·그림을 참조했을 때

### 5단계 · Word로 마무리 (가능한 경우)
```bash
python scripts/finalize_word.py 보고서.docx            # 목차·쪽 번호 갱신 + 보고서.pdf
```

### 6단계 · 규칙 자동 검사
```bash
python scripts/check_report.py 보고서.docx --links    # FAIL이 0개가 될 때까지 고친다
```

### 7단계 · 눈으로 확인 (자동 검사가 못 잡는 것)
```bash
python scripts/render_preview.py 보고서.pdf -o preview/
```
`preview/overview.png`와 각 쪽 이미지를 직접 열어 아래를 확인한다.
- [ ] 표지가 한 쪽에 들어가고 2쪽으로 넘치지 않는다.
- [ ] 목차의 쪽 번호가 실제 쪽과 일치한다.
- [ ] 표 머리행이나 캡션만 쪽 끝에 홀로 남지 않는다.
- [ ] 그림 주석이 다음 쪽으로 밀려나지 않는다.
- [ ] 차트 글자가 깨지지 않고(한글 폰트), 값 라벨이 원문 수치와 같다.
- [ ] 빈 쪽이 없다.

문제가 있으면 content.json이나 스크립트를 고친 뒤 4단계부터 다시 한다.

### 8단계 · 보고
사용자에게 다음을 알린다.
- 생성된 파일 경로(DOCX, PDF)와 쪽 수
- 검사 결과(PASS/FAIL 수)
- 기본값으로 정한 항목
- 확인하지 못해 제외한 정보가 있다면 그 내용

## 파일 구성
| 파일 | 역할 |
|---|---|
| `scripts/build_report.py` | content.json → DOCX (규칙 R1~R15 구현) |
| `scripts/finalize_word.py` | Word 자동화: 목차·필드 갱신, PDF 내보내기 |
| `scripts/check_report.py` | 규칙 자동 검사기 (`--links`로 링크 접속 확인) |
| `scripts/render_preview.py` | PDF → 쪽별 PNG + 한눈에 보기 이미지 |
| `references/report-rules.md` | 작성·디자인 규칙 R1~R16 |
| `references/content-schema.md` | content.json 형식 |
| `templates/content.template.json` | 새 보고서용 빈 템플릿 |
