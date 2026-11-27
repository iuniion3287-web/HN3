#!/usr/bin/env python3
"""DOCX 보고서가 references/report-rules.md 규칙을 지키는지 자동 검사한다.

사용법:
    python check_report.py 보고서.docx [--links] [--json]

    --links  참고문헌·본문의 외부 링크에 실제로 접속해 응답 코드를 확인한다(인터넷 필요).
종료 코드: 0 = 모든 필수 규칙 통과, 1 = 하나 이상 FAIL
"""
import argparse
import json
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

results = []


def report(rule, ok, msg, level="FAIL"):
    results.append({"rule": rule, "status": "PASS" if ok else level, "message": msg})


def ptext(p_el):
    return "".join(t.text or "" for t in p_el.iter(W + "t"))


def style_name(doc, p_el):
    ps = p_el.find(f"{W}pPr/{W}pStyle")
    if ps is None:
        return "Normal"
    sid = ps.get(W + "val")
    try:
        return doc.styles.get_by_id(sid, 1).name  # 1 = paragraph
    except Exception:
        return sid


def instr_texts(el):
    out = [it.text or "" for it in el.iter(W + "instrText")]
    out += [fs.get(W + "instr") or "" for fs in el.iter(W + "fldSimple")]
    return " ".join(out)


def check(path, check_links=False):
    doc = Document(path)
    body = doc.element.body
    paras = list(body.iter(W + "p"))
    top = list(body)  # top-level block order
    rels = doc.part.rels

    # ---------------- R15 파일명 & 문서 속성
    name = os.path.basename(path)
    report("R15", re.match(r".+_보고서_\d{8}\.docx$", name) is not None,
           f"파일명 '{name}' (형식: 주제_보고서_YYYYMMDD.docx)")
    cp = doc.core_properties
    report("R15", bool(cp.title and cp.author), f"문서 속성 제목='{cp.title}', 작성자='{cp.author}'")

    # ---------------- R5 용지·여백
    secs = doc.sections
    ok = all(abs(s.page_width.mm - 210) < 1 and abs(s.page_height.mm - 297) < 1 for s in secs)
    report("R5", ok, "모든 구역 A4(210×297mm)")
    m = [round(getattr(s, k).mm, 1) for s in secs for k in ("top_margin", "bottom_margin", "left_margin", "right_margin")]
    report("R5", all(abs(x - 25) <= 1 for x in m), f"여백 25mm (실측: {sorted(set(m))})")

    # ---------------- R6 본문 글꼴
    normal = doc.styles["Normal"]
    rpr = normal.element.find(f"{W}rPr")
    ea = rpr.find(f"{W}rFonts").get(W + "eastAsia") if rpr is not None and rpr.find(f"{W}rFonts") is not None else None
    size = normal.font.size.pt if normal.font.size else None
    if size is None:  # Word는 기본값과 같은 크기를 Normal에서 생략하고 docDefaults에만 남긴다
        sz = doc.styles.element.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}sz")
        size = int(sz.get(W + "val")) / 2 if sz is not None else None
    report("R6", ea == "맑은 고딕" and size == 11, f"본문(Normal) 글꼴 {ea}, {size}pt (기준: 맑은 고딕 11pt)")
    ls = normal.paragraph_format.line_spacing
    report("R6", ls is not None and 1.3 <= float(ls) <= 1.6, f"본문 줄 간격 {ls} (기준 1.4)")

    # ---------------- R2 표지 / R8 머리글·바닥글
    report("R2", len(secs) >= 2, f"표지와 본문이 별도 구역으로 분리됨 (구역 수 {len(secs)})")
    if len(secs) >= 2:
        s0 = secs[0]
        cover_hf = "".join(p.text for p in s0.header.paragraphs + s0.footer.paragraphs).strip()
        cover_fields = instr_texts(s0.footer._element) if not s0.footer.is_linked_to_previous else ""
        report("R2", cover_hf == "" and "PAGE" not in cover_fields, "표지에는 머리글·쪽 번호가 없음")
        sb = secs[1]
        htxt = " ".join(p.text for p in sb.header.paragraphs).strip()
        report("R8", bool(htxt), f"본문 머리글에 보고서 제목: '{htxt}'")
        f_instr = instr_texts(sb.footer._element)
        report("R8", "PAGE" in f_instr and ("SECTIONPAGES" in f_instr or "NUMPAGES" in f_instr),
               f"바닥글에 '현재 쪽 / 전체 쪽' 필드 (발견: {f_instr.split()})")
        align = sb.footer.paragraphs[0].alignment
        report("R8", align is not None and int(align) == 1, "쪽 번호 가운데 정렬")
        pg = sb._sectPr.find(f"{W}pgNumType")
        report("R8", pg is not None and pg.get(W + "start") == "1", "본문 쪽 번호가 1부터 시작")

    # ---------------- R3 목차
    all_instr = instr_texts(body)
    has_toc = re.search(r"\bTOC\b", all_instr) is not None
    report("R3", has_toc and "\\h" in all_instr, "자동 목차(TOC) 필드, 항목 클릭 시 이동(\\h)")
    toc_links = [h for h in body.iter(W + "hyperlink") if (h.get(W + "anchor") or "").startswith("_Toc")]
    report("R3", len(toc_links) > 0, f"목차가 채워져 있음 (목차 항목 {len(toc_links)}개)", level="WARN")

    # ---------------- R4/R7 제목
    heads = [(style_name(doc, p), ptext(p).strip()) for p in paras]
    h1 = [t for s, t in heads if s == "Heading 1"]
    h2 = [t for s, t in heads if s == "Heading 2"]
    report("R7", len(h1) >= 3, f"제목 1 수준 {len(h1)}개, 제목 2 수준 {len(h2)}개")
    bad1 = [t for t in h1 if t not in ("요약", "참고문헌") and not re.match(r"^\d+\. \S", t)]
    bad2 = [t for t in h2 if not re.match(r"^\d+\.\d+ \S", t)]
    report("R4", not bad1 and not bad2, "제목 번호 형식 '1.' / '1.1'" + (f" 위반: {bad1 + bad2}" if bad1 or bad2 else ""))
    nums = [int(t.split(".")[0]) for t in h1 if re.match(r"^\d+\.", t)]
    report("R4", nums == list(range(1, len(nums) + 1)), f"장 번호 연속 {nums}")
    report("R1", bool(h1) and h1[-1] == "참고문헌", "보고서 마지막 장은 '참고문헌'")
    report("R1", "요약" in h1 and h1.index("요약") == 0, "본문 첫 장은 '요약'(핵심 요약)")

    # ---------------- R13 핵심 요약 박스
    shaded = [p for p in paras if p.find(f"{W}pPr/{W}shd") is not None]
    report("R13", len(shaded) >= 2, f"음영 강조 박스 문단 {len(shaded)}개")

    # ---------------- R9 표 캡션·머리행
    tables = [el for el in top if el.tag == W + "tbl"]
    tbl_ok, tbl_msgs = True, []
    for i, tbl in enumerate(tables, 1):
        prev = tbl.getprevious()
        cap = ptext(prev).strip() if prev is not None and prev.tag == W + "p" else ""
        if not cap.startswith(f"표 {i}."):
            tbl_ok = False
            tbl_msgs.append(f"{i}번째 표 위 캡션 없음/번호 불일치('{cap[:20]}')")
        first_tr = tbl.find(f"{W}tr")
        if first_tr is None or first_tr.find(f"{W}trPr/{W}tblHeader") is None:
            tbl_ok = False
            tbl_msgs.append(f"표 {i} 머리행 반복 미설정")
        if first_tr is not None and first_tr.find(f".//{W}shd") is None:
            tbl_ok = False
            tbl_msgs.append(f"표 {i} 머리행 음영 없음")
    report("R9", tbl_ok and len(tables) > 0, f"표 {len(tables)}개: 위쪽 캡션 '표 n.', 머리행 반복·음영" + (" / " + "; ".join(tbl_msgs) if tbl_msgs else ""))

    # ---------------- R10 그림 캡션
    fig_ok, fig_msgs, k = True, [], 0
    for p in paras:
        if p.find(f".//{W}drawing") is None:
            continue
        if p.getparent().tag != W + "body":
            continue
        k += 1
        nxt = p.getnext()
        cap = ptext(nxt).strip() if nxt is not None else ""
        if not cap.startswith(f"그림 {k}."):
            fig_ok = False
            fig_msgs.append(f"{k}번째 그림 아래 캡션 없음('{cap[:20]}')")
    report("R10", fig_ok, f"그림 {k}개: 아래쪽 캡션 '그림 n.'" + (" / " + "; ".join(fig_msgs) if fig_msgs else ""))

    # ---------------- R11 / R12 인용과 참고문헌
    ref_idx = next((i for i, p in enumerate(paras) if style_name(doc, p) == "Heading 1" and ptext(p).strip() == "참고문헌"), None)
    report("R12", ref_idx is not None, "참고문헌 장 존재")
    if ref_idx is not None:
        body_ps, ref_ps = paras[:ref_idx], [p for p in paras[ref_idx + 1:] if ptext(p).strip()]
        # bookmarks & numbering
        ref_nums, ext_ok, bm_ok = [], True, True
        bm_names = {b.get(W + "name") for b in body.iter(W + "bookmarkStart")}
        for p in ref_ps:
            m = re.match(r"^\[(\d+)\]", ptext(p).strip())
            if not m:
                continue
            n = int(m.group(1))
            ref_nums.append(n)
            links = [h for h in p.iter(W + "hyperlink") if h.get(R_ID)]
            urls = [rels[h.get(R_ID)].target_ref for h in links if h.get(R_ID) in rels]
            if not any(u.startswith(("http://", "https://")) for u in urls):
                ext_ok = False
            if f"ref_{n}" not in bm_names:
                bm_ok = False
        report("R12", ref_nums == list(range(1, len(ref_nums) + 1)) and ref_nums,
               f"참고문헌 번호 [1]~[{len(ref_nums)}] 연속")
        report("R12", ext_ok, "모든 참고문헌에 클릭 가능한 외부 링크(DOI/URL)")
        report("R11", bm_ok, "모든 참고문헌 항목에 인용 이동용 북마크(ref_n)")

        # citations in body
        cited, unlinked = [], []
        for p in body_ps:
            for h in p.iter(W + "hyperlink"):
                a = h.get(W + "anchor") or ""
                if a.startswith("ref_"):
                    cited.append(int(a[4:]))
            plain = "".join(t.text or "" for t in p.iter(W + "t") if not any(anc.tag == W + "hyperlink" for anc in t.iterancestors()))
            unlinked += re.findall(r"\[\d+(?:,\s*\d+)*\]", plain)
        report("R11", not unlinked, "본문 인용 [n]이 모두 클릭 가능한 내부 링크" + (f" (링크 없는 인용: {unlinked[:5]})" if unlinked else ""))
        missing = sorted(set(cited) - set(ref_nums))
        unused = sorted(set(ref_nums) - set(cited))
        report("R11", not missing, "본문 인용 번호가 모두 참고문헌에 존재" + (f" (없음: {missing})" if missing else ""))
        report("R12", not unused, "인용되지 않은 참고문헌 없음" + (f" (미인용: {unused})" if unused else ""))
        first_order = []
        for n in cited:
            if n not in first_order:
                first_order.append(n)
        report("R12", first_order == sorted(first_order), "참고문헌 번호가 본문 첫 인용 순서와 일치")
        report("R11", len(cited) >= len(ref_nums), f"본문 인용 {len(cited)}회 / 참고문헌 {len(ref_nums)}개")

    # ---------------- R16 링크 실제 접속
    if check_links:
        import urllib.request
        import ssl
        urls = sorted({r.target_ref for r in rels.values() if r.is_external and r.target_ref.startswith("http")})
        bad = []
        ctx = ssl.create_default_context()
        for u in urls:
            try:
                req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0 (report-link-check)"})
                with urllib.request.urlopen(req, timeout=25, context=ctx) as resp:
                    if resp.status >= 400:
                        bad.append(f"{resp.status} {u}")
            except Exception as e:
                code = getattr(e, "code", None)
                # 일부 출판사는 자동 접속을 403으로 막지만 브라우저에서는 열린다
                if code in (401, 403, 429):
                    continue
                bad.append(f"{code or type(e).__name__} {u}")
        report("R16", not bad, f"외부 링크 {len(urls)}개 접속 확인" + (f" / 실패: {bad}" if bad else ""))

    return results


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--links", action="store_true")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    res = check(a.docx, a.links)
    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        for r in res:
            mark = {"PASS": "✔", "FAIL": "✘", "WARN": "!"}[r["status"]]
            print(f"{mark} [{r['rule']:>3}] {r['status']:<4} {r['message']}")
        n_fail = sum(r["status"] == "FAIL" for r in res)
        n_warn = sum(r["status"] == "WARN" for r in res)
        print(f"\n총 {len(res)}개 검사: PASS {len(res) - n_fail - n_warn}, WARN {n_warn}, FAIL {n_fail}")
    sys.exit(1 if any(r["status"] == "FAIL" for r in res) else 0)


if __name__ == "__main__":
    main()
