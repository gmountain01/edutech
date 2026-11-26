# -*- coding: utf-8 -*-
"""3권 진행 현황 사이트 데이터 생성 → index.html 의 /*DATA*/ 구간을 갈아 끼운다.
근거: 00_자료조사/교사AI도서_전수조사_20261028.xlsx(도서), 00_기획서/toc_cache.jsonl + 01_자료조사/toc_extra.jsonl(예스24 목차),
      01_자료조사/3·4권_시장조사_에듀테크·이미지동영상_20261125.docx(툴 정보)
실행: python build_site.py
"""
import json, re
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

# 툴: 시장조사 4장 표 + 추가 후보. a = 목차에서 찾을 이름(별칭)
TOOLS = [
 # 이름, 분류, 핵심 AI 기능, 국내/해외, 교사 무료 범위, 학생 연령·계정, 에듀집, 별칭, 링크
 ("구글 클래스룸", "LMS·수업 운영", "Gemini, Gems·NotebookLM을 학생에게 배정", "해외", "Workspace for Education 전 에디션 무료", "학교 계정", "등재", ["클래스룸", "Classroom"], None),
 ("Gemini for Education", "업무 비서", "학교 계정용 Gemini. 개인 계정과 데이터 처리가 다름", "해외", "Workspace for Education 포함", "개인 만 14+, 학교 계정은 관리자 설정", "등재(구글 묶음)", ["제미나이", "Gemini", "제미니"], None),
 ("NotebookLM", "자료 요약·리서치", "올린 자료 근거로 답변, 오디오·비디오 개요", "해외", "교육 에디션 무료", "클래스룸 학생 개인 노트북은 18+·고등교육만", "미확인(구글 묶음)", ["노트북LM", "NotebookLM", "노트북 LM"], None),
 ("MS Copilot Chat / Teams", "업무 비서", "Copilot Chat, 학습 에이전트", "해외", "A1·A3·A5 무료 (M365 Copilot은 월 18달러)", "만 13+", "M365 등재, Copilot 미확인", ["코파일럿", "Copilot", "Teams", "팀즈"], None),
 ("아이쌤GPT", "업무 비서(국산)", "HWP 작성·편집, 행정 자동화", "국내(아이스크림)", "미확인", "학생용 2026 하반기 예정", "등재", ["아이쌤"], None),
 ("테크빌 마이클(MyCl)", "업무 비서(국산)", "공문 서식 40종+, 여러 LLM 선택", "국내(테크빌)", "미확인 (약 2,700교, 인천 3만 명)", "학생판 준비 중", "등재", ["마이클", "MyCl"], None),
 ("클로바노트", "기록·요약", "음성 기록·요약", "국내(네이버)", "월 300분 무료", "미확인", "네이버 묶음", ["클로바노트", "클로바 노트", "CLOVA Note"], None),
 ("Brisk Teaching", "업무 비서", "크롬 확장. 피드백·지도안 등 20+ 도구", "해외", "Educator 무료(일일 한도)", "미확인", "등재", ["Brisk", "브리스크"], None),
 ("MagicSchool", "자료·문항 생성", "교사 도구 80종+", "해외", "교사 무료 플랜", "미확인", "등재", ["매직스쿨", "MagicSchool", "Magic School"], None),
 ("Diffit", "자료·문항 생성", "같은 지문을 수준별로, 활동지", "해외", "Basic 무료", "미확인", "없음", ["Diffit", "디핏"], None),
 ("Khanmigo 교사용", "업무 비서", "교사용 수업 준비 도구", "해외", "교사 무료(한국 포함)", "학생용은 미국만", "없음", ["Khanmigo", "칸미고"], None),
 ("캔바 교육용", "자료·발표", "Magic Write (이미지·영상 생성은 4권)", "해외", "K-12 무료", "미확인", "등재", ["캔바", "Canva"], None),
 ("감마", "자료·발표", "AI 슬라이드 생성", "해외", "가입 시 400크레딧 1회", "미확인", "없음", ["감마", "Gamma"], None),
 ("패들렛", "협업 보드", "Magic Padlet, AI 토론 보드", "해외", "무료 보드 3개", "미확인", "등재", ["패들렛", "Padlet"], None),
 ("북크리에이터", "전자책", "이미지 생성 (글쓰기 AI는 미확인)", "해외", "미확인", "미확인", "등재", ["북크리에이터", "북 크리에이터", "Book Creator"], None),
 ("띵커벨", "퀴즈·평가", "AI 문항 추천", "국내(아이스크림)", "미확인", "QR 참여", "등재", ["띵커벨"], None),
 ("ZEP 퀴즈", "퀴즈·평가", "PPT·PDF·유튜브로 AI 문항 생성", "국내", "AI는 Premium", "미확인", "등재", ["ZEP 퀴즈", "젭 퀴즈", "ZEP퀴즈"], None),
 ("카훗", "퀴즈·평가", "AI 문항 생성", "해외", "AI는 유료", "미확인", "등재", ["카훗", "Kahoot"], None),
 ("블루킷", "퀴즈·평가", "Khanmigo 연동 문항 생성", "해외", "기본 무료", "미확인", "등재", ["블루킷", "Blooket"], None),
 ("퀴즈앤", "퀴즈·평가", "AI 기능 미확인", "국내", "미확인", "미확인", "등재", ["퀴즈앤", "QuizN"], None),
 ("클래스팅 AI", "코스웨어", "AI 진단·튜터", "국내", "유료 (3,500교+)", "미확인", "등재", ["클래스팅"], None),
 ("EBS AI 단추플러스", "코스웨어(공공)", "진단→개념→심화→총괄, 학습분석 대시보드", "공공", "무료", "미확인", "등재", ["단추"], None),
 ("EBS AI 펭톡", "영어(공공)", "AI 영어 대화", "공공", "무료", "미확인", "등재", ["펭톡"], None),
 ("서울 SEN스쿨 / 채움아이", "교육청 플랫폼", "AI 맞춤 학습 / AI 서·논술형 평가(2026 110교 → 2027 전면)", "공공", "무료", "학교 계정", "SEN스쿨 등재", ["SEN스쿨", "채움아이"], None),
 ("경기 하이러닝", "교육청 플랫폼", "AI 서·논술형 평가 등", "공공", "무료", "학교 계정", "등재", ["하이러닝"], None),
 ("부산 AI 튜터 BeAT", "교육청 플랫폼", "미확인", "공공", "무료", "미확인", "등재", ["BeAT"], None),
 ("KERIS AI·디지털 교육자료 포털", "공공", "AIDT 이후 교육자료, 똑똑! 수학탐험대(초등 수학 AI)", "공공", "무료", "학교 계정", "등재", ["수학탐험대", "KERIS", "교육자료 포털"], None),
 ("하이클래스 / 클래스123", "학급경영", "AI 기능 미확인", "국내", "무료", "미확인", "등재", ["하이클래스", "클래스123"], None),
 ("뤼튼", "범용", "AI 기능 세부 미확인", "국내", "미확인", "미확인", "등재", ["뤼튼", "wrtn"], None),
 ("ZEP", "메타버스", "AI 기능 세부 미확인", "국내", "미확인", "미확인", "등재", ["ZEP", "젭"], None),
 ("웨이그라운드(옛 퀴지즈)", "퀴즈·평가", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["퀴지즈", "Quizizz", "웨이그라운드", "Wayground"], None),
 ("멘티미터", "참여·설문", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["멘티미터", "Mentimeter"], None),
 ("니어팟", "수업 진행", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["니어팟", "Nearpod"], None),
 ("에드퍼즐", "영상 수업", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["에드퍼즐", "Edpuzzle"], None),
 ("워드월", "퀴즈·게임", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["워드월", "Wordwall"], None),
 ("SchoolAI", "학생 AI", "AI 세부 미확인", "해외", "미확인", "미확인", "등재", ["SchoolAI", "스쿨AI"], None),
 ("클래스카드", "단어 학습", "AI 세부 미확인", "국내", "미확인", "미확인", "등재", ["클래스카드"], None),
]

SKIP_CAT = {"바이브코딩"}

# 공식 링크 (2026-11-26 확인). SEN스쿨/채움아이는 SEN스쿨 주소(채움아이 단독 주소 없음), BeAT는 공개 주소가 없어 부산교육청 보도자료
URLS = {
 "구글 클래스룸": "https://edu.google.com/intl/ALL_kr/products/classroom/",
 "Gemini for Education": "https://edu.google.com/ai/gemini-for-education/",
 "NotebookLM": "https://notebooklm.google",
 "MS Copilot Chat / Teams": "https://www.microsoft.com/en-us/education/products/copilot-in-education",
 "클로바노트": "https://clovanote.line.me/",
 "Brisk Teaching": "https://www.briskteaching.com",
 "MagicSchool": "https://www.magicschool.ai/",
 "Diffit": "https://web.diffit.me/",
 "Khanmigo 교사용": "https://www.khanmigo.ai/teachers",
 "캔바 교육용": "https://www.canva.com/ko_kr/education/",
 "감마": "https://gamma.app/",
 "패들렛": "https://padlet.com/",
 "북크리에이터": "https://bookcreator.com/",
 "띵커벨": "https://www.i-screammedia.com/www/business_edubank.html",
 "ZEP 퀴즈": "https://quiz.zep.us/",
 "카훗": "https://kahoot.com/ko/schools/kahoot-edu/",
 "블루킷": "https://www.blooket.com/",
 "퀴즈앤": "https://quizn.show/",
 "클래스팅 AI": "https://www.classting.ai/",
 "EBS AI 단추플러스": "https://ai-plus.ebs.co.kr",
 "EBS AI 펭톡": "https://pengtalk-student.ebse.co.kr/",
 "경기 하이러닝": "https://hi.goe.go.kr",
 "KERIS AI·디지털 교육자료 포털": "https://www.edunet.net",
 "하이클래스 / 클래스123": "https://www.hiclass.net",
 "뤼튼": "https://wrtn.ai/",
 "ZEP": "https://zep.us",
 "웨이그라운드(옛 퀴지즈)": "https://wayground.com",
 "멘티미터": "https://www.mentimeter.com/",
 "니어팟": "https://nearpod.com/",
 "에드퍼즐": "https://edpuzzle.com/",
 "워드월": "https://wordwall.net/",
 "SchoolAI": "https://www.schoolai.com",
 "클래스카드": "https://www.classcard.net",
 "아이쌤GPT": "https://issamgpt.com/",
 "테크빌 마이클(MyCl)": "https://mycl.io/",
 "서울 SEN스쿨 / 채움아이": "https://senedu.kr/",
 "부산 AI 튜터 BeAT": "https://www.pen.go.kr/main/na/ntt/selectNttInfo.do?mi=30397&nttSn=1150888",
}

# 화면에서 묶는 큰 분류 (TOOLS 의 세부 분류 → 묶음)
GROUP = {
 "업무 비서": "업무 비서·기록", "업무 비서(국산)": "업무 비서·기록", "기록·요약": "업무 비서·기록", "자료 요약·리서치": "업무 비서·기록", "범용": "업무 비서·기록",
 "자료·문항 생성": "수업 준비·자료", "자료·발표": "수업 준비·자료", "전자책": "수업 준비·자료",
 "협업 보드": "수업 참여·퀴즈", "퀴즈·평가": "수업 참여·퀴즈", "참여·설문": "수업 참여·퀴즈", "수업 진행": "수업 참여·퀴즈", "퀴즈·게임": "수업 참여·퀴즈", "영상 수업": "수업 참여·퀴즈", "메타버스": "수업 참여·퀴즈",
 "LMS·수업 운영": "수업 운영·학급경영", "학급경영": "수업 운영·학급경영", "학생 AI": "수업 운영·학급경영",
 "코스웨어": "맞춤학습·코스웨어", "코스웨어(공공)": "맞춤학습·코스웨어", "영어(공공)": "맞춤학습·코스웨어", "단어 학습": "맞춤학습·코스웨어",
 "교육청 플랫폼": "공공·교육청 플랫폼", "공공": "공공·교육청 플랫폼",
}


def load_books():
    ws = openpyxl.load_workbook(ROOT / "00_자료조사/교사AI도서_전수조사_20261028.xlsx", read_only=True)["전수조사DB"]
    rows = list(ws.iter_rows(values_only=True))
    H = rows[0]
    return [dict(zip(H, r)) for r in rows[1:]]


def load_toc():
    toc = {}
    for p in [ROOT / "00_기획서/toc_cache.jsonl", HERE / "01_자료조사/toc_extra.jsonl"]:
        for line in p.open(encoding="utf-8"):
            d = json.loads(line)
            # 예전 캐시는 목차 뒤에 예스24 '펼쳐보기접어보기' 버튼·스크립트가 붙어 있음
            t = (d.get("toc") or "").split("펼쳐보기")[0].lstrip(">").strip()
            if t:
                toc[d["url"]] = t
    return toc


# 목차 정리: 글은 원문 그대로, 기호·쪽 번호·부속 줄(머리말 등)을 걷어내고 줄마다 단계(1~4)를 붙인다 → [[단계, 글], ...]
FRONT = re.compile(r"^(머리말|머리글|추천사|프롤로그|에필로그|들어가며|들어가는 ?(말|글)|시작하며|시작하는 글|마치며|나가며|맺음말|작가의 말|"
                   r"이 책의 (구성|200% 활용법|특징)|책의 구성|더 알아보기|찾아보기|독자지원센터|참고문헌|저자 소개|감사의 글)(\s|$)")
LEAD = re.compile(r"^(┗|└|ㄴ|[-–·•●▶■□○*]|_+)\s*")
PAGE = re.compile(r"\s+[0-9]{1,3}(\s[0-9]{1,2})?$")


KINDS = [  # (단계 우선순위, 패턴) 위에서부터 먼저 맞는 것. 책마다 나온 것만 1~4단계로 다시 매긴다
    (0, re.compile(r"^\[?\s*(PART|Part|part)\s*\.?\s*([0-9]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)|^\[?\s*(제\s*)?[0-9]+\s*부(?![가-힣])")),
    (1, re.compile(r"^(CHAPTER|Chapter|chapter)\s*\.?\s*[0-9]+|^(제\s*)?[0-9]+\s*장(?![가-힣])|^부록")),
    (2, re.compile(r"^(SECTION|Section|LESSON|Lesson|_?Class|UNIT|Unit)\b|^[0-9]+\s*-\s*[0-9]+(?![0-9])|^[0-9]+\.[0-9]+\s")),
    (3, re.compile(r"^[0-9]{1,3}\s*[.|]?\s+\S|^(실습|Step|STEP)\s*[0-9]+")),
    (3.5, re.compile(r"^\[[^\]]+\]")),
    (4, re.compile(r"^([가-하]\.|[0-9]+\)|[①-⑳]|※|팁:)")),
]


def clean_toc(text):
    raw = [l.rstrip() for l in text.split("\n") if l.strip()]
    # 쪽 번호가 줄 끝에 붙은 책(30% 이상)만 떼기
    strip_pg = sum(bool(PAGE.search(l)) for l in raw) >= 0.3 * max(1, len(raw))
    rows = []  # [표기 단계 또는 None(바로 위 줄의 아래), 글]
    for l in raw:
        indent = len(l) - len(l.lstrip())
        t = l.strip()
        if re.fullmatch(r"[=\-_~·.\s]+", t):
            continue
        t = t.strip("|").strip()
        if strip_pg:
            t = PAGE.sub("", t)
        lead = LEAD.match(t)
        t = LEAD.sub("", t)
        t = re.sub(r"\s*\|\s*", " · ", t).strip(" ·")
        t = re.sub(r"\s{2,}", " ", t)
        if not t or FRONT.match(t):
            continue
        # 기호·들여쓰기로 시작한 줄과 번호 없는 줄은 표기 단계 없이 '바로 위 줄의 아래'
        kind = None if lead or indent >= 2 else next((k for k, p in KINDS if p.match(t)), None)
        rows.append([kind, t])
    # 표기 단계는 이 책에 나온 것만 1부터 다시 매기고, 나머지 줄은 위 표기 줄 + 1 (최대 5)
    rank = {k: i + 1 for i, k in enumerate(sorted({k for k, _ in rows if k is not None}))}
    out, cur = [], 0
    for k, t in rows:
        if k is None:
            out.append([min(cur + 1, 5), t])
        else:
            cur = min(rank[k], 4)
            out.append([cur, t])
    return out


def _check_toc():
    r = clean_toc("머리말\nPART 01. 시작\nPreview 들어가기\n01 챗GPT 시작하기\n    ┗ 대비해요.\n2장 둘째\n가. 세부\n===\n| 시작 | 기준\n찾아보기")
    assert r == [[1, "PART 01. 시작"], [2, "Preview 들어가기"], [3, "01 챗GPT 시작하기"], [4, "대비해요."],
                 [2, "2장 둘째"], [4, "가. 세부"], [5, "시작 · 기준"]], r


def hits(text, aliases):
    # 'ZEP'이 'ZEP 퀴즈' 안에서도 잡히는 건 허용(ZEP 퀴즈는 ZEP 서비스라서)
    return any(a in text for a in aliases)


def main():
    _check_toc()
    toc = load_toc()
    books = []
    for r in load_books():
        if r["구분"] != "A_본조사" or r["주제분류"] in SKIP_CAT:
            continue
        t = toc.get(r["예스24 링크"]) or ""
        hay = " ".join(str(x or "") for x in (r["도서명"], r["부제"], r["주요 툴"], t))
        used = [i for i, tool in enumerate(TOOLS) if hits(hay, tool[7])]
        if "에듀테크" not in (r["도서명"] or "") and len(used) < 2:
            continue
        books.append({
            "id": r["No"], "title": r["도서명"], "sub": r["부제"] or "", "author": r["저자"] or "",
            "authorType": r["저자유형"] or "", "pub": r["출판사"] or "", "date": str(r["출간일"] or "")[:10],
            "pages": r["쪽수"], "sales": r["판매지수"] or 0, "level": r["학교급"] or "", "cat": r["주제분류"],
            "summary": r["책소개(요약)"] or "", "url": r["예스24 링크"] or "", "toc": clean_toc(t), "tools": used,
        })
    books.sort(key=lambda b: -(b["sales"] or 0))
    tools = [dict(zip(["name", "cat", "ai", "origin", "free", "age", "zip", "alias", "url"], t)) for t in TOOLS]
    for i, tool in enumerate(tools):
        tool["url"] = tool["url"] or URLS.get(tool["name"])
        tool["group"] = GROUP.get(tool["cat"], "기타")
        tool["books"] = [b["id"] for b in books if i in b["tools"]]
    data = json.dumps({"tools": tools, "books": books}, ensure_ascii=False, separators=(",", ":"))
    data = data.replace("</", "<\\/")
    page = HERE / "index.html"
    s = page.read_text(encoding="utf-8")
    s, n = re.subn(r"/\*DATA\*/.*?/\*DATA\*/", lambda m: "/*DATA*/" + data + "/*DATA*/", s, flags=re.S)
    assert n == 1, "index.html 에 /*DATA*/ 구간이 없음"
    page.write_text(s, encoding="utf-8")
    # GitHub Pages 용 완성본: 아티팩트가 붙여 주던 문서 틀(doctype·head·기본 리셋)을 직접 붙인다
    (HERE / "docs").mkdir(exist_ok=True)
    (HERE / "docs/index.html").write_text(
        '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<style>body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style></head><body>\n'
        + s + '\n</body></html>\n', encoding="utf-8")
    print(f"툴 {len(tools)}종, 도서 {len(books)}권, 목차 있는 책 {sum(1 for b in books if b['toc'])}권")


if __name__ == "__main__":
    main()
