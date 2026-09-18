# -*- coding: utf-8 -*-
"""
상담·행정 채용공고 수집기 (청소년상담사 지원용)

하는 일
  1. 잡알리오 오픈API(공공기관)와 워크넷 오픈API(민간·지자체 포함)에서
     수도권(서울·경기·인천) 공고를 가져온다
  2. 상담 / 청소년·복지 / 교육기관 행정직만 남긴다
  3. '청소년상담사' 자격이 요건·우대에 적힌 공고에 ⭐ 를 붙인다
  4. 한 장짜리 웹페이지(docs/index.html)와 엑셀을 만든다

두 곳 다 인증키가 필요하다. 없는 쪽은 건너뛰고 있는 쪽만 모은다.
  ALIO_API_KEY  공공데이터포털(data.go.kr) 인증키 — 공공기관 채용정보 조회서비스
  WORK_API_KEY  워크넷(고용24) 오픈API 인증키
키는 깃허브에서는 저장소 시크릿으로, 내 노트북에서는 같은 폴더의 .env 로 준다.

주의
  - 워크넷은 http 로 부른다. https 로는 못 받는다 — 서버 인증서에 Authority Key Identifier 가
    빠져 있어 요즘 OpenSSL 이 검증을 거부한다(이 노트북도, 깃허브 서버도 같다).
    주고받는 값에 개인정보가 없어 http 로도 문제되지 않는다.
  - 로그는 공개 저장소에 그대로 올라간다. 인증키가 붙은 주소는 어떤 경우에도 찍지 않는다.
"""
import datetime as dt
import html
import json
import os
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

# 윈도우 콘솔은 기본이 cp949 라, 한글 로그를 찍다가 프로그램이 죽는다.
# (깃허브 서버는 UTF-8 이라 상관없지만, 내 노트북에서 돌릴 때 필요하다)
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, OSError):
    pass

# ─────────────────────────── 설정 ───────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE, "docs")
FILES_DIR = os.path.join(DOCS_DIR, "files")
DATA_DIR = os.path.join(BASE, "data")
STATE_PATH = os.path.join(DATA_DIR, "state_seen.json")
LOG_PATH = os.path.join(DATA_DIR, "log.txt")
HIST_PATH = os.path.join(DATA_DIR, "history.json")

SITE_URL = "https://hongyul67-cpu.github.io/sangdam-chaeyong/"

# 어느 지역을 볼지. 잡알리오 코드와 워크넷 코드가 서로 다르다.
REGIONS_ALIO = "R3010,R3017,R3011"          # 서울 · 경기 · 인천
REGIONS_WORK = "11000|41000|28000"          # 서울 · 경기 · 인천
REGION_LABEL = "서울 · 경기 · 인천"
# 잡알리오가 지역을 안 걸러 줄 때를 대비한 2차 그물(응답의 근무지 글자로 한 번 더 거른다)
REGION_WORDS = ["서울", "경기", "인천", "전국"]

# 잡알리오 NCS 대분류.
#   사회복지.종교  → 상담·복지 자리가 대부분 여기 들어온다
#   교육.자연.사회과학 → 청소년지도·교육 자리
#   경영.회계.사무 → 행정직. 너무 넓어서, 기관·제목이 상담/청소년/복지/교육에
#                   걸릴 때만 담는다(아래 SAMU_NEEDS_CONTEXT).
NCS_CODES = "R600007,R600004,R600002"

# 며칠 전 공고까지 훑을지. 3~4일에 한 번 도니 넉넉히 잡아도 겹칠 뿐 빠지지 않는다.
LOOKBACK_DAYS = 45

# 오래된 회차 리포트는 지운다.
KEEP_DAYS = 365

# ── 분야 이름표 ──────────────────────────────────────────────
#   공고 글자에서 뽑아 카드 위에 붙인다. 판별이 아니라 이름표라 조금 헐거워도 된다.
TAG_KEYWORDS = {
    "상담": ["상담", "심리", "치료", "카운슬", "정서", "위기개입", "솔루션위원회",
             "자살예방", "학교폭력", "심리검사", "놀이치료", "미술치료"],
    "청소년": ["청소년", "학교밖", "꿈드림", "위센터", "wee", "Wee", "방과후",
               "아동", "학생", "진로", "유스", "수련관", "청소년지도"],
    "복지": ["복지", "사회복지", "가족", "다문화", "돌봄", "자립", "드림스타트",
             "건강가정", "취약계층", "장애"],
    "행정": ["행정", "사무", "총무", "회계", "경리", "서무", "기획", "운영지원",
             "인사", "관리직", "경영지원"],
    "교육": ["교육", "강사", "교사", "평생교육", "학습", "교육지원", "장학"],
}

# ⭐ 이 낱말이 자격·우대에 있으면 '청소년상담사 자격이 쓰이는 자리'로 본다.
LICENSE_WORDS = ["청소년상담사", "청소년 상담사", "청소년상담복지", "청소년지도사",
                 "전문상담", "상담심리사", "임상심리사", "상담사 자격"]
# '사회복지사'는 일부러 뺐다 — 공공기관 공고 자격요건에 워낙 흔해서
# 국민건강보험공단 같은 데까지 상담 자리로 잡혔다(2026-09-18 실제 확인).
LICENSE_STAR = ["청소년상담사", "청소년 상담사"]   # 이건 있으면 별을 붙인다

# ── 담을지 말지를 가르는 두 관문 ─────────────────────────────
# 2026-09-18 에 처음 실제로 돌려 보고 다시 짰다. 그 전에는 공고 글자를 통째로 훑어
# 낱말을 찾았는데, 공공기관 공고문의 자격요건에는 '장애인복지법·국가유공자·행정'
# 같은 상투 문구가 반드시 들어가서 한국도로공사·한국전력까지 '복지·행정'으로
# 잡혔다. 87건 중 82건이 통과해 걸러내기가 없는 것과 같았다.
# 그래서 지금은 **제목과 기관명만** 본다. 자격요건 전문은 ⭐ 판정에만 쓴다.

# ① 하는 일이 상담·청소년인가 — 제목이나 NCS 분류에서 찾는다
JOB_WORDS = ["상담", "심리", "청소년", "동반자", "꿈드림", "학교밖", "위기청소년",
             "사례관리", "정서", "진로", "wee", "Wee", "WEE", "자살예방",
             "학교폭력", "방과후", "돌봄", "멘토"]

# ② 일할 곳이 청소년·상담 기관인가 — 기관명에서 찾는다
#   '복지'만으로는 안 된다. 한국보훈복지의료공단·한국산림복지진흥원까지 걸린다.
#   '여성'·'장학'도 뺐다 — 한국여성과학기술인육성재단·장학재단처럼 상담과
#   무관한 곳이 걸린다. 여성가족 계열은 '가족'으로 잡힌다.
ORG_WORDS = ["청소년", "상담", "가족", "아동", "심리", "복지관",
             "평생학습", "꿈드림", "디딤", "학교", "교육청"]

# 제목에 이 낱말이 있으면 뺀다. 상담사 자리가 아니다.
SKIP_WORDS = ["미화", "경비원", "청소원", "조리", "당직", "운전원", "시설관리원",
              "보안", "수위", "영양사", "간호", "의사", "약사", "사서보조"]

WORK_URL = "http://openapi.work.go.kr/opi/opi/opia/wantedApi.do"
ALIO_API = "https://apis.data.go.kr/1051000/recruitment"
ALIO_VIEW = "https://job.alio.go.kr/recruitview.do?idx={idx}"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

for _d in (DOCS_DIR, FILES_DIR, DATA_DIR):
    os.makedirs(_d, exist_ok=True)


def _load_env():
    """이 폴더에 .env 가 있으면 인증키를 읽어 온다.

    내 노트북에서 시험 삼아 돌릴 때 쓴다. 깃허브에서 돌 때는 저장소 시크릿이
    환경변수로 들어오므로 .env 가 없어도 된다.
    .env 는 .gitignore 에 걸려 있어 저장소에 올라가지 않는다 — 그대로 두어야 한다.

        ALIO_API_KEY=여기에붙여넣기
        WORK_API_KEY=여기에붙여넣기
    """
    path = os.path.join(BASE, ".env")
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8-sig") as f:      # 메모장이 붙이는 BOM 을 벗긴다
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k and v and not os.environ.get(k):     # 환경변수가 이미 있으면 그쪽이 우선
                os.environ[k] = v


_load_env()


def log(msg):
    line = "[%s] %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_stamp():
    """화면에 찍을 '마지막 갱신' 시각. 서버는 UTC 라 TZ 를 켜 두어야 한국시간이 된다."""
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")


# ─────────────────────────── 잡알리오 ───────────────────────────
def _alio_key():
    key = (os.environ.get("ALIO_API_KEY") or "").strip()
    if key and "%" in key:              # Encoding 표기면 한 번 풀어 준다
        key = urllib.parse.unquote(key)  # (아래서 다시 인코딩할 때 이중 인코딩을 막는다)
    return key


def _get(url, timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "ignore")


def alio_get(path, params):
    """잡알리오 오픈API 호출. 인증키가 붙은 주소는 절대 로그에 남기지 않는다."""
    key = _alio_key()
    if not key:
        return None
    p = dict(params)
    p["serviceKey"] = key
    p.setdefault("resultType", "json")
    try:
        return json.loads(_get(ALIO_API + path + "?" + urllib.parse.urlencode(p)))
    except Exception as e:                # noqa: BLE001 — 주소를 찍으면 키가 샌다
        log("  잡알리오 호출 실패(%s): %s" % (path, type(e).__name__))
        return None


def alio_fetch():
    """잡알리오에서 수도권 상담·복지·교육·사무 공고를 모아 온다."""
    if not _alio_key():
        log("잡알리오 인증키(ALIO_API_KEY) 없음 — 건너뜁니다")
        return []
    since = (dt.date.today() - dt.timedelta(days=LOOKBACK_DAYS)).isoformat()
    recs, page = {}, 1
    while page <= 20:
        data = alio_get("/list", {
            "ncsCdLst": NCS_CODES,
            "workRgnLst": REGIONS_ALIO,
            "ongoingYn": "Y",            # 아직 접수 중인 것만
            "pbancBgngYmd": since,       # 최근 것만 (옛 공고가 쏟아지는 걸 막는다)
            "numOfRows": 100,
            "pageNo": page,
        })
        if data is None:
            break
        if str(data.get("resultCode")) not in ("200", "0", "00"):
            log("  잡알리오 오류 %s: %s"
                % (data.get("resultCode"), str(data.get("resultMsg"))[:80]))
            break
        batch = data.get("result") or []
        for r in batch:
            recs[str(r.get("recrutPblntSn"))] = r
        total = int(data.get("totalCount") or 0)
        log("  잡알리오 p%d → %d건 (전체 %d)" % (page, len(batch), total))
        if not batch or len(recs) >= total:
            break
        page += 1
    return [alio_row(r) for r in recs.values()]


def _ymd(s):
    """'20260908' → '2026-09-08'"""
    s = re.sub(r"[^0-9]", "", str(s or ""))
    if len(s) != 8:
        return ""
    try:
        return dt.date(int(s[:4]), int(s[4:6]), int(s[6:])).isoformat()
    except ValueError:
        return ""


def alio_row(rec):
    idx = str(rec.get("recrutPblntSn"))
    상세 = ALIO_VIEW.format(idx=idx)
    자격 = re.sub(r"\s+", " ", rec.get("aplyQlfcCn") or "").strip()
    우대 = re.sub(r"\s+", " ", " ".join([rec.get("prefCn") or "",
                                         rec.get("prefCondCn") or ""])).strip()
    return {
        "_id": "alio:" + idx,
        "출처": "잡알리오(공공기관)",
        "기관명": rec.get("instNm") or "",
        "제목": rec.get("recrutPbancTtl") or "",
        "분야원문": rec.get("ncsCdNmLst") or "",
        "고용형태": rec.get("hireTypeNmLst") or "",
        "근무지": rec.get("workRgnNmLst") or "",
        "학력": rec.get("acbgCondNmLst") or "",
        "인원": str(rec.get("recrutNope") or ""),
        "자격": 자격[:1200],
        "우대": 우대[:800],
        "접수시작": _ymd(rec.get("pbancBgngYmd")),
        "접수마감": _ymd(rec.get("pbancEndYmd")),
        "링크": rec.get("srcUrl") or 상세,
        "상세": 상세,
        "대체인력": "예" if str(rec.get("replmprYn") or "").upper() == "Y" else "",
    }


# ─────────────────────────── 워크넷 ───────────────────────────
# ⚠ 주소가 http 인 것은 실수가 아니다. https 로 부르면 인증서 검증에서 막힌다
#   (CERTIFICATE_VERIFY_FAILED: Missing Authority Key Identifier — 서버 쪽 문제).
#   2026-09-17 에 이 노트북과 curl 양쪽에서 확인했다.
#
# 워크넷은 상담·청소년 자리를 가장 많이 담고 있다(시·군·구 청소년상담복지센터,
# 꿈드림, 건강가정지원센터 등). 다만 검색어를 주지 않으면 수도권 공고가 수만 건이라,
# 낱말로 좁혀서 여러 번 부르고 구인인증번호로 합친다.
WORK_KEYWORDS = ["청소년상담", "청소년지도", "청소년", "전문상담", "상담원",
                 "심리상담", "상담사", "학교밖청소년", "꿈드림", "건강가정",
                 "청소년수련", "위기청소년"]


def _work_key():
    return (os.environ.get("WORK_API_KEY") or "").strip()


def work_fetch():
    """워크넷에서 낱말별로 불러 합친다. 키가 없으면 빈 목록."""
    key = _work_key()
    if not key:
        log("워크넷 인증키(WORK_API_KEY) 없음 — 건너뜁니다")
        return []
    got, bad = {}, 0
    for kw in WORK_KEYWORDS:
        page = 1
        while page <= 5:
            q = urllib.parse.urlencode({
                "authKey": key, "callTp": "L", "returnType": "XML",
                "startPage": page, "display": 100,
                "region": REGIONS_WORK,
                "keyword": kw,
            }, safe="|")
            try:
                xml = _get(WORK_URL + "?" + q, timeout=30)
            except Exception as e:        # noqa: BLE001 — 주소에 키가 들어 있다
                log("  워크넷 호출 실패(%s): %s" % (kw, type(e).__name__))
                bad += 1
                break
            try:
                root = ET.fromstring(xml)
            except ET.ParseError:
                log("  워크넷 응답을 읽지 못했습니다(%s)" % kw)
                bad += 1
                break
            msg = root.findtext("message")
            if msg and root.findtext("messageCd") not in (None, "000"):
                log("  워크넷 오류: %s" % msg)
                return []                 # 인증키 문제면 더 불러 봐야 소용없다
            items = root.findall(".//wanted")
            for it in items:
                r = work_row(it)
                if r:
                    got[r["_id"]] = r
            log("  워크넷 '%s' p%d → %d건" % (kw, page, len(items)))
            if len(items) < 100:
                break
            page += 1
    if bad and not got:
        log("워크넷에서 아무것도 받지 못했습니다")
    return list(got.values())


def _t(el, *names):
    for n in names:
        v = el.findtext(n)
        if v and v.strip():
            return v.strip()
    return ""


def work_row(el):
    no = _t(el, "wantedAuthNo", "empSeqno")
    if not no:
        return None
    마감 = _t(el, "closeDt", "recrutClosDt")
    등록 = _t(el, "regDt", "regDtm")
    return {
        "_id": "work:" + no,
        "출처": "워크넷",
        "기관명": _t(el, "company", "coNm"),
        "제목": _t(el, "title", "wantedTitle"),
        "분야원문": _t(el, "jobsCdNm", "jobsNm", "occupation"),
        "고용형태": _t(el, "empTpNm", "empTpCdNm", "holidayTpNm"),
        "근무지": _t(el, "basicAddr", "region", "workRgnNm"),
        "학력": _t(el, "minEdubg", "maxEdubg"),
        "인원": _t(el, "empPersonCnt", "recrutNope"),
        "자격": re.sub(r"\s+", " ", _t(el, "jobCont", "career", "detailCont"))[:1200],
        "우대": re.sub(r"\s+", " ", _t(el, "prefCond", "etc"))[:800],
        "접수시작": _dashify(등록),
        "접수마감": _dashify(마감),
        "링크": _t(el, "wantedInfoUrl", "wantedMobileInfoUrl") or
                ("https://www.work24.go.kr/wk/a/b/1500/empDetailAuthView.do?wantedAuthNo=" + no),
        "상세": "",
        "대체인력": "",
    }


def _dashify(s):
    """'2026-09-17' 또는 '20260917' 어느 쪽이 와도 '2026-09-17' 로."""
    s = str(s or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    return _ymd(s)


# ────────────────────────── 고르기·이름표 ──────────────────────────
def tags_of(row):
    """분야 이름표. **자격요건 전문은 넣지 않는다** — 상투 문구 때문에 전부 '복지·행정'이 된다."""
    blob = " ".join([row.get("제목", ""), row.get("분야원문", ""), row.get("기관명", "")])
    return [t for t, words in TAG_KEYWORDS.items() if any(w in blob for w in words)]


def license_hit(row):
    """요건·우대에 상담 자격이 적혀 있으면 (별표여부, 찾은 낱말들)."""
    blob = " ".join([row.get("자격", ""), row.get("우대", ""),
                     row.get("제목", ""), row.get("분야원문", "")])
    found = [w for w in LICENSE_WORDS if w in blob]
    star = any(w in blob for w in LICENSE_STAR)
    return star, found


def keep(row):
    """담을지 말지. (담는다, 버린 이유)"""
    title = row.get("제목", "")
    if any(w in title for w in SKIP_WORDS):
        return False, "제외낱말"

    # 지역 2차 그물. 잡알리오는 근무지를 비워 보내는 일이 있어 비었으면 통과시킨다.
    지역 = row.get("근무지", "") + " " + row.get("기관명", "")
    if 지역.strip() and not any(w in 지역 for w in REGION_WORDS):
        return False, "지역밖"

    # ⭐ 공고에 '청소년상담사'가 실제로 적혀 있으면 무조건 담는다. 이 도구의 핵심이다.
    if row["_star"]:
        return True, ""

    # ① 하는 일이 상담·청소년인가 (제목·NCS분류)
    일 = row.get("제목", "") + " " + row.get("분야원문", "")
    if any(w in 일 for w in JOB_WORDS):
        return True, ""

    # ② 일할 곳이 청소년·상담 기관인가 (기관명)
    if any(w in row.get("기관명", "") for w in ORG_WORDS):
        return True, ""

    return False, "무관한 자리"


def normalize(row):
    row["_tags"] = tags_of(row)
    star, found = license_hit(row)
    row["_star"] = star
    row["_lic"] = found
    return row


def dday(마감, today):
    if not 마감:
        return None
    try:
        return (dt.date.fromisoformat(마감) - today).days
    except ValueError:
        return None


# ─────────────────────────── 저장 ───────────────────────────
def save_text(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


XLSX_COLS = ["분야", "청소년상담사", "기관명", "제목", "고용형태", "근무지", "학력",
             "인원", "접수시작", "접수마감", "남은날", "자격요건", "우대", "출처", "링크"]


def write_xlsx(rows, path, today, title="공고"):
    """전달용 엑셀. 가로 방향 · 너비 1페이지로 인쇄되게 맞춰 둔다."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        log("  openpyxl 이 없어 엑셀을 건너뜁니다")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = title[:30]
    ws.append(XLSX_COLS)
    head_fill = PatternFill("solid", fgColor="1F2937")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        n = dday(r.get("접수마감"), today)
        ws.append([
            " · ".join(r["_tags"]),
            "⭐" if r["_star"] else ("○" if r["_lic"] else ""),
            r.get("기관명", ""), r.get("제목", ""), r.get("고용형태", ""),
            r.get("근무지", ""), r.get("학력", ""), r.get("인원", ""),
            r.get("접수시작", ""), r.get("접수마감", ""),
            ("마감" if n is not None and n < 0 else
             ("오늘" if n == 0 else ("D-%d" % n if n is not None else ""))),
            r.get("자격", "")[:500], r.get("우대", "")[:300],
            r.get("출처", ""), r.get("링크", ""),
        ])
    widths = [14, 11, 22, 46, 12, 16, 12, 6, 11, 11, 8, 60, 34, 16, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    # 인쇄: 가로 방향 · 너비 1페이지 (세로는 몇 장이 되든 둔다)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    wb.save(path)


# ─────────────────────────── 화면 ───────────────────────────
CSS = """body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;margin:24px;color:#111;background:#fff}
h1{font-size:20px;margin:0 0 4px}.sub{color:#666;font-size:13px;margin-bottom:4px;line-height:1.6}
.stamp{color:#374151;font-size:12.5px;margin:0 0 14px}
h3{font-size:15px;margin:24px 0 10px;padding-bottom:5px;border-bottom:2px solid #111}
.card{border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:12px}
.card h2{font-size:15px;margin:0 0 6px;line-height:1.45}
.tag{display:inline-block;background:#dbeafe;color:#1e40af;border-radius:99px;
 padding:1px 9px;font-size:12px;margin-right:5px}
.tag.lic{background:#fef3c7;color:#92400e;font-weight:700}
.src{display:inline-block;background:#f3f4f6;color:#4b5563;border-radius:99px;
 padding:1px 9px;font-size:11.5px;margin-right:5px}
.d{color:#b91c1c;font-weight:700}
table{border-collapse:collapse;font-size:13px;margin-top:8px}
td{padding:2px 10px 2px 0;vertical-align:top}td.k{color:#666;white-space:nowrap}
.none{color:#666;padding:20px;border:1px dashed #d1d5db;border-radius:10px;text-align:center}
a{color:#1d4ed8}
.warn{border:2px solid #b91c1c;background:#fef2f2;color:#7f1d1d;border-radius:10px;
 padding:12px 16px;margin:0 0 16px;font-size:13px;line-height:1.65}
.warn b{font-size:14px}
.bar{margin:0 0 18px;display:flex;gap:8px;flex-wrap:wrap}
.bar button{font-family:inherit;font-size:13px;padding:8px 16px;border:1px solid #111;
 background:#111;color:#fff;border-radius:8px;cursor:pointer}
.bar .btnlink{font-size:13px;padding:8px 16px;border:1px solid #111;background:#fff;
 color:#111;border-radius:8px;text-decoration:none;display:inline-block}
.bar .btnlink:hover{background:#f3f4f6}
.tabs{display:flex;gap:6px;border-bottom:2px solid #111;margin:0 0 4px;flex-wrap:wrap}
.tabs button{font-family:inherit;font-size:14px;font-weight:700;padding:9px 18px;
 border:1px solid #d1d5db;border-bottom:none;background:#f3f4f6;color:#6b7280;
 border-radius:9px 9px 0 0;cursor:pointer;position:relative;top:2px}
.tabs button.on{background:#111;color:#fff;border-color:#111}
table.sum{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
table.sum th{background:#f3f4f6;border:1px solid #d1d5db;padding:6px 8px;
 text-align:left;white-space:nowrap;font-weight:700}
table.sum td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:middle}
table.sum tr.urgent{background:#fef2f2}
table.sum tr.star{background:#fffbeb}
table.sum td a{text-decoration:none;font-weight:700}
table.sum td a:hover{text-decoration:underline}
table.sum tr.gone{background:#f3f4f6;color:#9ca3af}
table.sum tr.gone td{text-decoration:line-through}
table.sum tr.gone td.st{text-decoration:none;color:#6b7280;font-weight:700}
.card.gone{background:#f9fafb;border-color:#e5e7eb;opacity:.62}
.card.gone h2{text-decoration:line-through;color:#9ca3af}
.card.star{border-color:#f59e0b;border-width:2px}
.gonetag{display:inline-block;background:#6b7280;color:#fff;border-radius:99px;
 padding:1px 9px;font-size:12px;font-weight:700;margin-left:6px}
.today{display:inline-block;background:#b91c1c;color:#fff;border-radius:99px;
 padding:1px 9px;font-size:12px;font-weight:700}
.newtag{display:inline-block;background:#059669;color:#fff;border-radius:99px;
 padding:1px 9px;font-size:11.5px;font-weight:700;margin-left:6px}
.asof{margin:10px 0 0;font-size:12.5px;color:#374151;background:#fffbeb;
 border:1px solid #fde68a;border-radius:8px;padding:8px 10px}
.hint{color:#666;font-size:12px;margin:6px 0 0}
/* 폰에서는 8열이 다 들어가지 않아 기관명이 한 글자씩 쪼개진다.
   덜 중요한 열(#·분야·고용형태)을 감추고 기관·제목에 자리를 준다. */
@media(max-width:640px){body{margin:12px}
 table.sum{font-size:12.5px}
 table.sum th,table.sum td{padding:6px 5px}
 table.sum th:nth-child(1),table.sum td:nth-child(1),
 table.sum th:nth-child(4),table.sum td:nth-child(4),
 table.sum th:nth-child(6),table.sum td:nth-child(6){display:none}
 table.sum td:nth-child(2){min-width:5.2em}
 table.sum th:nth-child(7),table.sum td:nth-child(7){white-space:nowrap;font-size:11.5px}
 .card h2{font-size:14.5px}
 td.k{width:4.6em}}
@media print{body{margin:0;font-size:11pt}.bar,.noprint,.tabs{display:none!important}
 [hidden]{display:block!important}
 .card{break-inside:avoid;page-break-inside:avoid;border:1px solid #999}
 .warn{border:2px solid #000;background:#fff;color:#000}
 h3{break-after:avoid}a{color:#000;text-decoration:none}}"""

# 만든 날 박아 둔 D-day 를 그대로 두면 사흘 뒤 열었을 때 이미 마감된 공고가
# '지금 지원 가능'으로 남는다. 열어 보는 날 기준으로 브라우저에서 다시 계산한다.
RECALC_JS = r"""<script>
(function(){
  function ymd(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
  var now=new Date(); var today=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  function left(v){ if(!/^\d{4}-\d{2}-\d{2}$/.test(v)) return null;
    var p=v.split('-'); var d=new Date(+p[0],+p[1]-1,+p[2]);
    return Math.round((d-today)/86400000); }
  var gone=0, open=0;
  document.querySelectorAll('tr[data-dl]').forEach(function(tr){
    var n=left(tr.getAttribute('data-dl')); var cell=tr.querySelector('td.st'); if(n===null) return;
    if(n<0){ gone++; tr.classList.add('gone'); tr.classList.remove('urgent');
      if(cell){cell.textContent=''; var b=document.createElement('span');
        b.className='gonetag'; b.textContent='마감됨'; cell.appendChild(b);} }
    else { open++;
      if(cell){ cell.textContent='';
        if(n===0){ var t=document.createElement('span'); t.className='today';
          t.textContent='오늘 마감'; cell.appendChild(t); }
        else { cell.textContent='D-'+n; } }
      if(n<=7) tr.classList.add('urgent'); }
  });
  document.querySelectorAll('.card[data-dl]').forEach(function(c){
    var n=left(c.getAttribute('data-dl')); if(n===null) return;
    var d=c.querySelector('.d');
    if(n<0){ c.classList.add('gone');
      if(d){ d.textContent=''; var b=document.createElement('span'); b.className='gonetag';
        b.textContent='마감됨 — 지금은 지원할 수 없습니다'; d.appendChild(b); } }
    else if(d){ d.textContent = (n===0? '오늘 마감' : 'D-'+n); }
  });
  var sub=document.querySelector('.stamp');
  if(sub){ var p=document.createElement('div'); p.className='asof';
    p.innerHTML='📅 <b>오늘('+ymd(today)+') 기준으로 다시 계산했습니다.</b> '+
      (gone? '이 중 <b>'+gone+'건은 이미 마감</b>되어 회색으로 표시했습니다. ':'')+
      '마감일·자격은 기관 공고가 기준이며, 이 표와 다를 수 있습니다.';
    sub.parentNode.insertBefore(p, sub.nextSibling); }
})();
</script>"""

esc = lambda s: html.escape(str(s or ""))


def sort_key(r, today):
    """마감 임박순. 마감일이 없는 건 뒤로."""
    n = dday(r.get("접수마감"), today)
    return (n is None, n if n is not None else 9999, r.get("기관명", ""))


def _cards(parts, rows, today, new_ids):
    for r in rows:
        n = dday(r.get("접수마감"), today)
        label = ("마감" if n is not None and n < 0 else
                 ("오늘 마감" if n == 0 else ("D-%d" % n if n is not None else "")))
        tags = "".join("<span class='tag'>%s</span>" % esc(t) for t in r["_tags"])
        if r["_lic"]:
            tags += "<span class='tag lic'>%s%s</span>" % (
                "⭐ " if r["_star"] else "", esc(" · ".join(r["_lic"][:3])))
        newtag = "<span class='newtag'>NEW</span>" if r["_id"] in new_ids else ""
        parts.append(
            "<div class='card%s' id='%s' data-dl='%s'><h2>%d. %s — %s%s</h2>"
            "<span class='src'>%s</span>%s %s<table>"
            % (" star" if r["_star"] else "", r["_hid"], esc(r.get("접수마감", "")),
               r["_no"], esc(r.get("기관명")), esc(r.get("제목")), newtag,
               esc(r.get("출처")), tags,
               "<span class='d'>%s</span>" % esc(label) if label else ""))
        for k in ("고용형태", "근무지", "학력", "인원", "접수시작", "접수마감"):
            if r.get(k):
                parts.append("<tr><td class='k'>%s</td><td>%s</td></tr>" % (esc(k), esc(r[k])))
        if r.get("대체인력") == "예":
            parts.append("<tr><td class='k'>비고</td><td class='d'>대체인력(육아휴직 등) 채용</td></tr>")
        for k in ("자격", "우대"):
            if r.get(k):
                parts.append("<tr><td class='k'>%s</td><td>%s</td></tr>"
                             % ("자격요건" if k == "자격" else "우대", esc(r[k])[:600]))
        if r.get("링크"):
            parts.append("<tr><td class='k'>지원</td><td><a href='%s' target='_blank' rel='noopener'>%s</a></td></tr>"
                         % (esc(r["링크"]), esc(r["링크"])))
        parts.append("</table></div>")


def _table(parts, rows, today, new_ids):
    parts.append("<table class='sum'><tr><th>#</th><th>기관</th><th>제목</th>"
                 "<th>분야</th><th>자격</th><th>고용형태</th><th>마감</th><th>상태</th></tr>")
    for r in rows:
        n = dday(r.get("접수마감"), today)
        label = ("마감" if n is not None and n < 0 else
                 ("오늘 마감" if n == 0 else ("D-%d" % n if n is not None else "")))
        urgent = n is not None and 0 <= n <= 7
        cls = " class='urgent'" if urgent else (" class='star'" if r["_star"] else "")
        parts.append(
            "<tr%s data-dl='%s'><td>%d</td><td><a href='#%s'>%s</a></td>"
            "<td>%s%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
            "<td class='st'>%s</td></tr>"
            % (cls, esc(r.get("접수마감", "")), r["_no"], r["_hid"], esc(r.get("기관명")),
               esc(r.get("제목"))[:60],
               "<span class='newtag'>NEW</span>" if r["_id"] in new_ids else "",
               esc(" · ".join(r["_tags"])),
               "⭐" if r["_star"] else ("○" if r["_lic"] else ""),
               esc(r.get("고용형태")), esc(r.get("접수마감")), esc(label)))
    parts.append("</table>")


def write_page(path, rows, today, updated, new_ids, mode, downloads=None, stats=None):
    """mode='open' 첫 화면 / mode='new' 이번 회차 신규."""
    rows = sorted(rows, key=lambda r: sort_key(r, today))
    star = [r for r in rows if r["_star"]]
    for i, r in enumerate(rows, 1):
        r["_no"] = i
        r["_hid"] = "c%d" % i

    if mode == "open":
        head = "상담 · 행정 채용공고 (%s)" % REGION_LABEL
        sub = ("%s 기준으로 <b>접수가 끝나지 않은 공고 %d건</b>입니다. "
               "마감이 가까운 순서로 놓았습니다. 그중 <b>⭐ 청소년상담사 자격이 "
               "요건·우대에 적힌 공고가 %d건</b>입니다.<br>"
               "<b>월요일·목요일 아침</b>에 저절로 새로 고쳐집니다 — 이 주소만 기억해 두세요."
               % (today.isoformat(), len(rows), len(star)))
    else:
        head = "새로 올라온 공고 (%s)" % today.isoformat()
        sub = ("이번에 <b>새로 확인된 공고 %d건</b>입니다. (⭐ 청소년상담사 관련 %d건)"
               % (len(rows), len(star)))

    parts = ["<!doctype html><meta charset='utf-8'>"
             "<meta name='viewport' content='width=device-width, initial-scale=1'>"
             "<title>%s</title><style>%s</style>" % (esc(head), CSS)]
    parts.append("<h1>%s</h1>" % esc(head))
    parts.append("<div class='sub'>%s</div>" % sub)
    parts.append("<div class='stamp'>🕒 마지막 갱신 %s (한국시간)%s</div>"
                 % (esc(updated), (" · " + esc(stats)) if stats else ""))

    parts.append(
        "<div class='warn'><b>⚠️ 지원하기 전에 반드시 원문 공고를 확인하세요.</b><br>"
        "이 표는 잡알리오·워크넷의 공고 요약을 자동으로 모아 정리한 것입니다. "
        "<b>자격요건·마감일·근무조건이 실제 공고와 다를 수 있고, 공고가 중간에 바뀌거나 "
        "취소되기도 합니다.</b> 각 공고의 <b>‘지원’ 링크</b>를 눌러 직접 확인하세요.<br>"
        "<b>‘자격’ 칸의 ⭐ 는 공고 글자에 ‘청소년상담사’가 들어 있다는 표시일 뿐</b>이며, "
        "필수 자격인지 우대 사항인지는 원문에서 확인해야 합니다.</div>")

    bar = ["<div class='bar'>",
           "<button onclick='window.print()'>🖨 인쇄 / PDF로 저장</button>"]
    for label, href in (downloads or []):
        bar.append("<a class='btnlink' href='%s' download>⬇ %s</a>" % (esc(href), esc(label)))
    arc = (SITE_URL + "archive.html") if mode == "open" else "archive.html"
    bar.append("<a class='btnlink' href='%s'>📁 지난 회차 보기</a>" % arc)
    if mode != "open":
        bar.append("<a class='btnlink' href='index.html'>🏠 지금 지원 가능한 공고</a>")
    bar.append("</div>")
    parts.append("".join(bar))

    if not rows:
        parts.append("<div class='none'>해당하는 공고가 없습니다.</div>")
        save_text(path, "\n".join(parts))
        return

    if star:
        parts.append("<h3>⭐ 청소년상담사 자격이 적힌 공고</h3>")
        _table(parts, star, today, new_ids)
        parts.append("<p class='hint'>공고 글자에 ‘청소년상담사’가 들어 있는 것만 따로 모았습니다. "
                     "아래 ‘한눈에 보기’에 다시 나옵니다.</p>")

    parts.append("<h3>한눈에 보기 — 전체 %d건</h3>" % len(rows))
    parts.append("<p class='hint'>줄을 누르면 아래 상세 내용으로 바로 갑니다.</p>")
    _table(parts, rows, today, new_ids)

    parts.append("<h3>상세</h3>")
    _cards(parts, rows, today, new_ids)

    parts.append(RECALC_JS)
    save_text(path, "\n".join(parts))


def write_archive():
    names = sorted((f for f in os.listdir(DOCS_DIR)
                    if re.match(r"^\d{4}-\d{2}-\d{2}\.html$", f)), reverse=True)
    items = "\n".join("<li><a href='%s'>%s</a></li>" % (esc(n), esc(n[:-5])) for n in names)
    save_text(os.path.join(DOCS_DIR, "archive.html"),
              "<!doctype html><meta charset='utf-8'>"
              "<meta name='viewport' content='width=device-width, initial-scale=1'>"
              "<title>지난 회차 모아보기</title>"
              "<style>body{font-family:'Malgun Gothic',sans-serif;margin:28px;color:#111}"
              "h1{font-size:19px}ul{line-height:2;padding-left:18px}a{color:#1d4ed8}"
              ".back{display:inline-block;margin-bottom:14px}</style>"
              "<a class='back' href='index.html'>← 지금 지원 가능한 공고</a>"
              "<h1>지난 회차 모아보기</h1><ul>%s</ul>" % (items or "<li>아직 없습니다.</li>"))


def prune_old(today):
    limit = (today - dt.timedelta(days=KEEP_DAYS)).isoformat()
    for f in sorted(os.listdir(DOCS_DIR)):
        m = re.match(r"^(\d{4}-\d{2}-\d{2})\.html$", f)
        if m and m.group(1) < limit:
            os.remove(os.path.join(DOCS_DIR, f))
            log("  오래된 리포트 삭제: %s" % f)


# ─────────────────────────── 본체 ───────────────────────────
def main():
    today = dt.date.today()
    updated = run_stamp()
    log("=" * 60)
    log("수집 시작 — %s (%s)" % (today.isoformat(), REGION_LABEL))

    raw = []
    raw += alio_fetch()
    raw += work_fetch()
    log("받아온 공고 %d건" % len(raw))

    kept, dropped = [], {}
    for r in raw:
        normalize(r)
        ok, why = keep(r)
        if ok:
            kept.append(r)
        else:
            dropped[why] = dropped.get(why, 0) + 1
    log("걸러서 남은 공고 %d건 (버린 이유: %s)"
        % (len(kept), ", ".join("%s %d" % kv for kv in sorted(dropped.items())) or "없음"))

    # 이미 본 공고인지 — 새 것에 NEW 를 붙이려고 기억해 둔다
    seen = load_json(STATE_PATH, {})
    new_ids = {r["_id"] for r in kept if r["_id"] not in seen}
    for r in kept:
        seen[r["_id"]] = today.isoformat()
    # 1년 넘게 안 보인 번호는 지운다(파일이 계속 커지는 걸 막는다)
    old = (today - dt.timedelta(days=365)).isoformat()
    seen = {k: v for k, v in seen.items() if v >= old}
    save_json(STATE_PATH, seen)
    log("이번에 새로 보인 공고 %d건" % len(new_ids))

    # 접수가 끝나지 않은 것만 첫 화면에
    live = [r for r in kept
            if (dday(r.get("접수마감"), today) is None) or dday(r.get("접수마감"), today) >= 0]

    stamp = today.isoformat()
    day_dir = os.path.join(FILES_DIR, stamp)
    os.makedirs(day_dir, exist_ok=True)
    write_xlsx(sorted(live, key=lambda r: sort_key(r, today)),
               os.path.join(day_dir, "상담행정공고_%s.xlsx" % stamp), today, "공고")
    downloads = [("엑셀로 받기", "files/%s/상담행정공고_%s.xlsx" % (stamp, stamp))]

    stats = "잡알리오·워크넷에서 %d건을 훑어 %d건" % (len(raw), len(kept))
    write_page(os.path.join(DOCS_DIR, "index.html"), live, today, updated,
               new_ids, "open", downloads, stats)

    # 신규가 하나도 없는 날은 회차 리포트를 만들지 않는다(빈 껍데기가 목록을 채운다)
    news = [r for r in live if r["_id"] in new_ids]
    if news:
        write_page(os.path.join(DOCS_DIR, "%s.html" % stamp), news, today, updated,
                   new_ids, "new",
                   [("이번 회차 엑셀", "files/%s/상담행정공고_%s.xlsx" % (stamp, stamp))])
        log("회차 리포트 작성: %s.html (%d건)" % (stamp, len(news)))
    else:
        log("신규 0건 — 회차 리포트는 만들지 않습니다")

    prune_old(today)
    write_archive()

    # 누적 이력(사람이 볼 일은 없지만, 나중에 통계를 내려면 필요하다)
    hist = load_json(HIST_PATH, {})
    for r in kept:
        hist[r["_id"]] = {k: v for k, v in r.items() if not k.startswith("_")}
        hist[r["_id"]]["분야"] = " · ".join(r["_tags"])
        hist[r["_id"]]["청소년상담사"] = "⭐" if r["_star"] else ("○" if r["_lic"] else "")
        hist[r["_id"]]["처음본날"] = seen.get(r["_id"], stamp)
    save_json(HIST_PATH, hist)

    save_text(os.path.join(DATA_DIR, "last_run.txt"),
              "%s / 받아온 %d건 / 남긴 %d건 / 신규 %d건\n"
              % (updated, len(raw), len(kept), len(new_ids)))
    log("끝 — 지금 지원 가능 %d건, 신규 %d건" % (len(live), len(new_ids)))


if __name__ == "__main__":
    main()
