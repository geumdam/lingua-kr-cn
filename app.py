"""Lingua — 수묵 산수화 감성의 한중 실무 사전."""

from __future__ import annotations

import base64
import html
import json
import random
import unicodedata
from datetime import date
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

ROOT = Path(__file__).resolve().parent
DATA_CANDIDATES = [
    ROOT / "src" / "data" / "dictionary_verified_v3.xlsx",
    ROOT / "dictionary_verified_v3.xlsx",
]
USER_PATH = ROOT / "user_data.json"
STYLE_PATH = ROOT / "styles.css"
BG_PHOTO = ROOT / "assets" / "bg_ink_wash.jpg"
BG_PATH = ROOT / "assets" / "bg_ink_wash.svg"
HOME_BG = ROOT / "assets" / "home_bg.jpg"
LOGO_PATH = ROOT / "assets" / "logo.png"
SITE_NAME = "译语类解"
SHEET = "한중 실무 사전"
CATEGORIES = ["무역·물류", "비즈니스 이메일", "취업·면접", "경제·금융", "IT·산업", "일반 비즈니스"]
CATEGORY_ZH = {
    "무역·물류": "贸易·物流",
    "비즈니스 이메일": "商务邮件",
    "취업·면접": "求职·面试",
    "경제·금융": "经济·金融",
    "IT·산업": "IT·产业",
    "일반 비즈니스": "一般商务",
}
TEXT = {
    "검색기록": ("검색기록", "检索记录"),
    "즐겨찾기": ("즐겨찾기", "收藏"),
    "단어장": ("단어장", "单词本"),
    "퀴즈": ("퀴즈", "测验"),
    "설정": ("설정", "设置"),
    "중문전환": ("中文", "한국어"),
    "검색자리": ("한국어 · 중국어 · 병음", "韩语 · 中文 · 拼音"),
    "홈검색": ("단어를 검색하세요", "搜索单词"),
    "검색실행": ("검색", "搜索"),
    "태그라인": ("찾아보고 익히세요", "查一词，懂更多"),
    "오늘의단어": ("오늘의 단어", "每日一词"),
    "최근검색": ("최근 검색", "最近搜索"),
    "지우기": ("지우기", "清除"),
    "닫기": ("닫기", "关闭"),
    "최근힌트": ("최근에 연 표제어입니다.", "最近打开的词条。"),
    "내단어장": ("내 단어장", "我的单词本"),
    "내 단어장": ("내 단어장", "我的单词本"),
    "틀린 단어": ("틀린 단어", "错题"),
    "비어있음": ("아직 비어 있습니다.", "还是空的。"),
    "병음표시": ("병음(Pinyin) 표시", "显示拼音"),
    "예문병음": ("예문 병음 표시", "显示例句拼音"),
    "발음속도": ("발음 속도", "语速"),
    "설정힌트": (
        "재생 버튼은 브라우저 음성 합성(한국어 ko-KR, 중국어 zh-CN)을 사용합니다. 속도는 이 슬라이더 값이 바로 반영됩니다.",
        "朗读使用浏览器语音合成（韩语 ko-KR、中文 zh-CN）。语速会立即反映这个滑块的值。",
    ),
    "단어장힌트": ("파트를 누르면 그 표제어가 열립니다.", "点开分类即可查看词条。"),
    "파트없음": ("이 파트에 표제어가 없습니다.", "这个分类还没有词条。"),
    "학습": ("{done}/{total} 학습", "{done}/{total} 已学"),
    "파트학습": ("이 파트 {done}/{total} 학습", "本部分 {done}/{total} 已学"),
    "뜻없음": ("뜻풀이가 없습니다.", "暂无释义。"),
    "예문없음": ("등록된 예문이 없습니다.", "暂无例句。"),
    "예문": ("실무 예문", "实务例句"),
    "관련어휘": ("관련 실무 어휘", "相关实务词汇"),
    "저장됨": ("✓ 단어장에 저장됨", "✓ 已存入单词本"),
    "저장": ("📖 내 단어장에 저장", "📖 存入单词本"),
    "이전": ("이전", "上一个"),
    "다음": ("다음", "下一个"),
    "목록": ("목록", "列表"),
    "다음단어": ("다음 단어", "下一个词"),
    "검색없음": ("해당하는 표제어가 없습니다.", "没有对应的词条。"),
    "전체": ("전체", "全部"),
    "분야별": ("분야별", "按领域"),
    "틀린단어": ("틀린 단어", "错题"),
    "문제": ("문제", "题"),
    "방향중한": ("중국어 → 한국어", "中文 → 韩语"),
    "방향한중": ("한국어 → 중국어", "韩语 → 中文"),
    "뜻고르기": ("한국어 뜻을 고르세요", "请选择韩语释义"),
    "중국어고르기": ("알맞은 중국어를 고르세요", "请选择对应中文"),
    "퀴즈진행": ("퀴즈 {n} / {total} · {scope} · {score}개 맞춤", "测验 {n} / {total} · {scope} · 答对 {score}题"),
    "퀴즈끝": ("퀴즈 · {scope} · {score} / {total}", "测验 · {scope} · {score} / {total}"),
    "퀴즈제목": ("퀴즈 · {scope}", "测验 · {scope}"),
    "퀴즈완료": ("이번 퀴즈를 마쳤습니다.", "这次测验做完了。"),
    "틀린목록": ("틀린 단어: {names}", "错题：{names}"),
    "틀린다시": ("틀린 단어만 다시", "只重做错题"),
    "다시풀기": ("다시 풀기", "再做一次"),
    "사전으로": ("사전으로", "返回词典"),
    "결과보기": ("결과 보기", "查看结果"),
    "다음문제": ("다음 문제", "下一题"),
    "정답": ("정답은 {answer}입니다.", "正确答案是 {answer}。"),
    "단어부족": (
        "내 단어장에 4개 이상 저장하면 그 단어로 퀴즈를 낼 수 있습니다.",
        "单词本至少保存 4 个词后，才能用这些词出题。",
    ),
    "파트제목": ("단어장 · {name} · {done}/{total} 학습", "单词本 · {name} · {done}/{total} 已学"),
}
PREFERRED = ["선하증권", "계약", "이력서", "신용장"]
EXAMPLE_FIELDS = [
    ("예문1 중국어", "예문1 병음", "예문1 한국어"),
    ("예문2 중국어", "예문2 병음", "예문2 한국어"),
]
TTS_RATE_MIN, TTS_RATE_MAX, TTS_RATE_STEP = 0.6, 2.0, 0.05

st.set_page_config(
    page_title="译语类解 · 한중 실무 사전",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def fold_pinyin(text: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(text).lower())
    plain = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    plain = plain.replace("ü", "v").replace("u:", "v")
    return "".join(plain.split())


def find_workbook() -> Path:
    for path in DATA_CANDIDATES:
        if path.exists():
            return path
    return DATA_CANDIDATES[0]


@st.cache_data(show_spinner="사전을 펼치는 중…")
def load_dictionary(path: str, modified: float) -> list[dict]:
    frame = pd.read_excel(path, sheet_name=SHEET).fillna("")
    records = frame.to_dict("records")
    cleaned = []
    for row in records:
        item = {key: ("" if value is None else str(value).strip()) for key, value in row.items()}
        if not item.get("번호"):
            continue
        item["번호"] = int(float(item["번호"]))
        item["py_fold"] = fold_pinyin(item.get("병음", ""))
        for zh_key, py_key, ko_key in EXAMPLE_FIELDS:
            item.setdefault(zh_key, "")
            item.setdefault(py_key, "")
            item.setdefault(ko_key, "")
        cleaned.append(item)
    return cleaned


def lang() -> str:
    return "zh" if st.session_state.get("lang") == "zh" else "ko"


def t(key: str, **kwargs: object) -> str:
    pair = TEXT.get(key)
    text = pair[1] if pair and lang() == "zh" else pair[0] if pair else key
    return text.format(**kwargs) if kwargs else text


def cat_label(name: str) -> str:
    if lang() == "zh":
        return CATEGORY_ZH.get(name, name)
    return name


def shown(name: str) -> str:
    if name in CATEGORY_ZH:
        return cat_label(name)
    return t(name)


def toggle_lang() -> None:
    st.session_state.lang = "ko" if lang() == "zh" else "zh"


def logo_src() -> str:
    if not LOGO_PATH.exists():
        return ""
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def brand_html(with_name: bool = True) -> str:
    image = logo_src()
    mark = f'<img class="brand-logo" alt="" src="{image}">' if image else ""
    name = f'<div class="logo">{SITE_NAME}</div>' if with_name else ""
    return f'<div class="brand">{mark}{name}</div>'


def default_settings() -> dict:
    return {
        "show_pinyin": True,
        "show_example_pinyin": True,
        "tts_rate": 0.95,
        "lang": "ko",
    }


def prefs() -> dict:
    """Canonical settings, kept out of widget keys so they survive unmounting."""
    current = st.session_state.get("prefs")
    if not isinstance(current, dict):
        current = default_settings()
        st.session_state.prefs = current
    return current


def pref(name: str):
    return prefs().get(name, default_settings()[name])


def save_pref(name: str) -> None:
    value = st.session_state.get(f"ui_{name}")
    if value is not None:
        prefs()[name] = value


def read_user_data() -> dict:
    if not USER_PATH.exists():
        return {"history": [], "favorites": [], "wordbook": [], "seen": [], "settings": default_settings()}
    try:
        raw = json.loads(USER_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"history": [], "favorites": [], "wordbook": [], "seen": [], "settings": default_settings()}
    stored = raw.get("settings") or {}
    settings = default_settings()
    for name in ("show_pinyin", "show_example_pinyin"):
        if name in stored:
            settings[name] = bool(stored[name])
    try:
        settings["tts_rate"] = min(
            TTS_RATE_MAX, max(TTS_RATE_MIN, float(stored.get("tts_rate", settings["tts_rate"])))
        )
    except (TypeError, ValueError):
        pass
    settings["lang"] = "zh" if stored.get("lang") == "zh" else "ko"
    return {
        "history": [int(item) for item in raw.get("history", [])],
        "favorites": [int(item) for item in raw.get("favorites", [])],
        "wordbook": [int(item) for item in raw.get("wordbook", [])],
        "seen": [int(item) for item in raw.get("seen", [])],
        "settings": settings,
    }


def persist() -> None:
    payload = {
        "history": [int(item) for item in st.session_state.history][:30],
        "favorites": [int(item) for item in st.session_state.favorites][:200],
        "wordbook": [int(item) for item in st.session_state.wordbook][:200],
        "seen": [int(item) for item in st.session_state.get("seen") or []][:3000],
        "settings": {
            "show_pinyin": bool(pref("show_pinyin")),
            "show_example_pinyin": bool(pref("show_example_pinyin")),
            "tts_rate": float(pref("tts_rate")),
            "lang": lang(),
        },
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if st.session_state.get("_snapshot") == text:
        return
    try:
        USER_PATH.write_text(text, encoding="utf-8")
    except OSError:
        # Read-only filesystem (e.g. a hosted deployment): keep working in
        # memory instead of failing the whole run.
        pass
    st.session_state._snapshot = text


def pick_default(records: list[dict]) -> int:
    by_ko = {row["한국어"]: row["번호"] for row in records}
    for name in PREFERRED:
        if name in by_ko:
            return by_ko[name]
    return records[0]["번호"]


def init_state(records: list[dict]) -> None:
    if st.session_state.get("ready"):
        return
    saved = read_user_data()
    settings = saved["settings"]
    st.session_state.history = saved["history"]
    st.session_state.favorites = saved["favorites"]
    st.session_state.wordbook = saved["wordbook"]
    st.session_state.seen = saved["seen"]
    st.session_state.prefs = dict(settings)
    st.session_state.lang = "zh" if settings.get("lang") == "zh" else "ko"
    st.session_state.selected_id = pick_default(records)
    st.session_state.panel = None
    st.session_state.browse_open = False
    st.session_state.category_pills = "전체"
    st.session_state.category_applied = "전체"
    st.session_state.query = ""
    st.session_state.mode = "dict"
    st.session_state.quiz = None
    st.session_state.quiz_level = "쉬움"
    st.session_state.quiz_scope = "전체"
    st.session_state.ready = True
    persist()


def open_part(name: str) -> None:
    st.session_state.mode = "dict"
    st.session_state.panel = None
    st.session_state.quiz = None
    st.session_state.category_pills = name
    st.session_state.category_applied = name
    st.session_state.reading = False


def toggle_book() -> None:
    st.session_state.panel = None
    st.session_state.quiz = None
    if st.session_state.get("mode") == "book":
        st.session_state.mode = "dict"
        return
    st.session_state.mode = "book"


def render_part_words(choice: str, words: list[dict]) -> None:
    with st.container(key="catlist"):
        done = sum(1 for row in words if row["번호"] in seen_ids())
        st.markdown(
            f'<p class="panel-title">{html.escape(t("파트제목", name=cat_label(choice), done=done, total=len(words)))}</p>',
            unsafe_allow_html=True,
        )
        if not words:
            st.markdown(f'<p class="empty-note">{html.escape(t("파트없음"))}</p>', unsafe_allow_html=True)
            return
        with st.container(key="partscroll", height=420):
            for row in words:
                st.button(
                    f"{row['한국어']}  ·  {row['중국어']}",
                    key=f"partword_{row['번호']}",
                    on_click=open_word,
                    args=(row["번호"],),
                )


def render_wordbook(by_cat: dict[str, list[dict]]) -> None:
    with st.container(key="wordbookhome"):
        st.markdown(f'<p class="panel-title">{html.escape(t("단어장"))}</p>', unsafe_allow_html=True)
        st.markdown(f'<p class="hint">{html.escape(t("단어장힌트"))}</p>', unsafe_allow_html=True)
        seen = seen_ids()
        for name in CATEGORIES:
            rows = by_cat.get(name, [])
            done = sum(1 for row in rows if row["번호"] in seen)
            st.button(
                f"{cat_label(name)}  ·  {done}/{len(rows)}",
                key=f"open_part_{name}",
                on_click=open_part,
                args=(name,),
            )


def seen_ids() -> set[int]:
    return {int(item) for item in st.session_state.get("seen") or []}


def mark_seen(word_id: int) -> None:
    word_id = int(word_id)
    current = [int(item) for item in st.session_state.get("seen") or []]
    st.session_state.seen = [word_id, *[item for item in current if item != word_id]][:3000]


def remember(word_id: int) -> None:
    word_id = int(word_id)
    st.session_state.selected_id = word_id
    history = [item for item in st.session_state.history if item != word_id]
    st.session_state.history = [word_id, *history][:30]
    st.session_state.browse_open = False
    mark_seen(word_id)


def open_word(word_id: int) -> None:
    remember(word_id)
    st.session_state.reading = True
    category = st.session_state.get("word_cats", {}).get(int(word_id))
    if category in CATEGORIES:
        st.session_state.category_applied = category


def show_part_list() -> None:
    st.session_state.reading = False


def step_word(step: int) -> None:
    ids = st.session_state.get("part_ids") or []
    current = int(st.session_state.selected_id)
    if current not in ids:
        return
    index = ids.index(current) + step
    if 0 <= index < len(ids):
        remember(ids[index])
        st.session_state.reading = True


def toggle_list(bucket: str, word_id: int) -> None:
    items = [item for item in st.session_state[bucket] if item != word_id]
    if word_id not in st.session_state[bucket]:
        items = [word_id, *items]
    st.session_state[bucket] = items[:200]


def toggle_panel(name: str) -> None:
    st.session_state.panel = None if st.session_state.panel == name else name


def open_from_panel(word_id: int) -> None:
    open_word(word_id)
    st.session_state.panel = None
    st.session_state.mode = "dict"


def open_from_search(word_id: int) -> None:
    """Open a lookup result from any screen (dictionary, wordbook or quiz)."""
    st.session_state.home = False
    st.session_state.panel = None
    st.session_state.quiz = None
    st.session_state.quiz_retry_ids = None
    st.session_state.mode = "dict"
    st.session_state.browse_open = False
    st.session_state.reading = True
    open_word(int(word_id))


def toggle_browse() -> None:
    st.session_state.browse_open = not st.session_state.browse_open


def _scene_layer(selector: str, encoded: str, fallback: str) -> str:
    uri = "data:image/jpeg;base64," + encoded
    return (
        f"{selector}{{background-image:none !important;background-color:{fallback} !important;}}"
        f"{selector}::before{{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;"
        f"background:url('{uri}') center/cover no-repeat;}}"
    )


# --- Search field styling -------------------------------------------------
# The page sits on a blue-grey ink wash. The field carries no border, no
# stroke and no shadow at all: a single blurred ::before layer behind the
# input lays down a drift of pale mist that thins to nothing at both ends.
SEARCH_FIELD_CSS = """
/* 안개 레이어 — 경계선 없는 배경 */
.st-key-headersearch,
.st-key-homeform {
  position: relative !important;
  isolation: isolate;
  overflow: visible !important;
  background: none !important;
  border: none !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  backdrop-filter: none !important;
  -webkit-backdrop-filter: none !important;
}

.st-key-headersearch::before,
.st-key-homeform::before {
  content: "";
  position: absolute;
  left: 50%;
  top: 50%;
  width: 120%;
  height: 140%;
  transform: translate(-50%, -50%);
  z-index: 0;
  pointer-events: none;
  background:
    radial-gradient(40% 62% at 18% 52%, rgba(248,250,252,0.95) 0%, rgba(242,246,250,0.74) 40%, rgba(235,241,246,0.50) 68%, rgba(235,241,246,0) 100%),
    radial-gradient(36% 58% at 42% 44%, rgba(248,250,252,0.95) 0%, rgba(242,246,250,0.72) 42%, rgba(235,241,246,0.48) 70%, rgba(235,241,246,0) 100%),
    radial-gradient(38% 64% at 66% 56%, rgba(248,250,252,0.94) 0%, rgba(242,246,250,0.70) 44%, rgba(235,241,246,0.46) 72%, rgba(235,241,246,0) 100%),
    radial-gradient(32% 54% at 88% 48%, rgba(246,249,252,0.88) 0%, rgba(239,244,248,0.58) 46%, rgba(235,241,246,0.32) 74%, rgba(235,241,246,0) 100%);
  filter: blur(18px);
  -webkit-filter: blur(18px);
}

/* 입력창·아이콘은 안개 위로 */
.st-key-headersearch > *,
.st-key-homeform > * {
  position: relative !important;
  z-index: 2;
}

.st-key-headersearch [data-testid="stTextInputRootElement"],
.st-key-homesearch [data-testid="stTextInputRootElement"],
.st-key-headersearch div[data-baseweb="input"],
.st-key-headersearch div[data-baseweb="base-input"],
.st-key-homesearch div[data-baseweb="input"],
.st-key-homesearch div[data-baseweb="base-input"],
.st-key-headersearch [data-testid="stTextInput"] > div,
.st-key-homesearch [data-testid="stTextInput"] > div {
  background: transparent !important;
  border: none !important;
  border-radius: 0 !important;
  box-shadow: none !important;
}

.st-key-headersearch input,
.st-key-homesearch input {
  background: transparent !important;
  border: none !important;
  border-radius: 0 !important;
  outline: none !important;
  box-shadow: none !important;
  color: #3a4550 !important;
  caret-color: #3a4550;
  text-shadow: none;
}

.st-key-headersearch input:focus,
.st-key-homesearch input:focus,
.st-key-headersearch input:focus-visible,
.st-key-homesearch input:focus-visible {
  background: transparent !important;
  border: none !important;
  outline: none !important;
  box-shadow: none !important;
}

.st-key-headersearch input::placeholder,
.st-key-homesearch input::placeholder {
  color: rgba(58, 69, 80, 0.55) !important;
}

/* 돋보기 / 검색 버튼도 선명하게, 테두리 없이 */
.st-key-headersearch [data-testid="stTextInput"]::before,
.st-key-homesearch [data-testid="stTextInput"]::before {
  z-index: 2;
  opacity: 0.6;
}

.st-key-headersearch [data-testid="stFormSubmitButton"] button,
.st-key-homesearch button {
  position: relative !important;
  z-index: 2;
  border: none !important;
  border-left: none !important;
  border-radius: 999px !important;
  background-color: transparent !important;
  box-shadow: none !important;
  opacity: 0.7;
}

.st-key-headersearch [data-testid="stFormSubmitButton"] button:hover,
.st-key-homesearch button:hover {
  background-color: rgba(255, 255, 255, 0.35) !important;
  opacity: 1;
}
"""


ZH_FONT_CSS = (
    # In the Chinese UI every Han glyph should fall through to the Kai face
    # (the headword/logo font) instead of Malgun Gothic, which is what the
    # Korean-first stack used to pick for simplified Chinese. Pretendard has
    # no Han glyphs, so Korean text is still rendered in Pretendard; the
    # Korean-capable fonts stay at the end of the list as an offline fallback.
    ':root{--sans:Pretendard,"LXGW WenKai","KaiTi","STKaiti","Kaiti SC","楷体",'
    '"Noto Sans KR","Malgun Gothic",sans-serif;}'
    '.stApp,.stApp button,.stApp p,.stApp label,.stApp input,.stApp textarea,'
    '.stApp [data-testid="stMarkdownContainer"],.stApp [data-testid="stWidgetLabel"]'
    '{font-family:var(--sans),var(--kaiti);}'
)


def inject_styles() -> None:
    css = STYLE_PATH.read_text(encoding="utf-8") if STYLE_PATH.exists() else ""
    if lang() == "zh":
        css += ZH_FONT_CSS
    css += SEARCH_FIELD_CSS
    css += (
        ".stApp{position:relative;background-image:none !important;}"
        "[data-testid='stAppViewContainer'],[data-testid='stMain']{"
        "background:transparent !important;position:relative;z-index:1;}"
        ".block-container{position:relative;z-index:1;min-height:100vh;"
        "background:transparent !important;background-image:none !important;}"
    )
    if BG_PHOTO.exists():
        encoded = base64.b64encode(BG_PHOTO.read_bytes()).decode("ascii")
        css += _scene_layer(".stApp:not(:has(.st-key-home))", encoded, "#c5d4de")
    elif BG_PATH.exists():
        svg = BG_PATH.read_text(encoding="utf-8")
        uri = "data:image/svg+xml;charset=UTF-8," + quote(svg)
        css += (
            ".stApp{background-image:url('" + uri + "') !important;"
            "background-size:cover !important;background-attachment:scroll !important;"
            "background-position:center !important;}"
        )
    else:
        css += ".stApp{background-image:linear-gradient(180deg,#f4f7f8 0%,#c5d5df 48%,#6d8794 100%) !important;}"
    if HOME_BG.exists():
        encoded = base64.b64encode(HOME_BG.read_bytes()).decode("ascii")
        css += _scene_layer(".stApp:has(.st-key-home)", encoded, "#e7eef2")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def render_head(current: dict, rate: float, show_pinyin: bool) -> None:
    ko = html.escape(current["한국어"])
    zh = html.escape(current["중국어"])
    zh_attr = html.escape(current["중국어"], quote=True)
    pinyin = html.escape(current["병음"]) if show_pinyin and current["병음"] else ""
    pinyin_html = f'<div class="py">{pinyin}</div>' if pinyin else ""
    components.html(
        f"""
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/lxgw-wenkai-webfont@1.7.0/lxgwwenkai-regular.css">
        <style>
          html, body {{ margin: 0; padding: 10px 0 0; background: transparent; overflow: visible;
            font-family: Pretendard, "Noto Sans KR", "Malgun Gothic", sans-serif; }}
          .line {{ display: flex; align-items: center; gap: 10px; }}
          .ko {{ font-size: 32px; font-weight: 700; color: #1e293b; letter-spacing: -0.03em; line-height: 1.15; }}
          .zh {{ font-family: "LXGW WenKai", "KaiTi", "STKaiti", "Kaiti SC", "楷体", serif;
            font-weight: 400; font-synthesis: none;
            font-size: 30px; color: #1e293b; line-height: 1.2; }}
          .py {{ margin-top: 2px; font-size: 14px; color: #64748b; font-family: "Noto Sans", sans-serif; }}
          button {{
            width: 36px; height: 36px; border-radius: 999px; cursor: pointer; flex: 0 0 auto;
            border: 1px solid rgba(30, 41, 59, 0.14);
            background: rgba(255, 255, 255, 0.78); color: #1e293b; font-size: 14px;
          }}
          button:hover {{ background: #fff; }}
        </style>
        <div class="line"><div class="ko">{ko}</div></div>
        <div class="line" style="margin-top:8px"><div><div class="zh">{zh}</div>{pinyin_html}</div>
          <button type="button" data-text="{zh_attr}" data-lang="zh-CN" onclick="speak(this)">🔊</button></div>
        <script>
          function speak(el) {{
            const synth = window.speechSynthesis;
            if (!synth) return;
            const utter = new SpeechSynthesisUtterance(el.dataset.text);
            utter.lang = el.dataset.lang;
            utter.rate = {float(rate)};
            synth.cancel();
            synth.speak(utter);
          }}
        </script>
        """,
        height=128,
        width=440,
    )


def search_ranked(records: list[dict], query: str, category: str) -> list[tuple[int, int]]:
    """Return (score, 번호) pairs, best match first."""
    pool = records if category in ("", "전체") else [row for row in records if row["범주"] == category]
    needle = query.strip()
    if not needle:
        return []
    folded = fold_pinyin(needle)
    lowered = needle.lower()
    ranked = []
    for row in pool:
        ko = row["한국어"]
        zh = row["중국어"]
        py = row["py_fold"]
        meaning = row["뜻"]
        if ko.lower() == lowered or zh == needle or (folded and py == folded):
            score = 100
        elif ko.lower().startswith(lowered) or zh.startswith(needle) or (folded and py.startswith(folded)):
            score = 60
        elif lowered in ko.lower() or needle in zh or (folded and folded in py):
            score = 30
        elif lowered in meaning.lower():
            score = 10
        else:
            continue
        ranked.append((score, row["번호"]))
    ranked.sort(key=lambda item: (-item[0], item[1]))
    return ranked[:6]


def search_ids(records: list[dict], query: str, category: str) -> list[int]:
    return [word_id for _, word_id in search_ranked(records, query, category)]


def related_words(pool: list[dict], current: dict, limit: int = 4) -> list[dict]:
    zh_chars = set(current["중국어"])
    ko_chars = set(current["한국어"].replace(" ", ""))
    ranked = []
    for row in pool:
        if row["번호"] == current["번호"]:
            continue
        overlap = len(zh_chars & set(row["중국어"])) * 3 + len(ko_chars & set(row["한국어"].replace(" ", "")))
        ranked.append((overlap, abs(row["번호"] - current["번호"]), row))
    ranked.sort(key=lambda item: (-item[0], item[1]))
    picked = [row for score, _, row in ranked if score > 0][:limit]
    if len(picked) < limit:
        used = {row["번호"] for row in picked}
        for _, _, row in ranked:
            if row["번호"] in used:
                continue
            picked.append(row)
            if len(picked) == limit:
                break
    return picked


QUIZ_COUNTS = (5, 15, 30)
QUIZ_SCOPES = ("전체", "분야별", "내 단어장")
OLD_QUIZ_LEVELS = {"쉬움": 5, "중간": 15, "어려움": 30}


def current_quiz_count() -> int:
    level = st.session_state.get("quiz_level", 5)
    if level in OLD_QUIZ_LEVELS:
        return OLD_QUIZ_LEVELS[level]
    try:
        count = int(level)
    except (TypeError, ValueError):
        return 5
    return count if count in QUIZ_COUNTS else 5


def set_quiz_count(count: int) -> None:
    if current_quiz_count() == int(count):
        return
    st.session_state.quiz_level = int(count)
    st.session_state.quiz = None
    st.session_state.quiz_retry_ids = None


def active_quiz_field() -> str:
    field = st.session_state.get("quiz_field") or ""
    if field in CATEGORIES:
        return field
    applied = st.session_state.get("category_applied") or ""
    if applied in CATEGORIES:
        return applied
    return CATEGORIES[0]


def set_quiz_scope(name: str) -> None:
    if name == "이 파트":
        name = "분야별"
    if st.session_state.get("quiz_scope") == name:
        return
    st.session_state.quiz_scope = name
    if name == "분야별":
        st.session_state.quiz_field = active_quiz_field()
    st.session_state.quiz = None
    st.session_state.quiz_retry_ids = None


def set_quiz_field(name: str) -> None:
    st.session_state.quiz_scope = "분야별"
    if st.session_state.get("quiz_field") == name:
        st.session_state.quiz = None
        st.session_state.quiz_retry_ids = None
        return
    st.session_state.quiz_field = name
    st.session_state.quiz = None
    st.session_state.quiz_retry_ids = None


def active_quiz_scope() -> str:
    scope = st.session_state.get("quiz_scope") or "전체"
    if scope == "이 파트":
        return "분야별"
    return scope if scope in QUIZ_SCOPES else "전체"


def quiz_pool(records: list[dict]) -> tuple[list[dict], str]:
    scope = active_quiz_scope()
    if scope == "내 단어장":
        saved = {int(item) for item in st.session_state.wordbook}
        return [row for row in records if row["번호"] in saved], "내 단어장"
    if scope == "분야별":
        field = active_quiz_field()
        return [row for row in records if row["범주"] == field], field
    return list(records), "전체"


def start_quiz(records: list[dict]) -> None:
    retry = [int(item) for item in (st.session_state.pop("quiz_retry_ids", None) or [])]
    by_id = {row["번호"]: row for row in records}
    asked: list[dict] | None = None
    if retry:
        asked = [by_id[item] for item in retry if item in by_id]
        pool = asked if len(asked) >= 4 else list(records)
        label = "틀린 단어"
        count = len(asked)
    else:
        pool, label = quiz_pool(records)
        count = current_quiz_count()
    if not asked and label == "내 단어장" and len(pool) < 4:
        st.session_state.quiz = {
            "questions": [],
            "index": 0,
            "picked": None,
            "score": 0,
            "direction": "zh",
            "wrong": [],
            "label": "내 단어장",
            "notice": "내 단어장에 4개 이상 저장하면 그 단어로 퀴즈를 낼 수 있습니다.",
        }
        return
    if asked is not None and not asked:
        asked = None
        pool, label = quiz_pool(records)
        count = current_quiz_count()
    rng = random.Random()
    questions = asked if asked is not None else rng.sample(pool, k=min(count, len(pool)))
    built = []
    for row in questions:
        others = [item for item in pool if item["번호"] != row["번호"] and item["한국어"] != row["한국어"]]
        if len(others) < 3:
            others = [item for item in records if item["번호"] != row["번호"]]
        distractors = rng.sample(others, k=min(3, len(others)))
        choices = distractors + [row]
        rng.shuffle(choices)
        built.append({"answer": row["번호"], "choices": [item["번호"] for item in choices]})
    st.session_state.quiz = {
        "questions": built,
        "index": 0,
        "picked": None,
        "score": 0,
        "direction": "zh",
        "wrong": [],
        "label": label,
        "notice": "",
    }


def pick_answer(choice_id: int) -> None:
    quiz = st.session_state.quiz
    if quiz is None or quiz["picked"] is not None:
        return
    quiz["picked"] = int(choice_id)
    current = quiz["questions"][quiz["index"]]
    answer_id = int(current["answer"])
    wrong = [int(item) for item in quiz.get("wrong") or [] if int(item) != answer_id]
    if int(choice_id) == answer_id:
        quiz["score"] += 1
    else:
        wrong.append(answer_id)
    quiz["wrong"] = wrong
    st.session_state.quiz = quiz


def next_quiz() -> None:
    quiz = st.session_state.quiz
    quiz["index"] += 1
    quiz["picked"] = None
    st.session_state.quiz = quiz


def reset_quiz() -> None:
    st.session_state.quiz = None
    st.session_state.quiz_retry_ids = None


def retry_wrong() -> None:
    quiz = st.session_state.quiz or {}
    st.session_state.quiz_retry_ids = [int(item) for item in quiz.get("wrong") or []]
    st.session_state.quiz = None


def flip_quiz_direction() -> None:
    quiz = st.session_state.quiz
    if not quiz:
        return
    quiz["direction"] = "ko" if quiz["direction"] == "zh" else "zh"
    quiz["picked"] = None
    st.session_state.quiz = quiz


def render_quiz(by_id: dict[int, dict]) -> None:
    quiz = st.session_state.quiz
    total = len(quiz["questions"])
    finished = quiz["index"] >= total and not quiz.get("notice")
    if quiz.get("label") == "틀린 단어":
        scope = "틀린 단어"
    elif active_quiz_scope() == "분야별":
        scope = active_quiz_field()
    else:
        scope = active_quiz_scope()
    selected_count = current_quiz_count()
    direction = quiz.get("direction") or "zh"
    with st.container(key="quizcard"):
        with st.container(key="quizbar"):
            with st.container(key="quiztop", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
                with st.container(key="quizscopes", horizontal=True, gap="small", wrap=False):
                    for name in QUIZ_SCOPES:
                        st.button(
                            t(name),
                            key=f"scope_{name}",
                            type="primary" if active_quiz_scope() == name else "secondary",
                            on_click=set_quiz_scope,
                            args=(name,),
                        )
                st.button(
                    t("방향중한") if direction == "zh" else t("방향한중"),
                    key="quiz_flip",
                    on_click=flip_quiz_direction,
                )
            with st.container(key="quizcounts", horizontal=True, gap="small", wrap=False):
                for count in QUIZ_COUNTS:
                    st.button(
                        f"{count}{t('문제')}",
                        key=f"count_{count}",
                        type="primary" if selected_count == count else "secondary",
                        on_click=set_quiz_count,
                        args=(count,),
                    )
            if active_quiz_scope() == "분야별":
                field = active_quiz_field()
                with st.container(key="quizfields", horizontal=True, gap="small", wrap=True):
                    for name in CATEGORIES:
                        st.button(
                            cat_label(name),
                            key=f"field_{name}",
                            type="primary" if field == name else "secondary",
                            on_click=set_quiz_field,
                            args=(name,),
                        )
        notice = quiz.get("notice") or ""
        if notice:
            st.markdown(
                f'<p class="quiz-progress">{html.escape(t("퀴즈제목", scope=shown(scope)))}</p>',
                unsafe_allow_html=True,
            )
            st.markdown(f'<p class="meaning quiz-center">{html.escape(t("단어부족") if notice else "")}</p>', unsafe_allow_html=True)
            return
        if finished:
            st.markdown(
                f'<p class="quiz-progress">{html.escape(t("퀴즈끝", scope=shown(scope), score=quiz["score"], total=total))}</p>',
                unsafe_allow_html=True,
            )
            st.markdown(f'<p class="quiz-word">{quiz["score"]} / {total}</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="meaning quiz-center">{html.escape(t("퀴즈완료"))}</p>', unsafe_allow_html=True)
            wrong_ids = [int(item) for item in quiz.get("wrong") or [] if int(item) in by_id]
            if wrong_ids:
                names = ", ".join(by_id[item]["한국어"] for item in wrong_ids)
                st.markdown(
                    f'<p class="hint quiz-center">{html.escape(t("틀린목록", names=names))}</p>',
                    unsafe_allow_html=True,
                )
                st.button(t("틀린다시"), key="quiz_retry", on_click=retry_wrong)
            st.button(t("다시풀기"), key="quiz_again", on_click=reset_quiz)
            st.button(t("사전으로"), key="quiz_back", on_click=toggle_quiz)
            return
        question = quiz["questions"][quiz["index"]]
        answer = by_id[question["answer"]]
        prompt = t("뜻고르기") if direction == "zh" else t("중국어고르기")
        st.markdown(
            f'<p class="quiz-progress">{html.escape(t("퀴즈진행", n=quiz["index"] + 1, total=total, scope=shown(scope), score=quiz["score"]))}</p>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div class="quiz-stage"><p class="quiz-prompt">{html.escape(prompt)}</p>'
            f'<span class="badge">{html.escape(cat_label(answer["범주"]))}</span></div>',
            unsafe_allow_html=True,
        )
        if direction == "zh":
            st.markdown(f'<p class="quiz-word">{html.escape(answer["중국어"])}</p>', unsafe_allow_html=True)
            if answer["병음"]:
                st.markdown(f'<p class="quiz-py">{html.escape(answer["병음"])}</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p class="quiz-word">{html.escape(answer["한국어"])}</p>', unsafe_allow_html=True)
        picked = quiz["picked"]
        letters = "ABCD"

        def choice_text(row: dict) -> str:
            if direction == "zh":
                return row["한국어"]
            pinyin = f"  {row['병음']}" if pref("show_pinyin") and row["병음"] else ""
            return f"{row['중국어']}{pinyin}"

        if picked is None:
            with st.container(key="quizchoices"):
                for index, choice_id in enumerate(question["choices"]):
                    letter = letters[index] if index < len(letters) else str(index + 1)
                    st.button(
                        f"{letter}   {choice_text(by_id[choice_id])}",
                        key=f"choice_{quiz['index']}_{choice_id}",
                        on_click=pick_answer,
                        args=(choice_id,),
                    )
        else:
            blocks = []
            for index, choice_id in enumerate(question["choices"]):
                row = by_id[choice_id]
                classes = ["q-choice"]
                if choice_id == picked:
                    classes.append("mine")
                if choice_id == question["answer"]:
                    classes.append("right")
                letter = letters[index] if index < len(letters) else str(index + 1)
                blocks.append(
                    f'<div class="{" ".join(classes)}"><span class="q-letter">{letter}</span>'
                    f"<span>{html.escape(choice_text(row))}</span></div>"
                )
            st.markdown("".join(blocks), unsafe_allow_html=True)
            correct = choice_text(answer)
            st.markdown(
                f'<p class="meaning quiz-center">{t("정답", answer=f"<b>{html.escape(correct)}</b>")}</p>',
                unsafe_allow_html=True,
            )
            last = quiz["index"] + 1 >= total
            st.button(t("결과보기") if last else t("다음문제"), key="quiz_next", on_click=next_quiz)


def neighbor(pool: list[dict], current: dict, step: int = 1) -> dict:
    ids = [row["번호"] for row in pool]
    if current["번호"] not in ids:
        return pool[0]
    index = ids.index(current["번호"])
    return pool[(index + step) % len(ids)]


def render_brand() -> None:
    st.markdown(brand_html(True), unsafe_allow_html=True)


def render_brand_mark() -> None:
    st.markdown(brand_html(False), unsafe_allow_html=True)


def toggle_quiz() -> None:
    st.session_state.panel = None
    if st.session_state.get("mode") == "quiz":
        st.session_state.mode = "dict"
        return
    st.session_state.mode = "quiz"
    st.session_state.quiz = None


def home_panel(name: str) -> None:
    st.session_state.home = False
    st.session_state.mode = "dict"
    st.session_state.panel = name
    st.session_state.query = ""
    st.session_state.quiz = None
    st.session_state.reading = False


def home_open(word_id: int) -> None:
    st.session_state.home = False
    st.session_state.mode = "dict"
    st.session_state.panel = None
    st.session_state.query = ""
    st.session_state.quiz = None
    st.session_state.reading = False
    st.session_state.category_applied = "전체"
    open_word(int(word_id))


def clear_history() -> None:
    st.session_state.history = []


def word_of_day(records: list[dict]) -> dict:
    return records[date.today().toordinal() % len(records)]


def go_home() -> None:
    st.session_state.home = True
    st.session_state.panel = None
    st.session_state.mode = "dict"
    st.session_state.quiz = None
    st.session_state.query = ""
    st.session_state.category_applied = "전체"
    st.session_state.category_pills = "전체"
    st.session_state.reading = False
    st.session_state.home_query = ""


def render_home(records: list[dict], by_id: dict[int, dict]) -> None:
    today = word_of_day(records)
    pinyin = (today.get("병음") or "").strip()
    spoken = f'<em>/ {html.escape(pinyin)} /</em>' if pinyin else ""
    recent = [by_id[item] for item in st.session_state.history if item in by_id][:3]
    with st.container(key="home"):
        with st.container(key="hometop", horizontal=True, vertical_alignment="center", wrap=False):
            st.markdown(brand_html(True).replace('class="brand"', 'class="home-brand"', 1), unsafe_allow_html=True)
            with st.container(key="homeactions", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
                st.button(t("중문전환"), key="home_lang", on_click=toggle_lang)
                with st.container(key="homeicons", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
                    st.button(t("검색기록"), key="home_history", on_click=home_panel, args=("history",))
                    st.button(t("즐겨찾기"), key="home_star", on_click=home_panel, args=("favorites",))
                    st.button(t("설정"), key="home_gear", on_click=home_panel, args=("settings",))
        with st.container(key="homehero"):
            st.markdown(
                f"""
<h1 class="home-title">译语类解</h1>
<p class="home-tag"><i></i>{html.escape(t("태그라인"))}<i></i></p>
                """,
                unsafe_allow_html=True,
            )
            with st.container(key="homeform"):
                with st.form("home_search", border=False):
                    with st.container(key="homesearch", horizontal=True, vertical_alignment="center", wrap=False):
                        st.text_input(
                            "검색",
                            key="home_query",
                            label_visibility="collapsed",
                            placeholder=t("홈검색"),
                        )
                        submitted = st.form_submit_button("검색")
            if submitted:
                query = (st.session_state.get("home_query") or "").strip()
                if query:
                    st.session_state.query = query
                    st.session_state.home = False
                    st.session_state.mode = "dict"
                    st.session_state.panel = None
                    st.session_state.category_applied = "전체"
                    st.session_state.reading = False
                    st.rerun()
            with st.container(key="wotd"):
                st.markdown(
                    f"""
<div class="wotd-card">
  <span class="wotd-ico" aria-hidden="true"></span>
  <span class="wotd-copy">
    <span class="wotd-kicker">{html.escape(t("오늘의단어"))}</span>
    <span class="wotd-word">{html.escape(today["한국어"])} {spoken}</span>
  </span>
  <span class="wotd-go" aria-hidden="true"></span>
</div>
                    """,
                    unsafe_allow_html=True,
                )
                st.button("오늘의 단어", key="wotd_open", on_click=home_open, args=(today["번호"],))
            if recent:
                with st.container(key="homerecent", horizontal=True, vertical_alignment="center", gap="small", wrap=True):
                    st.markdown(f'<span class="recent-label">{html.escape(t("최근검색"))}</span>', unsafe_allow_html=True)
                    for row in recent:
                        st.button(
                            row["한국어"],
                            key=f"home_recent_{row['번호']}",
                            on_click=home_open,
                            args=(row["번호"],),
                        )
                    st.button(t("지우기"), key="home_clear", on_click=clear_history)


def render_topbar() -> None:
    quiz_on = st.session_state.get("mode") == "quiz"
    with st.container(key="topbar", horizontal=True, vertical_alignment="center", wrap=False):
        with st.container(key="brandhome", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
            render_brand_mark()
            st.button(SITE_NAME, key="brand_home", on_click=go_home)
        with st.container(key="navrow", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
            st.button(
                t("검색기록"),
                key="nav_history",
                type="primary" if st.session_state.panel == "history" else "secondary",
                on_click=toggle_panel,
                args=("history",),
            )
            st.button(
                t("즐겨찾기"),
                key="nav_favorites",
                type="primary" if st.session_state.panel == "favorites" else "secondary",
                on_click=toggle_panel,
                args=("favorites",),
            )
            st.button(
                t("단어장"),
                key="nav_book",
                type="primary" if st.session_state.get("mode") == "book" else "secondary",
                on_click=toggle_book,
            )
            st.button(
                t("퀴즈"),
                key="nav_quiz",
                type="primary" if quiz_on else "secondary",
                on_click=toggle_quiz,
            )
            st.button(
                t("설정"),
                key="nav_settings",
                type="primary" if st.session_state.panel == "settings" else "secondary",
                on_click=toggle_panel,
                args=("settings",),
            )
            st.button(t("중문전환"), key="nav_lang", on_click=toggle_lang)


def render_word_buttons(
    ids: list[int],
    by_id: dict[int, dict],
    key_prefix: str,
    columns: int = 1,
    on_pick=None,
) -> None:
    pick = on_pick or remember
    visible = [by_id[word_id] for word_id in ids if word_id in by_id]
    if not visible:
        st.markdown(f'<p class="empty-note">{html.escape(t("비어있음"))}</p>', unsafe_allow_html=True)
        return
    if columns == 1:
        for row in visible[:12]:
            st.button(
                f"{row['한국어']}  ·  {row['중국어']}",
                key=f"{key_prefix}_{row['번호']}",
                on_click=pick,
                args=(row["번호"],),
            )
        return
    for start in range(0, len(visible), columns):
        cols = st.columns(columns)
        for offset, row in enumerate(visible[start : start + columns]):
            with cols[offset]:
                st.button(
                    row["한국어"],
                    key=f"{key_prefix}_{row['번호']}",
                    on_click=pick,
                    args=(row["번호"],),
                )


def render_panel(by_id: dict[int, dict]) -> None:
    panel = st.session_state.panel
    if not panel:
        return
    with st.container(key="drawer"):
        st.button(t("닫기"), key="close_panel", on_click=toggle_panel, args=(panel,))
        if panel == "history":
            st.markdown(f'<p class="panel-title">{html.escape(t("검색기록"))}</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="hint">{html.escape(t("최근힌트"))}</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.history, by_id, "hist", on_pick=open_from_panel)
        elif panel == "favorites":
            st.markdown(f'<p class="panel-title">{html.escape(t("즐겨찾기"))}</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.favorites, by_id, "favrow", on_pick=open_from_panel)
            st.markdown(f'<p class="panel-title">{html.escape(t("내단어장"))}</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.wordbook, by_id, "bookrow", on_pick=open_from_panel)
        else:
            st.markdown(f'<p class="panel-title">{html.escape(t("설정"))}</p>', unsafe_allow_html=True)
            st.toggle(
                t("병음표시"),
                value=bool(pref("show_pinyin")),
                key="ui_show_pinyin",
                on_change=save_pref,
                args=("show_pinyin",),
            )
            st.toggle(
                t("예문병음"),
                value=bool(pref("show_example_pinyin")),
                key="ui_show_example_pinyin",
                on_change=save_pref,
                args=("show_example_pinyin",),
            )
            st.slider(
                t("발음속도"),
                min_value=TTS_RATE_MIN,
                max_value=TTS_RATE_MAX,
                step=TTS_RATE_STEP,
                value=float(pref("tts_rate")),
                key="ui_tts_rate",
                on_change=save_pref,
                args=("tts_rate",),
            )
            st.markdown(f'<p class="hint">{html.escape(t("설정힌트"))}</p>', unsafe_allow_html=True)


def example_count(row: dict) -> int:
    count = 0
    for zh_key, py_key, ko_key in EXAMPLE_FIELDS:
        if any((row.get(zh_key, ""), row.get(py_key, ""), row.get(ko_key, ""))):
            count += 1
    return count


def render_examples(row: dict) -> None:
    shown = False
    for zh_key, py_key, ko_key in EXAMPLE_FIELDS:
        zh, py, ko = row.get(zh_key, ""), row.get(py_key, ""), row.get(ko_key, "")
        if not any((zh, py, ko)):
            continue
        if not shown:
            st.markdown(f'<p class="kicker">{html.escape(t("예문"))}</p>', unsafe_allow_html=True)
            shown = True
        parts = []
        if zh:
            parts.append(f'<p class="zh-ex">{html.escape(zh)}</p>')
        if pref("show_example_pinyin") and py:
            parts.append(f'<p class="py-ex">{html.escape(py)}</p>')
        if ko:
            parts.append(f'<p class="ko-ex">{html.escape(ko)}</p>')
        st.markdown("".join(parts), unsafe_allow_html=True)
    if not shown:
        st.markdown(f'<p class="empty-note">{html.escape(t("예문없음"))}</p>', unsafe_allow_html=True)


def render_card(current: dict, pool: list[dict], neighbors: tuple[dict | None, dict | None] | None = None) -> None:
    rate = float(pref("tts_rate"))
    with st.container(key="wordcard"):
        render_head(current, rate, bool(pref("show_pinyin")))
        st.markdown(f'<span class="badge">{html.escape(cat_label(current["범주"]))}</span>', unsafe_allow_html=True)
        if neighbors is not None:
            part_ids = [int(item) for item in st.session_state.get("part_ids") or []]
            seen = seen_ids()
            done = sum(1 for word_id in part_ids if word_id in seen)
            st.markdown(
                f'<p class="hint">{html.escape(t("파트학습", done=done, total=len(part_ids)))}</p>',
                unsafe_allow_html=True,
            )
        meaning = current["뜻"] or t("뜻없음")
        st.markdown(f'<p class="meaning">{html.escape(meaning)}</p>', unsafe_allow_html=True)
        loved = current["번호"] in st.session_state.favorites
        saved = current["번호"] in st.session_state.wordbook
        fav_col, book_col = st.columns(2)
        with fav_col:
            st.button(
                f"♥ {t('즐겨찾기')}" if loved else f"♡ {t('즐겨찾기')}",
                key="fav_on" if loved else "fav_off",
                on_click=toggle_list,
                args=("favorites", current["번호"]),
            )
        with book_col:
            st.button(
                t("저장됨") if saved else t("저장"),
                key="book_on" if saved else "book_off",
                on_click=toggle_list,
                args=("wordbook", current["번호"]),
            )
        st.markdown('<hr class="hairline"/>', unsafe_allow_html=True)
        st.markdown(f'<p class="kicker">{html.escape(t("관련어휘"))}</p>', unsafe_allow_html=True)
        related = related_words(pool, current)
        if related:
            with st.container(horizontal=True, wrap=True, gap="small", key="relatedrow"):
                for row in related:
                    st.button(
                        row["한국어"],
                        key=f"rel_{row['번호']}",
                        on_click=remember,
                        args=(row["번호"],),
                    )
        render_examples(current)
        if neighbors is not None:
            prev_row, next_row = neighbors
            with st.container(key="cardnav", horizontal=True, wrap=False, gap="small"):
                if prev_row:
                    st.button(
                        f"{t('이전')} · {prev_row['한국어']}",
                        key="card_prev",
                        on_click=step_word,
                        args=(-1,),
                    )
                else:
                    st.button(t("이전"), key="card_prev_off", disabled=True)
                if next_row:
                    st.button(
                        f"{t('다음')} · {next_row['한국어']}",
                        key="card_next",
                        on_click=step_word,
                        args=(1,),
                    )
                else:
                    st.button(t("다음"), key="card_next_off", disabled=True)
            st.button(t("목록"), key="back_list", on_click=show_part_list)
        if example_count(current) > 1:
            st.button("More >", key="more_next", on_click=toggle_browse)
    if example_count(current) > 1 and st.session_state.browse_open:
        with st.container(key="browse"):
            nxt = neighbor(pool, current, 1)
            st.markdown(
                f'<p class="panel-title">{html.escape(current["범주"])}</p>',
                unsafe_allow_html=True,
            )
            st.button(
                f"{t('다음단어')}  ·  {nxt['한국어']}",
                key=f"moreitem_{nxt['번호']}",
                on_click=remember,
                args=(nxt["번호"],),
            )
            upcoming = []
            cursor = current
            for _ in range(8):
                cursor = neighbor(pool, cursor, 1)
                if cursor["번호"] == current["번호"]:
                    break
                upcoming.append(cursor["번호"])
            render_word_buttons(upcoming, {row["번호"]: row for row in pool}, "browseitem", columns=2)




def main() -> None:
    workbook = find_workbook()
    if not workbook.exists():
        st.error(f"사전 파일을 찾지 못했습니다: {workbook}")
        return
    records = load_dictionary(str(workbook), workbook.stat().st_mtime)
    if not records:
        st.error("시트에 표제어가 없습니다.")
        return
    init_state(records)
    inject_styles()
    if "mode" not in st.session_state:
        st.session_state.mode = "dict"
    if "quiz" not in st.session_state:
        st.session_state.quiz = None
    if "quiz_level" not in st.session_state:
        st.session_state.quiz_level = "쉬움"
    if "quiz_scope" not in st.session_state:
        st.session_state.quiz_scope = "전체"
    if "seen" not in st.session_state:
        st.session_state.seen = read_user_data().get("seen", [])
    if st.session_state.get("quiz_scope") == "이 파트":
        st.session_state.quiz_scope = "분야별"
    if "home" not in st.session_state:
        st.session_state.home = True
    by_id = {row["번호"]: row for row in records}
    by_cat: dict[str, list[dict]] = {name: [] for name in CATEGORIES}
    for row in records:
        by_cat.setdefault(row["범주"], []).append(row)
    if st.session_state.selected_id not in by_id:
        st.session_state.selected_id = pick_default(records)
    st.session_state.part_heads = {name: rows[0]["번호"] for name, rows in by_cat.items() if rows}
    st.session_state.word_cats = {row["번호"]: row["범주"] for row in records}
    if st.query_params:
        st.query_params.clear()

    if st.session_state.home:
        render_home(records, by_id)
        persist()
        return

    with st.container(key="stickyhead"):
        render_topbar()
        with st.container(key="searchwrap"):
            # A form makes Enter submit every time (outside a form Streamlit
            # only reruns when the value actually changed) and gives the box a
            # visible submit button.
            with st.form("search_form", border=False):
                with st.container(key="headersearch", horizontal=True, vertical_alignment="center", wrap=False):
                    st.text_input(
                        "검색",
                        key="query",
                        label_visibility="collapsed",
                        placeholder=t("검색자리"),
                    )
                    submitted = st.form_submit_button(t("검색실행"))

    # Search results are rendered before the drawer so that searching keeps
    # working while 검색기록/즐겨찾기/설정 is open.
    query = (st.session_state.get("query") or "").strip()
    ranked = search_ranked(records, query, "전체")
    hits = [word_id for _, word_id in ranked]
    if submitted and ranked and ranked[0][0] >= 60:
        open_from_search(ranked[0][1])

    if query:
        with st.container(key="suggest", horizontal=True, wrap=True, gap="small"):
            if not hits:
                st.markdown(f'<p class="empty-note">{html.escape(t("검색없음"))}</p>', unsafe_allow_html=True)
            else:
                for word_id in hits:
                    row = by_id[word_id]
                    st.button(
                        f"{row['한국어']} · {row['중국어']}",
                        key=f"sug_{word_id}",
                        on_click=open_from_search,
                        args=(word_id,),
                    )

    render_panel(by_id)
    if st.session_state.panel:
        persist()
        return

    choice = st.session_state.get("category_applied") or "전체"
    if choice not in CATEGORIES:
        choice = "전체"

    if st.session_state.mode == "quiz":
        if not st.session_state.quiz:
            start_quiz(records)
        render_quiz(by_id)
    elif st.session_state.mode == "book":
        render_wordbook(by_cat)
    else:
        current = by_id[st.session_state.selected_id]
        if choice not in ("", "전체"):
            words = by_cat.get(choice, [])
            st.session_state.part_ids = [row["번호"] for row in words]
            if st.session_state.get("reading") and current["번호"] in st.session_state.part_ids:
                index = st.session_state.part_ids.index(current["번호"])
                prev_row = by_id[st.session_state.part_ids[index - 1]] if index > 0 else None
                next_row = by_id[st.session_state.part_ids[index + 1]] if index + 1 < len(words) else None
                render_card(current, words, (prev_row, next_row))
            else:
                render_part_words(choice, words)
        else:
            render_card(current, by_cat.get(current["범주"], records))
    persist()


if __name__ == "__main__":
    main()
else:
    main()
