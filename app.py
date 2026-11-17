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
SHEET = "한중 실무 사전"
CATEGORIES = ["무역·물류", "비즈니스 이메일", "취업·면접", "경제·금융", "IT·산업", "일반 비즈니스"]
PREFERRED = ["선하증권", "계약", "이력서", "신용장"]
EXAMPLE_FIELDS = [
    ("예문1 중국어", "예문1 병음", "예문1 한국어"),
    ("예문2 중국어", "예문2 병음", "예문2 한국어"),
]

st.set_page_config(
    page_title="Lingua · 한중 실무 사전",
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
def load_dictionary(path: str, _modified: float) -> list[dict]:
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


def default_settings() -> dict:
    return {"show_pinyin": True, "show_example_pinyin": True, "tts_rate": 0.95}


def read_user_data() -> dict:
    if not USER_PATH.exists():
        return {"history": [], "favorites": [], "wordbook": [], "seen": [], "settings": default_settings()}
    try:
        raw = json.loads(USER_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"history": [], "favorites": [], "wordbook": [], "seen": [], "settings": default_settings()}
    settings = default_settings()
    settings.update(raw.get("settings") or {})
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
            "show_pinyin": bool(st.session_state.show_pinyin),
            "show_example_pinyin": bool(st.session_state.show_example_pinyin),
            "tts_rate": float(st.session_state.tts_rate),
        },
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if st.session_state.get("_snapshot") == text:
        return
    USER_PATH.write_text(text, encoding="utf-8")
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
    st.session_state.show_pinyin = settings["show_pinyin"]
    st.session_state.show_example_pinyin = settings["show_example_pinyin"]
    st.session_state.tts_rate = float(settings["tts_rate"])
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
            f'<p class="panel-title">단어장 · {html.escape(choice)} · {done}/{len(words)} 학습</p>',
            unsafe_allow_html=True,
        )
        if not words:
            st.markdown('<p class="empty-note">이 파트에 표제어가 없습니다.</p>', unsafe_allow_html=True)
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
        st.markdown('<p class="panel-title">단어장</p>', unsafe_allow_html=True)
        st.markdown('<p class="hint">파트를 누르면 그 표제어가 열립니다.</p>', unsafe_allow_html=True)
        seen = seen_ids()
        for name in CATEGORIES:
            rows = by_cat.get(name, [])
            done = sum(1 for row in rows if row["번호"] in seen)
            st.button(
                f"{name}  ·  {done}/{len(rows)}",
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


def toggle_browse() -> None:
    st.session_state.browse_open = not st.session_state.browse_open


def inject_styles() -> None:
    css = STYLE_PATH.read_text(encoding="utf-8") if STYLE_PATH.exists() else ""
    if BG_PHOTO.exists():
        encoded = base64.b64encode(BG_PHOTO.read_bytes()).decode("ascii")
        css += (
            ".stApp{background-image:url('data:image/jpeg;base64," + encoded + "') !important;"
            "background-size:cover !important;background-attachment:fixed !important;"
            "background-position:center center !important;}"
        )
    elif BG_PATH.exists():
        svg = BG_PATH.read_text(encoding="utf-8")
        uri = "data:image/svg+xml;charset=UTF-8," + quote(svg)
        css += (
            ".stApp{background-image:url('" + uri + "') !important;"
            "background-size:cover !important;background-attachment:fixed !important;"
            "background-position:center !important;}"
        )
    else:
        css += ".stApp{background-image:linear-gradient(180deg,#f4f7f8 0%,#c5d5df 48%,#6d8794 100%) !important;}"
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def render_head(current: dict, rate: float, show_pinyin: bool) -> None:
    ko = html.escape(current["한국어"])
    zh = html.escape(current["중국어"])
    zh_attr = html.escape(current["중국어"], quote=True)
    pinyin = html.escape(current["병음"]) if show_pinyin and current["병음"] else ""
    pinyin_html = f'<div class="py">{pinyin}</div>' if pinyin else ""
    components.html(
        f"""
        <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@700&family=Noto+Serif+SC:wght@600&display=swap">
        <style>
          html, body {{ margin: 0; padding: 10px 0 0; background: transparent; overflow: visible;
            font-family: Pretendard, "Noto Sans KR", "PingFang SC", "Microsoft YaHei", "Malgun Gothic", sans-serif; }}
          .line {{ display: flex; align-items: center; gap: 10px; }}
          .ko {{ font-size: 32px; font-weight: 700; color: #1e293b; letter-spacing: -0.03em; line-height: 1.15; }}
          .zh {{ font-family: "Noto Serif SC", "Noto Sans SC", "PingFang SC", "Microsoft YaHei", serif;
            font-size: 30px; font-weight: 600; color: #1e293b; line-height: 1.2; }}
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


def search_ids(records: list[dict], query: str, category: str) -> list[int]:
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
    return [word_id for _, word_id in ranked[:6]]


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
                            name,
                            key=f"scope_{name}",
                            type="primary" if active_quiz_scope() == name else "secondary",
                            on_click=set_quiz_scope,
                            args=(name,),
                        )
                st.button(
                    "중국어 → 한국어" if direction == "zh" else "한국어 → 중국어",
                    key="quiz_flip",
                    on_click=flip_quiz_direction,
                )
            with st.container(key="quizcounts", horizontal=True, gap="small", wrap=False):
                for count in QUIZ_COUNTS:
                    st.button(
                        f"{count}문제",
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
                            name,
                            key=f"field_{name}",
                            type="primary" if field == name else "secondary",
                            on_click=set_quiz_field,
                            args=(name,),
                        )
        notice = quiz.get("notice") or ""
        if notice:
            st.markdown(f'<p class="quiz-progress">퀴즈 · {html.escape(scope)}</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="meaning quiz-center">{html.escape(notice)}</p>', unsafe_allow_html=True)
            return
        if finished:
            st.markdown(
                f'<p class="quiz-progress">퀴즈 · {html.escape(scope)} · {quiz["score"]} / {total}</p>',
                unsafe_allow_html=True,
            )
            st.markdown(f'<p class="quiz-word">{quiz["score"]} / {total}</p>', unsafe_allow_html=True)
            st.markdown('<p class="meaning quiz-center">이번 퀴즈를 마쳤습니다.</p>', unsafe_allow_html=True)
            wrong_ids = [int(item) for item in quiz.get("wrong") or [] if int(item) in by_id]
            if wrong_ids:
                names = ", ".join(by_id[item]["한국어"] for item in wrong_ids)
                st.markdown(f'<p class="hint quiz-center">틀린 단어: {html.escape(names)}</p>', unsafe_allow_html=True)
                st.button("틀린 단어만 다시", key="quiz_retry", on_click=retry_wrong)
            st.button("다시 풀기", key="quiz_again", on_click=reset_quiz)
            st.button("사전으로", key="quiz_back", on_click=toggle_quiz)
            return
        question = quiz["questions"][quiz["index"]]
        answer = by_id[question["answer"]]
        prompt = "한국어 뜻을 고르세요" if direction == "zh" else "알맞은 중국어를 고르세요"
        st.markdown(
            f'<p class="quiz-progress">퀴즈 {quiz["index"] + 1} / {total} · {html.escape(scope)} · {quiz["score"]}개 맞춤</p>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div class="quiz-stage"><p class="quiz-prompt">{html.escape(prompt)}</p>'
            f'<span class="badge">{html.escape(answer["범주"])}</span></div>',
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
            pinyin = f"  {row['병음']}" if st.session_state.show_pinyin and row["병음"] else ""
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
                f'<p class="meaning quiz-center">정답은 <b>{html.escape(correct)}</b>입니다.</p>',
                unsafe_allow_html=True,
            )
            last = quiz["index"] + 1 >= total
            st.button("결과 보기" if last else "다음 문제", key="quiz_next", on_click=next_quiz)


def neighbor(pool: list[dict], current: dict, step: int = 1) -> dict:
    ids = [row["번호"] for row in pool]
    if current["번호"] not in ids:
        return pool[0]
    index = ids.index(current["번호"])
    return pool[(index + step) % len(ids)]


BRAND_MARK = """
<div class="brand">
  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M5 4.5h10.5A2.5 2.5 0 0 1 18 7v12.2a1.8 1.8 0 0 0-1.8-1.7H5V4.5Z" stroke="#1E293B" stroke-width="1.6"/>
    <path d="M5 4.5A2.2 2.2 0 0 0 2.8 6.7V18a2.2 2.2 0 0 0 2.2 2.2H16" stroke="#1E293B" stroke-width="1.6"/>
    <path d="M8 8.2h6.2M8 11.4h6.2" stroke="#1E293B" stroke-width="1.4" stroke-linecap="round"/>
  </svg>
  <div class="logo">Lingua</div>
</div>
"""


def render_brand() -> None:
    st.markdown(BRAND_MARK, unsafe_allow_html=True)


def render_brand_mark() -> None:
    st.markdown(BRAND_MARK.replace('<div class="logo">Lingua</div>', ""), unsafe_allow_html=True)


def toggle_quiz() -> None:
    st.session_state.panel = None
    if st.session_state.get("mode") == "quiz":
        st.session_state.mode = "dict"
        return
    st.session_state.mode = "quiz"
    st.session_state.quiz = None


def daily_word(records: list[dict]) -> dict:
    by_ko = {row["한국어"]: row for row in records}
    pool = [by_ko[name] for name in PREFERRED if name in by_ko] or records
    return pool[date.today().toordinal() % len(pool)]


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


def open_from_home(word_id: int) -> None:
    open_word(word_id)
    st.session_state.home = False
    st.session_state.mode = "dict"
    st.session_state.panel = None
    st.session_state.query = ""


def home_panel(name: str) -> None:
    st.session_state.home = False
    st.session_state.mode = "dict"
    st.session_state.quiz = None
    st.session_state.panel = name


def clear_history() -> None:
    st.session_state.history = []


def render_home(records: list[dict], by_id: dict[int, dict]) -> None:
    with st.container(key="home"):
        with st.container(key="hometop", horizontal=True, vertical_alignment="center", wrap=False):
            render_brand()
            with st.container(key="homeicons", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
                st.button("◷", key="home_history", on_click=home_panel, args=("history",))
                st.markdown('<span class="home-rule"></span>', unsafe_allow_html=True)
                st.button("☆", key="home_favorites", on_click=home_panel, args=("favorites",))
                st.markdown('<span class="home-rule"></span>', unsafe_allow_html=True)
                st.button("⚙", key="home_settings", on_click=home_panel, args=("settings",))
        st.markdown(
            """
            <div class="home-hero">
              <h1 class="home-title">Lingua</h1>
              <p class="home-tag">— &nbsp; Look up. Learn more. &nbsp; —</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("home_search", border=False):
            with st.container(key="homesearch", horizontal=True, vertical_alignment="center", wrap=False):
                st.text_input(
                    "검색",
                    key="home_query",
                    label_visibility="collapsed",
                    placeholder="Search for a word...",
                )
                submitted = st.form_submit_button("→")
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
        word = daily_word(records)
        with st.container(key="wotd"):
            st.markdown(
                (
                    '<div class="wotd-card">'
                    '<div class="wotd-kicker">Word of the Day</div>'
                    '<div class="wotd-line">'
                    f'<span class="wotd-word">{html.escape(word["한국어"])}</span>'
                    f'<span class="wotd-py">{html.escape(str(word.get("병음") or ""))}</span>'
                    '<span class="wotd-go">›</span>'
                    "</div></div>"
                ),
                unsafe_allow_html=True,
            )
            st.button("오늘의 단어", key="wotd_open", on_click=open_from_home, args=(word["번호"],))
        recent = [by_id[item] for item in st.session_state.history if item in by_id][:3]
        if recent:
            with st.container(key="recent", horizontal=True, vertical_alignment="center", wrap=True, gap="small"):
                st.markdown(
                    '<p class="recent-label">Recent Searches</p>',
                    unsafe_allow_html=True,
                )
                for row in recent:
                    st.button(
                        row["한국어"],
                        key=f"recent_{row['번호']}",
                        on_click=open_from_home,
                        args=(row["번호"],),
                    )
                st.button("Clear", key="recent_clear", on_click=clear_history)


def render_topbar() -> None:
    quiz_on = st.session_state.get("mode") == "quiz"
    with st.container(key="topbar", horizontal=True, vertical_alignment="center", wrap=False):
        with st.container(key="brandhome", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
            render_brand_mark()
            st.button("Lingua", key="brand_home", on_click=go_home)
        with st.container(key="navrow", horizontal=True, vertical_alignment="center", gap="small", wrap=False):
            st.button(
                "검색기록",
                key="nav_history",
                type="primary" if st.session_state.panel == "history" else "secondary",
                on_click=toggle_panel,
                args=("history",),
            )
            st.button(
                "즐겨찾기",
                key="nav_favorites",
                type="primary" if st.session_state.panel == "favorites" else "secondary",
                on_click=toggle_panel,
                args=("favorites",),
            )
            st.button(
                "단어장",
                key="nav_book",
                type="primary" if st.session_state.get("mode") == "book" else "secondary",
                on_click=toggle_book,
            )
            st.button(
                "퀴즈",
                key="nav_quiz",
                type="primary" if quiz_on else "secondary",
                on_click=toggle_quiz,
            )
            st.button(
                "설정",
                key="nav_settings",
                type="primary" if st.session_state.panel == "settings" else "secondary",
                on_click=toggle_panel,
                args=("settings",),
            )


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
        st.markdown('<p class="empty-note">아직 비어 있습니다.</p>', unsafe_allow_html=True)
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
        st.button("닫기", key="close_panel", on_click=toggle_panel, args=(panel,))
        if panel == "history":
            st.markdown('<p class="panel-title">검색기록</p>', unsafe_allow_html=True)
            st.markdown('<p class="hint">최근에 연 표제어입니다.</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.history, by_id, "hist", on_pick=open_from_panel)
        elif panel == "favorites":
            st.markdown('<p class="panel-title">즐겨찾기</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.favorites, by_id, "favrow", on_pick=open_from_panel)
            st.markdown('<p class="panel-title">내 단어장</p>', unsafe_allow_html=True)
            render_word_buttons(st.session_state.wordbook, by_id, "bookrow", on_pick=open_from_panel)
        else:
            st.markdown('<p class="panel-title">설정</p>', unsafe_allow_html=True)
            st.toggle("병음(Pinyin) 표시", key="show_pinyin")
            st.toggle("예문 병음 표시", key="show_example_pinyin")
            st.slider("발음 속도", min_value=0.6, max_value=1.4, step=0.05, key="tts_rate")
            st.markdown(
                '<p class="hint">재생 버튼은 브라우저 음성 합성(한국어 ko-KR, 중국어 zh-CN)을 사용합니다. 속도는 이 슬라이더 값이 바로 반영됩니다.</p>',
                unsafe_allow_html=True,
            )


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
            st.markdown('<p class="kicker">실무 예문</p>', unsafe_allow_html=True)
            shown = True
        parts = []
        if zh:
            parts.append(f'<p class="zh-ex">{html.escape(zh)}</p>')
        if st.session_state.show_example_pinyin and py:
            parts.append(f'<p class="py-ex">{html.escape(py)}</p>')
        if ko:
            parts.append(f'<p class="ko-ex">{html.escape(ko)}</p>')
        st.markdown("".join(parts), unsafe_allow_html=True)
    if not shown:
        st.markdown('<p class="empty-note">등록된 예문이 없습니다.</p>', unsafe_allow_html=True)


def render_card(current: dict, pool: list[dict], neighbors: tuple[dict | None, dict | None] | None = None) -> None:
    rate = float(st.session_state.tts_rate)
    with st.container(key="wordcard"):
        render_head(current, rate, bool(st.session_state.show_pinyin))
        st.markdown(f'<span class="badge">{html.escape(current["범주"])}</span>', unsafe_allow_html=True)
        if neighbors is not None:
            part_ids = [int(item) for item in st.session_state.get("part_ids") or []]
            seen = seen_ids()
            done = sum(1 for word_id in part_ids if word_id in seen)
            st.markdown(
                f'<p class="hint">이 파트 {done}/{len(part_ids)} 학습</p>',
                unsafe_allow_html=True,
            )
        meaning = current["뜻"] or "뜻풀이가 없습니다."
        st.markdown(f'<p class="meaning">{html.escape(meaning)}</p>', unsafe_allow_html=True)
        loved = current["번호"] in st.session_state.favorites
        saved = current["번호"] in st.session_state.wordbook
        fav_col, book_col = st.columns(2)
        with fav_col:
            st.button(
                "♥ 즐겨찾기" if loved else "♡ 즐겨찾기",
                key="fav_on" if loved else "fav_off",
                on_click=toggle_list,
                args=("favorites", current["번호"]),
            )
        with book_col:
            st.button(
                "✓ 단어장에 저장됨" if saved else "📖 내 단어장에 저장",
                key="book_on" if saved else "book_off",
                on_click=toggle_list,
                args=("wordbook", current["번호"]),
            )
        st.markdown('<hr class="hairline"/>', unsafe_allow_html=True)
        st.markdown('<p class="kicker">관련 실무 어휘</p>', unsafe_allow_html=True)
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
                        f"이전 · {prev_row['한국어']}",
                        key="card_prev",
                        on_click=step_word,
                        args=(-1,),
                    )
                else:
                    st.button("이전", key="card_prev_off", disabled=True)
                if next_row:
                    st.button(
                        f"다음 · {next_row['한국어']}",
                        key="card_next",
                        on_click=step_word,
                        args=(1,),
                    )
                else:
                    st.button("다음", key="card_next_off", disabled=True)
            st.button("목록", key="back_list", on_click=show_part_list)
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
                f"다음 단어  ·  {nxt['한국어']}",
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
    inject_styles()
    workbook = find_workbook()
    if not workbook.exists():
        st.error(f"사전 파일을 찾지 못했습니다: {workbook}")
        return
    records = load_dictionary(str(workbook), workbook.stat().st_mtime)
    if not records:
        st.error("시트에 표제어가 없습니다.")
        return
    init_state(records)
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
            st.text_input(
                "검색",
                key="query",
                label_visibility="collapsed",
                placeholder="한국어 · 중국어 · 병음",
            )
    render_panel(by_id)
    if st.session_state.panel:
        persist()
        return

    choice = st.session_state.get("category_applied") or "전체"
    if choice not in CATEGORIES:
        choice = "전체"

    query = st.session_state.query.strip()
    hits = search_ids(records, query, "전체")
    token = f"{query}|전체"
    if query and hits and st.session_state.get("exact_token") != token:
        best = by_id[hits[0]]
        folded = fold_pinyin(query)
        exact = best["한국어"] == query or best["중국어"] == query or best["py_fold"] == folded
        if exact:
            st.session_state.exact_token = token
            open_word(best["번호"])

    if query:
        with st.container(key="suggest", horizontal=True, wrap=True, gap="small"):
            if not hits:
                st.markdown('<p class="empty-note">해당하는 표제어가 없습니다.</p>', unsafe_allow_html=True)
            else:
                for word_id in hits:
                    row = by_id[word_id]
                    st.button(
                        f"{row['한국어']} · {row['중국어']}",
                        key=f"sug_{word_id}",
                        on_click=open_word,
                        args=(word_id,),
                    )

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
