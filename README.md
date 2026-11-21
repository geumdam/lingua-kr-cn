# 译语类解 · 한중 실무 사전

수묵 산수화 감성의 한중 실무 사전입니다. Streamlit으로 만든 단일 페이지 앱이고,
표제어 2,736개를 `src/data/dictionary_verified_v3.xlsx`에서 읽습니다.

- 검색: 한글 · 중국어(간체) · 병음(성조·공백 무시)으로 찾고, 정확/접두 일치면 바로 표제어가 열립니다.
- 단어장 · 즐겨찾기 · 검색기록 · 학습 표시
- 퀴즈: 전체 / 분야별 / 내 단어장, 5·15·30문제, 中→韓 · 韓→中 방향 전환, 틀린 단어만 다시 풀기
- 설정: 병음·예문 병음 표시, 발음 속도(0.6~2.0배), 검색창 안개 UI
- 중국어 UI(中文) 전환 지원 — 표제어와 같은 서체(서예 계열)로 렌더링

## 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 열리면 <http://localhost:8501> 에서 사용할 수 있습니다.
Windows에서는 `실행.bat`을 더블클릭해도 됩니다.

## 구조

```
app.py                     Streamlit 앱 (UI·검색·퀴즈·설정)
styles.css                 앱 스타일 (수묵 배경, 카드, 검색창)
src/data/dictionary_verified_v3.xlsx   사전 데이터 (표제어 2,736개)
assets/                    배경 이미지 · 로고
.streamlit/config.toml     테마 · 브라우저 설정
```

## 배포

Streamlit Community Cloud에 이 저장소를 연결하면 `app.py`가 자동으로 실행됩니다.
사용자 데이터(`user_data.json`)는 저장소에 커밋하지 않으며(`.gitignore`), 호스팅 환경에서는
메모리에만 유지되고 재시작 시 초기화됩니다.
