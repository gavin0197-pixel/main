# 작업 규칙

- 대상: Windows PowerPoint VBA 매크로를 만들어 주는 Python(tkinter) GUI. 배경과 사용자 선호는 `docs/chatgpt-handover.md`.
- **버전 관리 (필수):** 기능이나 GUI를 수정해서 커밋할 때마다
  1. `ppt_macro_gen/__init__.py`의 `__version__`을 올린다 (기능 추가/변경 → 부 번호, 버그 수정 → 수 번호).
  2. `CHANGELOG.md` 맨 위에 `## [버전] - 날짜` 섹션을 추가한다 (`tests/test_version.py`가 검사).
  3. 버전 표시는 `__version__`에서 자동으로 GUI 제목/상단, EXE 파일명·속성, VBA 머리 주석에 들어가므로 하드코딩하지 않는다.
- 새 버전을 푸시하면 `.github/workflows/build-exe.yml`이 Windows EXE를 빌드하고 GitHub Release `v<버전>`을 만든다.
- 테스트: `python -m pytest`. 생성 VBA는 실제 PowerPoint에서 검증되지 않았으므로 "테스트 완료"라고 말하지 않는다.
- 사용자 요청 없이 사진 압축·크롭·비율 변경 등 동작을 바꾸지 않는다. 가로/세로, 간격/피치를 혼동하지 않는다.
