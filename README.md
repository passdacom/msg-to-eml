# MSG to EML Converter

Microsoft Outlook에서 생성되는 `.msg` 파일을 표준 `.eml` 형식으로 변환하는 프로그램입니다.

## 특징

- ✅ **완전한 오프라인 작동**: 외부 API 없이 순수 Python 라이브러리만 사용
- ✅ **데스크톱 GUI 앱**: 모던 다크 테마 UI
- ✅ **다중 파일 변환**: 여러 파일을 한 번에 변환
- ✅ **첨부파일 보존**: 원본 첨부파일을 그대로 유지
- ✅ **실행 파일 패키징**: PyInstaller로 .app/.exe 생성 가능

---

## 설치

```bash
# 의존성 설치
pip install -r requirements.txt
```

---

## 사용법

### 1. GUI 앱 (권장)

```bash
python gui_app.py
```

현대적인 다크 테마 UI에서 파일을 선택하고 변환할 수 있습니다.

### 2. 명령줄 (CLI)

```bash
# 단일 파일 변환
python msg_to_eml.py input.msg

# 출력 파일명 지정
python msg_to_eml.py input.msg -o output.eml

# 폴더 내 모든 MSG 파일 변환
python msg_to_eml.py ./emails/

# 하위 폴더 포함 재귀 변환
python msg_to_eml.py ./emails/ -r
```

---

## 실행 파일로 패키징 (.app / .exe)

PyInstaller를 사용하여 독립 실행 파일을 만들 수 있습니다.

### macOS (.app)

```bash
# 단일 파일 앱 생성
pyinstaller --onefile --windowed --name "MSG to EML Converter" gui_app.py

# 또는 폴더 형태로 생성 (더 빠른 시작)
pyinstaller --windowed --name "MSG to EML Converter" gui_app.py
```

생성된 앱은 `dist/` 폴더에 있습니다.

### Windows (.exe)

```bash
# 단일 exe 파일 생성
pyinstaller --onefile --windowed --name "MSG_to_EML_Converter" gui_app.py
```

### 아이콘 추가 (선택사항)

```bash
# macOS
pyinstaller --onefile --windowed --icon=icon.icns --name "MSG to EML Converter" gui_app.py

# Windows
pyinstaller --onefile --windowed --icon=icon.ico --name "MSG_to_EML_Converter" gui_app.py
```

---

## 파일 구조

```
msg-to-eml/
├── gui_app.py          # 데스크톱 GUI 앱 (권장)
├── msg_to_eml.py       # 핵심 변환 로직 + CLI
├── requirements.txt    # Python 의존성
├── README.md           # 사용 설명서
│
├── app.py              # 웹 버전 (선택사항)
├── templates/          # 웹 템플릿
└── static/             # 웹 정적 파일
```

---

## 변환 결과

변환된 `.eml` 파일은 다음 이메일 클라이언트에서 열 수 있습니다:

- Mozilla Thunderbird
- Apple Mail
- Gmail (첨부로 업로드)
- 기타 표준 이메일 클라이언트

---

## 기술 정보

### MSG 형식이란?

`.msg` 파일은 Microsoft Outlook에서 사용하는 독자적인 MAPI 형식입니다.

### EML 형식이란?

`.eml` 파일은 RFC 822 표준을 따르는 범용 이메일 형식입니다.

### 사용된 라이브러리

- **extract-msg**: MSG 파일 파싱
- **customtkinter**: 모던 GUI 프레임워크
- **pyinstaller**: 실행 파일 패키징

---

## 라이선스

MIT License
