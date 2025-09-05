# 단어시험 프로그램

엑셀에 정리된 단어장을 불러와서 **무작위 단어시험/답안 파일**을 자동으로 생성하는 파이썬 기반 프로그램입니다.  
영어학원 단어시험, 자기 학습용 단어테스트 제작에 활용할 수 있습니다.

## 주요 기능
- `src` 폴더의 **엑셀 파일(.xlsx/.xls)** 자동 인식
- 파일마다 다른 **열 구조(3열/4열/기타)**를 미리보기 후 직접 매핑 가능
- 매핑 정보(`schema_memory.txt`)를 기억하여 다음 실행 시 자동 적용
- **DAY 범위**를 지정하면 해당 단원의 단어만 시험지로 변환
- 무작위로 **단어/뜻 중 하나를 빈칸 처리** → 시험지/답안 분리 저장
- 결과 파일은 `result` 폴더에 **엑셀(.xlsx)**로 생성

## 폴더 구조
단어시험 프로그램/
├─ day_word_quiz.py # 메인 파이썬 코드
├─ day_word_quiz.exe # (PyInstaller로 빌드 시 생성)
├─ schema_memory.txt # 파일별 스키마 매핑 정보 자동 저장
├─ 원본/ # 단어장 원본 엑셀 파일들 (.xlsx/.xls)
└─ 파일/ # 생성된 시험지/답안 저장 위치

markdown
코드 복사

## 사용 방법 (Python 실행)
1. Python 3.9+ 가상환경 준비
   ```bash
   pip install -U pandas openpyxl

2. 원본/ 폴더에 단어장 엑셀 파일을 넣기

3. 실행
   python day_word_quiz.py

4. 순서
- 사용할 파일 번호 선택
- (처음인 경우) 스키마 매핑 설정
- DAY 범위 입력 → 예: 2 3 또는 4 7
- 파일/ 폴더에 [번호]_..._퀴즈_시험.xlsx, [번호]_..._퀴즈_답안.xlsx 생성

## EXE 빌드 방법
Python 없이 실행 가능한 단일 실행파일을 만들려면:

1. 패키지 설치
   pip install -U pyinstaller pandas openpyxl

2. 빌드
   pyinstaller day_word_quiz.py --onefile --clean --name day_word_quiz
3. 결과물은 dist/day_word_quiz.exe 에 생성
   exe 파일을 원본/ 파일/ 폴더와 함께 두고 실행하면 됩니다.

> 참고: 빌드된 exe는 실행 경로를 exe가 있는 폴더로 고정합니다.
> 따라서 반드시 원본/과 파일/ 폴더를 exe 옆에 두세요.

## 예시 화면
```
python-repl
코드 복사
[단어장 파일 목록]
1. Word Master 수능2000.xlsx
2. Word Master_고등 Basic.xlsx
...

파일 번호 (1~2) 또는 'Q' 종료: 2
[스키마 설정 필요] Word Master_고등 Basic

[시트 미리보기]
1:1   | 2:2       | 3:3           | 4:0
번호    | English   | 의미            | Day
1     | advise    | 충고하다          | Day 1
...

[스키마 선택]
A: 자동추정
3: 3열
4: 4열
M: 수동 매핑
선택 (A/3/4/M): M
...
생성완료 → 시험: 1번_Word Master_고등 Basic_Day1-2_퀴즈_시험.xlsx / 답안: 1번_Word Master_고등 Basic_Day1-2_퀴즈_답안.x 
```