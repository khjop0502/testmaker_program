import pandas as pd
import random
import re
import os
import sys
import json

if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 exe
    os.chdir(os.path.dirname(sys.executable))
else:
    # 평소 .py 실행
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


BASE_DIR_SRC = "src"
BASE_DIR_OUT = "result"
MEMO_PATH = "schema_memory.txt"  # JSON 저장 (파일명별 스키마 기억)

os.makedirs(BASE_DIR_OUT, exist_ok=True)

def normalize_cell(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = re.sub(r'\s+', ' ', s.strip())
    return s

def list_source_files(src_dir):
    if not os.path.isdir(src_dir):
        print(f'"{src_dir}" 폴더가 없습니다. 같은 위치에 만들어 주세요.')
        sys.exit(1)
    files = [f for f in os.listdir(src_dir)
             if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    files.sort()
    if not files:
        print(f'"{src_dir}" 폴더에 엑셀(.xlsx/.xls) 파일이 없습니다.')
        sys.exit(1)
    return files

def print_file_menu(files):
    print("\n[단어장 파일 목록]")
    for i, name in enumerate(files, start=1):
        print(f"{i}. {name}")
    print()

def parse_days(day_input:str):
    days = sorted([int(d) for d in day_input.strip().split()])
    if len(days) == 2 and days[1] > days[0]:
        days = list(range(days[0], days[1] + 1))
    return days

def load_memory():
    if not os.path.exists(MEMO_PATH):
        return {}
    try:
        with open(MEMO_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_memory(memo):
    with open(MEMO_PATH, "w", encoding="utf-8") as f:
        json.dump(memo, f, ensure_ascii=False, indent=2)

def preview_df(df, max_cols=8, max_rows=6):
    """앞부분 데이터가 있는 열/행을 간단 프리뷰 (열 맞춤 출력)
       반환: 프리뷰에 사용한 원본 컬럼 인덱스 리스트(cols)"""
    nonnull_counts = df.notna().sum().sort_values(ascending=False)
    cols = list(nonnull_counts.index[:max_cols])  # ← 원본 df의 컬럼 인덱스(0-based)들
    prev = df[cols].head(max_rows).fillna("")

    # 각 열의 최대 너비 계산 (제목/데이터 모두 포함)
    col_widths = {}
    for i, c in enumerate(cols, start=1):
        col_title = f"{i}:{c}"
        max_len = len(str(col_title))
        for val in prev[c]:
            max_len = max(max_len, len(str(val)))
        col_widths[c] = max_len + 8  # padding 2칸

    # 헤더 출력
    header = " | ".join(str(i).ljust(col_widths[c]) for i, c in enumerate(cols, start=1))
    print("\n[시트 미리보기] (최대 열:", max_cols, "/ 최대 행:", max_rows, ")")
    print(header)
    print("-" * len(header))

    # 데이터 행 출력
    for _, row in prev.iterrows():
        line = " | ".join(str(row[c]).ljust(col_widths[c]) for c in cols)
        print(line)
    print()

    return cols  # ★ 프리뷰에 사용된 원본 컬럼 인덱스(0-based)를 반환


def auto_guess_use_cols(df):
    """열 개수를 자동 추정(4열 이상이면 4, 아니면 3) - 안전한 기본값"""
    return 4 if df.shape[1] >= 4 else 3

def read_raw_sheet(xlsx_path):
    raw = pd.read_excel(xlsx_path, sheet_name=0, header=None)
    return raw

def ask_schema_interactively(base_name, raw_df):
    """미리보기 보여주고 스키마를 사용자에게 물어서 반환"""
    preview_cols = preview_df(raw_df)  # 프리뷰 열 순서(원본 인덱스 0-based)
    print("[설정 선택]")
    print("  A: 자동추정 (열 수로 3/4열 판단)")
    print("  3: 3열 (단원, 단어, 뜻)")
    print("  4: 4열 (분류, 단원, 단어, 뜻)")
    print("  M: 수동 매핑 (열 번호 직접 지정 - 프리뷰의 왼쪽 번호 1..N 기준)")
    choice = input("선택 (A/3/4/M): ").strip().upper()

    def ask_header_skip():
        ans = input("첫 행을 제목/머리글로 보고 건너뛸까요? (y/n, 기본 y): ").strip().lower()
        return 1 if ans in ("", "y", "yes") else 0

    if choice in ("", "A"):
        use_cols = auto_guess_use_cols(raw_df)
        header_skip = ask_header_skip()
        schema = {"mode": "auto", "use_cols": use_cols, "header_skip": header_skip}
        if use_cols == 3:
            schema["mapping"] = {"분류": -1, "단원": 1, "단어": 2, "뜻": 3}
        else:
            schema["mapping"] = {"분류": 1, "단원": 2, "단어": 3, "뜻": 4}
        return schema

    if choice == "3":
        header_skip = ask_header_skip()
        return {
            "mode": "fixed3",
            "use_cols": 3,
            "header_skip": header_skip,
            "mapping": {"분류": -1, "단원": 1, "단어": 2, "뜻": 3}
        }

    if choice == "4":
        header_skip = ask_header_skip()
        return {
            "mode": "fixed4",
            "use_cols": 4,
            "header_skip": header_skip,
            "mapping": {"분류": 1, "단원": 2, "단어": 3, "뜻": 4}
        }

    # 수동 매핑: 프리뷰 번호(1..len(preview_cols))를 받아 원본 인덱스(1-based)로 변환
    print("\n[수동 매핑] 프리뷰의 왼쪽 번호(1..N)를 입력하세요. 해당 열이 없으면 -1")
    header_skip = 1 if input("첫 행 스킵? (y=1 / n=0, 기본 y): ").strip().lower() in ("", "y") else 0

    def ask_col(role, default_preview_pos):
        v = input(f"'{role}' 열 (프리뷰 번호, 없으면 -1, 기본 {default_preview_pos}): ").strip()
        if v == "":
            sel = default_preview_pos
        else:
            try:
                sel = int(v)
            except:
                sel = default_preview_pos
        if sel == -1:
            return -1
        # 범위를 벗어나면 기본값
        if not (1 <= sel <= len(preview_cols)):
            return -1 if default_preview_pos == -1 else (preview_cols[default_preview_pos - 1] + 1)
        # 프리뷰 → 원본(0-based) → 1-based로 변환
        return preview_cols[sel - 1] + 1

    # 기본값 힌트: 흔한 순서(분류 없음, 단원은 'Day' 열, 단어=2, 뜻=3)를 프리뷰 번호로 가정
    # 안전하게 기본 -1/1/2/3 그대로 두되, 변환 함수가 알아서 처리
    b = ask_col("분류", -1)
    d = ask_col("단원", 1)
    w = ask_col("단어", 2)
    m = ask_col("뜻",   3)

    # use_cols는 매핑이 요구하는 최대 열까지 포함되도록 계산
    mapped_max_1based = max([x for x in (b, d, w, m) if isinstance(x, int) and x != -1] + [0])
    use_cols = max(3, mapped_max_1based) if b == -1 else max(4, mapped_max_1based)

    return {
        "mode": "manual",
        "use_cols": use_cols,
        "header_skip": header_skip,
        "mapping": {"분류": b, "단원": d, "단어": w, "뜻": m}
    }

# 안전한 숫자 추출 함수
def extract_day_number(x):
    if pd.isna(x):
        return None
    m = re.search(r'\d+', str(x))
    return int(m.group()) if m else None

def apply_schema(raw_df, schema):
    df = raw_df.copy()

    # --- 매핑이 요구하는 최대 열까지 자르도록 수정 (이전 답변에 안내했던 부분) ---
    mp = schema.get("mapping", {"분류": -1, "단원": 1, "단어": 2, "뜻": 3})
    mapped_max = max([v for v in mp.values() if isinstance(v, int) and v != -1] + [0])
    declared_use_cols = schema.get("use_cols", df.shape[1])
    use_cols = max(declared_use_cols, mapped_max)
    df = df.iloc[:, :min(use_cols, df.shape[1])]

    # 헤더 스킵
    skip = schema.get("header_skip", 1)
    if skip > 0 and df.shape[0] > 0:
        df = df.iloc[skip:, :]

    # 1-based -> 0-based 선택
    def pick(col_idx):
        if col_idx == -1:
            return pd.Series([""] * len(df))
        idx0 = col_idx - 1
        if 0 <= idx0 < df.shape[1]:
            return df.iloc[:, idx0]
        return pd.Series([""] * len(df))

    out = pd.DataFrame({
        "분류": pick(mp.get("분류", -1)),
        "단원": pick(mp.get("단원", 1)),
        "단어": pick(mp.get("단어", 2)),
        "뜻":   pick(mp.get("뜻",   3)),
    }).reset_index(drop=True)

    for col in ["분류", "단어", "뜻"]:
        out[col] = out[col].apply(normalize_cell)

    # --- 단원 숫자 추출 + 아래로 채우기---
    out["단원"] = out["단원"].apply(extract_day_number).ffill()
    # 숫자형으로 정리 (비어있으면 NA 유지)
    try:
        out["단원"] = out["단원"].astype("Int64")
    except Exception:
        # 판다스 버전에 따라 실패할 수 있음 -> 일반 int 캐스팅 시도 (NA가 없다는 가정)
        if out["단원"].isna().any():
            pass
        else:
            out["단원"] = out["단원"].astype(int)

    # --- 내용 없는 줄 제거: 단어/뜻 모두 빈칸이면 드랍 ---
    out = out[~((out["단어"] == "") & (out["뜻"] == ""))].reset_index(drop=True)

    # 카테고리형 유지
    out["분류"] = out["분류"].astype("category")
    return out



def get_schema_for_file(base_name, raw_df, memo, force_reset=False):
    if (not force_reset) and base_name in memo:
        sc = memo[base_name]
        preview_cols = preview_df(raw_df)
        print(f"\n[설정 불러옴] {base_name} → {sc['mode']} / use_cols={sc['use_cols']} / header_skip={sc['header_skip']} / mapping={sc['mapping']}")
        return sc

    print(f"\n[설정 필요] {base_name}")
    sc = ask_schema_interactively(base_name, raw_df)
    memo[base_name] = sc
    save_memory(memo)
    print("[저장 완료] schema_memory.txt 에 기록했습니다.\n")
    return sc

def main():
    files = list_source_files(BASE_DIR_SRC)
    print_file_menu(files)

    memo = load_memory()
    count = 1

    while True:
        try:
            sel_raw = input(f"파일 번호 (1~{len(files)}) 또는 'Q' 종료: ").strip()
            if sel_raw.upper() == "Q":
                print("종료합니다.")
                break
            sel = int(sel_raw)
            if not (1 <= sel <= len(files)):
                raise ValueError("목록에 없는 번호입니다.")

            excel_path = os.path.join(BASE_DIR_SRC, files[sel - 1])
            base_name = os.path.splitext(os.path.basename(excel_path))[0]

            raw_df = read_raw_sheet(excel_path)

            excel_path = os.path.join(BASE_DIR_SRC, files[sel - 1])
            base_name = os.path.splitext(os.path.basename(excel_path))[0]
            raw_df = read_raw_sheet(excel_path)

            # 저장된 설정 확인 → 있으면 먼저 보여주고, 그 다음 재설정 여부 묻기
            saved = memo.get(base_name)
            if saved:
                
                # 필요하면 프리뷰도 다시 보여줘서 참고하게
                _ = preview_df(raw_df)

                print(f"\n[저장된 설정 발견] {base_name}")
                print(f" mode        : {saved.get('mode')}")
                print(f" use_cols    : {saved.get('use_cols')}")
                print(f" header_skip : {saved.get('header_skip')}")
                print(f" mapping     : {saved.get('mapping')}")

                ans = input("저장된 설정을 사용할까요? (엔터=예) / 재설정하려면 'R': ").strip().upper()
                if ans == "R":
                    print("\n[재설정 진행]")
                    schema = ask_schema_interactively(base_name, raw_df)
                    memo[base_name] = schema
                    save_memory(memo)
                    print("[저장 완료] schema_memory.txt 에 기록했습니다.\n")
                else:
                    schema = saved
            else:
                # 저장된 설정이 없으면 바로 설정 진입
                print(f"\n[처음 사용하는 파일] {base_name} → 설정을 진행합니다.")
                schema = ask_schema_interactively(base_name, raw_df)
                memo[base_name] = schema
                save_memory(memo)
                print("[저장 완료] schema_memory.txt 에 기록했습니다.\n")

            day_input = input("DAY 번호 입력 (공백 구분, 예: 2 3 또는 4 7): ")
            days = parse_days(day_input)
            if not days:
                raise ValueError("DAY 번호를 올바르게 입력하세요.")

        except ValueError as e:
            print(f"입력 오류: {e}")
            input("\n[엔터를 눌러 종료합니다]")
            sys.exit(1)

        try:
            df = apply_schema(raw_df, schema)
        except Exception as e:
            print(f"정규화 중 오류: {e}")
            input("\n[엔터를 눌러 종료합니다]")
            sys.exit(1)

        filtered_df = df[df['단원'].isin(days)].reset_index(drop=True)
        if filtered_df.empty:
            print("해당 DAY 범위에 데이터가 없습니다. 다시 시도하세요.\n")
            continue

        shuffled_df = filtered_df.sample(frac=1, random_state=None).reset_index(drop=True)
        original_df = shuffled_df.copy()

        for idx in range(len(shuffled_df)):
            blank_col = random.choice(['단어', '뜻'])
            shuffled_df.at[idx, blank_col] = ""

        quiz_df = shuffled_df[['단어', '뜻']].copy()
        quiz_df.insert(0, '번호', range(1, len(quiz_df) + 1))

        day_tag = f"D{days[0]}" if len(days) == 1 else f"D{days[0]}-{days[-1]}"
        quiz_path = os.path.join(BASE_DIR_OUT, f"{count}번_{base_name}_{day_tag}_퀴즈_시험.xlsx")
        ans_path  = os.path.join(BASE_DIR_OUT, f"{count}번_{base_name}_{day_tag}_퀴즈_답안.xlsx")

        quiz_df.to_excel(quiz_path, index=False, engine="openpyxl")
        original_df.to_excel(ans_path, index=False, engine="openpyxl")

        print(f"생성완료 → 시험: {os.path.basename(quiz_path)} / 답안: {os.path.basename(ans_path)}\n")
        count += 1

if __name__ == "__main__":
    main()
