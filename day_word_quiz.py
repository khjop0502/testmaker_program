import pandas as pd
import random
import re
import os
import sys
import json

# 실행 위치 고정 (exe/py 공통)
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR_SRC = "src"
BASE_DIR_OUT = "result"
MEMO_PATH = "schema_memory.txt"  # JSON 저장 (파일명별 스키마 기억)

os.makedirs(BASE_DIR_OUT, exist_ok=True)

# -------------------- 유틸 --------------------

try:
    from wcwidth import wcswidth  # pip install wcwidth
except Exception:
    wcswidth = None

def display_len(s: str) -> int:
    s = "" if s is None else str(s)
    if wcswidth:
        w = wcswidth(s)
        return w if w >= 0 else len(s)
    # 폴백: 대략적인 2배폭 문자 보정 (간이)
    # 전각/한글/한자 등은 2칸으로 추정
    wide = sum(1 for ch in s if ord(ch) > 0x1100)
    return len(s) + wide  # 대략 보정

def ljust_display(s: str, width: int) -> str:
    s = "" if s is None else str(s)
    pad = max(0, width - display_len(s))
    return s + (" " * pad)

def normalize_cell(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = re.sub(r'\s+', ' ', s.strip())
    return s

def extract_day_number(x):
    if pd.isna(x):
        return None
    m = re.search(r'\d+', str(x))
    return int(m.group()) if m else None

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

# -------------------- 프리뷰 --------------------
def preview_df(df, max_cols=8, max_rows=4, show_row_labels=True):
    """
    프리뷰 출력 (열 안내 1줄 + 데이터 최대 4행)
    예:
      열 1: 분류 | 열 2: 단원 | 열 3: 단어 | 열 4: 뜻
      행 1 | 분류값 | 단원값 | 단어값 | 뜻값
      행 2 | ...
    반환: 프리뷰에 사용한 원본 컬럼 인덱스 리스트(cols, 0-based)
    """
    # 데이터 많은 열부터 상위 max_cols 선택
    nonnull_counts = df.notna().sum().sort_values(ascending=False)
    cols = list(nonnull_counts.index[:max_cols])  # 0-based 원본 인덱스
    prev = df[cols].head(max_rows).fillna("")

    # 각 열의 "라벨" 추출 (앞쪽 3행 내 첫 비어있지 않은 값)
    def first_nonempty(series):
        for v in series:
            s = "" if pd.isna(v) else str(v).strip()
            if s:
                return s
        return ""
    header_labels = []
    for c in cols:
        header_labels.append(first_nonempty(df[c].head(3)))

    # 열 너비 계산: 라벨/데이터 모두 고려 + 최소 폭 14
    col_widths = {}
    for i, c in enumerate(cols, start=1):
        label = f"열 {i}: {header_labels[i-1] or ''}".strip()
        max_len = display_len(label)
        for val in prev[c]:
            max_len = max(max_len, display_len(str(val)))
        col_widths[c] = max(max_len + 2, 14)

    # (선택) 행 라벨 폭
    row_label_w = 6 if show_row_labels else 0  # "행 10" 도 들어가도록 6칸

    # 헤더 라인: "열 1: 분류 | 열 2: 단원 | ..."
    header_line = (" " * row_label_w) + " | ".join(
        ljust_display(f"열 {i}: {header_labels[i-1]}", col_widths[c])
        for i, c in enumerate(cols, start=1)
    )

    print("\n[시트 미리보기]")
    print(header_line)

    # 데이터 행 출력 (최대 4행)
    for r_idx, (_, row) in enumerate(prev.iterrows(), start=1):
        row_prefix = (ljust_display(f"행 {r_idx}", row_label_w) + " | ") if show_row_labels else ""
        line = row_prefix + " | ".join(
            ljust_display(str(row[c]), col_widths[c]) for c in cols
        )
        print(line)
    print()

    return cols



def read_raw_sheet(xlsx_path):
    return pd.read_excel(xlsx_path, sheet_name=0, header=None)

# -------------------- 스키마 입력(수동만) --------------------
def guess_defaults(preview_cols, raw_df):
    """
    프리뷰 열들(0-based)에 대해 첫 행의 텍스트로 기본 매핑 후보 추정.
    반환: (분류, 단원, 단어, 뜻) → 프리뷰 번호(1..N) 또는 -1
    """
    header_texts = []
    for c in preview_cols:
        v = raw_df.iloc[0, c] if raw_df.shape[0] > 0 else ""
        s = "" if pd.isna(v) else str(v).strip().lower()
        header_texts.append(s)

    b, d, w, m = -1, 1, 2, 3  # 기본값

    for i, s in enumerate(header_texts, start=1):
        if not s:
            continue
        if any(k in s for k in ["day", "단원"]):
            d = i
        if any(k in s for k in ["english", "영어", "단어"]):
            w = i
        if any(k in s for k in ["뜻", "의미", "mean"]):
            m = i
        if any(k in s for k in ["분류", "품사", "category", "pos"]):
            b = i

    return b, d, w, m

def ask_schema_interactively(base_name, raw_df):
    """
    선택지 없이: 미리보기 → 헤더 스킵 여부 → 수동 매핑만 진행(프리뷰 번호 기반)
    """
    preview_cols = preview_df(raw_df)  # 0-based 원본 인덱스 리스트

    ans = input("첫 행을 제목/머리글로 보고 건너뛸까요? (y/n, 기본 y): ").strip().lower()
    header_skip = 1 if ans in ("", "y", "yes") else 0

    def_b, def_d, def_w, def_m = guess_defaults(preview_cols, raw_df)

    print("\n[열 매핑] 프리뷰의 왼쪽 번호(1..N)를 입력하세요. 해당 열이 없으면 -1")
    print("예) 분류=-1, 단원=Day/단원/챕터, 단어=English/영어, 뜻=의미/한글")

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
        if not (1 <= sel <= len(preview_cols)):
            return -1 if default_preview_pos == -1 else (preview_cols[default_preview_pos - 1] + 1)
        return preview_cols[sel - 1] + 1  # 1-based 원본 인덱스

    # ✅ 라벨을 더 직관적으로
    b = ask_col("분류(품사/카테고리)", def_b)
    d = ask_col("단원(단원/챕터/Day)",  def_d)
    w = ask_col("단어(영어/English)",   def_w)
    m = ask_col("뜻(의미/한글/Meaning)", def_m)

    mapped_max_1based = max([x for x in (b, d, w, m) if isinstance(x, int) and x != -1] + [0])
    use_cols = max(3, mapped_max_1based) if b == -1 else max(4, mapped_max_1based)

    schema = {
        "mode": "manual",
        "use_cols": use_cols,
        "header_skip": header_skip,
        "mapping": {"분류": b, "단원": d, "단어": w, "뜻": m}
    }

    print("\n[설정 요약]")
    print(f" 제목줄 스킵: {schema['header_skip']}행")
    print(f" 매핑(원본 1-based, -1=없음): {schema['mapping']}\n")
    return schema

# -------------------- 설정 미리보기(저장값 표시 + 예시행) --------------------
def _get_cell(raw_df, row_idx, col_1based):
    if col_1based is None or col_1based == -1:
        return ""
    c0 = col_1based - 1
    if not (0 <= row_idx < raw_df.shape[0]) or not (0 <= c0 < raw_df.shape[1]):
        return ""
    return normalize_cell(raw_df.iat[row_idx, c0])

def show_saved_schema_with_example(saved, raw_df):
    """
    저장된 스키마를 사람이 보기 쉽도록 출력.
    - 프리뷰 표 출력
    - header_skip 반영한 '예시 행'(= header_skip 이후 첫 행)을 1줄로 매칭해 보여줌
    """
    preview_df(raw_df)  # 컨텍스트용 표

    hs = int(saved.get("header_skip", 1) or 0)
    row_idx = hs  # 0-based: 헤더를 hs만큼 건너뛴 뒤 첫 데이터 행
    mp = saved.get("mapping", {})
    sample = {
        "분류": _get_cell(raw_df, row_idx, mp.get("분류", -1)),
        "단원": _get_cell(raw_df, row_idx, mp.get("단원", 1)),
        "단어": _get_cell(raw_df, row_idx, mp.get("단어", 2)),
        "뜻":   _get_cell(raw_df, row_idx, mp.get("뜻",   3)),
    }

    print("[저장된 설정]")
    print(f" 제목줄 스킵: {hs}행")
    print(f" 매핑(원본 1-based, -1=없음): {mp}")

    # 예시 행(분류는 -1이거나 값이 비면 생략)
    parts = []
    if mp.get("분류", -1) != -1 and sample["분류"] != "":
        parts.append(f"분류='{sample['분류']}'")
    parts.append(f"단원='{sample['단원']}'")
    parts.append(f"단어='{sample['단어']}'")
    parts.append(f"뜻='{sample['뜻']}'")
    print(f"[예시 행] (제목줄 {hs}행 스킵 → 미리보기 기준 {hs+1}행)\n")
    print(" " + " | ".join(parts) + "\n")

# -------------------- 스키마 적용 --------------------
def apply_schema(raw_df, schema):
    df = raw_df.copy()

    mp = schema.get("mapping", {"분류": -1, "단원": 1, "단어": 2, "뜻": 3})
    mapped_max = max([v for v in mp.values() if isinstance(v, int) and v != -1] + [0])
    declared_use_cols = schema.get("use_cols", df.shape[1])
    use_cols = max(declared_use_cols, mapped_max)
    df = df.iloc[:, :min(use_cols, df.shape[1])]

    skip = schema.get("header_skip", 1)
    if skip > 0 and df.shape[0] > 0:
        df = df.iloc[skip:, :]

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

    out["단원"] = out["단원"].apply(extract_day_number).ffill()
    try:
        out["단원"] = out["단원"].astype("Int64")
    except Exception:
        if out["단원"].isna().any():
            pass
        else:
            out["단원"] = out["단원"].astype(int)

    out = out[~((out["단어"] == "") & (out["뜻"] == ""))].reset_index(drop=True)
    out["분류"] = out["분류"].astype("category")
    return out

# -------------------- 메인 --------------------
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

            # 저장된 설정 있으면 먼저 보여주고, 사용/재설정 묻기
            saved = memo.get(base_name)
            if saved:
                show_saved_schema_with_example(saved, raw_df)
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
                # 없으면 바로 설정
                print(f"\n[처음 사용하는 파일] {base_name} → 설정을 진행합니다.")
                schema = ask_schema_interactively(base_name, raw_df)
                memo[base_name] = schema
                save_memory(memo)
                print("[저장 완료] schema_memory.txt 에 기록했습니다.\n")

            # DAY 입력
            day_input = input("DAY 번호 입력 (공백 구분, 예: 2 3 또는 4 7): ")
            days = parse_days(day_input)
            if not days:
                raise ValueError("DAY 번호를 올바르게 입력하세요.")

        except ValueError as e:
            print(f"입력 오류: {e}")
            input("\n[엔터를 눌러 종료합니다]")
            sys.exit(1)

        # 스키마 적용 및 출력
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

        # 단어/뜻 중 무작위로 빈칸 만들기
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
