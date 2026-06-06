import argparse
import io
import os
import pandas as pd

__version__ = '1.3.0'
__updated__ = '2026-06-06'
__features__ = [
    ('vlookup',  '以 ID 跨檔 VLOOKUP 合併，輸出 xlsx'),
    ('extract',  '從來源檔篩選指定欄位，輸出 xlsx'),
    ('txt2xlsx', '文字檔轉 xlsx，數字欄位自動加總'),
]
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── 共用樣式 ─────────────────────────────────────────────────────────────────

HEADER_FONT  = Font(name='Arial', bold=True, color='FFFFFF')
HEADER_FILL  = PatternFill('solid', start_color='2F5597')
DATA_FONT    = Font(name='Arial', size=11)
CENTER       = Alignment(horizontal='center', vertical='center')
LEFT         = Alignment(horizontal='left',   vertical='center')
RIGHT        = Alignment(horizontal='right',  vertical='center')
THIN         = Side(style='thin', color='BFBFBF')
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALT_FILL     = PatternFill('solid', start_color='DCE6F1')
WHITE_FILL   = PatternFill('solid', start_color='FFFFFF')


def detect_style(col: str, dtype) -> str:
    """依 dtype 判斷欄位格式類型：text / date / number"""
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return 'date'
    if pd.api.types.is_numeric_dtype(dtype):
        return 'number'
    return 'text'


def write_sheet(ws, df: pd.DataFrame):
    """將 DataFrame 寫入 worksheet，自動套用格式"""
    cols       = df.columns.tolist()
    col_styles = {col: detect_style(col, df[col].dtype) for col in cols}

    # 標題列
    for c, h in enumerate(cols, start=1):
        cell           = ws.cell(row=1, column=c, value=h)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        cell.border    = BORDER

    # 資料列
    for r, row in enumerate(df.itertuples(index=False), start=2):
        fill = ALT_FILL if r % 2 == 0 else WHITE_FILL
        for c, (col, val) in enumerate(zip(cols, row), start=1):
            if val is pd.NaT or (isinstance(val, float) and np.isnan(val)):
                val = None
            cell        = ws.cell(row=r, column=c, value=val)
            cell.font   = DATA_FONT
            cell.fill   = fill
            cell.border = BORDER

            style = col_styles[col]
            if style == 'date':
                cell.alignment     = CENTER
                cell.number_format = 'YYYY-MM-DD'
            elif style == 'number':
                cell.alignment     = RIGHT
                cell.number_format = '#,##0'
            else:  # text：保留原格式，前導零不消失
                cell.alignment     = LEFT
                cell.number_format = '@'

    # 欄寬
    width_map = {'text': 12, 'date': 13, 'number': 14}
    for c, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width_map[col_styles[col]]

    ws.row_dimensions[1].height = 22


def add_source_note(ws, row: int, note: str):
    cell      = ws.cell(row=row, column=1, value=f'{note}  |  excel_tools v{__version__}')
    cell.font = Font(name='Arial', italic=True, color='808080', size=9)


# ── 共用：讀取 .ods / .xlsx / .xls ──────────────────────────────────────────

def _read_excel_auto(path: str, dtype: dict | None = None) -> pd.DataFrame:
    """依副檔名自動選擇引擎讀取試算表，並保留前導零欄位。"""
    ext    = path.rsplit('.', 1)[-1].lower()
    engine = 'odf' if ext == 'ods' else None   # xlsx/xls 用 openpyxl（預設）

    df_raw = pd.read_excel(path, engine=engine, dtype=str)
    df     = pd.read_excel(path, engine=engine, dtype=dtype or {})

    # 有前導零的欄位強制保留文字
    for col in df.columns:
        if df_raw[col].astype(str).str.match(r'^0\d+$').any():
            df[col] = df_raw[col]

    # 移除 Total 合計列、全空列、備註列（由 txt2xlsx 加入的尾端資料）
    first = df.columns[0]
    mask  = df[first].astype(str).str.strip().isin(['Total', 'nan', '']) | df[first].isna()
    if mask.any():
        df = df.loc[:mask.idxmax() - 1]   # 保留第一個 Total/空列之前的資料

    df = df.dropna(how='all').reset_index(drop=True)

    # float 欄位若無小數則轉回 int（避免 303771.0 這樣的顯示）
    for col in df.select_dtypes(include='float64').columns:
        if df[col].dropna().apply(lambda x: x == int(x)).all():
            df[col] = df[col].astype('Int64')

    return df


# ── 功能 1：VLOOKUP ──────────────────────────────────────────────────────────

def run_vlookup(src1: str, src2: str, output: str):
    """
    以 test2 的 ID 查找 test1 的薪資資料，合併後輸出。

    Args:
        src1   : test1 檔案路徑
        src2   : test2 檔案路徑
        output : 輸出 xlsx 路徑
    """
    df1 = _read_excel_auto(src1, dtype={'ID': str})
    df2 = _read_excel_auto(src2, dtype={'ID': str})

    # 自動偵測並解析日期欄位
    for col in df2.columns:
        if 'date' in col.lower():
            try:
                df2[col] = pd.to_datetime(df2[col])
            except Exception:
                pass

    merged = df2.merge(df1, on='ID', how='left')
    # 欄位順序：先 df2 所有欄，再 df1 不重複的欄
    cols   = list(df2.columns) + [c for c in df1.columns if c not in df2.columns]
    merged = merged[cols]

    wb = Workbook()
    ws = wb.active
    ws.title = 'VLOOKUP Result'
    write_sheet(ws, merged)
    add_source_note(ws, len(merged) + 3,
                    f'Source: {src1} + {src2}  |  VLOOKUP on ID')
    try:
        wb.save(output)
    except PermissionError:
        raise PermissionError(f'無法儲存 {output}，請先關閉該 Excel 檔案後再執行。')
    print(f'[VLOOKUP] 已儲存 → {output}')
    print(merged.to_string(index=False))


# ── 功能 2：欄位篩選輸出 ──────────────────────────────────────────────────────

def run_column_extract(src: str, columns: list[str], output: str):
    """
    從 src 中指定欄位，篩選後輸出到新檔案。

    Args:
        src     : 來源檔案路徑（.ods / .xlsx）
        columns : 要保留的欄位名稱清單，例如 ['ID', 'Net Pay']
        output  : 輸出 xlsx 路徑
    """
    df = _read_excel_auto(src)

    # 檢查欄位是否存在
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f'找不到欄位：{missing}，可用欄位為 {df.columns.tolist()}')

    result = df[columns].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Extract'
    write_sheet(ws, result)
    add_source_note(ws, len(result) + 3,
                    f'Source: {src}  |  Columns: {", ".join(columns)}')
    try:
        wb.save(output)
    except PermissionError:
        raise PermissionError(f'無法儲存 {output}，請先關閉該 Excel 檔案後再執行。')
    print(f'[Extract] 已儲存 → {output}')
    print(result.to_string(index=False))


# ── 功能 3：文字檔轉 xlsx + 數字欄位加總 ─────────────────────────────────────

def _read_txt(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    讀取文字檔，自動偵測分隔符，清除尾端多餘分隔符。
    回傳 (df_raw 全文字版, df_typed 自動轉型版)。
    """
    with open(path, 'r', encoding='utf-8-sig') as f:
        raw_lines = f.readlines()

    # 偵測分隔符
    sample = ''.join(raw_lines[:5])
    sep = '|' if '|' in sample else ('\t' if '\t' in sample else ',')

    # 去除每列尾端多餘的分隔符
    cleaned = [line.rstrip().rstrip(sep) for line in raw_lines]
    content = '\n'.join(cleaned)

    df_raw   = pd.read_csv(io.StringIO(content), sep=sep, dtype=str,
                           skipinitialspace=True)
    df_typed = pd.read_csv(io.StringIO(content), sep=sep,
                           skipinitialspace=True)

    # 欄名去空白
    df_raw.columns   = df_raw.columns.str.strip()
    df_typed.columns = df_typed.columns.str.strip()

    # 內容去空白（raw）
    df_raw = df_raw.apply(lambda s: s.str.strip())

    return df_raw, df_typed


def run_txt_to_xlsx(folder: str, output_folder: str = None):
    """
    掃描 folder 內所有 .txt / .csv / .tsv，逐一：
      1. 自動偵測分隔符並讀取
      2. 辨識文字欄位（前導零）與數字欄位
      3. 在最後一列加入數字欄位合計（黃底粗體）
      4. 輸出為 <原檔名>_converted.xlsx

    Args:
        folder        : 來源資料夾
        output_folder : 輸出資料夾（預設同 folder）
    """
    if output_folder is None:
        output_folder = folder

    TEXT_EXT  = {'.txt', '.csv', '.tsv'}
    txt_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if os.path.splitext(f)[1].lower() in TEXT_EXT
    ]

    if not txt_files:
        print('[TXT2XLSX] 找不到文字檔（.txt / .csv / .tsv）')
        return

    TOTAL_FONT = Font(name='Arial', bold=True, size=11)
    TOTAL_FILL = PatternFill('solid', start_color='FFF2CC')

    for txt_path in txt_files:
        base     = os.path.splitext(os.path.basename(txt_path))[0]
        out_path = os.path.join(output_folder, f'{base}_converted.xlsx')

        df_raw, df_typed = _read_txt(txt_path)

        cols     = [c for c in df_raw.columns if c in df_typed.columns]
        df_raw   = df_raw[cols].dropna(how='all').reset_index(drop=True)
        df_typed = df_typed[cols].dropna(how='all').reset_index(drop=True)

        text_cols = {col for col in cols if df_raw[col].str.match(r'^0\d+$').any()}
        num_cols  = [col for col in cols
                     if col not in text_cols and pd.api.types.is_numeric_dtype(df_typed[col])]

        df = df_typed.copy()
        for col in text_cols:
            df[col] = df_raw[col]

        # 合計列資料
        sum_row = {
            col: (df_typed[col].sum() if col in num_cols
                  else ('Total' if col == cols[0] else ''))
            for col in cols
        }

        # 建立 xlsx
        wb = Workbook()
        ws = wb.active
        ws.title = 'Data'

        # 主資料列（不含合計）
        write_sheet(ws, df)

        # 合計列（緊接資料後，列號 = 資料列數 + 標題列 + 1）
        col_styles    = {col: detect_style(col, df[col].dtype) for col in cols}
        total_row_num = len(df) + 2

        for c, col in enumerate(cols, start=1):
            cell            = ws.cell(row=total_row_num, column=c, value=sum_row[col])
            cell.font       = TOTAL_FONT
            cell.fill       = TOTAL_FILL
            cell.border     = BORDER
            if col_styles[col] == 'number':
                cell.alignment     = RIGHT
                cell.number_format = '#,##0'
            else:
                cell.alignment     = LEFT
                cell.number_format = '@'

        add_source_note(ws, total_row_num + 2,
                        f'Source: {os.path.basename(txt_path)}')

        try:
            wb.save(out_path)
        except PermissionError:
            raise PermissionError(f'無法儲存 {out_path}，請先關閉該 Excel 檔案後再執行。')

        df_preview = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)
        print(f'[TXT2XLSX] {os.path.basename(txt_path)} → {out_path}')
        print(df_preview.to_string(index=False))
        print()


# ── CLI 入口 ──────────────────────────────────────────────────────────────────

BASE = r'C:\Users\ASUS\repo\vlookup'

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description=f'Excel 工具集  v{__version__}',
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        '--version', action='version',
        version=(
            f'excel_tools v{__version__}  ({__updated__})\n'
            + '\n'.join(f'  {cmd:<10}{desc}' for cmd, desc in __features__)
        ),
    )
    subparsers = parser.add_subparsers(dest='command', required=True)

    # -- vlookup 子命令 --
    p_vl = subparsers.add_parser('vlookup', help='以 test2 ID 查找 test1 薪資資料並輸出')
    p_vl.add_argument('--src1',   default=f'{BASE}/test1.ods',          help='來源檔案1 (預設: test1.ods)')
    p_vl.add_argument('--src2',   default=f'{BASE}/test2.ods',          help='來源檔案2 (預設: test2.ods)')
    p_vl.add_argument('--output', default=f'{BASE}/output_vlookup.xlsx', help='輸出路徑 (預設: output_vlookup.xlsx)')

    # -- extract 子命令 --
    p_ex = subparsers.add_parser('extract', help='從 test1 篩選指定欄位並輸出')
    p_ex.add_argument('--src',     default=f'{BASE}/test1.ods',           help='來源檔案 (預設: test1.ods)')
    p_ex.add_argument('--columns', default='ID,Net Pay',                  help='欄位名稱，逗號分隔 (預設: ID,Net Pay)')
    p_ex.add_argument('--output',  default=f'{BASE}/output_extract.xlsx', help='輸出路徑 (預設: output_extract.xlsx)')

    # -- txt2xlsx 子命令 --
    p_t2x = subparsers.add_parser('txt2xlsx', help='文字檔轉 xlsx，數字欄位自動加總')
    p_t2x.add_argument('--folder',        default=BASE, help=f'文字檔所在資料夾 (預設: {BASE})')
    p_t2x.add_argument('--output-folder', default=None, help='輸出資料夾 (預設: 同來源資料夾)')

    args = parser.parse_args()

    if args.command == 'vlookup':
        run_vlookup(args.src1, args.src2, args.output)

    elif args.command == 'extract':
        columns = [c.strip() for c in args.columns.split(',')]
        run_column_extract(args.src, columns, args.output)

    elif args.command == 'txt2xlsx':
        run_txt_to_xlsx(args.folder, args.output_folder)
