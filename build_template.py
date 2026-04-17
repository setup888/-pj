"""
執務表自動化テンプレート生成スクリプト (完全データ駆動版)

設計思想:
- ポジション定義シートで「ポジション × 時間帯 の ×」をマトリクス管理
- 名簿は氏名のみ
- 当日チェックで「その日のポジション + 追加除外時間帯」を1人1行
- VBAにはハードコードなし、すべてシートで調整可

生成後の手順:
  1. マクロ有効ブック (.xlsm) で保存
  2. VBAコードシートの内容を標準モジュールに貼付
  3. ボタン2つ (GenerateShift / ShowDateFromHistory) 配置
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date

OUT_PATH = "執務表自動化.xlsx"

TIME_SLOTS = [
    "8:40〜9", "9〜10", "10〜11", "11〜12", "12〜13", "13〜14",
    "14〜15", "15〜16", "16〜17", "17〜18", "18〜19", "19〜20",
    "20〜21", "21〜22", "22〜23", "23〜24",
    "0〜1", "1〜2", "2〜3", "3〜4", "4〜5", "5〜6",
    "6〜7", "7〜8", "8〜8:40",
]
N_SLOTS = len(TIME_SLOTS)

# ユーザー指定のポジション + 共通 duty
# 各ポジションの × は後でマトリクスで設定 (ユーザー編集可)
POSITIONS_INITIAL = [
    "ポンプ隊",
    "救助隊",
    "はしご隊",
    "救急隊",
    "伝令",
    "通信担当",
    "情報担当",
    "情報員",
    "署隊長伝令",
    "残留",
    "署隊本部支援員",
    "その他",
    # 以下は duty (時間帯制限あり)
    "当直士長",
    "食当",
    "休暇",
    "研修/出向",
]

# デフォルト × 設定 (ユーザー編集可、初期値)
POSITION_DEFAULT_BLOCKS = {
    "救急隊": set(TIME_SLOTS),
    "休暇": set(TIME_SLOTS),
    "研修/出向": set(TIME_SLOTS),
    "食当": {"14〜15", "15〜16", "16〜17"},
    "当直士長": {"18〜19", "19〜20", "6〜7", "7〜8"},
}

# セルスタイル
thin = Side(border_style="thin", color="000000")
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
FILL_HEADER = PatternFill("solid", fgColor="DCE6F1")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_OUTPUT = PatternFill("solid", fgColor="E2EFDA")
FILL_BLOCK = PatternFill("solid", fgColor="BFBFBF")
FONT_BOLD = Font(bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_header(cell):
    cell.font = FONT_BOLD
    cell.alignment = ALIGN_CENTER
    cell.fill = FILL_HEADER
    cell.border = BORDER_ALL


def style_input(cell):
    cell.alignment = ALIGN_CENTER
    cell.fill = FILL_INPUT
    cell.border = BORDER_ALL


def style_output(cell):
    cell.alignment = ALIGN_CENTER
    cell.fill = FILL_OUTPUT
    cell.border = BORDER_ALL


# =============================================================
# 名簿 (氏名のみ)
# =============================================================
def build_roster(ws):
    ws.title = "名簿"
    headers = ["No", "氏名"]
    for i, h in enumerate(headers, start=1):
        style_header(ws.cell(row=1, column=i, value=h))
    N = 30
    for i in range(N):
        r = 2 + i
        ws.cell(row=r, column=1, value=i + 1).alignment = ALIGN_CENTER
        ws.cell(row=r, column=1).border = BORDER_ALL
        cell = ws.cell(row=r, column=2)
        cell.border = BORDER_ALL
        cell.fill = FILL_INPUT
        cell.alignment = ALIGN_LEFT

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 22

    # サンプル
    samples = ["例: 原田 陽一郎", "例: 梅村 侑志", "例: 鍋谷 昇"]
    for i, name in enumerate(samples):
        ws.cell(row=2 + i, column=2, value=name)


# =============================================================
# ポジション定義 (マトリクス: ポジション × 時間帯)
# =============================================================
def build_positions(ws):
    ws.title = "ポジション定義"
    ws["A1"] = "ポジション × 時間帯 の割当可否設定"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = ("× = そのポジションの人はこの時間帯に 執務表 を割り当てない。"
                "空白 = 割当可。ここを編集するだけでルール変更可能。")
    ws["A2"].font = Font(italic=True, size=9)

    # ヘッダ行 (行4)
    style_header(ws.cell(row=4, column=1, value="ポジション"))
    for j, slot in enumerate(TIME_SLOTS, start=2):
        style_header(ws.cell(row=4, column=j, value=slot))

    # データ行 (50行バッファ)
    N_POS_ROWS = 50
    for i in range(N_POS_ROWS):
        r = 5 + i
        for c in range(1, 2 + N_SLOTS):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER_ALL
            cell.alignment = ALIGN_CENTER
            if c == 1:
                cell.fill = FILL_INPUT

    # 初期値を埋める
    for i, pos in enumerate(POSITIONS_INITIAL):
        r = 5 + i
        ws.cell(row=r, column=1, value=pos)
        blocks = POSITION_DEFAULT_BLOCKS.get(pos, set())
        for j, slot in enumerate(TIME_SLOTS, start=2):
            if slot in blocks:
                c = ws.cell(row=r, column=j, value="×")
                c.fill = FILL_BLOCK

    # ドロップダウン: × / 空白 (他のマスへ)
    dv = DataValidation(type="list", formula1='"×"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B5:{get_column_letter(1 + N_SLOTS)}{4 + N_POS_ROWS}")

    ws.column_dimensions["A"].width = 18
    for j in range(2, 2 + N_SLOTS):
        ws.column_dimensions[get_column_letter(j)].width = 8

    ws.freeze_panes = "B5"


# =============================================================
# 当日チェック (1行1人)
# =============================================================
def build_daily_input(ws):
    ws.title = "当日チェック"
    ws["A1"] = "当日チェック"
    ws["A1"].font = Font(bold=True, size=14)

    # 上部: 当番日 / 休日フラグ
    ws["A3"] = "当番日 (yyyy/m/d)"
    ws["A3"].font = FONT_BOLD
    ws["A3"].border = BORDER_ALL
    ws["B3"] = date(2026, 4, 16)
    ws["B3"].number_format = "yyyy/m/d"
    style_input(ws["B3"])

    ws["A4"] = "休日フラグ (1=休日/0=平日)"
    ws["A4"].font = FONT_BOLD
    ws["A4"].border = BORDER_ALL
    ws["B4"] = 0
    style_input(ws["B4"])
    ws["C4"] = "※ このシートで各隊員の『今日のポジション』を選び、必要なら除外時間帯を追加"
    ws["C4"].font = Font(italic=True, size=9)
    ws.merge_cells("C4:J4")

    # マトリクスヘッダ (行6)
    headers = [
        "No", "氏名", "今日のポジション",
        "除外1 開始", "除外1 終了",
        "除外2 開始", "除外2 終了",
        "除外3 開始", "除外3 終了",
        "備考",
    ]
    for i, h in enumerate(headers, start=1):
        style_header(ws.cell(row=6, column=i, value=h))

    N = 30
    for i in range(N):
        r = 7 + i
        roster_row = 2 + i
        # No
        ws.cell(row=r, column=1, value=i + 1).alignment = ALIGN_CENTER
        ws.cell(row=r, column=1).border = BORDER_ALL
        # 氏名 (formula)
        nm = ws.cell(row=r, column=2,
                     value=f'=IF(名簿!B{roster_row}="","",名簿!B{roster_row})')
        nm.border = BORDER_ALL
        nm.alignment = ALIGN_LEFT
        # ポジション + 除外列
        for c in range(3, 11):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER_ALL
            cell.alignment = ALIGN_CENTER
            if c < 10:
                cell.fill = FILL_INPUT

    # ドロップダウン: ポジション (ポジション定義 A5:A54)
    dv_pos = DataValidation(type="list",
                            formula1="=ポジション定義!$A$5:$A$54",
                            allow_blank=True)
    ws.add_data_validation(dv_pos)
    dv_pos.add(f"C7:C{6 + N}")

    # ドロップダウン: 時間帯 (除外1-3 開始/終了)
    slot_list = ",".join(TIME_SLOTS)
    dv_slot = DataValidation(type="list", formula1=f'"{slot_list}"', allow_blank=True)
    ws.add_data_validation(dv_slot)
    dv_slot.add(f"D7:I{6 + N}")

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    for c in "DEFGHI":
        ws.column_dimensions[c].width = 10
    ws.column_dimensions["J"].width = 22

    ws.freeze_panes = "C7"


# =============================================================
# 履歴
# =============================================================
def build_history(ws):
    ws.title = "履歴"
    ws["A1"] = "過去当番の一次指定者 (GenerateShift 実行時に自動追記)"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "月初に前月分をここに貼り付け可。列: 当番日/時間帯/通信指令/受付。1当番=25行。"
    ws["A2"].font = Font(italic=True, size=9)

    headers = ["当番日", "時間帯", "通信指令", "受付"]
    for i, h in enumerate(headers, start=1):
        style_header(ws.cell(row=4, column=i, value=h))

    for r in range(5, 5 + 400):
        ws.cell(row=r, column=1).number_format = "yyyy/m/d"
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).border = BORDER_ALL

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.freeze_panes = "A5"


# =============================================================
# 執務表
# =============================================================
def build_output(ws):
    ws.title = "執務表"
    ws["A1"] = "執　務　表"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:G1")
    ws["A1"].alignment = ALIGN_CENTER

    # 表示日付
    ws["A2"] = "表示日付"
    ws["A2"].font = FONT_BOLD
    ws["A2"].border = BORDER_ALL
    ws["A2"].alignment = ALIGN_CENTER
    ws["B2"].number_format = "yyyy/m/d"
    ws["B2"].alignment = ALIGN_CENTER
    ws["B2"].border = BORDER_ALL
    ws["B2"].fill = FILL_INPUT
    ws["C2"] = "← 過去日を入力して ShowDateFromHistory ボタンで過去の表示可"
    ws["C2"].alignment = ALIGN_LEFT
    ws["C2"].font = Font(italic=True, size=9)
    ws.merge_cells("C2:G2")

    # 2段ヘッダ
    ws["A3"] = "事務区分"
    ws["B3"] = "通信指令"
    ws["E3"] = "受付"
    ws.merge_cells("B3:D3")
    ws.merge_cells("E3:G3")
    ws["A4"] = "時間"
    ws["B4"] = "一次指定者"
    ws["C4"] = "指定変更 従事者"
    ws["D4"] = "時間"
    ws["E4"] = "一次指定者"
    ws["F4"] = "指定変更 従事者"
    ws["G4"] = "時間"
    for addr in ["A3", "B3", "E3", "A4", "B4", "C4", "D4", "E4", "F4", "G4"]:
        style_header(ws[addr])
    ws.merge_cells("A3:A4")

    for i, slot in enumerate(TIME_SLOTS, start=5):
        ws.cell(row=i, column=1, value=slot).border = BORDER_ALL
        ws.cell(row=i, column=1).alignment = ALIGN_CENTER
        for col in (2, 5):
            style_output(ws.cell(row=i, column=col))
        for col in (3, 4, 6, 7):
            ws.cell(row=i, column=col).border = BORDER_ALL

    last = 5 + N_SLOTS
    ws.cell(row=last, column=1, value="二次指定者").font = FONT_BOLD
    ws.cell(row=last, column=1).border = BORDER_ALL
    for col in range(2, 8):
        ws.cell(row=last, column=col).border = BORDER_ALL

    ws.cell(row=last + 1, column=1, value="備考").font = FONT_BOLD
    ws.merge_cells(start_row=last + 1, start_column=2,
                   end_row=last + 4, end_column=7)
    ws.cell(row=last + 1, column=2,
            value="・夜20時以降は前日と同じ時間帯で同じ人を避ける\n"
                  "・10-17時は一人1回以上\n"
                  "・12勤と17勤は別人\n"
                  "・深夜(22-4時)は一人1回").alignment = Alignment(wrap_text=True, vertical="top")
    for rr in range(last + 1, last + 5):
        for cc in range(1, 8):
            ws.cell(row=rr, column=cc).border = BORDER_ALL

    ws.column_dimensions["A"].width = 10
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 14
    ws["A" + str(last + 6)] = "↑ 一次指定者は GenerateShift で自動生成。指定変更欄は当日手書き。"


# =============================================================
# 使い方
# =============================================================
def build_help(ws):
    ws.title = "使い方"
    lines = [
        "【執務表自動化 使い方】",
        "",
        "◆ 初回のみ",
        " 1. ポジション定義シートを確認 (必要なら × を編集・ポジション追加)",
        " 2. 名簿シートに隊員の氏名を入力 (最大30名)",
        " 3. マクロ有効ブック(.xlsm) で保存",
        " 4. Alt+F11 → ThisWorkbook 右クリック → 挿入 → 標準モジュール",
        " 5. VBAコード シートの A12 以下を全コピー → 貼付",
        " 6. 執務表シートにボタン2つ配置",
        "    - ボタン1: マクロ「GenerateShift」",
        "    - ボタン2: マクロ「ShowDateFromHistory」",
        "",
        "◆ 月初の準備 (月1回)",
        " 1. 履歴シートに前月分をコピペ (列: 当番日/時間帯/通信/受付、1当番=25行)",
        "",
        "◆ 毎当番の作業",
        " 1. 当日チェックシート",
        "    - B3 当番日を更新",
        "    - 各隊員の「今日のポジション」をドロップダウンから選択",
        "    - 半日不在等あれば 除外1〜3 の時間帯を入れる",
        " 2. 執務表シート「本日を生成」ボタン",
        " 3. 印刷 or Word様式にコピペ",
        "",
        "◆ 過去の日を見たい",
        " - 執務表 B2 に日付 → ShowDateFromHistory ボタン",
        "",
        "◆ ポジションのルール変更",
        " - ポジション定義シートの × を編集するだけ",
        " - VBAの書き換え不要",
        "",
        "◆ よく使うポジション (初期登録済み)",
        " ポンプ隊 / 救助隊 / はしご隊 / 救急隊 / 伝令 / 通信担当 / 情報担当 /",
        " 情報員 / 署隊長伝令 / 残留 / 署隊本部支援員 / その他",
        " 当直士長 / 食当 / 休暇 / 研修・出向",
        "",
        "◆ 新ポジション追加",
        " - ポジション定義シートに行追加 → 時間帯に × を入れる",
        " - 当日チェックのドロップダウンに自動で出る",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 90


# =============================================================
# VBAコード (手動貼付用)
# =============================================================
def build_vba_code(ws):
    ws.title = "VBAコード"
    ws["A1"] = "■ VBA マクロ導入手順"
    ws["A1"].font = Font(bold=True, size=14)

    instructions = [
        "① このファイルを「名前を付けて保存」→ Excel マクロ有効ブック(.xlsm)",
        "② Alt + F11 で VBエディタ",
        "③ 左ペイン ThisWorkbook 右クリック → 挿入 → 標準モジュール",
        "④ 下記のコードを全て選択してコピー (A12セル→Ctrl+Shift+End→Ctrl+C)",
        "⑤ 追加した標準モジュールに貼付 (Ctrl+V)",
        "⑥ 執務表シートにボタン2つ配置し GenerateShift / ShowDateFromHistory を割当",
    ]
    for i, line in enumerate(instructions, start=2):
        ws.cell(row=i, column=1, value=line)

    header_row = len(instructions) + 3
    ws.cell(row=header_row, column=1,
            value="■ ここから下をコピー ↓").font = Font(bold=True, color="FF0000")

    start_row = header_row + 1
    try:
        with open("ShiftScheduler.bas", "r", encoding="utf-8") as f:
            code = f.read()
    except FileNotFoundError:
        ws.cell(row=start_row, column=1,
                value="(ShiftScheduler.bas が見つかりません)")
        return

    lines = code.split("\n")
    if lines and lines[0].startswith("Attribute VB_Name"):
        lines = lines[1:]
    for i, line in enumerate(lines):
        cell = ws.cell(row=start_row + i, column=1, value=line)
        cell.font = Font(name="Consolas", size=10)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)

    ws.column_dimensions["A"].width = 110


def main():
    wb = Workbook()
    ws_output = wb.active
    build_output(ws_output)

    ws_input = wb.create_sheet()
    build_daily_input(ws_input)

    ws_pos = wb.create_sheet()
    build_positions(ws_pos)

    ws_roster = wb.create_sheet()
    build_roster(ws_roster)

    ws_history = wb.create_sheet()
    build_history(ws_history)

    ws_help = wb.create_sheet()
    build_help(ws_help)

    ws_vba = wb.create_sheet()
    build_vba_code(ws_vba)

    wb._sheets = [ws_output, ws_input, ws_pos, ws_roster, ws_history, ws_help, ws_vba]
    wb.save(OUT_PATH)
    print(f"生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
