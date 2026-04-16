"""
執務表自動化テンプレート生成スクリプト

実行すると「執務表自動化.xlsx」を生成する。
生成後の手順:
  1. Excel で開き「名前を付けて保存」→ 形式「Excel マクロ有効ブック (.xlsm)」で保存
  2. Alt+F11 で VBE を開き、ShiftScheduler.bas をインポート
  3. 執務表シートのボタン（自分で配置）に Sub GenerateShift を割り当て
  詳細は README.md を参照
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Side, Font, PatternFill, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT_PATH = "執務表自動化.xlsx"

# 時間枠: 画像の様式に合わせる。8:40 から翌 8:40 まで。
TIME_SLOTS = [
    "8:40〜9",
    "9〜10",
    "10〜11",
    "11〜12",
    "12〜13",
    "13〜14",
    "14〜15",
    "15〜16",
    "16〜17",
    "17〜18",
    "18〜19",
    "19〜20",
    "20〜21",
    "21〜22",
    "22〜23",
    "23〜24",
    "0〜1",
    "1〜2",
    "2〜3",
    "3〜4",
    "4〜5",
    "5〜6",
    "6〜7",
    "7〜8",
    "8〜8:40",
]

# 役職カテゴリ（勤務可否や優先度判定に使う）
ROLE_CATEGORIES = [
    "指揮者",       # 中隊長・小隊長クラス（指揮専従）
    "情報員",       # 情報担当
    "伝令",         # 伝令
    "通信担当",     # 通信専従
    "機関員",       # 指揮車機関員＝受付残留
    "一般",         # 一般隊員（通信・受付どちらも可）
    "警防力",       # 毎日勤務（日中×、夜18-22優先）
    "救助隊",       # 訓練期間中は日中×、17時以降に割当
]

# ---- セルスタイル ----
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


def build_roster(ws):
    """名簿シート: 30行のバッファ付き。氏名・役職カテゴリ・備考。"""
    ws.title = "名簿"
    headers = ["No", "氏名", "役職カテゴリ", "備考（訓練期間/毎日勤務 等）"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        style_header(c)
    for r in range(2, 32):  # 30名バッファ
        ws.cell(row=r, column=1, value=r - 1).alignment = ALIGN_CENTER
        ws.cell(row=r, column=1).border = BORDER_ALL
        for c in range(2, 5):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER_ALL
            cell.alignment = ALIGN_LEFT
            if c == 2:
                cell.fill = FILL_INPUT
    # 役職カテゴリのドロップダウン
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(ROLE_CATEGORIES) + '"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add(f"C2:C31")
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 32

    # サンプル行（消してOK）
    sample = [
        ("例: 原田 陽一郎", "指揮者", "中隊長"),
        ("例: 梅村 侑志", "指揮者", "小隊長"),
        ("例: 鍋谷 昇", "情報員", ""),
        ("例: 尾坂 友梨", "通信担当", ""),
        ("例: 和田 浩司", "伝令", ""),
        ("例: 金子 卓磨", "機関員", ""),
    ]
    for i, (name, role, note) in enumerate(sample, start=2):
        ws.cell(row=i, column=2, value=name)
        ws.cell(row=i, column=3, value=role)
        ws.cell(row=i, column=4, value=note)


def build_daily_input(ws):
    """当日入力シート: 日付、食当、休暇、研修、訓練期間など。"""
    ws.title = "当日入力"

    ws["A1"] = "当日入力"
    ws["A1"].font = Font(bold=True, size=14)

    rows = [
        ("当番日 (yyyy/m/d)", "2026/4/16"),
        ("曜日", "木"),
        ("部", "3部"),
        ("当直主任", "氏名"),
        ("当直副主任", "氏名"),
        ("休日フラグ (1=休日/0=平日)", 0),
    ]
    r = 3
    for label, val in rows:
        ws.cell(row=r, column=1, value=label).font = FONT_BOLD
        ws.cell(row=r, column=1).border = BORDER_ALL
        c = ws.cell(row=r, column=2, value=val)
        style_input(c)
        if label.startswith("当番日"):
            c.number_format = "yyyy/m/d"
        r += 1

    # リスト入力（食当、休暇、研修、救助訓練期間）
    ws.cell(row=r + 1, column=1, value="下記は氏名を縦に記入。名簿と完全一致させること。").font = Font(italic=True)

    categories = [
        ("食当担当者 (14-17時×)", "D"),
        ("休暇者", "F"),
        ("研修・出向者", "H"),
        ("救助隊訓練期間フラグ者 (10-17時×, 17-19時優先)", "J"),
        ("警防力(毎日勤務)", "L"),
    ]
    for label, col in categories:
        header_cell = ws.cell(row=r + 2, column=openpyxl_col(col), value=label)
        style_header(header_cell)
        ws.merge_cells(start_row=r + 2, start_column=openpyxl_col(col),
                       end_row=r + 2, end_column=openpyxl_col(col) + 1)
        for i in range(8):
            cell = ws.cell(row=r + 3 + i, column=openpyxl_col(col))
            style_input(cell)
            ws.cell(row=r + 3 + i, column=openpyxl_col(col) + 1).border = BORDER_ALL

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    for col in "DEFGHIJKLM":
        ws.column_dimensions[col].width = 14


def openpyxl_col(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def build_history(ws):
    """履歴シート: 過去の当番日の一次指定者を累積保存 (長形式)"""
    ws.title = "履歴"
    ws["A1"] = "過去当番の一次指定者 (GenerateShift 実行時に自動追記される)"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = ("月初に前月分をここに貼り付ける。"
                "列構成: 当番日 / 時間帯 / 通信指令 / 受付。"
                "1当番につき25行。")
    ws["A2"].font = Font(italic=True, size=10)

    headers = ["当番日", "時間帯", "通信指令", "受付"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        style_header(c)

    # 日付列は日付フォーマット
    for r in range(5, 5 + 400):  # 400行（約16当番分）のプレ確保
        ws.cell(row=r, column=1).number_format = "yyyy/m/d"
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).border = BORDER_ALL

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    # 上部をフリーズ
    ws.freeze_panes = "A5"


def build_settings(ws):
    """設定シート: 時間枠ごとの禁止ルール。役職×時間枠で × を付ける。"""
    ws.title = "設定"
    ws["A1"] = "時間枠 × 役職 の割当可否設定 ( × = 割当禁止 / 空白 = 可 / △ = 可能なら回避 )"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "この表を変えれば制約ルールを調整できる。"

    headers = ["時間帯"] + ROLE_CATEGORIES
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        style_header(c)

    # 画像と規則文から起こしたデフォルト禁止ルール
    # キー: 役職カテゴリ, 値: × を付ける時間帯 (substring match)
    defaults = {
        "指揮者": [
            ("8:40〜9", "残留"), ("9〜10", "指揮者"),
            ("8〜8:40", "指揮者"),
        ],
        "情報員": [
            ("10〜11", "情報員"),
        ],
        "通信担当": [
            ("10〜11", "通担/伝令"),
        ],
        "伝令": [
            ("10〜11", "通担/伝令"),
        ],
        # 食当担当者 ( 当日入力で指定 ) は 14-17時 × → VBA側で動的適用
        # 当直は 18-19, 6-8 × → 下に一般役職で ×
        "機関員": [],  # 受付残留なので基本自由、実運用で調整
        "一般": [],
        "警防力": [
            ("8:40〜9", "×"), ("9〜10", "×"), ("10〜11", "×"), ("11〜12", "×"),
            ("12〜13", "×"), ("13〜14", "×"), ("14〜15", "×"), ("15〜16", "×"),
            ("16〜17", "×"), ("17〜18", "×"),
            # 18-22 は優先 △
            ("0〜1", "深夜×"), ("1〜2", "深夜×"), ("2〜3", "深夜×"), ("3〜4", "深夜×"),
            ("4〜5", "深夜×"), ("5〜6", "深夜×"),
            ("6〜7", "×"), ("7〜8", "×"), ("8〜8:40", "×"),
        ],
        "救助隊": [
            # 訓練期間フラグで 10-17 × → VBAで動的適用。ここは空。
        ],
    }

    # 全員に効く共通禁止 (食当, 当直, 救急→今回対象外) は一覧化のみ
    # 時間帯行を埋める
    for i, slot in enumerate(TIME_SLOTS, start=5):
        ws.cell(row=i, column=1, value=slot).border = BORDER_ALL
        ws.cell(row=i, column=1).alignment = ALIGN_CENTER
        for j, role in enumerate(ROLE_CATEGORIES, start=2):
            cell = ws.cell(row=i, column=j)
            cell.border = BORDER_ALL
            cell.alignment = ALIGN_CENTER
            marks = defaults.get(role, [])
            for s, mark in marks:
                if s == slot:
                    cell.value = mark
                    if "×" in mark:
                        cell.fill = FILL_BLOCK

    # 備考: 下部に規則サマリ
    r = 5 + len(TIME_SLOTS) + 2
    rules = [
        "■ 制約ルールまとめ（VBAで処理）",
        "・食当担当者: 14-17時×（当日入力シートで指定）",
        "・当直士長: 18-19時, 6-8時×, 22-23時△",
        "・救助隊訓練期間: 10-17時×, 17-19時優先",
        "・警防力(毎日勤務): 日中×(休日フラグ=1なら緩和), 18-22時優先, 深夜×",
        "・10-17時は一人最低1回勤務（R・L隊員ペアで）",
        "・12〜13時勤と17〜18時勤は別人",
        "・22-4時勤は深夜なので一人1回まで",
        "・前日の一つ上にずらす（基本ローテ）",
        "・夜20時以降は前日と同じ人にしない（同時間帯重複防止）",
    ]
    for line in rules:
        ws.cell(row=r, column=1, value=line)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(ROLE_CATEGORIES) + 1)
        r += 1

    ws.column_dimensions["A"].width = 12
    for j in range(2, len(ROLE_CATEGORIES) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 11


def build_output(ws):
    """執務表シート: 画像と同じ様式。一次指定者を VBA で埋める。"""
    ws.title = "執務表"

    ws["A1"] = "執　務　表"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:G1")
    ws["A1"].alignment = ALIGN_CENTER

    # 表示日付 (履歴再表示用、生成時はマクロが書き込む)
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

    # 2段ヘッダ: 事務区分 / 通信指令 (一次指定者 | 指定変更[従事者|時間]) / 受付 (〃)
    # 「指定変更」列は手書き用で空欄。
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
            cell = ws.cell(row=i, column=col)
            style_output(cell)
        for col in (3, 4, 6, 7):
            cell = ws.cell(row=i, column=col)
            cell.border = BORDER_ALL

    # 二次指定者行
    last = 5 + len(TIME_SLOTS)
    ws.cell(row=last, column=1, value="二次指定者").font = FONT_BOLD
    ws.cell(row=last, column=1).border = BORDER_ALL
    for col in range(2, 8):
        ws.cell(row=last, column=col).border = BORDER_ALL

    # 備考欄
    ws.cell(row=last + 1, column=1, value="備考").font = FONT_BOLD
    ws.merge_cells(start_row=last + 1, start_column=2,
                   end_row=last + 4, end_column=7)
    ws.cell(row=last + 1, column=2,
            value="・夜20時以降は前日と同じ時間帯で同じ人を避ける\n・10-17時は一人1回以上\n・12勤と17勤は別人\n・深夜(22-4時)は一人1回").alignment = Alignment(wrap_text=True, vertical="top")
    for rr in range(last + 1, last + 5):
        for cc in range(1, 8):
            ws.cell(row=rr, column=cc).border = BORDER_ALL

    # 列幅
    ws.column_dimensions["A"].width = 10
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 14

    # 注意書き
    ws["A" + str(last + 6)] = "↑ 一次指定者は VBA マクロ「GenerateShift」で自動生成。指定変更欄は手書き用（出場時の代打記録）。"


def build_help(ws):
    ws.title = "使い方"
    lines = [
        "【執務表自動化 使い方】",
        "",
        "◆ 月初の準備 (月1回だけ)",
        " 1. 履歴シートに前月分の一次指定者をコピペ (当番日×25行ずつ縦に並ぶ)",
        "    - 前月末の当番日が残っていれば、最初の当番でも自動でローテできる",
        "",
        "◆ 毎当番の作業 (ボタン1回)",
        " 1. 当日入力シート",
        "    - 当番日を yyyy/m/d 形式で更新",
        "    - 食当・休暇・研修・救助訓練・警防力の氏名を更新",
        "    - 当直主任/副主任を更新",
        " 2. 執務表シートのボタン「本日を生成」を押す",
        "    → 履歴の最新日を「前日」として読み込み、自動で一次指定者が埋まる",
        "    → 生成結果は履歴に自動追記される",
        " 3. 指定変更欄は空欄のまま。出場時の代打を手書き",
        " 4. 印刷 or Word様式にコピペ",
        "",
        "◆ 過去の日を見たい場合",
        " - 執務表シートの「表示日付 (B2)」に yyyy/m/d を書いて",
        "   ボタン「表示日付を反映」を押す (マクロ ShowDateFromHistory)",
        "",
        "◆ 名簿の更新",
        " 1. 名簿シートに3部員の氏名と役職カテゴリを入力 (最大30名)",
        "   - 役職カテゴリは C列のドロップダウンから選択",
        "   - 指揮者/情報員/伝令/通信担当/機関員/一般/警防力/救助隊 の8区分",
        "",
        "◆ 制約を変えたい場合",
        " - 禁止時間帯 → 設定シートのマスを × / △ / 空白 で編集",
        " - 食当・救助訓練期間・警防力 → 当日入力シートで変える",
        "",
        "■ 初回セットアップ (VBAマクロ導入)",
        " a) このファイルを「名前を付けて保存」→ Excel マクロ有効ブック(.xlsm)",
        " b) Alt+F11 で VBE を開く",
        " c) [ファイル]→[ファイルのインポート] で ShiftScheduler.bas を選ぶ",
        " d) 執務表シートに [開発]タブ→[挿入]→[ボタン] で図形を2つ置き",
        "    - ボタン1: マクロ「GenerateShift」",
        "    - ボタン2: マクロ「ShowDateFromHistory」",
        "",
        "◆ 履歴シートの扱い",
        " - 自動追記なので基本触らない",
        " - 月初に前月分を貼り付ける (列: 当番日/時間帯/通信指令/受付、1当番=25行)",
        " - 間違えて追記された日があれば、該当25行を削除すればOK",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 90


def main():
    wb = Workbook()
    # 最初のシートを 執務表 にする (起動時にここを見せる)
    ws_output = wb.active
    build_output(ws_output)

    ws_input = wb.create_sheet()
    build_daily_input(ws_input)

    ws_roster = wb.create_sheet()
    build_roster(ws_roster)

    ws_history = wb.create_sheet()
    build_history(ws_history)

    ws_settings = wb.create_sheet()
    build_settings(ws_settings)

    ws_help = wb.create_sheet()
    build_help(ws_help)

    # シート順: 執務表 / 当日入力 / 名簿 / 履歴 / 設定 / 使い方
    wb._sheets = [ws_output, ws_input, ws_roster, ws_history, ws_settings, ws_help]

    wb.save(OUT_PATH)
    print(f"生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
