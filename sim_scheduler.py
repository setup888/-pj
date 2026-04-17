"""
VBA ShiftScheduler.bas のロジックを Python に移植したドライランシミュレータ
(完全データ駆動版)。

ポジション×時間帯のブロックセット + 当日チェックの「ポジション+除外時間帯」
でメンバー可否を決定する。
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Tuple
from collections import defaultdict

TIME_SLOTS = [
    "8:40〜9", "9〜10", "10〜11", "11〜12", "12〜13", "13〜14",
    "14〜15", "15〜16", "16〜17", "17〜18", "18〜19", "19〜20",
    "20〜21", "21〜22", "22〜23", "23〜24",
    "0〜1", "1〜2", "2〜3", "3〜4", "4〜5", "5〜6",
    "6〜7", "7〜8", "8〜8:40",
]
N = len(TIME_SLOTS)
SLOT_MAP = {s: i for i, s in enumerate(TIME_SLOTS)}

S_10_11 = 2
S_12_13 = 4
S_17_18 = 9
S_18_19 = 10
S_20_21 = 12
S_22_23 = 14
S_4_5 = 20


# ポジション定義 (userが編集するもの。初期値と同じ)
POSITIONS: Dict[str, set] = {
    "ポンプ隊": set(),
    "救助隊": set(),
    "はしご隊": set(),
    "救急隊": set(range(N)),       # 全時間×
    "伝令": set(),
    "通信担当": set(),
    "情報担当": set(),
    "情報員": set(),
    "署隊長伝令": set(),
    "残留": set(),
    "署隊本部支援員": set(),
    "その他": set(),
    "当直主任": {10, 11, 22, 23},    # 18-19, 19-20, 6-7, 7-8
    "当直副主任": {10, 11, 22, 23},
    "食当": {6, 7, 8},               # 14-17
    "休暇": set(range(N)),
    "研修/出向": set(range(N)),
}


@dataclass
class DailyInput:
    date: date
    休日: int = 0
    positions: Dict[str, str] = field(default_factory=dict)   # name -> position
    exclusions: Dict[str, List[Tuple[int, int]]] = field(default_factory=dict)


def is_blocked(name: str, slot: int, daily: DailyInput) -> bool:
    pos = daily.positions.get(name)
    if pos and pos in POSITIONS and slot in POSITIONS[pos]:
        return True
    for s, e in daily.exclusions.get(name, []):
        if s <= slot <= e:
            return True
    return False


def is_late_night(slot): return S_22_23 <= slot <= S_4_5 - 1


def daytime_count(name, assign, cur):
    cnt = 0
    for i in range(S_10_11, S_17_18):
        if i == cur: break
        for c in (0, 1):
            if assign[c][i] == name: cnt += 1
    return cnt


def late_count(name, assign, cur):
    cnt = 0
    for i in range(S_22_23, S_4_5):
        if i == cur: break
        for c in (0, 1):
            if assign[c][i] == name: cnt += 1
    return cnt


def score(name, col, slot, prev_day, assign, work_count):
    s = work_count.get(name, 0) * 10
    prev_name = prev_day.get(slot, ("", ""))[col]
    if prev_name and prev_name == name:
        s += 1000 if slot >= S_20_21 else 50
    if slot + 1 <= N - 1:
        if prev_day.get(slot + 1, ("", ""))[col] == name:
            s -= 5
    if S_10_11 <= slot <= S_17_18 - 1:
        if daytime_count(name, assign, slot) == 0:
            s -= 200
    if is_late_night(slot):
        if late_count(name, assign, slot) >= 1:
            s += 500
    if slot == S_17_18 and assign[col][S_12_13] == name: s += 500
    if slot == S_12_13 and assign[col][S_17_18] == name: s += 500
    if slot > 0 and assign[col][slot - 1] == name: s += 50
    if slot < N - 1 and assign[col][slot + 1] == name: s += 50
    return s


def pick(col, slot, members, daily, prev_day, assign, work_count):
    cands = [n for n in members if not is_blocked(n, slot, daily)]
    if col == 1:
        cands = [n for n in cands if assign[0][slot] != n]
    if not cands: return ""
    best_n, best_s = "", float("inf")
    for n in cands:
        sc = score(n, col, slot, prev_day, assign, work_count)
        if sc < best_s:
            best_s, best_n = sc, n
    if best_n:
        work_count[best_n] = work_count.get(best_n, 0) + 1
    return best_n


def slot_order():
    return (list(range(S_10_11, S_17_18)) + list(range(9, 16))
            + list(range(16, 24)) + [24, 0, 1])


def all_blocked_daytime(name, daily):
    return all(is_blocked(name, i, daily) for i in range(S_10_11, S_17_18))


def generate(roster, daily, prev_day):
    members = [n for n in roster
               if any(not is_blocked(n, i, daily) for i in range(N))]
    assign = [[""] * N, [""] * N]
    work_count = {m: 0 for m in members}
    for slot in slot_order():
        for col in (0, 1):
            assign[col][slot] = pick(col, slot, members, daily, prev_day, assign, work_count)
    # 12/17 distinct
    for col in (0, 1):
        if assign[col][S_12_13] and assign[col][S_12_13] == assign[col][S_17_18]:
            for alt in members:
                if alt != assign[col][S_12_13] \
                        and not is_blocked(alt, S_17_18, daily) \
                        and assign[1-col][S_17_18] != alt:
                    assign[col][S_17_18] = alt
                    break
    return assign


def validate(assign, roster, daily, prev_day):
    msgs = []
    for i in range(N):
        for col in (0, 1):
            if not assign[col][i]:
                msgs.append(f"{TIME_SLOTS[i]}/{'通信' if col == 0 else '受付'} 空欄")
    covered = defaultdict(int)
    for i in range(S_10_11, S_17_18):
        for col in (0, 1):
            if assign[col][i]: covered[assign[col][i]] += 1
    for n in roster:
        if covered[n] == 0 and not all_blocked_daytime(n, daily):
            msgs.append(f"{n} は 10-17時未勤務")
    for i in range(S_20_21, N):
        for col in (0, 1):
            if assign[col][i] and prev_day.get(i, ("", ""))[col] == assign[col][i]:
                msgs.append(f"{TIME_SLOTS[i]}/{'通信' if col == 0 else '受付'} 前日と同じ ({assign[col][i]})")
    return msgs


def print_assign(assign, header=""):
    print(f"\n=== {header} ===")
    print(f"{'時間帯':<10} | {'通信':<14} | {'受付':<14}")
    print("-" * 50)
    for i, s in enumerate(TIME_SLOTS):
        print(f"{s:<10} | {assign[0][i]:<14} | {assign[1][i]:<14}")


def main():
    roster = [
        "原田 陽一郎", "梅村 侑志", "小西 隼人", "鍋谷 昇", "村山 哲也",
        "長田 智紀", "和田 浩司", "長友 亮澄", "尾坂 友梨", "山川 敦史",
        "金子 卓磨", "永井 恵理", "中村 太一", "藤井 惇平", "伊藤 祥輝",
        "飯塚 佑介", "後藤 直人",
    ]

    # 1日目 ポジション配置 (警防態勢)
    daily1 = DailyInput(
        date=date(2026, 4, 16),
        positions={
            "原田 陽一郎": "当直主任",
            "永井 恵理": "当直副主任",
            "梅村 侑志": "食当",
            "村山 哲也": "食当",
            "小西 隼人": "ポンプ隊",
            "鍋谷 昇": "はしご隊",
            "長田 智紀": "伝令",
            "和田 浩司": "伝令",
            "長友 亮澄": "情報員",
            "尾坂 友梨": "通信担当",
            "山川 敦史": "情報担当",
            "金子 卓磨": "残留",
            "中村 太一": "救助隊",
            "藤井 惇平": "ポンプ隊",
            "伊藤 祥輝": "ポンプ隊",
            "飯塚 佑介": "その他",
            "後藤 直人": "署隊長伝令",
        },
        exclusions={
            "鍋谷 昇": [(SLOT_MAP["9〜10"], SLOT_MAP["12〜13"])],  # 方面訓練
        },
    )
    prev_day = {i: ("", "") for i in range(N)}
    assign1 = generate(roster, daily1, prev_day)
    print_assign(assign1, f"Day 1 ({daily1.date})")
    msgs = validate(assign1, roster, daily1, prev_day)
    if msgs:
        print("\n⚠ 警告:")
        for m in msgs: print("  -", m)
    else:
        print("\n✓ 制約違反なし")

    # 2日目 (ポジション変わる)
    prev_day2 = {i: (assign1[0][i], assign1[1][i]) for i in range(N)}
    daily2 = DailyInput(
        date=date(2026, 4, 19),
        positions={
            "梅村 侑志": "当直主任",
            "永井 恵理": "当直副主任",
            "小西 隼人": "食当",
            "原田 陽一郎": "ポンプ隊",
            "鍋谷 昇": "伝令",
            "村山 哲也": "通信担当",
            "長田 智紀": "情報担当",
            "和田 浩司": "情報員",
            "長友 亮澄": "残留",
            "尾坂 友梨": "はしご隊",
            "山川 敦史": "ポンプ隊",
            "金子 卓磨": "署隊長伝令",
            "中村 太一": "休暇",
            "藤井 惇平": "救助隊",
            "伊藤 祥輝": "救急隊",
            "飯塚 佑介": "その他",
            "後藤 直人": "署隊本部支援員",
        },
    )
    assign2 = generate(roster, daily2, prev_day2)
    print_assign(assign2, f"Day 2 ({daily2.date})")
    msgs = validate(assign2, roster, daily2, prev_day2)
    if msgs:
        print("\n⚠ 警告:")
        for m in msgs: print("  -", m)
    else:
        print("\n✓ 制約違反なし")

    # 夜20時以降比較
    print("\n\n=== 夜20時以降 D1→D2 通信比較 ===")
    dup = 0
    for i in range(S_20_21, N):
        d1, d2 = assign1[0][i], assign2[0][i]
        mark = "×" if d1 and d1 == d2 else ""
        if mark: dup += 1
        print(f"{TIME_SLOTS[i]:<10} | {d1:<14} | {d2:<14} | {mark}")
    print(f"夜20時以降の重複: {dup} 件")


if __name__ == "__main__":
    main()
