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

# 時刻境界 (26)
TIME_MARKER_MAP = {
    "8:40": 0, "9時": 1, "10時": 2, "11時": 3, "12時": 4, "13時": 5,
    "14時": 6, "15時": 7, "16時": 8, "17時": 9, "18時": 10, "19時": 11,
    "20時": 12, "21時": 13, "22時": 14, "23時": 15, "0時": 16, "1時": 17,
    "2時": 18, "3時": 19, "4時": 20, "5時": 21, "6時": 22, "7時": 23,
    "8時": 24, "8:40(翌)": 25,
}


def marker_range(from_label: str, to_label: str) -> Tuple[int, int]:
    """時刻境界ラベルから スロット範囲 [start, end] を返す。
    例: ("9時", "17時") → (1, 8) → slot 9〜10 から 16〜17 をブロック
    """
    s = TIME_MARKER_MAP[from_label]
    e = TIME_MARKER_MAP[to_label]
    if e < s:
        s, e = e, s
    return (s, e - 1)

S_10_11 = 2
S_12_13 = 4
S_14_15 = 6
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
    "救急隊": set(range(N)),
    "日中救急": set(range(0, 10)),
    "夜救急": set(range(10, N)),
    "伝令": set(),
    "通信担当": set(),
    "情報担当": set(),
    "情報員": set(),
    "署隊長伝令": set(),
    "残留": set(),
    "署隊本部支援員": set(),
    "その他": set(),
    "当直": {10, 11, 22, 23},
    "食当": {6, 7, 8},
    "休暇": set(range(N)),
    "研修/出向": set(range(N)),
    "警防力": set(range(N)),
}

# カラム専属: "comm" = 通信のみ, "recv" = 受付のみ
POSITION_COL_LOCK: Dict[str, str] = {
    "残留": "comm",
    "署隊長伝令": "recv",
}


RANK_VALUE = {"司令補": 4, "士長": 3, "副士長": 2, "消防士": 1}


@dataclass
class DailyInput:
    date: date
    休日: int = 0
    positions: Dict[str, List[str]] = field(default_factory=dict)
    exclusions: Dict[str, List[Tuple[int, int]]] = field(default_factory=dict)


def is_blocked(name: str, slot: int, daily: DailyInput, col: int = -1) -> bool:
    # 複数ポジションの × を合算
    for pos in daily.positions.get(name, []):
        if pos in POSITIONS and slot in POSITIONS[pos]:
            return True

    # カラム専属チェック
    if col >= 0:
        for pos in daily.positions.get(name, []):
            lock = POSITION_COL_LOCK.get(pos)
            if lock == "comm" and col == 1:
                return True
            if lock == "recv" and col == 0:
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


def min_gap_same_col(name, col, slot, assign):
    min_g = 999
    for i in range(N):
        if i == slot: continue
        if assign[col][i] == name:
            d = abs(i - slot)
            if d < min_g: min_g = d
    return min_g


def is_exclusive_for_col(name, col, daily):
    for pos in daily.positions.get(name, []):
        lock = POSITION_COL_LOCK.get(pos)
        if col == 0 and lock == "comm": return True
        if col == 1 and lock == "recv": return True
    return False


def score(name, col, slot, prev_day, assign, work_count, roster, daily=None):
    s = work_count.get(name, 0) * 10

    # 階級による 通信/受付 優先
    rv = RANK_VALUE.get(roster.get(name, ""), 0)
    if col == 0:  # 通信
        if rv == 4 or rv == 3: s -= 30
        elif rv == 2: s += 30
        elif rv == 1: s += 300
    else:         # 受付
        if rv == 4 or rv == 3: s += 20
        elif rv == 2 or rv == 1: s -= 30

    prev_name = prev_day.get(slot, ("", ""))[col]
    if prev_name and prev_name == name:
        s += 1000 if slot >= S_20_21 else 50
    if slot + 1 <= N - 1:
        if prev_day.get(slot + 1, ("", ""))[col] == name:
            s -= 5
    if S_10_11 <= slot <= S_17_18 - 1:
        if daytime_count(name, assign, slot) == 0:
            s -= 200

    # (f2) 専属カラム保持者ローテ優先
    if daily and is_exclusive_for_col(name, col, daily):
        s -= 40

    # (f3) 食当者の 10-14時優先
    if daily and S_10_11 <= slot <= S_14_15 - 1:
        if "食当" in daily.positions.get(name, []):
            s -= 60

    if is_late_night(slot):
        if late_count(name, assign, slot) >= 1:
            s += 500
    if slot == S_17_18 and assign[col][S_12_13] == name: s += 500
    if slot == S_12_13 and assign[col][S_17_18] == name: s += 500
    # 近接ペナルティ ±4
    for dist, pen in [(1, 200), (2, 100), (3, 50), (4, 30)]:
        if slot - dist >= 0 and assign[col][slot - dist] == name: s += pen
        if slot + dist < N and assign[col][slot + dist] == name: s += pen
    return s


def pick(col, slot, members, daily, prev_day, assign, work_count, roster):
    base = [n for n in members if not is_blocked(n, slot, daily, col)]
    if col == 1:
        base = [n for n in base if assign[0][slot] != n]
    if not base: return ""

    # 2段階候補フィルタ: Strict (gap≥4) → Relaxed (gap≥2) → All
    strict = [n for n in base if min_gap_same_col(n, col, slot, assign) >= 4]
    relaxed = [n for n in base if min_gap_same_col(n, col, slot, assign) >= 2]
    cands = strict if strict else (relaxed if relaxed else base)

    best_n, best_s = "", float("inf")
    for n in cands:
        sc = score(n, col, slot, prev_day, assign, work_count, roster, daily)
        if sc < best_s:
            best_s, best_n = sc, n
    if best_n:
        work_count[best_n] = work_count.get(best_n, 0) + 1
    return best_n


def slot_order():
    return (list(range(S_10_11, S_17_18)) + list(range(9, 16))
            + list(range(16, 24)) + [24, 0, 1])


def all_blocked_daytime(name, daily):
    # 通信でも受付でも入れないなら True
    return all(
        is_blocked(name, i, daily, 0) and is_blocked(name, i, daily, 1)
        for i in range(S_10_11, S_17_18)
    )


def find_member_with_position(members, daily, position_name):
    for n in members:
        if position_name in daily.positions.get(n, []):
            return n
    return None


def apply_fixed_slots(assign, members, daily, work_count):
    """Phase A: 残留 → 通信 slot 0,1 / 署隊長伝令 → 受付 slot 1,24 / 伝令or通信担当 → slot 2"""
    zan = find_member_with_position(members, daily, "残留")
    if zan:
        if not is_blocked(zan, 0, daily, 0):
            assign[0][0] = zan
            work_count[zan] = work_count.get(zan, 0) + 1
        if not is_blocked(zan, 1, daily, 0):
            assign[0][1] = zan
            work_count[zan] = work_count.get(zan, 0) + 1

    de = find_member_with_position(members, daily, "署隊長伝令")
    if de:
        if not is_blocked(de, 1, daily, 1):
            assign[1][1] = de
            work_count[de] = work_count.get(de, 0) + 1
        if not is_blocked(de, 24, daily, 1):
            assign[1][24] = de
            work_count[de] = work_count.get(de, 0) + 1

    den = find_member_with_position(members, daily, "伝令")
    tsu = find_member_with_position(members, daily, "通信担当")
    if den and not is_blocked(den, S_10_11, daily, 0):
        assign[0][S_10_11] = den
        work_count[den] = work_count.get(den, 0) + 1
    elif tsu and not is_blocked(tsu, S_10_11, daily, 0):
        assign[0][S_10_11] = tsu
        work_count[tsu] = work_count.get(tsu, 0) + 1

    recv = None
    if tsu and tsu != assign[0][S_10_11]:
        recv = tsu
    elif den and den != assign[0][S_10_11]:
        recv = den
    if recv and not is_blocked(recv, S_10_11, daily, 1):
        assign[1][S_10_11] = recv
        work_count[recv] = work_count.get(recv, 0) + 1


def apply_night_slide(assign, members, daily, prev_day, roster, work_count):
    """Phase B: slot 14..20 (22時〜5時) = 前日 slot+1 を上に1つスライド"""
    for slot in range(S_22_23, S_4_5 + 1):  # 14..20
        for col in (0, 1):
            if assign[col][slot]:
                continue
            if slot + 1 >= N: continue
            prev_name = prev_day.get(slot + 1, ("", ""))[col]
            if not prev_name: continue
            if prev_name not in roster: continue
            if is_blocked(prev_name, slot, daily, col): continue
            if col == 1 and assign[0][slot] == prev_name: continue
            assign[col][slot] = prev_name
            work_count[prev_name] = work_count.get(prev_name, 0) + 1


def generate(roster, daily, prev_day):
    members = [
        n for n in roster
        if any(
            not is_blocked(n, i, daily, 0) or not is_blocked(n, i, daily, 1)
            for i in range(N)
        )
    ]
    assign = [[""] * N, [""] * N]
    work_count = {m: 0 for m in members}

    # Phase A: 固定枠
    apply_fixed_slots(assign, members, daily, work_count)

    # Phase B: 深夜スライド
    apply_night_slide(assign, members, daily, prev_day, roster, work_count)

    # Phase C: 残りスコアリング
    for slot in slot_order():
        for col in (0, 1):
            if col == 1 and slot == 0:
                assign[1][0] = ""
                continue
            if assign[col][slot]:
                continue  # 固定枠・スライドで埋まってる
            assign[col][slot] = pick(col, slot, members, daily, prev_day,
                                     assign, work_count, roster)

    # 12/17 distinct
    for col in (0, 1):
        if assign[col][S_12_13] and assign[col][S_12_13] == assign[col][S_17_18]:
            for alt in members:
                if alt != assign[col][S_12_13] \
                        and not is_blocked(alt, S_17_18, daily, col) \
                        and assign[1-col][S_17_18] != alt:
                    assign[col][S_17_18] = alt
                    break
    return assign


def validate(assign, roster, daily, prev_day):
    msgs = []
    for i in range(N):
        for col in (0, 1):
            if col == 1 and i == 0:
                continue  # 受付 8:40〜9 は斜線
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
    # 名簿: 氏名 -> 階級
    roster = {
        "原田 陽一郎": "司令補",
        "梅村 侑志": "司令補",
        "小西 隼人": "士長",
        "鍋谷 昇": "士長",
        "村山 哲也": "士長",
        "長田 智紀": "士長",
        "和田 浩司": "副士長",
        "長友 亮澄": "副士長",
        "尾坂 友梨": "副士長",
        "山川 敦史": "副士長",
        "金子 卓磨": "副士長",
        "永井 恵理": "副士長",
        "中村 太一": "消防士",
        "藤井 惇平": "消防士",
        "伊藤 祥輝": "消防士",
        "飯塚 佑介": "消防士",
        "後藤 直人": "消防士",
    }

    # 1日目 ポジション配置 (警防態勢) - 複数ポジション対応
    daily1 = DailyInput(
        date=date(2026, 4, 16),
        positions={
            "原田 陽一郎": ["ポンプ隊"],
            "梅村 侑志": ["救助隊"],
            "小西 隼人": ["ポンプ隊", "当直"],
            "鍋谷 昇": ["はしご隊"],
            "村山 哲也": ["食当"],
            "長田 智紀": ["伝令"],            # 10-11固定候補
            "和田 浩司": ["通信担当"],         # 10-11固定候補
            "長友 亮澄": ["情報員"],
            "尾坂 友梨": ["情報担当"],
            "山川 敦史": ["残留"],            # 通信専属、朝2時間固定
            "金子 卓磨": ["署隊長伝令"],       # 受付専属、朝&翌朝受付固定
            "永井 恵理": ["ポンプ隊"],
            "中村 太一": ["救助隊"],
            "藤井 惇平": ["はしご隊"],
            "伊藤 祥輝": ["夜救急"],
            "飯塚 佑介": ["その他"],
            "後藤 直人": ["署隊本部支援員"],
        },
        exclusions={
            "鍋谷 昇": [marker_range("9時", "13時")],  # 方面訓練
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

    # 2日目 (ポジション・当直がぐるぐる変わる)
    prev_day2 = {i: (assign1[0][i], assign1[1][i]) for i in range(N)}
    daily2 = DailyInput(
        date=date(2026, 4, 19),
        positions={
            "梅村 侑志": ["ポンプ隊", "当直"],             # 当直に切替
            "永井 恵理": ["残留"],
            "小西 隼人": ["食当", "救助隊"],               # 食当+救助
            "原田 陽一郎": ["ポンプ隊"],
            "鍋谷 昇": ["伝令"],
            "村山 哲也": ["通信担当"],
            "長田 智紀": ["情報担当"],
            "和田 浩司": ["情報員"],
            "長友 亮澄": ["残留"],
            "尾坂 友梨": ["はしご隊"],
            "山川 敦史": ["ポンプ隊"],
            "金子 卓磨": ["署隊長伝令"],
            "中村 太一": ["休暇"],                         # 休暇チェック
            "藤井 惇平": ["救助隊"],
            "伊藤 祥輝": ["救急隊"],                       # 全日救急
            "飯塚 佑介": ["その他"],
            "後藤 直人": ["署隊本部支援員"],
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
