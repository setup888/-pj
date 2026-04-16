"""
VBA ShiftScheduler.bas のロジックを Python に移植したドライランシミュレータ。
サンプルデータで生成し、制約違反・偏りを確認する。

使い方:
    python3 sim_scheduler.py
"""

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, List, Tuple, Optional
from collections import defaultdict
import random

TIME_SLOTS = [
    "8:40〜9", "9〜10", "10〜11", "11〜12", "12〜13", "13〜14",
    "14〜15", "15〜16", "16〜17", "17〜18", "18〜19", "19〜20",
    "20〜21", "21〜22", "22〜23", "23〜24",
    "0〜1", "1〜2", "2〜3", "3〜4", "4〜5", "5〜6",
    "6〜7", "7〜8", "8〜8:40",
]
N = len(TIME_SLOTS)  # 25

S_8_9, S_9_10, S_10_11 = 0, 1, 2
S_12_13 = 4
S_14_15 = 6
S_17_18 = 9
S_18_19 = 10
S_20_21 = 12
S_22_23 = 14
S_0_1 = 16
S_4_5 = 20
S_6_7 = 22
S_8_840 = 24


@dataclass
class DailyInput:
    date: date
    休日: int = 0
    休暇: set = field(default_factory=set)
    食当: set = field(default_factory=set)
    当直主任: str = ""
    当直副主任: str = ""
    除外: Dict[str, List[Tuple[int, int]]] = field(default_factory=dict)  # name -> [(startIdx, endIdx)]


def is_blocked(name: str, slot: int, roster: Dict[str, str],
               daily: DailyInput) -> bool:
    role = roster.get(name, "")

    # 2) 休暇 → 全時間 ×
    if name in daily.休暇:
        return True

    # 3) 食当 → 14-17時 ×
    if name in daily.食当:
        if S_14_15 <= slot <= S_14_15 + 2:  # 14,15,16
            return True

    # 4) 警防力 → 日中×/深夜×
    if role == "警防力":
        if daily.休日 != 1:
            if S_8_9 <= slot <= S_17_18:
                return True
        if S_0_1 <= slot <= S_6_7 - 1:
            return True

    # 5) 当直主任/副主任 は 18-19, 6-8 ×
    if name in (daily.当直主任, daily.当直副主任) and name:
        if slot in (S_18_19, S_18_19 + 1, S_6_7, S_6_7 + 1):
            return True

    # 6) 除外要件
    for s, e in daily.除外.get(name, []):
        if s <= slot <= e:
            return True

    return False


def is_late_night(slot: int) -> bool:
    return S_22_23 <= slot <= S_4_5 - 1


def daytime_count(name: str, assign, current_slot: int) -> int:
    cnt = 0
    for i in range(S_10_11, S_17_18):
        if i >= current_slot:
            break
        for c in (0, 1):
            if assign[c][i] == name:
                cnt += 1
    return cnt


def late_night_count(name: str, assign, current_slot: int) -> int:
    cnt = 0
    for i in range(S_22_23, S_4_5):
        if i >= current_slot:
            break
        for c in (0, 1):
            if assign[c][i] == name:
                cnt += 1
    return cnt


def score_candidate(name: str, col: int, slot: int, roster: Dict[str, str],
                    daily: DailyInput, prev_day: Dict[int, Tuple[str, str]],
                    assign, work_count: Dict[str, int]) -> float:
    score = 0.0

    # (a)
    score += work_count.get(name, 0) * 10

    # (b) 前日同時刻
    prev_name = prev_day.get(slot, ("", ""))[col]
    if prev_name and prev_name == name:
        if slot >= S_20_21:
            score += 1000
        else:
            score += 50

    # (c) 前日 slot+1 を軽く優先
    if slot + 1 <= N - 1:
        prev_next = prev_day.get(slot + 1, ("", ""))[col]
        if prev_next == name:
            score -= 5

    # (d) 警防力 18-22時優先
    if roster.get(name) == "警防力":
        if S_18_19 <= slot <= S_20_21 + 1:
            score -= 20

    # (f) 10-17時未勤務者を強優先
    if S_10_11 <= slot <= S_17_18 - 1:
        if daytime_count(name, assign, slot) == 0:
            score -= 200

    # (g) 深夜1回超え
    if is_late_night(slot):
        if late_night_count(name, assign, slot) >= 1:
            score += 500

    # (h) 12勤/17勤 別人
    if slot == S_17_18 and assign[col][S_12_13] == name:
        score += 500
    if slot == S_12_13 and assign[col][S_17_18] == name:
        score += 500

    # (i) 直前連続ペナルティ
    if slot > 0 and assign[col][slot - 1] == name:
        score += 3

    return score


def pick_assignee(col: int, slot: int, members: List[str], roster, daily,
                  prev_day, assign, work_count) -> str:
    candidates = []
    for name in members:
        if is_blocked(name, slot, roster, daily):
            continue
        if col == 1 and assign[0][slot] == name:
            continue
        candidates.append(name)

    if not candidates:
        return ""

    best_name, best_score = "", float("inf")
    for c in candidates:
        s = score_candidate(c, col, slot, roster, daily, prev_day, assign, work_count)
        if s < best_score:
            best_score = s
            best_name = c
    if best_name:
        work_count[best_name] = work_count.get(best_name, 0) + 1
    return best_name


def enforce_distinct_12_17(assign, members, roster, daily):
    for col in (0, 1):
        if assign[col][S_12_13] and assign[col][S_12_13] == assign[col][S_17_18]:
            for alt in members:
                if alt == assign[col][S_12_13]:
                    continue
                if is_blocked(alt, S_17_18, roster, daily):
                    continue
                if assign[1 - col][S_17_18] == alt:
                    continue
                assign[col][S_17_18] = alt
                break


def build_slot_order():
    # daytime 10-17 を最優先、次に 17-24, 0-8, 8-8:40, 最後に 8:40-10
    return (
        list(range(S_10_11, S_17_18))    # 2..8
        + list(range(9, 16))              # 9..15 (17-24)
        + list(range(16, 24))             # 16..23 (0-8)
        + [24]                            # 8〜8:40
        + [0, 1]                          # 8:40〜9, 9〜10
    )


def generate(roster: Dict[str, str], daily: DailyInput,
             prev_day: Dict[int, Tuple[str, str]]) -> List[List[str]]:
    members = [n for n in roster if n not in daily.休暇]
    assign = [[""] * N, [""] * N]
    work_count = {m: 0 for m in members}
    for slot in build_slot_order():
        for col in (0, 1):
            assign[col][slot] = pick_assignee(col, slot, members, roster, daily,
                                              prev_day, assign, work_count)
    enforce_distinct_12_17(assign, members, roster, daily)
    return assign


def validate(assign, roster, daily, prev_day) -> List[str]:
    msgs = []
    for i in range(N):
        for col in (0, 1):
            if not assign[col][i]:
                msgs.append(f"{TIME_SLOTS[i]}/{'通信' if col == 0 else '受付'} 空欄")
    # 10-17カバー
    covered = defaultdict(int)
    for i in range(S_10_11, S_17_18):
        for col in (0, 1):
            if assign[col][i]:
                covered[assign[col][i]] += 1
    for k, role in roster.items():
        if k in daily.休暇 or role == "警防力":
            continue
        if covered[k] == 0:
            msgs.append(f"{k} は 10-17時に未勤務")
    # 夜20以降 前日同一
    for i in range(S_20_21, N):
        for col in (0, 1):
            if assign[col][i] and prev_day.get(i, ("", ""))[col] == assign[col][i]:
                msgs.append(f"{TIME_SLOTS[i]}/{'通信' if col == 0 else '受付'} 前日と同じ ({assign[col][i]})")
    return msgs


def print_assign(assign, header=""):
    print(f"\n=== {header} ===")
    print(f"{'時間帯':<10} | {'通信':<12} | {'受付':<12}")
    print("-" * 46)
    for i, slot in enumerate(TIME_SLOTS):
        print(f"{slot:<10} | {assign[0][i]:<12} | {assign[1][i]:<12}")


def count_assignments(assign) -> Dict[str, int]:
    cnt = defaultdict(int)
    for col in (0, 1):
        for name in assign[col]:
            if name:
                cnt[name] += 1
    return cnt


def main():
    # サンプル名簿 17名
    roster = {
        "原田 陽一郎": "指揮者",
        "梅村 侑志": "指揮者",
        "小西 隼人": "指揮者",
        "鍋谷 昇": "情報員",
        "村山 哲也": "情報員",
        "長田 智紀": "情報員",
        "和田 浩司": "伝令",
        "長友 亮澄": "伝令",
        "尾坂 友梨": "通信担当",
        "山川 敦史": "通信担当",
        "金子 卓磨": "機関員",
        "永井 恵理": "機関員",
        "中村 太一": "一般",
        "藤井 惇平": "一般",
        "伊藤 祥輝": "一般",
        "飯塚 佑介": "一般",
        "後藤 直人": "一般",
    }

    # 前日実績なし、初回ドライラン
    prev_day = {i: ("", "") for i in range(N)}

    # 1日目
    daily1 = DailyInput(
        date=date(2026, 4, 16),
        休日=0,
        食当={"梅村 侑志", "村山 哲也"},
        当直主任="原田 陽一郎",
        当直副主任="永井 恵理",
        除外={
            "鍋谷 昇": [(S_9_10, S_12_13)],  # 方面訓練 9-13
        },
    )
    print("【1日目】", daily1.date)
    print("食当:", daily1.食当, "休暇:", daily1.休暇)
    print("当直:", daily1.当直主任, "/", daily1.当直副主任)
    print("除外:", daily1.除外)
    assign1 = generate(roster, daily1, prev_day)
    print_assign(assign1, f"Day 1 ({daily1.date})")

    cnt1 = count_assignments(assign1)
    print("\n勤務回数 (day 1):")
    for name in sorted(roster):
        print(f"  {name:<15}: {cnt1.get(name, 0)}")

    msgs1 = validate(assign1, roster, daily1, prev_day)
    if msgs1:
        print("\n⚠ 警告:")
        for m in msgs1:
            print(" -", m)
    else:
        print("\n✓ 制約違反なし")

    # 2日目 (前日=1日目)
    prev_day2 = {i: (assign1[0][i], assign1[1][i]) for i in range(N)}
    daily2 = DailyInput(
        date=date(2026, 4, 19),
        休日=0,
        食当={"小西 隼人"},
        休暇={"中村 太一"},
        当直主任="梅村 侑志",
        当直副主任="永井 恵理",
        除外={
            "藤井 惇平": [(S_10_11, S_17_18 - 1)],  # 救助訓練
        },
    )
    print("\n\n【2日目】", daily2.date)
    assign2 = generate(roster, daily2, prev_day2)
    print_assign(assign2, f"Day 2 ({daily2.date})")

    cnt2 = count_assignments(assign2)
    print("\n勤務回数 (day 2):")
    for name in sorted(roster):
        print(f"  {name:<15}: {cnt2.get(name, 0)}")

    msgs2 = validate(assign2, roster, daily2, prev_day2)
    if msgs2:
        print("\n⚠ 警告:")
        for m in msgs2:
            print(" -", m)
    else:
        print("\n✓ 制約違反なし")

    # 夜20時以降の前日重複チェック (手動)
    print("\n\n=== 夜20時以降 前日比較 ===")
    print(f"{'時間帯':<10} | {'D1通信':<12} | {'D2通信':<12} | 重複")
    print("-" * 60)
    dup = 0
    for i in range(S_20_21, N):
        d1_c = assign1[0][i]
        d2_c = assign2[0][i]
        same = "×" if d1_c and d1_c == d2_c else ""
        if same: dup += 1
        print(f"{TIME_SLOTS[i]:<10} | {d1_c:<12} | {d2_c:<12} | {same}")
    print(f"夜20時以降の重複: {dup} 件")


if __name__ == "__main__":
    main()
