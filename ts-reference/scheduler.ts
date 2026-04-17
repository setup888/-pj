/**
 * 執務表 一次指定者 自動生成アルゴリズム
 *
 * 純関数（React hooks 等依存なし）。
 * Python/VBA 版で動作検証済みのロジックを TypeScript に移植。
 */

import type {
  Member,
  PositionDef,
  PreviousDay,
  Assignment,
  GenerateOptions,
  GenerateResult,
  ValidationIssue,
  Column,
  Rank,
} from "./types";
import { RANK_VALUE } from "./types";
import { N_SLOTS, SLOT_INDEX, TIME_SLOTS } from "./time-markers";

const {
  S_10_11,
  S_12_13,
  S_17_18,
  S_20_21,
  S_22_23,
  S_4_5,
} = SLOT_INDEX;

// =============================================================
// ブロック判定
// =============================================================
export function isBlocked(
  member: Member,
  slot: number,
  positions: Map<string, PositionDef>
): boolean {
  if (member.absent) return true;

  // 全ポジションの × を合算
  for (const posName of member.positions) {
    const def = positions.get(posName);
    if (def && def.blockedSlots.has(slot)) return true;
  }

  // 追加除外 (時間帯範囲)
  if (member.exclusions) {
    for (const ex of member.exclusions) {
      const s = Math.min(ex.fromBoundary, ex.toBoundary);
      const e = Math.max(ex.fromBoundary, ex.toBoundary);
      // boundary 範囲 [s, e) → slot 範囲 [s, e-1]
      if (slot >= s && slot <= e - 1) return true;
    }
  }

  return false;
}

// =============================================================
// スロット処理順: 10-17 を最優先
// =============================================================
function buildSlotProcessOrder(): number[] {
  const order: number[] = [];
  for (let i = S_10_11; i < S_17_18; i++) order.push(i); // 2..8
  for (let i = 9; i <= 15; i++) order.push(i); // 17-24
  for (let i = 16; i <= 23; i++) order.push(i); // 0-8
  order.push(24); // 8〜8:40
  order.push(0); // 8:40〜9
  order.push(1); // 9〜10
  return order;
}

function isLateNight(slot: number): boolean {
  return slot >= S_22_23 && slot <= S_4_5 - 1;
}

function countDaytime(name: string, cells: string[][], current: number): number {
  let cnt = 0;
  for (let i = S_10_11; i < S_17_18; i++) {
    if (i === current) break;
    for (let c = 0; c < 2; c++) {
      if (cells[c][i] === name) cnt++;
    }
  }
  return cnt;
}

function countLateNight(name: string, cells: string[][], current: number): number {
  let cnt = 0;
  for (let i = S_22_23; i < S_4_5; i++) {
    if (i === current) break;
    for (let c = 0; c < 2; c++) {
      if (cells[c][i] === name) cnt++;
    }
  }
  return cnt;
}

// =============================================================
// スコアリング
// =============================================================
function scoreCandidate(
  member: Member,
  col: Column,
  slot: number,
  previousDay: PreviousDay,
  cells: string[][],
  workCount: Map<string, number>
): number {
  let score = 0;

  // (a) 総勤務回数バランス
  score += (workCount.get(member.name) ?? 0) * 10;

  // (a2) 階級による通信/受付 優先
  const rv = RANK_VALUE[member.rank] ?? 0;
  if (col === 0) {
    // 通信
    if (rv === 4 || rv === 3) score -= 30;
    else if (rv === 2) score += 30;
    else if (rv === 1) score += 300;
  } else {
    // 受付
    if (rv === 4 || rv === 3) score += 20;
    else if (rv === 2 || rv === 1) score -= 30;
  }

  // (b) 前日同時刻と同一人物
  const prev = previousDay.bySlot.get(slot) ?? ["", ""];
  const prevName = prev[col];
  if (prevName && prevName === member.name) {
    if (slot >= S_20_21) score += 1000;
    else score += 50;
  }

  // (c) 前日の 1 つ上 (slot+1) を軽優先
  if (slot + 1 < N_SLOTS) {
    const pn = previousDay.bySlot.get(slot + 1) ?? ["", ""];
    if (pn[col] === member.name) score -= 5;
  }

  // (f) 10-17時未勤務者を強優先
  if (slot >= S_10_11 && slot <= S_17_18 - 1) {
    if (countDaytime(member.name, cells, slot) === 0) score -= 200;
  }

  // (g) 深夜(22-4時)は一人1回
  if (isLateNight(slot)) {
    if (countLateNight(member.name, cells, slot) >= 1) score += 500;
  }

  // (h) 12勤と17勤は別人
  if (slot === S_17_18 && cells[col][S_12_13] === member.name) score += 500;
  if (slot === S_12_13 && cells[col][S_17_18] === member.name) score += 500;

  // (i) 連続割当ペナルティ (前後どちらかが同一)
  if (slot > 0 && cells[col][slot - 1] === member.name) score += 50;
  if (slot < N_SLOTS - 1 && cells[col][slot + 1] === member.name) score += 50;

  return score;
}

// =============================================================
// メイン
// =============================================================
export function generate(options: GenerateOptions): GenerateResult {
  const { members, positions, previousDay, date } = options;

  // 1日全スロット × でないメンバーのみ割当対象
  const available = members.filter((m) =>
    Array.from({ length: N_SLOTS }, (_, i) => i).some((slot) => !isBlocked(m, slot, positions))
  );

  const cells: string[][] = [Array(N_SLOTS).fill(""), Array(N_SLOTS).fill("")];
  const workCount = new Map<string, number>();
  for (const m of available) workCount.set(m.name, 0);

  const order = buildSlotProcessOrder();

  for (const slot of order) {
    for (let col = 0 as Column; col < 2; col = (col + 1) as Column) {
      // 受付 8:40〜9 は斜線 (誰も割り当てない)
      if (col === 1 && slot === 0) {
        cells[1][0] = "";
        continue;
      }

      const candidates = available.filter((m) => {
        if (isBlocked(m, slot, positions)) return false;
        // 通信と受付は同時刻で別人
        if (col === 1 && cells[0][slot] === m.name) return false;
        return true;
      });
      if (candidates.length === 0) {
        cells[col][slot] = "";
        continue;
      }

      let bestName = "";
      let bestScore = Infinity;
      for (const c of candidates) {
        const s = scoreCandidate(c, col, slot, previousDay, cells, workCount);
        if (s < bestScore) {
          bestScore = s;
          bestName = c.name;
        }
      }

      cells[col][slot] = bestName;
      workCount.set(bestName, (workCount.get(bestName) ?? 0) + 1);
    }
  }

  // 後処理: 12勤と17勤が同じなら差し替え
  enforceDistinct12_17(cells, available, positions);

  const assignment: Assignment = {
    cells: [cells[0].slice(), cells[1].slice()],
    generatedAt: new Date().toISOString(),
    date,
  };

  const issues = validate(assignment, members, previousDay, positions);
  return { assignment, issues };
}

function enforceDistinct12_17(
  cells: string[][],
  members: Member[],
  positions: Map<string, PositionDef>
): void {
  for (let col = 0; col < 2; col++) {
    if (cells[col][S_12_13] && cells[col][S_12_13] === cells[col][S_17_18]) {
      for (const m of members) {
        if (m.name === cells[col][S_12_13]) continue;
        if (isBlocked(m, S_17_18, positions)) continue;
        if (cells[1 - col][S_17_18] === m.name) continue;
        cells[col][S_17_18] = m.name;
        break;
      }
    }
  }
}

// =============================================================
// 検証
// =============================================================
export function validate(
  assignment: Assignment,
  members: Member[],
  previousDay: PreviousDay,
  positions: Map<string, PositionDef>
): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const { cells } = assignment;

  // 空欄チェック (受付 8:40〜9 は除く)
  for (let i = 0; i < N_SLOTS; i++) {
    for (let col = 0; col < 2; col++) {
      if (col === 1 && i === 0) continue;
      if (!cells[col][i]) {
        issues.push({
          kind: "empty_slot",
          slot: i,
          message: `${TIME_SLOTS[i]} / ${col === 0 ? "通信" : "受付"} が空欄 (割当候補なし)`,
        });
      }
    }
  }

  // 10-17時 全員カバー (日中 全スロット × の人は除外)
  const covered = new Map<string, number>();
  for (const m of members) covered.set(m.name, 0);
  for (let i = S_10_11; i < S_17_18; i++) {
    for (let col = 0; col < 2; col++) {
      const n = cells[col][i];
      if (n && covered.has(n)) covered.set(n, (covered.get(n) ?? 0) + 1);
    }
  }
  for (const m of members) {
    const c = covered.get(m.name) ?? 0;
    if (c === 0) {
      // 10-17 の全スロットでブロックされていれば警告しない
      const blockedAll = Array.from({ length: S_17_18 - S_10_11 }, (_, k) => k + S_10_11)
        .every((slot) => isBlocked(m, slot, positions));
      if (!blockedAll) {
        issues.push({
          kind: "daytime_uncovered",
          name: m.name,
          message: `${m.name} は 10-17時に未勤務`,
        });
      }
    }
  }

  // 夜20時以降 前日同時刻と同一人物
  for (let i = S_20_21; i < N_SLOTS; i++) {
    const prev = previousDay.bySlot.get(i) ?? ["", ""];
    for (let col = 0; col < 2; col++) {
      if (cells[col][i] && cells[col][i] === prev[col]) {
        issues.push({
          kind: "night_duplicate_with_prev",
          slot: i,
          name: cells[col][i],
          message: `${TIME_SLOTS[i]} / ${col === 0 ? "通信" : "受付"} が前日と同じ (${cells[col][i]})`,
        });
      }
    }
  }

  // 12勤と17勤が同じ
  for (let col = 0; col < 2; col++) {
    if (cells[col][S_12_13] && cells[col][S_12_13] === cells[col][S_17_18]) {
      issues.push({
        kind: "12_17_same_person",
        name: cells[col][S_12_13],
        message: `12勤と17勤が同じ人物 (${cells[col][S_12_13]})`,
      });
    }
  }

  return issues;
}
