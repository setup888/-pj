/**
 * scheduler.ts テストシナリオ
 *
 * 実行方法:
 *   npx tsx ts-reference/test-scheduler.ts
 *   (または package.json の scripts に登録)
 *
 * Python 版 (sim_scheduler.py) と同じサンプルで動作検証。
 */

import type { Member, PreviousDay, TimeRangeBoundary } from "./types";
import { generate } from "./scheduler";
import { buildPositionMap, DEFAULT_POSITIONS } from "./positions";
import { TIME_SLOTS, N_SLOTS, MARKER_TO_BOUNDARY } from "./time-markers";

function markerRange(from: string, to: string): TimeRangeBoundary {
  const f = MARKER_TO_BOUNDARY.get(from);
  const t = MARKER_TO_BOUNDARY.get(to);
  if (f === undefined || t === undefined) {
    throw new Error(`Invalid marker: ${from} or ${to}`);
  }
  return { fromBoundary: f, toBoundary: t };
}

const positions = buildPositionMap(DEFAULT_POSITIONS);

const members: Member[] = [
  { id: "1", name: "原田 陽一郎", rank: "司令補", positions: ["ポンプ隊", "当直"] },
  { id: "2", name: "梅村 侑志", rank: "司令補", positions: ["救助隊", "食当"] },
  { id: "3", name: "小西 隼人", rank: "士長", positions: ["ポンプ隊"] },
  { id: "4", name: "鍋谷 昇", rank: "士長", positions: ["はしご隊"],
    exclusions: [markerRange("9時", "13時")] },
  { id: "5", name: "村山 哲也", rank: "士長", positions: ["食当"] },
  { id: "6", name: "長田 智紀", rank: "士長", positions: ["伝令"] },
  { id: "7", name: "和田 浩司", rank: "副士長", positions: ["伝令"] },
  { id: "8", name: "長友 亮澄", rank: "副士長", positions: ["情報員"] },
  { id: "9", name: "尾坂 友梨", rank: "副士長", positions: ["通信担当"] },
  { id: "10", name: "山川 敦史", rank: "副士長", positions: ["情報担当"] },
  { id: "11", name: "金子 卓磨", rank: "副士長", positions: ["日中救急", "残留"] },
  { id: "12", name: "永井 恵理", rank: "副士長", positions: ["残留"] },
  { id: "13", name: "中村 太一", rank: "消防士", positions: ["救助隊"] },
  { id: "14", name: "藤井 惇平", rank: "消防士", positions: ["ポンプ隊"] },
  { id: "15", name: "伊藤 祥輝", rank: "消防士", positions: ["夜救急"] },
  { id: "16", name: "飯塚 佑介", rank: "消防士", positions: ["その他"] },
  { id: "17", name: "後藤 直人", rank: "消防士", positions: ["署隊長伝令"] },
];

// 前日実績なし
const emptyPrev: PreviousDay = {
  bySlot: new Map(
    Array.from({ length: N_SLOTS }, (_, i) => [i, ["", ""]]) as [number, [string, string]][]
  ),
};

function printAssignment(cells: string[][], label: string) {
  console.log(`\n=== ${label} ===`);
  console.log(`${"時間帯".padEnd(8)} | ${"通信".padEnd(14)} | ${"受付".padEnd(14)}`);
  console.log("-".repeat(50));
  for (let i = 0; i < N_SLOTS; i++) {
    const t = TIME_SLOTS[i];
    const c0 = cells[0][i].padEnd(14);
    const c1 = cells[1][i].padEnd(14);
    console.log(`${t.padEnd(8)} | ${c0} | ${c1}`);
  }
}

function runDay(label: string, ms: Member[], prev: PreviousDay, date: string) {
  const result = generate({ members: ms, positions, previousDay: prev, date });
  printAssignment(result.assignment.cells, label);
  if (result.issues.length) {
    console.log("\n⚠ Issues:");
    for (const i of result.issues) console.log("  -", i.message);
  } else {
    console.log("\n✓ No issues");
  }
  return result.assignment;
}

// Day 1
const day1 = runDay("Day 1 (2026-04-16)", members, emptyPrev, "2026-04-16");

// Day 2 (ポジション変更)
const members2: Member[] = members.map((m) => ({ ...m, exclusions: undefined }));
members2[0].positions = ["ポンプ隊"]; // 原田
members2[1].positions = ["ポンプ隊", "当直"]; // 梅村
members2[2].positions = ["食当", "救助隊"]; // 小西
members2[3].positions = ["伝令"]; // 鍋谷
members2[4].positions = ["通信担当"]; // 村山
members2[10].positions = ["署隊長伝令"]; // 金子
members2[12].positions = ["休暇"]; // 中村
members2[14].positions = ["救急隊"]; // 伊藤

const prev2: PreviousDay = {
  bySlot: new Map(
    Array.from({ length: N_SLOTS }, (_, i) => [i, [day1.cells[0][i], day1.cells[1][i]]]) as [
      number,
      [string, string]
    ][]
  ),
};
const day2 = runDay("Day 2 (2026-04-19)", members2, prev2, "2026-04-19");

// 夜20時以降 前日重複チェック
console.log("\n=== 夜20時以降 D1→D2 通信比較 ===");
let dup = 0;
for (let i = 12; i < N_SLOTS; i++) {
  const d1 = day1.cells[0][i];
  const d2 = day2.cells[0][i];
  const mark = d1 && d1 === d2 ? "×" : "";
  if (mark) dup++;
  console.log(`${TIME_SLOTS[i].padEnd(8)} | ${d1.padEnd(14)} | ${d2.padEnd(14)} | ${mark}`);
}
console.log(`夜20時以降の重複: ${dup} 件`);
