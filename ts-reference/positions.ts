/**
 * デフォルトのポジション定義
 *
 * fire-duty に持ち込む際は、ユーザーが設定画面で編集できる状態にしておく。
 * localStorage or DB でユーザー編集版を保持し、ここは初期値として使う。
 */

import type { PositionDef } from "./types";
import { TIME_SLOTS } from "./time-markers";

/** スロット番号を作るヘルパー */
function slotsFrom(labels: string[]): Set<number> {
  const index = new Map(TIME_SLOTS.map((s, i) => [s, i]));
  const s = new Set<number>();
  for (const label of labels) {
    const i = index.get(label);
    if (i !== undefined) s.add(i);
  }
  return s;
}

const ALL_DAY = new Set<number>(Array.from({ length: 25 }, (_, i) => i));

/**
 * 初期ポジション一覧 (名前) と ブロック時間帯
 */
export const DEFAULT_POSITIONS: PositionDef[] = [
  // 基本隊 (通常ブロックなし)
  { name: "ポンプ隊", blockedSlots: new Set() },
  { name: "救助隊", blockedSlots: new Set() },
  { name: "はしご隊", blockedSlots: new Set() },
  { name: "伝令", blockedSlots: new Set() },
  { name: "通信担当", blockedSlots: new Set() },
  { name: "情報担当", blockedSlots: new Set() },
  { name: "情報員", blockedSlots: new Set() },
  { name: "署隊長伝令", blockedSlots: new Set() },
  { name: "残留", blockedSlots: new Set() },
  { name: "署隊本部支援員", blockedSlots: new Set() },
  { name: "その他", blockedSlots: new Set() },

  // 全時間ブロック (執務表対象外)
  { name: "救急隊", blockedSlots: ALL_DAY },
  { name: "警防力", blockedSlots: ALL_DAY },
  { name: "休暇", blockedSlots: ALL_DAY },
  { name: "研修/出向", blockedSlots: ALL_DAY },

  // 時間帯ブロック
  {
    name: "日中救急",
    blockedSlots: slotsFrom([
      "8:40〜9", "9〜10", "10〜11", "11〜12", "12〜13",
      "13〜14", "14〜15", "15〜16", "16〜17", "17〜18",
    ]),
  },
  {
    name: "夜救急",
    blockedSlots: slotsFrom([
      "18〜19", "19〜20", "20〜21", "21〜22", "22〜23",
      "23〜24", "0〜1", "1〜2", "2〜3", "3〜4",
      "4〜5", "5〜6", "6〜7", "7〜8", "8〜8:40",
    ]),
  },
  {
    name: "食当",
    blockedSlots: slotsFrom(["14〜15", "15〜16", "16〜17"]),
  },
  {
    name: "当直",
    blockedSlots: slotsFrom(["18〜19", "19〜20", "6〜7", "7〜8"]),
  },
];

export function buildPositionMap(defs: PositionDef[]): Map<string, PositionDef> {
  const m = new Map<string, PositionDef>();
  for (const d of defs) m.set(d.name, d);
  return m;
}
