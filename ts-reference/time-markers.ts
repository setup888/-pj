/**
 * 時間帯 (25 slot) と時刻境界 (26 boundary) の定義
 *
 * スロット境界 i は スロット i-1 の終わりかつ スロット i の始まり。
 * スロット 0 = 8:40〜9 (boundary 0..1)
 * スロット 24 = 8〜8:40 (boundary 24..25)
 */

export const TIME_SLOTS: readonly string[] = [
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
] as const;

export const N_SLOTS = TIME_SLOTS.length; // 25

/**
 * 時刻境界ラベル (UI 表示用)。インデックスが boundary。
 * 0 = 当番開始 8:40, 25 = 当番終了 翌 8:40
 */
export const TIME_MARKERS: readonly string[] = [
  "8:40",
  "9時",
  "10時",
  "11時",
  "12時",
  "13時",
  "14時",
  "15時",
  "16時",
  "17時",
  "18時",
  "19時",
  "20時",
  "21時",
  "22時",
  "23時",
  "0時",
  "1時",
  "2時",
  "3時",
  "4時",
  "5時",
  "6時",
  "7時",
  "8時",
  "8:40(翌)",
] as const;

export const N_BOUNDARIES = TIME_MARKERS.length; // 26

/** ラベル → boundary index */
export const MARKER_TO_BOUNDARY: ReadonlyMap<string, number> = new Map(
  TIME_MARKERS.map((m, i) => [m, i])
);

/**
 * "9時" から "17時" のような指定を slot index の範囲 [start, end] に変換。
 * 例: fromBoundary=1 (9時), toBoundary=9 (17時) → [1, 8] (slot 9〜10〜...〜16〜17)
 * 範囲が逆なら自動で入れ替え。
 */
export function boundaryRangeToSlots(
  fromBoundary: number,
  toBoundary: number
): { start: number; end: number } | null {
  let s = fromBoundary;
  let e = toBoundary;
  if (e < s) [s, e] = [e, s];
  if (e <= s) return null;
  return { start: s, end: e - 1 };
}

/**
 * 各スロット内の時間的インデックス (スコアリング用)。
 * 特別な定数は ALGORITHM 側で使う。
 */
export const SLOT_INDEX = {
  S_8_9: 0, //  8:40〜9
  S_9_10: 1, //  9〜10
  S_10_11: 2,
  S_11_12: 3,
  S_12_13: 4,
  S_13_14: 5,
  S_14_15: 6,
  S_15_16: 7,
  S_16_17: 8,
  S_17_18: 9,
  S_18_19: 10,
  S_19_20: 11,
  S_20_21: 12,
  S_21_22: 13,
  S_22_23: 14,
  S_23_24: 15,
  S_0_1: 16,
  S_1_2: 17,
  S_2_3: 18,
  S_3_4: 19,
  S_4_5: 20,
  S_5_6: 21,
  S_6_7: 22,
  S_7_8: 23,
  S_8_840: 24,
} as const;
