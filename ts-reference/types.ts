/**
 * 執務表 (Duty Roster) 型定義
 *
 * fire-duty の既存 Member / Position と噛み合う設計。
 * 外部依存なし、React/Vue/素のJS いずれでも使える。
 */

// =============================================================
// 階級 (rank)
// =============================================================
export type Rank = "司令補" | "士長" | "副士長" | "消防士" | "";

export const RANK_VALUE: Record<Rank, number> = {
  司令補: 4,
  士長: 3,
  副士長: 2,
  消防士: 1,
  "": 0,
};

// =============================================================
// 隊員 (fire-duty の Member と互換)
// =============================================================
export interface Member {
  id: string;
  name: string;
  rank: Rank;
  /** その日の警防態勢ポジション。複数可（例: ["ポンプ隊", "当直"]） */
  positions: string[];
  /**
   * その日の追加除外時間帯（ad-hoc）
   * 例: 方面訓練で 9時〜13時不在 → [{ fromBoundary: 1, toBoundary: 5 }]
   */
  exclusions?: TimeRangeBoundary[];
  /** 休暇/出向等で丸一日不在 */
  absent?: boolean;
}

// =============================================================
// ポジション定義
// =============================================================
export interface PositionDef {
  name: string;
  /** ブロックするスロット index (0..24) の集合 */
  blockedSlots: Set<number>;
}

// =============================================================
// 時間帯
// =============================================================
export type Slot = number; // 0..24 (25 枠)
export type Boundary = number; // 0..25 (26 境界)

export interface TimeRangeBoundary {
  /** 開始境界 inclusive (e.g., 1 = 9時) */
  fromBoundary: Boundary;
  /** 終了境界 exclusive (e.g., 9 = 17時) */
  toBoundary: Boundary;
}

// =============================================================
// 割当
// =============================================================
/** 通信=0, 受付=1 */
export type Column = 0 | 1;

export interface Assignment {
  /** [通信のslot配列, 受付のslot配列], 各 length=25。"" = 空欄 */
  cells: [string[], string[]];
  /** 生成日時 ISO */
  generatedAt: string;
  /** 当番日 (yyyy-mm-dd) */
  date: string;
}

export interface PreviousDay {
  /** slot index -> [通信名, 受付名] */
  bySlot: Map<number, [string, string]>;
}

// =============================================================
// 生成オプション
// =============================================================
export interface GenerateOptions {
  members: Member[];
  positions: Map<string, PositionDef>;
  previousDay: PreviousDay;
  /** 当番日 yyyy-mm-dd */
  date: string;
  /** 休日フラグ */
  isHoliday?: boolean;
}

// =============================================================
// 検証結果
// =============================================================
export interface ValidationIssue {
  kind:
    | "empty_slot"
    | "daytime_uncovered"
    | "night_duplicate_with_prev"
    | "12_17_same_person";
  message: string;
  /** 該当 slot (あれば) */
  slot?: number;
  /** 該当氏名 (あれば) */
  name?: string;
}

export interface GenerateResult {
  assignment: Assignment;
  issues: ValidationIssue[];
}
