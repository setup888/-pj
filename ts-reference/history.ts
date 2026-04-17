/**
 * 履歴管理 (localStorage永続化)
 *
 * 過去の生成結果を保存し、前日の一次指定者を自動で参照可能にする。
 */

import type { Assignment, PreviousDay } from "./types";
import { N_SLOTS } from "./time-markers";

const STORAGE_KEY = "fire-duty.duty-roster.history";

export interface HistoryRecord {
  date: string; // yyyy-mm-dd
  cells: [string[], string[]]; // 通信, 受付
  generatedAt: string;
}

export function loadHistory(): HistoryRecord[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as HistoryRecord[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveHistory(records: HistoryRecord[]): void {
  // 日付順ソート (降順)
  records.sort((a, b) => b.date.localeCompare(a.date));
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

export function upsertRecord(record: HistoryRecord): HistoryRecord[] {
  const list = loadHistory().filter((r) => r.date !== record.date);
  list.push(record);
  saveHistory(list);
  return list;
}

export function findByDate(date: string): HistoryRecord | null {
  return loadHistory().find((r) => r.date === date) ?? null;
}

/**
 * 指定日より前の直近の履歴から PreviousDay を構築
 */
export function buildPreviousDay(beforeDate: string): PreviousDay {
  const list = loadHistory()
    .filter((r) => r.date < beforeDate)
    .sort((a, b) => b.date.localeCompare(a.date));
  const latest = list[0];
  const map = new Map<number, [string, string]>();
  if (!latest) {
    for (let i = 0; i < N_SLOTS; i++) map.set(i, ["", ""]);
    return { bySlot: map };
  }
  for (let i = 0; i < N_SLOTS; i++) {
    map.set(i, [latest.cells[0][i] ?? "", latest.cells[1][i] ?? ""]);
  }
  return { bySlot: map };
}

export function assignmentToRecord(a: Assignment): HistoryRecord {
  return {
    date: a.date,
    cells: [a.cells[0].slice(), a.cells[1].slice()],
    generatedAt: a.generatedAt,
  };
}

/**
 * 特定日のレコード削除
 */
export function deleteByDate(date: string): HistoryRecord[] {
  const list = loadHistory().filter((r) => r.date !== date);
  saveHistory(list);
  return list;
}

/**
 * 全履歴クリア (デバッグ用)
 */
export function clearAllHistory(): void {
  localStorage.removeItem(STORAGE_KEY);
}
