/**
 * DutyRosterView 参考実装 (React)
 *
 * fire-duty の App 構成に合わせて調整してください。
 * 既存の Member / Leave (休暇) データを props で受け取る想定。
 *
 * このファイルは TypeScript + React のスケルトンです。
 * コンパイラ設定は fire-duty のものに従ってください (strict, JSX など)。
 */

import React, { useMemo, useState } from "react";
import type {
  Member,
  PositionDef,
  Assignment,
  ValidationIssue,
  TimeRangeBoundary,
} from "./types";
import { generate } from "./scheduler";
import {
  TIME_SLOTS,
  TIME_MARKERS,
  MARKER_TO_BOUNDARY,
  N_SLOTS,
} from "./time-markers";
import { DEFAULT_POSITIONS, buildPositionMap } from "./positions";
import {
  assignmentToRecord,
  buildPreviousDay,
  upsertRecord,
} from "./history";

interface Props {
  /** fire-duty の Member 一覧 (名前・階級を含むもの) */
  rosterMembers: Array<{
    id: string;
    name: string;
    rank: "司令補" | "士長" | "副士長" | "消防士";
  }>;
  /** 当番日 yyyy-mm-dd */
  date: string;
  /** その日の警防態勢 (氏名 -> ポジション配列) */
  positionAssignments: Map<string, string[]>;
  /** 休暇登録された氏名 */
  absentNames: Set<string>;
  /** ユーザー定義ポジション (空なら DEFAULT_POSITIONS を使用) */
  customPositions?: PositionDef[];
}

export function DutyRosterView({
  rosterMembers,
  date,
  positionAssignments,
  absentNames,
  customPositions,
}: Props) {
  // 1人ごとの追加除外 (方面訓練等)
  const [exclusions, setExclusions] = useState<Map<string, TimeRangeBoundary[]>>(new Map());
  const [assignment, setAssignment] = useState<Assignment | null>(null);
  const [issues, setIssues] = useState<ValidationIssue[]>([]);

  const positions = useMemo(
    () => buildPositionMap(customPositions ?? DEFAULT_POSITIONS),
    [customPositions]
  );

  const members: Member[] = useMemo(
    () =>
      rosterMembers.map((m) => ({
        id: m.id,
        name: m.name,
        rank: m.rank,
        positions: positionAssignments.get(m.name) ?? [],
        exclusions: exclusions.get(m.name),
        absent: absentNames.has(m.name),
      })),
    [rosterMembers, positionAssignments, absentNames, exclusions]
  );

  const handleGenerate = () => {
    const prev = buildPreviousDay(date);
    const result = generate({ members, positions, previousDay: prev, date });
    setAssignment(result.assignment);
    setIssues(result.issues);
    upsertRecord(assignmentToRecord(result.assignment));
  };

  const handleAddExclusion = (
    name: string,
    fromLabel: string,
    toLabel: string
  ) => {
    const f = MARKER_TO_BOUNDARY.get(fromLabel);
    const t = MARKER_TO_BOUNDARY.get(toLabel);
    if (f === undefined || t === undefined) return;
    setExclusions((prev) => {
      const next = new Map(prev);
      const current = next.get(name) ?? [];
      next.set(name, [...current, { fromBoundary: f, toBoundary: t }]);
      return next;
    });
  };

  return (
    <div style={{ padding: 16, fontFamily: "sans-serif" }}>
      <h2>執務表 — {date}</h2>

      <button
        onClick={handleGenerate}
        style={{
          padding: "12px 24px",
          fontSize: 16,
          fontWeight: "bold",
          background: "#d9534f",
          color: "white",
          border: "none",
          borderRadius: 8,
          cursor: "pointer",
        }}
      >
        🔥 本日を生成
      </button>

      {/* ExclusionEditor: 半日不在の時刻範囲を追加する UI */}
      <details style={{ marginTop: 16 }}>
        <summary>追加の除外時間帯を設定</summary>
        <ExclusionEditor
          members={rosterMembers}
          onAdd={handleAddExclusion}
        />
      </details>

      {/* 生成結果表示 */}
      {assignment && (
        <>
          <AssignmentTable cells={assignment.cells} />
          {issues.length > 0 && (
            <div style={{ marginTop: 16, color: "#b00" }}>
              <strong>⚠ 注意点:</strong>
              <ul>
                {issues.map((i, idx) => (
                  <li key={idx}>{i.message}</li>
                ))}
              </ul>
            </div>
          )}
        </>
      )}
    </div>
  );
}

function AssignmentTable({ cells }: { cells: [string[], string[]] }) {
  return (
    <table
      style={{
        borderCollapse: "collapse",
        marginTop: 16,
        width: "100%",
        maxWidth: 600,
      }}
    >
      <thead>
        <tr style={{ background: "#dce6f1" }}>
          <th style={th}>時間帯</th>
          <th style={th}>通信指令</th>
          <th style={th}>受付</th>
        </tr>
      </thead>
      <tbody>
        {TIME_SLOTS.map((slot, i) => (
          <tr key={i}>
            <td style={td}>{slot}</td>
            <td style={td}>{cells[0][i]}</td>
            <td
              style={{
                ...td,
                ...(i === 0 ? { background: "repeating-linear-gradient(45deg, #eee, #eee 4px, transparent 4px, transparent 8px)" } : {}),
              }}
            >
              {i === 0 ? "" : cells[1][i]}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function ExclusionEditor({
  members,
  onAdd,
}: {
  members: Array<{ id: string; name: string }>;
  onAdd: (name: string, fromLabel: string, toLabel: string) => void;
}) {
  const [name, setName] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");

  return (
    <div style={{ marginTop: 8, padding: 12, border: "1px solid #ccc", borderRadius: 6 }}>
      <div>
        氏名:
        <select value={name} onChange={(e) => setName(e.target.value)}>
          <option value="">-</option>
          {members.map((m) => (
            <option key={m.id} value={m.name}>
              {m.name}
            </option>
          ))}
        </select>
      </div>
      <div>
        から:
        <select value={from} onChange={(e) => setFrom(e.target.value)}>
          <option value="">-</option>
          {TIME_MARKERS.map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
        まで:
        <select value={to} onChange={(e) => setTo(e.target.value)}>
          <option value="">-</option>
          {TIME_MARKERS.map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
        <button
          onClick={() => {
            if (name && from && to) {
              onAdd(name, from, to);
              setName("");
              setFrom("");
              setTo("");
            }
          }}
        >
          追加
        </button>
      </div>
    </div>
  );
}

const th: React.CSSProperties = {
  border: "1px solid #999",
  padding: "6px 10px",
  textAlign: "center",
  fontWeight: "bold",
};
const td: React.CSSProperties = {
  border: "1px solid #999",
  padding: "6px 10px",
  textAlign: "center",
};
