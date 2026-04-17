# fire-duty への執務表機能統合ガイド

このディレクトリは `setup888/fire-duty` に **執務表 (Duty Roster) タブ** を追加するためのリファレンス実装。
Claude Code で作業する際、このフォルダの内容を参考にして fire-duty に統合してください。

## ファイル構成

| ファイル | 役割 | fire-duty での配置例 |
|---|---|---|
| `types.ts` | Member / Position / Assignment 型定義 | `src/features/duty-roster/types.ts` |
| `time-markers.ts` | 25スロット + 26境界の定数 | `src/features/duty-roster/time-markers.ts` |
| `positions.ts` | デフォルトポジション × 時間帯ブロック | `src/features/duty-roster/positions.ts` |
| `scheduler.ts` | 純関数のスケジューリングアルゴリズム | `src/features/duty-roster/scheduler.ts` |
| `history.ts` | localStorage 永続化 | `src/features/duty-roster/history.ts` |
| `test-scheduler.ts` | 動作検証用テスト | `tests/duty-roster.test.ts` (vitest等) |
| `DutyRosterView.example.tsx` | React UI スケルトン | `src/features/duty-roster/DutyRosterView.tsx` |

## 統合手順

### 1. ファイルをコピー

```bash
# fire-duty リポジトリで
mkdir -p src/features/duty-roster
# -pj の ts-reference/* をコピー (.example を外す)
```

### 2. fire-duty の既存 Member 型との対応

fire-duty の Member:

```ts
interface Member {
  id: string;
  name: string;
  stationId: string;
  division: 1 | 2 | 3;
  assignedVehicle: string;
  position: string;         // ← 単一。執務表用は複数にしたい
  rank: Rank;
  qualifications: Qualifications;
}
```

執務表では **その日のポジション (複数)** が必要。fire-duty 側で:

- 警防態勢画面の編集データ = 当日のポジション assignments
- これを `Map<string, string[]>` (氏名 → ポジション配列) に整形
- `DutyRosterView` の props `positionAssignments` に渡す

fire-duty の「目YD 指揮担当=平野」といった割当は、ポジション名 (`指揮担当` or 隊名 `目YD`) として扱うか、ユーザーが設定画面でマッピングを決める形が柔軟。

### 3. 休暇データの接続

fire-duty の休暇登録 → `absentNames: Set<string>` として DutyRosterView に渡す。

```ts
const absentNames = new Set(
  leaveRecords
    .filter((l) => l.date === dutyDate)
    .map((l) => l.memberName)
);
```

### 4. App.tsx にタブ追加

```tsx
// 既存
<TabBar>
  <Tab name="ホーム" />
  <Tab name="態勢" />
  <Tab name="休暇" />
  <Tab name="職員" />
  <Tab name="設定" />
</TabBar>

// 追加
<Tab name="執務表">
  <DutyRosterView
    rosterMembers={members}
    date={currentDutyDate}
    positionAssignments={todayPositions}
    absentNames={todayAbsent}
    customPositions={settings.positions}
  />
</Tab>
```

### 5. ポジション定義の設定画面追加

`DEFAULT_POSITIONS` を初期値として、設定画面で編集できるように。
localStorage に保存:

```ts
const POS_KEY = "fire-duty.duty-roster.positions";
localStorage.setItem(POS_KEY, JSON.stringify(customPositions));
```

### 6. AI 連携 (既存の Claude API ボタン)

既存の AI ボタンから「執務表を生成」も選べるように:

```ts
const prompt = `
あなたは消防部隊の執務表を作る専門家です。
以下の条件で、通信指令/受付の一次指定者を 25 時間帯分作ってください:

名簿:
${members.map((m) => `- ${m.name} (${m.rank})`).join("\n")}

本日のポジション:
${Array.from(positionAssignments.entries())
  .map(([n, pos]) => `- ${n}: ${pos.join(", ")}`)
  .join("\n")}

休暇: ${Array.from(absentNames).join(", ")}

ルール:
- 司令補/士長は 通信 優先、副士長/消防士は 受付 優先
- 消防士は 通信 原則不可
- 受付 8:40〜9 は斜線 (空欄)
- 10-17時は一人1回以上、12勤と17勤は別人
- 夜20時以降は前日同時刻と別人
- 深夜22-4時は一人1回

以下の JSON 形式で返してください:
{ "cells": [[通信25個], [受付25個]] }
`;

// Claude API 呼び出し → JSON パース → Assignment として反映
```

手動スケジューラと AI どちらも使える選択式にすると運用が柔軟。

### 7. 印刷対応 (Word様式そっくりの印刷)

`AssignmentTable` に print スタイル追加:

```css
@media print {
  body { margin: 0; }
  .no-print { display: none; }
  table { page-break-inside: avoid; }
}
```

A4 縦向きに収まるよう調整。

### 8. テスト

```bash
# vitest 導入済みなら
vitest src/features/duty-roster/

# または tsx で直接
npx tsx src/features/duty-roster/test-scheduler.ts
```

## 注意点

### fire-duty の階級型

fire-duty の `Rank` 型と本リファレンスの `Rank` 型が一致するか確認。
もし違えば `types.ts` の RANK_VALUE マップを合わせる。

### ポジション名の命名規則

fire-duty 側の呼称 (例: `目YD大隊長`) と 執務表向けのブロック対象 (例: `ポンプ隊`) が
一致しない可能性あり。設定画面で **警防態勢ポジション → 執務表ポジション** のマッピングを持つか、
または最初から同一命名で統一するか決める。

### 休暇以外の不在

fire-duty は 5 種別 (年休/事故/研修/交出/欠入)。執務表では
「その日出勤していない = 割当不能」とだけ判定すればよいので:

```ts
const absent = new Set(
  leaveRecords
    .filter((l) => l.date === date && l.kind !== "欠入")
    .map((l) => l.memberName)
);
```

### 履歴の保存先

localStorage はブラウザローカル。PC 間で同期したい場合は Cloudflare KV や
Google Drive (OAuth 済み) に同期する層を別途追加。

## 参考

### 元の Excel 版

- `setup888/-pj` (.xlsx + VBA) の仕様に基づく
- VBA 版でロジック検証済み
- Python 版 (`sim_scheduler.py`) で動作確認済み
- TypeScript 版 (`scheduler.ts`) は 3 世代目の移植

### 動作確認結果

17名サンプルで:
- 階級別 通信/受付 振り分け ✓
- 複数ポジション対応 (例: ポンプ隊+当直) ✓
- 除外時間帯 (方面訓練 9時〜13時等) ✓
- 夜20時以降の前日重複回避 ✓
- 12勤と17勤の別人化 ✓
- 深夜一人1回 ✓
- 受付 8:40〜9 斜線 ✓

## Claude Code へのプロンプト例

fire-duty リポジトリで Claude Code に指示する際の雛形:

```
github.com/setup888/-pj の ts-reference/ にある執務表アルゴリズムを
fire-duty に統合して。

1. ts-reference/*.ts を src/features/duty-roster/ にコピー
2. fire-duty の既存 Member 型と型定義を整合させる
3. App.tsx のタブに「執務表」を追加し、DutyRosterView を表示
4. 警防態勢画面から当日のポジションデータを流し込む
5. 休暇登録データを absentNames として渡す
6. 生成結果は localStorage (fire-duty.duty-roster.history) に保存

参考ファイル:
- ts-reference/INTEGRATION.md (このファイル)
- ts-reference/test-scheduler.ts (動作確認テスト)
```
