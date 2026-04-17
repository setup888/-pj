# 執務表自動化 汎用アプリ 仕様書

## 1. ビジョン

消防現場で**警防態勢・休暇管理・執務表作成に膨大な時間**が割かれている。
AIで効率化したいが、**消防組織への AI 直接導入はハードルが高い**（セキュリティ、承認プロセス、心理的抵抗）。

→ **アプリの皮を被せ、中身にAIを組み込む**ことで障壁を下げる。
→ ただし **どの署でもフィット**する必要あり（署ごとに運用ルールが微妙に違う）。

## 2. 設計原則

### 2.1 単一の中心思想
> **「すべての勤務制約は ポジション × 時間帯 のマトリクスで表現できる」**

食当・当直・救急隊・警防力・非常勤・休暇・研修 … 全て **ポジション** として扱い、
それぞれが「どの時間帯にブロックされるか」の設定だけを持つ。
ハードコードしない。

### 2.2 普遍ロジックと署別パラメータの分離

| 分類 | 内容 | 例 |
|---|---|---|
| **普遍コア** | どの署でも必要 | 時間帯×人の排他割当、カバレッジ、ローテーション |
| **署別Config** | 設定JSONで可変 | 階級名、ペア制約、斜線スロット、時間帯区切り |

### 2.3 AI はオプション、ローカルで7割動く
- 手動スケジューラだけで基本動作が完結
- Claude API 等は「AIに手直しさせる」「候補を増やす」用途
- 通信断でもローカルで使える

### 2.4 データは端末ローカル中心
- localStorage / IndexedDB で基本動作
- クラウド同期（Cloudflare KV, Google Drive 等）はオプション層として後付け

---

## 3. 用語定義

| 用語 | 意味 |
|---|---|
| **執務表** | 通信指令/受付の一次指定者を 25 時間帯 × 2 列で割り当てた表 |
| **当番日** | 3部制（署により異なる）の当該部が当番の日 |
| **ポジション** | その日の警防態勢で個々の隊員に割り当てられる役割（例: ポンプ隊, 当直, 食当, 救助隊, 非常勤=通信残留 等）|
| **時間帯 (slot)** | 執務表の 25 区分（8:40〜9, 9〜10, ... 8〜8:40） |
| **時刻境界 (boundary)** | 除外指定用の 26 時点（8:40, 9時, 10時, ... 8:40(翌)） |
| **ブロック** | 割当不可の時間帯 |
| **除外要件** | 特定日の特定人について追加でブロックする時間帯（方面訓練、研修等） |

---

## 4. データモデル

### 4.1 Member

```ts
interface Member {
  id: string;
  name: string;
  rank: string;                        // 署ごとに名称は可変
  employmentType: "fulltime" | "parttime";
  qualifications?: string[];            // 将来: 資格制約用

  // 当日の情報 (別ストアで管理する場合も)
  positions: string[];                  // その日のポジション (複数)
  exclusions?: TimeRangeBoundary[];     // 半日不在等 ad-hoc
  absent?: boolean;                      // 丸一日不在 (休暇等)
}
```

### 4.2 PositionDef

```ts
interface PositionDef {
  name: string;
  blockedSlots: Set<number>;            // ブロックする slot index (0..24)
  // 将来拡張:
  description?: string;
  category?: "基本隊" | "当直関連" | "特殊" | "不在";
}
```

### 4.3 SchedulerConfig (署別パラメータ)

```ts
interface SchedulerConfig {
  stationName: string;                   // "目黒消防署"

  // 時間帯定義 (署ごとに8:40始まりでない可能性)
  slots: string[];                       // ["8:40〜9", "9〜10", ...]
  boundaries: string[];                  // ["8:40", "9時", ...]

  // 階級
  rankOrder: string[];                   // ["司令補", "士長", "副士長", "消防士"]
  rankScoring: Record<string, {
    participates: boolean;               // 通常勤務対象か
    commWeight: number;                  // 通信側スコア (低いほど優先)
    recvWeight: number;                  // 受付側スコア
  }>;

  // ペア制約
  pairScoring: Record<string, number>;   // "通信rank|受付rank" → 重み

  // 不変制約の重み (細かい調整用)
  weights: {
    workCountBalance: number;            // 均等化: 10
    prevDaySameTime_night: number;       // 夜20以降 前日重複: 1000
    prevDaySameTime_day: number;         // 日中 前日重複: 50
    prevRotationBonus: number;           // 前日+1ずらし優先: -5
    daytimeUnworked: number;             // 10-17未勤務優先: -200
    lateNightDuplicate: number;          // 深夜2回目: 500
    distinct_12_17: number;              // 12勤/17勤重複: 500
    consecutiveSlot: number;             // 連続割当: 50
  };

  // 特殊スロット
  noAssignSlots: Array<{ col: 0|1; slot: number; label?: string }>;
  // 例: [{col: 1, slot: 0, label: "斜線"}] (目黒: 受付 8:40〜9)

  // カバレッジ判定時間帯
  daytimeCoverageRange: [number, number]; // [2, 9] = 10-17時

  // 物理レイアウト
  roomLayout: "separate" | "same-room";   // 通信と受付の物理配置

  // 当番サイクル
  dutyCycle: {
    type: "3-shift" | "4-shift" | "custom";
    pattern?: string[];                   // 曜日ローテ等
  };
}
```

### 4.4 DailyInput

```ts
interface DailyInput {
  date: string;                         // yyyy-mm-dd
  isHoliday: boolean;
  memberStates: Map<string, {
    positions: string[];                 // その日のポジション
    exclusions: TimeRangeBoundary[];    // ad-hoc 除外
    absent: boolean;
  }>;
}
```

---

## 5. 署別差異パターン（実例）

### 目黒署
- 通信/受付が **別室** → 同時刻別人必須
- 階級優先: 司令補/士長→通信、副士長/消防士→受付
- 受付 8:40〜9 は斜線（空欄）
- 司令/司令補は 通常 勤務対象外
- 士長以下で回す
- 非常勤は 通信残留ポジションで組み込まれる

### 想定される他署パターン
- **同室の署**: 閑散時間帯は兼任可能（同時刻同人でもOK）
- **人員少の署**: 司令補も勤務に入る（`司令補.participates = true`）
- **消防士ペア回避の署**: `pairScoring["消防士|消防士"] = +500`
- **士長-消防士ペア推奨**: `pairScoring["士長|消防士"] = -30`
- **9:00開始の署**: 時間帯定義を変更（`slots`, `boundaries`）
- **4部制の署**: `dutyCycle.type = "4-shift"`

---

## 6. 普遍ロジック（ハードコード可）

これは署に関係なく動く:
- 時間帯×人の排他割当（ポジションブロック + 除外）
- 候補からスコア最小を選ぶ
- `workCount` で均等化
- `daytimeCoverage`未カバー優先
- `distinct_12_17` 後処理
- 履歴の累積
- 前日からの自動取得

---

## 7. AI 統合設計

### 7.1 段階

| Phase | AI関与度 | 実装 |
|---|---|---|
| Phase 1 | なし | 手動スケジューラのみ。既に ts-reference/ で完了 |
| Phase 2 | 補助 | 「AIに手直しさせる」ボタン。違和感ある配置を見直し |
| Phase 3 | 主役 | 警防態勢の写真→OCR→自動で執務表まで。API → 構造化 JSON |

### 7.2 プロンプトテンプレート (Phase 2 例)

```
あなたは消防署の執務表を調整する専門家です。
以下のスケジューラ生成結果を見て、違和感がある箇所を指摘し、
可能なら修正案を提示してください。

制約:
- 通信と受付は同時刻で別人
- 10-17時は一人1回以上
- 夜20時以降は前日と同じ人を避ける
- 階級: {rankOrder}
- ペア制約: {pairScoring}

名簿:
{members}

前日:
{previousDay}

今日の割当:
{assignment}

JSON 形式で返答: { "issues": [...], "suggestions": [...] }
```

### 7.3 Phase 3: カメラOCR

スマホで警防態勢表を撮影 → Claude Vision API → JSON:

```json
{
  "date": "2026-04-06",
  "members": [
    { "position": "目YD大隊長", "name": "横山" },
    ...
  ]
}
```

これを DailyInput に変換 → スケジューラ実行。

---

## 8. 実装フェーズ

### Phase 1: MVP (fire-duty 統合) — 完了寄り
- [x] ts-reference/ でアルゴリズム移植済み
- [ ] fire-duty に `features/duty-roster/` として統合
- [ ] 執務表タブ追加、既存データと連携
- [ ] 生成 → localStorage 保存
- [ ] 印刷 CSS

### Phase 2: 汎用化
- [ ] `SchedulerConfig` の仕様策定（本ドキュメント）
- [ ] 設定画面: 階級・ペア制約・重みを JSON 編集
- [ ] デフォルト設定 = 目黒パターン
- [ ] 他署プリセット機能（「A署パターンを読み込む」）

### Phase 3: AI 統合
- [ ] Claude API 呼び出し層
- [ ] 「AI に手直しさせる」ボタン
- [ ] 警防態勢写真 OCR → 自動入力
- [ ] 過去パターン学習（ローカル履歴から）

### Phase 4: マルチ署展開
- [ ] 設定のエクスポート・インポート
- [ ] Cloudflare KV 等での署別設定ストア（オプション）
- [ ] 認証（既存 Google OAuth 活用）

---

## 9. テスト方針

### 9.1 普遍ロジックのユニットテスト
- `scheduler.ts` の純関数は TypeScript 直接テスト（vitest）
- Python 版 `sim_scheduler.py` と同じシナリオで結果一致確認

### 9.2 Config バリエーションテスト
- 目黒設定で生成 → 既存動作と一致
- 消防士ペア回避設定で生成 → 消防士×2 のスロットが発生しないこと
- 司令補参加設定で生成 → 司令補が通信に入ること

### 9.3 E2E テスト (fire-duty 側)
- 警防態勢入力 → 執務表生成 → 印刷プレビュー
- 履歴からの前日取得
- 設定画面で署パターン切替

---

## 10. 今後の要件追加ガイドライン

現場から新ルール出た時の扱い方:
1. **まず Config で表現できないか確認**（ほぼ全てできるはず）
2. Config で無理なら **ポジション定義を拡張**
3. それでも無理なら **普遍ロジックか検討**（普遍=どの署でも妥当なら実装OK）
4. 普遍ルールに紛れ込ませる「一署だけのルール」は**絶対に避ける**

---

## 11. 参考

- 実装リファレンス: `ts-reference/`（純関数 + Reactサンプル）
- 動作検証: `sim_scheduler.py`（Python版）、`test-scheduler.ts`（TS版）
- 既存 Excel 版: `build_template.py`, `ShiftScheduler.bas`（同一ロジック）
- fire-duty 既存 SPEC.md との関係: 本ドキュメントは **執務表機能の追加仕様**。
  fire-duty 本体仕様と重複しないよう設計。
