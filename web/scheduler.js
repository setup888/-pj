// =============================================================
// 執務表 スケジューリングロジック (sim_scheduler.py / ShiftScheduler.bas のJS移植)
//   完全データ駆動:
//     - ポジション×時間帯のブロックマトリクス
//     - 当日の「ポジション + 追加除外時間帯」
//   スコアリング規則はVBA/Pythonと同一。
// =============================================================

export const TIME_SLOTS = [
  "8:40〜9", "9〜10", "10〜11", "11〜12", "12〜13", "13〜14",
  "14〜15", "15〜16", "16〜17", "17〜18", "18〜19", "19〜20",
  "20〜21", "21〜22", "22〜23", "23〜24",
  "0〜1", "1〜2", "2〜3", "3〜4", "4〜5", "5〜6",
  "6〜7", "7〜8", "8〜8:40",
];
export const N = TIME_SLOTS.length;

export const TIME_MARKERS = [
  "8:40", "9時", "10時", "11時", "12時", "13時", "14時", "15時", "16時", "17時",
  "18時", "19時", "20時", "21時", "22時", "23時", "0時", "1時", "2時", "3時",
  "4時", "5時", "6時", "7時", "8時", "8:40(翌)",
];
export const TIME_MARKER_MAP = Object.fromEntries(
  TIME_MARKERS.map((m, i) => [m, i])
);

export const S_10_11 = 2;
export const S_12_13 = 4;
export const S_17_18 = 9;
export const S_20_21 = 12;
export const S_22_23 = 14;
export const S_4_5 = 20;

export function markerRange(fromLabel, toLabel) {
  let s = TIME_MARKER_MAP[fromLabel];
  let e = TIME_MARKER_MAP[toLabel];
  if (s == null || e == null) return null;
  if (e < s) [s, e] = [e, s];
  if (e === s) return null;
  return [s, e - 1];
}

// positions: { posName: Set(slotIdx) } — ブロック対象スロットの集合
export function isBlocked(name, slot, daily, positions) {
  const posList = daily.positions[name] || [];
  for (const pos of posList) {
    const blocks = positions[pos];
    if (blocks && blocks.has(slot)) return true;
  }
  for (const [s, e] of daily.exclusions[name] || []) {
    if (s <= slot && slot <= e) return true;
  }
  return false;
}

function isLateNight(slot) {
  return slot >= S_22_23 && slot <= S_4_5 - 1;
}

function daytimeCount(name, assign, cur) {
  let cnt = 0;
  for (let i = S_10_11; i < S_17_18; i++) {
    if (i === cur) break;
    for (let c = 0; c < 2; c++) if (assign[c][i] === name) cnt++;
  }
  return cnt;
}

function lateCount(name, assign, cur) {
  let cnt = 0;
  for (let i = S_22_23; i < S_4_5; i++) {
    if (i === cur) break;
    for (let c = 0; c < 2; c++) if (assign[c][i] === name) cnt++;
  }
  return cnt;
}

function scoreCandidate(name, col, slot, prevDay, assign, workCount) {
  let s = (workCount[name] || 0) * 10;
  const prevPair = prevDay[slot] || ["", ""];
  if (prevPair[col] && prevPair[col] === name) {
    s += slot >= S_20_21 ? 1000 : 50;
  }
  if (slot + 1 <= N - 1) {
    const p = (prevDay[slot + 1] || ["", ""])[col];
    if (p === name) s -= 5;
  }
  if (slot >= S_10_11 && slot <= S_17_18 - 1) {
    if (daytimeCount(name, assign, slot) === 0) s -= 200;
  }
  if (isLateNight(slot)) {
    if (lateCount(name, assign, slot) >= 1) s += 500;
  }
  if (slot === S_17_18 && assign[col][S_12_13] === name) s += 500;
  if (slot === S_12_13 && assign[col][S_17_18] === name) s += 500;
  if (slot > 0 && assign[col][slot - 1] === name) s += 50;
  if (slot < N - 1 && assign[col][slot + 1] === name) s += 50;
  return s;
}

function pick(col, slot, members, daily, positions, prevDay, assign, workCount) {
  let cands = members.filter((n) => !isBlocked(n, slot, daily, positions));
  if (col === 1) cands = cands.filter((n) => assign[0][slot] !== n);
  if (cands.length === 0) return "";
  let best = "";
  let bestScore = Infinity;
  for (const n of cands) {
    const sc = scoreCandidate(n, col, slot, prevDay, assign, workCount);
    if (sc < bestScore) {
      bestScore = sc;
      best = n;
    }
  }
  if (best) workCount[best] = (workCount[best] || 0) + 1;
  return best;
}

function slotOrder() {
  const order = [];
  for (let i = S_10_11; i < S_17_18; i++) order.push(i); // 10-17
  for (let i = 9; i <= 15; i++) order.push(i);            // 17-24
  for (let i = 16; i <= 23; i++) order.push(i);           // 0-8
  order.push(24, 0, 1);                                    // 8-8:40, 8:40-9, 9-10
  return order;
}

function allBlockedDaytime(name, daily, positions) {
  for (let i = S_10_11; i < S_17_18; i++) {
    if (!isBlocked(name, i, daily, positions)) return false;
  }
  return true;
}

export function generate(roster, daily, positions, prevDay) {
  const members = roster.filter((n) => {
    for (let i = 0; i < N; i++) {
      if (!isBlocked(n, i, daily, positions)) return true;
    }
    return false;
  });
  const assign = [Array(N).fill(""), Array(N).fill("")];
  const workCount = {};
  for (const m of members) workCount[m] = 0;
  for (const slot of slotOrder()) {
    for (let col = 0; col < 2; col++) {
      assign[col][slot] = pick(col, slot, members, daily, positions, prevDay, assign, workCount);
    }
  }
  // 12/17 distinct
  for (let col = 0; col < 2; col++) {
    if (assign[col][S_12_13] && assign[col][S_12_13] === assign[col][S_17_18]) {
      for (const alt of members) {
        if (alt === assign[col][S_12_13]) continue;
        if (isBlocked(alt, S_17_18, daily, positions)) continue;
        if (assign[1 - col][S_17_18] === alt) continue;
        assign[col][S_17_18] = alt;
        break;
      }
    }
  }
  return assign;
}

export function validate(assign, roster, daily, positions, prevDay) {
  const msgs = [];
  for (let i = 0; i < N; i++) {
    for (let col = 0; col < 2; col++) {
      if (!assign[col][i]) {
        msgs.push(`${TIME_SLOTS[i]} / ${col === 0 ? "通信" : "受付"} が空欄`);
      }
    }
  }
  const covered = {};
  for (const n of roster) covered[n] = 0;
  for (let i = S_10_11; i < S_17_18; i++) {
    for (let col = 0; col < 2; col++) {
      const n = assign[col][i];
      if (n && n in covered) covered[n]++;
    }
  }
  for (const n of roster) {
    if (covered[n] === 0 && !allBlockedDaytime(n, daily, positions)) {
      msgs.push(`${n} は 10-17時に未勤務`);
    }
  }
  for (let i = S_20_21; i < N; i++) {
    for (let col = 0; col < 2; col++) {
      const cur = assign[col][i];
      const pv = (prevDay[i] || ["", ""])[col];
      if (cur && cur === pv) {
        msgs.push(`${TIME_SLOTS[i]} / ${col === 0 ? "通信" : "受付"} が前日と同じ (${cur})`);
      }
    }
  }
  return msgs;
}
