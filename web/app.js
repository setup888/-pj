// =============================================================
// 執務表 Web版 エントリポイント
//   state は localStorage に永続化。
//   scheduler.js にロジック本体。
// =============================================================
import {
  TIME_SLOTS, TIME_MARKERS, N, markerRange,
  generate, validate,
} from "./scheduler.js";

const LS_KEY = "shumuhyo_web_state_v1";

// -------- 初期データ --------
const DEFAULT_POSITION_ORDER = [
  "ポンプ隊", "救助隊", "はしご隊", "救急隊",
  "日中救急", "夜救急",
  "伝令", "通信担当", "情報担当", "情報員",
  "署隊長伝令", "残留", "署隊本部支援員", "その他",
  "当直", "食当", "休暇", "研修/出向",
];

function allSlots() { return TIME_SLOTS.map((_, i) => i); }
function rangeSlots(from, toExcl) {
  const a = []; for (let i = from; i < toExcl; i++) a.push(i); return a;
}

const DEFAULT_POSITION_BLOCKS = {
  "救急隊": allSlots(),
  "休暇": allSlots(),
  "研修/出向": allSlots(),
  "食当": [6, 7, 8],            // 14-17
  "当直": [10, 11, 22, 23],     // 18-20, 6-8
  "日中救急": rangeSlots(0, 10),   // 8:40-18
  "夜救急": rangeSlots(10, N),     // 18-翌8:40
};

const DEFAULT_ROSTER = [
  "原田 陽一郎", "梅村 侑志", "小西 隼人", "鍋谷 昇", "村山 哲也",
  "長田 智紀", "和田 浩司", "長友 亮澄", "尾坂 友梨", "山川 敦史",
  "金子 卓磨", "永井 恵理", "中村 太一", "藤井 惇平", "伊藤 祥輝",
  "飯塚 佑介", "後藤 直人",
];

// -------- state --------
let state = load();
function load() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) {
      const s = JSON.parse(raw);
      if (s && s.version === 1) return s;
    }
  } catch (_) {}
  return defaultState();
}
function defaultState() {
  const positions = {};
  for (const p of DEFAULT_POSITION_ORDER) {
    positions[p] = DEFAULT_POSITION_BLOCKS[p] ? [...DEFAULT_POSITION_BLOCKS[p]] : [];
  }
  const todayStr = new Date().toISOString().slice(0, 10);
  return {
    version: 1,
    roster: [...DEFAULT_ROSTER],
    positions, // posName -> [slotIdx]
    positionOrder: [...DEFAULT_POSITION_ORDER],
    daily: {
      date: todayStr,
      holiday: 0,
      rows: {}, // name -> { kyuka, touchoku, shokutou, pos1, pos2, exc1from, exc1to, exc2from, exc2to, memo }
    },
    output: {
      date: todayStr,
      assign: [Array(N).fill(""), Array(N).fill("")],
      secondary: ["", "", "", "", "", ""],
      note: "",
    },
    history: [], // [{date, slot, comm, recep}]
  };
}
function save() {
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}

// -------- ヘルパ: daily -> scheduler input --------
function buildSchedulerDaily() {
  const positions = {};
  for (const row of state.daily.rows ? Object.entries(state.daily.rows) : []) { /* noop */ }

  const posByName = {};
  const excByName = {};
  for (const name of state.roster) {
    const r = state.daily.rows[name] || {};
    const pl = [];
    if (r.kyuka) pl.push("休暇");
    if (r.touchoku) pl.push("当直");
    if (r.shokutou) pl.push("食当");
    if (r.pos1) pl.push(r.pos1);
    if (r.pos2) pl.push(r.pos2);
    if (pl.length) posByName[name] = pl;
    const exc = [];
    for (const [a, b] of [["exc1from", "exc1to"], ["exc2from", "exc2to"]]) {
      if (r[a] && r[b]) {
        const rng = markerRange(r[a], r[b]);
        if (rng) exc.push(rng);
      }
    }
    if (exc.length) excByName[name] = exc;
  }
  return {
    date: state.daily.date,
    holiday: state.daily.holiday,
    positions: posByName,
    exclusions: excByName,
  };
}
function buildPositionsAsSets() {
  const out = {};
  for (const [k, arr] of Object.entries(state.positions)) {
    out[k] = new Set(arr);
  }
  return out;
}
function loadPrevDayFromHistory(beforeDate) {
  const d = {};
  for (let i = 0; i < N; i++) d[i] = ["", ""];
  const before = new Date(beforeDate).getTime();
  const dates = [...new Set(state.history.map((h) => h.date))]
    .filter((x) => new Date(x).getTime() < before)
    .sort();
  if (dates.length === 0) return d;
  const prev = dates[dates.length - 1];
  for (const h of state.history) {
    if (h.date !== prev) continue;
    const idx = TIME_SLOTS.indexOf(h.slot);
    if (idx >= 0) d[idx] = [h.comm || "", h.recep || ""];
  }
  return d;
}

// =============================================================
// タブ切替
// =============================================================
document.querySelectorAll(".tab").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

// =============================================================
// 執務表タブ
// =============================================================
function renderOutput() {
  const tbody = document.querySelector("#table-output tbody");
  tbody.innerHTML = "";
  for (let i = 0; i < N; i++) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${TIME_SLOTS[i]}</td>
      <td class="output" data-col="0" data-slot="${i}">${esc(state.output.assign[0][i] || "")}</td>
      <td contenteditable="true" data-edit="comm-change-${i}"></td>
      <td>${TIME_SLOTS[i]}</td>
      <td class="output" data-col="1" data-slot="${i}">${esc(state.output.assign[1][i] || "")}</td>
      <td contenteditable="true" data-edit="recep-change-${i}"></td>
      <td>${TIME_SLOTS[i]}</td>`;
    tbody.appendChild(tr);
  }
  document.getElementById("out-date").value = state.output.date || state.daily.date;
}

document.getElementById("out-date").addEventListener("change", (e) => {
  state.output.date = e.target.value;
  save();
});

// =============================================================
// 当日チェック タブ
// =============================================================
function renderDaily() {
  document.getElementById("daily-date").value = state.daily.date;
  document.getElementById("daily-holiday").value = String(state.daily.holiday || 0);
  const tbody = document.querySelector("#table-daily tbody");
  tbody.innerHTML = "";
  state.roster.forEach((name, idx) => {
    const r = state.daily.rows[name] || {};
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td style="text-align:left;">${esc(name)}</td>
      <td class="check-cell"><input type="checkbox" data-f="kyuka" ${r.kyuka ? "checked" : ""}></td>
      <td class="check-cell"><input type="checkbox" data-f="touchoku" ${r.touchoku ? "checked" : ""}></td>
      <td class="check-cell"><input type="checkbox" data-f="shokutou" ${r.shokutou ? "checked" : ""}></td>
      <td class="input-cell">${posSelect("pos1", r.pos1)}</td>
      <td class="input-cell">${posSelect("pos2", r.pos2)}</td>
      <td class="input-cell">${markerSelect("exc1from", r.exc1from)}</td>
      <td class="input-cell">${markerSelect("exc1to", r.exc1to)}</td>
      <td class="input-cell">${markerSelect("exc2from", r.exc2from)}</td>
      <td class="input-cell">${markerSelect("exc2to", r.exc2to)}</td>
      <td class="input-cell"><input type="text" data-f="memo" value="${esc(r.memo || "")}"></td>`;
    tr.dataset.name = name;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("[data-f]").forEach((el) => {
    el.addEventListener("change", onDailyChange);
    el.addEventListener("input", onDailyChange);
  });
}
function posSelect(field, val) {
  const opts = ['<option value=""></option>']
    .concat(state.positionOrder.map((p) => `<option value="${esc(p)}"${p === val ? " selected" : ""}>${esc(p)}</option>`));
  return `<select data-f="${field}">${opts.join("")}</select>`;
}
function markerSelect(field, val) {
  const opts = ['<option value=""></option>']
    .concat(TIME_MARKERS.map((m) => `<option value="${esc(m)}"${m === val ? " selected" : ""}>${esc(m)}</option>`));
  return `<select data-f="${field}">${opts.join("")}</select>`;
}
function onDailyChange(e) {
  const tr = e.target.closest("tr");
  const name = tr.dataset.name;
  if (!state.daily.rows[name]) state.daily.rows[name] = {};
  const f = e.target.dataset.f;
  const v = e.target.type === "checkbox" ? e.target.checked : e.target.value;
  state.daily.rows[name][f] = v;
  save();
}
document.getElementById("daily-date").addEventListener("change", (e) => {
  state.daily.date = e.target.value;
  save();
});
document.getElementById("daily-holiday").addEventListener("change", (e) => {
  state.daily.holiday = Number(e.target.value);
  save();
});

// =============================================================
// ポジション定義
// =============================================================
function renderPositions() {
  const head = document.getElementById("pos-head");
  head.innerHTML = '<th class="pos-name">ポジション</th>'
    + TIME_SLOTS.map((s) => `<th>${esc(s)}</th>`).join("")
    + "<th></th>";
  const tbody = document.querySelector("#table-positions tbody");
  tbody.innerHTML = "";
  state.positionOrder.forEach((pos, rowIdx) => {
    const blocks = new Set(state.positions[pos] || []);
    const tr = document.createElement("tr");
    tr.dataset.pos = pos;
    const nameCell = `<td class="input-cell pos-name"><input type="text" value="${esc(pos)}" data-f="name"></td>`;
    const slotCells = TIME_SLOTS.map((_, i) =>
      `<td class="block-toggle${blocks.has(i) ? " block" : ""}" data-slot="${i}">${blocks.has(i) ? "×" : ""}</td>`
    ).join("");
    const delCell = `<td><button class="row-del" title="削除">✕</button></td>`;
    tr.innerHTML = nameCell + slotCells + delCell;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("td.block-toggle").forEach((td) => {
    td.addEventListener("click", () => {
      const tr = td.closest("tr");
      const pos = tr.dataset.pos;
      const slot = Number(td.dataset.slot);
      const arr = state.positions[pos] || [];
      const i = arr.indexOf(slot);
      if (i >= 0) arr.splice(i, 1); else arr.push(slot);
      state.positions[pos] = arr;
      if (arr.indexOf(slot) >= 0) {
        td.classList.add("block");
        td.textContent = "×";
      } else {
        td.classList.remove("block");
        td.textContent = "";
      }
      save();
    });
  });
  tbody.querySelectorAll("input[data-f=name]").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      const oldName = inp.closest("tr").dataset.pos;
      const newName = e.target.value.trim();
      if (!newName || newName === oldName) return;
      if (state.positions[newName]) { alert("同名のポジションが既に存在します"); e.target.value = oldName; return; }
      state.positions[newName] = state.positions[oldName];
      delete state.positions[oldName];
      const idx = state.positionOrder.indexOf(oldName);
      if (idx >= 0) state.positionOrder[idx] = newName;
      save();
      renderPositions();
      renderDaily();
    });
  });
  tbody.querySelectorAll(".row-del").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tr = btn.closest("tr");
      const pos = tr.dataset.pos;
      if (!confirm(`ポジション「${pos}」を削除しますか？`)) return;
      delete state.positions[pos];
      state.positionOrder = state.positionOrder.filter((p) => p !== pos);
      save();
      renderPositions();
      renderDaily();
    });
  });
}
document.getElementById("btn-pos-add").addEventListener("click", () => {
  let name = prompt("新規ポジション名");
  if (!name) return;
  name = name.trim();
  if (!name || state.positions[name]) { alert("無効か同名が存在"); return; }
  state.positions[name] = [];
  state.positionOrder.push(name);
  save();
  renderPositions();
  renderDaily();
});
document.getElementById("btn-pos-reset").addEventListener("click", () => {
  if (!confirm("ポジション定義を初期値に戻しますか？")) return;
  const def = defaultState();
  state.positions = def.positions;
  state.positionOrder = def.positionOrder;
  save();
  renderPositions();
  renderDaily();
});

// =============================================================
// 名簿
// =============================================================
function renderRoster() {
  const tbody = document.querySelector("#table-roster tbody");
  tbody.innerHTML = "";
  state.roster.forEach((name, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td class="input-cell" style="text-align:left;"><input type="text" value="${esc(name)}"></td>
      <td><button class="row-del" title="削除">✕</button></td>`;
    tr.dataset.idx = i;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("input").forEach((inp) => {
    inp.addEventListener("change", (e) => {
      const idx = Number(inp.closest("tr").dataset.idx);
      const newName = e.target.value.trim();
      const oldName = state.roster[idx];
      if (!newName) { e.target.value = oldName; return; }
      state.roster[idx] = newName;
      if (state.daily.rows[oldName]) {
        state.daily.rows[newName] = state.daily.rows[oldName];
        delete state.daily.rows[oldName];
      }
      save();
      renderDaily();
    });
  });
  tbody.querySelectorAll(".row-del").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.closest("tr").dataset.idx);
      const name = state.roster[idx];
      if (!confirm(`「${name}」を名簿から削除しますか？`)) return;
      state.roster.splice(idx, 1);
      delete state.daily.rows[name];
      save();
      renderRoster();
      renderDaily();
    });
  });
}
document.getElementById("btn-roster-add").addEventListener("click", () => {
  state.roster.push("");
  save();
  renderRoster();
  renderDaily();
});

// =============================================================
// 履歴
// =============================================================
function renderHistory() {
  const tbody = document.querySelector("#table-history tbody");
  tbody.innerHTML = "";
  state.history.slice().reverse().forEach((h, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${esc(h.date)}</td>
      <td>${esc(h.slot)}</td>
      <td>${esc(h.comm)}</td>
      <td>${esc(h.recep)}</td>
      <td><button class="row-del" data-date="${esc(h.date)}">✕</button></td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll(".row-del").forEach((btn) => {
    btn.addEventListener("click", () => {
      const d = btn.dataset.date;
      if (!confirm(`${d} の履歴を全削除？`)) return;
      state.history = state.history.filter((h) => h.date !== d);
      save();
      renderHistory();
    });
  });
}
document.getElementById("btn-history-clear").addEventListener("click", () => {
  if (!confirm("すべての履歴を削除しますか？")) return;
  state.history = [];
  save();
  renderHistory();
});

// =============================================================
// アクション: 生成 / 過去表示 / エクスポート / インポート / 初期化 / 印刷
// =============================================================
document.getElementById("btn-generate").addEventListener("click", () => {
  const daily = buildSchedulerDaily();
  const positions = buildPositionsAsSets();
  const prevDay = loadPrevDayFromHistory(daily.date);
  if (!state.roster.filter((n) => n).length) { alert("名簿が空です"); return; }
  if (!daily.date) { alert("当番日を入力してください"); return; }
  const assign = generate(state.roster.filter((n) => n), daily, positions, prevDay);
  state.output.date = daily.date;
  state.output.assign = assign;
  // 履歴: 同日削除 → 追記
  state.history = state.history.filter((h) => h.date !== daily.date);
  for (let i = 0; i < N; i++) {
    state.history.push({
      date: daily.date, slot: TIME_SLOTS[i],
      comm: assign[0][i], recep: assign[1][i],
    });
  }
  save();
  renderOutput();
  renderHistory();
  const msgs = validate(assign, state.roster.filter((n) => n), daily, positions, prevDay);
  const w = document.getElementById("warnings");
  w.textContent = msgs.length ? "以下の注意点があります:\n - " + msgs.join("\n - ") : "";
  document.querySelector('[data-tab="output"]').click();
});

document.getElementById("btn-show-history").addEventListener("click", () => {
  const d = document.getElementById("out-date").value;
  if (!d) { alert("表示日付を入れてください"); return; }
  const rows = state.history.filter((h) => h.date === d);
  if (rows.length === 0) { alert(`${d} の履歴がありません`); return; }
  const assign = [Array(N).fill(""), Array(N).fill("")];
  for (const h of rows) {
    const i = TIME_SLOTS.indexOf(h.slot);
    if (i >= 0) { assign[0][i] = h.comm || ""; assign[1][i] = h.recep || ""; }
  }
  state.output.date = d;
  state.output.assign = assign;
  save();
  renderOutput();
  document.querySelector('[data-tab="output"]').click();
});

document.getElementById("btn-export").addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `shumuhyo_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});
document.getElementById("btn-import").addEventListener("click", () => {
  document.getElementById("file-import").click();
});
document.getElementById("file-import").addEventListener("change", (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const s = JSON.parse(rd.result);
      if (!s || s.version !== 1) throw new Error("不正な形式");
      if (!confirm("現在のデータを上書きします。よろしいですか？")) return;
      state = s;
      save();
      renderAll();
    } catch (err) {
      alert("読み込み失敗: " + err.message);
    }
  };
  rd.readAsText(f);
});
document.getElementById("btn-reset").addEventListener("click", () => {
  if (!confirm("すべて初期値に戻します。よろしいですか？")) return;
  state = defaultState();
  save();
  renderAll();
});
document.getElementById("btn-print").addEventListener("click", () => window.print());

// =============================================================
// ユーティリティ
// =============================================================
function esc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}
function renderAll() {
  renderOutput();
  renderDaily();
  renderPositions();
  renderRoster();
  renderHistory();
}
renderAll();
