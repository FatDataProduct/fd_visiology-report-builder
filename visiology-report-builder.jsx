import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import * as XLSX from "sheetjs";

/* ═══════════════════════════════════════════
   CONSTANTS & HELPERS
   ═══════════════════════════════════════════ */

const uid = () => Math.random().toString(36).slice(2, 10);
const clone = (o) => JSON.parse(JSON.stringify(o));

const MARKER_TYPES = [
  { value: "cell",   label: "Ячейка",  icon: "◎", color: "#03FF94", bg: "#03FF9418", desc: "Единичное значение из виджета" },
  { value: "table",  label: "Таблица", icon: "▦", color: "#00DBFF", bg: "#00DBFF18", desc: "Табличное представление" },
  { value: "column", label: "Столбец", icon: "▥", color: "#B4FFDF", bg: "#B4FFDF18", desc: "Отдельный столбец данных" },
];

const markerName = (type, idx) => {
  const p = type === "table" ? "TBL" : type === "cell" ? "VAL" : "COL";
  return `{{VIS_${p}_${String(idx).padStart(2, "0")}}}`;
};

const COND_OPERATORS = [
  { value: "lt", label: "< меньше" },
  { value: "gt", label: "> больше" },
  { value: "eq", label: "= равно" },
  { value: "lte", label: "≤ не больше" },
  { value: "gte", label: "≥ не меньше" },
  { value: "neq", label: "≠ не равно" },
  { value: "contains", label: "содержит" },
];

const A4_ROWS_APPROX = 45;

/* ═══════════════════════════════════════════
   BLOCK TEMPLATES
   ═══════════════════════════════════════════ */

const BLOCK_TEMPLATES = [
  {
    id: "header", label: "Шапка отчёта", icon: "◈", desc: "Название, дата, период",
    rows: [
      [{ v: "РЕГЛАМЕНТНЫЙ ОТЧЁТ", b: true, fs: 16, cs: 4, al: "center", bg: "#1a365d", fc: "#ffffff" }, null, null, null],
      [{ v: "" , cs: 4 }, null, null, null],
      [{ v: "Дата формирования:", b: true, bg: "#edf2f7" }, { v: "{{ДАТА}}", mk: true, mt: "cell" }, { v: "Отчётный период:", b: true, bg: "#edf2f7" }, { v: "{{ПЕРИОД}}", mk: true, mt: "cell" }],
      [{ v: "" , cs: 4 }, null, null, null],
    ],
  },
  {
    id: "kpi", label: "KPI-панель", icon: "◉", desc: "Ключевые показатели",
    rows: [
      [{ v: "Показатель 1", b: true, bg: "#edf2f7", al: "center" }, { v: "Показатель 2", b: true, bg: "#edf2f7", al: "center" }, { v: "Показатель 3", b: true, bg: "#edf2f7", al: "center" }, { v: "Показатель 4", b: true, bg: "#edf2f7", al: "center" }],
      [{ v: "{{MK}}", mk: true, mt: "cell", al: "center", fs: 14, b: true }, { v: "{{MK}}", mk: true, mt: "cell", al: "center", fs: 14, b: true }, { v: "{{MK}}", mk: true, mt: "cell", al: "center", fs: 14, b: true }, { v: "{{MK}}", mk: true, mt: "cell", al: "center", fs: 14, b: true }],
    ],
  },
  {
    id: "table", label: "Таблица данных", icon: "▦", desc: "Маркер «Таблица»",
    rows: [
      [{ v: "Колонка A", b: true, bg: "#1a365d", fc: "#fff", al: "center" }, { v: "Колонка B", b: true, bg: "#1a365d", fc: "#fff", al: "center" }, { v: "Колонка C", b: true, bg: "#1a365d", fc: "#fff", al: "center" }, { v: "Колонка D", b: true, bg: "#1a365d", fc: "#fff", al: "center" }],
      [{ v: "{{MK}}", mk: true, mt: "table", cs: 4, rs: 4, bg: "#f7fafc", mOpts: { showHeaders: true, shiftRows: true } }, null, null, null],
      [null, null, null, null],
      [null, null, null, null],
      [null, null, null, null],
    ],
  },
  {
    id: "column", label: "Столбец данных", icon: "▥", desc: "Маркер «Столбец»",
    rows: [
      [{ v: "Категория", b: true, bg: "#edf2f7" }, { v: "{{MK}}", mk: true, mt: "column", b: true, bg: "#edf2f7", mOpts: { colIndex: 0 } }, { v: "", bg: "#edf2f7" }, { v: "", bg: "#edf2f7" }],
      [{ v: "Строка 1" }, { v: "" }, { v: "" }, { v: "" }],
      [{ v: "Строка 2" }, { v: "" }, { v: "" }, { v: "" }],
      [{ v: "Строка 3" }, { v: "" }, { v: "" }, { v: "" }],
    ],
  },
  {
    id: "formula", label: "Итоговая строка", icon: "Σ", desc: "Ячейки с формулами",
    rows: [
      [{ v: "ИТОГО:", b: true, bg: "#edf2f7" }, { v: "=SUM(B2:B10)", fm: true }, { v: "=SUM(C2:C10)", fm: true }, { v: "=SUM(D2:D10)", fm: true }],
    ],
  },
  {
    id: "separator", label: "Разделитель", icon: "─", desc: "Пустая строка",
    rows: [[{ v: "", cs: 4 }, null, null, null]],
  },
  {
    id: "footer", label: "Подвал", icon: "◻", desc: "Подписи, примечания",
    rows: [
      [{ v: "", cs: 4 }, null, null, null],
      [{ v: "Подготовил: _______________", cs: 2 }, null, { v: "Утвердил: _______________", cs: 2 }, null],
      [{ v: "Дата: «__» ________ 20__ г.", cs: 2, fc: "#718096" }, null, { v: "Дата: «__» ________ 20__ г.", cs: 2, fc: "#718096" }, null],
    ],
  },
];

/* ═══════════════════════════════════════════
   TEMPLATE LIBRARY
   ═══════════════════════════════════════════ */

const TEMPLATE_LIBRARY = [
  {
    id: "financial", name: "Финансовый отчёт", desc: "Выручка, расходы, прибыль, детализация",
    sheets: [{
      id: uid(), name: "Финансы", blocks: [
        { id: uid(), tid: "header", label: "Шапка отчёта", rows: clone(BLOCK_TEMPLATES[0].rows) },
        {
          id: uid(), tid: "custom", label: "Финансовые показатели",
          rows: [
            [{ v: "Показатель", b: true, bg: "#1a365d", fc: "#fff" }, { v: "План", b: true, bg: "#1a365d", fc: "#fff", al: "center" }, { v: "Факт", b: true, bg: "#1a365d", fc: "#fff", al: "center" }, { v: "% выполнения", b: true, bg: "#1a365d", fc: "#fff", al: "center" }],
            [{ v: "Выручка", b: true }, { v: "{{VIS_VAL_01}}", mk: true, mt: "cell", al: "right", mOpts: { rowIndex: 0, colIndex: 1 } }, { v: "{{VIS_VAL_02}}", mk: true, mt: "cell", al: "right", mOpts: { rowIndex: 0, colIndex: 2 } }, { v: "=C3/B3*100", fm: true, al: "right" }],
            [{ v: "Расходы" }, { v: "{{VIS_VAL_03}}", mk: true, mt: "cell", al: "right" }, { v: "{{VIS_VAL_04}}", mk: true, mt: "cell", al: "right" }, { v: "=C4/B4*100", fm: true, al: "right" }],
            [{ v: "Прибыль", b: true, bg: "#edf2f7" }, { v: "=B3-B4", fm: true, al: "right", bg: "#edf2f7" }, { v: "=C3-C4", fm: true, al: "right", bg: "#edf2f7" }, { v: "=C5/B5*100", fm: true, al: "right", bg: "#edf2f7" }],
          ],
        },
        { id: uid(), tid: "separator", label: "Разделитель", rows: clone(BLOCK_TEMPLATES[5].rows) },
        {
          id: uid(), tid: "custom", label: "Детализация",
          rows: [
            [{ v: "{{VIS_TBL_01}}", mk: true, mt: "table", cs: 4, rs: 5, bg: "#f7fafc", mOpts: { showHeaders: true, shiftRows: true } }, null, null, null],
            [null, null, null, null], [null, null, null, null], [null, null, null, null], [null, null, null, null],
          ],
        },
        { id: uid(), tid: "footer", label: "Подвал", rows: clone(BLOCK_TEMPLATES[6].rows) },
      ],
    }],
  },
  {
    id: "kpi_monthly", name: "Ежемесячный KPI", desc: "Сводка KPI по подразделениям",
    sheets: [{
      id: uid(), name: "KPI", blocks: [
        { id: uid(), tid: "header", label: "Шапка", rows: [
          [{ v: "ЕЖЕМЕСЯЧНЫЙ ОТЧЁТ ПО KPI", b: true, fs: 16, cs: 4, al: "center", bg: "#065f46", fc: "#fff" }, null, null, null],
          [{ v: "", cs: 4 }, null, null, null],
          [{ v: "Месяц:", b: true, bg: "#ecfdf5" }, { v: "{{VIS_VAL_01}}", mk: true, mt: "cell" }, { v: "Подразделение:", b: true, bg: "#ecfdf5" }, { v: "{{VIS_VAL_02}}", mk: true, mt: "cell" }],
          [{ v: "", cs: 4 }, null, null, null],
        ]},
        { id: uid(), tid: "custom", label: "KPI Grid", rows: [
          [{ v: "KPI", b: true, bg: "#065f46", fc: "#fff" }, { v: "Целевое", b: true, bg: "#065f46", fc: "#fff", al: "center" }, { v: "Фактическое", b: true, bg: "#065f46", fc: "#fff", al: "center" }, { v: "Статус", b: true, bg: "#065f46", fc: "#fff", al: "center" }],
          [{ v: "Выручка" }, { v: "{{VIS_VAL_03}}", mk: true, mt: "cell", al: "right" }, { v: "{{VIS_VAL_04}}", mk: true, mt: "cell", al: "right" }, { v: "" }],
          [{ v: "Кол-во клиентов" }, { v: "{{VIS_VAL_05}}", mk: true, mt: "cell", al: "right" }, { v: "{{VIS_VAL_06}}", mk: true, mt: "cell", al: "right" }, { v: "" }],
          [{ v: "NPS" }, { v: "{{VIS_VAL_07}}", mk: true, mt: "cell", al: "right" }, { v: "{{VIS_VAL_08}}", mk: true, mt: "cell", al: "right" }, { v: "" }],
          [{ v: "Срок обработки" }, { v: "{{VIS_VAL_09}}", mk: true, mt: "cell", al: "right" }, { v: "{{VIS_VAL_10}}", mk: true, mt: "cell", al: "right" }, { v: "" }],
        ]},
        { id: uid(), tid: "footer", label: "Подвал", rows: clone(BLOCK_TEMPLATES[6].rows) },
      ],
    }],
  },
  {
    id: "statistics", name: "Статистическая форма", desc: "Универсальная табличная форма",
    sheets: [{
      id: uid(), name: "Статистика", blocks: [
        { id: uid(), tid: "header", label: "Шапка", rows: [
          [{ v: "СТАТИСТИЧЕСКИЙ ОТЧЁТ", b: true, fs: 14, cs: 4, al: "center", bg: "#7c3aed", fc: "#fff" }, null, null, null],
          [{ v: "", cs: 4 }, null, null, null],
          [{ v: "Организация:", b: true }, { v: "", cs: 3 }, null, null],
          [{ v: "Период:", b: true }, { v: "{{VIS_VAL_01}} — {{VIS_VAL_02}}", mk: true, mt: "cell", cs: 3 }, null, null],
          [{ v: "", cs: 4 }, null, null, null],
        ]},
        { id: uid(), tid: "custom", label: "Основная таблица", rows: [
          [{ v: "{{VIS_TBL_01}}", mk: true, mt: "table", cs: 4, rs: 8, bg: "#faf5ff", mOpts: { showHeaders: true, shiftRows: true } }, null, null, null],
          [null, null, null, null], [null, null, null, null], [null, null, null, null],
          [null, null, null, null], [null, null, null, null], [null, null, null, null], [null, null, null, null],
        ]},
        { id: uid(), tid: "footer", label: "Подвал", rows: clone(BLOCK_TEMPLATES[6].rows) },
      ],
    }],
  },
];

/* ═══════════════════════════════════════════
   XLSX EXPORT (SheetJS)
   ═══════════════════════════════════════════ */

function exportXLSX(sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const allRows = [];
    const merges = [];
    let rowOff = 0;

    sheet.blocks.forEach((block) => {
      block.rows.forEach((row, ri) => {
        const r = [];
        row.forEach((c, ci) => {
          if (!c) { r.push(""); return; }
          r.push(c.fm ? c.v : (c.v || ""));
          if ((c.cs > 1 || c.rs > 1)) {
            merges.push({ s: { r: rowOff + ri, c: ci }, e: { r: rowOff + ri + (c.rs || 1) - 1, c: ci + (c.cs || 1) - 1 } });
          }
        });
        allRows.push(r);
      });
      rowOff += block.rows.length;
    });

    const ws = XLSX.utils.aoa_to_sheet(allRows);
    ws["!merges"] = merges;
    ws["!cols"] = [{ wch: 26 }, { wch: 22 }, { wch: 22 }, { wch: 22 }];

    // Write formulas
    rowOff = 0;
    sheet.blocks.forEach((block) => {
      block.rows.forEach((row, ri) => {
        row.forEach((c, ci) => {
          if (!c) return;
          const addr = XLSX.utils.encode_cell({ r: rowOff + ri, c: ci });
          if (c.fm && c.v?.startsWith("=")) {
            ws[addr] = { t: "s", f: c.v.slice(1), v: c.v };
          }
        });
      });
      rowOff += block.rows.length;
    });

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
  });
  XLSX.writeFile(wb, "visiology_template.xlsx");
}

/* ═══════════════════════════════════════════
   XLSX IMPORT
   ═══════════════════════════════════════════ */

function importXLSX(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const mg = ws["!merges"] || [];
    const maxC = Math.max(4, ...data.map((r) => r.length));
    const rows = data.map((row, ri) => {
      const cells = [];
      for (let ci = 0; ci < maxC; ci++) {
        const away = mg.some((m) => ri >= m.s.r && ri <= m.e.r && ci >= m.s.c && ci <= m.e.c && !(ri === m.s.r && ci === m.s.c));
        if (away) { cells.push(null); continue; }
        const merge = mg.find((m) => ri === m.s.r && ci === m.s.c);
        const val = row[ci] !== undefined ? String(row[ci]) : "";
        const cell = { v: val };
        if (val.startsWith("=")) cell.fm = true;
        if (/\{\{.*?\}\}/.test(val)) { cell.mk = true; cell.mt = val.includes("TBL") ? "table" : val.includes("COL") ? "column" : "cell"; }
        if (merge) {
          if (merge.e.c - merge.s.c > 0) cell.cs = merge.e.c - merge.s.c + 1;
          if (merge.e.r - merge.s.r > 0) cell.rs = merge.e.r - merge.s.r + 1;
        }
        cells.push(cell);
      }
      return cells;
    });
    return { id: uid(), name, blocks: [{ id: uid(), tid: "imported", label: "Импорт", rows }] };
  });
}

/* ═══════════════════════════════════════════
   CELL EDITOR
   ═══════════════════════════════════════════ */

function CellEditor({ cell, onSave, onClose }) {
  const [s, setS] = useState({
    v: cell?.v || "", b: cell?.b || false, bg: cell?.bg || "", fc: cell?.fc || "",
    mk: cell?.mk || false, mt: cell?.mt || "cell", fs: cell?.fs || 11, al: cell?.al || "left",
    fm: cell?.fm || false, mOpts: cell?.mOpts || {}, condRules: cell?.condRules || [],
  });
  const set = (k, val) => setS((p) => ({ ...p, [k]: val }));
  const setMO = (k, val) => setS((p) => ({ ...p, mOpts: { ...p.mOpts, [k]: val } }));
  const addCR = () => set("condRules", [...s.condRules, { op: "lt", threshold: "0", bg: "#fee2e2", fc: "#991b1b" }]);
  const rmCR = (i) => set("condRules", s.condRules.filter((_, j) => j !== i));
  const updCR = (i, k, val) => { const r = [...s.condRules]; r[i] = { ...r[i], [k]: val }; set("condRules", r); };

  return (
    <div style={overlay} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={modal}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 18 }}>
          <div style={secTitle}>Редактор ячейки</div>
          <button onClick={onClose} style={{ ...tinyBtn, fontSize: 16 }}>✕</button>
        </div>

        <label style={lbl}>Значение</label>
        <input value={s.v} onChange={(e) => set("v", e.target.value)} style={inp} placeholder="Текст, маркер или =формула" />

        <div style={{ display: "flex", gap: 14, marginTop: 10, alignItems: "center" }}>
          <label style={{ ...lbl, margin: 0, display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <input type="checkbox" checked={s.fm} onChange={(e) => set("fm", e.target.checked)} style={chk} />
            <span style={{ color: "#00DBFF" }}>Формула Excel</span>
          </label>
        </div>
        {s.fm && <div style={{ fontSize: 9, color: "#5A8A7A", marginTop: 3 }}>Пример: =SUM(B2:B10), =C3/B3*100</div>}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 }}>
          <div>
            <label style={lbl}>Фон</label>
            <div style={{ display: "flex", gap: 4 }}>
              <input type="color" value={s.bg || "#ffffff"} onChange={(e) => set("bg", e.target.value)} style={cpick} />
              <input value={s.bg} onChange={(e) => set("bg", e.target.value)} style={{ ...inp, flex: 1 }} placeholder="#hex" />
            </div>
          </div>
          <div>
            <label style={lbl}>Цвет текста</label>
            <div style={{ display: "flex", gap: 4 }}>
              <input type="color" value={s.fc || "#000000"} onChange={(e) => set("fc", e.target.value)} style={cpick} />
              <input value={s.fc} onChange={(e) => set("fc", e.target.value)} style={{ ...inp, flex: 1 }} placeholder="#hex" />
            </div>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginTop: 10 }}>
          <div><label style={lbl}>Размер</label><input type="number" value={s.fs} onChange={(e) => set("fs", +e.target.value)} style={inp} min={8} max={36} /></div>
          <div><label style={lbl}>Выравнивание</label><select value={s.al} onChange={(e) => set("al", e.target.value)} style={inp}><option value="left">← Лево</option><option value="center">Центр</option><option value="right">Право →</option></select></div>
          <div><label style={lbl}>Жирный</label><button onClick={() => set("b", !s.b)} style={{ ...inp, cursor: "pointer", fontWeight: s.b ? 900 : 400, background: s.b ? "#03FF94" : "#000D0A", color: s.b ? "#000D0A" : "#5A8A7A", textAlign: "center", border: s.b ? "1px solid #03FF94" : "1px solid #007359" }}>{s.b ? "B ✓" : "B"}</button></div>
        </div>

        {/* Marker */}
        <div style={{ marginTop: 14, padding: 12, background: "#0A1A14", borderRadius: 9, border: "1px solid #007359" }}>
          <label style={{ ...lbl, margin: 0, display: "flex", alignItems: "center", gap: 8, cursor: "pointer", marginBottom: s.mk ? 10 : 0 }}>
            <input type="checkbox" checked={s.mk} onChange={(e) => set("mk", e.target.checked)} style={chk} />
            <span style={{ color: "#03FF94", fontWeight: 700 }}>МАРКЕР VISIOLOGY</span>
          </label>
          {s.mk && (<>
            <select value={s.mt} onChange={(e) => set("mt", e.target.value)} style={{ ...inp, marginBottom: 10 }}>
              {MARKER_TYPES.map((t) => <option key={t.value} value={t.value}>{t.icon} {t.label} — {t.desc}</option>)}
            </select>
            {s.mt === "cell" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <div><label style={lbl}>Индекс строки</label><input type="number" value={s.mOpts.rowIndex ?? ""} onChange={(e) => setMO("rowIndex", e.target.value === "" ? undefined : +e.target.value)} style={inp} min={0} placeholder="0" /></div>
                <div><label style={lbl}>Индекс столбца</label><input type="number" value={s.mOpts.colIndex ?? ""} onChange={(e) => setMO("colIndex", e.target.value === "" ? undefined : +e.target.value)} style={inp} min={0} placeholder="0" /></div>
              </div>
            )}
            {s.mt === "table" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11, color: "#B4FFDF", cursor: "pointer" }}>
                  <input type="checkbox" checked={s.mOpts.showHeaders ?? true} onChange={(e) => setMO("showHeaders", e.target.checked)} style={chk} />
                  Отображать заголовки столбцов
                </label>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11, color: "#B4FFDF", cursor: "pointer" }}>
                  <input type="checkbox" checked={s.mOpts.shiftRows ?? true} onChange={(e) => setMO("shiftRows", e.target.checked)} style={chk} />
                  Смещать строки ниже (динамическая высота)
                </label>
              </div>
            )}
            {s.mt === "column" && (
              <div><label style={lbl}>Индекс столбца данных</label><input type="number" value={s.mOpts.colIndex ?? 0} onChange={(e) => setMO("colIndex", +e.target.value)} style={inp} min={0} /></div>
            )}
          </>)}
        </div>

        {/* Conditional formatting */}
        <div style={{ marginTop: 12, padding: 12, background: "#0A1A14", borderRadius: 9, border: "1px solid #007359" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: s.condRules.length ? 8 : 0 }}>
            <span style={{ ...lbl, margin: 0, color: "#00DBFF" }}>УСЛОВНОЕ ФОРМАТИРОВАНИЕ</span>
            <button onClick={addCR} style={{ ...tinyBtn, color: "#00DBFF", border: "1px solid #007359", borderRadius: 5, padding: "2px 8px" }}>+ правило</button>
          </div>
          {s.condRules.map((rule, i) => (
            <div key={i} style={{ display: "flex", gap: 5, alignItems: "center", marginBottom: 5 }}>
              <select value={rule.op} onChange={(e) => updCR(i, "op", e.target.value)} style={{ ...inp, width: 110, fontSize: 10 }}>
                {COND_OPERATORS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
              <input value={rule.threshold} onChange={(e) => updCR(i, "threshold", e.target.value)} style={{ ...inp, width: 60 }} />
              <input type="color" value={rule.bg} onChange={(e) => updCR(i, "bg", e.target.value)} style={cpick} title="Фон" />
              <input type="color" value={rule.fc} onChange={(e) => updCR(i, "fc", e.target.value)} style={cpick} title="Текст" />
              <button onClick={() => rmCR(i)} style={{ ...tinyBtn, color: "#ef4444" }}>✕</button>
            </div>
          ))}
        </div>

        <div style={{ display: "flex", gap: 8, marginTop: 18, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={btnSec}>Отмена</button>
          <button onClick={() => onSave({ ...cell, ...s })} style={btnPri}>Сохранить</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════
   MARKER PANEL
   ═══════════════════════════════════════════ */

function MarkerPanel({ sheets }) {
  const markers = [];
  sheets.forEach((sh) => sh.blocks.forEach((bl) => bl.rows.forEach((row) => row.forEach((c) => {
    if (c?.mk) markers.push({ sheet: sh.name, type: c.mt, value: c.v, opts: c.mOpts });
  }))));
  return (
    <div style={{ background: "#0A1A14", borderRadius: 8, padding: 10, border: "1px solid #007359" }}>
      <div style={{ ...secTitle, fontSize: 9, color: "#03FF94", marginBottom: 6 }}>МАРКЕРЫ ({markers.length})</div>
      {markers.length === 0 ? <div style={{ color: "#3D6457", fontSize: 9 }}>Нет маркеров</div> : (
        <div style={{ maxHeight: 160, overflowY: "auto" }}>
          {markers.map((m, i) => {
            const t = MARKER_TYPES.find((x) => x.value === m.type) || MARKER_TYPES[0];
            return (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 5, padding: "4px 0", borderBottom: "1px solid #007359", fontSize: 9 }}>
                <span style={{ padding: "1px 5px", borderRadius: 4, fontSize: 8, fontWeight: 700, background: t.bg, color: t.color, border: `1px solid ${t.color}44` }}>{t.value.slice(0, 3).toUpperCase()}</span>
                <span style={{ color: "#B4FFDF", fontFamily: "'TT Commons Pro','Inter',monospace", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.value}</span>
                {m.opts && Object.keys(m.opts).length > 0 && <span style={{ color: "#007359", fontSize: 7 }}>⚙</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════
   PREVIEW CELL
   ═══════════════════════════════════════════ */

function PCell({ cell, onClick }) {
  if (!cell) return <td style={{ ...tdS, background: "#F0FFF8" }} />;
  const isM = cell.mk, isF = cell.fm;
  return (
    <td style={{
      ...tdS,
      fontWeight: cell.b ? 700 : 400, fontSize: cell.fs || 11, textAlign: cell.al || "left",
      background: isM ? "#03FF9412" : (cell.bg || "#fff"),
      color: isM ? "#007359" : isF ? "#00DBFF" : (cell.fc || "#1a202c"),
      border: isM ? "1.5px dashed #03FF94" : isF ? "1.5px dashed #00DBFF" : "1px solid #CEF3F2",
      cursor: "pointer", position: "relative",
    }} colSpan={cell.cs || 1} rowSpan={cell.rs || 1} onClick={onClick}>
      {isM && <span style={{ position: "absolute", top: 0, right: 2, fontSize: 7, color: "#03FF94", fontWeight: 800 }}>{MARKER_TYPES.find((t) => t.value === cell.mt)?.icon || "◎"}</span>}
      {isF && <span style={{ position: "absolute", top: 0, left: 2, fontSize: 7, color: "#00DBFF", fontWeight: 800 }}>ƒ</span>}
      {cell.condRules?.length > 0 && <span style={{ position: "absolute", bottom: 0, right: 2, fontSize: 5, color: "#03FF94" }}>◆</span>}
      <span style={{ fontFamily: (isM || isF) ? "'TT Commons Pro','Inter',monospace" : "inherit", fontSize: (isM || isF) ? Math.min(cell.fs || 11, 10) : (cell.fs || 11) }}>
        {cell.v || "\u00A0"}
      </span>
    </td>
  );
}

/* ═══════════════════════════════════════════
   TEMPLATE CHOOSER
   ═══════════════════════════════════════════ */

function TemplateChooser({ onSelect, onClose }) {
  return (
    <div style={overlay} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ ...modal, width: 520 }}>
        <div style={{ ...secTitle, marginBottom: 16 }}>Библиотека шаблонов</div>
        {TEMPLATE_LIBRARY.map((tpl) => (
          <div key={tpl.id} onClick={() => onSelect(tpl)} style={{
            padding: 14, background: "#0A1A14", border: "1px solid #007359", borderRadius: 10, cursor: "pointer", marginBottom: 8,
            transition: "border-color .15s, box-shadow .15s",
          }}
          onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#03FF94"; e.currentTarget.style.boxShadow = "0 0 12px rgba(3,255,148,0.15)"; }}
          onMouseLeave={(e) => { e.currentTarget.style.borderColor = "#007359"; e.currentTarget.style.boxShadow = "none"; }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#FFFFFF" }}>{tpl.name}</div>
            <div style={{ fontSize: 10, color: "#B4FFDF", marginTop: 3 }}>{tpl.desc}</div>
            <div style={{ fontSize: 9, color: "#5A8A7A", marginTop: 5 }}>
              {tpl.sheets.length} лист · {tpl.sheets.reduce((a, s) => a + s.blocks.length, 0)} блоков · {tpl.sheets.reduce((a, s) => a + s.blocks.reduce((b, bl) => b + bl.rows.flat().filter((c) => c?.mk).length, 0), 0)} маркеров
            </div>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}><button onClick={onClose} style={btnSec}>Закрыть</button></div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════
   ONBOARDING MODAL
   ═══════════════════════════════════════════ */

function OnboardingModal({ onStart, onTemplate }) {
  const steps = [
    { n: "1", color: "#03FF94", title: "Добавьте блоки", desc: "Выбирайте из палитры слева — Шапка, KPI, Таблица, Формулы" },
    { n: "2", color: "#00DBFF", title: "Настройте маркеры", desc: "Кликайте на ячейки и привязывайте маркеры {{VIS_*}} к данным Visiology" },
    { n: "3", color: "#B4FFDF", title: "Экспортируйте", desc: "Скачайте готовый .xlsx — он будет заполнен реальными данными платформы" },
  ];
  return (
    <div style={overlay}>
      <div style={{ ...modal, width: 460, textAlign: "center" }}>
        {/* Logo */}
        <div style={{ width: 58, height: 58, borderRadius: 16, background: "linear-gradient(135deg, #03FF94, #00DBFF)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, fontWeight: 900, color: "#000D0A", margin: "0 auto 16px", boxShadow: "0 0 24px rgba(3,255,148,0.35)" }}>V</div>
        <div style={{ fontSize: 18, fontWeight: 800, color: "#FFFFFF", marginBottom: 4, letterSpacing: -0.3 }}>Visiology Report Builder</div>
        <div style={{ fontSize: 11, color: "#B4FFDF", marginBottom: 24 }}>Конструктор шаблонов регламентных отчётов</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 24, textAlign: "left" }}>
          {steps.map((s) => (
            <div key={s.n} style={{ display: "flex", gap: 13, padding: "11px 14px", background: "#0A1A14", borderRadius: 10, border: "1px solid #007359", alignItems: "flex-start" }}>
              <div style={{ width: 28, height: 28, borderRadius: 7, background: s.color + "1E", border: "1px solid " + s.color + "66", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: s.color, flexShrink: 0, fontWeight: 800 }}>{s.n}</div>
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#FFFFFF", marginBottom: 2 }}>{s.title}</div>
                <div style={{ fontSize: 10, color: "#B4FFDF", lineHeight: 1.6 }}>{s.desc}</div>
              </div>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onTemplate} style={{ ...btnPri, flex: 1 }}>📚 Выбрать шаблон</button>
          <button onClick={onStart} style={{ ...btnSec, flex: 1 }}>Начать с нуля →</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════
   MAIN APP
   ═══════════════════════════════════════════ */

export default function App() {
  const [sheets, _setSheets] = useState(() => [{ id: uid(), name: "Лист 1", blocks: [] }]);
  const [tab, setTab] = useState(0);
  const [editing, setEditing] = useState(null);
  const [showTpl, setShowTpl] = useState(false);
  const [showImp, setShowImp] = useState(false);
  const [impText, setImpText] = useState("");
  const [note, setNote] = useState(null);
  const [dragBlock, setDragBlock] = useState(null);
  const [dragOver, setDragOver] = useState(null);
  const [showOnboarding, setShowOnboarding] = useState(
    () => typeof window !== "undefined" && !localStorage.getItem("vrb_onboarded")
  );
  const [histStats, setHistStats] = useState({ idx: 0, total: 1 });
  const fRef = useRef(null);
  const historyRef = useRef(null);
  const histIdxRef = useRef(0);
  if (historyRef.current === null) { historyRef.current = [clone(sheets)]; }

  const setAndPush = useCallback((updater) => {
    _setSheets((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      const h = historyRef.current.slice(0, histIdxRef.current + 1);
      h.push(clone(next));
      if (h.length > 50) h.shift();
      historyRef.current = h;
      histIdxRef.current = h.length - 1;
      setHistStats({ idx: histIdxRef.current, total: h.length });
      return next;
    });
  }, []);

  const flash = (m, ok = true) => { setNote({ m, ok }); setTimeout(() => setNote(null), 2500); };

  const undo = () => {
    if (histIdxRef.current > 0) {
      histIdxRef.current--;
      const newSheets = historyRef.current[histIdxRef.current];
      _setSheets(clone(newSheets));
      setTab((t) => Math.min(t, newSheets.length - 1));
      setHistStats({ idx: histIdxRef.current, total: historyRef.current.length });
      flash("↩ Отменено");
    }
  };

  const redo = () => {
    if (histIdxRef.current < historyRef.current.length - 1) {
      histIdxRef.current++;
      const newSheets = historyRef.current[histIdxRef.current];
      _setSheets(clone(newSheets));
      setTab((t) => Math.min(t, newSheets.length - 1));
      setHistStats({ idx: histIdxRef.current, total: historyRef.current.length });
      flash("↪ Повторено");
    }
  };

  useEffect(() => {
    const onKey = (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      if (e.key === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
      if (e.key === "y" || (e.key === "z" && e.shiftKey)) { e.preventDefault(); redo(); }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const sh = sheets[tab] || sheets[0];
  const updSh = useCallback((i, fn) => setAndPush((p) => p.map((s, j) => j === i ? fn(s) : s)), [setAndPush]);

  const addBlock = (tid) => {
    const tpl = BLOCK_TEMPLATES.find((t) => t.id === tid);
    if (!tpl) return;
    const mc = sh.blocks.reduce((c, b) => c + b.rows.flat().filter((x) => x?.mk).length, 0);
    const rows = clone(tpl.rows);
    let mi = mc;
    rows.forEach((r) => r.forEach((c) => { if (c?.mk && c.v === "{{MK}}") c.v = markerName(c.mt || "cell", ++mi); }));
    updSh(tab, (s) => ({ ...s, blocks: [...s.blocks, { id: uid(), tid: tpl.id, label: tpl.label, rows }] }));
  };

  const rmBlock = (i) => updSh(tab, (s) => ({ ...s, blocks: s.blocks.filter((_, j) => j !== i) }));
  const moveBlock = (f, t) => { if (t < 0 || t >= sh.blocks.length) return; updSh(tab, (s) => { const b = [...s.blocks]; const [m] = b.splice(f, 1); b.splice(t, 0, m); return { ...s, blocks: b }; }); };
  const addRow = (bi) => updSh(tab, (s) => { const b = clone(s.blocks); const cols = Math.max(4, ...b[bi].rows.map((r) => r.length)); b[bi].rows.push(Array.from({ length: cols }, () => ({ v: "" }))); return { ...s, blocks: b }; });
  const addCol = (bi) => updSh(tab, (s) => { const b = clone(s.blocks); b[bi].rows = b[bi].rows.map((r) => [...r, { v: "" }]); return { ...s, blocks: b }; });
  const delRow = (bi, ri) => updSh(tab, (s) => { const b = clone(s.blocks); if (b[bi].rows.length <= 1) return s; b[bi].rows.splice(ri, 1); return { ...s, blocks: b }; });
  const updateCell = (bi, ri, ci, d) => { updSh(tab, (s) => { const b = clone(s.blocks); b[bi].rows[ri][ci] = d; return { ...s, blocks: b }; }); setEditing(null); };
  const addSheet = () => { setAndPush((p) => [...p, { id: uid(), name: `Лист ${p.length + 1}`, blocks: [] }]); setTab(sheets.length); };
  const rmSheet = (i) => { if (sheets.length <= 1) return; setAndPush((p) => p.filter((_, j) => j !== i)); if (tab >= i && tab > 0) setTab(tab - 1); };

  const applyTpl = (tpl) => { setAndPush(clone(tpl.sheets).map((s) => ({ ...s, id: uid() }))); setTab(0); setShowTpl(false); flash(`«${tpl.name}» загружен`); };

  const handleFile = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader();
    r.onload = (ev) => { try { const imp = importXLSX(ev.target.result); if (imp?.length) { setAndPush(imp); setTab(0); flash("XLSX импортирован"); } else flash("Ошибка чтения", false); } catch (err) { flash("Ошибка: " + err.message, false); } };
    r.readAsArrayBuffer(f); e.target.value = "";
  };

  const handleImpJSON = () => { try { const d = JSON.parse(impText); if (d.sheets) { setAndPush(d.sheets); setTab(0); setShowImp(false); setImpText(""); flash("JSON импортирован"); } else flash("Неверный формат", false); } catch { flash("Ошибка JSON", false); } };

  const palDrag = (e, tid) => e.dataTransfer.setData("tplId", tid);
  const prevDrop = (e) => { e.preventDefault(); const tid = e.dataTransfer.getData("tplId"); if (tid) addBlock(tid); };

  const totalRows = useMemo(() => sh.blocks.reduce((a, b) => a + b.rows.length, 0), [sh.blocks]);

  return (
    <div style={{ minHeight: "100vh", background: "#000D0A", fontFamily: "'TT Commons Pro','Inter',-apple-system,BlinkMacSystemFont,sans-serif", color: "#FFFFFF" }}>
      {note && <div style={{ position: "fixed", top: 14, right: 14, zIndex: 9999, padding: "10px 18px", borderRadius: 9, background: note.ok ? "#007359" : "#7f1d1d", color: "#FFFFFF", fontSize: 11, fontWeight: 600, boxShadow: "0 0 20px rgba(3,255,148,0.2), 0 8px 28px rgba(0,0,0,0.5)", animation: "fi .2s ease", border: note.ok ? "1px solid #03FF9466" : "none" }}>{note.m}</div>}
      {showTpl && <TemplateChooser onSelect={applyTpl} onClose={() => setShowTpl(false)} />}
      {showOnboarding && (
        <OnboardingModal
          onStart={() => { setShowOnboarding(false); localStorage.setItem("vrb_onboarded", "1"); }}
          onTemplate={() => { setShowOnboarding(false); localStorage.setItem("vrb_onboarded", "1"); setShowTpl(true); }}
        />
      )}
      {showImp && <div style={overlay} onClick={() => setShowImp(false)}><div onClick={(e) => e.stopPropagation()} style={modal}><div style={{ ...secTitle, marginBottom: 12 }}>Импорт JSON</div><textarea value={impText} onChange={(e) => setImpText(e.target.value)} style={{ ...inp, height: 160, resize: "vertical", fontFamily: "'TT Commons Pro','Inter',monospace", fontSize: 10 }} placeholder='{"sheets":[...]}' /><div style={{ display: "flex", gap: 8, marginTop: 12, justifyContent: "flex-end" }}><button onClick={() => setShowImp(false)} style={btnSec}>Отмена</button><button onClick={handleImpJSON} style={btnPri}>Импорт</button></div></div></div>}
      {editing && <CellEditor cell={sh.blocks[editing.b]?.rows[editing.r]?.[editing.c]} onSave={(d) => updateCell(editing.b, editing.r, editing.c, d)} onClose={() => setEditing(null)} />}
      <input ref={fRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />

      {/* HEADER */}
      <header style={{ padding: "11px 22px", borderBottom: "1px solid #007359", display: "flex", alignItems: "center", justifyContent: "space-between", background: "rgba(0,13,10,0.95)", backdropFilter: "blur(12px)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "linear-gradient(135deg, #03FF94, #00DBFF)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 900, color: "#000D0A", boxShadow: "0 0 12px rgba(3,255,148,0.4)" }}>V</div>
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#FFFFFF", letterSpacing: -0.3 }}>Visiology Report Builder</div>
            <div style={{ fontSize: 8, color: "#5A8A7A", letterSpacing: 1, textTransform: "uppercase" }}>Конструктор шаблонов регламентных отчётов</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 5, alignItems: "center" }}>
          <button onClick={undo} disabled={!histStats.idx} title="Отменить (Ctrl+Z)" style={{ ...hBtn, opacity: histStats.idx ? 1 : 0.32, fontSize: 14, padding: "4px 9px" }}>↩</button>
          <button onClick={redo} disabled={histStats.idx >= histStats.total - 1} title="Повторить (Ctrl+Y)" style={{ ...hBtn, opacity: histStats.idx < histStats.total - 1 ? 1 : 0.32, fontSize: 14, padding: "4px 9px" }}>↪</button>
          <div style={{ width: 1, height: 18, background: "#007359", margin: "0 3px" }} />
          <button onClick={() => setShowTpl(true)} style={{ ...hBtn, background: "#1C3D36", color: "#03FF94", borderColor: "#007359" }}>📚 Шаблоны</button>
          <button onClick={() => fRef.current?.click()} style={hBtn}>📂 XLSX</button>
          <button onClick={() => setShowImp(true)} style={hBtn}>↓ JSON</button>
          <div style={{ width: 1, height: 18, background: "#007359", margin: "0 3px" }} />
          <button onClick={() => { const j = JSON.stringify({ version: "1.0", app: "vrb", sheets }, null, 2); const b = new Blob([j], { type: "application/json" }); const u = URL.createObjectURL(b); const a = document.createElement("a"); a.href = u; a.download = "template.json"; a.click(); URL.revokeObjectURL(u); flash("JSON ↑"); }} style={hBtn}>↑ JSON</button>
          <button onClick={() => { exportXLSX(sheets); flash("XLSX экспортирован"); }} style={{ ...hBtn, background: "#03FF94", color: "#000D0A", border: "none", fontWeight: 800, boxShadow: "0 0 10px rgba(3,255,148,0.3)" }}>↑ XLSX</button>
        </div>
      </header>

      {/* GRID */}
      <div style={{ display: "grid", gridTemplateColumns: "268px 1fr", height: "calc(100vh - 53px)" }}>
        {/* SIDEBAR */}
        <aside style={{ borderRight: "1px solid #007359", overflowY: "auto", background: "#0A1A14" }}>
          <div style={{ padding: 12 }}>
            <div style={{ ...secTitle, fontSize: 9, marginBottom: 8 }}>БЛОКИ</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
              {BLOCK_TEMPLATES.map((t) => (
                <div key={t.id} draggable onDragStart={(e) => palDrag(e, t.id)} onClick={() => addBlock(t.id)}
                  style={{ padding: "7px 10px", borderRadius: 8, background: "#1C3D36", border: "1px solid #007359", cursor: "grab", display: "flex", alignItems: "center", gap: 8, transition: "all .15s" }}
                  onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#03FF94"; e.currentTarget.style.boxShadow = "0 0 10px rgba(3,255,148,0.12)"; }}
                  onMouseLeave={(e) => { e.currentTarget.style.borderColor = "#007359"; e.currentTarget.style.boxShadow = "none"; }}>
                  <span style={{ fontSize: 15, width: 22, textAlign: "center", color: "#03FF94", opacity: 0.7 }}>{t.icon}</span>
                  <div><div style={{ fontSize: 10, fontWeight: 600, color: "#FFFFFF" }}>{t.label}</div><div style={{ fontSize: 8, color: "#5A8A7A" }}>{t.desc}</div></div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ padding: "8px 12px", borderTop: "1px solid #007359" }}>
            <div style={{ ...secTitle, fontSize: 9, marginBottom: 6 }}>СТРУКТУРА «{sh.name}»</div>
            {sh.blocks.length === 0 ? <div style={{ color: "#3D6457", fontSize: 9, textAlign: "center", padding: "14px 0" }}>Пусто</div> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                {sh.blocks.map((bl, bi) => (
                  <div key={bl.id} draggable
                    onDragStart={() => setDragBlock(bi)} onDragOver={(e) => { e.preventDefault(); setDragOver(bi); }}
                    onDrop={() => { if (dragBlock !== null && dragBlock !== bi) moveBlock(dragBlock, bi); setDragBlock(null); setDragOver(null); }}
                    onDragEnd={() => { setDragBlock(null); setDragOver(null); }}
                    style={{
                      padding: "5px 8px", borderRadius: 6, background: dragBlock === bi ? "#1C3D36" : "#0A1A14",
                      border: dragOver === bi ? "1px solid #03FF94" : "1px solid #007359",
                      display: "flex", alignItems: "center", justifyContent: "space-between", fontSize: 9, cursor: "grab",
                      transition: "border-color .1s",
                    }}>
                    <span style={{ color: "#B4FFDF", display: "flex", alignItems: "center", gap: 3 }}>
                      <span style={{ opacity: 0.3 }}>⠿</span>
                      {BLOCK_TEMPLATES.find((t) => t.id === bl.tid)?.icon || "◇"} {bl.label}
                      <span style={{ color: "#3D6457" }}>({bl.rows.length})</span>
                    </span>
                    <div style={{ display: "flex", gap: 2 }}>
                      <button onClick={() => addRow(bi)} style={tinyBtn}>+r</button>
                      <button onClick={() => addCol(bi)} style={tinyBtn}>+c</button>
                      <button onClick={() => rmBlock(bi)} style={{ ...tinyBtn, color: "#ef4444" }}>✕</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div style={{ padding: "8px 12px", borderTop: "1px solid #007359" }}><MarkerPanel sheets={sheets} /></div>

          <div style={{ padding: "8px 12px", borderTop: "1px solid #007359" }}>
            <div style={{ ...secTitle, fontSize: 9, marginBottom: 4, color: "#5A8A7A" }}>ПЕЧАТЬ A4</div>
            <div style={{ fontSize: 9, color: "#5A8A7A" }}>≈ {totalRows} строк · {Math.max(1, Math.ceil(totalRows / A4_ROWS_APPROX))} стр.</div>
            <div style={{ marginTop: 4, height: 3, borderRadius: 2, background: "#1C3D36", overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${Math.min(100, (totalRows / A4_ROWS_APPROX) * 100)}%`, background: totalRows > A4_ROWS_APPROX ? "#00DBFF" : "#03FF94", borderRadius: 2, transition: "width .3s" }} />
            </div>
          </div>
        </aside>

        {/* PREVIEW */}
        <main style={{ overflowY: "auto", padding: 22, background: "#F0FFF8" }} onDragOver={(e) => e.preventDefault()} onDrop={prevDrop}>
          <div style={{ display: "flex", alignItems: "center", gap: 2, marginBottom: 0 }}>
            {sheets.map((s, i) => (
              <div key={s.id} style={{ display: "flex", alignItems: "center" }}>
                <button onClick={() => setTab(i)} style={{
                  padding: "5px 14px", fontSize: 10, fontWeight: tab === i ? 700 : 400,
                  background: tab === i ? "#FFFFFF" : "#CEF3F2",
                  color: tab === i ? "#000D0A" : "#5A8A7A",
                  border: tab === i ? "1px solid #007359" : "1px solid transparent",
                  borderBottom: tab === i ? "1px solid #FFFFFF" : "1px solid #007359",
                  borderRadius: "6px 6px 0 0", cursor: "pointer", transition: "all .12s",
                }}>{s.name}</button>
                {sheets.length > 1 && <button onClick={() => rmSheet(i)} style={{ ...tinyBtn, color: "#5A8A7A", fontSize: 8 }}>✕</button>}
              </div>
            ))}
            <button onClick={addSheet} style={{ background: "none", border: "1px dashed #007359", borderRadius: "6px 6px 0 0", padding: "5px 11px", fontSize: 10, color: "#5A8A7A", cursor: "pointer" }}>+</button>
          </div>

          <div style={{ background: "#fff", borderRadius: "0 6px 6px 6px", boxShadow: "0 2px 16px rgba(0,115,89,0.08)", padding: 24, minHeight: 500, border: "1px solid #007359", position: "relative" }}>
            {totalRows > A4_ROWS_APPROX && Array.from({ length: Math.floor(totalRows / A4_ROWS_APPROX) }, (_, i) => (
              <div key={i} style={{ position: "absolute", left: 0, right: 0, top: (i + 1) * A4_ROWS_APPROX * 22 + 24, borderTop: "2px dashed #007359", zIndex: 5 }}>
                <span style={{ position: "absolute", right: 0, top: -12, fontSize: 8, color: "#007359", background: "#fff", padding: "0 3px" }}>стр. {i + 2}</span>
              </div>
            ))}

            {sh.blocks.length === 0 ? (
              <div style={{ padding: "40px 24px" }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#007359", marginBottom: 4, textAlign: "center" }}>С чего начать?</div>
                <div style={{ fontSize: 10, color: "#5A8A7A", marginBottom: 20, textAlign: "center" }}>Выберите вариант или перетащите блок из палитры слева</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, maxWidth: 480, margin: "0 auto 22px" }}>
                  {[
                    { icon: "⊞", color: "#00DBFF", title: "Добавьте блок", desc: "Кликните или перетащите блок из палитры слева" },
                    { icon: "◎", color: "#03FF94", title: "Настройте маркеры", desc: "Кликните ячейку → выберите тип маркера {{VIS_*}}" },
                    { icon: "↑", color: "#B4FFDF", title: "Экспортируйте", desc: "Скачайте .xlsx — Visiology заполнит данными" },
                  ].map((s, i) => (
                    <div key={i} style={{ background: "#F0FFF8", border: `1.5px solid ${s.color}55`, borderRadius: 12, padding: "14px 10px", textAlign: "center" }}>
                      <div style={{ fontSize: 22, marginBottom: 6, color: s.color }}>{s.icon}</div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#007359", marginBottom: 5 }}>{s.title}</div>
                      <div style={{ fontSize: 9, color: "#5A8A7A", lineHeight: 1.5 }}>{s.desc}</div>
                    </div>
                  ))}
                </div>
                <div style={{ textAlign: "center", display: "flex", gap: 8, justifyContent: "center" }}>
                  <button onClick={() => setShowTpl(true)} style={{ ...btnPri, fontSize: 11, padding: "7px 18px" }}>📚 Открыть готовый шаблон</button>
                  <button onClick={() => setShowOnboarding(true)} style={{ ...btnSec, fontSize: 10, padding: "7px 14px" }}>? Как это работает</button>
                </div>
              </div>
            ) : sh.blocks.map((bl, bi) => {
              const rOff = sh.blocks.slice(0, bi).reduce((a, b) => a + b.rows.length, 0);
              return (
                <table key={bl.id} style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed" }}>
                  <tbody>
                    {bl.rows.map((row, ri) => (
                      <tr key={ri}>
                        <td style={{ width: 24, padding: "1px 3px", fontSize: 7, color: "#007359", textAlign: "right", verticalAlign: "middle", border: "none", background: "#F0FFF8", userSelect: "none" }}>{rOff + ri + 1}</td>
                        {row.map((cell, ci) => cell === null ? null : <PCell key={ci} cell={cell} onClick={() => setEditing({ b: bi, r: ri, c: ci })} />)}
                        <td style={{ width: 18, border: "none", background: "#F0FFF8" }}><button onClick={() => delRow(bi, ri)} style={{ ...tinyBtn, fontSize: 8, color: "#007359", opacity: 0.5 }}>✕</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              );
            })}
          </div>

          <div style={{ marginTop: 10, display: "flex", gap: 14, flexWrap: "wrap", padding: "0 2px", fontSize: 9, color: "#5A8A7A" }}>
            <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ display: "inline-block", width: 10, height: 10, background: "#03FF9412", border: "1.5px dashed #03FF94", borderRadius: 2 }} />Маркер</span>
            <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ display: "inline-block", width: 10, height: 10, background: "#00DBFF12", border: "1.5px dashed #00DBFF", borderRadius: 2 }} />Формула</span>
            <span style={{ color: "#03FF94" }}>◆ Усл. форматирование</span>
            <span style={{ color: "#5A8A7A" }}>Кликните на ячейку для редактирования</span>
          </div>
        </main>
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:'TT Commons Pro','Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:#007359;border-radius:3px}
        ::-webkit-scrollbar-thumb:hover{background:#03FF94}
        @keyframes fi{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
        @keyframes glow-in{from{box-shadow:0 0 0 rgba(3,255,148,0)}to{box-shadow:0 0 16px rgba(3,255,148,0.25)}}
        button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #03FF94;outline-offset:2px}
      `}</style>
    </div>
  );
}

/* ─── BRAND TOKEN MAP (FatData Brandbook) ───────────────────────────────────
   #000D0A  Black Green     — page background
   #0A1A14  Deep Forest ×   — sidebar / input background
   #1C3D36  Deep Forest     — card / modal surface
   #007359  Emerald Core    — borders, secondary accents
   #03FF94  Neon Green      — primary accent, CTA, markers
   #00DBFF  Electric Cyan   — secondary accent, formulas
   #B4FFDF  Soft Mint       — secondary text, subtle accents
   #CEF3F2  Pale Mint Gray  — light surfaces, table borders
   #FFFFFF  Pure White      — primary text
   Font     TT Commons Pro → Inter fallback
────────────────────────────────────────────────────────────────────────── */
const overlay={position:"fixed",inset:0,background:"rgba(0,13,10,0.78)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000};
const modal={background:"#1C3D36",borderRadius:16,padding:24,width:420,border:"1px solid #007359",boxShadow:"0 0 48px rgba(3,255,148,0.07), 0 24px 64px rgba(0,0,0,0.7)",maxHeight:"90vh",overflowY:"auto"};
const secTitle={fontSize:10,fontWeight:700,color:"#03FF94",letterSpacing:1.8,textTransform:"uppercase"};
const lbl={display:"block",fontSize:9,fontWeight:600,color:"#B4FFDF",letterSpacing:0.5,marginBottom:3,textTransform:"uppercase"};
const inp={width:"100%",padding:"7px 10px",background:"#000D0A",border:"1px solid #007359",borderRadius:7,color:"#FFFFFF",fontSize:11,outline:"none",fontFamily:"'TT Commons Pro','Inter',-apple-system,sans-serif"};
const chk={width:13,height:13,accentColor:"#03FF94"};
const cpick={width:28,height:26,border:"none",borderRadius:4,cursor:"pointer",padding:0};
const btnPri={padding:"7px 18px",background:"#03FF94",color:"#000D0A",border:"none",borderRadius:8,fontSize:11,fontWeight:700,cursor:"pointer",letterSpacing:0.2};
const btnSec={padding:"7px 18px",background:"#1C3D36",color:"#B4FFDF",border:"1px solid #007359",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer"};
const hBtn={padding:"5px 11px",background:"#1C3D36",color:"#B4FFDF",border:"1px solid #007359",borderRadius:6,fontSize:10,fontWeight:600,cursor:"pointer"};
const tinyBtn={background:"none",border:"none",color:"#5A8A7A",fontSize:9,cursor:"pointer",padding:"1px 4px",fontFamily:"'TT Commons Pro','Inter',monospace"};
const tdS={padding:"4px 7px",fontSize:11,verticalAlign:"middle",border:"1px solid #CEF3F2",fontFamily:"'TT Commons Pro','Inter',-apple-system,sans-serif",minWidth:44,transition:"background .1s"};
