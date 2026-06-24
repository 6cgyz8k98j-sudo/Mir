const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, WidthType, BorderStyle,
  Header, Footer, PageNumber, ShadingType, convertInchesToTwip
} = require("docx");
const fs = require("fs");
const path = require("path");

// ─── Данные ────────────────────────────────────────────────────────────────

const PERIOD = "Май 2025";
const GENERATED = new Date().toLocaleDateString("ru-RU");

const CORP_BLUE = "1F4E79";
const ACCENT    = "2E75B6";
const WHITE     = "FFFFFF";
const GREEN_BG  = "E2EFDA";
const YELLOW_BG = "FFEB9C";
const RED_BG    = "FFC7CE";

const shpdData = [
  { city: "Алматы",   point: "ТЦ Mega",        plan: 120, fact: 115 },
  { city: "Алматы",   point: "ТЦ APORT",        plan: 80,  fact: 88  },
  { city: "Астана",   point: "ТЦ Хан Шатыр",   plan: 100, fact: 72  },
  { city: "Астана",   point: "Офис центральный", plan: 60,  fact: 61  },
  { city: "Шымкент",  point: "ТЦ Гранд Парк",  plan: 50,  fact: 38  },
];

const tvData = [
  { city: "Алматы",   point: "ТЦ Mega",        plan: 90,  fact: 95  },
  { city: "Алматы",   point: "ТЦ APORT",        plan: 60,  fact: 55  },
  { city: "Астана",   point: "ТЦ Хан Шатыр",   plan: 75,  fact: 60  },
  { city: "Астана",   point: "Офис центральный", plan: 45,  fact: 48  },
  { city: "Шымкент",  point: "ТЦ Гранд Парк",  plan: 40,  fact: 30  },
];

const fmsData = [
  { city: "Алматы",   point: "ТЦ Mega",        plan: 50,  fact: 52  },
  { city: "Алматы",   point: "ТЦ APORT",        plan: 35,  fact: 28  },
  { city: "Астана",   point: "ТЦ Хан Шатыр",   plan: 40,  fact: 35  },
  { city: "Астана",   point: "Офис центральный", plan: 30,  fact: 31  },
  { city: "Шымкент",  point: "ТЦ Гранд Парк",  plan: 25,  fact: 18  },
];

const employees = [
  { name: "Калибеков Ерасыл",   position: "Хантер", pct: 100 },
  { name: "Нурсултанов Абзал",  position: "Хантер", pct: 80  },
  { name: "Байораз Нурлан",     position: "Хантер", pct: 65  },
];

// ─── Утилиты ────────────────────────────────────────────────────────────────

function pct(plan, fact) {
  return plan > 0 ? Math.round((fact / plan) * 100) : 0;
}

function txt(text, opts = {}) {
  return new TextRun({ text: String(text), font: "Arial", size: opts.size || 22,
    bold: opts.bold || false, color: opts.color || "000000" });
}

function cell(content, opts = {}) {
  const shading = opts.shading
    ? { fill: opts.shading, type: ShadingType.CLEAR, color: "auto" }
    : undefined;
  return new TableCell({
    children: [new Paragraph({
      children: Array.isArray(content) ? content : [txt(content, opts)],
      alignment: opts.align || AlignmentType.LEFT,
    })],
    shading,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
  });
}

function headerRow(cols) {
  return new TableRow({
    children: cols.map(c => cell(c, { bold: true, color: WHITE, shading: CORP_BLUE, align: AlignmentType.CENTER })),
    tableHeader: true,
  });
}

function sectionTitle(text) {
  return new Paragraph({
    children: [txt(text, { bold: true, size: 28, color: CORP_BLUE })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
  });
}

// ─── Таблица продаж по точкам ───────────────────────────────────────────────

function salesTable(data) {
  const COLS = [800, 3000, 1000, 1000, 1200, 1000];
  const rows = [
    headerRow(["№", "Точка продаж", "План", "Факт", "% вып.", "Остаток"]),
    ...data.map((r, i) => {
      const p = pct(r.plan, r.fact);
      const bg = p >= 100 ? GREEN_BG : p >= 75 ? YELLOW_BG : RED_BG;
      return new TableRow({ children: [
        cell(i + 1,          { align: AlignmentType.CENTER, width: COLS[0] }),
        cell(`${r.city} — ${r.point}`, { width: COLS[1] }),
        cell(r.plan,         { align: AlignmentType.CENTER, width: COLS[2] }),
        cell(r.fact,         { align: AlignmentType.CENTER, width: COLS[3] }),
        cell(`${p}%`,        { align: AlignmentType.CENTER, width: COLS[4], shading: bg }),
        cell(r.plan - r.fact,{ align: AlignmentType.CENTER, width: COLS[5] }),
      ]});
    }),
  ];
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

// ─── Сводная таблица ────────────────────────────────────────────────────────

function summaryTable() {
  function sumField(arr, f) { return arr.reduce((s, r) => s + r[f], 0); }
  const services = [
    { name: "ШПД", data: shpdData },
    { name: "TV",  data: tvData   },
    { name: "FMS", data: fmsData  },
  ];
  const COLS = [1500, 1500, 1500, 1500];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      headerRow(["Услуга", "План", "Факт", "% выполнения"]),
      ...services.map(s => {
        const plan = sumField(s.data, "plan");
        const fact = sumField(s.data, "fact");
        const p = pct(plan, fact);
        const bg = p >= 100 ? GREEN_BG : p >= 75 ? YELLOW_BG : RED_BG;
        return new TableRow({ children: [
          cell(s.name, { bold: true, width: COLS[0] }),
          cell(plan,   { align: AlignmentType.CENTER, width: COLS[1] }),
          cell(fact,   { align: AlignmentType.CENTER, width: COLS[2] }),
          cell(`${p}%`,{ align: AlignmentType.CENTER, width: COLS[3], shading: bg }),
        ]});
      }),
    ],
  });
}

// ─── Таблица сотрудников ────────────────────────────────────────────────────

function employeesTable() {
  const COLS = [600, 3000, 1500, 1200, 1000];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      headerRow(["№", "Сотрудник", "Должность", "% вып.", "Статус"]),
      ...employees.map((e, i) => {
        const bg = e.pct >= 100 ? GREEN_BG : e.pct >= 75 ? YELLOW_BG : RED_BG;
        const status = e.pct >= 100 ? "✓" : e.pct >= 75 ? "~" : "!";
        return new TableRow({ children: [
          cell(i + 1,      { align: AlignmentType.CENTER, width: COLS[0] }),
          cell(e.name,     { width: COLS[1] }),
          cell(e.position, { width: COLS[2] }),
          cell(`${e.pct}%`,{ align: AlignmentType.CENTER, width: COLS[3], shading: bg }),
          cell(status,     { align: AlignmentType.CENTER, width: COLS[4], shading: bg, bold: true }),
        ]});
      }),
    ],
  });
}

// ─── Документ ───────────────────────────────────────────────────────────────

async function generateReport() {
  const doc = new Document({
    sections: [{
      headers: {
        default: new Header({ children: [
          new Paragraph({
            children: [txt("КОНФИДЕНЦИАЛЬНО", { bold: true, color: "C00000" })],
            alignment: AlignmentType.RIGHT,
          }),
        ]}),
      },
      footers: {
        default: new Footer({ children: [
          new Paragraph({
            children: [
              txt("Страница "),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 20 }),
              txt(" из "),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 20 }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ]}),
      },
      children: [
        // Титул
        new Paragraph({
          children: [txt("ОТЧЁТ ПО ПРОДАЖАМ", { bold: true, size: 36, color: CORP_BLUE })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 400, after: 120 },
        }),
        new Paragraph({
          children: [txt(`Период: ${PERIOD}`, { size: 24 })],
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          children: [txt(`Дата формирования: ${GENERATED}`, { size: 24 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Сводные показатели
        sectionTitle("Сводные показатели"),
        summaryTable(),

        // ШПД
        sectionTitle("Раздел 1 — ШПД (Широкополосный доступ в интернет)"),
        salesTable(shpdData),

        // TV
        sectionTitle("Раздел 2 — TV (Кабельное телевидение)"),
        salesTable(tvData),

        // FMS
        sectionTitle("Раздел 3 — FMS (Фиксированная мобильная связь)"),
        salesTable(fmsData),

        // Сотрудники
        sectionTitle("Раздел 4 — План продаж по сотрудникам"),
        employeesTable(),

        new Paragraph({
          children: [
            txt("🟢 Зелёный — ≥100%   🟡 Жёлтый — 75–99%   🔴 Красный — <75%", { size: 18, color: "595959" })
          ],
          spacing: { before: 120 },
        }),
      ],
    }],
  });

  const outDir = path.join(__dirname, "output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, `sales_report_${PERIOD.replace(" ", "_")}.docx`);
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  console.log(`✅ Отчёт сохранён: ${outPath}`);
}

generateReport().catch(console.error);
