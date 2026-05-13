#!/usr/bin/env node

/**
 * DOCX report generator for Lab 2 — Algorithms & Data Structures
 * Topic: Algorithms for working with data structures
 *
 * Usage: npm run report
 */

import {
  Document, Packer, Paragraph, TextRun, Header,
  AlignmentType, PageNumber, HeadingLevel,
  PageBreak,
} from "docx";
import { readFileSync, writeFileSync, readdirSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// ─── Constants ───────────────────────────────────────────────────────

const MM_TO_DXA = 56.693;
const PT_TO_HALF_PT = 2;

const FONT = "Times New Roman";
const FONT_CODE = "Courier New";
const BODY_SIZE = 14 * PT_TO_HALF_PT;
const CODE_SIZE = 10 * PT_TO_HALF_PT;
const TITLE_SIZE = 14 * PT_TO_HALF_PT;
const LINE_SPACING_15 = 360;

const margins = {
  top: Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left: Math.round(30 * MM_TO_DXA),
  right: Math.round(15 * MM_TO_DXA),
};

// ─── Helpers (from lab-report.mjs template) ──────────────────────────

function titleRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: TITLE_SIZE,
    bold: opts.bold ?? false,
    ...opts,
  });
}

function bodyRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: BODY_SIZE,
    ...opts,
  });
}

function centeredParagraph(runs, spacing = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function emptyLine() {
  return centeredParagraph(titleRun(""));
}

function sectionHeading(number, title, opts = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: opts.pageBreakBefore ?? false,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [
      new TextRun({
        text: [number, title].filter(Boolean).join(" ").toUpperCase(),
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function subsectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      new TextRun({
        text: `${number} ${title}`,
        font: FONT,
        size: BODY_SIZE,
        bold: true,
      }),
    ],
  });
}

function bodyParagraph(text, opts = {}) {
  return new Paragraph({
    spacing: {
      before: opts.before ?? 0,
      after: opts.after ?? 0,
      line: LINE_SPACING_15,
      lineRule: "auto",
    },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(text)],
  });
}

function codeParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    children: [
      new TextRun({
        text,
        font: FONT_CODE,
        size: CODE_SIZE,
      }),
    ],
  });
}

function listingCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: Math.round(12.5 * MM_TO_DXA) },
    children: [
      bodyRun(`Лістинг ${number} — ${title}`),
    ],
  });
}

// ─── Code block helper ───────────────────────────────────────────────

function codeBlock(text) {
  return text.split("\n").map((line) => codeParagraph(line));
}

// ─── Source code reader ──────────────────────────────────────────────

function readSourceFiles() {
  const srcDir = join(__dirname, "..", "..", "src");
  const files = [];

  function walk(dir, prefix = "") {
    const entries = readdirSync(dir, { withFileTypes: true }).sort((a, b) =>
      a.name.localeCompare(b.name),
    );
    for (const entry of entries) {
      const fullPath = join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath, prefix ? `${prefix}/${entry.name}` : entry.name);
      } else if (entry.name.endsWith(".ts")) {
        const relativeName = prefix ? `${prefix}/${entry.name}` : entry.name;
        files.push({
          name: relativeName,
          content: readFileSync(fullPath, "utf-8"),
        });
      }
    }
  }

  walk(srcDir);
  return files;
}

// ─── Read structure source files ─────────────────────────────────────

const srcDir = join(__dirname, "..", "..", "src", "structures");
const stackCode = readFileSync(join(srcDir, "stack.ts"), "utf-8").trimEnd();
const queueCode = readFileSync(join(srcDir, "queue.ts"), "utf-8").trimEnd();
const linkedListCode = readFileSync(join(srcDir, "linked-list.ts"), "utf-8").trimEnd();

// ─── TITLE PAGE ──────────────────────────────────────────────────────

const titlePageParagraphs = [
  centeredParagraph(titleRun("Міністерство освіти і науки України")),
  centeredParagraph(titleRun("Харківський національний університет радіоелектроніки")),
  emptyLine(),
  centeredParagraph(titleRun("Кафедра програмної інженерії")),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("ЗВІТ", { bold: true })),
  centeredParagraph(titleRun("з лабораторної роботи № 2")),
  centeredParagraph(titleRun("з дисципліни «Алгоритми та структури даних»")),
  centeredParagraph(titleRun("на тему: «Алгоритми роботи зі структурами даних»")),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("Варіант 9")),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Виконав: ст. гр. ПЗПІ-25-6")],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Коновалов О. О.")],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Перевірив: Олійник О. О.")],
  }),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  emptyLine(),
  centeredParagraph(titleRun("Харків — 2026")),
];

// ─── BODY SECTIONS ───────────────────────────────────────────────────

const programOutput = `Лабораторна робота №2 — Структури даних
Варіант 9

=== Завдання 1: Стек на основі масиву ===

PUSH 5 елементів:
  push(10)
  push(20)
  push(30)
  push(40)
  push(50)
Stack [bottom → top]: 10, 20, 30, 40, 50

POP 2 елементи:
  pop() → 50
  pop() → 40

Залишок стеку:
Stack [bottom → top]: 10, 20, 30

=== Завдання 2: Черга на основі масиву ===

ENQUEUE 5 елементів:
  enqueue(10)
  enqueue(20)
  enqueue(30)
  enqueue(40)
  enqueue(50)
Queue [head → tail]: 10, 20, 30, 40, 50

DEQUEUE 2 елементи:
  dequeue() → 10
  dequeue() → 20

Залишок черги:
Queue [head → tail]: 30, 40, 50

=== Завдання 3: Зв'язковий список ===

Створення списку з 5 елементів:
  append(10)
  append(20)
  append(30)
  append(40)
  append(50)

Список:
List: 10 → 20 → 30 → 40 → 50 → null`;

const bodyParagraphs = [
  new Paragraph({ children: [new PageBreak()] }),

  // 1 МЕТА РОБОТИ
  sectionHeading("1", "Мета роботи"),
  bodyParagraph(
    "Навчитися реалізовувати базові структури даних (стек, чергу, зв'язковий список) та операції з ними.",
  ),

  // 2 ЗАВДАННЯ
  sectionHeading("2", "Завдання"),
  bodyParagraph(
    "Реалізувати наступні структури даних без використання стандартних класів бібліотек:",
  ),
  bodyParagraph(
    "1) Стек на основі масиву з операціями PUSH та POP. Записати 5 елементів, виштовхнути 2, роздрукувати залишок.",
  ),
  bodyParagraph(
    "2) Чергу на основі масиву з операціями enqueue та dequeue.",
  ),
  bodyParagraph(
    "3) Зв'язковий список. Створити список із п'яти елементів та роздрукувати його.",
  ),

  // 3 ХІД РОБОТИ
  sectionHeading("3", "Хід роботи"),

  // 3.1 Стек на основі масиву
  subsectionHeading("3.1", "Стек на основі масиву"),
  bodyParagraph(
    "Стек — це структура даних, що працює за принципом LIFO (Last In, First Out). Реалізація виконана на основі масиву фіксованої ємності та індексу top, що вказує на верхній елемент. Операція push збільшує top на одиницю та записує значення, операція pop зчитує значення та зменшує top. При спробі додати елемент до переповненого стеку генерується виняток overflow, при вилученні з порожнього — underflow.",
  ),

  listingCaption("3.1", "Реалізація стеку"),
  ...codeBlock(stackCode),

  // 3.2 Черга на основі масиву
  subsectionHeading("3.2", "Черга на основі масиву"),
  bodyParagraph(
    "Черга — це структура даних, що працює за принципом FIFO (First In, First Out). Реалізація виконана як циклічний буфер з використанням індексів head, tail та лічильника size. Операція enqueue додає елемент у позицію tail, операція dequeue вилучає елемент із позиції head. Обидва індекси обертаються за модулем ємності масиву, що забезпечує ефективне використання пам'яті.",
  ),

  listingCaption("3.2", "Реалізація черги"),
  ...codeBlock(queueCode),

  // 3.3 Зв'язковий список
  subsectionHeading("3.3", "Зв'язковий список"),
  bodyParagraph(
    "Однозв'язковий список — це динамічна структура даних, де кожен вузол зберігає значення та посилання на наступний вузол. Операція append проходить список до останнього вузла та додає новий елемент. Операція print обходить список від head до null, збираючи значення для виведення.",
  ),

  listingCaption("3.3", "Реалізація зв'язкового списку"),
  ...codeBlock(linkedListCode),

  // 4 РЕЗУЛЬТАТИ
  sectionHeading("4", "Результати"),
  bodyParagraph("Результати виконання програми наведено нижче."),

  listingCaption("4.1", "Вивід програми"),
  ...codeBlock(programOutput),

  // 5 ВИСНОВКИ
  sectionHeading("5", "Висновки"),
  bodyParagraph(
    "У ході лабораторної роботи було реалізовано три базові структури даних — стек, чергу та зв'язковий список — без використання стандартних бібліотечних класів. Стек реалізовано на основі масиву з принципом LIFO, чергу — як циклічний буфер з принципом FIFO, список — як однозв'язкову динамічну структуру. Результати демонструють коректну роботу всіх операцій.",
  ),
];

// ─── ДОДАТОК А — Вихідний код програми ──────────────────────────────

const appendixParagraphs = [
  sectionHeading("", "Додаток А", { pageBreakBefore: true }),
  centeredParagraph(bodyRun("Вихідний код програми", { bold: true })),
  emptyLine(),
];

const sourceFiles = readSourceFiles();
let listingCounter = 1;
for (const file of sourceFiles) {
  appendixParagraphs.push(
    listingCaption(`А.${listingCounter}`, file.name),
  );
  appendixParagraphs.push(...codeBlock(file.content.trimEnd()));
  appendixParagraphs.push(emptyLine());
  listingCounter++;
}

// ─── DOCUMENT ────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: FONT,
          size: BODY_SIZE,
          language: { value: "uk-UA" },
        },
        paragraph: {
          spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
        },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: {
          spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
          outlineLevel: 1,
        },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { ...margins, header: 708, footer: 708 },
        },
        titlePage: true,
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  children: [PageNumber.CURRENT],
                  font: FONT,
                  size: BODY_SIZE,
                }),
              ],
            }),
          ],
        }),
      },
      children: [
        ...titlePageParagraphs,
        ...bodyParagraphs,
        ...appendixParagraphs,
      ],
    },
  ],
});

// ─── Output ──────────────────────────────────────────────────────────

const outputPath = join(__dirname, "Звіт_ЛР2_Коновалов_ПЗПІ-25-6.docx");
const buffer = await Packer.toBuffer(doc);
writeFileSync(outputPath, buffer);
console.log(`Created: ${outputPath}`);
