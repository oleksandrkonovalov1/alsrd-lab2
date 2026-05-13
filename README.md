# alsrd-lab2

Лабораторна робота №2 з дисципліни "Алгоритми та структури даних"

**Тема:** Алгоритми роботи зі структурами даних

**Варіант 9**

## Автор

- **Коновалов Олександр**, група ПЗПІ-25-6, oleksandr.konovalov1@nure.ua

## Технології

- TypeScript
- Node.js (tsx)
- docx (генерація звіту)
- VCS: Git + GitHub

## Опис проєкту

Реалізація базових структур даних без використання стандартних бібліотечних класів:
- Стек на основі масиву (PUSH/POP)
- Черга на основі масиву (enqueue/dequeue, циклічний буфер)
- Однозв'язковий список (append, print)

## Запуск

```bash
git clone https://github.com/oleksandrkonovalov1/alsrd-lab2.git
cd alsrd-lab2
npm install

# Запустити демонстрацію
npm start

# Згенерувати звіт
npm run report
```

## Структура

```
alsrd-lab2/
├── src/
│   ├── structures/
│   │   ├── stack.ts
│   │   ├── queue.ts
│   │   └── linked-list.ts
│   ├── demo-stack.ts
│   ├── demo-queue.ts
│   ├── demo-list.ts
│   └── main.ts
├── reports/lr2/
│   ├── generate-report.mjs
│   ├── screenshots/
│   └── Звіт_ЛР2_Коновалов_ПЗПІ-25-6.docx
├── package.json
├── tsconfig.json
├── .gitignore
├── README.md
└── LICENSE
```

## Ліцензія

MIT License
