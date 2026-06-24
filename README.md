# Корпоративная цифровая платформа

Три модуля согласно ТЗ.

## Структура

```
├── report/
│   └── generate.js     # Генератор отчёта .docx (Node.js + docx v9)
├── portal/
│   └── index.html      # Портал абонента — выбор менеджера
├── crm/
│   └── index.html      # CRM-дашборд входящих заявок
└── package.json
```

## Запуск

### Модуль 1 — Генератор отчёта Word
```bash
npm install
npm run generate-report
# Файл сохраняется в report/output/
```

### Модуль 2 — Портал абонента
Открыть `portal/index.html` в браузере.

### Модуль 3 — CRM-дашборд
Открыть `crm/index.html` в браузере.
