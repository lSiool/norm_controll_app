# Thesis Auto-Corrector

Автоматическая нормоконтроль дипломных работ согласно **ПР V-08-2022**  
НАО «Карагандинский технический университет имени Абылкаса Сагинова»

---

## Архитектура

```
thesis-corrector/
├── app/
│   ├── main.py                        # FastAPI: /correct, /rules, /download
│   └── corrector/
│       └── thesis_corrector.py        # Движок нормоконтроля
├── data/
│   └── norm_control_rules.json        # Правила, извлечённые из ПР V-08-2022
└── requirements.txt
```

---

## Что проверяется и исправляется автоматически

| Правило   | Что делает                                         | Авто |
|-----------|----------------------------------------------------|------|
| 5.3       | Шрифт Times New Roman 14pt во всём тексте          | ✅   |
| 5.5       | Поля: лево 30, верх 20, право 10, низ 25 мм        | ✅   |
| 5.5       | Абзацный отступ 7.5 мм, выравнивание по ширине     | ✅   |
| 7.1.10    | Точка в конце заголовка (запрещено)                | ⚠️  |
| 7.1.10    | Перенос слов в заголовке (запрещено)               | ⚠️  |
| 7.3.1     | Формат подписи «Рисунок N – Наименование»          | ⚠️  |
| 7.5.1     | Формат заголовка «Таблица N – Наименование»        | ⚠️  |
| 7.1.19.2  | Ссылки на литературу [N], не (N)                   | ⚠️  |
| 7.2.4     | Символ Ø в тексте → писать «диаметр»               | ⚠️  |

✅ — исправляется автоматически  
⚠️ — помечается в отчёте, требует ручной проверки

---

## Запуск

```bash
pip install -r requirements.txt
cd app
uvicorn main:app --reload --port 8000
```

Открыть документацию API: http://localhost:8000/docs

---

## Использование API

```bash
# Загрузить дипломную работу и получить отчёт
curl -X POST http://localhost:8000/correct \
  -F "file=@thesis.docx" | python -m json.tool

# Скачать исправленный файл
curl -o corrected.docx http://localhost:8000/download/{job_id}/docx

# Скачать текстовый отчёт
curl -o report.txt http://localhost:8000/download/{job_id}/report
```

---

## Структура norm_control_rules.json

```
page_layout          — формат А4, поля (мм)
typography           — шрифт, размер, межстрочный
paragraph_formatting — отступ, выравнивание
document_structure   — обязательные разделы
page_numbering       — позиция, стиль номеров
headings             — нумерация, правила заголовков
illustrations        — нумерация, формат подписей
tables               — нумерация, формат заголовков
formulas             — нумерация, расположение
bibliography         — ГОСТ 7.32, формат [N]
appendices           — обозначение А–Т, правила
footnotes            — знак, позиция
text_style_rules     — запрещённые конструкции
volume_limits        — 70/120 страниц
```

---

## Расширение правил (Phase 1)

Если появится новая инструкция нормоконтроля в PDF/Word/TXT:

```python
from app.parser.pdf_to_rules import parse_rules_from_pdf
new_rules = parse_rules_from_pdf("new_standard.pdf")
# → data/norm_control_rules.json обновляется через Claude API
```
