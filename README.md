# Trip expenses · Франция → Исландия → Швейцария (июль 2026)

Офлайн-репорт по тратам трипа: единый JSON → Excel + HTML-дашборд.

## Быстрый старт

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
# локальные (gitignored) файлы — один раз после клона:
cp data/meta.example.json data/meta.json
cp data/excluded.example.json data/excluded.json
# затем заполни meta.json / excluded.json своими данными
.venv/bin/python scripts/build.py
open trip_dashboard.html          # полная версия (со счётом NBG)
open trip_dashboard_public.html   # для шаринга (без банковского блока)
```

## Структура

```
data/
  transactions.json        # источник правды по тратам трипа
  meta.example.json        # шаблон → скопировать в meta.json
  excluded.example.json    # шаблон → скопировать в excluded.json
  excluded.json            # non-trip / internal / reverted (gitignore)
  meta.json                # KPI-мета, этапы, financing, checksum (gitignore)
templates/
  dashboard.html           # шаблон (плейсхолдеры __DATA__ / __META__ / …)
scripts/
  build.py                 # пересборка xlsx + обоих HTML + проверка checksum
raw/                       # сюда класть исходные выписки (gitignore)
trip_report.xlsx
trip_dashboard.html
trip_dashboard_public.html
PLAN.md
```

## Что умеет дашборд

- KPI: всего, shared ÷ travelers, личные, пре-трип / в поездке, €/день по фактическим дням
- Фильтры: этап, категория, источник, даты, поиск, пре-трип, личные
- Donut + дни + этапы; клик фильтрует
- Аналитика: еда по странам, фиксированные vs переменные, топ-10
- Вкладка «Исключено» и «Финансирование» — только в полной версии
- Экспорт отфильтрованной таблицы в CSV

## Править данные

1. Отредактируй `data/transactions.json` (или `excluded.json` / `meta.json`).
2. Запусти `scripts/build.py` — он сверит checksum и пересоберёт артефакты.

Ключевые поля транзакции: `dt`, `src`, `desc`, `label`, `category`, `stage`, `pretrip`, `amount_eur`, опционально `personal`, `trip_day` (если settlement ≠ день поездки), `time_unknown`.

## Сверка

`meta.checksum.trip_total` должен равняться сумме `-amount_eur` по всем транзакциям. Сейчас: **9 826,98 €** (shared 9 477,07 € → **3 159,02 €/чел**, личное 349,91 €).


## GitHub Pages

Публичная версия лежит в `docs/index.html`: без financing/account_picture и без номеров карт/счёта в `src` и `checksum.by_source`.

После пуша в GitHub: **Settings → Pages → Deploy from branch → `main` / `/docs`**.

Приватные файлы (`trip_dashboard.html`, `trip_report.xlsx`, `data/meta.json`, `data/excluded.json`) в `.gitignore` — только локально.
