#!/usr/bin/env python3
"""
Парсер календаря событий ГЭС-2 (ges-2.org).

Использует тот же GraphQL-эндпоинт, который дергает сам сайт для рендера
календаря (найден через анализ сетевых запросов браузера, operation
"newMaterialsInterval", Apollo persisted query). Никакого HTML не парсит —
получает структурированные данные напрямую.

Использование:
    python3 fetch_calendar.py                       # события с сегодня по конец года, JSON в stdout
    python3 fetch_calendar.py --from 2026-08-01 --to 2026-12-31
    python3 fetch_calendar.py --csv > calendar.csv
    python3 fetch_calendar.py --out calendar-data.json

Зависимости: только стандартная библиотека Python 3 (urllib, json).
"""

import argparse
import csv
import json
import sys
import urllib.parse
import urllib.request
from datetime import datetime, date

GRAPHQL_URL = "https://ges-2.org/graphql"
# Persisted query hash для operationName=newMaterialsInterval.
# Если сайт обновит фронтенд и хэш перестанет работать, его нужно заново
# подсмотреть в Network-логе браузера при открытии https://ges-2.org/calendar
# (запрос содержит extensions.persistedQuery.sha256Hash).
PERSISTED_QUERY_HASH = "38cd05024f96c697b38f6c56d7db07275e1d0dec751f7183e4cee72da2d01aaf"


def fetch_raw(first_future=500, first_now=200, first_past=0, ges2_page_full_path=""):
    variables = {
        "firstFuture": first_future, "skipFuture": 0,
        "firstNow": first_now, "skipNow": 0,
        "firstPast": first_past, "skipPast": 0,
        "withDeepProjects": True,
        "ges2PageFullPath": ges2_page_full_path,
        "accessibilityGroupSlugs": [], "ageRestrictions": [], "areaSlugs": [],
        "citySlugs": [], "subtypeSlugs": [], "recommendedAges": [],
        "materialType": "", "modalitySlug": "", "seriesSlug": "",
        "tagSlug": "", "topicSlug": "",
    }
    extensions = {"persistedQuery": {"version": 1, "sha256Hash": PERSISTED_QUERY_HASH}}
    params = {
        "operationName": "newMaterialsInterval",
        "variables": json.dumps(variables, ensure_ascii=False),
        "extensions": json.dumps(extensions, ensure_ascii=False),
    }
    url = GRAPHQL_URL + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "ges2-calendar-tool/1.0"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.load(resp)


def public_path(full_path):
    """GraphQL отдаёт fullPath с внутренним префиксом сайта ("ges2/..."),
    который не существует в реальной публичной ссылке (проверено: /ges2/<slug>
    отдаёт 404, /<slug> — 200). "projects/..." — настоящий сегмент пути,
    его трогать не нужно."""
    if full_path and full_path.startswith("ges2/"):
        return full_path[len("ges2/"):]
    return full_path


def compact(materials):
    out = []
    for m in materials:
        out.append({
            "id": m.get("id"),
            "type": m.get("materialType"),
            "subtype": (m.get("subtype") or {}).get("title"),
            "title": (m.get("translation") or {}).get("title"),
            "path": public_path(m.get("fullPath")),
            "starts": m.get("projectStartsAt"),
            "ends": m.get("projectEndsAt"),
            "paid": m.get("paid"),
        })
    return out


def parse_dt(s):
    if not s:
        return None
    return datetime.fromisoformat(s)


def filter_range(items, date_from, date_to):
    result = []
    for it in items:
        s = parse_dt(it["starts"])
        e = parse_dt(it["ends"])
        if s is None:
            continue
        if s.date() > date_to:
            continue
        effective_end = e if e is not None else s
        if effective_end.date() < date_from:
            continue
        result.append(it)
    result.sort(key=lambda x: x["starts"])
    return result


def main():
    ap = argparse.ArgumentParser(description="Календарь событий ГЭС-2 из GraphQL API сайта")
    ap.add_argument("--from", dest="date_from", default=date.today().isoformat(),
                     help="дата начала диапазона, YYYY-MM-DD (по умолчанию — сегодня)")
    ap.add_argument("--to", dest="date_to", default=f"{date.today().year}-12-31",
                     help="дата конца диапазона, YYYY-MM-DD (по умолчанию — конец текущего года)")
    ap.add_argument("--csv", action="store_true", help="вывести в формате CSV вместо JSON")
    ap.add_argument("--out", help="записать результат в файл вместо stdout")
    ap.add_argument("--page", default="", help="ограничить конкретным разделом сайта (ges2PageFullPath), по умолчанию — весь сайт")
    args = ap.parse_args()

    date_from = date.fromisoformat(args.date_from)
    date_to = date.fromisoformat(args.date_to)

    raw = fetch_raw(ges2_page_full_path=args.page)
    if "errors" in raw:
        print("GraphQL error:", json.dumps(raw["errors"], ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    all_items = compact(raw["data"].get("future", []) + raw["data"].get("now", []))
    events = filter_range(all_items, date_from, date_to)

    out_stream = open(args.out, "w", encoding="utf-8") if args.out else sys.stdout

    if args.csv:
        writer = csv.DictWriter(out_stream, fieldnames=["starts", "ends", "type", "subtype", "title", "paid", "path"])
        writer.writeheader()
        for it in events:
            writer.writerow({
                "starts": it["starts"], "ends": it["ends"] or "",
                "type": it["type"], "subtype": it["subtype"] or "",
                "title": it["title"], "paid": it["paid"],
                "path": f"https://ges-2.org/{it['path']}",
            })
    else:
        json.dump(events, out_stream, ensure_ascii=False, indent=2)
        out_stream.write("\n")

    if args.out:
        out_stream.close()
        print(f"Записано {len(events)} событий в {args.out}", file=sys.stderr)
    else:
        print(f"# Событий: {len(events)}", file=sys.stderr)


if __name__ == "__main__":
    main()
