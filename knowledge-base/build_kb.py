#!/usr/bin/env python3
"""
Собирает автономный index.html базы знаний из ges2_kb_data.json + template.html.

Использование:
    python3 knowledge-base/build_kb.py
    python3 knowledge-base/build_kb.py --data ../docs/baza_znaniy/ges2_kb_data.json --out index.html

Перед сборкой валидирует единственное жёсткое правило целостности из README:
    status == "answered"  =>  answer и source заполнены
    status == "pending"   =>  answer пустой
Сборка останавливается с понятной ошибкой, если правило нарушено — так
испорченные данные не попадают в готовый файл незаметно.
"""

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).parent


def validate(data):
    errors = []
    seen_ids = set()
    for e in data["entries"]:
        eid = e.get("id", "???")
        if eid in seen_ids:
            errors.append(f"{eid}: дублирующийся id")
        seen_ids.add(eid)

        if e["status"] == "answered":
            if not e.get("answer"):
                errors.append(f"{eid}: status=answered, но answer пустой")
            if not e.get("source"):
                errors.append(f"{eid}: status=answered, но source не указан")
        elif e["status"] == "pending":
            if e.get("answer"):
                errors.append(f"{eid}: status=pending, но answer не пустой")
        else:
            errors.append(f"{eid}: неизвестный status '{e['status']}'")

    counted_answered = sum(1 for e in data["entries"] if e["status"] == "answered")
    counted_pending = sum(1 for e in data["entries"] if e["status"] == "pending")
    meta_counts = data["meta"].get("status_counts", {})
    if meta_counts.get("answered") != counted_answered:
        errors.append(
            f"meta.status_counts.answered={meta_counts.get('answered')}, "
            f"а по факту записей answered: {counted_answered} (обновите meta или проверьте данные)"
        )
    if meta_counts.get("pending") != counted_pending:
        errors.append(
            f"meta.status_counts.pending={meta_counts.get('pending')}, "
            f"а по факту записей pending: {counted_pending} (обновите meta или проверьте данные)"
        )

    return errors


def build(data_path: Path, template_path: Path, out_path: Path):
    data = json.loads(data_path.read_text(encoding="utf-8"))

    errors = validate(data)
    if errors:
        print("Сборка остановлена — данные не прошли проверку целостности:", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)
        sys.exit(1)

    total = len(data["entries"])
    answered = sum(1 for e in data["entries"] if e["status"] == "answered")
    pending = total - answered

    kb_json = json.dumps(data, ensure_ascii=False)
    # Защита от преждевременного закрытия <script>, если в данных когда-нибудь
    # появится буквальная подстрока "</" перед "script" (или похожая) —
    # безопасно для JSON: экранированный "/" декодируется обратно в "/".
    kb_json_safe = kb_json.replace("</", "<\\/")

    template = template_path.read_text(encoding="utf-8")
    html = (
        template
        .replace("__KB_DATA_JSON__", kb_json_safe)
        .replace("__GENERATED_DATE__", data["meta"]["generated"])
        .replace("__TOTAL__", str(total))
        .replace("__ANSWERED__", str(answered))
        .replace("__PENDING__", str(pending))
    )

    out_path.write_text(html, encoding="utf-8")
    print(f"OK: {out_path} собран из {total} записей ({answered} отвечено, {pending} уточняется).")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(HERE / ".." / "docs" / "baza_znaniy" / "ges2_kb_data.json"))
    ap.add_argument("--template", default=str(HERE / "template.html"))
    ap.add_argument("--out", default=str(HERE / "index.html"))
    args = ap.parse_args()
    build(Path(args.data), Path(args.template), Path(args.out))


if __name__ == "__main__":
    main()
