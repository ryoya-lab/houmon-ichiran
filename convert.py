#!/usr/bin/env python3
"""訪問内容一覧表.xlsx をスマホ閲覧用 index.html に変換する。

使い方:
    python convert.py 訪問内容一覧表.xlsx [出力先index.html]
"""
import json
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook

DAYS = ["月", "火", "水", "木", "金", "土", "日"]


def parse_sheet(ws, sheet_name):
    client = {"name": sheet_name.strip(), "policy": "", "visits": {d: [] for d in DAYS}, "notes": []}
    current_day = None
    for row in ws.iter_rows(min_row=1):
        b, c, d, e, f = (row[i].value if i < len(row) else None for i in range(1, 6))
        b = str(b).strip() if b else ""
        c = str(c).strip() if c else ""
        if b == "【援助方針】":
            client["policy"] = str(row[2].value or "").strip()
            continue
        if c == "曜日":  # ヘッダー行
            continue
        if b.startswith("※"):
            client["notes"].append(b.lstrip("※").strip())
            continue
        day = next((x for x in DAYS if c.startswith(x)), None)
        if day:
            current_day = day
        if current_day and d:
            client["visits"][current_day].append({
                "time": str(d).strip(),
                "code": str(e or "").strip(),
                "text": str(f or "").strip(),
            })
    return client


def main():
    if len(sys.argv) < 2:
        sys.exit("使い方: python convert.py 訪問内容一覧表.xlsx [index.html]")
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent / "index.html"

    wb = load_workbook(src, data_only=True)
    clients = [parse_sheet(wb[name], name) for name in wb.sheetnames]

    jst = timezone(timedelta(hours=9))
    updated = datetime.now(jst).strftime("%Y/%m/%d %H:%M")

    template = (Path(__file__).parent / "template.html").read_text(encoding="utf-8")
    html = template.replace("/*__DATA__*/[]", json.dumps(clients, ensure_ascii=False))
    html = html.replace("__UPDATED__", updated)
    dst.write_text(html, encoding="utf-8")
    print(f"OK: {dst}（利用者{len(clients)}名、更新 {updated}）")


if __name__ == "__main__":
    main()
