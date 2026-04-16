#!/usr/bin/env python3
"""複数ユーザーの部品情報を収集し、統合CSVを出力するツール。

想定フロー:
1. ベースURLまたはAPI URLテンプレートに対してユーザーごとにアクセス
2. JSON または HTML テーブルを解析して部品レコードへ正規化
3. 同一部品をユーザー横断でマージして CSV 出力
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib import parse, request


DEFAULT_BASE_URL = "http://egl620043.jpn.mds.honda.com/11G/3V5_WHS/#no-back"


class FirstTableParser(HTMLParser):
    """最初のtableを抽出する軽量HTMLパーサ。"""

    def __init__(self) -> None:
        super().__init__()
        self.in_table = False
        self.table_depth = 0
        self.in_cell = False
        self.current_cell: list[str] = []
        self.current_row: list[str] = []
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "table":
            if not self.in_table:
                self.in_table = True
            self.table_depth += 1
        if not self.in_table:
            return
        if tag == "tr":
            self.current_row = []
        elif tag in {"th", "td"}:
            self.in_cell = True
            self.current_cell = []

    def handle_data(self, data: str) -> None:
        if self.in_table and self.in_cell:
            self.current_cell.append(data.strip())

    def handle_endtag(self, tag: str) -> None:
        if not self.in_table:
            return
        if tag in {"th", "td"} and self.in_cell:
            self.in_cell = False
            value = " ".join(x for x in self.current_cell if x)
            self.current_row.append(value)
            self.current_cell = []
        elif tag == "tr" and self.current_row:
            if any(cell for cell in self.current_row):
                self.rows.append(self.current_row)
            self.current_row = []
        elif tag == "table":
            self.table_depth -= 1
            if self.table_depth <= 0:
                self.in_table = False


def coerce_number(value: Any) -> float | int | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace(",", "")
    if not s:
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return None


@dataclass
class PartRecord:
    user: str
    part_no: str
    part_name: str
    quantity: float
    location: str
    raw: dict[str, Any]


def canonical_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def pick_value(row: dict[str, Any], aliases: list[str], default: Any = "") -> Any:
    normalized = {canonical_key(k): v for k, v in row.items()}
    for alias in aliases:
        key = canonical_key(alias)
        if key in normalized and normalized[key] not in (None, ""):
            return normalized[key]
    return default


def normalize_record(row: dict[str, Any], user: str) -> PartRecord | None:
    part_no = str(pick_value(row, ["part_no", "partno", "part number", "品番", "部品番号"], "")).strip()
    if not part_no:
        return None
    part_name = str(pick_value(row, ["part_name", "name", "description", "部品名"], "")).strip()
    qty_raw = pick_value(row, ["quantity", "qty", "在庫数", "数量"], 0)
    quantity = coerce_number(qty_raw)
    location = str(pick_value(row, ["location", "shelf", "bin", "保管場所"], "")).strip()
    return PartRecord(
        user=user,
        part_no=part_no,
        part_name=part_name,
        quantity=float(quantity) if quantity is not None else 0.0,
        location=location,
        raw=row,
    )


def parse_json_records(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        return [x for x in payload if isinstance(x, dict)]
    if isinstance(payload, dict):
        for key in ("items", "parts", "data", "results", "rows"):
            value = payload.get(key)
            if isinstance(value, list):
                return [x for x in value if isinstance(x, dict)]
        return [payload]
    return []


def parse_html_records(text: str) -> list[dict[str, Any]]:
    parser = FirstTableParser()
    parser.feed(text)
    rows = parser.rows
    if len(rows) < 2:
        return []
    headers = rows[0]
    data_rows = rows[1:]
    records: list[dict[str, Any]] = []
    for row in data_rows:
        padded = row + [""] * max(0, len(headers) - len(row))
        records.append(dict(zip(headers, padded, strict=False)))
    return records


def build_url(base_url: str, user: str, user_param: str) -> str:
    parsed = parse.urlsplit(base_url)
    if parsed.scheme == "file":
        return base_url
    current = parse.parse_qsl(parsed.query, keep_blank_values=True)
    current = [(k, v) for k, v in current if k != user_param]
    current.append((user_param, user))
    query = parse.urlencode(current)
    return parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, query, parsed.fragment))


def fetch_text(url: str, timeout: int, headers: dict[str, str]) -> str:
    req = request.Request(url, headers=headers)
    with request.urlopen(req, timeout=timeout) as res:
        charset = res.headers.get_content_charset() or "utf-8"
        return res.read().decode(charset, errors="replace")


def fetch_records(url: str, timeout: int, headers: dict[str, str], forced_format: str) -> list[dict[str, Any]]:
    body = fetch_text(url, timeout=timeout, headers=headers)
    if forced_format in {"json", "auto"}:
        try:
            return parse_json_records(json.loads(body))
        except json.JSONDecodeError:
            if forced_format == "json":
                raise
    if forced_format in {"html", "auto"}:
        return parse_html_records(body)
    raise ValueError(f"未対応のformat: {forced_format}")


def merge_records(records: list[PartRecord]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], dict[str, Any]] = defaultdict(
        lambda: {
            "part_no": "",
            "part_name": "",
            "total_quantity": 0.0,
            "users": set(),
            "locations": set(),
            "user_quantity": defaultdict(float),
        }
    )

    for r in records:
        key = (r.part_no, r.part_name)
        g = grouped[key]
        g["part_no"] = r.part_no
        g["part_name"] = r.part_name
        g["total_quantity"] += r.quantity
        g["users"].add(r.user)
        if r.location:
            g["locations"].add(r.location)
        g["user_quantity"][r.user] += r.quantity

    output: list[dict[str, Any]] = []
    for value in grouped.values():
        output.append(
            {
                "part_no": value["part_no"],
                "part_name": value["part_name"],
                "total_quantity": value["total_quantity"],
                "users": ";".join(sorted(value["users"])),
                "locations": ";".join(sorted(value["locations"])),
                "user_quantity": json.dumps(dict(sorted(value["user_quantity"].items())), ensure_ascii=False),
            }
        )
    return sorted(output, key=lambda x: (x["part_no"], x["part_name"]))


def write_csv(rows: list[dict[str, Any]], path: Path) -> None:
    if not rows:
        raise RuntimeError("CSVに出力できるデータがありません。")

    fieldnames = ["part_no", "part_name", "total_quantity", "users", "locations", "user_quantity"]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def parse_headers(header_args: list[str]) -> dict[str, str]:
    headers: dict[str, str] = {}
    for h in header_args:
        if ":" not in h:
            raise ValueError(f"ヘッダー指定が不正です: {h} (例: 'Cookie: session=...')")
        key, value = h.split(":", 1)
        headers[key.strip()] = value.strip()
    return headers


def main() -> int:
    parser = argparse.ArgumentParser(description="複数ユーザーの部品情報を集約しCSV出力します。")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="取得対象URL（ユーザーはクエリに追加）")
    parser.add_argument("--users", required=True, help="対象ユーザーIDをカンマ区切りで指定 (例: u001,u002)")
    parser.add_argument("--user-param", default="user", help="URLに付けるユーザー識別パラメータ名")
    parser.add_argument("--format", choices=["auto", "json", "html"], default="auto", help="レスポンス形式")
    parser.add_argument("--header", action="append", default=[], help="HTTPヘッダーを追加。複数指定可。例: --header 'Cookie: a=b'")
    parser.add_argument("--timeout", type=int, default=30, help="HTTPタイムアウト秒")
    parser.add_argument("--output", default="parts_merged.csv", help="出力CSVパス")
    args = parser.parse_args()

    users = [u.strip() for u in args.users.split(",") if u.strip()]
    if not users:
        parser.error("--users に1件以上のユーザーを指定してください。")

    headers = parse_headers(args.header)
    all_records: list[PartRecord] = []

    for user in users:
        url = build_url(args.base_url, user, args.user_param)
        try:
            source_rows = fetch_records(url, args.timeout, headers, args.format)
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] user={user} の取得に失敗: {exc}", file=sys.stderr)
            continue

        if not source_rows:
            print(f"[WARN] user={user} から部品データを抽出できませんでした。", file=sys.stderr)
            continue

        count = 0
        for row in source_rows:
            rec = normalize_record(row, user)
            if rec is None:
                continue
            all_records.append(rec)
            count += 1

        print(f"[INFO] user={user}: {count} 件の部品を取り込み。", file=sys.stderr)

    if not all_records:
        print("[ERROR] 有効な部品データが0件のため終了します。", file=sys.stderr)
        return 1

    merged = merge_records(all_records)
    output = Path(args.output)
    write_csv(merged, output)
    print(f"[OK] {len(merged)}件を {output} に出力しました。", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
