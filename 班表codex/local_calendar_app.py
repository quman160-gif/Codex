#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import html
import json
import mimetypes
import re
import urllib.parse
import webbrowser
from dataclasses import replace
from datetime import datetime, time, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from sync_schedule import (
    STORE_NAME,
    ShiftEvent,
    iter_schedule_events,
    render_calendar,
)


APP_DIR = Path(__file__).resolve().parent
EXPORT_DIR = APP_DIR / "local_exports"
ICS_DIR = EXPORT_DIR / "ics"
CSV_DIR = EXPORT_DIR / "csv"
MANIFEST_PATH = EXPORT_DIR / "manifest.json"
FIELDNAMES = [
    "Subject",
    "Start Date",
    "Start Time",
    "End Date",
    "End Time",
    "All Day Event",
    "Description",
    "Location",
    "Private",
]


def find_workbooks() -> list[Path]:
    return sorted(
        path for path in APP_DIR.glob("*.xlsx") if path.is_file() and not path.name.startswith("~$")
    )


def safe_filename(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "-", value.strip())
    value = re.sub(r"\s+", "-", value)
    return value or "calendar"


def roc_month_label(events: list[ShiftEvent]) -> str:
    months = sorted({(event.work_date.year - 1911, event.work_date.month) for event in events})
    if not months:
        return "班表"
    labels = [f"{year}{month:02d}" for year, month in months]
    return labels[0] if len(labels) == 1 else f"{labels[0]}-{labels[-1]}"


def datefmt(value: Any) -> str:
    return value.strftime("%m/%d/%Y")


def timefmt(value: time | None) -> str:
    return value.strftime("%I:%M %p") if value else ""


def csv_row(event: ShiftEvent, include_person: bool = False) -> dict[str, str]:
    subject = f"{event.person}｜{event.summary}" if include_person else event.summary
    return {
        "Subject": subject,
        "Start Date": datefmt(event.work_date),
        "Start Time": "" if event.all_day else timefmt(event.start),
        "End Date": datefmt(event.work_date),
        "End Time": "" if event.all_day else timefmt(event.end),
        "All Day Event": "True" if event.all_day else "False",
        "Description": f"人員：{event.person}\n班別：{event.code}\n來源：{event.source}",
        "Location": STORE_NAME,
        "Private": "False",
    }


def write_csv(path: Path, events: list[ShiftEvent], include_person: bool = False) -> None:
    sorted_events = sorted(events, key=lambda item: (item.work_date, item.person, item.start or time.min, item.summary))
    with path.open("w", encoding="utf-8", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=FIELDNAMES)
        writer.writeheader()
        for event in sorted_events:
            writer.writerow(csv_row(event, include_person))


def write_text_file(path: Path, content: str) -> None:
    with path.open("w", encoding="utf-8", newline="") as output:
        output.write(content)


def with_person_in_summary(events: list[ShiftEvent]) -> list[ShiftEvent]:
    return [replace(event, summary=f"{event.person}｜{event.summary}") for event in events]


def export_for_workbook(workbook: Path) -> dict[str, Any]:
    events, warnings = iter_schedule_events(workbook)
    generated_at = datetime.now(timezone.utc)
    label = roc_month_label(events)

    ICS_DIR.mkdir(parents=True, exist_ok=True)
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    for folder in (ICS_DIR, CSV_DIR):
        for old_file in folder.iterdir():
            if old_file.is_file():
                old_file.unlink()

    by_person: dict[str, list[ShiftEvent]] = {}
    for event in events:
        by_person.setdefault(event.person, []).append(event)

    people: list[dict[str, Any]] = []
    for person in sorted(by_person):
        person_events = by_person[person]
        base = f"{safe_filename(person)}-{label}"
        ics_path = ICS_DIR / f"{base}.ics"
        csv_path = CSV_DIR / f"{base}.csv"
        write_text_file(ics_path, render_calendar(person, person_events, generated_at))
        write_csv(csv_path, person_events)
        people.append(
            {
                "person": person,
                "events": len(person_events),
                "ics": f"/exports/ics/{urllib.parse.quote(ics_path.name)}",
                "csv": f"/exports/csv/{urllib.parse.quote(csv_path.name)}",
            }
        )

    all_ics = ICS_DIR / f"全部人-{label}.ics"
    all_csv = CSV_DIR / f"全部人-{label}.csv"
    write_text_file(all_ics, render_calendar("全部人", with_person_in_summary(events), generated_at))
    write_csv(all_csv, events, include_person=True)

    manifest = {
        "generated_at": generated_at.isoformat(),
        "workbook": workbook.name,
        "label": label,
        "events": len(events),
        "people": people,
        "all": {
            "events": len(events),
            "ics": f"/exports/ics/{urllib.parse.quote(all_ics.name)}",
            "csv": f"/exports/csv/{urllib.parse.quote(all_csv.name)}",
        },
        "warnings": warnings,
    }
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest


def load_manifest() -> dict[str, Any] | None:
    if not MANIFEST_PATH.exists():
        return None
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def render_page(status: str = "") -> str:
    workbooks = find_workbooks()
    manifest = load_manifest()
    options = "".join(
        f"<option value=\"{html.escape(path.name)}\">{html.escape(path.name)}</option>" for path in workbooks
    )
    workbook_control = (
        f"""
        <form action="/generate" method="post" class="toolbar">
          <label>
            <span>Excel 班表</span>
            <select name="workbook">{options}</select>
          </label>
          <button type="submit">產生檔案</button>
          <a class="secondary" href="https://calendar.google.com/calendar/u/0/r/settings/export" target="_blank" rel="noreferrer">Google 匯入頁</a>
        </form>
        """
        if workbooks
        else "<p class=\"notice\">目前資料夾內沒有 .xlsx 班表。</p>"
    )

    result_html = ""
    if manifest:
        person_rows = "".join(
            f"""
            <tr>
              <td>{html.escape(item["person"])}</td>
              <td>{item["events"]}</td>
              <td><a href="{item["ics"]}">ICS</a></td>
              <td><a href="{item["csv"]}">CSV</a></td>
            </tr>
            """
            for item in manifest["people"]
        )
        warnings = "".join(f"<li>{html.escape(item)}</li>" for item in manifest["warnings"])
        warning_html = f"<ul class=\"warnings\">{warnings}</ul>" if warnings else ""
        result_html = f"""
        <section class="results">
          <div class="summary">
            <div>
              <span>來源</span>
              <strong>{html.escape(manifest["workbook"])}</strong>
            </div>
            <div>
              <span>月份</span>
              <strong>{html.escape(manifest["label"])}</strong>
            </div>
            <div>
              <span>活動</span>
              <strong>{manifest["events"]}</strong>
            </div>
          </div>
          <table>
            <thead>
              <tr><th>人員</th><th>活動數</th><th>匯入 ICS</th><th>匯入 CSV</th></tr>
            </thead>
            <tbody>
              <tr>
                <td>全部人</td>
                <td>{manifest["all"]["events"]}</td>
                <td><a href="{manifest["all"]["ics"]}">ICS</a></td>
                <td><a href="{manifest["all"]["csv"]}">CSV</a></td>
              </tr>
              {person_rows}
            </tbody>
          </table>
          {warning_html}
        </section>
        """

    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{STORE_NAME}本機日曆工具</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #1f2933;
      --muted: #64748b;
      --line: #d8dee8;
      --fill: #f7f9fc;
      --accent: #0f766e;
      --accent-strong: #115e59;
      --warn: #b91c1c;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: #ffffff;
    }}
    main {{
      width: min(1040px, calc(100% - 32px));
      margin: 32px auto;
    }}
    header {{
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 24px;
      border-bottom: 1px solid var(--line);
      padding-bottom: 18px;
      margin-bottom: 22px;
    }}
    h1 {{
      font-size: 28px;
      line-height: 1.2;
      margin: 0 0 6px;
      font-weight: 750;
    }}
    p {{
      margin: 0;
      color: var(--muted);
      font-size: 15px;
    }}
    .toolbar {{
      display: flex;
      flex-wrap: wrap;
      align-items: end;
      gap: 12px;
      background: var(--fill);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 20px;
    }}
    label {{
      display: grid;
      gap: 6px;
      min-width: min(420px, 100%);
      font-size: 13px;
      color: var(--muted);
    }}
    select {{
      height: 40px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0 10px;
      font: inherit;
      color: var(--ink);
      background: #ffffff;
    }}
    button, .secondary {{
      height: 40px;
      border-radius: 6px;
      padding: 0 14px;
      border: 1px solid transparent;
      font: inherit;
      font-weight: 650;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      white-space: nowrap;
    }}
    button {{
      background: var(--accent);
      color: #ffffff;
      cursor: pointer;
    }}
    button:hover {{ background: var(--accent-strong); }}
    .secondary {{
      background: #ffffff;
      color: var(--accent-strong);
      border-color: var(--line);
    }}
    .status {{
      color: var(--accent-strong);
      margin: 0 0 12px;
    }}
    .notice {{
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--fill);
    }}
    .summary {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }}
    .summary div {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #ffffff;
    }}
    .summary span {{
      display: block;
      font-size: 13px;
      color: var(--muted);
      margin-bottom: 6px;
    }}
    .summary strong {{
      font-size: 18px;
      word-break: break-word;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      border: 1px solid var(--line);
    }}
    th, td {{
      text-align: left;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      font-size: 15px;
    }}
    th {{
      background: var(--fill);
      color: var(--muted);
      font-size: 13px;
      font-weight: 700;
    }}
    td a {{
      color: var(--accent-strong);
      font-weight: 700;
    }}
    .warnings {{
      color: var(--warn);
      margin-top: 16px;
    }}
    @media (max-width: 720px) {{
      main {{ width: min(100% - 20px, 1040px); margin: 20px auto; }}
      header {{ display: block; }}
      .summary {{ grid-template-columns: 1fr; }}
      table {{ display: block; overflow-x: auto; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>{STORE_NAME}本機日曆工具</h1>
        <p>Excel 轉成可手動匯入 Google 日曆的 ICS / CSV。</p>
      </div>
    </header>
    {f'<p class="status">{html.escape(status)}</p>' if status else ''}
    {workbook_control}
    {result_html}
  </main>
</body>
</html>
"""


class CalendarAppHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            params = urllib.parse.parse_qs(parsed.query)
            status = "已產生最新檔案。" if params.get("generated") else ""
            self.send_html(render_page(status))
            return
        if parsed.path.startswith("/exports/"):
            self.send_export(parsed.path)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        if self.path != "/generate":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length).decode("utf-8")
        params = urllib.parse.parse_qs(body)
        workbook_name = params.get("workbook", [""])[0]
        workbook = APP_DIR / workbook_name
        if not workbook.exists() or workbook.suffix.lower() != ".xlsx":
            self.send_html(render_page("找不到選取的 Excel 班表。"), status=HTTPStatus.BAD_REQUEST)
            return

        export_for_workbook(workbook)
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", "/?generated=1")
        self.end_headers()

    def send_html(self, content: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        payload = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_export(self, url_path: str) -> None:
        relative = urllib.parse.unquote(url_path.removeprefix("/exports/"))
        target = (EXPORT_DIR / relative).resolve()
        if not str(target).startswith(str(EXPORT_DIR.resolve())) or not target.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        payload = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(target.name)}")
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args: Any) -> None:
        return


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the local shift calendar app.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), CalendarAppHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"本機日曆工具已啟動：{url}")
    print("按 Ctrl+C 可以停止。")
    if not args.no_browser:
        webbrowser.open(url)
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
