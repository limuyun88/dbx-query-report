#!/usr/bin/env python3
"""Render already verified DBX aggregate results without database access."""

from __future__ import annotations

import argparse
import html
import io
import json
import os
import re
import sys
import tempfile
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

MAX_INPUT_BYTES = 5_000_000
MAX_ROWS = 10_000
MAX_DATASETS = 20
MAX_COLUMNS = 30
TYPES = {"text", "integer", "decimal", "percent"}


def clean_text(value: object) -> str:
    """Remove characters XML 1.0 cannot represent."""
    text = str(value)
    return "".join(
        char for char in text
        if ord(char) in (9, 10, 13)
        or 32 <= ord(char) <= 0xD7FF
        or 0xE000 <= ord(char) <= 0xFFFD
        or 0x10000 <= ord(char) <= 0x10FFFF
    )


def xml_text(value: object) -> str:
    return xml_escape(clean_text(value), {'"': "&quot;"})


def cell_ref(column: int, row: int) -> str:
    letters = ""
    while column:
        column, digit = divmod(column - 1, 26)
        letters = chr(65 + digit) + letters
    return f"{letters}{row}"


def number(value: object, kind: str) -> Decimal:
    if isinstance(value, bool) or not isinstance(value, (int, float, str, Decimal)):
        raise ValueError(f"{kind} 单元格不是数值")
    try:
        parsed = Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(f"{kind} 单元格不是有效数值") from exc
    if not parsed.is_finite():
        raise ValueError("不能导出 NaN 或无穷大")
    if kind == "integer" and parsed != parsed.to_integral_value():
        raise ValueError("integer 列包含非整数")
    if abs(parsed) >= Decimal("1e15") or len(parsed.normalize().as_tuple().digits) > 15:
        raise ValueError("数值超出 Excel 可可靠表示的 15 位有效数字；请改为文本")
    return parsed


def validate(payload: object) -> dict:
    if not isinstance(payload, dict):
        raise ValueError("输入必须是 JSON 对象")
    for key in ("title", "source", "period", "definition"):
        if not isinstance(payload.get(key), str) or not payload[key].strip():
            raise ValueError(f"缺少非空文本字段 {key}")
    if payload.get("complete") is not True or payload.get("executed") is not True:
        raise ValueError("只接受已执行且已确认未截断的完整结果")
    sql = payload.get("sql")
    if isinstance(sql, str):
        sql = [sql]
    if not isinstance(sql, list) or not sql or any(not isinstance(s, str) or not s.strip() for s in sql):
        raise ValueError("sql 必须是非空 SQL 字符串或字符串数组")
    datasets = payload.get("datasets")
    if not isinstance(datasets, list) or not 1 <= len(datasets) <= MAX_DATASETS:
        raise ValueError("datasets 数量须为 1 至 20")
    names = set()
    for dataset in datasets:
        if not isinstance(dataset, dict) or not isinstance(dataset.get("title"), str) or not dataset["title"].strip():
            raise ValueError("结果集缺少标题")
        cols, rows = dataset.get("columns"), dataset.get("rows")
        if not isinstance(cols, list) or not 1 <= len(cols) <= MAX_COLUMNS:
            raise ValueError("每个结果集须有 1 至 30 列")
        if not isinstance(rows, list) or len(rows) > MAX_ROWS:
            raise ValueError("每个结果集最多 10000 行")
        keys = set()
        for col in cols:
            if not isinstance(col, dict) or not all(isinstance(col.get(k), str) and col[k] for k in ("key", "label", "type")):
                raise ValueError("列定义缺少 key、label 或 type")
            if col["key"] in keys or col["type"] not in TYPES:
                raise ValueError("列 key 重复或类型无效")
            keys.add(col["key"])
        for row in rows:
            if not isinstance(row, dict) or set(row) != keys:
                raise ValueError("结果行的列必须与 columns 完全一致")
            for col in cols:
                value = row[col["key"]]
                if value is None:
                    continue
                if col["type"] != "text":
                    number(value, col["type"])
                elif not isinstance(value, (str, int, float)):
                    raise ValueError("text 列包含不支持的值")
        chart = dataset.get("chart")
        if chart is not None:
            if not isinstance(chart, dict) or chart.get("category") not in keys or chart.get("value") not in keys:
                raise ValueError("chart 必须引用现有列")
            if "group" in chart and chart["group"] not in keys:
                raise ValueError("chart.group 必须引用现有列")
            kinds = {col["key"]: col["type"] for col in cols}
            if kinds[chart["value"]] not in ("integer", "decimal", "percent"):
                raise ValueError("chart.value 须为数值列")
        if "note" in dataset and not isinstance(dataset["note"], str):
            raise ValueError("note 必须是文本")
        raw_name = re.sub(r"[\\/*?:\[\]]", "_", dataset["title"]).strip("' ")[:31] or "Sheet"
        candidate = raw_name
        suffix = 2
        while candidate.casefold() in names:
            ending = f"_{suffix}"
            candidate = raw_name[: 31 - len(ending)] + ending
            suffix += 1
        dataset["_sheet_name"] = candidate
        names.add(candidate.casefold())
    payload["sql"] = sql
    return payload


def xlsx_cell(value: object, kind: str, column: int, row: int, header: bool = False) -> str:
    ref = cell_ref(column, row)
    if value is None:
        return f'<c r="{ref}"/>'
    if kind == "text":
        style = ' s="1"' if header else ""
        return f'<c r="{ref}" t="inlineStr"{style}><is><t xml:space="preserve">{xml_text(value)}</t></is></c>'
    style = {"integer": 2, "decimal": 3, "percent": 4}[kind]
    return f'<c r="{ref}" s="{style}"><v>{number(value, kind)}</v></c>'


def xlsx_sheet(dataset: dict, payload: dict) -> bytes:
    cols, rows = dataset["columns"], dataset["rows"]
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             f'<dimension ref="A1:{cell_ref(len(cols), max(7, 7 + len(rows)))}"/>',
             '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>',
             '<cols>']
    for i, col in enumerate(cols, 1):
        width = 28 if col["type"] == "text" else 19
        lines.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    lines += ['</cols>', '<sheetData>']
    metadata = [
        (1, dataset["title"]),
        (2, "来源：" + payload["source"]),
        (3, "期间：" + payload["period"]),
        (4, "口径：" + payload["definition"]),
    ]
    for row_idx, value in metadata:
        lines.append(f'<row r="{row_idx}">{xlsx_cell(value, "text", 1, row_idx)}</row>')
    lines.append('<row r="6">' + "".join(xlsx_cell(col["label"], "text", i, 6, True) for i, col in enumerate(cols, 1)) + '</row>')
    for row_idx, row in enumerate(rows, 7):
        lines.append(f'<row r="{row_idx}">' + "".join(
            xlsx_cell(row[col["key"]], col["type"], i, row_idx)
            for i, col in enumerate(cols, 1)
        ) + '</row>')
    if not rows:
        lines.append('<row r="7">' + xlsx_cell("无符合条件记录", "text", 1, 7) + '</row>')
    if dataset.get("note"):
        row_idx = 8 + len(rows)
        lines.append(f'<row r="{row_idx}">{xlsx_cell("说明：" + dataset["note"], "text", 1, row_idx)}</row>')
    lines += ['</sheetData>', '</worksheet>']
    return "".join(lines).encode("utf-8")


def build_xlsx(payload: dict) -> bytes:
    out = io.BytesIO()
    sheets = payload["datasets"]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as book:
        overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]
        for i in range(1, len(sheets) + 1):
            overrides.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        book.writestr("[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides) + '</Types>')
        book.writestr("_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
        workbook = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>',
        ]
        relationships = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ]
        for i, dataset in enumerate(sheets, 1):
            workbook.append(f'<sheet name="{xml_text(dataset["_sheet_name"])}" sheetId="{i}" r:id="rId{i}"/>')
            relationships.append(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
            book.writestr(f"xl/worksheets/sheet{i}.xml", xlsx_sheet(dataset, payload))
        workbook.append('</sheets></workbook>')
        relationships.append(
            f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        relationships.append('</Relationships>')
        book.writestr("xl/workbook.xml", "".join(workbook))
        book.writestr("xl/_rels/workbook.xml.rels", "".join(relationships))
        book.writestr("xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00########"/>'
            '<numFmt numFmtId="165" formatCode="0.00%"/></numFmts>'
            '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font>'
            '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts>'
            '<fills count="3"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF243C5A"/><bgColor indexed="64"/></patternFill></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/>'
            '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    return out.getvalue()


def display(value: object, kind: str) -> str:
    if value is None:
        return "—"
    if kind == "text":
        return str(value)
    numeric = number(value, kind)
    if kind == "integer":
        return f"{numeric:,.0f}"
    if kind == "percent":
        return f"{numeric * 100:,.2f}%"
    return f"{numeric:,.2f}"


def chart_html(dataset: dict) -> str:
    chart = dataset.get("chart")
    if not chart:
        return ""
    kinds = {col["key"]: col["type"] for col in dataset["columns"]}
    grouped = {}
    for row in dataset["rows"]:
        if row[chart["value"]] is not None:
            group = str(row[chart["group"]]) if chart.get("group") else ""
            grouped.setdefault(group, []).append(
                (row[chart["category"]], number(row[chart["value"]], kinds[chart["value"]]))
            )
    if not grouped:
        return ""
    scale = max(abs(value) for pairs in grouped.values() for _, value in pairs) or Decimal(1)
    parts = []
    for group, pairs in grouped.items():
        bars = []
        for category, value in pairs:
            width = min(100, float(abs(value) / scale * 100))
            color = "#b34747" if value < 0 else "#5079b8"
            bars.append(
                '<div class="barrow"><span>' + html.escape(str(category)) + '</span>'
                '<div class="track"><i style="width:' + f"{width:.2f}" + '%;background:' + color + '"></i></div>'
                '<strong>' + html.escape(display(value, kinds[chart["value"]])) + '</strong></div>'
            )
        heading = "<h3>" + html.escape(group) + "</h3>" if chart.get("group") else ""
        parts.append(heading + '<div class="bars" aria-label="数据条形图，精确值见下表">' + "".join(bars) + '</div>')
    return "".join(parts)


def build_html(payload: dict) -> str:
    sections = []
    for dataset in payload["datasets"]:
        headings = "".join("<th>" + html.escape(col["label"]) + "</th>" for col in dataset["columns"])
        body = "".join(
            "<tr>" + "".join(
                "<td>" + html.escape(display(row[col["key"]], col["type"])) + "</td>"
                for col in dataset["columns"]
            ) + "</tr>"
            for row in dataset["rows"]
        )
        if not body:
            body = f'<tr><td colspan="{len(dataset["columns"])}">无符合条件记录</td></tr>'
        note = '<p class="note">' + html.escape(dataset["note"]) + '</p>' if dataset.get("note") else ""
        sections.append(
            "<section><h2>" + html.escape(dataset["title"]) + "</h2>"
            + chart_html(dataset) + '<div class="scroll"><table><thead><tr>' + headings
            + "</tr></thead><tbody>" + body + "</tbody></table></div>" + note + "</section>"
        )
    css = (
        "body{margin:0;background:#f3f5f8;color:#243047;font:15px/1.55 Arial,sans-serif}"
        "main{max-width:1100px;margin:32px auto;padding:0 18px}h1{font-size:26px}h2{font-size:20px}"
        "section{background:white;border:1px solid #dce3ec;border-radius:12px;padding:22px;margin:20px 0}"
        ".meta,.note{color:#54647a}.scroll{overflow-x:auto}table{border-collapse:collapse;width:100%;min-width:520px}"
        "th,td{padding:10px 12px;text-align:left;border-bottom:1px solid #e5eaf0;white-space:nowrap}"
        "th{background:#edf2f7}tbody tr:nth-child(even){background:#f8fafc}"
        ".barrow{display:grid;grid-template-columns:minmax(100px,2fr) minmax(80px,3fr) auto;gap:10px;align-items:center;margin:9px 0}"
        ".barrow span{overflow-wrap:anywhere}.barrow strong{font-variant-numeric:tabular-nums;white-space:nowrap}"
        ".track{height:10px;background:#edf2f7;border-radius:6px}.track i{display:block;height:100%;border-radius:6px}"
        "@media(max-width:600px){main{padding:0 10px;margin:15px auto}section{padding:15px}"
        ".barrow{grid-template-columns:1fr auto}.track{grid-column:1/-1;grid-row:2}}"
    )
    return (
        '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<meta http-equiv="Content-Security-Policy" content="default-src &#39;none&#39;; style-src &#39;unsafe-inline&#39;">'
        "<title>" + html.escape(payload["title"]) + "</title><style>" + css + "</style></head><body><main>"
        "<h1>" + html.escape(payload["title"]) + "</h1>"
        '<p class="meta">来源：' + html.escape(payload["source"])
        + "　期间：" + html.escape(payload["period"]) + "</p>"
        + "".join(sections) + '<section><h2>统计口径</h2><p>'
        + html.escape(payload["definition"]) + "</p></section></main></body></html>"
    )


def export(payload: dict, destination: Path) -> None:
    if destination.exists():
        raise ValueError("输出目录已存在；请换一个新目录，避免覆盖")
    if not destination.parent.is_dir():
        raise ValueError("输出目录的父目录不存在")
    xlsx = build_xlsx(payload)
    page = build_html(payload).encode("utf-8")
    sql = ("\n\n".join(payload["sql"]) + "\n").encode("utf-8")
    with tempfile.TemporaryDirectory(prefix=".dbx-report-", dir=destination.parent) as temp:
        staging = Path(temp)
        (staging / "report.xlsx").write_bytes(xlsx)
        (staging / "report.html").write_bytes(page)
        (staging / "report.sql").write_bytes(sql)
        os.rename(staging, destination)


def main() -> int:
    parser = argparse.ArgumentParser(description="将已核实的 DBX 汇总结果导出为 XLSX、HTML 和 SQL")
    parser.add_argument("--input", required=True, type=Path, help="UTF-8 JSON 文件")
    parser.add_argument("--out-dir", required=True, type=Path, help="尚不存在的输出目录")
    args = parser.parse_args()
    try:
        if args.input.stat().st_size > MAX_INPUT_BYTES:
            raise ValueError("输入文件超过 5 MB")
        payload = validate(json.loads(args.input.read_text(encoding="utf-8")))
        export(payload, args.out_dir.resolve())
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"导出失败：{exc}", file=sys.stderr)
        return 1
    print(f"已生成：{args.out_dir.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
