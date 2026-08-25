"""Small local web interface for selection notices and timetable snapshots."""

from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, redirect, render_template_string, request, url_for
from werkzeug.utils import secure_filename

from .notice import (
    REQUIRED_FIELDS,
    confirm_notice,
    fetch_notice_text,
    parse_notice,
    save_notice,
)
from .timetable import entries_to_dict, import_timetable


INDEX_TEMPLATE = """<!doctype html>
<html lang="zh-CN">
<head><meta charset="utf-8"><title>教务选课规划</title>
<style>body{font-family:system-ui;margin:2rem;max-width:70rem}section{border:1px solid #ddd;padding:1rem;margin:1rem 0}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:.4rem;text-align:left}.muted{color:#666}.error{color:#a00}</style>
</head>
<body>
<h1>教务选课规划</h1>
<section>
<h2>选课窗口</h2>
{% if notice %}
<p>状态：<strong>{{ notice.status }}</strong>；{{ notice.title or "未命名通知" }}</p>
<p>学期：{{ notice.term or "待补充" }}；类型：{{ notice.selection_type or "待补充" }}</p>
<p>时间：{{ notice.opens_at or "待补充" }} 至 {{ notice.closes_at or "待补充" }}</p>
{% if notice.missing_fields %}<p class="error">待补充：{{ notice.missing_fields|join("、") }}</p>{% endif %}
{% if notice.status != "confirmed" %}<form method="post" action="{{ url_for('confirm_notice_route') }}"><button>确认选课窗口</button></form>{% endif %}
{% else %}<p class="muted">尚未导入选课通知。</p>{% endif %}
<form method="post" action="{{ url_for('create_notice') }}">
<label>通知链接 <input name="source_url" size="60"></label><br>
<label>通知正文或人工录入<br><textarea name="text" rows="6" cols="80" required></textarea></label><br>
<button>导入通知</button>
</form>
</section>
<section>
<h2>当前课表</h2>
<form id="timetable-upload" method="post" action="{{ url_for('upload_timetable') }}" enctype="multipart/form-data">
<input type="file" name="timetable" accept=".xls,.xlsx" required><button>导入课表</button>
</form>
{% if timetable %}<p>学期：{{ timetable.term }}；课程数：{{ timetable.entries|length }}</p>
<p class="muted">当前快照：{{ timetable.source_name }}，导入时间：{{ timetable.imported_at }}</p>
<p class="muted">再次上传会替换当前课表快照，请勾选确认。</p>
<label><input type="checkbox" name="replace_existing" value="1" form="timetable-upload">确认替换当前课表</label>
<table><tr><th>课程</th><th>星期</th><th>节次</th><th>周次</th><th>单双周</th><th>地点</th></tr>
{% for item in timetable.entries %}<tr><td>{{ item.course_name }}</td><td>星期{{ item.weekday }}</td><td>{{ item.start_period }}-{{ item.end_period }}</td><td>{{ item.week_start }}-{{ item.week_end }}</td><td>{{ item.week_parity }}</td><td>{{ item.location }}</td></tr>{% endfor %}
</table>{% else %}<p class="muted">尚未导入课表。</p>{% endif %}
</section>
</body></html>"""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _read_json(path: Path) -> dict | None:
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _notice_for_view(path: Path) -> dict | None:
    notice = _read_json(path)
    if notice is None:
        return None
    return {
        **notice,
        "missing_fields": [field for field in REQUIRED_FIELDS if not notice.get(field)],
    }


def create_app(private_root: Path | str = ".private/academic-selection") -> Flask:
    root = Path(private_root)
    root.mkdir(parents=True, exist_ok=True)
    app = Flask(__name__)
    app.config["PRIVATE_ROOT"] = root

    @app.get("/")
    def index():
        notice_data = _notice_for_view(root / "selection-notice.json")
        timetable_data = _read_json(root / "current-timetable.json")
        return render_template_string(
            INDEX_TEMPLATE,
            notice=notice_data,
            timetable=timetable_data,
        )

    @app.post("/notices")
    def create_notice():
        text = request.form.get("text", "").strip()
        source_url = request.form.get("source_url", "").strip()
        if not text and source_url:
            try:
                text = fetch_notice_text(source_url)
            except (OSError, ValueError) as error:
                return f"通知链接读取失败，请粘贴正文：{error}", 400
        if not text:
            return "通知正文或通知链接不能为空", 400
        notice = parse_notice(
            text,
            source_url=source_url,
            source_kind="official" if source_url else "manual",
        )
        save_notice(root / "selection-notice.json", notice)
        return redirect(url_for("index"))

    @app.post("/notices/confirm")
    def confirm_notice_route():
        from .notice import load_notice

        path = root / "selection-notice.json"
        if not path.is_file():
            return "尚未导入选课通知", 404
        try:
            save_notice(path, confirm_notice(load_notice(path)))
        except ValueError as error:
            return str(error), 400
        return redirect(url_for("index"))

    @app.post("/timetable")
    def upload_timetable():
        upload = request.files.get("timetable")
        if upload is None or not upload.filename:
            return "请选择 XLS 或 XLSX 课表", 400
        filename = secure_filename(upload.filename)
        if Path(filename).suffix.lower() not in {".xls", ".xlsx"}:
            return "课表必须是 .xls 或 .xlsx 文件", 400
        current_snapshot = root / "current-timetable.json"
        if current_snapshot.is_file() and request.form.get("replace_existing") != "1":
            return "已有当前课表，请确认替换后再上传", 409
        root.joinpath("imports").mkdir(exist_ok=True)
        with tempfile.NamedTemporaryFile(
            dir=root / "imports", suffix=Path(filename).suffix, delete=False
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            upload.save(temporary_path)
            notice = _read_json(root / "selection-notice.json")
            expected_term = notice.get("term") if notice else None
            entries = import_timetable(temporary_path, expected_term=expected_term)
            snapshot = {
                "term": entries[0].term,
                "imported_at": _now(),
                "source_name": filename,
                "entries": entries_to_dict(entries),
            }
            current_snapshot.write_text(
                json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except (OSError, ValueError) as error:
            return str(error), 400
        finally:
            temporary_path.unlink(missing_ok=True)
        return redirect(url_for("index"))

    return app
