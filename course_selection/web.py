"""Local workbench for selection notices, timetable snapshots, and read-only results."""

from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, flash, redirect, render_template_string, request, url_for

from .notice import (
    REQUIRED_FIELDS,
    confirm_notice,
    fetch_notice_text,
    fetch_notice_text_in_browser,
    parse_notice,
    save_notice,
)
from .timetable import entries_to_dict, import_timetable

INDEX_TEMPLATE = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>选课工作台</title>
<style>
:root{--ink:#172033;--muted:#647089;--line:#dfe5ee;--paper:#fff;--wash:#f5f7fb;--accent:#275df5;--dark:#1644c7;--good:#087443;--bad:#b42318}*{box-sizing:border-box}body{margin:0;background:var(--wash);color:var(--ink);font:15px/1.55 system-ui,-apple-system,"Segoe UI","Microsoft YaHei",sans-serif}button,input,textarea{font:inherit}button{cursor:pointer;border:0;border-radius:10px;padding:10px 16px;background:var(--accent);color:white;font-weight:700}button:hover{background:var(--dark)}.secondary{background:#e9eefb;color:#2347a9}.shell{max-width:1180px;margin:auto;padding:28px 22px 56px}.top{display:flex;align-items:flex-end;justify-content:space-between;gap:20px;margin-bottom:26px}.kicker{color:var(--accent);font-weight:800;letter-spacing:.04em;font-size:12px}.top h1{font-size:32px;line-height:1.15;margin:5px 0}.subtitle,.muted{color:var(--muted)}.badge{padding:6px 10px;border-radius:99px;background:#eaf0ff;color:#2347a9;font-weight:700;font-size:13px;white-space:nowrap}.flash{padding:12px 14px;border-radius:12px;margin:12px 0;background:#eaf8f0;color:var(--good);border:1px solid #b5e1c8}.flash.error{background:#fff0ef;color:var(--bad);border-color:#f1b6b0}.grid{display:grid;grid-template-columns:1.05fr .95fr;gap:18px}.panel{background:var(--paper);border:1px solid var(--line);border-radius:16px;padding:22px;box-shadow:0 8px 24px #24375a0d}.wide{grid-column:1/-1}.panel-head{display:flex;justify-content:space-between;gap:15px;margin-bottom:17px}.panel h2{font-size:20px;margin:0 0 4px}.panel p{margin:7px 0}.status{font-weight:800;color:var(--accent)}.meta{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:15px 0}.meta div{background:var(--wash);border-radius:11px;padding:11px}.meta small{display:block;color:var(--muted);font-size:12px}.meta strong{display:block;margin-top:2px}label{display:block;font-weight:700;margin:12px 0 6px}input[type=url],input[type=file],textarea{width:100%;border:1px solid #cbd4e3;border-radius:10px;padding:10px;background:white;color:var(--ink)}textarea{resize:vertical}.actions{display:flex;flex-wrap:wrap;gap:9px;margin-top:12px}.hint{font-size:13px;color:var(--muted)}.replace{display:flex;align-items:center;gap:8px;font-weight:500}.replace input{width:16px;height:16px}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:12px;margin-top:16px}table{border-collapse:collapse;width:100%;min-width:680px;background:white}td,th{padding:11px 12px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap}th{font-size:12px;color:var(--muted);background:#f8faff}.empty{padding:24px 10px;text-align:center;color:var(--muted);background:var(--wash);border-radius:12px}code{background:#edf1f8;padding:2px 5px;border-radius:5px}@media(max-width:800px){.shell{padding:20px 14px 40px}.top{display:block}.badge{display:inline-block;margin-top:12px}.grid{grid-template-columns:1fr}.wide{grid-column:auto}.meta{grid-template-columns:1fr}.top h1{font-size:27px}}
</style></head><body><main class="shell"><header class="top"><div><div class="kicker">ACADEMIC SELECTION DESK</div><h1>选课工作台</h1><p class="subtitle">先确认窗口，再导入课表，最后只读查看可选课程。</p></div><span class="badge">本地运行 · 数据不上传</span></header>
{% with messages = get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="flash {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}
<div class="grid"><section class="panel"><div class="panel-head"><div><h2>01 · 选课窗口</h2><p class="muted">从学校通知中提取学期和开放时间。</p></div>{% if notice and notice.status == "confirmed" %}<span class="badge">已确认</span>{% endif %}</div>
{% if notice %}<p><strong>{{ notice.title or "未命名通知" }}</strong></p><div class="meta"><div><small>学期</small><strong>{{ notice.term or "待补充" }}</strong></div><div><small>类型</small><strong>{{ notice.selection_type or "待补充" }}</strong></div><div><small>时间</small><strong>{{ notice.opens_at or "待补充" }}<br>{{ notice.closes_at or "待补充" }}</strong></div></div>{% if notice.missing_fields %}<div class="flash error">还需要补充：{{ notice.missing_fields|join("、") }}</div>{% endif %}{% if notice.status != "confirmed" %}<form method="post" action="{{ url_for('confirm_notice_route') }}"><button>确认选课窗口</button></form>{% endif %}{% else %}<div class="empty">还没有选课通知。粘贴学校官网链接后即可导入。</div>{% endif %}
<form method="post" action="{{ url_for('create_notice') }}"><label>学校通知链接</label><input type="url" name="source_url" placeholder="https://…/page.html"><div class="actions"><button type="submit" formaction="{{ url_for('import_notice_browser') }}" class="secondary">用已登录浏览器读取</button><button type="submit">读取公开页面</button></div><label>或粘贴通知正文</label><textarea name="text" rows="4" placeholder="页面需要 WebVPN 登录时，可直接粘贴正文"></textarea><p class="hint">浏览器读取会复用成绩查询的登录会话，不会把认证信息交给普通 HTTP 请求。</p></form></section>
<section class="panel"><div class="panel-head"><div><h2>02 · 当前课表</h2><p class="muted">导入后用于冲突检查和智能规划。</p></div>{% if timetable %}<span class="badge">{{ timetable.entries|length }} 门记录</span>{% endif %}</div><form id="timetable-upload" method="post" action="{{ url_for('upload_timetable') }}" enctype="multipart/form-data"><label>学校导出的课表文件</label><input type="file" name="timetable" accept=".xls,.xlsx" required><div class="actions"><button>导入课表</button></div></form>
{% if timetable %}<div class="meta"><div><small>学期</small><strong>{{ timetable.term }}</strong></div><div><small>来源</small><strong>{{ timetable.source_name }}</strong></div><div><small>导入时间</small><strong>{{ timetable.imported_at }}</strong></div></div><label class="replace"><input type="checkbox" name="replace_existing" value="1" form="timetable-upload">确认替换当前课表</label><p class="hint">支持学校个人课表和班级课表的 .xls / .xlsx 文件。若提示学期不匹配，请先检查通知中的学期。</p><div class="table-wrap"><table><tr><th>课程</th><th>星期</th><th>节次</th><th>周次</th><th>单双周</th><th>地点</th></tr>{% for item in timetable.entries %}<tr><td>{{ item.course_name }}</td><td>星期{{ item.weekday }}</td><td>{{ item.start_period }}-{{ item.end_period }}</td><td>{{ item.week_start }}-{{ item.week_end }}</td><td>{{ item.week_parity }}</td><td>{{ item.location }}</td></tr>{% endfor %}</table></div>{% else %}<div class="empty">选择学校导出的 .xls 或 .xlsx 文件开始导入。</div>{% endif %}</section>
<section class="panel wide"><div class="panel-head"><div><h2>03 · 选课入口（只读）</h2><p class="muted">探测课程接口，不会点击选课、退课或提交。</p></div>{% if selection %}<span class="badge">{{ selection.status }}</span>{% endif %}</div>{% if selection %}<p>状态：<strong class="status">{{ selection.status }}</strong>；接口：{{ selection.request_url or "暂无" }}</p>{% if selection.sections %}<div class="table-wrap"><table><tr><th>课程</th><th>学分</th><th>教师</th><th>时间</th><th>容量</th><th>已选</th></tr>{% for item in selection.sections %}<tr><td>{{ item.name }}</td><td>{{ item.credits }}</td><td>{{ item.teacher }}</td><td>{{ item.time }}</td><td>{{ item.capacity }}</td><td>{{ item.selected }}</td></tr>{% endfor %}</table></div>{% else %}<div class="empty">当前没有可呈现的课程班结果。</div>{% endif %}{% else %}<div class="empty">尚未探索教务选课入口。运行 <code>uv run python -m course_selection explore-entry</code> 开始只读探测。</div>{% endif %}</section></div></main></body></html>"""


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
    return {**notice, "missing_fields": [field for field in REQUIRED_FIELDS if not notice.get(field)]}


def create_app(private_root: Path | str = ".private/academic-selection") -> Flask:
    root = Path(private_root)
    root.mkdir(parents=True, exist_ok=True)
    app = Flask(__name__)
    app.secret_key = "local-academic-selection"
    app.config["PRIVATE_ROOT"] = root

    @app.get("/")
    def index():
        return render_template_string(INDEX_TEMPLATE, notice=_notice_for_view(root / "selection-notice.json"), timetable=_read_json(root / "current-timetable.json"), selection=_read_json(root / "selection-entry.json"))

    @app.post("/notices")
    def create_notice():
        text = request.form.get("text", "").strip()
        source_url = request.form.get("source_url", "").strip()
        if not text and source_url:
            try:
                text = fetch_notice_text(source_url)
            except (OSError, ValueError) as error:
                flash(f"公开页面读取失败，请粘贴正文或改用已登录浏览器：{error}", "error")
                return redirect(url_for("index"))
        if not text:
            flash("请填写学校通知链接，或粘贴通知正文。", "error")
            return redirect(url_for("index"))
        save_notice(root / "selection-notice.json", parse_notice(text, source_url=source_url, source_kind="official" if source_url else "manual"))
        flash("通知已导入，请检查字段后确认选课窗口。", "success")
        return redirect(url_for("index"))

    @app.post("/notices/browser")
    def import_notice_browser():
        source_url = request.form.get("source_url", "").strip()
        if not source_url:
            flash("请先填写学校通知链接。", "error")
            return redirect(url_for("index"))
        try:
            text = fetch_notice_text_in_browser(source_url)
            save_notice(root / "selection-notice.json", parse_notice(text, source_url=source_url, source_kind="browser"))
            flash("已通过登录浏览器读取通知，请检查字段后确认。", "success")
        except Exception as error:
            flash(f"浏览器读取失败：{error}", "error")
        return redirect(url_for("index"))

    @app.post("/notices/confirm")
    def confirm_notice_route():
        from .notice import load_notice
        path = root / "selection-notice.json"
        if not path.is_file():
            flash("尚未导入选课通知。", "error")
            return redirect(url_for("index"))
        try:
            save_notice(path, confirm_notice(load_notice(path)))
        except ValueError as error:
            flash(str(error), "error")
            return redirect(url_for("index"))
        flash("选课窗口已确认。", "success")
        return redirect(url_for("index"))

    @app.post("/timetable")
    def upload_timetable():
        upload = request.files.get("timetable")
        if upload is None or not upload.filename:
            flash("请选择 XLS 或 XLSX 课表。", "error")
            return redirect(url_for("index"))
        original_name = Path(upload.filename).name
        suffix = Path(original_name).suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            flash("课表必须是 .xls 或 .xlsx 文件。", "error")
            return redirect(url_for("index"))
        current_snapshot = root / "current-timetable.json"
        if current_snapshot.is_file() and request.form.get("replace_existing") != "1":
            flash("已有当前课表；如需替换，请勾选“确认替换当前课表”。", "error")
            return redirect(url_for("index"))
        root.joinpath("imports").mkdir(exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=root / "imports", suffix=suffix, delete=False) as temporary:
            temporary_path = Path(temporary.name)
        try:
            upload.save(temporary_path)
            notice = _read_json(root / "selection-notice.json")
            expected_term = notice.get("term") if notice else None
            entries = import_timetable(temporary_path, expected_term=expected_term)
            snapshot = {"term": entries[0].term, "imported_at": _now(), "source_name": original_name, "entries": entries_to_dict(entries)}
            current_snapshot.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as error:
            flash(f"课表导入失败：{error}", "error")
            return redirect(url_for("index"))
        finally:
            temporary_path.unlink(missing_ok=True)
        flash(f"课表导入成功，共识别 {len(entries)} 条课程记录。", "success")
        return redirect(url_for("index"))

    return app
