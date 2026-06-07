import secrets
from datetime import datetime
import os
from pathlib import Path
from urllib.parse import quote

from flask import Flask, Response, jsonify, render_template, request

from analysis_engine import AnalysisOptions, analyze_store_data, extract_turnover_filter_options
from sku_splitter import transform_parent_sku_workbook
from xlsx_toolkit import SheetData, load_uploaded_workbook, workbook_to_xlsx_bytes


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

ANALYSIS_EXPORT_CACHE = {}
SKU_EXPORT_CACHE = {}


def build_default_form() -> dict:
    return {
        "group_filters": ["__all__"],
        "site_filter": "__all__",
        "min_sales_30d": "3",
        "min_turnover_rate": "10",
        "low_impressions_threshold": "100",
        "low_page_views_threshold": "10",
        "low_conversion_threshold": "0.02",
        "max_results": "30",
        "enable_ebay_scrape": "on",
        "enable_openai": "",
        "openai_model": "gpt-4.1-mini",
        "ebay_site_domain": "www.ebay.co.uk",
    }


def build_empty_filter_options() -> dict:
    return {"groups": [], "sites": []}


def build_sku_splitter_form() -> dict:
    return {"preview_limit": "120"}


def coerce_float(value: str, fallback: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return fallback


def coerce_int(value: str, fallback: int) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return fallback


def build_options(form) -> AnalysisOptions:
    group_filters = form.getlist("group_filters") if hasattr(form, "getlist") else []
    if not group_filters:
        group_filters = ["__all__"]
    if "__all__" in group_filters and len(group_filters) > 1:
        group_filters = [value for value in group_filters if value != "__all__"]

    return AnalysisOptions(
        group_filters=group_filters,
        site_filter=form.get("site_filter", "__all__"),
        min_sales_30d=coerce_float(form.get("min_sales_30d"), 3.0),
        min_turnover_rate=coerce_float(form.get("min_turnover_rate"), 10.0),
        low_impressions_threshold=coerce_float(form.get("low_impressions_threshold"), 100.0),
        low_page_views_threshold=coerce_float(form.get("low_page_views_threshold"), 10.0),
        low_conversion_threshold=coerce_float(form.get("low_conversion_threshold"), 0.02),
        max_results=coerce_int(form.get("max_results"), 30),
        enable_ebay_scrape=form.get("enable_ebay_scrape") == "on",
        enable_openai=form.get("enable_openai") == "on",
        openai_api_key=form.get("openai_api_key", ""),
        openai_model=form.get("openai_model", "gpt-4.1-mini"),
        ebay_site_domain=form.get("ebay_site_domain", "www.ebay.co.uk").strip() or "www.ebay.co.uk",
    )


def build_download_headers(filename: str) -> dict:
    safe_ascii_name = "download.xlsx"
    return {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": (
            f'attachment; filename="{safe_ascii_name}"; '
            f"filename*=UTF-8''{quote(filename)}"
        ),
    }


def get_default_downloads_dir() -> Path:
    return Path.home() / "Downloads"


def save_cached_sku_export(cached: dict) -> Path:
    downloads_dir = get_default_downloads_dir()
    downloads_dir.mkdir(parents=True, exist_ok=True)
    output_path = downloads_dir / cached["filename"]
    output_path.write_bytes(workbook_to_xlsx_bytes(cached["sheets"]))
    return output_path


def render_link_analysis_page(
    form=None,
    filter_options=None,
    results=None,
    error=None,
    export_token=None,
    generated_at=None,
):
    return render_template(
        "index.html",
        form=form or build_default_form(),
        filter_options=filter_options or build_empty_filter_options(),
        results=results,
        error=error,
        export_token=export_token,
        generated_at=generated_at,
    )


def render_sku_splitter_page(
    form=None,
    error=None,
    summary=None,
    preview_rows=None,
    export_token=None,
    generated_at=None,
    source_filename=None,
    saved_file_path=None,
):
    return render_template(
        "sku_splitter.html",
        form=form or build_sku_splitter_form(),
        error=error,
        summary=summary,
        preview_rows=preview_rows,
        export_token=export_token,
        generated_at=generated_at,
        source_filename=source_filename,
        saved_file_path=saved_file_path,
    )


@app.route("/", methods=["GET"])
def index():
    return render_sku_splitter_page()


@app.route("/link-analysis", methods=["GET"])
def link_analysis_index():
    return render_link_analysis_page()


@app.route("/filter-options", methods=["POST"])
def filter_options():
    try:
        turnover_file = request.files.get("turnover_file")
        if not turnover_file or not turnover_file.filename:
            raise ValueError("请先上传动销表。")
        turnover_workbook = load_uploaded_workbook(turnover_file)
        return jsonify(extract_turnover_filter_options(turnover_workbook))
    except Exception as exc:
        return jsonify({"error": str(exc), "groups": [], "sites": []}), 400


@app.route("/analyze", methods=["POST"])
def analyze():
    form_data = build_default_form()
    form_data.update(request.form.to_dict())
    form_data["group_filters"] = request.form.getlist("group_filters") or ["__all__"]
    filter_options = build_empty_filter_options()

    try:
        turnover_file = request.files.get("turnover_file")
        traffic_file = request.files.get("traffic_file")
        if not turnover_file or not turnover_file.filename:
            raise ValueError("请先上传动销表。")
        if not traffic_file or not traffic_file.filename:
            raise ValueError("请先上传流量报表。")

        turnover_workbook = load_uploaded_workbook(turnover_file)
        traffic_workbook = load_uploaded_workbook(traffic_file)
        filter_options = extract_turnover_filter_options(turnover_workbook)
        results = analyze_store_data(turnover_workbook, traffic_workbook, build_options(request.form))

        export_token = secrets.token_urlsafe(16)
        ANALYSIS_EXPORT_CACHE[export_token] = {
            "generated_at": datetime.now(),
            "summary": results["summary"],
            "rows": results["rows"],
            "export_rows": results["export_rows"],
        }

        return render_link_analysis_page(
            form=form_data,
            filter_options=filter_options,
            results=results,
            error=None,
            export_token=export_token,
            generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
    except Exception as exc:
        return render_link_analysis_page(
            form=form_data,
            filter_options=filter_options,
            results=None,
            error=str(exc),
            export_token=None,
            generated_at=None,
        )


@app.route("/export/<token>", methods=["GET"])
def export_report(token: str):
    cached = ANALYSIS_EXPORT_CACHE.get(token)
    if not cached:
        return Response("导出记录已失效，请重新分析。", status=404, content_type="text/plain; charset=utf-8")

    summary = cached["summary"]
    export_rows = cached["export_rows"]
    summary_rows = [
        ["指标", "数值"],
        ["动销筛选后 SKU 数", str(summary["turnover_count"])],
        ["流量主键记录数", str(summary["traffic_count"])],
        ["建议优化 SKU 数", str(summary["candidate_count"])],
        ["无流量数据", str(summary["no_traffic_count"])],
        ["流量偏低", str(summary["low_traffic_count"])],
        ["转化偏低", str(summary["low_conversion_count"])],
        ["动销工作表", str(summary["turnover_sheet"])],
        ["流量工作表", str(summary["traffic_sheet"])],
    ]
    for index, rule_line in enumerate(summary.get("rule_lines", []), start=1):
        summary_rows.append([f"规则 {index}", rule_line])

    workbook_bytes = workbook_to_xlsx_bytes(
        [
            SheetData(name="优化建议", rows=export_rows),
            SheetData(name="分析摘要", rows=summary_rows),
        ]
    )

    filename = f"store-link-diagnosis-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
    return Response(workbook_bytes, headers=build_download_headers(filename))


@app.route("/sku-splitter", methods=["GET", "POST"])
def sku_splitter():
    if request.method == "GET":
        return render_sku_splitter_page()

    form_data = build_sku_splitter_form()
    form_data.update(request.form.to_dict())

    try:
        source_file = request.files.get("source_file")
        if not source_file or not source_file.filename:
            raise ValueError("请先上传利润 Excel 文件。")

        preview_limit = max(20, min(coerce_int(request.form.get("preview_limit"), 120), 300))
        workbook = load_uploaded_workbook(source_file)
        result = transform_parent_sku_workbook(workbook, preview_limit=preview_limit)

        export_token = secrets.token_urlsafe(16)
        SKU_EXPORT_CACHE[export_token] = {
            "generated_at": datetime.now(),
            "sheets": result["sheets"],
            "filename": result["output_filename"],
        }

        form_data["preview_limit"] = str(preview_limit)
        return render_sku_splitter_page(
            form=form_data,
            error=None,
            summary=result["summary"],
            preview_rows=result["preview_rows"],
            export_token=export_token,
            generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            source_filename=source_file.filename,
        )
    except Exception as exc:
        return render_sku_splitter_page(
            form=form_data,
            error=str(exc),
            summary=None,
            preview_rows=None,
            export_token=None,
            generated_at=None,
            source_filename=None,
        )


@app.route("/sku-splitter/export/<token>", methods=["GET"])
def export_sku_splitter(token: str):
    cached = SKU_EXPORT_CACHE.get(token)
    if not cached:
        return Response("导出记录已失效，请重新处理文件。", status=404, content_type="text/plain; charset=utf-8")

    if request.args.get("download") != "1":
        saved_path = save_cached_sku_export(cached)
        html = f"""<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>文件已保存</title>
    <style>
      body {{
        margin: 0;
        font-family: 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
        background: #f6f1e7;
        color: #24180b;
      }}
      main {{
        max-width: 760px;
        margin: 48px auto;
        padding: 32px;
        background: #fffdf8;
        border: 1px solid rgba(74, 55, 26, 0.12);
        border-radius: 24px;
        box-shadow: 0 18px 48px rgba(63, 37, 12, 0.12);
      }}
      h1 {{ margin-top: 0; }}
      p {{ line-height: 1.7; }}
      code {{
        background: rgba(36, 24, 11, 0.06);
        padding: 2px 6px;
        border-radius: 6px;
        word-break: break-all;
      }}
      a {{
        display: inline-block;
        margin-top: 18px;
        color: #8d3c12;
        text-decoration: none;
        font-weight: 600;
      }}
    </style>
  </head>
  <body>
    <main>
      <h1>文件已保存到下载目录</h1>
      <p>由于当前应用内浏览器不支持直接下载，系统已经帮你把处理结果保存到本地：</p>
      <p><code>{saved_path}</code></p>
      <a href="/sku-splitter">返回处理页面</a>
    </main>
  </body>
</html>"""
        return Response(html, content_type="text/html; charset=utf-8")

    workbook_bytes = workbook_to_xlsx_bytes(cached["sheets"])
    filename = cached["filename"]
    return Response(workbook_bytes, headers=build_download_headers(filename))


@app.route("/sku-splitter/save/<token>", methods=["POST"])
def save_sku_splitter_to_downloads(token: str):
    cached = SKU_EXPORT_CACHE.get(token)
    if not cached:
        return jsonify({"error": "导出记录已失效，请重新处理文件。"}), 404

    output_path = save_cached_sku_export(cached)

    return jsonify(
        {
            "ok": True,
            "saved_path": str(output_path),
            "filename": cached["filename"],
        }
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5055")), debug=False)
