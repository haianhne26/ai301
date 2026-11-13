"""Flask Localhost Web Server (Port 5000) for DaiThanh PartnerAI."""

from __future__ import annotations

import io
import json
from pathlib import Path
from typing import Any

from flask import Flask, Response, jsonify, make_response, render_template, request, send_file

from agent.compliance_checker import ComplianceChecker
from agent.excel_workflow import ExcelWorkflowManager
from agent.extractor import SAMPLE_PRESETS
from agent.pipeline_manager import DATA_DIR, RUNTIME_DATA_DIR, PipelineManager, SAMPLE_FILES_DIR

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["JSON_AS_ASCII"] = False
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["MAX_CONTENT_LENGTH"] = 128 * 1024 * 1024  # 128 MB upload cap

manager = PipelineManager()
excel_manager = ExcelWorkflowManager(manager, RUNTIME_DATA_DIR / "excel_jobs.json", SAMPLE_FILES_DIR)


@app.errorhandler(413)
def handle_payload_too_large(_err: Any) -> Any:
    return jsonify({"error": "Tệp hoặc dữ liệu tải lên vượt quá giới hạn dung lượng 128 MB."}), 413


def _compact_processed_row(pr: dict[str, Any]) -> dict[str, Any]:
    """Project a processed/KO Excel row to the compact UI payload."""
    prof = pr.get("profile") or {}
    legal = prof.get("legal") or {}
    loc = prof.get("location") or {}
    comm = prof.get("commercial") or {}
    war = prof.get("warranty") or {}
    fin = prof.get("financial") or {}
    sc = pr.get("scoring_report") or {}
    return {
        "row_num": pr.get("row_num"),
        "dossier_id": pr.get("dossier_id"),
        "include_in_export": bool(pr.get("include_in_export")),
        "auto_approved": bool(pr.get("auto_approved")),
        "batch_duplicate_mst": bool(pr.get("batch_duplicate_mst")),
        "anomalies": pr.get("anomalies") or [],
        "profile": {
            "legal": {
                "company_name": legal.get("company_name", ""),
                "tax_code": legal.get("tax_code", ""),
            },
            "location": {
                "region": loc.get("region", ""),
                "district": loc.get("district", ""),
            },
            "commercial": {
                "monthly_commitment_billion_vnd": comm.get("monthly_commitment_billion_vnd", 0),
            },
            "warranty": {
                "technicians_count": war.get("technicians_count", 0),
            },
            "financial": {
                "payment_term_days": fin.get("payment_term_days", 0),
                "credit_limit_requested_million_vnd": fin.get("credit_limit_requested_million_vnd", 0),
            },
        },
        "scoring_report": {
            "final_score": sc.get("final_score", 0),
            "tier": sc.get("tier", "Tier D"),
        },
    }


def _serialize_excel_job(job: dict[str, Any] | None) -> dict[str, Any] | None:
    """Return compact UI-safe Excel workflow payload so thousands of rows serialize cleanly."""
    if not job:
        return None
    skip_keys = {"all_sheet_tables", "raw_rows", "report_html", "cleaned_rows", "processed_rows", "approved_rows", "ko_rows"}
    payload = {k: v for k, v in job.items() if k not in skip_keys}
    payload["has_report"] = bool(job.get("report_html"))
    payload["approved_count"] = len(job.get("approved_rows") or [])
    payload["approved_row_nums"] = job.get("approved_row_nums") or []
    payload["verify_alert"] = job.get("verify_alert") or {}
    payload["raw_rows_preview"] = (job.get("raw_rows") or [])[:25]
    payload["cleaned_rows"] = [
        {
            "row_num": r.get("row_num"),
            "included": bool(r.get("included")),
            "clean_status": r.get("clean_status"),
            "cleaning_actions": r.get("cleaning_actions") or [],
            "row_errors": r.get("row_errors") or [],
            "profile": {
                "legal": {
                    "company_name": (r.get("profile") or {}).get("legal", {}).get("company_name", ""),
                    "tax_code": (r.get("profile") or {}).get("legal", {}).get("tax_code", ""),
                },
                "location": {
                    "region": (r.get("profile") or {}).get("location", {}).get("region", ""),
                    "district": (r.get("profile") or {}).get("location", {}).get("district", ""),
                },
                "warranty": {
                    "technicians_count": (r.get("profile") or {}).get("warranty", {}).get("technicians_count", 0),
                },
                "financial": {
                    "payment_term_days": (r.get("profile") or {}).get("financial", {}).get("payment_term_days", 0),
                },
            },
        }
        for r in (job.get("cleaned_rows") or [])
    ]
    payload["ko_rows"] = [_compact_processed_row(pr) for pr in (job.get("ko_rows") or [])]
    payload["processed_rows"] = [_compact_processed_row(pr) for pr in (job.get("processed_rows") or [])]
    return payload


def _load_server_config() -> dict[str, Any]:
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _clear_session_workspace_data() -> None:
    """Clear transactional dossiers & Excel workflow jobs while retaining configs/policies/partners."""
    manager.clear_all_dossiers()
    excel_manager.clear_all_jobs()


@app.route("/", methods=["GET"])
def index() -> Response:
    _clear_session_workspace_data()
    resp = make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/api/session/reset", methods=["POST"])
def api_session_reset() -> Any:
    _clear_session_workspace_data()
    return jsonify({
        "reset": True,
        "latest_excel_job": None,
        "dossiers": [],
        "analytics": manager.get_dashboard_analytics(),
    })


@app.route("/api/bootstrap", methods=["GET"])
def api_bootstrap() -> Any:
    sample_files = [
        {
            "key": key,
            "label": item["label"],
            "badge": item["badge"],
            "filename": item["filename"],
            "file_type": "DOCX" if item["filename"].endswith(".docx") else "PDF",
        }
        for key, item in SAMPLE_PRESETS.items()
    ]
    return jsonify({
        "llm_status": manager.llm.get_status(),
        "policies": manager.policies,
        "existing_partners": manager.existing_partners,
        "presets": SAMPLE_PRESETS,
        "sample_files": sample_files,
        "excel_sample_files": [
            {"filename": "mau_nhap_excel_chuan.xlsx", "label": "Mẫu chuẩn (3 đối tác)", "tag": "Sẵn sàng"},
            {"filename": "mau_nhap_excel_lam_sach_va_loi.xlsx", "label": "Mẫu Test Làm sạch & Lỗi", "tag": "2 Sheets"},
        ],
        "latest_excel_job": _serialize_excel_job(excel_manager.get_latest_job()),
        "dossiers": manager.list_dossier_summaries(),
        "analytics": manager.get_dashboard_analytics(),
    })


@app.route("/api/dossiers", methods=["GET"])
def api_list_dossiers() -> Any:
    return jsonify({
        "dossiers": manager.list_dossier_summaries(),
        "analytics": manager.get_dashboard_analytics(),
    })


@app.route("/api/dossiers/<dossier_id>", methods=["GET"])
def api_get_dossier(dossier_id: str) -> Any:
    dossier = manager.get_dossier(dossier_id)
    if not dossier:
        return jsonify({"error": f"Không tìm thấy hồ sơ {dossier_id}"}), 404
    return jsonify({"dossier": dossier})


@app.route("/api/dossiers/<dossier_id>", methods=["DELETE"])
def api_delete_dossier(dossier_id: str) -> Any:
    deleted = manager.delete_dossier(dossier_id)
    if not deleted:
        return jsonify({"error": f"Không tìm thấy hồ sơ {dossier_id}"}), 404
    return jsonify({
        "deleted": True,
        "dossiers": manager.list_dossier_summaries(),
        "analytics": manager.get_dashboard_analytics(),
    })


@app.route("/api/step1/extract", methods=["POST"])
def api_step1_extract() -> Any:
    try:
        if "file" in request.files and request.files["file"].filename:
            uploaded = request.files["file"]
            filename = uploaded.filename or "upload.txt"
            allowed_ext = (".pdf", ".docx", ".txt")
            if not filename.lower().endswith(allowed_ext):
                return jsonify({"error": "Chỉ hỗ trợ định dạng .PDF, .DOCX hoặc .TXT."}), 400
            file_bytes = uploaded.read()
            result = manager.extract_preview(filename=filename, file_bytes=file_bytes)
        else:
            payload = request.get_json(silent=True) or {}
            raw_text = str(payload.get("raw_text", "")).strip()
            result = manager.extract_preview(raw_text=raw_text)
        return jsonify(result)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/api/step1/submit", methods=["POST"])
def api_step1_submit() -> Any:
    payload = request.get_json(silent=True) or {}
    profile = payload.get("profile")
    if not isinstance(profile, dict):
        return jsonify({"error": "Thiếu dữ liệu hồ sơ đối tác (profile)."}), 400
    dossier = manager.create_dossier(
        profile=profile,
        intake_mode=str(payload.get("intake_mode", "form")),
        source_info=str(payload.get("source_info", "Biểu mẫu trực tuyến")),
        raw_text=str(payload.get("raw_text", "")),
    )
    return jsonify({
        "dossier": dossier,
        "dossiers": manager.list_dossier_summaries(),
        "analytics": manager.get_dashboard_analytics(),
    }), 201


@app.route("/api/dossiers/<dossier_id>/step2", methods=["POST"])
def api_run_step2(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        dossier = manager.run_compliance_step(dossier_id, payload.get("profile"))
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404


@app.route("/api/dossiers/<dossier_id>/step", methods=["POST"])
def api_advance_step(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        target_step = int(payload.get("step", 2))
        dossier = manager.advance_to_step(dossier_id, target_step)
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except (KeyError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/api/dossiers/<dossier_id>/step3/verify", methods=["POST"])
def api_step3_verify(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        dossier = manager.update_verification_and_rescore(dossier_id, payload)
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404


@app.route("/api/dossiers/<dossier_id>/step5/assess", methods=["POST"])
def api_step5_assess(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        force_refresh = bool(payload.get("force_refresh", False))
        dossier = manager.generate_strategic_assessment(dossier_id, force_refresh=force_refresh)
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404


@app.route("/api/dossiers/<dossier_id>/step5/decide", methods=["POST"])
def api_step5_decide(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        decision = str(payload.get("decision", "")).strip()
        conflict_plan = str(payload.get("conflict_resolution_plan", "")).strip()
        director_notes = str(payload.get("director_notes", "")).strip()
        dossier = manager.submit_director_decision(
            dossier_id=dossier_id,
            decision=decision,
            conflict_resolution_plan=conflict_plan,
            director_notes=director_notes,
        )
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except (KeyError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/api/dossiers/<dossier_id>/step6/email", methods=["POST"])
def api_step6_email(dossier_id: str) -> Any:
    try:
        payload = request.get_json(silent=True) or {}
        dossier = manager.update_handover_email(
            dossier_id=dossier_id,
            subject=payload.get("subject"),
            body=payload.get("body"),
            regenerate=bool(payload.get("regenerate", False)),
        )
        return jsonify({
            "dossier": dossier,
            "dossiers": manager.list_dossier_summaries(),
            "analytics": manager.get_dashboard_analytics(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404


@app.route("/api/export/excel", methods=["GET"])
def api_export_excel() -> Any:
    xlsx_bytes = manager.export_excel_bytes()
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="DaiThanh_DanhSach_DoiTac_PhanPhoi.xlsx",
    )


@app.route("/api/samples/<path:filename>", methods=["GET"])
def api_download_sample(filename: str) -> Any:
    safe_dir = SAMPLE_FILES_DIR.resolve()
    target = (safe_dir / filename).resolve()
    if target.parent != safe_dir or not target.exists():
        return jsonify({"error": "Không tìm thấy tệp mẫu."}), 404
    return send_file(str(target), as_attachment=True, download_name=target.name)


# ----------------- EXCEL BATCH WORKFLOW -----------------


@app.route("/api/excel/latest", methods=["GET"])
def api_excel_latest() -> Any:
    job = excel_manager.get_latest_job()
    return jsonify({"job": _serialize_excel_job(job)})


@app.route("/api/excel/jobs/<job_id>", methods=["GET"])
def api_excel_get_job(job_id: str) -> Any:
    job = excel_manager.get_job(job_id)
    if not job:
        return jsonify({"error": f"Không tìm thấy phiên nhập Excel {job_id}"}), 404
    return jsonify({"job": _serialize_excel_job(job)})


@app.route("/api/excel/upload", methods=["POST"])
def api_excel_upload() -> Any:
    try:
        requested_sheet: str | None = None
        filename = ""
        file_bytes = b""

        if "file" in request.files and request.files["file"].filename:
            up = request.files["file"]
            filename = up.filename or "upload.xlsx"
            file_bytes = up.read()
            requested_sheet = (request.form.get("sheet_name") or "").strip() or None
        else:
            payload = request.get_json(silent=True) or {}
            sample_name = (payload.get("sample_filename") or request.form.get("sample_filename") or "").strip()
            requested_sheet = (payload.get("sheet_name") or "").strip() or None
            if not sample_name:
                return jsonify({"error": "Vui lòng chọn tệp Excel (.xlsx/.csv) hoặc chọn tệp mẫu."}), 400
            safe_dir = SAMPLE_FILES_DIR.resolve()
            target = (safe_dir / sample_name).resolve()
            if target.parent != safe_dir or not target.exists():
                return jsonify({"error": f"Không tìm thấy tệp mẫu '{sample_name}'."}), 404
            filename = target.name
            file_bytes = target.read_bytes()

        job = excel_manager.start_from_upload(filename, file_bytes, requested_sheet=requested_sheet)
        return jsonify({"job": _serialize_excel_job(job)})
    except Exception as exc:
        return jsonify({"error": f"Lỗi nạp tệp Excel: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/sheet", methods=["POST"])
def api_excel_select_sheet(job_id: str) -> Any:
    payload = request.get_json(force=True, silent=True) or {}
    sheet_name = str(payload.get("sheet_name") or "").strip()
    if not sheet_name:
        return jsonify({"error": "Thiếu tên sheet cần chọn."}), 400
    try:
        job = excel_manager.select_sheet(job_id, sheet_name)
        return jsonify({"job": _serialize_excel_job(job)})
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"Lỗi chuyển sheet: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/cleaning", methods=["POST"])
def api_excel_cleaning(job_id: str) -> Any:
    payload = request.get_json(force=True, silent=True) or {}
    outcome = str(payload.get("outcome") or "OK").strip().upper()
    if outcome not in ("OK", "KO"):
        return jsonify({"error": "Giá trị outcome phải là OK hoặc KO."}), 400
    row_updates = payload.get("row_updates")
    try:
        job = excel_manager.submit_cleaning_decision(job_id, outcome, row_updates=row_updates)
        return jsonify({
            "job": _serialize_excel_job(job),
            "analytics": manager.get_dashboard_analytics(),
            "dossiers": manager.list_dossier_summaries(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"Lỗi xử lý bước Cleaning: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/retry-sheets", methods=["POST"])
def api_excel_retry_sheets(job_id: str) -> Any:
    try:
        job = excel_manager.return_error_to_sheets(job_id)
        return jsonify({"job": _serialize_excel_job(job)})
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404


@app.route("/api/excel/jobs/<job_id>/verify", methods=["POST"])
def api_excel_verify(job_id: str) -> Any:
    payload = request.get_json(force=True, silent=True) or {}
    outcome = str(payload.get("outcome") or "OK").strip().upper()
    if outcome not in ("OK", "KO"):
        return jsonify({"error": "Giá trị outcome phải là OK hoặc KO."}), 400
    feedback_note = str(payload.get("feedback_note") or "")
    row_toggles = payload.get("row_toggles")
    try:
        job = excel_manager.submit_verify_decision(
            job_id,
            outcome=outcome,
            feedback_note=feedback_note,
            row_toggles=row_toggles if isinstance(row_toggles, dict) else None,
        )
        return jsonify({
            "job": _serialize_excel_job(job),
            "analytics": manager.get_dashboard_analytics(),
            "dossiers": manager.list_dossier_summaries(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Lỗi duyệt & xuất dữ liệu: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/feedback", methods=["POST"])
def api_excel_feedback(job_id: str) -> Any:
    payload = request.get_json(force=True, silent=True) or {}
    try:
        job = excel_manager.apply_feedback_and_reprocess(
            job_id,
            note=str(payload.get("note") or ""),
            exclude_bad_debt=bool(payload.get("exclude_bad_debt", False)),
            clamp_payment_term_30=bool(payload.get("clamp_payment_term_30", False)),
            clamp_credit_500=bool(payload.get("clamp_credit_500", False)),
            exclude_tier_d=bool(payload.get("exclude_tier_d", False)),
            fix_ktv_minimum=bool(payload.get("fix_ktv_minimum", False)),
            approve_override_warnings=bool(payload.get("approve_override_warnings", False)),
            excluded_rows=payload.get("excluded_rows") if isinstance(payload.get("excluded_rows"), list) else None,
        )
        return jsonify({
            "job": _serialize_excel_job(job),
            "analytics": manager.get_dashboard_analytics(),
            "dossiers": manager.list_dossier_summaries(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"Lỗi xử lý Feedback: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/finish", methods=["POST"])
def api_excel_finish(job_id: str) -> Any:
    try:
        job = excel_manager.dismiss_ko_and_complete(job_id)
        return jsonify({
            "job": _serialize_excel_job(job),
            "analytics": manager.get_dashboard_analytics(),
            "dossiers": manager.list_dossier_summaries(),
        })
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"Lỗi hoàn tất phiên Excel: {exc}"}), 500


@app.route("/api/excel/jobs/<job_id>/report.html", methods=["GET"])
def api_excel_report_html(job_id: str) -> Any:
    job = excel_manager.get_job(job_id)
    if not job or not job.get("report_html"):
        return Response("<h3>Chưa có báo cáo HTML cho phiên làm việc này.</h3>", status=404, mimetype="text/html; charset=utf-8")
    resp = Response(job["report_html"], status=200, mimetype="text/html; charset=utf-8")
    if request.args.get("download") == "1":
        resp.headers["Content-Disposition"] = f'attachment; filename="BaoCao_DaiLy_{job_id}.html"'
    return resp


@app.route("/api/geofence/preview", methods=["POST"])
def api_geofence_preview() -> Any:
    payload = request.get_json(silent=True) or {}
    lat = float(payload.get("latitude", 21.0285))
    lng = float(payload.get("longitude", 105.8542))
    district = str(payload.get("district", ""))
    province = str(payload.get("province", ""))
    checker = ComplianceChecker(manager.policies, manager.existing_partners)
    dummy_profile = {
        "location": {
            "latitude": lat,
            "longitude": lng,
            "district": district,
            "province": province,
        }
    }
    geo = checker._check_geofencing(dummy_profile)
    return jsonify(geo)


@app.route("/api/settings/gemini", methods=["POST"])
def api_settings_gemini() -> Any:
    payload = request.get_json(silent=True) or {}
    action = payload.get("action", "save")
    if action == "save":
        api_key = str(payload.get("api_key", ""))
        model = str(payload.get("model", "")).strip() or None
        status = manager.llm.update_api_key(api_key, model=model)
        return jsonify({"status": status})
    if action == "test":
        return jsonify(manager.llm.test_connection())
    return jsonify({"error": "Hành động không hợp lệ."}), 400


if __name__ == "__main__":
    cfg = _load_server_config()
    host = str(cfg.get("HOST", "0.0.0.0"))
    port = int(cfg.get("PORT", 5000))
    debug = bool(cfg.get("DEBUG", False))
    print(f"DaiThanh PartnerAI Server is running at http://localhost:{port}")
    app.run(host=host, port=port, debug=debug)
