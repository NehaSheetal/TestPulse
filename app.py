from __future__ import annotations
from io import BytesIO
from pathlib import Path
from datetime import date
import json
import pandas as pd
import streamlit as st

from runner import (
    DEFAULT_APP_URL, SESSION_FILE, STORAGE_STATE_FILE, clear_saved_session,
    load_cases_from_excel, load_fusion_cases_from_excel,
    open_login_session, open_application_session, run_test_case, run_fusion_test_case,
)
from database import database_enabled, init_db, save_test_run, load_history

BASE_DIR = Path(__file__).resolve().parent
REQUIRED_COLUMNS = ["Test Case ID", "Test Case Name", "Module", "Priority", "Expected Result"]
TEMPLATE_COLUMNS = REQUIRED_COLUMNS + ["Description", "Steps", "Automation Key", "Project Name", "Test Date", "Hours", "Allow Save"]
RULES_TEXT = """TESTPULSE — EXCEL TEST SUITE RULES

A workbook can use either of these formats:

FORMAT A — Row-based test cases
Required columns: Test Case ID, Test Case Name, Module, Priority, Expected Result.
Optional columns include Description, Steps, Automation Key, Project Name, Test Date, Hours, Allow Save.
Supported existing scenario keys: login, old_entry, weekend_entry, holiday_entry, over_24_hours, search_project, generic_steps.

FORMAT B — Step-based test cases
Use a worksheet named Fusion_Test_Steps with columns: TEST_CASE_ID, APPLICATION, TEST_CASE_NAME, STEP_NUMBER, ACTION, FIELD_NAME, TEST_DATA, EXPECTED_RESULT, NOTES.
Each test case is a group of rows ordered by STEP_NUMBER. Actions and values are read from Excel.

SAFETY
- Never store passwords, one-time codes, cookies, or session tokens in Excel.
- Saving/creating records is disabled by default. Enable it only for a dedicated test environment and after reviewing the target data.
- Use unique test data to avoid creating duplicate records.
- TestPulse does not invent missing values.
"""


def _missing_case_data(case):
    key = str(case.get("Automation Key", "")).strip().lower()
    required = {"old_entry": ["Project Name"], "weekend_entry": ["Project Name", "Hours"], "holiday_entry": ["Project Name", "Test Date", "Hours"], "over_24_hours": ["Project Name", "Hours"], "search_project": ["Project Name"]}.get(key, [])
    missing = [name for name in required if str(case.get(name, "")).strip() == "" or pd.isna(case.get(name, ""))]
    if key == "holiday_entry" and "Test Date" not in missing:
        try: date.fromisoformat(str(case.get("Test Date")).strip()[:10])
        except Exception: missing.append("Test Date (YYYY-MM-DD)")
    if key in ("weekend_entry", "holiday_entry", "over_24_hours") and "Hours" not in missing:
        try: float(case.get("Hours"))
        except Exception: missing.append("Hours (number)")
    return missing

def _blocked_result(case, missing):
    return {"test_case_id": str(case.get("Test Case ID", "UNKNOWN")), "test_case_name": str(case.get("Test Case Name", "")), "status": "BLOCKED", "duration_seconds": 0, "actual_result": "Not run: required test data is missing from Excel.", "error_message": "Missing Excel values: " + ", ".join(missing), "screenshot_path": "", "video_path": "", "logs": ["BLOCKED before browser launch.", "Missing Excel values: " + ", ".join(missing)], "executed_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}

st.set_page_config(page_title="TestPulse | Automated Testing", page_icon="⚡", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root{--ink:#18263d;--navy:#10213b;--muted:#66758d;--blue:#315eea;--teal:#24bdb2;--line:#dce5f1;--canvas:#f5f7fb;--shadow:0 4px 15px rgba(16,36,70,.045)}
html,body,[data-testid="stAppViewContainer"],.stApp{background:var(--canvas)!important;color:var(--ink)!important;font-family:'DM Sans',sans-serif!important}
[data-testid="stHeader"]{background:transparent!important;height:1.2rem}.block-container{max-width:1480px;padding:.65rem 1.25rem 1.5rem!important}[data-testid="stVerticalBlock"]{gap:.55rem}
h1,h2,h3,h4{font-family:'Manrope',sans-serif!important;color:var(--navy)!important;letter-spacing:-.025em!important}h2{font-size:1.08rem!important;font-weight:800!important;margin:.75rem 0 .15rem!important;padding:0 0 .42rem;border-bottom:1px solid var(--line)}h3{font-size:.94rem!important;font-weight:800!important}
p,li{line-height:1.5}.stMarkdown p{color:#4a5a72}
.hero{background:linear-gradient(115deg,#0d1b31 0%,#16345f 58%,#2854c7 100%);padding:24px 27px;border-radius:16px;color:#fff;margin:0 0 4px;box-shadow:0 9px 24px rgba(16,35,73,.11);border:1px solid #ffffff20}.hero-title{font-family:'Manrope',sans-serif;font-size:clamp(29px,3vw,40px);font-weight:800;letter-spacing:-.055em;line-height:1.08;color:#fff;margin:8px 0}.hero-kicker{display:inline-block;font-size:9px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#d2e2ff;background:#ffffff0d;border:1px solid #ffffff25;padding:5px 9px;border-radius:30px}.hero-copy{color:#e0e9f8;font-size:13px;max-width:760px;line-height:1.6}.metric{background:#fff;border:1px solid var(--line);border-radius:11px;padding:12px 14px;min-height:78px;box-shadow:var(--shadow)}.metric-label{font-size:9px;color:#728098;font-weight:800;letter-spacing:.1em;text-transform:uppercase}.metric-value{font-family:'Manrope',sans-serif;font-size:24px;font-weight:800;color:#122440;margin-top:5px;line-height:1}.metric-foot{font-size:9px;color:#8290a4;margin-top:5px}.panel{background:#fff;border:1px solid var(--line);border-radius:11px;padding:14px 16px;box-shadow:var(--shadow)}
.stButton>button,.stDownloadButton>button{border-radius:8px!important;min-height:37px;font-weight:700!important;border:1px solid #d8e1ef!important;box-shadow:0 2px 5px #13294b08}.stButton>button[kind="primary"]{background:linear-gradient(105deg,#2859e8,#416ff0)!important;border-color:#3268f5!important;color:#fff!important}.stTextInput input,[data-baseweb="select"]>div{border-radius:8px!important;background:#fff!important;color:#203047!important;border-color:#d7e0ee!important;min-height:38px}[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] p{color:#2b3b55!important;font-size:.8rem!important;font-weight:700!important;opacity:1!important}[data-testid="stFileUploader"]{background:#fff;border:1px dashed #8ca8e8;border-radius:10px;padding:8px 10px}[data-testid="stDataFrame"]{border:1px solid #dfe6f0;border-radius:9px;overflow:hidden;background:#fff}[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:4px;border-bottom:1px solid #e1e7f0}[data-testid="stTabs"] [data-baseweb="tab"]{border-radius:7px 7px 0 0;padding:7px 11px;font-weight:700;font-size:.8rem}[data-testid="stExpander"]{background:#fff;border:1px solid #dfe6f0;border-radius:10px;overflow:hidden}[data-testid="stAlert"]{border-radius:9px}hr{border-color:#dfe6f0!important}footer{visibility:hidden}
@media(max-width:700px){.hero{padding:18px}.block-container{padding:.5rem .8rem 1.2rem!important}}
</style>
""", unsafe_allow_html=True)

for key, default in [
    ("execution_results", []), ("loaded_cases", []), ("uploaded_workbook_name", ""),
    ("workbook_type", ""), ("app_url", ""), ("fusion_connected", False),
    ("connected_url", ""), ("last_upload_signature", ""),
    ("preview_validated", False), ("validated_workbook_signature", ""),
]:
    st.session_state.setdefault(key, default)

try:
    init_db()
    db_ok = database_enabled()
    db_error = ""
except Exception as exc:
    db_ok, db_error = False, str(exc)

st.markdown("""
<div class="hero">
  <div class="hero-kicker">✦ QUALITY ENGINEERING · CONTROL CENTER</div>
  <div class="hero-title">TestPulse <span style="color:#7ce7e0">Quality Lab</span></div>
  <div class="hero-copy">Upload a test plan, connect to your application, execute scenarios, and review the evidence. One workspace for multiple applications.</div>
</div>
""", unsafe_allow_html=True)

results = st.session_state.execution_results
passed = sum(r.get("status") == "PASS" for r in results)
failed = sum(r.get("status") == "FAIL" for r in results)
blocked = sum(r.get("status") == "BLOCKED" for r in results)
rate = round(100 * passed / len(results)) if results else 0
metrics = st.columns(4)
for col, label, value, foot in [
    (metrics[0], "TEST CASES", len(st.session_state.loaded_cases), "Loaded from the current workbook"),
    (metrics[1], "EXECUTED", len(results), "Results in this session"),
    (metrics[2], "PASS RATE", f"{rate}%", "Of recorded results"),
    (metrics[3], "NEEDS ATTENTION", failed + blocked, f"{failed} failed · {blocked} blocked"),
]:
    with col:
        st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-foot">{foot}</div></div>', unsafe_allow_html=True)

st.markdown("## 01 · Prepare your test suite")
st.markdown('<div class="panel"><b>Upload one Excel workbook</b><br><span style="color:#65758d;font-size:13px">TestPulse detects the supported workbook structure. No test cases are loaded until you upload a valid workbook.</span></div>', unsafe_allow_html=True)
with st.expander("Download a blank workbook template and rules"):
    template_df = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    xbuf = BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
        template_df.to_excel(writer, index=False, sheet_name="Test Cases")
    c1, c2 = st.columns(2)
    c1.download_button("Download row-based template", xbuf.getvalue(), "testpulse_test_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    c2.download_button("Download workbook rules", RULES_TEXT.encode("utf-8"), "TestPulse_Excel_Rules.txt", "text/plain", use_container_width=True)

upload = st.file_uploader("Upload test cases (.xlsx)", type=["xlsx"], key="unified_workbook_upload")
cases = []
if upload is not None:
    try:
        upload.seek(0)
        fusion_rows = load_fusion_cases_from_excel(upload)
        grouped = {}
        for row in fusion_rows:
            grouped.setdefault(str(row["TEST_CASE_ID"]), []).append(row)
        cases = []
        for test_id, steps in grouped.items():
            steps = sorted(steps, key=lambda r: int(r["STEP_NUMBER"]))
            last_expected = next((str(s.get("EXPECTED_RESULT", "")).strip() for s in reversed(steps) if str(s.get("EXPECTED_RESULT", "")).strip()), "Step-based test")
            cases.append({
                "Test Case ID": test_id,
                "Test Case Name": str(steps[0].get("TEST_CASE_NAME", "")),
                "Module": str(steps[0].get("APPLICATION", "Application")),
                "Priority": "Normal",
                "Expected Result": last_expected,
                "Automation Key": "fusion_steps",
                "Steps": f"{len(steps)} Excel-defined steps",
                "_fusion_steps": steps,
            })
        st.session_state.workbook_type = "step"
        st.session_state.loaded_cases = cases
        st.session_state.uploaded_workbook_name = upload.name
        st.session_state.last_upload_signature = f"{upload.name}:{len(fusion_rows)}"
        st.success(f"Workbook validated · {len(cases)} test case(s), {len(fusion_rows)} Excel-defined steps.")
    except Exception as step_error:
        try:
            upload.seek(0)
            cases = load_cases_from_excel(upload)
            st.session_state.workbook_type = "row"
            st.session_state.loaded_cases = cases
            st.session_state.uploaded_workbook_name = upload.name
            st.session_state.last_upload_signature = f"{upload.name}:{len(cases)}"
            st.success(f"Workbook validated · {len(cases)} test case(s).")
        except Exception as row_error:
            cases = []
            st.session_state.loaded_cases = []
            st.session_state.workbook_type = ""
            st.session_state.uploaded_workbook_name = upload.name
            st.error("This workbook format was not recognized. Check the downloaded rules or use the required column names.")
            with st.expander("Validation details"):
                st.write(f"Step-based format: {step_error}")
                st.write(f"Row-based format: {row_error}")
else:
    st.session_state.loaded_cases = []
    st.session_state.workbook_type = ""
    st.session_state.uploaded_workbook_name = ""
    st.session_state.preview_validated = False
    st.session_state.validated_workbook_signature = ""
    st.session_state.fusion_connected = False
    st.session_state.connected_url = ""

if cases:
    st.markdown("### Test case preview & editor")
    st.caption("Edit the values directly in this table. Click Validate changes before connecting or running tests. Any further edit requires validation again.")

    def _invalidate_preview():
        st.session_state.preview_validated = False

    if st.session_state.workbook_type == "step":
        st.caption("Detected format: step-based workbook · edit individual steps below")
        step_rows = []
        for case in cases:
            for step in case.get("_fusion_steps", []):
                step_rows.append({
                    "TEST_CASE_ID": str(step.get("TEST_CASE_ID", "")),
                    "APPLICATION": str(step.get("APPLICATION", "")),
                    "TEST_CASE_NAME": str(step.get("TEST_CASE_NAME", "")),
                    "STEP_NUMBER": int(step.get("STEP_NUMBER", 0)),
                    "ACTION": str(step.get("ACTION", "")),
                    "FIELD_NAME": str(step.get("FIELD_NAME", "")),
                    "TEST_DATA": str(step.get("TEST_DATA", "")),
                    "EXPECTED_RESULT": str(step.get("EXPECTED_RESULT", "")),
                    "NOTES": str(step.get("NOTES", "")),
                })
        editor_key = "testpulse_step_editor_" + str(st.session_state.last_upload_signature).replace(" ", "_")
        edited = st.data_editor(pd.DataFrame(step_rows), key=editor_key, on_change=_invalidate_preview,
            use_container_width=True, hide_index=True, num_rows="fixed",
            column_config={"STEP_NUMBER": st.column_config.NumberColumn("STEP_NUMBER", min_value=1, step=1)})
        edited_steps = edited.to_dict(orient="records")
        grouped = {}
        for row in edited_steps:
            row["TEST_CASE_ID"] = str(row.get("TEST_CASE_ID", "")).strip()
            if row["TEST_CASE_ID"]:
                grouped.setdefault(row["TEST_CASE_ID"], []).append(row)
        cases = []
        for test_id, steps in grouped.items():
            try:
                steps.sort(key=lambda r: int(r.get("STEP_NUMBER", 0)))
            except Exception:
                pass
            expected = next((str(x.get("EXPECTED_RESULT", "")).strip() for x in reversed(steps) if str(x.get("EXPECTED_RESULT", "")).strip()), "Step-based test")
            cases.append({"Test Case ID": test_id, "Test Case Name": str(steps[0].get("TEST_CASE_NAME", "")),
                "Module": str(steps[0].get("APPLICATION", "Application")), "Priority": "Normal",
                "Expected Result": expected, "Automation Key": "fusion_steps", "Steps": f"{len(steps)} Excel-defined steps", "_fusion_steps": steps})
    else:
        st.caption("Detected format: row-based workbook")
        visible_rows = [{k: v for k, v in c.items() if not k.startswith("_")} for c in cases]
        preview_df = pd.DataFrame(visible_rows)
        editor_key = "testpulse_row_editor_" + str(st.session_state.last_upload_signature).replace(" ", "_")
        edited = st.data_editor(preview_df, key=editor_key, on_change=_invalidate_preview,
            use_container_width=True, hide_index=True, num_rows="fixed")
        cases = edited.to_dict(orient="records")

    st.session_state.loaded_cases = cases
    validate_clicked = st.button("✓ Validate edited test cases", type="primary", use_container_width=True)
    if validate_clicked:
        problems = []
        ids = [str(c.get("Test Case ID", "")).strip() for c in cases]
        if not cases:
            problems.append("No test cases are present in the preview.")
        if any(not cid for cid in ids):
            problems.append("Every test case must have a Test Case ID.")
        if len([cid for cid in ids if cid]) != len(set(cid for cid in ids if cid)):
            problems.append("Test Case IDs must be unique.")
        if st.session_state.workbook_type == "step":
            for c in cases:
                steps = c.get("_fusion_steps", [])
                if not str(c.get("Test Case Name", "")).strip():
                    problems.append(f"{c.get('Test Case ID')}: Test Case Name is required.")
                if not steps:
                    problems.append(f"{c.get('Test Case ID')}: at least one step is required.")
                for step in steps:
                    if not str(step.get("ACTION", "")).strip():
                        problems.append(f"{c.get('Test Case ID')}: ACTION cannot be blank.")
                    try:
                        int(step.get("STEP_NUMBER", 0))
                    except Exception:
                        problems.append(f"{c.get('Test Case ID')}: STEP_NUMBER must be a whole number.")
        else:
            required = ["Test Case ID", "Test Case Name", "Module", "Priority", "Expected Result"]
            for c in cases:
                for col in required:
                    if not str(c.get(col, "")).strip():
                        problems.append(f"{c.get('Test Case ID') or '(missing ID)'}: {col} is required.")
                problems.extend([f"{c.get('Test Case ID')}: missing or invalid {x}." for x in _missing_case_data(c)])
        if problems:
            st.session_state.preview_validated = False
            st.error("Validation failed. Fix the items below and validate again.")
            for problem in problems:
                st.markdown(f"- {problem}")
        else:
            st.session_state.preview_validated = True
            st.session_state.validated_workbook_signature = st.session_state.last_upload_signature
            st.success(f"Validation passed for {len(cases)} test case(s). You can now connect and execute.")

    if st.session_state.preview_validated and st.session_state.validated_workbook_signature == st.session_state.last_upload_signature:
        st.success("Test cases validated. You may proceed to application connection.")
    else:
        st.warning("Validation required: connect and execute controls remain locked until the preview passes validation.")
else:
    st.info("Upload a valid Excel workbook to preview test cases and enable execution.")

st.markdown("## 02 · Connect to your application")
with st.form("application_connection_form", clear_on_submit=False):
    app_url = st.text_input("Application URL", value=st.session_state.app_url, placeholder="https://your-application.example.com")
    col_user, col_pass = st.columns(2)
    username = col_user.text_input("Username", key="runtime_username", placeholder="Enter username")
    password = col_pass.text_input("Password", key="runtime_password", type="password", placeholder="Enter password")
    connect_clicked = st.form_submit_button("Connect / sign in", type="primary", use_container_width=True, disabled=(not cases or not st.session_state.preview_validated or st.session_state.validated_workbook_signature != st.session_state.last_upload_signature))

if connect_clicked:
    if not cases:
        st.error("Upload a valid test workbook before connecting.")
    elif not st.session_state.preview_validated or st.session_state.validated_workbook_signature != st.session_state.last_upload_signature:
        st.error("Validate the test case preview successfully before connecting.")
    elif not app_url.strip().startswith(("https://", "http://")):
        st.error("Enter a valid application URL beginning with http:// or https://.")
    else:
        st.session_state.app_url = app_url.strip()
        with st.spinner("Opening the application in Chrome..."):
            try:
                if st.session_state.workbook_type == "row":
                    answer = open_login_session(app_url=app_url.strip(), username=username.strip(), password=password, show_browser=True, timeout_seconds=60)
                    st.session_state.fusion_connected = False
                else:
                    answer = open_application_session(app_url=app_url.strip(), username=username.strip(), password=password)
                    st.session_state.fusion_connected = True
                    st.session_state.connected_url = app_url.strip()
                st.success(answer.get("message", "Application opened."))
                st.info("If your organization uses SSO, MFA, or a verification prompt, finish it in the visible Chrome window before running tests.")
            except Exception as exc:
                st.error(f"Could not connect to the application: {exc}")

if st.session_state.workbook_type == "row":
    if SESSION_FILE.exists() and STORAGE_STATE_FILE.exists():
        try:
            info = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
            st.success(f"Saved browser session detected · {info.get('authenticated_url', 'URL available')}")
        except Exception:
            st.success("Saved browser session files detected.")
    else:
        st.caption("Connect to the application before running browser tests.")
elif st.session_state.workbook_type == "step":
    if st.session_state.fusion_connected and st.session_state.connected_url == st.session_state.app_url:
        st.success("Application browser connection started. Complete any pending sign-in or MFA in Chrome before running tests.")
    else:
        st.caption("Connect to the application before running browser tests.")

st.markdown("## 03 · Execute test plan")
if cases:
    all_ids = [str(c.get("Test Case ID", "")) for c in cases]
    selected_ids = st.multiselect("Choose test cases", all_ids, default=all_ids[:1], format_func=lambda cid: next(f"{c.get('Test Case ID')} — {c.get('Test Case Name')}" for c in cases if str(c.get("Test Case ID", "")) == cid))
    allow_changes = st.checkbox("Allow tests to save or create records", value=False, help="Disabled by default. Enable only for a dedicated test environment after reviewing the data and avoiding duplicates.")
    show_browser = st.checkbox("Keep Chrome visible during execution", value=True)
    r1, r2 = st.columns(2)
    validation_ready = st.session_state.preview_validated and st.session_state.validated_workbook_signature == st.session_state.last_upload_signature
    run_selected = r1.button("▶ Run selected tests", type="primary", use_container_width=True, disabled=not validation_ready)
    run_all = r2.button("▶ Run all tests", use_container_width=True, disabled=not validation_ready)
    if run_selected or run_all:
        is_connected = (SESSION_FILE.exists() and STORAGE_STATE_FILE.exists()) if st.session_state.workbook_type == "row" else (st.session_state.fusion_connected and st.session_state.connected_url == st.session_state.app_url)
        if not st.session_state.preview_validated or st.session_state.validated_workbook_signature != st.session_state.last_upload_signature:
            st.error("Validate the edited test case preview before execution.")
        elif not is_connected:
            st.error("Connect/sign in to the application first.")
        else:
            queue = cases if run_all else [c for c in cases if str(c.get("Test Case ID", "")) in selected_ids]
            if not queue:
                st.error("Select at least one test case.")
            else:
                progress = st.progress(0)
                live = st.empty()
                completed = []
                for i, case in enumerate(queue, 1):
                    cid = str(case.get("Test Case ID", "UNKNOWN"))
                    live.info(f"Running {i}/{len(queue)} · {cid} — {case.get('Test Case Name', '')}")
                    try:
                        if st.session_state.workbook_type == "step":
                            cfg = {"app_url": st.session_state.app_url, "username": username, "password": password, "allow_save": allow_changes, "headless": not show_browser}
                            result = run_fusion_test_case(case.get("_fusion_steps", []), cfg)
                        else:
                            key = str(case.get("Automation Key", "")).strip().lower()
                            allow_save_for_case = allow_changes and str(case.get("Allow Save", "")).strip().lower() in ("true", "yes", "1")
                            cfg = {
                                "app_url": st.session_state.app_url.strip(),
                                "project_name": str(case.get("Project Name", "")).strip(),
                                "holiday_date": str(case.get("Test Date", "")).strip()[:10],
                                "normal_hours": float(case.get("Hours")) if key in ("weekend_entry", "holiday_entry") and str(case.get("Hours", "")).strip() else 0,
                                "excessive_hours": float(case.get("Hours")) if key == "over_24_hours" and str(case.get("Hours", "")).strip() else 0,
                                "allow_save": allow_save_for_case,
                                "headless": not show_browser,
                            }
                            missing = _missing_case_data(case)
                            if missing:
                                result = _blocked_result(case, missing)
                            else:
                                result = run_test_case(case, cfg)
                    except Exception as exc:
                        result = {"test_case_id": cid, "test_case_name": str(case.get("Test Case Name", "")), "status": "FAIL", "duration_seconds": 0, "actual_result": "Automation could not complete the scenario.", "error_message": str(exc), "screenshot_path": "", "video_path": "", "logs": [f"FAIL: {exc}"], "executed_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}
                    completed.append(result)
                    progress.progress(i / len(queue))
                existing = {r.get("test_case_id"): r for r in st.session_state.execution_results}
                existing.update({r.get("test_case_id"): r for r in completed})
                st.session_state.execution_results = list(existing.values())
                try:
                    if db_ok:
                        save_test_run(completed, st.session_state.uploaded_workbook_name or "uploaded_test_suite.xlsx", st.session_state.app_url.strip())
                except Exception as exc:
                    st.warning(f"Results are available, but PostgreSQL save failed: {exc}")
                live.success(f"Finished {len(completed)} test(s). Results are shown below.")
                st.rerun()
else:
    st.caption("Upload a valid workbook to enable execution.")

st.markdown("## 04 · Results & evidence")
results = st.session_state.execution_results
if results:
    summary = pd.DataFrame([{"Test Case ID": r.get("test_case_id", ""), "Test Case Name": r.get("test_case_name", ""), "Status": r.get("status", ""), "Duration (s)": r.get("duration_seconds", 0), "Actual Result": r.get("actual_result", ""), "Executed At": r.get("executed_at", ""), "Error": r.get("error_message", "")} for r in results[::-1]])
    st.dataframe(summary, use_container_width=True, hide_index=True)
    options = [f"{r.get('test_case_id')} — {r.get('test_case_name')}" for r in results[::-1]]
    chosen = st.selectbox("Inspect evidence for", options)
    result = next(item for item in results[::-1] if f"{item.get('test_case_id')} — {item.get('test_case_name')}" == chosen)
    st.markdown(f"### {result.get('test_case_id')} · {result.get('test_case_name')}")
    st.write(f"**Status:** {result.get('status', '')}")
    st.write(result.get("actual_result", ""))
    if result.get("error_message"):
        st.error(result["error_message"])
    tab1, tab2, tab3 = st.tabs(["📸 Screenshot", "🎞 Video", "☷ Execution steps"])
    with tab1:
        p = Path(result.get("screenshot_path", "")) if result.get("screenshot_path") else None
        if p and p.is_file():
            st.image(str(p), caption=f"Screenshot — {result.get('test_case_id')}", use_container_width=True)
            st.caption(f"Saved at: {p}")
        else:
            st.info("No screenshot was recorded for this result.")
    with tab2:
        p = Path(result.get("video_path", "")) if result.get("video_path") else None
        if p and p.is_file() and p.stat().st_size:
            st.video(str(p))
            st.caption(f"Saved at: {p}")
        else:
            st.info("No video was recorded for this result.")
    with tab3:
        logs = result.get("logs", [])
        if logs:
            for i, entry in enumerate(logs, 1):
                st.markdown(f"**Step {i}**")
                st.code(str(entry), language="text")
        else:
            st.info("No execution logs were recorded.")
    st.download_button("⬇ Export results as CSV", summary.to_csv(index=False).encode("utf-8"), "testpulse_results.csv", "text/csv")
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Results")
    st.download_button("⬇ Export results as Excel", out.getvalue(), "testpulse_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("No test results yet. Run tests to view screenshots, videos, and execution logs here.")

with st.expander("PostgreSQL history and setup status"):
    if db_ok:
        st.success("PostgreSQL is configured. Completed test runs are saved to the history database.")
        try:
            history = load_history(limit=100)
            if history:
                st.dataframe(pd.DataFrame(history), use_container_width=True, hide_index=True)
            else:
                st.caption("No saved runs yet.")
        except Exception as exc:
            st.warning(f"Could not read history: {exc}")
    else:
        st.warning("PostgreSQL is not configured. The dashboard can run locally, but history will not persist across app restarts.")
        if db_error:
            st.code(db_error)
    st.caption("Browser session files and evidence stay local. Credentials are entered at runtime and are not written to Excel.")

st.markdown("---")
st.caption("TESTPULSE  /  Upload. Test. See the proof.")

