from __future__ import annotations
from io import BytesIO
from pathlib import Path
from datetime import date
import json
import pandas as pd
import streamlit as st

from runner import (
    DEFAULT_APP_URL, SESSION_FILE, STORAGE_STATE_FILE, clear_saved_session,
    load_cases_from_excel, open_login_session, run_test_case,
)
from database import database_enabled, init_db, save_test_run, load_history

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK = BASE_DIR / "timesheet_test_cases.xlsx"
REQUIRED_COLUMNS = ["Test Case ID", "Test Case Name", "Module", "Priority", "Expected Result"]
TEMPLATE_COLUMNS = REQUIRED_COLUMNS + ["Description", "Steps", "Automation Key", "Project Name", "Test Date", "Hours", "Allow Save"]
RULES_TEXT = """TESTPULSE — EXCEL TEST SUITE RULES

1. Keep the required columns exactly named: Test Case ID, Test Case Name, Module, Priority, Expected Result.
2. Every test row must have a unique Test Case ID and non-empty required fields.
3. Optional columns: Description, Steps, Automation Key, Project Name, Test Date, Hours, Allow Save.
4. Automation Key values supported by this runner: login, old_entry, weekend_entry, holiday_entry, over_24_hours, search_project, generic_steps.
5. Project Name is required for old_entry, weekend_entry, holiday_entry, and search_project.
6. Hours is required for weekend_entry and holiday_entry. For over_24_hours, enter the invalid hours value in Hours.
7. Test Date is required for holiday_entry and must use YYYY-MM-DD (for example, 2026-12-25). Use the real date you intend to test.
8. Allow Save is optional and defaults to false. Only set true when you explicitly intend the runner to submit a change in a dedicated test environment.
9. Do not enter passwords, one-time codes, session cookies, or other secrets in Excel.
10. This runner does not invent missing test data. Rows with missing scenario-specific values are blocked before browser execution.
11. Generic steps are separated with semicolons. Supported examples: Open application; Wait 2 seconds; Click Add Project; Verify text Employee Timesheet.
12. Save the workbook as .xlsx and upload it in TestPulse. Review the preview and missing-data warnings before running.
"""

st.set_page_config(page_title="TestPulse | Automated Testing", page_icon="⚡", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
:root{--navy:#07162c;--blue:#356dff;--cyan:#4de2d0;--ink:#152742;--muted:#6e7f98;--line:#e4eaf4;--paper:#f4f7fc;--shadow:0 12px 32px rgba(15,36,74,.07)}
html,body,[data-testid="stAppViewContainer"],.stApp{background:radial-gradient(ellipse at 8% 0%,#e2eaff 0%,transparent 24%),radial-gradient(ellipse at 100% 26%,#e8f8ff 0%,transparent 20%),var(--paper)!important;color:var(--ink)!important;font-family:'DM Sans',sans-serif}
[data-testid="stHeader"]{background:transparent!important}
.block-container{max-width:1540px;padding:1.2rem 2.3rem 3.2rem!important}
h1,h2,h3,h4{font-family:'Manrope',sans-serif!important;color:var(--navy)!important;letter-spacing:-.035em!important}
h2{font-size:1.48rem!important;font-weight:800!important;margin:2.2rem 0 .6rem!important}
h3{font-size:1.05rem!important;font-weight:800!important}
p,li,label{line-height:1.65}
.hero{position:relative;isolation:isolate;overflow:hidden;background:linear-gradient(118deg,#07162c 0%,#102b53 48%,#2459db 100%);padding:37px 40px 33px;border-radius:28px;color:#fff;margin:4px 0 24px;box-shadow:0 24px 65px rgba(18,45,100,.22);border:1px solid #ffffff22;min-height:260px}
.hero:before{content:"";position:absolute;z-index:-1;width:440px;height:440px;right:5%;top:-240px;border-radius:50%;background:radial-gradient(circle,#56a5ff45 0%,#56a5ff00 68%);pointer-events:none}
.hero:after{content:"";position:absolute;z-index:-1;width:330px;height:330px;right:-70px;bottom:-220px;border:1px solid #ffffff2b;border-radius:50%;box-shadow:0 0 0 25px #ffffff08,0 0 0 55px #ffffff06,0 0 0 85px #ffffff04;pointer-events:none}
.hero-grid{display:grid;grid-template-columns:minmax(0,1fr) 250px;gap:25px;align-items:center}
.hero-kicker{display:inline-flex;align-items:center;gap:8px;font-size:10px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#b9d5ff;background:#ffffff0c;border:1px solid #ffffff27;padding:8px 12px;border-radius:30px}
.hero-title{font-family:'Manrope',sans-serif;font-size:clamp(32px,3.2vw,48px);font-weight:800;letter-spacing:-.06em;line-height:1.05;color:#fff;margin:18px 0 12px}
.hero-copy{color:#dce8ff;font-size:14px;max-width:760px;line-height:1.8}
.hero-meta{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:22px}
.chip{display:inline-block;border:1px solid #ffffff30;background:#ffffff0c;border-radius:30px;padding:7px 11px;font-size:10px;font-weight:700;letter-spacing:.04em;color:#f4f8ff}
.hero-console{border:1px solid #ffffff27;background:linear-gradient(145deg,#ffffff19,#ffffff08);backdrop-filter:blur(14px);border-radius:20px;padding:19px;box-shadow:inset 0 1px #ffffff1b,0 15px 35px #06132935}
.console-top{display:flex;align-items:center;justify-content:space-between;color:#d9e7ff;font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase}
.live-dot{display:inline-block;width:7px;height:7px;background:#4de2b2;border-radius:50%;box-shadow:0 0 0 5px #4de2b21b;margin-right:7px}
.console-number{font:800 36px 'Manrope',sans-serif;color:#fff;letter-spacing:-.05em;margin:15px 0 0}
.console-caption{font-size:11px;color:#bfd1ef;margin-top:2px}
.console-line{height:6px;background:#ffffff18;border-radius:20px;overflow:hidden;margin:15px 0 12px}
.console-line span{display:block;width:72%;height:100%;background:linear-gradient(90deg,#4de2d0,#70a6ff);border-radius:20px}
.console-foot{display:flex;justify-content:space-between;color:#d9e7ff;font-size:10px}
.metric{position:relative;overflow:hidden;background:linear-gradient(145deg,#fff 0%,#fafdff 100%);border:1px solid #e1e9f5;border-radius:19px;padding:19px 20px 17px;min-height:119px;box-shadow:var(--shadow);transition:transform .2s ease,box-shadow .2s ease}
.metric:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(180deg,#50d9d0,#3e6fff)}
.metric:after{content:"";position:absolute;width:95px;height:95px;right:-42px;top:-42px;border-radius:50%;background:#eaf1ff}
.metric-label{position:relative;z-index:1;font-size:10px;color:#72829a;font-weight:800;letter-spacing:.12em;text-transform:uppercase}
.metric-value{position:relative;z-index:1;font-family:'Manrope',sans-serif;font-size:32px;font-weight:800;letter-spacing:-.055em;color:#0b1f3d;margin-top:10px;line-height:1.05}
.metric-foot{position:relative;z-index:1;font-size:11px;color:#8492a6;margin-top:8px}
.panel{background:#fff;border:1px solid var(--line);border-radius:18px;padding:22px 24px;margin:10px 0 16px;box-shadow:var(--shadow)}
.stMarkdown p{color:#465975}
[data-testid="stCaptionContainer"]{color:#7c8ba2!important}
.stButton>button,.stDownloadButton>button{border-radius:12px!important;min-height:45px;font-weight:750!important;border:1px solid #dce5f2!important;box-shadow:0 3px 8px #13294b0a;transition:all .16s ease!important}
.stButton>button:hover,.stDownloadButton>button:hover{border-color:#9bb8ff!important;transform:translateY(-2px);box-shadow:0 9px 22px #2f66f51c!important}
.stButton>button[kind="primary"]{background:linear-gradient(105deg,#2859e8,#4c7dff)!important;border-color:#3268f5!important;color:#fff!important;box-shadow:0 9px 22px #3268f535!important}
.stButton>button[kind="primary"] p{color:#fff!important}
.stTextInput input,.stTextArea textarea,.stDateInput input,[data-baseweb="select"]>div{border-radius:11px!important;background:#fff!important;color:#203047!important;border-color:#dbe4f1!important}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] p{color:#273a58!important;font-weight:700!important;opacity:1!important}
[data-testid="stFileUploader"]{background:linear-gradient(145deg,#fff,#f8fbff);border:1px dashed #91aff8;border-radius:16px;padding:14px}
[data-testid="stFileUploader"] section{background:transparent!important}
[data-testid="stDataFrame"]{border:1px solid #e1e8f3;border-radius:15px;overflow:hidden;background:#fff;box-shadow:0 6px 20px #15294b08}
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:7px;border-bottom:1px solid #e4eaf4}
[data-testid="stTabs"] [data-baseweb="tab"]{border-radius:10px 10px 0 0;padding:10px 16px;font-weight:700}
[data-testid="stTabs"] [aria-selected="true"]{color:#2d60ec!important;border-bottom-color:#3268f5!important}
[data-testid="stExpander"]{background:#fff;border:1px solid #e4eaf4;border-radius:14px;overflow:hidden}
[data-testid="stAlert"]{border-radius:13px}
div[role="radiogroup"] label,div[role="radiogroup"] label p{color:#203047!important;opacity:1!important}
[data-testid="stProgressBar"]>div>div{background:linear-gradient(90deg,#4de2d0,#356dff)!important}
hr{border-color:#e5ebf4!important;margin:2rem 0 1rem!important}
footer{visibility:hidden}
@media(max-width:900px){.hero-grid{grid-template-columns:1fr}.hero-console{max-width:360px}.block-container{padding:1rem 1rem 2rem!important}}
@media(max-width:600px){.hero{padding:25px 22px;border-radius:20px}.hero-title{font-size:31px}.metric{padding:15px;min-height:100px}.metric-value{font-size:26px}}
</style>
""", unsafe_allow_html=True)

for key, default in [("execution_results", []), ("loaded_cases", []), ("uploaded_workbook_name", ""), ("app_url", DEFAULT_APP_URL)]:
    st.session_state.setdefault(key, default)

try:
    init_db()
    db_ok = database_enabled()
    db_error = ""
except Exception as exc:
    db_ok, db_error = False, str(exc)

# Header: HTML is deliberately a standalone markdown string, not mixed with code blocks.
st.markdown("""
<div class="hero">
  <div class="hero-grid">
    <div>
      <div class="hero-kicker">✦ QUALITY ENGINEERING · CONTROL CENTER</div>
      <div class="hero-title">TestPulse <span style="color:#7ce7e0">Quality Lab</span></div>
      <div class="hero-copy">From test plan to proof. Orchestrate Oracle APEX checks, capture every important interaction, and turn execution results into evidence your team can trust.</div>
      <div class="hero-meta"><span class="chip">ORACLE APEX</span><span class="chip">PLAYWRIGHT AUTOMATION</span><span class="chip">POSTGRESQL HISTORY</span><span class="chip">EVIDENCE-LED QA</span></div>
    </div>
    <div class="hero-console">
      <div class="console-top"><span>Test environment</span><span><i class="live-dot"></i>Ready to run</span></div>
      <div class="console-number">QA<span style="color:#7ce7e0">/</span>01</div>
      <div class="console-caption">Controlled browser execution · Evidence capture</div>
      <div class="console-line"><span></span></div>
      <div class="console-foot"><span>Automation engine</span><span>Playwright</span></div>
      <div class="console-foot" style="margin-top:8px"><span>Result persistence</span><span>PostgreSQL</span></div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

results = st.session_state.execution_results
passed = sum(r.get("status") == "PASS" for r in results)
failed = sum(r.get("status") == "FAIL" for r in results)
blocked = sum(r.get("status") == "BLOCKED" for r in results)
rate = round(100 * passed / len(results)) if results else 0
cols = st.columns(4)
for col, label, value, foot in [
    (cols[0], "TEST CASES", len(st.session_state.loaded_cases), "Loaded from Excel"),
    (cols[1], "EXECUTED", len(results), "Results in this session"),
    (cols[2], "PASS RATE", f"{rate}%", "Of recorded results"),
    (cols[3], "NEEDS ATTENTION", failed + blocked, f"{failed} failed · {blocked} blocked"),
]:
    with col:
        st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-foot">{foot}</div></div>', unsafe_allow_html=True)

st.markdown("## 01  Prepare your test suite")
left, right = st.columns([1, 1], gap="large")
with left:
    template_df = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    # Seed template with the existing project workbook's test cases when available, without inventing scenario values.
    if WORKBOOK.exists():
        try:
            existing = pd.read_excel(WORKBOOK, engine="openpyxl").fillna("")
            for col in TEMPLATE_COLUMNS:
                if col not in existing.columns:
                    existing[col] = ""
            template_df = existing[TEMPLATE_COLUMNS]
        except Exception:
            pass
    xbuf = BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
        template_df.to_excel(writer, index=False, sheet_name="Test Cases")
    st.download_button("⬇ Download Excel template", xbuf.getvalue(), "testpulse_test_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.download_button("⬇ Download Excel rules", RULES_TEXT.encode("utf-8"), "TestPulse_Excel_Rules.txt", "text/plain", use_container_width=True)
with right:
    st.markdown('<div class="panel"><h3>Designed for reliable runs</h3><p><b>01 · Required fields</b><br>Test Case ID, Test Case Name, Module, Priority and Expected Result.</p><p><b>02 · Scenario inputs</b><br>Project Name, Test Date and Hours are supplied in the workbook. Missing inputs are flagged before execution.</p><p><b>03 · Safe by default</b><br>Save actions stay disabled unless you explicitly enable them. Never include passwords or session tokens in Excel.</p></div>', unsafe_allow_html=True)

upload = st.file_uploader("Upload your test suite (.xlsx)", type=["xlsx"], key="testpulse_workbook")
if upload is not None:
    try:
        loaded = load_cases_from_excel(upload)
        st.session_state.loaded_cases = loaded
        st.session_state.uploaded_workbook_name = upload.name
        st.success(f"Validated {len(loaded)} test cases from {upload.name}.")
    except Exception as exc:
        st.session_state.loaded_cases = []
        st.session_state.uploaded_workbook_name = upload.name
        st.error(f"Workbook validation failed: {exc}")
elif not st.session_state.loaded_cases and WORKBOOK.exists():
    try:
        st.session_state.loaded_cases = load_cases_from_excel(WORKBOOK)
        st.session_state.uploaded_workbook_name = WORKBOOK.name
    except Exception as exc:
        st.warning(f"Included workbook could not be loaded: {exc}")

cases = st.session_state.loaded_cases
if cases:
    display_cols = [c for c in TEMPLATE_COLUMNS if any(c in row for row in cases)]
    st.markdown("### Test case preview")
    st.caption(f"Active workbook: {st.session_state.uploaded_workbook_name or 'Excel workbook'} · {len(cases)} rows")
    st.dataframe(pd.DataFrame([{c: row.get(c, "") for c in display_cols} for row in cases]), use_container_width=True, hide_index=True)
else:
    st.info("Upload an Excel workbook to preview and run its test cases.")


def missing_case_data(case):
    key = str(case.get("Automation Key", "")).strip().lower()
    required = {
        "old_entry": ["Project Name"],
        "weekend_entry": ["Project Name", "Hours"],
        "holiday_entry": ["Project Name", "Test Date", "Hours"],
        "over_24_hours": ["Project Name", "Hours"],
        "search_project": ["Project Name"],
    }.get(key, [])
    missing = [name for name in required if str(case.get(name, "")).strip() == "" or pd.isna(case.get(name, ""))]
    if key == "holiday_entry" and "Test Date" not in missing:
        try:
            date.fromisoformat(str(case.get("Test Date")).strip()[:10])
        except Exception:
            missing.append("Test Date (YYYY-MM-DD)")
    if key in ("weekend_entry", "holiday_entry", "over_24_hours") and "Hours" not in missing:
        try:
            float(case.get("Hours"))
        except Exception:
            missing.append("Hours (number)")
    return missing

if cases:
    missing_rows = []
    for case in cases:
        missing = missing_case_data(case)
        if missing:
            missing_rows.append({"Test Case ID": case.get("Test Case ID", ""), "Test Case Name": case.get("Test Case Name", ""), "Missing Excel values": ", ".join(missing)})
    if missing_rows:
        with st.expander(f"Excel data checks · {len(missing_rows)} test case(s) need values", expanded=True):
            st.warning("These tests will be marked BLOCKED until the required values are supplied in Excel. TestPulse will not invent project names, dates, or hours.")
            st.dataframe(pd.DataFrame(missing_rows), use_container_width=True, hide_index=True)

st.markdown("## 02  Connect to Oracle APEX")
st.caption("Sign in manually in the Chrome window. TestPulse does not request or store your APEX password.")
st.session_state.app_url = st.text_input("Oracle APEX application URL", value=st.session_state.app_url)
sc1, sc2, sc3 = st.columns([1, 1, 2])
with sc1:
    login_clicked = st.button("🔐 Open Chrome & sign in", type="primary", use_container_width=True)
with sc2:
    if st.button("Clear saved session", use_container_width=True):
        clear_saved_session()
        st.success("Saved local browser session cleared.")
        st.rerun()
with sc3:
    if SESSION_FILE.exists() and STORAGE_STATE_FILE.exists():
        try:
            session_info = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
            st.success(f"Saved browser session found · {session_info.get('saved_at', 'time unavailable')}")
        except Exception:
            st.success("Saved browser session files found.")
    else:
        st.warning("Sign in before running browser tests.")
if login_clicked:
    if not st.session_state.app_url.strip().startswith(("https://", "http://")):
        st.error("Enter a valid http/https application URL.")
    else:
        st.info("Chrome will open. Complete sign-in manually, then open Employee Timesheet. Keep this dashboard open.")
        with st.spinner("Waiting for Employee Timesheet..."):
            try:
                answer = open_login_session(st.session_state.app_url.strip(), show_browser=True, timeout_seconds=300)
                st.success(answer["message"])
                st.rerun()
            except Exception as exc:
                st.error(f"Sign-in was not confirmed: {exc}")

st.markdown("## 03  Execute test plan")
if cases:
    all_ids = [str(c.get("Test Case ID", "")) for c in cases]
    selected_ids = st.multiselect("Choose test cases (used by Run selected)", all_ids, default=all_ids[:1], format_func=lambda cid: next(f"{c.get('Test Case ID')} — {c.get('Test Case Name')}" for c in cases if str(c.get("Test Case ID", "")) == cid))
    safe_save = st.checkbox("Allow tests to submit changes in APEX", value=False, help="Keep disabled for safer testing. Enable only when you intentionally want the selected scenarios to submit changes in a dedicated test environment.")
    show_browser = st.checkbox("Show Chrome while tests run", value=True)
    run1, run2 = st.columns(2)
    run_selected = run1.button("▶ Run selected tests", type="primary", use_container_width=True)
    run_all = run2.button("▶ Run all tests", use_container_width=True)
    if run_selected or run_all:
        if not (SESSION_FILE.exists() and STORAGE_STATE_FILE.exists()):
            st.error("Sign in first using Open Chrome & sign in.")
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
                    missing = missing_case_data(case)
                    if missing:
                        result = {"test_case_id": cid, "test_case_name": str(case.get("Test Case Name", "")), "status": "BLOCKED", "duration_seconds": 0, "actual_result": "Not run: required test data is missing from Excel.", "error_message": "Missing Excel values: " + ", ".join(missing), "screenshot_path": "", "video_path": "", "logs": ["BLOCKED before browser launch.", "Missing Excel values: " + ", ".join(missing)], "executed_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}
                    else:
                        key = str(case.get("Automation Key", "")).strip().lower()
                        cfg = {
                            "app_url": st.session_state.app_url.strip(),
                            "project_name": str(case.get("Project Name", "")).strip(),
                            "holiday_date": str(case.get("Test Date", "")).strip()[:10],
                            "normal_hours": float(case.get("Hours")) if key in ("weekend_entry", "holiday_entry") else 0,
                            "excessive_hours": float(case.get("Hours")) if key == "over_24_hours" else 0,
                            "allow_save": bool(safe_save and str(case.get("Allow Save", "")).strip().lower() in ("true", "yes", "1")),
                            "headless": not show_browser,
                        }
                        try:
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
    st.caption("Upload a valid Excel suite to enable test execution.")

st.markdown("## 04  Results & evidence")
results = st.session_state.execution_results
if results:
    summary = pd.DataFrame([{"Test Case ID": r.get("test_case_id", ""), "Test Case Name": r.get("test_case_name", ""), "Status": r.get("status", ""), "Duration (s)": r.get("duration_seconds", 0), "Actual Result": r.get("actual_result", ""), "Executed At": r.get("executed_at", ""), "Error": r.get("error_message", "")} for r in results[::-1]])
    st.dataframe(summary, use_container_width=True, hide_index=True)
    options = [f"{r.get('test_case_id')} — {r.get('test_case_name')}" for r in results[::-1]]
    chosen = st.selectbox("Inspect evidence for", options)
    r = next(item for item in results[::-1] if f"{item.get('test_case_id')} — {item.get('test_case_name')}" == chosen)
    st.markdown(f"### {r.get('test_case_id')} · {r.get('test_case_name')}")
    st.write(f"**Status:** {r.get('status', '')}")
    st.write(r.get("actual_result", ""))
    if r.get("error_message"):
        st.error(r["error_message"])
    tab1, tab2, tab3 = st.tabs(["📸 Screenshot", "🎞 Video", "☷ Execution steps"])
    with tab1:
        p = Path(r.get("screenshot_path", "")) if r.get("screenshot_path") else None
        if p and p.is_file():
            st.image(str(p), caption=f"Screenshot — {r.get('test_case_id')}", use_container_width=True)
            st.caption(f"Saved at: {p}")
        else:
            st.info("No screenshot was recorded for this result (for example, a test blocked before browser launch).")
    with tab2:
        p = Path(r.get("video_path", "")) if r.get("video_path") else None
        if p and p.is_file() and p.stat().st_size:
            st.video(str(p))
            st.caption(f"Saved at: {p}")
        else:
            st.info("No video was recorded for this result.")
    with tab3:
        logs = r.get("logs", [])
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
    st.info("No test results yet. Run selected tests or the full suite above.")

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
    st.caption("Oracle APEX remains the target application. Playwright handles browser automation and local evidence. PostgreSQL stores test history. Browser session files and evidence stay local.")

st.markdown("---")
st.caption("TESTPULSE  /  Upload. Test. See the proof.  ·  Browser state is sensitive and stays local.")
