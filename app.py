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
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root{--ink:#18263d;--navy:#10213b;--muted:#66758d;--blue:#315eea;--teal:#24bdb2;--line:#dce5f1;--canvas:#f5f7fb;--shadow:0 4px 15px rgba(16,36,70,.045)}
html,body,[data-testid="stAppViewContainer"],.stApp{background:var(--canvas)!important;color:var(--ink)!important;font-family:'DM Sans',sans-serif!important}
[data-testid="stHeader"]{background:transparent!important;height:1.2rem}.block-container{max-width:1480px;padding:.5rem 1.25rem 1.3rem!important}[data-testid="stVerticalBlock"]{gap:.48rem}
h1,h2,h3,h4{font-family:'Manrope',sans-serif!important;color:var(--navy)!important;letter-spacing:-.025em!important}
h2{font-size:1.04rem!important;font-weight:800!important;margin:.75rem 0 .1rem!important;padding:0 0 .42rem;border-bottom:1px solid var(--line)}h3{font-size:.94rem!important;font-weight:800!important;margin:.1rem 0 .35rem!important}
p,li{line-height:1.48}.stMarkdown p{color:#4a5a72}
.hero{position:relative;isolation:isolate;overflow:hidden;background:linear-gradient(115deg,#0d1b31 0%,#16345f 58%,#2854c7 100%);padding:21px 25px;border-radius:16px;color:#fff;margin:0 0 3px;box-shadow:0 9px 24px rgba(16,35,73,.11);border:1px solid #ffffff20}
.hero:before{content:"";position:absolute;z-index:-1;width:330px;height:330px;right:8%;top:-245px;border-radius:50%;background:radial-gradient(circle,#60a5fa38 0%,transparent 68%);pointer-events:none}.hero:after{content:"";position:absolute;z-index:-1;width:210px;height:210px;right:-70px;bottom:-165px;border:1px solid #ffffff24;border-radius:50%;box-shadow:0 0 0 20px #ffffff06,0 0 0 42px #ffffff05;pointer-events:none}
.hero-grid{display:grid;grid-template-columns:minmax(0,1fr) 205px;gap:20px;align-items:center}.hero-kicker{display:inline-flex;align-items:center;gap:7px;font-size:9px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#d2e2ff;background:#ffffff0d;border:1px solid #ffffff25;padding:5px 9px;border-radius:30px}.hero-title{font-family:'Manrope',sans-serif;font-size:clamp(29px,3vw,39px);font-weight:800;letter-spacing:-.055em;line-height:1.08;color:#fff;margin:10px 0 7px}.hero-copy{color:#e0e9f8;font-size:12.5px;max-width:740px;line-height:1.6}.hero-meta{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:12px}.chip{display:inline-block;border:1px solid #ffffff2c;background:#ffffff0c;border-radius:30px;padding:5px 8px;font-size:8px;font-weight:800;letter-spacing:.035em;color:#f4f8ff}
.hero-console{border:1px solid #ffffff26;background:linear-gradient(145deg,#ffffff19,#ffffff08);border-radius:13px;padding:13px 14px;box-shadow:inset 0 1px #ffffff16}.console-top{display:flex;align-items:center;justify-content:space-between;color:#d9e7ff;font-size:8px;font-weight:800;letter-spacing:.07em;text-transform:uppercase}.live-dot{display:inline-block;width:6px;height:6px;background:#52e0b3;border-radius:50%;box-shadow:0 0 0 4px #4de2b21a;margin-right:5px}.console-number{font:800 27px 'Manrope',sans-serif;color:#fff;letter-spacing:-.05em;margin:8px 0 0}.console-caption{font-size:9px;color:#c3d3ee;margin-top:1px;line-height:1.5}.console-line{height:4px;background:#ffffff18;border-radius:20px;overflow:hidden;margin:10px 0 8px}.console-line span{display:block;width:72%;height:100%;background:linear-gradient(90deg,#42d6c7,#7ea7ff);border-radius:20px}.console-foot{display:flex;justify-content:space-between;color:#d9e7ff;font-size:9px}
.metric{position:relative;overflow:hidden;background:#fff;border:1px solid var(--line);border-radius:11px;padding:11px 13px 10px;min-height:79px;box-shadow:var(--shadow)}.metric:before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:linear-gradient(180deg,#35cbbd,#426cf1)}.metric:after{content:"";position:absolute;width:52px;height:52px;right:-29px;top:-30px;border-radius:50%;background:#edf2ff}.metric-label{position:relative;z-index:1;font-size:8px;color:#728098;font-weight:800;letter-spacing:.1em;text-transform:uppercase}.metric-value{position:relative;z-index:1;font-family:'Manrope',sans-serif;font-size:24px;font-weight:800;letter-spacing:-.05em;color:#122440;margin-top:5px;line-height:1}.metric-foot{position:relative;z-index:1;font-size:9px;color:#8290a4;margin-top:5px}
.panel,.info-panel{background:#fff;border:1px solid var(--line);border-radius:11px;padding:13px 15px;margin:0;box-shadow:var(--shadow)}.info-panel{height:100%;background:linear-gradient(145deg,#fff,#f9fbff)}.info-row{display:grid;grid-template-columns:25px 1fr;gap:8px;margin-top:9px;align-items:start}.info-num{width:25px;height:25px;display:flex;align-items:center;justify-content:center;border-radius:7px;background:#edf3ff;color:#315eea;font-size:9px;font-weight:800}.info-title{font-size:10px;font-weight:800;color:#20324f;margin-bottom:2px}.info-copy{font-size:10px;color:#65758d;line-height:1.45}.download-card{background:#fff;border:1px solid var(--line);border-radius:9px;padding:9px 10px;min-height:59px;box-shadow:0 2px 8px #15294b05}.download-title{font-size:10px;font-weight:800;color:#1c2e49;margin-bottom:2px}.download-copy{font-size:9px;color:#77869b;line-height:1.35}
.stButton>button,.stDownloadButton>button{border-radius:8px!important;min-height:36px;font-weight:700!important;font-size:.8rem!important;border:1px solid #d8e1ef!important;box-shadow:0 2px 5px #13294b08;transition:all .16s ease!important;padding:.35rem .65rem!important}.stButton>button:hover,.stDownloadButton>button:hover{border-color:#91aaf5!important;box-shadow:0 4px 11px #2f66f51b!important;transform:translateY(-1px)}.stButton>button[kind="primary"]{background:linear-gradient(105deg,#2859e8,#416ff0)!important;border-color:#3268f5!important;color:#fff!important;box-shadow:0 4px 11px #3268f523!important}.stButton>button[kind="primary"] p{color:#fff!important}
.stTextInput input,.stTextArea textarea,.stDateInput input,[data-baseweb="select"]>div{border-radius:8px!important;background:#fff!important;color:#203047!important;border-color:#d7e0ee!important;min-height:38px}[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] p{color:#2b3b55!important;font-size:.8rem!important;font-weight:700!important;opacity:1!important}[data-testid="stFileUploader"]{background:#fff;border:1px dashed #8ca8e8;border-radius:10px;padding:7px 10px}[data-testid="stFileUploader"] section{background:transparent!important}[data-testid="stDataFrame"]{border:1px solid #dfe6f0;border-radius:9px;overflow:hidden;background:#fff;box-shadow:0 3px 10px #15294b06}[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:4px;border-bottom:1px solid #e1e7f0}[data-testid="stTabs"] [data-baseweb="tab"]{border-radius:7px 7px 0 0;padding:7px 11px;font-weight:700;font-size:.8rem}[data-testid="stTabs"] [aria-selected="true"]{color:#2d60ec!important;border-bottom-color:#3268f5!important}[data-testid="stExpander"]{background:#fff;border:1px solid #dfe6f0;border-radius:10px;overflow:hidden}[data-testid="stAlert"]{border-radius:9px;padding:.55rem .75rem}div[role="radiogroup"] label,div[role="radiogroup"] label p{color:#203047!important;opacity:1!important}[data-testid="stProgressBar"]>div>div{background:linear-gradient(90deg,#38cbbd,#356dff)!important}hr{border-color:#dfe6f0!important;margin:.65rem 0 .3rem!important}footer{visibility:hidden}
@media(max-width:900px){.hero-grid{grid-template-columns:1fr}.hero-console{max-width:100%}.block-container{padding:.5rem .8rem 1.2rem!important}}@media(max-width:600px){.hero{padding:18px 17px;border-radius:13px}.hero-title{font-size:28px}.metric{padding:10px;min-height:76px}.metric-value{font-size:22px}}
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
left, right = st.columns([1.02, 1], gap="medium")
with left:
    st.markdown('<div class="panel"><h3>Start with a structured workbook</h3><p style="font-size:11px;margin:0;color:#6b7a90">Organize scenarios, expected outcomes and safe execution settings in one place.</p></div>', unsafe_allow_html=True)
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
    dl1, dl2 = st.columns(2, gap="small")
    with dl1:
        st.markdown('<div class="download-card"><div class="download-title">01 · Excel template</div><div class="download-copy">Ready-to-fill test case structure</div></div>', unsafe_allow_html=True)
        st.download_button("↓ Download template", xbuf.getvalue(), "testpulse_test_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with dl2:
        st.markdown('<div class="download-card"><div class="download-title">02 · Workbook rules</div><div class="download-copy">Required columns and safety guidance</div></div>', unsafe_allow_html=True)
        st.download_button("↓ Download rules", RULES_TEXT.encode("utf-8"), "TestPulse_Excel_Rules.txt", "text/plain", use_container_width=True)
with right:
    info_html = """<div class='info-panel'><h3>Built for reliable execution</h3>
      <div class='info-row'><div class='info-num'>01</div><div><div class='info-title'>Required fields</div><div class='info-copy'>Test Case ID, Test Case Name, Module, Priority and Expected Result.</div></div></div>
      <div class='info-row'><div class='info-num'>02</div><div><div class='info-title'>Scenario inputs</div><div class='info-copy'>Project, date and hours are read from Excel. Missing inputs are flagged before execution.</div></div></div>
      <div class='info-row'><div class='info-num'>03</div><div><div class='info-title'>Safe by default</div><div class='info-copy'>Saving remains disabled unless explicitly enabled. Never put passwords or session tokens in Excel.</div></div></div>
    </div>"""
    st.markdown(info_html, unsafe_allow_html=True)

st.markdown("### Upload workbook")
st.caption("Drop an .xlsx test suite below. TestPulse validates it before enabling execution.")
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
