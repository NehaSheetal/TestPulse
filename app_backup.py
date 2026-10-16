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
st.set_page_config(page_title="NovaTest | APEX Quality Lab", page_icon="✦", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
:root{--navy:#102542;--blue:#2563eb;--sky:#60a5fa;--ink:#203047;--muted:#738198;--line:#e3eaf4;--paper:#f4f7fc;--white:#fff}
html,body,[data-testid="stAppViewContainer"],.stApp{background:var(--paper)!important;color:var(--ink)!important;font-family:'DM Sans',sans-serif}
[data-testid="stSidebar"]{background:#102542!important;border-right:1px solid #203958}
[data-testid="stSidebar"] *{color:#eaf2ff!important}
.block-container{max-width:1500px;padding:1.8rem 2rem 3rem!important}
h1,h2,h3,h4{font-family:'Manrope',sans-serif;color:var(--navy)!important;letter-spacing:-.025em}
p,label,span{font-family:'DM Sans',sans-serif}
.hero{position:relative;overflow:hidden;background:linear-gradient(125deg,#102542 0%,#163b70 60%,#2563eb 100%);padding:28px 30px;border-radius:22px;color:white;margin-bottom:22px;box-shadow:0 14px 38px #10254220}
.hero:after{content:"";position:absolute;width:260px;height:260px;border:1px solid #ffffff26;border-radius:50%;right:-55px;top:-100px;box-shadow:0 0 0 28px #ffffff0a,0 0 0 58px #ffffff08}
.hero-kicker{font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#9dc5ff}
.hero-title{font-family:'Manrope',sans-serif;font-size:35px;font-weight:800;line-height:1.2;color:#fff;margin:8px 0}
.hero-copy{color:#d7e7ff;font-size:14px;max-width:690px}
.hero-chip{display:inline-block;border:1px solid #ffffff35;background:#ffffff12;border-radius:30px;padding:6px 10px;font-size:11px;margin:16px 6px 0 0;color:#edf5ff}
.metric{background:#fff;border:1px solid var(--line);border-radius:16px;padding:17px 18px;box-shadow:0 4px 18px #1c37600a;min-height:112px}
.metric-label{font-size:12px;color:var(--muted);font-weight:600}
.metric-value{font-family:'Manrope',sans-serif;font-size:29px;font-weight:800;color:var(--navy);margin-top:8px}
.metric-foot{font-size:11px;color:#8997aa;margin-top:3px}
.section-kicker{font-size:10px;font-weight:800;letter-spacing:.14em;color:#2563eb;text-transform:uppercase;margin-bottom:3px}
.panel{background:white;border:1px solid var(--line);border-radius:16px;padding:20px 22px;margin:10px 0 18px}
.small-note{font-size:12px;color:#718096}
.stButton>button,.stDownloadButton>button{border-radius:10px!important;min-height:42px;font-weight:700!important;border:1px solid #d6e1f0!important;transition:all .15s ease}
.stButton>button[kind="primary"]{background:linear-gradient(120deg,#2563eb,#3478f6)!important;border-color:#2563eb!important;color:white!important;box-shadow:0 5px 14px #2563eb25}
.stButton>button:hover,.stDownloadButton>button:hover{transform:translateY(-1px);border-color:#8bb6ff!important}
.stTextInput input,.stTextArea textarea,.stNumberInput input{border-radius:10px!important;background:#fff!important}


/* Fix form label visibility */
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label {
    color: #102542 !important;
    opacity: 1 !important;
}

/* Fix text inside input fields */
.stTextInput input,
.stTextArea textarea,
.stNumberInput input,
.stDateInput input {
    color: #203047 !important;
    -webkit-text-fill-color: #203047 !important;
    background-color: #ffffff !important;
}

/* Fix placeholder text */
.stTextInput input::placeholder,
.stNumberInput input::placeholder,
.stDateInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #738198 !important;
    -webkit-text-fill-color: #738198 !important;
    opacity: 1 !important;
}


[data-baseweb="select"]>div{border-radius:10px!important;background:white!important}
[data-testid="stFileUploader"]{background:#f9fbff;border:1px dashed #a8c7f8;border-radius:13px;padding:12px}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:12px;overflow:hidden}
.stTabs [aria-selected="true"]{color:#2563eb!important}
.stTabs [data-baseweb="tab"] {
    color: #334155 !important;
    opacity: 1 !important;
}

.stTabs [data-baseweb="tab"][aria-selected="true"] {
    color: #2563eb !important;
    font-weight: 600 !important;
}
hr{border-color:#e5ebf4}
footer{visibility:hidden}
</style>
""", unsafe_allow_html=True)

for key, default in [("execution_results", []), ("loaded_cases", []), ("uploaded_workbook_name", "")]:
    st.session_state.setdefault(key, default)

try:
    init_db()
    db_ok = database_enabled()
except Exception as exc:
    db_ok, db_error = False, str(exc)
else:
    db_error = ""

with st.sidebar:
    st.markdown("## ✦ NovaTest")
    st.caption("APEX QUALITY LAB")
    st.markdown("---")
    page = st.radio("WORKSPACE", ["Overview", "Test suite", "Run tests", "Results & evidence", "History & setup"], label_visibility="visible")
    st.markdown("---")
    st.markdown("**SYSTEM STATUS**")
    st.write("🟢 Streamlit interface")
    st.write("🟢 Playwright runner installed when setup completes")
    st.write(("🟢 PostgreSQL connected" if db_ok else "🟡 PostgreSQL not connected"))
    st.caption("NovaTest · Private local workspace")

results = st.session_state.execution_results
passed = sum(r.get("status") == "PASS" for r in results)
failed = sum(r.get("status") == "FAIL" for r in results)
blocked = sum(r.get("status") == "BLOCKED" for r in results)
rate = round(100 * passed / len(results)) if results else 0

st.markdown("""
<div class="hero">
  <div class="hero-kicker">✦ Quality engineering workspace</div>
  <div class="hero-title">NovaTest <span style="color:#8fc0ff">Quality Lab</span></div>
  <div class="hero-copy">A focused automation command centre for your Oracle APEX Timesheet application. Run repeatable browser checks, inspect evidence, and keep a durable history of results.</div>
  <span class="hero-chip">ORACLE APEX</span><span class="hero-chip">PLAYWRIGHT AUTOMATION</span><span class="hero-chip">POSTGRESQL HISTORY</span>
</div>
""", unsafe_allow_html=True)

m = st.columns(4)
for col, label, value, foot in [
    (m[0],"TEST CASES",len(st.session_state.loaded_cases),"In the active workbook"),
    (m[1],"EXECUTED",len(results),"Results in this session"),
    (m[2],"PASS RATE",f"{rate}%","Based on completed runs"),
    (m[3],"NEEDS ATTENTION",failed+blocked,f"{failed} failed · {blocked} blocked"),
]:
    with col:
        st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-foot">{foot}</div></div>',unsafe_allow_html=True)

if page == "Overview":
    st.markdown("### Your testing cockpit")
    st.caption("A guided workspace for the Timesheet app — designed to make test execution, diagnosis, and evidence review straightforward.")
    a,b = st.columns([1.25,1], gap="large")
    with a:
        st.markdown('<div class="panel"><div class="section-kicker">GET STARTED</div><h3>From setup to first test</h3><p>1. Load your Excel test plan<br>2. Sign in through the visible Chrome window<br>3. Run one test before running the full suite<br>4. Review screenshots, logs, and saved history</p></div>',unsafe_allow_html=True)
        if st.button("Open test suite →",type="primary",use_container_width=True):
            st.session_state["nav_override"]="Test suite"; st.rerun()
    with b:
        st.markdown('<div class="panel"><div class="section-kicker">CONNECTIONS</div><h3>Environment health</h3>',unsafe_allow_html=True)
        st.write("🟢 **Oracle APEX target**")
        st.caption("The test runner uses your configured application URL.")
        st.write("🟢 **Local evidence folder**")
        st.caption("Screenshots, logs, and videos are stored under `test_evidence/`.")
        st.write(("🟢 **PostgreSQL history**" if db_ok else "🟡 **PostgreSQL history not active**"))
        st.caption("Configure `DATABASE_URL` in `.env` to persist runs across restarts.")
        st.markdown("</div>",unsafe_allow_html=True)
    st.markdown("### Latest results")
    if results:
        st.dataframe(pd.DataFrame([{"Test ID":r["test_case_id"],"Test":r["test_case_name"],"Status":r["status"],"Duration (s)":r["duration_seconds"],"Executed":r["executed_at"]} for r in results[-8:][::-1]]),use_container_width=True,hide_index=True)
    else:
        st.info("No tests have run in this session yet. Start in **Test suite**, then open **Run tests**.")

elif page == "Test suite":
    st.markdown("### Test suite")
    st.caption("Upload an Excel workbook or use the included starter workbook. The workbook is validated before it can run.")
    default_xlsx = WORKBOOK
    l,r = st.columns([1.35,1],gap="large")
    with l:
        upload = st.file_uploader("Upload test plan (.xlsx)",type=["xlsx"],key="case_workbook")
        if upload is not None:
            try:
                st.session_state.loaded_cases = load_cases_from_excel(upload)
                st.session_state.uploaded_workbook_name = upload.name
                st.success(f"Validated {len(st.session_state.loaded_cases)} test cases.")
            except Exception as exc:
                st.error(f"Workbook validation failed: {exc}")
        elif not st.session_state.loaded_cases and default_xlsx.exists():
            try:
                st.session_state.loaded_cases=load_cases_from_excel(default_xlsx)
                st.session_state.uploaded_workbook_name=default_xlsx.name
            except Exception as exc: st.error(f"Starter workbook could not be loaded: {exc}")
    with r:
        st.markdown('<div class="panel"><div class="section-kicker">WORKBOOK FORMAT</div><h3>Keep your test plan tidy</h3><p>Required: Test Case ID, Test Case Name, Module, Priority, Expected Result.</p><p>Optional: Description, Steps, Automation Key.</p><p class="small-note">Do not put usernames, passwords, or session tokens in Excel.</p></div>',unsafe_allow_html=True)
        if default_xlsx.exists():
            st.download_button("⬇ Download starter workbook",default_xlsx.read_bytes(),"timesheet_test_cases.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    if st.session_state.loaded_cases:
        st.caption(f"Active workbook: {st.session_state.uploaded_workbook_name or 'Uploaded workbook'}")
        st.dataframe(pd.DataFrame([{"Test Case ID":c.get("Test Case ID",""),"Test Case Name":c.get("Test Case Name",""),"Module":c.get("Module",""),"Priority":c.get("Priority",""),"Automation Key":c.get("Automation Key",""),"Expected Result":c.get("Expected Result","")} for c in st.session_state.loaded_cases]),use_container_width=True,hide_index=True)
    else: st.warning("Upload a workbook to enable test execution.")

elif page == "Run tests":
    st.markdown("### Run tests")
    st.caption("Sign in manually using Chrome. NovaTest never asks for or stores your APEX password.")
    with st.container(border=True):
        c1,c2 = st.columns(2,gap="large")
        with c1:
            app_url=st.text_input("Oracle APEX application URL",value=DEFAULT_APP_URL)
            project_name=st.text_input("Existing project name",value="Data Migration and Validation")
            holiday_date=st.date_input("Holiday date to test",value=date.today())
        with c2:
            normal_hours=st.number_input("Regular / weekend / holiday hours",min_value=0.5,max_value=8.0,value=2.0,step=0.5)
            excessive_hours=st.number_input("Hours for over-limit validation",min_value=24.5,max_value=100.0,value=25.0,step=1.0)
            allow_save=st.checkbox("Allow changes to be saved in the APEX application",value=False,help="Leave off for safer dry runs. Turn on only when using a test account and you intend to change records.")
            show_browser=st.checkbox("Show Chrome browser during test",value=True)
    if SESSION_FILE.exists() and STORAGE_STATE_FILE.exists():
        try:
            sess=json.loads(SESSION_FILE.read_text(encoding="utf-8"))
            st.success(f"Saved sign-in state found · {sess.get('saved_at','timestamp unavailable')}")
        except Exception: st.success("Saved sign-in state files found.")
    else: st.warning("Sign in once before running tests.")
    x,y,z=st.columns([1,1,1])
    with x: login_clicked=st.button("🔐 Open Chrome & sign in",type="primary",use_container_width=True)
    with y: clear_clicked=st.button("Clear saved session",use_container_width=True)
    with z: refresh_clicked=st.button("Refresh session status",use_container_width=True)
    if clear_clicked:
        clear_saved_session(); st.success("Local saved browser state cleared."); st.rerun()
    if refresh_clicked: st.rerun()
    if login_clicked:
        if not app_url.strip().startswith(("https://","http://")): st.error("Enter a valid http/https URL.")
        elif not show_browser: st.error("Enable Show Chrome browser during test.")
        else:
            st.info("A Chrome window will open. Sign in normally, then use the APEX navigation menu to open Employee Timesheet if needed. Keep this Streamlit page open.")
            with st.spinner("Waiting for Employee Timesheet to be detected…"):
                try:
                    answer=open_login_session(app_url.strip(),show_browser=True,timeout_seconds=300)
                    st.success(answer["message"]); st.caption(answer.get("authenticated_url","")); st.rerun()
                except Exception as exc: st.error(f"Sign-in was not confirmed: {exc}")
    st.markdown("#### Choose a run")
    cases=st.session_state.loaded_cases
    if not cases and WORKBOOK.exists():
        try:
            cases=load_cases_from_excel(WORKBOOK); st.session_state.loaded_cases=cases; st.session_state.uploaded_workbook_name=WORKBOOK.name
        except Exception as exc: st.error(str(exc))
    if cases:
        ids=[c["Test Case ID"] for c in cases]
        selected=st.selectbox("Test case",ids,format_func=lambda q:next(f'{c["Test Case ID"]} — {c["Test Case Name"]}' for c in cases if c["Test Case ID"]==q))
        chosen=next(c for c in cases if c["Test Case ID"]==selected)
        st.info(f"**Expected result:** {chosen.get('Expected Result','')}")
        a,b=st.columns(2)
        run_one=a.button("▶ Run selected test",type="primary",use_container_width=True)
        run_all=b.button("▶ Run full suite",use_container_width=True)
    else: run_one=run_all=False; selected=None
    if run_one or run_all:
        ready=SESSION_FILE.exists() and STORAGE_STATE_FILE.exists()
        if not ready: st.error("First click Open Chrome & sign in.")
        elif not project_name.strip(): st.error("Enter the project name.")
        else:
            queue=cases if run_all else [c for c in cases if c["Test Case ID"]==selected]
            cfg={"app_url":app_url.strip(),"project_name":project_name.strip(),"holiday_date":holiday_date.isoformat(),"normal_hours":float(normal_hours),"excessive_hours":float(excessive_hours),"allow_save":bool(allow_save),"headless":not show_browser}
            bar=st.progress(0); status=st.empty(); completed=[]
            for i,case in enumerate(queue,1):
                status.info(f"Running {i}/{len(queue)} · {case['Test Case ID']} — {case['Test Case Name']}")
                try: result=run_test_case(case,cfg)
                except Exception as exc:
                    result={"test_case_id":case.get("Test Case ID","UNKNOWN"),"test_case_name":case.get("Test Case Name",""),"status":"FAIL","duration_seconds":0,"actual_result":"Runner error","error_message":str(exc),"screenshot_path":"","video_path":"","logs":[str(exc)],"executed_at":pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")}
                completed.append(result); bar.progress(i/len(queue))
            prior={r["test_case_id"]:r for r in st.session_state.execution_results}
            prior.update({r["test_case_id"]:r for r in completed})
            st.session_state.execution_results=list(prior.values())
            try:
                if db_ok: save_test_run(completed,st.session_state.uploaded_workbook_name or "timesheet_test_cases.xlsx",app_url)
            except Exception as exc: st.warning(f"Results are visible now, but database save failed: {exc}")
            status.success(f"Finished {len(completed)} test(s). Review Results & evidence.")
            st.rerun()
    st.caption("Safety default: saving timesheet changes is disabled. The runner may still open forms and prepare values; enable saving only for a test account.")

elif page == "Results & evidence":
    st.markdown("### Results & evidence")
    if not results: st.info("No results in the current Streamlit session. For saved database runs, open History & setup.")
    else:
        df=pd.DataFrame([{"Test Case ID":r["test_case_id"],"Test Case Name":r["test_case_name"],"Status":r["status"],"Duration (s)":r["duration_seconds"],"Actual Result":r["actual_result"],"Executed At":r["executed_at"],"Error":r["error_message"]} for r in results[::-1]])
        st.dataframe(df,use_container_width=True,hide_index=True)
        opts=[f'{r["test_case_id"]} — {r["test_case_name"]}' for r in results[::-1]]
        pick=st.selectbox("Inspect a result",opts)
        r=next(x for x in results[::-1] if f'{x["test_case_id"]} — {x["test_case_name"]}'==pick)
        status_color={"PASS":"#16845b","FAIL":"#c0394b","BLOCKED":"#a36a05"}.get(r["status"],"#52657d")
        st.markdown(f'<div class="panel"><div class="section-kicker">TEST DETAIL</div><h3>{r["test_case_id"]} · {r["test_case_name"]}</h3><p><b>Status:</b> <span style="color:{status_color};font-weight:800">{r["status"]}</span></p><p>{r["actual_result"]}</p></div>',unsafe_allow_html=True)
        if r.get("error_message"): st.error(r["error_message"])
        t1,t2,t3=st.tabs(["📸 Screenshot","🎞 Video","⌁ Execution logs"])

        with t1:
            screenshot_path = r.get("screenshot_path", "")
            p = Path(screenshot_path) if screenshot_path else None

            if p and p.is_file():
                st.image(
                    str(p),
                    caption=f"Screenshot — {r.get('test_case_id', 'Selected test')}",
                    use_container_width=True,
                )
                st.caption(f"Saved at: {p}")
            else:
                st.warning("Screenshot file was not found for this test result.")
                st.caption(f"Recorded screenshot path: {screenshot_path or 'Not available'}")

        with t2:
            video_path = r.get("video_path", "")
            p = Path(video_path) if video_path else None

            if p and p.is_file() and p.stat().st_size > 0:
                st.video(str(p))
                st.caption(f"Video: {p.name}")
                st.caption(f"Saved at: {p}")
            else:
                st.warning("Video file was not found for this test result.")
                st.caption(f"Recorded video path: {video_path or 'Not available'}")

        with t3:
            execution_logs = r.get("logs", [])

            if execution_logs:
                st.markdown("### Test Execution Steps")
                for index, log_entry in enumerate(execution_logs, start=1):
                    st.markdown(f"**Step {index}**")
                    st.code(str(log_entry), language="text")
            else:
                st.warning("No execution logs were recorded for this test.")
        st.download_button("⬇ Export results as CSV",df.to_csv(index=False).encode("utf-8"),"novatest_results.csv","text/csv")
        buf=BytesIO()
        with pd.ExcelWriter(buf,engine="openpyxl") as writer: df.to_excel(writer,sheet_name="Results",index=False)
        st.download_button("⬇ Export results as Excel",buf.getvalue(),"novatest_results.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "History & setup":
    st.markdown("### History & setup")
    st.caption("PostgreSQL is an optional history store for NovaTest. It does not replace or modify the Oracle APEX database.")
    if db_ok:
        st.success("PostgreSQL is connected. New test runs will be stored in the configured database.")
        try:
            hist=load_history(limit=100)
            if hist: st.dataframe(pd.DataFrame(hist),use_container_width=True,hide_index=True)
            else: st.info("The database is connected, but no test runs have been saved yet.")
        except Exception as exc: st.error(f"Could not read history: {exc}")
    else:
        st.warning("PostgreSQL is not connected. The app can still run locally, but history will not survive an app restart.")
        if db_error: st.code(db_error)
    st.markdown("#### What goes where?")
    st.markdown("""
    - **Oracle APEX:** your real timesheet app, users, projects, approvals, and business rules.
    - **NovaTest + Playwright:** browser actions, screenshots, videos, logs, and test outcomes.
    - **PostgreSQL:** optional durable test-run history and test-result records.
    - **GitHub:** source files and workbook template, not passwords, browser sessions, evidence, or database data.
    """)
    st.markdown("#### First-time setup checklist")
    st.markdown("""
    1. Install Python 3.12 and VS Code.
    2. Open this folder in VS Code.
    3. Run `setup_windows.bat`.
    4. Copy `.env.example` to `.env` and add your PostgreSQL connection string only if PostgreSQL is configured.
    5. Run `run_app.bat`.
    6. In NovaTest, open **Run tests** and sign in using Chrome.
    """)
    st.markdown("#### Database connection test")
    if st.button("Test PostgreSQL connection",use_container_width=False):
        try:
            from database import test_connection
            msg=test_connection()
            st.success(msg)
        except Exception as exc: st.error(f"Connection failed: {exc}")

st.markdown("---")
st.caption("NovaTest Quality Lab · Built for careful, evidence-first testing. Browser state is sensitive and stays local.")
