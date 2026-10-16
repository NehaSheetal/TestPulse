from pathlib import Path
from datetime import datetime, date, timedelta
import json, re, time
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BASE_DIR = Path(__file__).resolve().parent
EVIDENCE_DIR = BASE_DIR / "test_evidence"
EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
SESSION_FILE = BASE_DIR / "apex_test_session.json"
STORAGE_STATE_FILE = BASE_DIR / "apex_storage_state.json"
DEFAULT_APP_URL = "https://oracleapex.com/ords/r/ns_351/timesheet-management/"
REQUIRED_COLUMNS = ["Test Case ID", "Test Case Name", "Module", "Priority", "Expected Result"]
PAGE_MARKERS = ("employee timesheet", "search projects", "add new project")


def load_cases_from_excel(source):
    df = pd.read_excel(source, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))
    if df.empty:
        raise ValueError("The workbook contains no test cases.")
    ids = df["Test Case ID"].astype(str).str.strip()
    if df["Test Case ID"].isna().any() or ids.eq("").any():
        raise ValueError("Every row must have a Test Case ID.")
    if ids.duplicated().any():
        raise ValueError("Test Case IDs must be unique.")
    for col in REQUIRED_COLUMNS:
        if df[col].isna().any() or df[col].astype(str).str.strip().eq("").any():
            raise ValueError(f"Column '{col}' contains blank values.")
    records = df.fillna("").to_dict(orient="records")
    for row in records:
        row["Test Case ID"] = str(row["Test Case ID"]).strip()
        row["Test Case Name"] = str(row["Test Case Name"]).strip()
        row["Automation Key"] = str(row.get("Automation Key", "")).strip() or infer_automation_key(row["Test Case ID"], row["Test Case Name"])
    return records


def infer_automation_key(test_id, test_name):
    s = f"{test_id} {test_name}".lower()
    for marker, key in [
        ("tc-001", "login"), ("login", "login"),
        ("tc-002", "old_entry"), ("three weeks", "old_entry"), ("3 weeks", "old_entry"),
        ("tc-003", "weekend_entry"), ("weekend", "weekend_entry"),
        ("tc-004", "holiday_entry"), ("holiday", "holiday_entry"),
        ("tc-005", "over_24_hours"), ("24 hours", "over_24_hours"), ("over 24", "over_24_hours"),
        ("tc-006", "search_project"), ("search project", "search_project"),
    ]:
        if marker in s: return key
    return "generic_steps"


def _log(logs, msg): logs.append(f"{datetime.now():%H:%M:%S}  {msg}")

def _body(page):
    try: return page.locator("body").inner_text(timeout=3000)
    except Exception: return ""

def _visible(locator, timeout=800):
    try:
        if locator.count() and locator.first.is_visible():
            locator.first.wait_for(state="visible", timeout=timeout)
            return locator.first if locator.first.is_enabled() else None
    except Exception: pass
    return None

def _click_named(page, names, timeout=2500):
    for name in names:
        for loc in (page.get_by_role("button",name=re.compile(re.escape(name),re.I)),page.get_by_role("link",name=re.compile(re.escape(name),re.I)),page.get_by_text(re.compile(re.escape(name),re.I),exact=False)):
            target=_visible(loc)
            if target:
                try: target.click(timeout=timeout); return name
                except Exception: pass
    return None

def _on_timesheet(page):
    text=_body(page).lower()
    return any(x in text for x in PAGE_MARKERS)

def _go_timesheet(page, logs=None):
    if _on_timesheet(page): return True
    clicked=_click_named(page,["Navigation Menu","Open Navigation Menu","Menu"])
    if not clicked:
        for sel in ["#t_Button_navControl","button[aria-label*='Navigation Menu' i]","button[title*='Navigation Menu' i]"]:
            target=_visible(page.locator(sel))
            if target:
                try: target.click(timeout=1500); clicked="navigation menu"; break
                except Exception: pass
    if clicked:
        page.wait_for_timeout(350)
        if _click_named(page,["Employee Timesheet"],3000):
            try: page.wait_for_load_state("domcontentloaded",timeout=12000)
            except Exception: pass
            page.wait_for_timeout(800)
            if _on_timesheet(page):
                if logs is not None: _log(logs,f"Opened Employee Timesheet: {page.url}")
                return True
    # Do not construct a URL from a guessed friendly path; retain APEX navigation/session.
    return _on_timesheet(page)

def open_login_session(app_url=DEFAULT_APP_URL, show_browser=True, timeout_seconds=300):
    """Open a visible Chrome window for manual login, then save Playwright storage state."""
    if not show_browser: raise RuntimeError("Enable 'Show browser window' to sign in manually.")
    with sync_playwright() as p:
        try:
            browser=p.chromium.launch(headless=False,channel="chrome")
        except Exception as exc:
            raise RuntimeError("Could not launch installed Google Chrome. Install Chrome, close any managed launch restrictions, and try again. Details: "+str(exc))
        context=browser.new_context(viewport={"width":1440,"height":950},ignore_https_errors=False)
        page=context.new_page()
        page.set_default_timeout(10000)
        try:
            # DOMContentLoaded avoids waiting for APEX network requests that may never become idle.
            page.goto(app_url,wait_until="domcontentloaded",timeout=45000)
            started=time.monotonic()
            last_navigation_attempt=0
            while time.monotonic()-started < timeout_seconds:
                page.wait_for_timeout(800)
                if _on_timesheet(page): break
                body=_body(page).lower()
                # APEX may land on a blank home page after login. Try its visible navigation
                # menu periodically instead of waiting forever on network-idle.
                if time.monotonic()-last_navigation_attempt > 4 and not any(x in body for x in ("password","sign in","username")):
                    _go_timesheet(page)
                    last_navigation_attempt=time.monotonic()
                # If the user is still on a login form, leave the browser open for manual input.
            else:
                current=_body(page).lower()
                if any(x in current for x in ("password","sign in","username")):
                    raise RuntimeError("The login page was still visible after the wait. Complete sign-in in Chrome, then make sure Employee Timesheet is open before the timeout. If clicking Sign In spins forever, test the APEX URL directly in normal Chrome; this is likely an APEX/network/session issue rather than a Streamlit page-load wait.")
                raise RuntimeError("Could not confirm Employee Timesheet within the timeout. In Chrome, complete sign-in and choose Navigation Menu → Employee Timesheet. If the page itself keeps loading, check the APEX app directly in Chrome.")
            page.wait_for_timeout(500)
            context.storage_state(path=str(STORAGE_STATE_FILE))
            SESSION_FILE.write_text(json.dumps({"app_url":app_url,"authenticated_url":page.url,"saved_at":datetime.now().isoformat(timespec="seconds")},indent=2),encoding="utf-8")
            return {"ok":True,"message":"Employee Timesheet detected; browser session state saved locally.","authenticated_url":page.url}
        finally:
            context.close(); browser.close()

def clear_saved_session():
    SESSION_FILE.unlink(missing_ok=True); STORAGE_STATE_FILE.unlink(missing_ok=True)

def _session():
    if not SESSION_FILE.exists() or not STORAGE_STATE_FILE.exists():
        raise RuntimeError("No saved login session. Open Run tests and click Open Chrome & sign in first.")
    return json.loads(SESSION_FILE.read_text(encoding="utf-8"))

def _week_start(body):
    m=re.search(r"(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{4}",body)
    if not m:return None
    try:return datetime.strptime(m.group(1),"%d %b %Y").date()
    except ValueError:return None

def _navigate_week(page,target,logs):
    for _ in range(24):
        start=_week_start(_body(page))
        if start and start<=target<=start+timedelta(days=6):
            _log(logs,f"Selected week containing {target:%d %b %Y}.");return
        name="Previous Week" if (start and target<start) or (not start and target<date.today()) else "Next Week"
        if not _click_named(page,[name],1800): raise RuntimeError(f"Could not find {name} control on Employee Timesheet.")
        page.wait_for_timeout(350)
    raise RuntimeError("Could not navigate to the requested week.")

def _search_project(page,project,logs):
    field=None
    for loc in [page.get_by_placeholder(re.compile("search project",re.I)),page.get_by_label(re.compile("search project",re.I)),page.locator("input[placeholder*='Search project' i]"),page.locator("input[aria-label*='Search project' i]")]:
        field=_visible(loc,1000)
        if field:break
    if not field:
        if project.casefold() in _body(page).casefold():return True
        raise RuntimeError("Could not identify the Search Projects input.")
    field.fill(project);page.wait_for_timeout(650)
    if project.casefold() in _body(page).casefold():_log(logs,f"Project found: {project}");return True
    raise RuntimeError(f"Project '{project}' did not appear after searching.")

def _open_cell(page,project,target,logs):
    _navigate_week(page,target,logs);_search_project(page,project,logs)
    match=page.get_by_text(re.compile(re.escape(project),re.I),exact=False).first
    match.wait_for(state="visible",timeout=5000)
    row=match.locator("xpath=ancestor::tr[1]")
    if row.count()==0:row=match.locator("xpath=ancestor::*[self::div or self::li][1]")
    col_index,header=None,""
    for selector in ["th","[role='columnheader']"]:
        headers=page.locator(selector).all_inner_texts()
        for i,value in enumerate(headers):
            txt=value.lower().replace("\n"," ")
            if target.strftime("%a").lower() in txt and str(target.day) in txt and target.strftime("%b").lower() in txt:
                col_index,header=i,value;break
        if col_index is not None:break
    if col_index is None:raise RuntimeError(f"Could not identify date column for {target:%a %d %b}.")
    cells=row.locator("td")
    if col_index>=cells.count():col_index-=1
    if col_index<0 or col_index>=cells.count():raise RuntimeError(f"Could not map date column '{header}' to a project cell.")
    cells.nth(col_index).click(timeout=4000);page.wait_for_timeout(350)
    _log(logs,f"Opened {target:%A %d %b %Y} for {project}.")

def _hours(page):
    for loc in [page.get_by_label(re.compile(r"hours|hour|time",re.I)),page.get_by_placeholder(re.compile(r"hours|hour|time",re.I)),page.locator("input[name*='HOUR' i], input[id*='HOUR' i]"),page.locator("input[type='number']"),page.locator("input[type='text']"),page.locator("input:not([type])")]:
        field=_visible(loc,1000)
        if field:return field
    raise RuntimeError("Could not identify the hours field in the selected cell dialog.")

def _save(page,allow_save,logs):
    if not allow_save:_log(logs,"Save skipped because saving is disabled.");return "BLOCKED"
    clicked=_click_named(page,["Save","Apply Changes","Submit"],3000)
    if not clicked:raise RuntimeError("Could not find Save/Apply Changes/Submit.")
    page.wait_for_timeout(1000);body=_body(page).lower()
    if any(x in body for x in ["cannot exceed","must not exceed","maximum","required","invalid","error"]):return "VALIDATION"
    return "SAVED"

def _generic(page,case,logs):
    steps=[x.strip() for x in re.split(r"[;\n]+",str(case.get("Steps","") or "")) if x.strip()]
    if not steps:raise RuntimeError("This row has no Steps and no supported Automation Key.")
    for step in steps:
        low=step.lower()
        if low in ("open application","launch application","navigate to application"):pass
        elif low.startswith("click ") and not _click_named(page,[step[6:].strip()],3000):raise RuntimeError(f"Could not click '{step[6:].strip()}'.")
        elif low.startswith("open ") and not _click_named(page,[step[5:].strip()],3000):raise RuntimeError(f"Could not open '{step[5:].strip()}'.")
        elif low.startswith("wait "):
            m=re.search(r"wait\s+(\d+)",low);page.wait_for_timeout(int(m.group(1))*1000 if m else 1000)
        elif low.startswith("verify text "):page.get_by_text(step[len("verify text "):].strip().strip("'\""),exact=False).first.wait_for(state="visible",timeout=8000)
        elif low.startswith("enter ") and " into " in low:
            m=re.match(r"""enter\s+["']?(.+?)["']?\s+into\s+(.+)$""",step,re.I)
            if not m:raise RuntimeError(f"Invalid step: {step}")
            value,label=m.group(1).strip().strip("'\""),m.group(2).strip().strip("'\"")
            target=_visible(page.get_by_label(label,exact=False),1500) or _visible(page.get_by_placeholder(label,exact=False),1500)
            if not target:raise RuntimeError(f"Could not find field '{label}'.")
            target.fill(value)
        else:raise RuntimeError(f"Unsupported generic step: {step}")
        _log(logs,f"Completed: {step}")

def _scenario(page,case,cfg,logs):
    key=str(case.get("Automation Key","")).lower() or infer_automation_key(case.get("Test Case ID",""),case.get("Test Case Name",""))
    if key=="login":
        if _on_timesheet(page):return "PASS","Authenticated Employee Timesheet page is visible."
        raise RuntimeError("Employee Timesheet was not visible after restoring the session.")
    if key=="search_project":
        _navigate_week(page,date.today(),logs);_search_project(page,cfg["project_name"],logs);return "PASS",f"Project '{cfg['project_name']}' was found."

    if key == "old_entry":
        target = date.today() - timedelta(days=21)

        # Open a timesheet cell from three weeks ago.
        _open_cell(page, cfg["project_name"], target, logs)

        # A week older than two weeks should be read-only.
        try:
            popup = page.get_by_text(
                "Read-Only Week",
                exact=True
            )
            popup.wait_for(state="visible", timeout=5000)

            body_text = _body(page)
            expected_message = (
                "Timesheet entries can only be edited for the "
                "current week and the previous two weeks."
            )

            if expected_message not in body_text:
                alternate_message = (
                    "This timesheet is read-only. You cannot "
                    "enter or edit hours for weeks older than 2 weeks."
                )
                if alternate_message not in body_text:
                    raise RuntimeError(
                        "Read-Only Week popup appeared, but its "
                        "message did not match the expected text."
                    )

            got_it = page.get_by_role(
                "button",
                name="Got it",
                exact=True
            )
            got_it.wait_for(state="visible", timeout=3000)
            got_it.click()

            logs.append(
                f"Verified read-only restriction for {target} "
                "and clicked Got it."
            )
            return "PASS", (
                f"Correctly displayed the Read-Only Week popup "
                f"for {target} and closed it using Got it."
            )

        except Exception as exc:
            raise RuntimeError(
                f"Read-only week validation failed for {target}: {exc}"
            ) from exc
    if key=="weekend_entry":
        today=date.today();target=today-timedelta(days=((today.weekday()-5)%7 or 7))
        _open_cell(page,cfg["project_name"],target,logs);_hours(page).fill(str(cfg["normal_hours"]));saved=_save(page,cfg["allow_save"],logs)
        return ("BLOCKED",f"Prepared weekend entry for {target}; not saved.") if saved=="BLOCKED" else ("PASS",f"Weekend entry action completed for {target}; verify approval status.")
    if key=="holiday_entry":
        target=date.fromisoformat(cfg["holiday_date"]);_open_cell(page,cfg["project_name"],target,logs);_hours(page).fill(str(cfg["normal_hours"]));saved=_save(page,cfg["allow_save"],logs)
        return ("BLOCKED",f"Prepared holiday entry for {target}; not saved.") if saved=="BLOCKED" else ("PASS",f"Holiday entry action completed for {target}; verify approval status.")
    if key=="over_24_hours":
        target=date.today();_open_cell(page,cfg["project_name"],target,logs);_hours(page).fill(str(cfg["excessive_hours"]))
        if not cfg["allow_save"]:
            body=_body(page).lower()
            if any(x in body for x in ["24 hours","maximum","cannot exceed","invalid","error"]):return "PASS","Application displayed excessive-hours validation."
            return "BLOCKED","Excessive hours entered but not submitted because saving is disabled."
        saved=_save(page,True,logs)
        if saved=="VALIDATION":return "PASS","Application displayed a validation message for excessive hours."
        raise RuntimeError("No validation message detected after submitting excessive hours.")
    if key=="generic_steps":_generic(page,case,logs);return "PASS","All generic steps completed."
    raise RuntimeError(f"Unsupported Automation Key '{key}'.")

def run_test_case(case,settings):
    started=time.monotonic();logs=[];status="FAIL";actual="";error=""
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    folder=EVIDENCE_DIR/f"{re.sub(r'[^A-Za-z0-9_-]+','_',str(case.get('Test Case ID','test')))}_{stamp}"
    folder.mkdir(parents=True,exist_ok=True);shot=folder/"final_screenshot.png";video_dir=folder/"video";video_dir.mkdir(exist_ok=True);video=""
    info=_session()
    with sync_playwright() as p:
        launch_args={"headless":bool(settings.get("headless",False))}
        if not launch_args["headless"]:launch_args["channel"]="chrome"
        try:browser=p.chromium.launch(**launch_args)
        except Exception as exc:raise RuntimeError("Could not launch browser for test execution. Check Chrome/Playwright installation: "+str(exc))
        context=browser.new_context(viewport={"width":1440,"height":950},record_video_dir=str(video_dir),record_video_size={"width":1440,"height":950},storage_state=str(STORAGE_STATE_FILE))
        page=context.new_page();page.set_default_timeout(12000)
        try:
            # APEX may keep polling in the background; do not wait for networkidle.
            page.goto(info.get("authenticated_url") or settings.get("app_url") or DEFAULT_APP_URL,wait_until="domcontentloaded",timeout=45000)
            page.wait_for_timeout(600)
            if not _on_timesheet(page) and not _go_timesheet(page,logs):raise RuntimeError("Saved session did not open Employee Timesheet. It may have expired; sign in again.")
            status,actual=_scenario(page,case,settings,logs);_log(logs,f"Test result: {status} — {actual}")
        except Exception as exc:
            error=str(exc);actual="Automation could not complete the scenario.";status="FAIL";_log(logs,f"FAIL: {error}")
        finally:
            try:page.screenshot(path=str(shot),full_page=True,timeout=10000);_log(logs,f"Screenshot saved: {shot}")
            except Exception as exc:_log(logs,f"Screenshot failed: {exc}")
            context.close();browser.close()
    videos=list(video_dir.glob("*.webm"))
    if videos:video=str(videos[0])
    return {"test_case_id":str(case.get("Test Case ID","UNKNOWN")),"test_case_name":str(case.get("Test Case Name","")),"status":status,"duration_seconds":round(time.monotonic()-started,2),"actual_result":actual,"error_message":error,"screenshot_path":str(shot) if shot.exists() else "","video_path":video,"logs":logs,"executed_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
