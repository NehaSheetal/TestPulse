from pathlib import Path

from datetime import datetime, date, timedelta

import json, re, time

import pandas as pd

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

import subprocess

import imageio_ffmpeg



BASE_DIR = Path(__file__).resolve().parent

EVIDENCE_DIR = BASE_DIR / "test_evidence"

EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)

SESSION_FILE = BASE_DIR / "apex_test_session.json"

STORAGE_STATE_FILE = BASE_DIR / "apex_storage_state.json"

DEFAULT_APP_URL = "https://oracleapex.com/ords/r/ns_351/timesheet-management/"

REQUIRED_COLUMNS = ["Test Case ID", "Test Case Name", "Module", "Priority", "Expected Result"]

PAGE_MARKERS = ("employee timesheet", "search projects", "add new project")











def load_fusion_cases_from_excel(source):

    """Load Fusion test steps from the dedicated Fusion workbook."""

    required_columns = [

        "TEST_CASE_ID",

        "APPLICATION",

        "TEST_CASE_NAME",

        "STEP_NUMBER",

        "ACTION",

        "FIELD_NAME",

        "TEST_DATA",

        "EXPECTED_RESULT",

        "NOTES",

    ]



    sheets = pd.read_excel(source, sheet_name=None, engine="openpyxl")

    if "Fusion_Test_Steps" not in sheets:

        raise ValueError(

            "Fusion workbook must contain a 'Fusion_Test_Steps' sheet."

        )



    df = sheets["Fusion_Test_Steps"].copy()

    df.columns = [str(column).strip() for column in df.columns]



    missing = [column for column in required_columns if column not in df.columns]

    if missing:

        raise ValueError(

            "Fusion workbook is missing columns: " + ", ".join(missing)

        )



    if df.empty:

        raise ValueError("The Fusion_Test_Steps sheet contains no steps.")



    df = df[required_columns].copy()



    for column in ["TEST_CASE_ID", "APPLICATION", "TEST_CASE_NAME", "ACTION"]:

        if df[column].isna().any() or df[column].astype(str).str.strip().eq("").any():

            raise ValueError(

                f"Fusion column '{column}' contains blank values."

            )



    df["STEP_NUMBER"] = pd.to_numeric(df["STEP_NUMBER"], errors="raise")

    if (df["STEP_NUMBER"] < 1).any():

        raise ValueError("Fusion STEP_NUMBER values must be at least 1.")



    if df.duplicated(subset=["TEST_CASE_ID", "STEP_NUMBER"]).any():

        raise ValueError(

            "Each Fusion test case must have unique step numbers."

        )



    return df.fillna("").sort_values(

        ["TEST_CASE_ID", "STEP_NUMBER"]

    ).to_dict(orient="records")















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





def open_login_session(

    app_url=DEFAULT_APP_URL,

    username="",

    password="",

    show_browser=True,

    timeout_seconds=60

):

    """Sign in to Oracle APEX using credentials supplied by TestPulse."""



    if not username.strip() or not password:

        raise RuntimeError("Enter your APEX username and password.")



    with sync_playwright() as p:

        browser = p.chromium.connect_over_cdp(

            "http://localhost:9222"

        )



        if not browser.contexts:

            raise RuntimeError("No Chrome context found on port 9222.")



        context = browser.contexts[0]

        pages = context.pages

        page = pages[0] if pages else context.new_page()



        try:

            page.goto(

                app_url,

                wait_until="domcontentloaded",

                timeout=60000

            )

            page.wait_for_timeout(1000)



            password_field = page.locator(

                "input[type='password']"

            ).first



            if password_field.is_visible(timeout=3000):

                username_field = page.locator(

                    "input[autocomplete='username'], "

                    "input[name*='USERNAME' i], "

                    "input[id*='USERNAME' i], "

                    "input[type='email'], "

                    "input[type='text']"

                ).first



                if not username_field.is_visible():

                    raise RuntimeError(

                        "Could not identify the APEX username field."

                    )



                username_field.fill(username)

                password_field.fill(password)



                submit = page.get_by_role(

                    "button",

                    name=re.compile(r"sign in|log in|login|submit", re.I)

                )



                if submit.count() > 0 and submit.first.is_visible():

                    submit.first.click()

                else:

                    password_field.press("Enter")



            started = time.monotonic()



            while time.monotonic() - started < timeout_seconds:

                page.wait_for_timeout(1000)



                if _on_timesheet(page):

                    context.storage_state(

                        path=str(STORAGE_STATE_FILE)

                    )



                    SESSION_FILE.write_text(

                        json.dumps({

                            "app_url": app_url,

                            "authenticated_url": page.url,

                            "saved_at": datetime.now().isoformat(

                                timespec="seconds"

                            )

                        }, indent=2),

                        encoding="utf-8"

                    )



                    return {

                        "ok": True,

                        "message": "APEX login successful. Session saved.",

                        "authenticated_url": page.url

                    }



            raise RuntimeError(

                "Login was not confirmed. Check your credentials, "

                "login page, or any MFA/SSO prompt."

            )



        finally:

            # Keep the user's existing Chrome window open.

            pass

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







def _open_cell(page, project, target, logs):

    # Close any existing Timesheet Entry dialog that may block cell clicks.

    dialog = page.locator(

        ".ui-dialog[aria-modal='true']"

    ).filter(has=page.locator('iframe[title="Timesheet Entry"]'))



    try:

        if dialog.count() > 0 and dialog.last.is_visible():

            close_button = dialog.last.locator(

                "button.ui-dialog-titlebar-close"

            )

            if close_button.count() > 0:

                close_button.click(timeout=3000)

                page.wait_for_timeout(500)

    except Exception as exc:

        _log(logs, f"Existing dialog close attempt: {exc}")



    # Navigate to the week containing the requested date.

    _navigate_week(page, target, logs)



    # Find the requested project row.

    project_locator = page.get_by_text(

        re.compile(re.escape(project), re.I),

        exact=False

    )



    project_visible = False

    for index in range(project_locator.count()):

        candidate = project_locator.nth(index)

        try:

            if candidate.is_visible():

                project_visible = True

                break

        except Exception:

            continue



    if not project_visible:

        _search_project(page, project, logs)



    match = page.get_by_text(

        re.compile(re.escape(project), re.I),

        exact=False

    ).first



    match.wait_for(state="visible", timeout=10000)

    row = match.locator("xpath=ancestor::tr[1]")



    if row.count() == 0:

        raise RuntimeError(

            f"Could not find a table row for project '{project}'."

        )



    # The APEX timesheet columns are: Project, Monday...Sunday, Weekly Total.

    # Monday=0 through Sunday=6; the first td is the project name.

    weekday_index = target.weekday()

    cell_index = weekday_index + 1



    cells = row.locator("td")



    if cells.count() <= cell_index:

        raise RuntimeError(

            f"Project row has {cells.count()} cells; "

            f"cannot select {target:%A} at index {cell_index}."

        )



    cell = cells.nth(cell_index)

    cell.wait_for(state="visible", timeout=5000)

    cell.click(timeout=10000)



    # Wait for the modal iframe to open.

    page.locator(

        'iframe[title="Timesheet Entry"]'

    ).wait_for(state="visible", timeout=10000)



    _log(

        logs,

        f"Opened {target:%A %d %b %Y} for {project}."

    )





def _hours(page):

    """Find the Hours input in the APEX Page 4 Timesheet Entry dialog."""



    # Wait up to 10 seconds for the Page 4 dialog to load.

    for _ in range(40):

        for frame in page.frames:

            try:

                field = frame.locator("#P4_HOURS")



                if field.count() > 0 and field.is_visible() and field.is_enabled():

                    return field



            except Exception:

                continue



        page.wait_for_timeout(250)



    # Include diagnostic details if the expected field is missing.

    frame_details = [

        {

            "url": frame.url,

            "name": frame.name,

            "hours_field_count": frame.locator("#P4_HOURS").count(),

        }

        for frame in page.frames

    ]



    raise RuntimeError(

        "Could not find the Page 4 Hours input (#P4_HOURS). "

        "The Timesheet Entry dialog may not have opened. "

        f"Browser frames: {frame_details!r}"

    )





def _save(page, allow_save, logs):

    if not allow_save:

        _log(logs, "Saving is disabled in TestPulse.")

        return "BLOCKED"



    # The Timesheet Entry page is displayed inside an APEX iframe.

    frame = None

    for current_frame in page.frames:

        try:

            if current_frame.locator("#P4_HOURS").count() > 0:

                frame = current_frame

                break

        except Exception:

            continue



    if frame is None:

        raise RuntimeError(

            "Timesheet Entry iframe or P4_HOURS field was not found."

        )



    # Find the Save button inside the iframe.

    save_button = frame.get_by_role(

        "button",

        name=re.compile(r"save|apply changes|submit", re.I)

    ).first



    if save_button.count() == 0:

        save_button = frame.locator(

            "button[type='submit'], input[type='submit']"

        ).first



    if save_button.count() == 0:

        raise RuntimeError(

            "Could not find a Save button inside the Timesheet Entry iframe."

        )



    save_button.click(timeout=10000)

    page.wait_for_timeout(1500)



    _log(logs, "Clicked Save in the Timesheet Entry dialog.")

    return "SAVED"



def _generic(page,case,logs):

    steps=[x.strip() for x in re.split(r"[;**\n**]+",str(case.get("Steps","") or "")) if x.strip()]

    if not steps:raise RuntimeError("This row has no Steps and no supported Automation Key.")

    for step in steps:

        low=step.lower()

        if low in ("open application","launch application","navigate to application"):pass

        elif low.startswith("click ") and not _click_named(page,[step[6:].strip()],3000):raise RuntimeError(f"Could not click '{step[6:].strip()}'.")

        elif low.startswith("open ") and not _click_named(page,[step[5:].strip()],3000):raise RuntimeError(f"Could not open '{step[5:].strip()}'.")

        elif low.startswith("wait "):

            m=re.search(r"wait\s+(\d+)",low);page.wait_for_timeout(int(m.group(1)) * 1000 if m else 1000)

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



        # Navigate to the historical week without searching for a project.

        _navigate_week(page, target, logs)



        # Click a date cell in the timesheet grid to trigger the restriction.

        weekday_index = target.weekday()

        cell_index = weekday_index + 1

        clicked = False



        rows = page.locator("table tbody tr")



        for row_index in range(rows.count()):

            row = rows.nth(row_index)



            try:

                if not row.is_visible():

                    continue



                cells = row.locator("td")



                if cells.count() <= cell_index:

                    continue



                cell = cells.nth(cell_index)



                if cell.is_visible():

                    cell.click(timeout=5000)

                    clicked = True

                    _log(

                        logs,

                        f"Clicked a timesheet date cell for {target:%d %b %Y}."

                    )

                    break



            except Exception as exc:

                _log(logs, f"Skipping a grid row: {exc}")



        if not clicked:

            raise RuntimeError(

                "Could not find a date cell to trigger the Read-Only Week popup."

            )



        page.wait_for_timeout(1000)



        expected_messages = [

            "Timesheet entries can only be edited for the current week and the previous two weeks.",

            "This timesheet is read-only. You cannot enter or edit hours for weeks older than 2 weeks."

        ]



        popup_found = False

        popup_closed = False

        popup_text = ""



        # Detect the popup and click Got it.

        for frame in page.frames:

            try:

                body_text = frame.locator("body").inner_text(timeout=3000)



                if (

                    "Read-Only Week" in body_text

                    or any(message in body_text for message in expected_messages)

                ):

                    popup_found = True

                    popup_text = body_text



                    got_it = frame.get_by_role(

                        "button", name="Got it", exact=True

                    )



                    for i in range(got_it.count()):

                        button = got_it.nth(i)



                        if button.is_visible():

                            button.click(timeout=5000)

                            page.wait_for_timeout(500)

                            popup_closed = True

                            _log(logs, "Clicked Got it to close the Read-Only Week popup.")

                            break



                    break



            except Exception as exc:

                _log(logs, f"Popup check skipped for a frame: {exc}")



        if not popup_found:

            raise RuntimeError(

                "The Read-Only Week popup was not detected after clicking the date cell."

            )



        if not any(message in popup_text for message in expected_messages):

            raise RuntimeError(

                "The popup appeared, but its message did not match the expected restriction."

            )



        if not popup_closed:

            raise RuntimeError(

                "The popup was detected, but the Got it button was not clicked."

            )



        _log(logs, f"Verified the read-only restriction for {target:%d %b %Y}.")



        return "PASS", (

            f"Read-only restriction verified for {target:%d %b %Y}; "

            "Got it was clicked and the popup was closed."

        )





    if key == "weekend_entry":

        # Find the most recent Saturday, including today if it is Saturday.

        today = date.today()

        target = today - timedelta(days=(today.weekday() - 5) % 7)



        # Open the selected weekend cell.

        _open_cell(page, cfg["project_name"], target, logs)



        # Wait for the Page 4 Hours field to become available.

        hours_field = _hours(page)

        hours_field.fill(str(cfg["normal_hours"]))



        # Find the Page 4 description field in the same frame.

        description_field = None

        for frame in page.frames:

            try:

                candidate = frame.locator("#P4_DESCRIPTION")

                if (

                    candidate.count() > 0

                    and candidate.is_visible()

                    and candidate.is_enabled()

                ):

                    description_field = candidate

                    break

            except Exception:

                continue



        if description_field is None:

            raise RuntimeError(

                "Page 4 opened, but the Description field "

                "(#P4_DESCRIPTION) was not found."

            )



        description_field.fill("Automated weekend entry test")



        # Avoid submitting changes unless explicitly allowed.

        if not cfg["allow_save"]:

            return (

                "BLOCKED",

                f"Weekend entry prepared for {target}; saving is disabled."

            )



        saved = _save(page, True, logs)



        if saved == "BLOCKED":

            return "BLOCKED", "Weekend entry was not saved."



        # Wait for the dialog to close after submission.

        page.wait_for_timeout(1500)



        _log(logs, f"Weekend entry submitted for {target}.")



        return (

            "PASS",

            f"Weekend entry submitted for {target}. "

            "Pending status still needs to be verified."

        )

    if key=="holiday_entry":

        target=date.fromisoformat(cfg["holiday_date"]);_open_cell(page,cfg["project_name"],target,logs);_hours(page).fill(str(cfg["normal_hours"]));saved=_save(page,cfg["allow_save"],logs)

        return ("BLOCKED",f"Prepared holiday entry for {target}; not saved.") if saved=="BLOCKED" else ("PASS",f"Holiday entry action completed for {target}; verify approval status.")



    if key == "over_24_hours":

        target = date.today()



        _open_cell(page, cfg["project_name"], target, logs)



        hours_field = _hours(page)

        hours_field.fill(str(cfg["excessive_hours"]))



        # Check validation before submitting.

        frame = hours_field.locator("xpath=ancestor::form")

        dialog_text = ""



        for current_frame in page.frames:

            try:

                if current_frame.locator("#P4_HOURS").count() > 0:

                    dialog_text = current_frame.locator("body").inner_text()

                    break

            except Exception:

                continue



        if not cfg["allow_save"]:

            if any(msg in dialog_text.lower() for msg in [

                "between 0 and 24",

                "cannot exceed",

                "maximum",

                "invalid"

            ]):

                return "PASS", "Application displayed excessive-hours validation."

            return "BLOCKED", "Excessive hours entered; saving is disabled."



        saved = _save(page, True, logs)



        # Read validation text from the dialog iframe.

        for current_frame in page.frames:

            try:

                if current_frame.locator("#P4_HOURS").count() > 0:

                    current_frame.wait_for_timeout(500)

                    dialog_text = current_frame.locator("body").inner_text()

                    break

            except Exception:

                continue



        if any(msg in dialog_text.lower() for msg in [

            "between 0 and 24",

            "cannot exceed",

            "maximum",

            "future dates cannot be entered"

        ]):

            # Close the Timesheet Entry popup after confirming validation.

            close_button = page.locator(

                '.ui-dialog[aria-modal="true"] button.ui-dialog-titlebar-close'

            ).last



            if close_button.is_visible():

                close_button.click(timeout=5000)

                page.wait_for_timeout(500)



            _log(logs, "Validation detected; closed the Timesheet Entry popup.")



            return "PASS", (

                "Application displayed validation for excessive hours and "

                "the Timesheet Entry popup was closed."

            )

        raise RuntimeError(

            "No expected validation message detected in the Timesheet Entry dialog."

        )

    if key=="generic_steps":_generic(page,case,logs);return "PASS","All generic steps completed."

    raise RuntimeError(f"Unsupported Automation Key '{key}'.")







def run_test_case(case, settings):

    started = time.monotonic()

    logs = []

    status = "FAIL"

    actual = ""

    error = ""



    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    folder = EVIDENCE_DIR / (

        f"{re.sub(r'[^A-Za-z0-9_-]+','_',str(case.get('Test Case ID','test')))}_{stamp}"

    )

    folder.mkdir(parents=True, exist_ok=True)



    shot = folder / "final_screenshot.png"

    video_path = folder / "test_run.mp4"

    info = _session()



    import subprocess



    ffmpeg_process = None

    browser = None

    context = None

    page = None



    with sync_playwright() as p:

        try:

            browser = p.chromium.connect_over_cdp("http://localhost:9222")



            if not browser.contexts:

                raise RuntimeError("No Chrome context found on port 9222.")



            context = browser.contexts[0]

            pages = context.pages

            page = pages[0] if pages else context.new_page()

            page.set_default_timeout(12000)

            page.bring_to_front()



            app_url = (

                info.get("authenticated_url")

                or settings.get("app_url")

                or DEFAULT_APP_URL

            )



            if not _on_timesheet(page):

                page.goto(app_url, wait_until="domcontentloaded", timeout=60000)



            page.wait_for_timeout(1000)



            if not _on_timesheet(page) and not _go_timesheet(page, logs):

                raise RuntimeError(

                    "Employee Timesheet was not detected in connected Chrome."

                )



            # Record the visible desktop while the test runs.

            try:

                ffmpeg_process = subprocess.Popen(

                    [

                        imageio_ffmpeg.get_ffmpeg_exe(), "-y",

                        "-f", "gdigrab",

                        "-framerate", "15",

                        "-i", "desktop",

                        "-c:v", "libx264",

                        "-preset", "ultrafast",

                        "-pix_fmt", "yuv420p",

                        str(video_path)

                    ],

                    stdin=subprocess.PIPE,

                    stdout=subprocess.DEVNULL,

                    stderr=subprocess.DEVNULL

                )

                time.sleep(1)

                _log(logs, "Desktop video recording started.")

            except Exception as exc:

                _log(logs, f"Video recording could not start: {exc}")

                ffmpeg_process = None



            status, actual = _scenario(page, case, settings, logs)

            _log(logs, f"Test result: {status} — {actual}")



        except Exception as exc:

            error = str(exc)

            actual = "Automation could not complete the scenario."

            status = "FAIL"

            _log(logs, f"FAIL: {error}")



        finally:

            if page is not None:

                try:

                    page.bring_to_front()

                    page.screenshot(

                        path=str(shot),

                        full_page=True,

                        timeout=15000

                    )

                    _log(logs, f"Screenshot saved: {shot}")

                except Exception as exc:

                    _log(logs, f"Screenshot failed: {exc}")



            if ffmpeg_process is not None:

                try:

                    ffmpeg_process.stdin.write(b"q\n")

                    ffmpeg_process.stdin.flush()

                    ffmpeg_process.wait(timeout=15)

                except Exception as exc:

                    ffmpeg_process.kill()

                    _log(logs, f"Video recording stop issue: {exc}")



            # Do not close the connected Chrome browser or its context.



    return {

        "test_case_id": str(case.get("Test Case ID", "UNKNOWN")),

        "test_case_name": str(case.get("Test Case Name", "")),

        "status": status,

        "duration_seconds": round(time.monotonic() - started, 2),

        "actual_result": actual,

        "error_message": error,

        "screenshot_path": str(shot) if shot.is_file() else "",

        "video_path": (

            str(video_path)

            if video_path.is_file() and video_path.stat().st_size > 0

            else ""

        ),

        "logs": logs,

        "executed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    }

# ---------------------------------------------------------------------------
# Generic application support. Existing Timesheet/APEX functions above remain
# unchanged; these functions are used for step-based workbooks from other apps.
# ---------------------------------------------------------------------------

def open_application_session(app_url, username="", password="", timeout_seconds=12):
    """Open an application and use runtime credentials when a standard login form is available.

    Searches the main document and child frames because some Oracle Fusion identity
    providers render the login controls inside an iframe. SSO/MFA may still require
    manual completion. Credentials are never written to disk or logged.
    """
    app_url = str(app_url or "").strip()
    username = str(username or "").strip()
    password = str(password or "")
    if not app_url.startswith(("https://", "http://")):
        raise RuntimeError("Enter a valid application URL beginning with http:// or https://.")

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
        except Exception as exc:
            raise RuntimeError(
                "Could not connect to Chrome on port 9222. Start Chrome using the TestPulse remote-debugging command, then try again."
            ) from exc
        if not browser.contexts:
            raise RuntimeError("No Chrome context found on port 9222.")
        context = browser.contexts[0]
        pages = context.pages
        page = pages[-1] if pages else context.new_page()
        page.set_default_timeout(10000)
        page.bring_to_front()
        page.goto(app_url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(1500)

        def visible_enabled(locator):
            try:
                return locator.count() > 0 and locator.first.is_visible() and locator.first.is_enabled()
            except Exception:
                return False

        # Search the top-level document and all identity-provider frames.
        frames = list(page.frames)
        password_field = None
        password_frame = None
        for frame in frames:
            for selector in [
                "input[type='password']",
                "input[autocomplete='current-password']",
                "input[name*='password' i]",
                "input[id*='password' i]",
            ]:
                candidate = frame.locator(selector)
                if visible_enabled(candidate):
                    password_field = candidate.first
                    password_frame = frame
                    break
            if password_field is not None:
                break

        if password_field is not None:
            if not username or not password:
                raise RuntimeError(
                    "A login form is visible, but the TestPulse Username and Password fields are empty. "
                    "Enter them in the dashboard and click Connect / sign in again."
                )

            user_field = None
            for selector in [
                "input[autocomplete='username']",
                "input[autocomplete='email']",
                "input[type='email']",
                "input[name*='user' i]",
                "input[id*='user' i]",
                "input[name*='email' i]",
                "input[id*='email' i]",
                "input[type='text']",
            ]:
                candidate = password_frame.locator(selector)
                if visible_enabled(candidate):
                    user_field = candidate.first
                    break

            if user_field is None:
                raise RuntimeError(
                    "Found the Fusion password field, but could not identify the username field on the same login form. "
                    "The identity provider may use a custom SSO page; complete that sign-in manually."
                )

            user_field.fill(username)
            password_field.fill(password)

            submit = None
            for locator in [
                password_frame.get_by_role("button", name=re.compile(r"sign in|log in|login|continue|submit|next", re.I)),
                password_frame.locator("input[type='submit']"),
                password_frame.locator("button[type='submit']"),
            ]:
                if visible_enabled(locator):
                    submit = locator.first
                    break
            if submit is not None:
                submit.click()
            else:
                password_field.press("Enter")

            page.wait_for_timeout(max(1500, min(int(timeout_seconds), 8) * 1000))

            # A remaining password field usually means invalid credentials or another
            # identity-provider step (for example MFA). Never claim login succeeded here.
            still_password = False
            for frame in page.frames:
                if visible_enabled(frame.locator("input[type='password']")):
                    still_password = True
                    break
            if still_password:
                return {
                    "ok": True,
                    "message": "TestPulse filled and submitted the credentials. A password form is still visible; check the browser for an error, SSO, or MFA prompt and complete it if required.",
                    "authenticated_url": page.url,
                }
            return {
                "ok": True,
                "message": "TestPulse submitted the dashboard credentials. Confirm the Fusion home page is loaded; complete any SSO/MFA prompt in Chrome if shown.",
                "authenticated_url": page.url,
            }

        # No standard password form was detected. This can mean an existing session,
        # an SSO landing page, or a custom identity flow. Do not pretend credentials
        # were used when no form was found.
        body_text = ""
        try:
            body_text = page.locator("body").inner_text(timeout=2000).lower()
        except Exception:
            pass
        if any(marker in body_text for marker in ["sign in", "sign-in", "username", "password", "oracle identity"]):
            return {
                "ok": True,
                "message": "Fusion opened, but TestPulse could not detect a standard username/password form to fill. This may be a custom SSO flow; complete sign-in manually in Chrome.",
                "authenticated_url": page.url,
            }
        return {
            "ok": True,
            "message": "Application opened. No standard password form was detected, so credentials were not entered. Confirm the application is already signed in or complete SSO/MFA in Chrome.",
            "authenticated_url": page.url,
        }


def _fusion_find_field(page, label):
    """Find a form field using accessible labels, placeholders, or common attributes."""
    label = str(label or "").strip()
    if not label:
        return None
    candidates = [
        page.get_by_label(re.compile(re.escape(label), re.I)),
        page.get_by_placeholder(re.compile(re.escape(label), re.I)),
        page.locator(f"input[aria-label*='{label}' i]"),
        page.locator(f"textarea[aria-label*='{label}' i]"),
        page.locator(f"input[name*='{label}' i]"),
        page.locator(f"input[id*='{label}' i]"),
        page.locator(f"textarea[name*='{label}' i]"),
        page.locator(f"textarea[id*='{label}' i]"),
        page.get_by_role("combobox", name=re.compile(re.escape(label), re.I)),
    ]
    for candidate in candidates:
        try:
            if candidate.count() and candidate.first.is_visible() and candidate.first.is_enabled():
                return candidate.first
        except Exception:
            continue
    # Oracle Fusion forms sometimes expose the field name as nearby visible text.
    try:
        label_text = page.get_by_text(re.compile(r"^\s*" + re.escape(label) + r"\s*[*:]?$", re.I)).first
        if label_text.count() and label_text.is_visible():
            parent = label_text.locator("xpath=..")
            for selector in ["input", "textarea", "select", "[role='combobox']"]:
                candidate = parent.locator(selector).first
                if candidate.count() and candidate.is_visible() and candidate.is_enabled():
                    return candidate
            parent2 = label_text.locator("xpath=../..")
            for selector in ["input", "textarea", "select", "[role='combobox']"]:
                candidate = parent2.locator(selector).first
                if candidate.count() and candidate.is_visible() and candidate.is_enabled():
                    return candidate
    except Exception:
        pass
    return None



def _fusion_fill_or_select(page, field_name, value, logs):
    target = _fusion_find_field(page, field_name)

    if target is None:
        raise RuntimeError(
            f"Could not identify an editable field or dropdown for "
            f"'{field_name}'. Check the visible form label and workbook FIELD_NAME."
        )

    value = str(value).strip()

    # Format dates as MM/DD/YY, for example 10/23/26.
    if "date" in field_name.lower():
        parsed_date = None

        for date_format in (
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%d/%m/%Y",
            "%d-%b-%Y",
            "%d-%B-%Y",
        ):
            try:
                parsed_date = datetime.strptime(value, date_format)
                break
            except ValueError:
                continue

        if parsed_date is not None:
            value = parsed_date.strftime("%m/%d/%y")
            _log(
                logs,
                f"Formatted {field_name} as {value} (MM/DD/YY)."
            )

    try:
        tag = target.evaluate("el => el.tagName.toLowerCase()")
    except Exception:
        tag = ""

    try:
        role = (target.get_attribute("role") or "").lower()
    except Exception:
        role = ""

    try:
        aria_has_popup = (
            target.get_attribute("aria-haspopup") or ""
        ).lower()
    except Exception:
        aria_has_popup = ""

    # ---------------------------------------------------------
    # SPECIAL HANDLING: Oracle Fusion Legal Employer dropdown
    # ---------------------------------------------------------
    
    # SPECIAL HANDLING: Oracle Fusion Legal Employer dropdown
    
    # Legal Employer: try entering the workbook value directly.
    if field_name.strip().casefold() == "legal employer":
        value = str(value).strip()

        try:
            target.click(timeout=3000)
            target.press("Control+A", timeout=2000)
            target.fill(value, timeout=3000)

            _log(
                logs,
                f"Entered '{value}' into the Legal Employer field."
            )

            # Check what the field actually displays.
            try:
                displayed_value = target.input_value(timeout=1500)
            except Exception:
                displayed_value = (
                    target.get_attribute("value") or ""
                )

            if value.casefold() not in displayed_value.casefold():
                raise RuntimeError(
                    f"Legal Employer did not retain '{value}'. "
                    f"Displayed value: '{displayed_value}'. "
                    "This field may require dropdown selection."
                )

            _log(
                logs,
                f"Verified Legal Employer text: '{displayed_value}'."
            )

        except Exception as exc:
            raise RuntimeError(
                f"Could not enter Legal Employer '{value}' manually: {exc}"
            ) from exc

    elif tag == "select":
        try:
            target.select_option(label=value)
        except Exception:
            target.select_option(value=value)

    elif tag == "select":
        try:
            target.select_option(label=value)
        except Exception:
            target.select_option(value=value)

    # ---------------------------------------------------------
    # STANDARD HTML SELECT
    # ---------------------------------------------------------
    elif tag == "select":
        try:
            target.select_option(label=value)
        except Exception:
            target.select_option(value=value)

    # ---------------------------------------------------------
    # GENERIC COMBOBOX
    # ---------------------------------------------------------
    elif role == "combobox" or "combobox" in aria_has_popup:
        target.click()
        page.wait_for_timeout(700)

        selected = False

        for frame in page.frames:
            try:
                candidates = [
                    frame.get_by_role(
                        "option", name=value, exact=True
                    ),
                    frame.get_by_text(value, exact=True),
                ]

                for locator in candidates:
                    for index in range(locator.count()):
                        option = locator.nth(index)

                        try:
                            if not option.is_visible():
                                continue

                            if (
                                option.inner_text(timeout=1000).strip().casefold()
                                != value.casefold()
                            ):
                                continue

                            option.click(timeout=3000)
                            page.wait_for_timeout(500)
                            selected = True
                            break

                        except Exception as exc:
                            _log(
                                logs,
                                f"Dropdown option click failed: {exc}"
                            )

                    if selected:
                        break

            except Exception:
                continue

            if selected:
                break

        if not selected:
            raise RuntimeError(
                f"Could not find and select dropdown option "
                f"'{value}' for '{field_name}'."
            )

    # ---------------------------------------------------------
    # TEXT INPUT / AUTOCOMPLETE
    # ---------------------------------------------------------
    else:
        try:
            target.fill(value)
        except Exception:
            target.click()
            target.press("Control+A")
            target.type(value)

        try:
            autocomplete = target.get_attribute("aria-autocomplete")
        except Exception:
            autocomplete = None

        if value and autocomplete:
            suggestion = page.get_by_role(
                "option",
                name=re.compile(re.escape(value), re.I)
            )

            try:
                if suggestion.count() and suggestion.first.is_visible():
                    suggestion.first.click()
            except Exception:
                pass

    _log(logs, f"Entered/selected value for field '{field_name}'.")





def _fusion_click(page, name, logs):
    name = str(name or "").strip()

    if not name:
        raise RuntimeError(
            "A click/select step has no FIELD_NAME or target text in Excel."
        )

    normalized = re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()

    # Special handling for Oracle Fusion's Navigator menu.
    is_navigation_menu = (
        "navigation menu" in normalized
        or "three lines" in normalized
        or normalized in {"hamburger menu", "main menu"}
    )

    if is_navigation_menu:
        try:
            navigator = page.locator("a#pt1\\:_UISmmLink")

            for index in range(navigator.count()):
                target = navigator.nth(index)
                if target.is_visible() and target.is_enabled():
                    target.click(timeout=4000)
                    page.wait_for_timeout(700)
                    _log(logs, "Clicked Oracle Fusion Navigator link.")
                    return
        except Exception as exc:
            _log(logs, f"Exact Navigator selector failed: {exc}")

        selectors = [
            "button[aria-label*='navigation' i]",
            "button[title*='navigation' i]",
            "[role='button'][aria-label*='navigation' i]",
            "button[aria-label*='menu' i]",
            "button[title*='menu' i]",
            "[role='button'][title*='menu' i]",
            "button[id*='menu' i]",
            "button[class*='menu' i]",
            "[role='button'][class*='menu' i]",
        ]

        for selector in selectors:
            try:
                matches = page.locator(selector)

                for index in range(matches.count()):
                    target = matches.nth(index)

                    if target.is_visible() and target.is_enabled():
                        target.click(timeout=2000)
                        page.wait_for_timeout(500)
                        _log(
                            logs,
                            f"Clicked navigation menu using selector: {selector}"
                        )
                        return
            except Exception as exc:
                _log(logs, f"Navigator selector attempt failed: {exc}")
    
    # Handle My Client Groups before generic locators time out.
    if normalized == "my client groups":
        page.wait_for_timeout(500)

        for frame in page.frames:
            target = frame.locator(
                '[id="pt1:_UISnvr:0:nvgcil_groupNode_workforce_management"]'
            )

            if target.count() == 0:
                target = frame.locator(
                    'a[title*="My Client Groups" i]'
                )

            for index in range(target.count()):
                item = target.nth(index)

                try:
                    if not item.is_visible():
                        continue

                    item.scroll_into_view_if_needed(timeout=2000)

                    # Fusion's sticky header can intercept normal clicks.
                    item.click(timeout=3000, force=True)
                    page.wait_for_timeout(800)

                    _log(
                        logs,
                        "Clicked My Client Groups using the "
                        "Fusion navigation element."
                    )
                    return

                except Exception as exc:
                    _log(
                        logs,
                        f"My Client Groups targeted click failed: {exc}"
                    )
    
    # Fast, Fusion-specific handling for My Client Groups.
    if normalized == "my client groups":
        page.wait_for_timeout(300)

        for frame in page.frames:
            # This exact ID was confirmed in your execution logs.
            expand_link = frame.locator(
                '[id="pt1:_UISnvr:0:nvgcil_groupNode_workforce_management"]'
            )

            try:
                if expand_link.count() and expand_link.first.is_visible():
                    item = expand_link.first

                    # Force avoids the sticky-header interception that
                    # caused repeated 3.5-second Playwright timeouts.
                    item.click(timeout=2000, force=True)
                    page.wait_for_timeout(400)

                    _log(
                        logs,
                        "Clicked My Client Groups using its exact "
                        "Fusion navigation element."
                    )
                    return

            except Exception as exc:
                _log(
                    logs,
                    f"Exact My Client Groups click failed: {exc}"
                )

        # Fallback: target the exact visible text directly.
        for frame in page.frames:
            try:
                matches = frame.get_by_text(
                    "My Client Groups", exact=True
                )

                for index in range(matches.count() - 1, -1, -1):
                    item = matches.nth(index)

                    if item.is_visible():
                        item.click(timeout=2000, force=True)
                        page.wait_for_timeout(400)

                        _log(
                            logs,
                            "Clicked My Client Groups using exact "
                            "Fusion text fallback."
                        )
                        return

            except Exception as exc:
                _log(
                    logs,
                    f"My Client Groups text fallback failed: {exc}"
                )

    # First, retain the existing standard click strategies.
    clicked = _click_named(page, [name], 3500)

    if clicked:
        _log(logs, f"Clicked '{name}'.")
        return

    for locator in [
        page.get_by_role(
            "button", name=re.compile(re.escape(name), re.I)
        ),
        page.get_by_role(
            "link", name=re.compile(re.escape(name), re.I)
        ),
        page.get_by_role(
            "menuitem", name=re.compile(re.escape(name), re.I)
        ),
        page.get_by_text(
            re.compile(r"^\s*" + re.escape(name) + r"\s*$", re.I)
        ),
    ]:
        try:
            for index in range(locator.count() - 1, -1, -1):
                target = locator.nth(index)

                if target.is_visible() and target.is_enabled():
                    target.click(timeout=3500)
                    _log(logs, f"Clicked '{name}'.")
                    return
        except Exception as exc:
            _log(logs, f"Standard click attempt failed for '{name}': {exc}")

    # Oracle Fusion-specific fallback for My Client Groups.
    if normalized == "my client groups":
        page.wait_for_timeout(1000)
        diagnostics = []

        for frame_index, frame in enumerate(page.frames):
            try:
                exact_text = frame.get_by_text(
                    "My Client Groups", exact=True
                )
                candidates = [
                    exact_text,
                    frame.locator(
                        "[aria-label='My Client Groups']"
                    ),
                    frame.locator(
                        "[title='My Client Groups']"
                    ),
                ]

                for locator in candidates:
                    for index in range(locator.count()):
                        target = locator.nth(index)

                        try:
                            if not target.is_visible():
                                continue

                            details = target.evaluate(
                                """el => ({
                                    tag: el.tagName,
                                    role: el.getAttribute('role'),
                                    text: (el.innerText || el.textContent || '').trim(),
                                    cls: typeof el.className === 'string'
                                        ? el.className : '',
                                    expanded: el.getAttribute('aria-expanded')
                                })"""
                            )

                            diagnostics.append(
                                f"Frame {frame_index}: {details}"
                            )

                            if not target.is_enabled():
                                continue

                            target.scroll_into_view_if_needed(
                                timeout=3000
                            )
                            target.click(timeout=4000)

                            page.wait_for_timeout(1000)
                            _log(
                                logs,
                                "Clicked 'My Client Groups' using "
                                f"Fusion fallback in frame {frame_index}."
                            )
                            return

                        except Exception as exc:
                            diagnostics.append(
                                f"Frame {frame_index}, candidate {index}: {exc}"
                            )

            except Exception as exc:
                diagnostics.append(
                    f"Frame {frame_index} inspection failed: {exc}"
                )

        for detail in diagnostics[:15]:
            _log(logs, f"My Client Groups diagnostic: {detail}")
    
    # Oracle Fusion fallback for navigation links such as
    # "Hire an Employee".
    if normalized == "hire an employee":
        page.wait_for_timeout(1000)
        diagnostics = []

        for frame_index, frame in enumerate(page.frames):
            try:
                # Try exact text and common Fusion link/button structures.
                candidates = [
                    frame.get_by_text(
                        "Hire an Employee", exact=True
                    ),
                    frame.get_by_role(
                        "link",
                        name="Hire an Employee",
                        exact=True
                    ),
                    frame.get_by_role(
                        "button",
                        name="Hire an Employee",
                        exact=True
                    ),
                    frame.locator(
                        "[title='Hire an Employee']"
                    ),
                    frame.locator(
                        "[aria-label='Hire an Employee']"
                    ),
                ]

                for locator in candidates:
                    for index in range(locator.count()):
                        target = locator.nth(index)

                        try:
                            if not target.is_visible():
                                continue

                            info = target.evaluate(
                                """el => ({
                                    tag: el.tagName,
                                    role: el.getAttribute('role'),
                                    text: (el.innerText || el.textContent || '').trim(),
                                    href: el.getAttribute('href'),
                                    cls: typeof el.className === 'string'
                                        ? el.className : ''
                                })"""
                            )
                            diagnostics.append(
                                f"Frame {frame_index}: {info}"
                            )

                            target.scroll_into_view_if_needed(
                                timeout=2000
                            )

                            # Normal click first. If an overlay intercepts
                            # it, continue to another candidate.
                            try:
                                target.click(timeout=2500)
                            except Exception:
                                continue

                            page.wait_for_timeout(1000)
                            _log(
                                logs,
                                "Clicked 'Hire an Employee' using "
                                f"Fusion fallback in frame {frame_index}."
                            )
                            return

                        except Exception as exc:
                            diagnostics.append(
                                f"Frame {frame_index}, candidate {index}: {exc}"
                            )

            except Exception as exc:
                diagnostics.append(
                    f"Frame {frame_index} inspection failed: {exc}"
                )

        for detail in diagnostics[:15]:
            _log(logs, f"Hire an Employee diagnostic: {detail}")

    # Targeted fallback for the Hire an Employee task on the New Person page.
    if name.strip().casefold() == "hire an employee":
        for frame in page.frames:
            try:
                task = frame.get_by_text("Hire an Employee", exact=True)

                for index in range(task.count()):
                    item = task.nth(index)

                    if not item.is_visible():
                        continue

                    item.scroll_into_view_if_needed(timeout=2000)

                    try:
                        item.click(timeout=3000)
                    except Exception:
                        # Fusion task links can have an overlapping container.
                        item.click(timeout=3000, force=True)

                    page.wait_for_timeout(1000)
                    _log(
                        logs,
                        "Clicked 'Hire an Employee' using the New Person task-list fallback."
                    )
                    return

            except Exception as exc:
                _log(
                    logs,
                    f"Hire an Employee task-list fallback: {exc}"
                )

    raise RuntimeError(f"Could not click or select '{name}'.")


def run_fusion_test_case(case_steps, settings):
    """Execute one Excel-defined step-based case in the current Chrome session.

    All test actions and test data are read from the workbook. A Save/Submit
    step is blocked unless allow_save=True in TestPulse.
    """
    started = time.monotonic()
    logs = []
    status = "FAIL"
    actual = ""
    error = ""
    if not case_steps:
        raise ValueError("No Excel steps were provided for this test case.")
    first = case_steps[0]
    test_id = str(first.get("TEST_CASE_ID", "UNKNOWN"))
    test_name = str(first.get("TEST_CASE_NAME", ""))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = EVIDENCE_DIR / f"{re.sub(r'[^A-Za-z0-9_-]+', '_', test_id)}_{stamp}"
    folder.mkdir(parents=True, exist_ok=True)
    shot = folder / "final_screenshot.png"
    video_path = folder / "test_run.mp4"
    ffmpeg_process = None
    page = None
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            if not browser.contexts:
                raise RuntimeError("No Chrome context found on port 9222. Open Chrome using the TestPulse remote-debugging command first.")
            context = browser.contexts[0]
            pages = context.pages
            page = pages[-1] if pages else context.new_page()
            page.set_default_timeout(10000)
            page.bring_to_front()
            app_url = str(settings.get("app_url", "")).strip()
            if app_url and page.url in ("about:blank", ""):
                page.goto(app_url, wait_until="domcontentloaded", timeout=60000)
            try:
                ffmpeg_process = subprocess.Popen(
                    [imageio_ffmpeg.get_ffmpeg_exe(), "-y", "-f", "gdigrab", "-framerate", "12", "-i", "desktop", "-c:v", "libx264", "-preset", "ultrafast", "-pix_fmt", "yuv420p", str(video_path)],
                    stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                time.sleep(0.7)
                _log(logs, "Desktop video recording started.")
            except Exception as exc:
                _log(logs, f"Video recording could not start: {exc}")
                ffmpeg_process = None

            for row in sorted(case_steps, key=lambda item: int(item.get("STEP_NUMBER", 0))):
                action = str(row.get("ACTION", "")).strip()
                field = str(row.get("FIELD_NAME", "")).strip()
                value = str(row.get("TEST_DATA", "")).strip()
                expected = str(row.get("EXPECTED_RESULT", "")).strip()
                low = action.lower()
                _log(logs, f"Step {row.get('STEP_NUMBER')}: {action}" + (f" — {field}" if field else ""))

                if low in {"open url", "navigate", "navigate to application", "open application"}:
                    target_url = app_url or value
                    if not target_url.startswith(("https://", "http://")):
                        raise RuntimeError("The workbook's Open URL step needs a valid application URL supplied in TestPulse.")
                    page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(600)
                elif "credential" in low or ("sign in" in low and "verify" not in low):
                    password_field = page.locator("input[type='password']").first
                    try:
                        login_visible = password_field.count() > 0 and password_field.is_visible()
                    except Exception:
                        login_visible = False
                    if login_visible:
                        username = str(settings.get("username", ""))
                        password = str(settings.get("password", ""))
                        user_field = None
                        for candidate in [page.locator("input[autocomplete='username']"), page.locator("input[type='email']"), page.locator("input[name*='user' i]"), page.locator("input[id*='user' i]"), page.locator("input[type='text']")]:
                            try:
                                if candidate.count() and candidate.first.is_visible() and candidate.first.is_enabled():
                                    user_field = candidate.first
                                    break
                            except Exception:
                                continue
                        if user_field is not None and username:
                            user_field.fill(username)
                        if password:
                            password_field.fill(password)
                        submit = page.get_by_role("button", name=re.compile(r"sign in|log in|login|continue|submit|next", re.I))
                        if submit.count() and submit.first.is_visible():
                            submit.first.click()
                        elif password:
                            password_field.press("Enter")
                        page.wait_for_timeout(1200)
                    else:
                        _log(logs, "No password field is visible; continuing with the current browser session (complete any SSO/MFA prompt manually).")
                elif low.startswith("verify"):
                    if "login" in field.lower() or "login" in expected.lower():
                        pwd = page.locator("input[type='password']")
                        try:
                            password_visible = pwd.count() > 0 and pwd.first.is_visible()
                        except Exception:
                            password_visible = False
                        if not password_visible and page.url not in ("about:blank", ""):
                            _log(logs, "Login status accepted: no visible password field remains; confirm any SSO/MFA step in Chrome if applicable.")
                        else:
                            raise RuntimeError("Login could not be verified because the sign-in form is still visible.")
                    else:
                        check_text = value or expected or field
                        if not check_text:
                            raise RuntimeError("A Verify step needs expected text in TEST_DATA or EXPECTED_RESULT.")
                        target = page.get_by_text(re.compile(re.escape(check_text), re.I), exact=False)
                        if not target.count() or not target.first.is_visible():
                            raise RuntimeError(f"Expected text/status was not verified: '{check_text}'.")
                        _log(logs, f"Verified expected text: {check_text}")
                elif low.startswith("wait"):
                    match = re.search(r"(\d+)", value or action)
                    page.wait_for_timeout(min(int(match.group(1)) if match else 1, 30) * 1000)
                elif any(token in low or token in field.lower() for token in ("save", "submit", "create")):
                    target_name = field or re.sub(r"^(click|select)\s*", "", action, flags=re.I).strip()
                    if not settings.get("allow_save", False):
                        status = "BLOCKED"
                        actual = f"Stopped before the '{target_name}' action because saving/creating records is disabled. Review the target environment and test data, then explicitly enable changes only when safe."
                        _log(logs, actual)
                        break
                    _fusion_click(page, target_name, logs)
                    page.wait_for_timeout(1200)
                elif low.startswith("click"):
                    _fusion_click(page, field or re.sub(r"^click\s*", "", action, flags=re.I), logs)
                    page.wait_for_timeout(300)
                elif low.startswith("select"):
                    if not field:
                        raise RuntimeError("A Select step needs FIELD_NAME in the workbook.")
                    if not value:
                        _fusion_click(page, field, logs)
                    else:
                        _fusion_fill_or_select(page, field, value, logs)
                    page.wait_for_timeout(250)
                elif "enter" in low or low.startswith("fill") or low.startswith("type"):
                    if not field:
                        raise RuntimeError("An Enter/Fill step needs FIELD_NAME in the workbook.")
                    if value.lower() in {"today's date", "today", "current date"}:
                        value = date.today().isoformat()
                    if field.lower() in {"username / password", "username/password", "login credentials"}:
                        # Credentials are provided securely in TestPulse, never from the workbook.
                        if page.locator("input[type='password']").count():
                            password_field = page.locator("input[type='password']").first
                            username_field = page.locator("input[autocomplete='username'], input[type='email'], input[name*='user' i], input[id*='user' i], input[type='text']").first
                            if username_field.count() and settings.get("username"):
                                username_field.fill(str(settings["username"]))
                            if settings.get("password"):
                                password_field.fill(str(settings["password"]))
                        else:
                            _log(logs, "Credentials are runtime-only; the browser does not currently show a password field.")
                    else:
                        _fusion_fill_or_select(page, field, value, logs)
                elif low.startswith("select/enter") or low.startswith("enter/select"):
                    if not field:
                        raise RuntimeError("A Select/Enter step needs FIELD_NAME in the workbook.")
                    _fusion_fill_or_select(page, field, value, logs)
                else:
                    raise RuntimeError(f"Unsupported Excel ACTION '{action}'. Use a supported action such as Click, Select, Enter text, Open URL, Wait, or Verify.")

            if status != "BLOCKED":
                status = "PASS"
                actual = "All workbook-defined steps completed."
                _log(logs, actual)
        except Exception as exc:
            status = "FAIL"
            error = str(exc)
            actual = "Automation could not complete the scenario."
            _log(logs, f"FAIL: {error}")
        finally:
            if page is not None:
                try:
                    page.bring_to_front()
                    page.screenshot(path=str(shot), full_page=True, timeout=15000)
                    _log(logs, f"Screenshot saved: {shot}")
                except Exception as exc:
                    _log(logs, f"Screenshot failed: {exc}")
            if ffmpeg_process is not None:
                try:
                    ffmpeg_process.stdin.write(b"q\n")
                    ffmpeg_process.stdin.flush()
                    ffmpeg_process.wait(timeout=15)
                except Exception as exc:
                    try:
                        ffmpeg_process.kill()
                    except Exception:
                        pass
                    _log(logs, f"Video recording stop issue: {exc}")

    return {
        "test_case_id": test_id,
        "test_case_name": test_name,
        "status": status,
        "duration_seconds": round(time.monotonic() - started, 2),
        "actual_result": actual,
        "error_message": error,
        "screenshot_path": str(shot) if shot.is_file() else "",
        "video_path": str(video_path) if video_path.is_file() and video_path.stat().st_size > 0 else "",
        "logs": logs,
        "executed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

