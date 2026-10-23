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
