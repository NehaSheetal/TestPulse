# NovaTest Quality Lab — Windows setup guide

NovaTest is a local Streamlit dashboard that uses Playwright to test your existing Oracle APEX Timesheet app. PostgreSQL is an **optional history database**. It does not replace or modify the APEX database.

## What each part does

- **Streamlit**: the dashboard you open in your browser.
- **Python**: runs the dashboard and test logic.
- **Playwright**: opens Chrome/Chromium and performs browser actions.
- **Oracle APEX**: the application being tested; keep its data and business rules there.
- **PostgreSQL**: optionally stores test-run summaries and individual results so history survives restarting Streamlit.
- **Git/GitHub**: synchronizes source code between computers. Git does not automatically synchronize a PostgreSQL database or browser login.

## 1. Install the tools

1. Install **Python 3.12 (64-bit)** from https://www.python.org/downloads/windows/
   - During setup, tick **Add python.exe to PATH** if shown.
   - Finish the installation.
2. Install **Visual Studio Code** from https://code.visualstudio.com/
3. Install **Google Chrome** from https://www.google.com/chrome/
4. Optional for PostgreSQL history: install PostgreSQL for Windows from https://www.postgresql.org/download/windows/. The official installer usually offers pgAdmin too. Keep the default port `5432`; choose a strong password for the database administrator and store it somewhere safe. Do not send it to anyone.
5. Optional cross-computer sync: install Git from https://git-scm.com/download/win and create a private repository on https://github.com/.

## 2. Create the project folder

1. Extract the downloaded ZIP to a folder you can find, for example `C:\Users\YOUR_USER\Documents\NovaTest`.
2. Open VS Code.
3. Choose **File → Open Folder…** and select the extracted `NovaTest` folder.
4. In VS Code, choose **Terminal → New Terminal**. The terminal should open inside the project folder.
5. Run `py --version`. If Python 3.12 is installed, it should show `Python 3.12.x`.

## 3. Install project packages

Double-click `setup_windows.bat` in File Explorer, or run `setup_windows.bat` from the VS Code terminal using `.\setup_windows.bat`.

It creates a project-specific `.venv`, installs the packages in `requirements.txt`, and installs Playwright's Chromium browser. This project uses the versions/ranges in `requirements.txt`:
- streamlit>=1.39
- pandas>=2.2
- openpyxl>=3.1
- playwright>=1.50
- psycopg[binary]>=3.2
- python-dotenv>=1.0

The project is pinned to Python 3.12 for the setup script. You can later select `.venv` in VS Code using **Ctrl+Shift+P → Python: Select Interpreter**.

## 4. Run NovaTest

Double-click `run_app.bat`. Streamlit should start and show a local URL, usually `http://localhost:8501`. Open that URL in your browser. Keep the terminal window open while using the app; closing it stops the app.

## 5. Sign in to Oracle APEX

1. In NovaTest, open **Run tests**.
2. Check the application URL. The default is your Timesheet Management URL.
3. Click **Open Chrome & sign in**.
4. In the separate Chrome window, sign in normally. NovaTest does not collect your password.
5. If APEX lands on a blank home page, use its navigation menu and open **Employee Timesheet**.
6. The runner saves a local Playwright storage-state file. Treat it like a password: do not share it or commit it to GitHub. If authentication expires, clear the saved session and sign in again.

**Important about the earlier infinite spinner:** the runner uses `domcontentloaded`, not `networkidle`, because APEX pages may continue background network activity. It periodically checks for Employee Timesheet and tries the navigation menu. If clicking **Sign In** itself spins forever in a normal Chrome window opened directly to the APEX URL, this is likely an APEX/session/network problem; this local test tool cannot repair the server-side login. First verify that the application can be opened and signed into directly in normal Chrome.

## 6. Configure PostgreSQL (optional)

If you don't have time to set up PostgreSQL right now, leave `DATABASE_URL=` empty. NovaTest still runs, but the result history remains in Streamlit session memory and won't survive an app restart. Configure PostgreSQL when ready.

1. Install PostgreSQL and pgAdmin from the official PostgreSQL download page.
2. Open **pgAdmin** and connect to your local server. Use the password you chose during installation.
3. Right-click **Login/Group Roles → Create → Login/Group Role**. Name it `novatest_user`; under **Definition**, set a strong password; under **Privileges**, allow login. Do not grant superuser.
4. Right-click **Databases → Create → Database…**. Name it `novatest_db`; set owner to `novatest_user` if available.
5. If you created the database with a different owner, use Query Tool connected to `novatest_db` and run:
   `GRANT CONNECT, CREATE ON DATABASE novatest_db TO novatest_user;`
6. Copy `.env.example` to `.env` (the setup script creates it). Open `.env` in VS Code and set:
   `DATABASE_URL=postgresql://novatest_user:URL_ENCODED_PASSWORD@localhost:5432/novatest_db`
   Replace the example password. If it contains characters such as `@`, `:`, `/`, `#`, or `%`, URL-encode the password or choose a strong password without URL-reserved characters. Never commit `.env`.
7. Save `.env`, stop Streamlit with **Ctrl+C** in the terminal, then run `run_app.bat` again.
8. Open **History & setup → Test PostgreSQL connection**. On the first successful connection, NovaTest creates the two tables from `schema.sql`.
9. Run a test, then open **History & setup** to see saved history.

Database concepts:
- A **database** is a named container for related data.
- A **table** is like a spreadsheet.
- A **row** is one record; a **column** is one field.
- A **primary key** uniquely identifies a record.
- A **foreign key** links a result to the run it belongs to.

## 7. Use it on a second computer with GitHub

### First computer: one-time repository setup

1. Create a **private** GitHub repository named `novatest-quality-lab`. Do not initialize it with a README if you plan to push this folder as-is.
2. Open the VS Code terminal in this project folder and run:

```powershell
git --version
git init
git add .
git status
```

3. Confirm `.env`, `apex_storage_state.json`, `apex_test_session.json`, `.venv`, and `test_evidence` files are not staged. If any secrets are staged, stop and unstage them before committing.
4. Continue:

```powershell
git commit -m "Initial NovaTest project"
git branch -M main
git remote add origin https://github.com/YOUR_GITHUB_USERNAME/novatest-quality-lab.git
git push -u origin main
```

Replace `YOUR_GITHUB_USERNAME`. GitHub may open a browser for authentication; do not paste a password or token into source files.

### Second computer

1. Install Python 3.12, VS Code, Chrome, and Git.
2. Open VS Code terminal in a folder where you want the project.
3. Run:

```powershell
git clone https://github.com/YOUR_GITHUB_USERNAME/novatest-quality-lab.git
cd novatest-quality-lab
```

4. Double-click `setup_windows.bat`.
5. Copy `.env.example` to `.env` and set the second computer's local PostgreSQL connection, or use the same hosted PostgreSQL URL only if you have configured a secure hosted database.
6. Run `run_app.bat` and sign in to APEX again. Browser authentication state is intentionally not shared via Git.

### Everyday workflow

Before starting work on a computer:
`git pull`

After making and testing changes:
`git add .`
`git commit -m "Describe your change"`
`git push`

Git synchronizes code, not your local database contents. For the same history on both computers, use one securely hosted PostgreSQL database, or export/import the database separately. For a beginner and a deadline, start with local PostgreSQL only if time permits; later, move history to a managed PostgreSQL provider and restrict network access. Do not expose a database to the public internet with an unrestricted firewall rule.

## 8. Safety notes

- Saving changes in APEX is **off by default**. Test with a test account and only enable saving when you intentionally want to change timesheet records.
- A test case may be marked PASS when a UI action completes; weekend/holiday approval outcomes still require checking the APEX approval state. Review evidence before treating it as proof of business-rule correctness.
- Screenshots and video can contain employee or project data. Keep `test_evidence/` private and do not commit it.
- A private GitHub repository is not a secret vault. Never commit real credentials, `.env`, browser state, cookies, or tokens.
- This project has been assembled from your supplied source files and workbook. Browser automation selectors are best-effort and may need adjustment against the live APEX UI. A generated project ZIP cannot prove that login or the tests work against your live application until you run them.
