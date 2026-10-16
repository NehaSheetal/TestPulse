"""Optional PostgreSQL persistence for NovaTest. No connection is attempted without DATABASE_URL."""
import os, json
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR=Path(__file__).resolve().parent
load_dotenv(BASE_DIR/".env")

def database_enabled():
    return bool(os.getenv("DATABASE_URL","").strip())

def _connect():
    if not database_enabled():
        raise RuntimeError("DATABASE_URL is not configured. Copy .env.example to .env and set DATABASE_URL.")
    try:
        import psycopg
        return psycopg.connect(os.environ["DATABASE_URL"],connect_timeout=5)
    except ImportError as exc:
        raise RuntimeError("PostgreSQL driver is missing. Run: python -m pip install 'psycopg[binary]>=3.2'") from exc

def init_db():
    if not database_enabled(): return
    sql=Path(__file__).with_name("schema.sql").read_text(encoding="utf-8")
    with _connect() as conn:
        with conn.cursor() as cur: cur.execute(sql)

def test_connection():
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT current_database(), current_user, version()")
            row=cur.fetchone()
    return f"Connected to database '{row[0]}' as '{row[1]}'. PostgreSQL responded successfully."

def save_test_run(results,workbook_name,app_url):
    if not results:return
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("""INSERT INTO test_runs(workbook_name, app_url, total_tests, passed_tests, failed_tests, blocked_tests)
                           VALUES (%s,%s,%s,%s,%s,%s) RETURNING run_id""",
                        (workbook_name,app_url,len(results),sum(r["status"]=="PASS" for r in results),sum(r["status"]=="FAIL" for r in results),sum(r["status"]=="BLOCKED" for r in results)))
            run_id=cur.fetchone()[0]
            for r in results:
                cur.execute("""INSERT INTO test_results(run_id,test_case_id,test_case_name,status,duration_seconds,actual_result,error_message,screenshot_path,video_path,logs,executed_at)
                               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s::jsonb,%s)""",
                            (run_id,r.get("test_case_id"),r.get("test_case_name"),r.get("status"),r.get("duration_seconds"),r.get("actual_result"),r.get("error_message"),r.get("screenshot_path"),r.get("video_path"),json.dumps(r.get("logs",[])),r.get("executed_at")))
        conn.commit()

def load_history(limit=100):
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("""SELECT r.run_id, r.started_at, r.workbook_name, r.total_tests, r.passed_tests, r.failed_tests, r.blocked_tests,
                           t.test_case_id, t.test_case_name, t.status, t.duration_seconds, t.actual_result, t.error_message
                           FROM test_runs r LEFT JOIN test_results t ON t.run_id=r.run_id
                           ORDER BY r.run_id DESC, t.result_id LIMIT %s""",(int(limit),))
            rows=cur.fetchall()
    cols=["Run ID","Started at","Workbook","Total","Passed","Failed","Blocked","Test ID","Test name","Status","Duration (s)","Actual result","Error"]
    return [dict(zip(cols,row)) for row in rows]
