CREATE TABLE IF NOT EXISTS test_runs (
    run_id BIGSERIAL PRIMARY KEY,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    workbook_name TEXT NOT NULL DEFAULT '',
    app_url TEXT NOT NULL DEFAULT '',
    total_tests INTEGER NOT NULL DEFAULT 0,
    passed_tests INTEGER NOT NULL DEFAULT 0,
    failed_tests INTEGER NOT NULL DEFAULT 0,
    blocked_tests INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS test_results (
    result_id BIGSERIAL PRIMARY KEY,
    run_id BIGINT NOT NULL REFERENCES test_runs(run_id) ON DELETE CASCADE,
    test_case_id TEXT NOT NULL,
    test_case_name TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL CHECK (status IN ('PASS','FAIL','BLOCKED')),
    duration_seconds NUMERIC(10,2) NOT NULL DEFAULT 0,
    actual_result TEXT NOT NULL DEFAULT '',
    error_message TEXT NOT NULL DEFAULT '',
    screenshot_path TEXT NOT NULL DEFAULT '',
    video_path TEXT NOT NULL DEFAULT '',
    logs JSONB NOT NULL DEFAULT '[]'::jsonb,
    executed_at TIMESTAMP NULL
);
CREATE INDEX IF NOT EXISTS idx_test_results_run_id ON test_results(run_id);
CREATE INDEX IF NOT EXISTS idx_test_runs_started_at ON test_runs(started_at DESC);
