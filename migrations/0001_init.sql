-- One row per uploaded day (re-uploading the same date replaces it).
CREATE TABLE reports (
  id           TEXT PRIMARY KEY,
  report_date  TEXT NOT NULL,            -- YYYY-MM-DD
  period_start INTEGER,                  -- first sale (ms, naive local time)
  period_end   INTEGER,
  file_name    TEXT NOT NULL,
  file_size    INTEGER,
  file_key     TEXT,                     -- R2 key of the original upload
  lines_key    TEXT,                     -- R2 key of parsed sale lines (gzip JSON)
  status       TEXT NOT NULL DEFAULT 'uploading',  -- uploading | ready
  rules        TEXT NOT NULL,            -- JSON: rules used for the stored summary
  totals       TEXT NOT NULL,            -- JSON: KPI totals
  uploaded_by  TEXT,
  uploaded_at  INTEGER NOT NULL
);
CREATE INDEX reports_by_date ON reports (report_date DESC);

-- Per-branch summary for each day: drives history, trends and repeat offenders.
CREATE TABLE store_daily (
  report_id     TEXT NOT NULL REFERENCES reports (id) ON DELETE CASCADE,
  report_date   TEXT NOT NULL,
  store         TEXT NOT NULL,
  status        TEXT NOT NULL,           -- High | Medium | Low | Clean
  invoices      INTEGER, mugs INTEGER, sales REAL,
  direct_lines  INTEGER, direct_excess INTEGER,
  runs          INTEGER, runs_high INTEGER, run_excess INTEGER,
  pct_over      REAL,
  PRIMARY KEY (report_id, store)
);
CREATE INDEX store_daily_by_store ON store_daily (store, report_date);
CREATE INDEX store_daily_by_date ON store_daily (report_date);

-- Review decisions on individual flags (split runs and direct violations).
CREATE TABLE reviews (
  flag_key     TEXT PRIMARY KEY,
  report_id    TEXT NOT NULL REFERENCES reports (id) ON DELETE CASCADE,
  kind         TEXT NOT NULL,            -- run | direct
  store        TEXT NOT NULL,
  status       TEXT NOT NULL,            -- open | genuine | violation | false_positive
  note         TEXT,
  reviewed_by  TEXT,
  reviewed_at  INTEGER NOT NULL
);
CREATE INDEX reviews_by_report ON reviews (report_id);
