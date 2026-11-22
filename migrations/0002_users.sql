-- App-managed accounts. New sign-ups are 'pending' until an admin approves them.
CREATE TABLE users (
  id            TEXT PRIMARY KEY,
  email         TEXT NOT NULL UNIQUE,          -- stored lower-case
  name          TEXT NOT NULL,
  pw_hash       TEXT NOT NULL,                 -- PBKDF2-SHA256, base64
  pw_salt       TEXT NOT NULL,                 -- base64
  pw_iter       INTEGER NOT NULL,
  role          TEXT NOT NULL DEFAULT 'viewer',  -- admin | viewer
  area_manager  TEXT,                          -- viewer scope (name as in the mapping)
  status        TEXT NOT NULL DEFAULT 'pending', -- pending | active | disabled
  failed_logins INTEGER NOT NULL DEFAULT 0,
  locked_until  INTEGER NOT NULL DEFAULT 0,
  created_at    INTEGER NOT NULL,
  last_login    INTEGER
);

CREATE TABLE sessions (
  token_hash  TEXT PRIMARY KEY,               -- SHA-256 of the cookie token
  user_id     TEXT NOT NULL REFERENCES users (id) ON DELETE CASCADE,
  created_at  INTEGER NOT NULL,
  expires_at  INTEGER NOT NULL
);
CREATE INDEX sessions_by_user ON sessions (user_id);

-- Branch hierarchy captured at upload, so area-manager views and trends filter in SQL.
ALTER TABLE store_daily ADD COLUMN am TEXT;
ALTER TABLE store_daily ADD COLUMN rom TEXT;
ALTER TABLE store_daily ADD COLUMN city TEXT;
ALTER TABLE store_daily ADD COLUMN zone TEXT;
CREATE INDEX store_daily_by_am ON store_daily (am, report_date);
