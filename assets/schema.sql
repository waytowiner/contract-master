PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS runs (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  started_at    TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
  finished_at   TEXT,
  mode          TEXT NOT NULL DEFAULT 'standard',
  config_json   TEXT,
  total_files   INTEGER DEFAULT 0,
  succeeded     INTEGER DEFAULT 0,
  failed        INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS files (
  id                INTEGER PRIMARY KEY AUTOINCREMENT,
  path              TEXT NOT NULL UNIQUE,
  file_name         TEXT NOT NULL,
  ext               TEXT,
  size_bytes        INTEGER,
  sha256            TEXT,
  page_count        INTEGER,
  status            TEXT NOT NULL DEFAULT 'pending',
  error_message     TEXT,
  first_seen_run_id INTEGER REFERENCES runs(id),
  updated_at        TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_files_sha256 ON files(sha256);
CREATE INDEX IF NOT EXISTS idx_files_status ON files(status);

CREATE TABLE IF NOT EXISTS ocr_pages (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  file_id       INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
  page_no       INTEGER NOT NULL,
  text          TEXT,
  avg_confidence REAL,
  blocks_json   TEXT,
  created_at    TEXT DEFAULT (datetime('now', 'localtime')),
  UNIQUE(file_id, page_no)
);

CREATE TABLE IF NOT EXISTS contracts (
  id                INTEGER PRIMARY KEY AUTOINCREMENT,
  group_name        TEXT,
  source_root       TEXT NOT NULL DEFAULT '',
  file_id           INTEGER REFERENCES files(id),
  run_id            INTEGER REFERENCES runs(id),
  contract_no       TEXT,
  title             TEXT,
  contract_type     TEXT,
  party_a           TEXT,
  party_b           TEXT,
  sign_place        TEXT,
  sign_date         TEXT,
  start_date        TEXT,
  end_date          TEXT,
  amount            REAL,
  amount_in_words   TEXT,
  tax_rate          REAL,
  currency          TEXT DEFAULT 'CNY',
  payment_terms     TEXT,
  archive_location  TEXT,
  status            TEXT NOT NULL DEFAULT 'unverified',
  created_at        TEXT DEFAULT (datetime('now', 'localtime')),
  updated_at        TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_contracts_no ON contracts(contract_no);
CREATE INDEX IF NOT EXISTS idx_contracts_root_group ON contracts(source_root, group_name);
CREATE INDEX IF NOT EXISTS idx_contracts_party_a ON contracts(party_a);
CREATE INDEX IF NOT EXISTS idx_contracts_party_b ON contracts(party_b);
CREATE INDEX IF NOT EXISTS idx_contracts_end_date ON contracts(end_date);
CREATE INDEX IF NOT EXISTS idx_contracts_amount ON contracts(amount);
CREATE INDEX IF NOT EXISTS idx_contracts_status ON contracts(status);

CREATE TABLE IF NOT EXISTS field_sources (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  contract_id   INTEGER NOT NULL REFERENCES contracts(id) ON DELETE CASCADE,
  field_name    TEXT NOT NULL,
  value         TEXT,
  source_value  TEXT,
  file_id       INTEGER REFERENCES files(id),
  page_no       INTEGER,
  source_page_no INTEGER,
  snippet       TEXT,
  bbox_json     TEXT,
  confidence    REAL,
  method        TEXT,
  needs_review  INTEGER NOT NULL DEFAULT 0,
  verified_by   TEXT,
  verified_at   TEXT,
  UNIQUE(contract_id, field_name)
);

CREATE INDEX IF NOT EXISTS idx_sources_review ON field_sources(needs_review);

CREATE TABLE IF NOT EXISTS review_items (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  review_key      TEXT UNIQUE,
  contract_id     INTEGER NOT NULL REFERENCES contracts(id) ON DELETE CASCADE,
  field_name      TEXT,
  current_value   TEXT,
  suggested_value TEXT,
  reason          TEXT NOT NULL,
  severity        TEXT NOT NULL DEFAULT 'normal',
  status          TEXT NOT NULL DEFAULT 'pending',
  resolved_value  TEXT,
  resolved_by     TEXT,
  resolved_at     TEXT
);

CREATE INDEX IF NOT EXISTS idx_review_status ON review_items(status);

CREATE TABLE IF NOT EXISTS issues (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  contract_id INTEGER REFERENCES contracts(id) ON DELETE CASCADE,
  file_id     INTEGER REFERENCES files(id),
  issue_type  TEXT NOT NULL,
  severity    TEXT NOT NULL DEFAULT 'normal',
  detail      TEXT,
  status      TEXT NOT NULL DEFAULT 'open',
  created_at  TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE VIRTUAL TABLE IF NOT EXISTS contract_fts USING fts5(
  title, party_a, party_b, full_text,
  tokenize='trigram'
);

CREATE TRIGGER IF NOT EXISTS trg_contracts_updated
AFTER UPDATE ON contracts
BEGIN
  UPDATE contracts SET updated_at = datetime('now', 'localtime') WHERE id = NEW.id;
END;
