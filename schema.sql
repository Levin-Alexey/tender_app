CREATE TABLE IF NOT EXISTS clients (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    name text NOT NULL UNIQUE,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS clients_name_key ON clients(name);

CREATE TABLE IF NOT EXISTS calculations (
    id uuid PRIMARY KEY,
    client_id uuid NOT NULL REFERENCES clients(id),
    status text NOT NULL DEFAULT 'pending',
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS documents (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    task_id uuid NOT NULL,
    calculation_id uuid NOT NULL REFERENCES calculations(id) ON DELETE CASCADE,
    document_url text NOT NULL,
    document_name text NOT NULL,
    source_object_key text,
    status text NOT NULL DEFAULT 'pending',
    extracted_data jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE documents
    ADD COLUMN IF NOT EXISTS source_object_key text;

ALTER TABLE documents
    DROP CONSTRAINT IF EXISTS documents_task_id_key;

DROP INDEX IF EXISTS documents_task_id_key;

CREATE INDEX IF NOT EXISTS documents_task_id_idx ON documents(task_id);
CREATE INDEX IF NOT EXISTS documents_calculation_id_idx
    ON documents(calculation_id);
