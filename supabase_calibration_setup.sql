-- ============================================
-- CALIBRATION CERTIFICATE GENERATOR
-- Supabase Database Setup
-- Generated: 2026-03-31
--
-- Run in Supabase SQL Editor:
-- Dashboard > SQL Editor > New Query > Paste > Run
-- ============================================

-- Certificate generation jobs (history)
CREATE TABLE certificate_jobs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    status VARCHAR(30) NOT NULL DEFAULT 'pending'
        CHECK (status IN ('pending', 'processing', 'completed', 'failed')),

    -- Equipment info (extracted from uploaded Excel)
    equipment_model VARCHAR(200),
    serial_number VARCHAR(100),
    manufacturer VARCHAR(200),
    calibration_date VARCHAR(50),
    calibration_due VARCHAR(50),
    calibration_tech VARCHAR(200),
    customer VARCHAR(200),

    -- Input files metadata
    input_files JSONB DEFAULT '[]'::jsonb,
    -- e.g. [{"name": "B2140 Acc.xlsx", "type": "acc_xlsx", "path": "uploads/..."}]

    -- Output file in Supabase Storage
    output_file_path TEXT,
    output_file_name VARCHAR(255),

    -- Error tracking
    error_message TEXT,

    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    completed_at TIMESTAMPTZ
);

CREATE INDEX idx_jobs_status ON certificate_jobs(status);
CREATE INDEX idx_jobs_created ON certificate_jobs(created_at DESC);
CREATE INDEX idx_jobs_serial ON certificate_jobs(serial_number);
CREATE INDEX idx_jobs_model ON certificate_jobs(equipment_model);

-- Enable RLS
ALTER TABLE certificate_jobs ENABLE ROW LEVEL SECURITY;

-- Open access (no auth for now)
CREATE POLICY "Allow all access" ON certificate_jobs
    FOR ALL USING (true) WITH CHECK (true);

-- ============================================
-- STORAGE BUCKET
-- ============================================
-- After running this SQL, create ONE storage bucket
-- in Supabase Dashboard > Storage:
--
--   Bucket name: calibration-files
--   Public: YES
--
-- Files will be organized as:
--   calibration-files/uploads/{job_id}/   -- input files
--   calibration-files/certificates/{job_id}/  -- generated PDFs
-- ============================================
