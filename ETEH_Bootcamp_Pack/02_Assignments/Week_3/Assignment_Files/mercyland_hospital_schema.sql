-- =====================================================================
-- Mercyland Specialist Hospital - inpatient database (SQLite)
-- ETEH Bootcamp - Week 3 (SQL / ODBC) and Capstone 3
-- =====================================================================
PRAGMA foreign_keys = ON;
DROP VIEW IF EXISTS v_admissions;
DROP TABLE IF EXISTS bed_census;
DROP TABLE IF EXISTS admissions;
DROP TABLE IF EXISTS patients;
DROP TABLE IF EXISTS doctors;
DROP TABLE IF EXISTS wards;

CREATE TABLE wards (
    ward_code    TEXT PRIMARY KEY,
    ward         TEXT NOT NULL,
    first_bed    INTEGER,
    last_bed     INTEGER,
    open_beds    INTEGER NOT NULL
);
CREATE TABLE doctors (
    consultant_id   TEXT PRIMARY KEY,
    consultant_name TEXT NOT NULL,
    department      TEXT NOT NULL,
    grade           TEXT
);
CREATE TABLE patients (
    patient_no      TEXT PRIMARY KEY,
    patient_name    TEXT NOT NULL,
    gender          TEXT,
    blood_group     TEXT,
    first_episode   DATE,
    admissions      INTEGER
);
CREATE TABLE admissions (
    episode_id        TEXT PRIMARY KEY,
    patient_no        TEXT NOT NULL REFERENCES patients(patient_no),
    age               INTEGER,
    diagnosis_group   TEXT,
    department        TEXT,
    admission_route   TEXT CHECK (admission_route IN ('Emergency','Urgent','Elective')),
    admit_date        DATE NOT NULL,
    discharge_date    DATE,
    los_days          INTEGER,
    ward_code         TEXT REFERENCES wards(ward_code),
    bed_no            INTEGER,
    consultant_id     TEXT REFERENCES doctors(consultant_id),
    hmo               TEXT,
    bill_amount_ngn   REAL,
    medication        TEXT,
    lab_result        TEXT,
    readmitted_30d    TEXT
);
CREATE TABLE bed_census (
    census_date       DATE PRIMARY KEY,
    inpatients        INTEGER,
    open_beds         INTEGER,
    occupancy_pct     REAL
);
CREATE INDEX ix_adm_date ON admissions(admit_date);
CREATE INDEX ix_adm_patient ON admissions(patient_no);
CREATE INDEX ix_adm_dept ON admissions(department);

CREATE VIEW v_admissions AS
SELECT a.*, p.patient_name, p.gender, p.blood_group, d.consultant_name, w.ward,
       strftime('%Y', a.admit_date) AS adm_year, strftime('%Y-%m', a.admit_date) AS adm_month
FROM admissions a
JOIN patients p ON p.patient_no = a.patient_no
JOIN doctors  d ON d.consultant_id = a.consultant_id
JOIN wards    w ON w.ward_code  = a.ward_code;
