-- =====================================================================
-- Lifecare General Hospital - inpatient database (SQLite)
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
    ward_id      TEXT PRIMARY KEY,
    ward         TEXT NOT NULL,
    first_room   INTEGER,
    last_room    INTEGER,
    staffed_beds INTEGER NOT NULL
);
CREATE TABLE doctors (
    doctor_id    TEXT PRIMARY KEY,
    doctor_name  TEXT NOT NULL,
    department   TEXT NOT NULL,
    grade        TEXT
);
CREATE TABLE patients (
    patient_id      TEXT PRIMARY KEY,
    patient_name    TEXT NOT NULL,
    gender          TEXT,
    blood_type      TEXT,
    first_admission DATE,
    admissions      INTEGER
);
CREATE TABLE admissions (
    admission_id      TEXT PRIMARY KEY,
    patient_id        TEXT NOT NULL REFERENCES patients(patient_id),
    age               INTEGER,
    medical_condition TEXT,
    department        TEXT,
    admission_type    TEXT CHECK (admission_type IN ('Emergency','Urgent','Elective')),
    admission_date    DATE NOT NULL,
    discharge_date    DATE,
    length_of_stay    INTEGER,
    ward_id           TEXT REFERENCES wards(ward_id),
    room_number       INTEGER,
    doctor_id         TEXT REFERENCES doctors(doctor_id),
    hmo               TEXT,
    billing_amount    REAL,
    medication        TEXT,
    test_result       TEXT,
    readmit_30d       TEXT
);
CREATE TABLE bed_census (
    census_date       DATE PRIMARY KEY,
    patients_in_house INTEGER,
    staffed_beds      INTEGER,
    occupancy_pct     REAL
);
CREATE INDEX ix_adm_date ON admissions(admission_date);
CREATE INDEX ix_adm_patient ON admissions(patient_id);
CREATE INDEX ix_adm_dept ON admissions(department);

CREATE VIEW v_admissions AS
SELECT a.*, p.patient_name, p.gender, p.blood_type, d.doctor_name, w.ward,
       strftime('%Y', a.admission_date) AS adm_year, strftime('%Y-%m', a.admission_date) AS adm_month
FROM admissions a
JOIN patients p ON p.patient_id = a.patient_id
JOIN doctors  d ON d.doctor_id  = a.doctor_id
JOIN wards    w ON w.ward_id    = a.ward_id;
