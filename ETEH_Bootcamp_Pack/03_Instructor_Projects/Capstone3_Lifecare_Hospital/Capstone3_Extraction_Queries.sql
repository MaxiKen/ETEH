-- Lifecare Hospital - Capstone 3 extraction queries (SQLite; run against 01_Section_Examples/Week3_Data_Files/lifecare_hospital.db)
-- Every query below was executed during the build; row counts are recorded on sheet 2 of the solution workbook.

-- Q1 Control totals by year (agree these with Health Records and Finance before any analysis)
SELECT substr(admission_date, 1, 4)              AS year,
       COUNT(*)                                   AS admissions,
       ROUND(SUM(billing_amount), 2)              AS billing_ngn,
       SUM(CASE WHEN readmit_30d = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
       ROUND(AVG(length_of_stay), 2)              AS alos
FROM admissions
GROUP BY year
ORDER BY year;

-- Q2 FY2023 admissions extract for Excel - NO names or patient IDs (data minimisation)
SELECT a.admission_id,
       a.admission_date,
       substr(a.admission_date, 1, 7)             AS month,
       'Q' || ((CAST(substr(a.admission_date, 6, 2) AS INTEGER) + 2) / 3) AS quarter,
       w.ward,
       a.department,
       d.grade                                    AS doctor_grade,
       a.admission_type,
       a.age,
       CASE WHEN a.age < 18 THEN '0-17' WHEN a.age < 40 THEN '18-39'
            WHEN a.age < 60 THEN '40-59' WHEN a.age < 75 THEN '60-74' ELSE '75+' END AS age_band,
       a.medical_condition,
       a.length_of_stay,
       a.hmo,
       a.billing_amount,
       a.readmit_30d
FROM admissions a
JOIN wards   w ON w.ward_id   = a.ward_id
JOIN doctors d ON d.doctor_id = a.doctor_id
WHERE a.admission_date BETWEEN '2023-01-01' AND '2023-12-31'
ORDER BY a.admission_date, a.admission_id;

-- Q3 Daily bed census (whole history, for capacity trends)
SELECT census_date,
       substr(census_date, 1, 4) AS year,
       substr(census_date, 1, 7) AS month,
       patients_in_house,
       staffed_beds,
       occupancy_pct
FROM bed_census
ORDER BY census_date;

-- Q4 Cross-check for the Excel safety table
SELECT department,
       COUNT(*) AS discharges,
       SUM(CASE WHEN readmit_30d = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
       ROUND(AVG(CASE WHEN readmit_30d = 'Yes' THEN 1.0 ELSE 0 END), 4) AS rate
FROM admissions
WHERE admission_date LIKE '2023%'
GROUP BY department
ORDER BY rate DESC;

-- Q5 Patients admitted 3+ times in FY2023 (counts only - the named list stays in the clinical system)
SELECT admissions_2023, COUNT(*) AS patients
FROM (SELECT patient_id, COUNT(*) AS admissions_2023
      FROM admissions
      WHERE admission_date LIKE '2023%'
      GROUP BY patient_id
      HAVING COUNT(*) >= 3)
GROUP BY admissions_2023
ORDER BY admissions_2023;
