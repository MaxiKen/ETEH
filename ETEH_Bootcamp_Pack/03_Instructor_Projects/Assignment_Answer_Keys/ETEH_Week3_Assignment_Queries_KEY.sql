-- ETEH Week 3 Assignment - ANSWER KEY SQL (instructor only)

-- Q2.1  How many admissions were there in 2023?   -> expected: 12626
SELECT COUNT(*) AS admissions
FROM admissions
WHERE strftime('%Y', admit_date) = '2023';

-- Q2.2  What was the average length of stay in 2023 (2 dp)?   -> expected: 5.02
SELECT ROUND(AVG(los_days), 2) AS alos
FROM admissions
WHERE strftime('%Y', admit_date) = '2023';

-- Q2.3  Which department had the most admissions in 2023?   -> expected: Internal Medicine
SELECT department, COUNT(*) AS n
FROM admissions
WHERE strftime('%Y', admit_date) = '2023'
GROUP BY department
ORDER BY n DESC
LIMIT 1;

-- Q2.4  Total 2023 billing for Emergency admissions (₦, nearest naira).   -> expected: 6170775526.0
SELECT ROUND(SUM(bill_amount_ngn), 0) AS emergency_billing
FROM admissions
WHERE strftime('%Y', admit_date) = '2023'
  AND admission_route = 'Emergency';

-- Q2.5  JOIN: which ward (by NAME) had the highest 30-day readmission rate in 2023?   -> expected: Baobab Ward
SELECT w.ward, ROUND(AVG(a.readmitted_30d = 'Yes'), 4) AS readmit_rate
FROM admissions a
JOIN wards w ON w.ward_code = a.ward_code
WHERE strftime('%Y', a.admit_date) = '2023'
GROUP BY w.ward
ORDER BY readmit_rate DESC
LIMIT 1;

-- Q2.6  JOIN: which doctor (by NAME) handled the most admissions in 2023?   -> expected: Dr Onyinye I. Lawal
SELECT d.consultant_name, COUNT(*) AS n
FROM admissions a
JOIN doctors d ON d.consultant_id = a.consultant_id
WHERE strftime('%Y', a.admit_date) = '2023'
GROUP BY d.consultant_name
ORDER BY n DESC
LIMIT 1;

-- Q2.7  HAVING: how many doctors handled MORE than 400 admissions in 2023?   -> expected: 24
SELECT COUNT(*) AS busy_doctors FROM (
  SELECT consultant_id
  FROM admissions
  WHERE strftime('%Y', admit_date) = '2023'
  GROUP BY consultant_id
  HAVING COUNT(*) > 400
);

-- Q2.8  How many patients have been admitted three or more times (all years)?   -> expected: 593
SELECT COUNT(*) AS frequent_patients FROM (
  SELECT patient_no
  FROM admissions
  GROUP BY patient_no
  HAVING COUNT(*) >= 3
);

-- Q2.9  Average daily bed occupancy in 2023 from bed_census (as a %, 1 dp).   -> expected: 86.9
SELECT ROUND(AVG(occupancy_pct) * 100, 1) AS avg_occupancy_pct
FROM bed_census
WHERE strftime('%Y', census_date) = '2023';

-- Q2.10  SUBQUERY: how many 2023 admissions were billed ABOVE the 2023 average bill?   -> expected: 4067
SELECT COUNT(*) AS above_avg
FROM admissions
WHERE strftime('%Y', admit_date) = '2023'
  AND bill_amount_ngn > (SELECT AVG(bill_amount_ngn)
                        FROM admissions
                        WHERE strftime('%Y', admit_date) = '2023');
