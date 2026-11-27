-- ETEH Week 3 Assignment - ANSWER KEY SQL (instructor only)

-- Q2.1  How many admissions were there in 2023?   -> expected: 10918
SELECT COUNT(*) AS admissions
FROM admissions
WHERE strftime('%Y', admission_date) = '2023';

-- Q2.2  What was the average length of stay in 2023 (2 dp)?   -> expected: 15.49
SELECT ROUND(AVG(length_of_stay), 2) AS alos
FROM admissions
WHERE strftime('%Y', admission_date) = '2023';

-- Q2.3  Which department had the most admissions in 2023?   -> expected: Endocrinology
SELECT department, COUNT(*) AS n
FROM admissions
WHERE strftime('%Y', admission_date) = '2023'
GROUP BY department
ORDER BY n DESC
LIMIT 1;

-- Q2.4  Total 2023 billing for Emergency admissions (₦, nearest naira).   -> expected: 4703210503.0
SELECT ROUND(SUM(billing_amount), 0) AS emergency_billing
FROM admissions
WHERE strftime('%Y', admission_date) = '2023'
  AND admission_type = 'Emergency';

-- Q2.5  JOIN: which ward (by NAME) had the highest 30-day readmission rate in 2023?   -> expected: Sapphire Ward
SELECT w.ward, ROUND(AVG(a.readmit_30d = 'Yes'), 4) AS readmit_rate
FROM admissions a
JOIN wards w ON w.ward_id = a.ward_id
WHERE strftime('%Y', a.admission_date) = '2023'
GROUP BY w.ward
ORDER BY readmit_rate DESC
LIMIT 1;

-- Q2.6  JOIN: which doctor (by NAME) handled the most admissions in 2023?   -> expected: Dr Hadiza Uzor
SELECT d.doctor_name, COUNT(*) AS n
FROM admissions a
JOIN doctors d ON d.doctor_id = a.doctor_id
WHERE strftime('%Y', a.admission_date) = '2023'
GROUP BY d.doctor_name
ORDER BY n DESC
LIMIT 1;

-- Q2.7  HAVING: how many doctors handled MORE than 400 admissions in 2023?   -> expected: 8
SELECT COUNT(*) AS busy_doctors FROM (
  SELECT doctor_id
  FROM admissions
  WHERE strftime('%Y', admission_date) = '2023'
  GROUP BY doctor_id
  HAVING COUNT(*) > 400
);

-- Q2.8  How many patients have been admitted three or more times (all years)?   -> expected: 259
SELECT COUNT(*) AS frequent_patients FROM (
  SELECT patient_id
  FROM admissions
  GROUP BY patient_id
  HAVING COUNT(*) >= 3
);

-- Q2.9  Average daily bed occupancy in 2023 from bed_census (as a %, 1 dp).   -> expected: 89.5
SELECT ROUND(AVG(occupancy_pct) * 100, 1) AS avg_occupancy_pct
FROM bed_census
WHERE strftime('%Y', census_date) = '2023';

-- Q2.10  SUBQUERY: how many 2023 admissions were billed ABOVE the 2023 average bill?   -> expected: 5521
SELECT COUNT(*) AS above_avg
FROM admissions
WHERE strftime('%Y', admission_date) = '2023'
  AND billing_amount > (SELECT AVG(billing_amount)
                        FROM admissions
                        WHERE strftime('%Y', admission_date) = '2023');
