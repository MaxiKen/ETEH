# Dirty Dataset Issue Manifest

Every file below is a deliberately corrupted copy of its Clean counterpart.
Use them for data-cleaning, QA and exploratory-analysis exercises.

## healthcare_records_dirty.csv
- Mixed-case patient names (e.g. 'bObby jacksOn', 'LESILE TERRY').
- ~4% missing values injected in blood_type, medical_condition, insurance_provider, doctor.
- ~15% of billing_amount values formatted as currency text ('$12,345.67').
- ~30% of date_of_admission rewritten in mixed formats (DD/MM/YYYY, 'Mon DD, YYYY', YYYY/MM/DD).
- Inconsistent gender labels (male/Male/MALE).
- ~2.5% exact duplicate rows appended.
- ~0.5% billing_amount values turned negative (sign-error).

## salary_data_dirty.csv
- Inconsistent education_level spellings/casing (bachelors, Bachelor's , PHD, masters...).
- ~15% of salary values formatted as currency text ('$95,000.00').
- Mixed gender labels (M, F, Male, female).
- Random leading/trailing whitespace in job_title.
- ~3% missing values in age and years_of_experience.
- ~2% exact duplicate rows appended.

## superstore_sales_dirty.csv
- ~30% of ship_date rewritten in mixed formats (DD-Mon-YYYY, MM/DD/YYYY, YYYYMMDD).
- ~4% missing values in postal_code and region.
- Random whitespace added to city names.
- ~15% of sales formatted as currency text.
- ~1% profit values turned negative (returned orders).
- ~2% exact duplicate rows appended.

## student_performance_dirty.csv
- ~1% exam_score values pushed above 100 (impossible).
- Inconsistent gender labels.
- Inconsistent Yes/No labels in extracurricular_activities.
- ~4% missing in hours_studied, sleep_hours, previous_scores.
- ~2% duplicate rows appended.

## diabetes_prediction_dirty.csv
- Inconsistent smoking_history labels (No Info, no info, NOT CURRENT, Ever...).
- ~4% missing in bmi and hbA1c_level.
- Mixed gender casing.
- ~1.5% duplicate rows appended.

## hotel_reservations_dirty.csv
- Inconsistent booking_status labels (canceled/Cancelled/not_canceled...).
- ~4% missing in type_of_meal_plan and avg_price_per_room.
- Whitespace added to market_segment_type.
- ~0.5% negative lead_time values (impossible).
- ~2% duplicate rows appended.

## student_depression_dirty.csv
- Inconsistent Yes/No labels in suicidal_thoughts.
- Mixed-case dietary_habits.
- ~4% missing in cgpa, sleep_duration, financial_stress.
- ~2% duplicate rows appended.

## social_media_addiction_dirty.csv
- Mixed gender casing.
- ~4% missing in usage hours and mental health score.
- Mixed-case platform names.
- ~2% duplicate rows appended.

