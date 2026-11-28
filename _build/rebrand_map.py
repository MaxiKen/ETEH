"""Edition 3 dataset replacement: three Kaggle-derived datasets are replaced by original data from new organisations.

    Lifecare General Hospital (+ bed census, diabetes screening)  ->  Mercyland Specialist Hospital, Abuja
    Crestview University (+ wellbeing and digital surveys)       ->  Kingsway University, Ilorin
    Eko Suites Hotel, Victoria Island                            ->  Atlantic Pearl Hotel, Lekki

The data itself is generated from scratch by neworgs.py (no third-party source rows). This map renames every
organisation name, file name, column name and code value in the builders, word-exact and case-sensitive.
Run `python3 _build/rebrand.py` to apply it; applying it twice is harmless (no new name is also an old name).
"""

ORGS = [
    # longest first
    ("Lifecare General Hospital", "Mercyland Specialist Hospital"),
    ("Lifecare Hospital", "Mercyland Hospital"),
    ("Lifecare_Admissions_2019_2024", "Mercyland_Inpatient_Episodes_2019_2024"),
    ("Lifecare_Admissions_2023_DIRTY", "Mercyland_Episodes_2023_DIRTY"),
    ("Lifecare_Admissions_2023_pipe", "Mercyland_Episodes_2023_pipe"),
    ("Lifecare_Daily_Bed_Census", "Mercyland_Daily_Bed_Census"),
    ("Lifecare_Diabetes_Screening", "Mercyland_Diabetes_Screening"),
    ("Lifecare_Doctors", "Mercyland_Consultants"),
    ("Lifecare_Patients", "Mercyland_Patients"),
    ("Lifecare_Wards", "Mercyland_Wards"),
    ("Lifecare_EMR_Export_2023", "Mercyland_EMR_Export_2023"),
    ("Lifecare_Screening_Clinic_Extract_RAW", "Mercyland_Screening_Clinic_Extract_RAW"),
    ("Project3_Lifecare_Diabetes_Screening", "Project3_Mercyland_Diabetes_Screening"),
    ("Capstone3_Lifecare_Hospital", "Capstone3_Mercyland_Hospital"),
    ("Capstone3_Lifecare_SOLUTION", "Capstone3_Mercyland_SOLUTION"),
    ("lifecare_hospital_schema", "mercyland_hospital_schema"),
    ("lifecare_hospital", "mercyland_hospital"),
    ("Lifecare", "Mercyland"),
    ("lifecare", "mercyland"),
    ("LIFECARE", "MERCYLAND"),
    ("LGH-P", "MSH-P"),
    ("LGH", "MSH"),
    ("LDS-", "MSS-"),
    ("tLcDirty", "tMsDirty"),
    ("Crestview University", "Kingsway University"),
    ("Crestview_Student_Performance", "Kingsway_Student_Results"),
    ("Crestview_Wellbeing_Survey_SENSITIVE", "Kingsway_Student_Welfare_Survey_SENSITIVE"),
    ("Crestview_Digital_Wellbeing_Survey", "Kingsway_Digital_Habits_Survey"),
    ("Crestview_Registry_Export_RAW", "Kingsway_Registry_Export_RAW"),
    ("Capstone2_Crestview_University", "Capstone2_Kingsway_University"),
    ("Capstone2_Crestview_SOLUTION", "Capstone2_Kingsway_SOLUTION"),
    ("Crestview Wellbeing", "Kingsway Welfare"),
    ("Crestview", "Kingsway"),
    ("crestview", "kingsway"),
    ("CRESTVIEW", "KINGSWAY"),
    ("CU/", "KU/"),
    ("DWS-", "KDH-"),
    ("Eko Suites Hotel, Victoria Island", "Atlantic Pearl Hotel, Lekki"),
    ("Eko Suites Hotel", "Atlantic Pearl Hotel"),
    ("Eko Suites", "Atlantic Pearl"),
    ("EkoSuites_Hotel_Bookings_2023_2024", "AtlanticPearl_Reservations_2023_2024"),
    ("EkoSuites_PMS_Export_RAW", "AtlanticPearl_PMS_Export_RAW"),
    ("Project2_EkoSuites_Hotel", "Project2_AtlanticPearl_Hotel"),
    ("EkoSuites", "AtlanticPearl"),
    ("Victoria Island", "Lekki"),
    ("EKO-", "APH-"),
]

COLUMNS = [
    # hospital
    ("Admission_ID", "Episode_ID"),
    ("Patient_ID", "Patient_No"),
    ("Blood_Type", "Blood_Group"),
    ("Medical_Condition", "Diagnosis_Group"),
    ("Admission_Type", "Admission_Route"),
    ("Admission_Date", "Admit_Date"),
    ("Length_of_Stay", "LOS_Days"),
    ("Ward_ID", "Ward_Code"),
    ("Room_Number", "Bed_No"),
    ("First_Room", "First_Bed"),
    ("Last_Room", "Last_Bed"),
    ("Doctor_ID", "Consultant_ID"),
    ("Doctor_Name", "Consultant_Name"),
    ("Billing_Amount", "Bill_Amount_NGN"),
    ("Test_Result", "Lab_Result"),
    ("Readmit_30d", "Readmitted_30d"),
    ("Patients_In_House", "Inpatients"),
    ("Staffed_Beds", "Open_Beds"),
    ("First_Admission", "First_Episode"),
    # screening
    ("Screening_ID", "Screen_ID"),
    ("Heart_Disease", "Heart_Condition"),
    ("Smoking_History", "Smoking_Status"),
    ("HbA1c_Level", "HbA1c_Pct"),
    ("Blood_Glucose_mg_dL", "Glucose_mg_dL"),
    # university
    ("Student_ID", "Matric_No"),
    ("Hours_Studied_Per_Week", "Study_Hours_Wk"),
    ("Attendance_Pct", "Lecture_Attendance_Pct"),
    ("Sleep_Hours", "Sleep_Hrs_Night"),
    ("Previous_Score", "Last_Sem_Score"),
    ("Tutoring_Sessions", "Tutorials_Attended"),
    ("Physical_Activity_Hrs", "Exercise_Hrs_Wk"),
    ("Parental_Involvement", "Guardian_Support"),
    ("Access_to_Resources", "Library_Access"),
    ("Extracurricular", "Club_Member"),
    ("Motivation_Level", "Motivation"),
    ("Internet_Access", "Home_Internet"),
    ("Family_Income", "Household_Income"),
    ("Teacher_Quality", "Lecturer_Rating"),
    ("School_Type", "Secondary_School"),
    ("Peer_Influence", "Peer_Effect"),
    ("Learning_Disability", "Learning_Support"),
    ("Parental_Education", "Guardian_Education"),
    ("Distance_from_Home", "Commute_Band"),
    ("Exam_Score", "Final_Score"),
    ("Academic_Pressure_1to5", "Study_Pressure_1to5"),
    ("CGPA_5pt", "CGPA"),
    ("Study_Satisfaction_1to5", "Course_Satisfaction_1to5"),
    ("Sleep_Duration", "Sleep_Band"),
    ("Dietary_Habits", "Diet_Quality"),
    ("Study_Hours_Per_Day", "Daily_Study_Hrs"),
    ("Financial_Stress_1to5", "Money_Worry_1to5"),
    ("Family_History_Mental_Illness", "Family_MH_History"),
    ("Suicidal_Thoughts_Ever", "Self_Harm_Thoughts"),
    ("Depression_Screen_Positive", "PHQ9_Positive"),
    ("Avg_Daily_Usage_Hours", "Daily_Screen_Hrs"),
    ("Most_Used_Platform", "Top_App"),
    ("Affects_Academic_Performance", "Hurts_Studies"),
    ("Sleep_Hours_Per_Night", "Sleep_Hrs"),
    ("Mental_Health_Score", "Wellbeing_Score"),
    ("Conflicts_Over_Social_Media", "Phone_Conflicts"),
    ("Addicted_Score", "Dependence_Score"),
    # hotel
    ("Booking_ID", "Reservation_ID"),
    ("Booking_Date", "Booked_On"),
    ("Arrival_Date", "Check_In"),
    ("Arrival_Year", "Check_In_Year"),
    ("Total_Nights", "Nights"),
    ("Meal_Plan", "Board_Basis"),
    ("Parking_Required", "Car_Park"),
    ("Room_Type", "Room_Category"),
    ("Lead_Time_Days", "Days_Ahead"),
    ("Market_Segment", "Sales_Channel"),
    ("Repeat_Guest", "Returning_Guest"),
    ("Previous_Cancellations", "Prior_Cancellations"),
    ("Previous_Stays", "Prior_Stays"),
    ("Avg_Room_Rate_NGN", "Room_Rate_NGN"),
    ("Special_Requests", "Guest_Requests"),
    ("Booking_Status", "Reservation_Status"),
    ("Booking_Value_NGN", "Reservation_Value_NGN"),
]

VALUES = [
    # hospital departments, diagnoses, wards, payers, medicines
    ("Orthopaedics & Rheumatology", "Obstetrics & Gynaecology"),
    ("Endocrinology", "Internal Medicine"),
    ("Oncology", "Surgery"),
    ("Respiratory", "Paediatrics"),
    ("Cancer", "Trauma"),
    ("Obesity", "Malaria"),
    ("Asthma", "Pneumonia"),
    ("Arthritis", "Childbirth"),
    ("Emerald Ward", "Iroko Ward"), ("Sapphire Ward", "Mahogany Ward"), ("Ruby Ward", "Obeche Ward"), ("Topaz Ward", "Baobab Ward"),
    ("EMERALD WARD", "IROKO WARD"), ("SAPPHIRE WARD", "MAHOGANY WARD"), ("RUBY WARD", "OBECHE WARD"), ("TOPAZ WARD", "BAOBAB WARD"),
    ("Saphire Ward", "Mahogony Ward"), ("Ruby  Ward", "Obeche  Ward"), ("Topaz ward", "Baobab ward"),
    ("Emerald", "Iroko"), ("Sapphire", "Mahogany"), ("Ruby", "Obeche"), ("Topaz", "Baobab"),
    ("emerald", "iroko"), ("sapphire", "mahogany"), ("saphire", "mahogony"), ("ruby", "obeche"), ("topaz", "baobab"),
    ("AXA Mansard Health", "Avon HMO"), ("Reliance HMO", "Clearline HMO"), ("Leadway Health", "Total Health Trust"),
    ("Hygeia HMO", "Hallmark HMO"),
    ("Lipitor", "Metformin"), ("Aspirin", "Amlodipine"), ("Penicillin", "Ceftriaxone"), ("Ibuprofen", "Artemether-Lumefantrine"),
    # hotel status
    ("HONOURED", "STAYED"), ("Honoured", "Stayed"), ("honoured", "stayed"),
]


def pairs():
    """All (old, new) pairs plus snake_case variants of column names (SQL / Power Query keys)."""
    out = list(ORGS) + list(COLUMNS) + list(VALUES)
    for o, n in COLUMNS:
        lo, ln = o.lower(), n.lower()
        if lo != o and "_" in o:
            out.append((lo, ln))
    return out
