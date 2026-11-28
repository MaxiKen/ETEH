"""Original teaching data for three fictional organisations (edition 3). No third-party rows are used.

    Mercyland Specialist Hospital, Abuja  - inpatient episodes 2019-2024, consultants, wards, daily bed census,
                                            and the community diabetes screening programme
    Kingsway University, Ilorin           - student results 2024/25, student welfare survey (SENSITIVE), digital habits
    Atlantic Pearl Hotel, Lekki           - room reservations, July 2023 - December 2024

Every generator is seeded, so the files are identical on every run. Patterns are built in on purpose so the lessons
have something real to find: a second-quarter bed surge, department differences in length of stay and readmission,
exam-score drivers with a small high-achiever cluster, a lead-time cancellation gradient and a Simpson's paradox.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

SEED = 20261127

# ======================================================================================================
# Mercyland Specialist Hospital
# ======================================================================================================
WARDS = [("W1", "Iroko Ward", 101, 160, 60), ("W2", "Mahogany Ward", 201, 250, 50),
         ("W3", "Obeche Ward", 301, 350, 50), ("W4", "Baobab Ward", 401, 440, 40)]
TOTAL_BEDS = sum(w[4] for w in WARDS)          # 200 staffed beds
DX = {  # diagnosis group -> (department, share, mean LOS days, main medicine, home ward)
    "Malaria": ("Internal Medicine", 0.22, 3.6, "Artemether-Lumefantrine", "W1"),
    "Diabetes": ("Internal Medicine", 0.12, 6.4, "Metformin", "W1"),
    "Hypertension": ("Cardiology", 0.15, 5.6, "Amlodipine", "W4"),
    "Pneumonia": ("Paediatrics", 0.15, 4.4, "Ceftriaxone", "W3"),
    "Childbirth": ("Obstetrics & Gynaecology", 0.21, 2.5, "Oxytocin", "W3"),
    "Trauma": ("Surgery", 0.15, 7.2, "Paracetamol", "W2"),
}
PAYERS = ["NHIA (National Health Insurance)", "Avon HMO", "Clearline HMO", "Total Health Trust", "Hallmark HMO"]
PAYER_P = [0.32, 0.20, 0.18, 0.16, 0.14]
BLOOD = ["O+", "A+", "B+", "AB+", "O-", "A-", "B-", "AB-"]
BLOOD_P = [0.52, 0.21, 0.19, 0.03, 0.025, 0.01, 0.01, 0.005]


def _daily_rate(days):
    """Expected admissions per day: growth, a second-quarter surge (early rains), a December dip, weekday pattern, 2020 dip."""
    t = (days - days[0]).days / 365.25
    m = days.month
    season = np.select([m.isin([4, 5, 6]), m.isin([7, 8, 9]), m == 12, m.isin([1, 2])], [1.11, 1.03, 0.91, 0.96], 1.0)
    wd = np.array([1.12, 1.04, 1.02, 1.00, 0.98, 0.90, 0.84])[days.dayofweek]
    covid = np.where((days >= "2020-04-01") & (days <= "2020-07-31"), 0.72, 1.0)
    return 30.2 * (1.03 ** t) * season * wd * covid


def gen_hospital(namer):
    """namer(key, gender) -> unique Nigerian name. Returns episodes, patients, consultants, wards, census (new column names)."""
    r = np.random.default_rng(SEED + 101)
    days = pd.date_range("2019-04-08", "2024-05-07", freq="D")
    n_day = r.poisson(_daily_rate(days))
    adm = np.repeat(days.values, n_day)
    n = len(adm)
    adm = pd.to_datetime(adm)
    mon = adm.month.values
    keys = list(DX)
    base_p = np.array([DX[k][1] for k in keys])
    # malaria rises in the rainy season (April - October)
    rainy = np.isin(mon, [4, 5, 6, 7, 8, 9, 10])
    p_rain = base_p.copy(); p_rain[0] *= 1.45; p_rain /= p_rain.sum()
    p_dry = base_p.copy(); p_dry[0] *= 0.65; p_dry /= p_dry.sum()
    u = r.random(n)
    dx = np.where(rainy, np.searchsorted(np.cumsum(p_rain), u), np.searchsorted(np.cumsum(p_dry), u))
    dx = np.clip(dx, 0, len(keys) - 1)
    dxn = np.array(keys)[dx]
    dept = np.array([DX[k][0] for k in keys])[dx]
    # sex and age by diagnosis
    male_p = np.array([0.48, 0.47, 0.46, 0.53, 0.0, 0.70])[dx]
    gender = np.where(r.random(n) < male_p, "Male", "Female")
    age = np.select(
        [dxn == "Pneumonia", dxn == "Childbirth", dxn == "Hypertension", dxn == "Diabetes", dxn == "Trauma"],
        [np.clip(r.gamma(1.6, 2.2, n), 0, 14), np.clip(r.normal(29, 5.5, n), 16, 46), np.clip(r.normal(59, 12, n), 30, 95),
         np.clip(r.normal(55, 12, n), 25, 95), np.clip(r.normal(34, 13, n), 16, 85)],
        np.clip(r.normal(36, 17, n), 15, 90)).round().astype(int)
    route_p = {"Malaria": [.50, .44, .06], "Diabetes": [.30, .40, .30], "Hypertension": [.36, .40, .24],
               "Pneumonia": [.56, .40, .04], "Childbirth": [.28, .30, .42], "Trauma": [.74, .21, .05]}
    ur = r.random(n)
    route = np.empty(n, dtype=object)
    for k, p in route_p.items():
        m_ = dxn == k
        c = np.cumsum(p)
        route[m_] = np.array(["Emergency", "Urgent", "Elective"])[np.searchsorted(c, ur[m_])]
    # length of stay: log-normal around the diagnosis mean, longer for emergencies and older patients
    mu = np.array([DX[k][2] for k in keys])[dx] * np.where(route == "Emergency", 1.15, 1.0) * (1 + np.clip(age - 60, 0, None) / 120)
    los = np.maximum(1, np.round(r.lognormal(np.log(mu) - 0.18, 0.6))).astype(int)
    los = np.minimum(los, 60)
    dis = adm + pd.to_timedelta(los, unit="D")
    # patients: chronic conditions come back; some within 30 days
    pid = np.empty(n, dtype=object)
    last_by_dx = {k: [] for k in keys}
    serial = 0
    back_p = {"Malaria": .05, "Diabetes": .20, "Hypertension": .17, "Pneumonia": .05, "Childbirth": .03, "Trauma": .04}
    order = np.argsort(adm.values, kind="stable")
    meta = {}
    for i in order:
        k = dxn[i]
        pool = last_by_dx[k]
        if pool and r.random() < back_p[k]:
            j = pool[int(r.integers(max(0, len(pool) - 900), len(pool)))]
            pid[i] = pid[j]
            gender[i] = gender[j]
            age[i] = max(age[i], age[j])
        else:
            serial += 1
            pid[i] = f"MSH-P{serial:06d}"
        pool.append(i)
        meta.setdefault(pid[i], gender[i])
    # readmission within 30 days of this discharge (observed in the data) + follow-up registry
    df = pd.DataFrame({"pid": pid, "adm": adm, "dis": dis})
    df = df.sort_values(["pid", "adm"])
    nxt = df.groupby("pid")["adm"].shift(-1)
    obs = ((nxt - df.dis).dt.days.between(0, 30)).reindex(range(n)).fillna(False).values
    z = (-3.55 + 0.012 * age + 0.45 * (route == "Emergency") + 0.06 * los
         + pd.Series(dxn).map({"Diabetes": 0.40, "Hypertension": 0.15, "Trauma": 0.10, "Pneumonia": 0.1, "Malaria": -0.1,
                               "Childbirth": -0.9}).values)
    readm = np.where(obs | (r.random(n) < 1 / (1 + np.exp(-z))), "Yes", "No")
    # consultants
    cons, roster, cid = [], {}, 1
    for dname in sorted(set(v[0] for v in DX.values())):
        roster[dname] = []
        for j in range(6 if dname in ("Internal Medicine", "Obstetrics & Gynaecology") else 4):
            g = "Female" if (cid * 7) % 3 else "Male"
            nm = "Dr " + namer(f"cons{cid}", g)
            roster[dname].append((f"CON{cid:03d}", nm))
            cons.append(dict(Consultant_ID=f"CON{cid:03d}", Consultant_Name=nm, Department=dname,
                             Grade="Consultant" if j < 3 else "Senior Registrar"))
            cid += 1
    pick = [roster[d][int(x * len(roster[d]))] for d, x in zip(dept, r.random(n))]
    # wards: home ward, 15% overflow to Baobab / Iroko
    home = np.array([DX[k][4] for k in keys])[dx]
    over = r.random(n) < 0.15
    ward = np.where(over, np.where(home == "W4", "W1", "W4"), home)
    wmap = {w[0]: w for w in WARDS}
    bed = np.array([wmap[w][2] + int(x * wmap[w][4]) for w, x in zip(ward, r.random(n))])
    # bill: daily rate by department + procedure fee, 19% a year medical inflation, payer mix, noise
    day_rate = pd.Series(dept).map({"Internal Medicine": 38000, "Cardiology": 52000, "Paediatrics": 34000,
                                    "Obstetrics & Gynaecology": 41000, "Surgery": 64000}).values
    fee = np.select([dxn == "Childbirth", dxn == "Trauma"],
                    [np.where(r.random(n) < 0.36, 520000, 160000), np.where(r.random(n) < 0.55, 680000, 210000)], 90000)
    infl = 1.19 ** ((adm - pd.Timestamp("2019-01-01")).days / 365.25)
    bill = np.round((fee + day_rate * los) * infl * r.lognormal(0, 0.18, n), 2)
    med = np.where(r.random(n) < 0.82, np.array([DX[k][3] for k in keys])[dx], "Paracetamol")
    lab = np.where(r.random(n) < np.select([dxn == "Childbirth", dxn == "Malaria"], [0.2, 0.75], 0.55), "Abnormal",
                   np.where(r.random(n) < 0.88, "Normal", "Inconclusive"))
    names = {p: namer(p, g) for p, g in meta.items()}
    ep = pd.DataFrame({"Admit_Date": adm, "Patient_No": pid, "Age": age, "Gender": gender,
                       "Blood_Group": None, "Diagnosis_Group": dxn, "Department": dept, "Admission_Route": route,
                       "Discharge_Date": dis, "LOS_Days": los, "Ward_Code": ward, "Bed_No": bed,
                       "Consultant_ID": [p_[0] for p_ in pick], "Consultant_Name": [p_[1] for p_ in pick],
                       "HMO": np.array(PAYERS)[np.searchsorted(np.cumsum(PAYER_P), r.random(n)).clip(0, 4)],
                       "Bill_Amount_NGN": bill, "Medication": med, "Lab_Result": lab, "Readmitted_30d": readm})
    bg = {p: BLOOD[min(int(np.searchsorted(np.cumsum(BLOOD_P), x)), 7)] for p, x in zip(meta, r.random(len(meta)))}
    ep["Blood_Group"] = ep.Patient_No.map(bg)
    ep["Gender"] = ep.Patient_No.map(meta)
    ep["Patient_Name"] = ep.Patient_No.map(names)
    ep["Ward"] = ep.Ward_Code.map({w[0]: w[1] for w in WARDS})
    # the published extract covers admissions from 8 May 2019 (the month before is warm-up for the census)
    warm = ep[ep.Admit_Date < "2019-05-08"]
    ep = ep[ep.Admit_Date >= "2019-05-08"].sort_values(["Admit_Date", "Patient_No"]).reset_index(drop=True)
    ep.insert(0, "Episode_ID", [f"ADM{d.year}-{i:05d}" for i, d in enumerate(ep.Admit_Date, 1)])
    ep = ep[["Episode_ID", "Patient_No", "Patient_Name", "Age", "Gender", "Blood_Group", "Diagnosis_Group", "Department",
             "Admission_Route", "Admit_Date", "Discharge_Date", "LOS_Days", "Ward_Code", "Ward", "Bed_No", "Consultant_ID",
             "Consultant_Name", "HMO", "Bill_Amount_NGN", "Medication", "Lab_Result", "Readmitted_30d"]]
    allspells = pd.concat([warm[["Admit_Date", "Discharge_Date"]], ep[["Admit_Date", "Discharge_Date"]]])
    cday = pd.date_range("2019-06-07", "2024-05-07", freq="D")
    a_ct = allspells.groupby("Admit_Date").size().reindex(pd.date_range("2019-04-01", "2024-08-31"), fill_value=0)
    d_ct = allspells.groupby("Discharge_Date").size().reindex(a_ct.index, fill_value=0)
    inh = (a_ct - d_ct).cumsum().reindex(cday).astype(int)
    census = pd.DataFrame({"Census_Date": cday.date, "Inpatients": inh.values, "Open_Beds": TOTAL_BEDS})
    census["Occupancy_Pct"] = (census.Inpatients / census.Open_Beds).round(4)
    ep["Admit_Date"] = ep.Admit_Date.dt.date
    ep["Discharge_Date"] = pd.to_datetime(ep.Discharge_Date).dt.date
    patients = ep.groupby("Patient_No").agg(Patient_Name=("Patient_Name", "first"), Gender=("Gender", "first"),
                                            Blood_Group=("Blood_Group", "first"), First_Episode=("Admit_Date", "min"),
                                            Admissions=("Episode_ID", "count")).reset_index()
    wards = pd.DataFrame([dict(Ward_Code=a, Ward=b, First_Bed=c, Last_Bed=d, Open_Beds=e) for a, b, c, d, e in WARDS])
    return ep, patients, pd.DataFrame(cons), wards, census


def gen_screening(n=60000):
    """Mercyland community diabetes screening programme, 2023-2024 (adults)."""
    r = np.random.default_rng(SEED + 202)
    gender = np.where(r.random(n) < 0.56, "Female", "Male")
    age = np.clip(r.gamma(5.5, 8.2, n) + 14, 18, 90).round().astype(int)
    bmi = np.clip(r.normal(26.4 + 0.035 * (age - 40) + 1.2 * (gender == "Female"), 5.2), 15, 58).round(1)
    smoke = np.where(gender == "Male", _cat(r, n, ["Never", "Former", "Current", "Not Recorded"], [0.56, 0.12, 0.16, 0.16]),
                     _cat(r, n, ["Never", "Former", "Current", "Not Recorded"], [0.79, 0.03, 0.03, 0.15]))
    hyp_z = -5.2 + 0.06 * age + 0.07 * (bmi - 25)
    hyp = np.where(r.random(n) < 1 / (1 + np.exp(-hyp_z)), "Yes", "No")
    heart = np.where(r.random(n) < 1 / (1 + np.exp(-(-6.4 + 0.055 * age + 0.6 * (hyp == "Yes") + 0.5 * (smoke == "Current")))),
                     "Yes", "No")
    risk = -6.05 + 0.045 * age + 0.11 * (bmi - 25) + 0.55 * (hyp == "Yes") + 0.4 * (heart == "Yes") + 0.3 * (smoke == "Current")
    diab = r.random(n) < 1 / (1 + np.exp(-risk))
    hba1c = np.where(diab, np.clip(r.normal(7.6, 1.1, n), 5.6, 13.5), np.clip(r.normal(5.35, 0.45, n) + 0.006 * (age - 40), 4.0, 6.9))
    gluc = np.where(diab, np.clip(r.normal(178, 42, n), 105, 380), np.clip(r.normal(98, 15, n) + 0.15 * (age - 40), 65, 165))
    sites = np.array(["Gwarinpa", "Kubwa", "Nyanya", "Lugbe", "Karu", "Bwari"])
    site = sites[np.searchsorted(np.cumsum([0.22, 0.2, 0.17, 0.16, 0.15, 0.10]), r.random(n)).clip(0, 5)]
    date = pd.Timestamp("2023-01-09") + pd.to_timedelta(r.integers(0, 720, n), unit="D")
    out = pd.DataFrame({"Screen_ID": [f"MSS-{i:06d}" for i in range(1, n + 1)], "Screen_Date": date.date, "Site": site,
                        "Gender": gender, "Age": age, "Hypertension": hyp, "Heart_Condition": heart, "Smoking_Status": smoke,
                        "BMI": bmi, "HbA1c_Pct": hba1c.round(1), "Glucose_mg_dL": gluc.round().astype(int),
                        "Diabetes": np.where(diab, "Yes", "No")})
    return out.sort_values(["Screen_Date", "Screen_ID"]).reset_index(drop=True).assign(
        Screen_ID=lambda d: [f"MSS-{i:06d}" for i in range(1, len(d) + 1)])


# ======================================================================================================
# Kingsway University
# ======================================================================================================
FACULTIES = {
    "Science": ["Computer Science", "Microbiology", "Physics", "Biochemistry"],
    "Social Sciences": ["Economics", "Accounting", "Mass Communication", "Political Science"],
    "Engineering": ["Civil Engineering", "Electrical Engineering", "Mechanical Engineering"],
    "Arts & Education": ["English", "History", "Education"],
}
DEPT_CODE = {"Computer Science": "CSC", "Microbiology": "MCB", "Physics": "PHY", "Biochemistry": "BCH",
             "Economics": "ECO", "Accounting": "ACC", "Mass Communication": "MAC", "Political Science": "POL",
             "Civil Engineering": "CVE", "Electrical Engineering": "EEE", "Mechanical Engineering": "MEE",
             "English": "ENG", "History": "HIS", "Education": "EDU"}
LMH = np.array(["Low", "Medium", "High"])


def _cat(r, n, labels, p):
    return np.asarray(labels)[np.searchsorted(np.cumsum(p), r.random(n)).clip(0, len(labels) - 1)]


def gen_students(n=5840):
    r = np.random.default_rng(SEED + 303)
    depts = [(f, d) for f, lst in FACULTIES.items() for d in lst]
    w = np.array([9, 6, 5, 6, 9, 9, 7, 6, 6, 7, 6, 5, 4, 5], float)
    ch = r.choice(len(depts), n, p=w / w.sum())
    level = r.choice([100, 200, 300, 400], n, p=[0.30, 0.27, 0.23, 0.20])
    gender = np.where(r.random(n) < 0.52, "Female", "Male")
    hours = np.clip(np.round(r.normal(19, 6.2, n)), 1, 44).astype(int)
    att = r.integers(60, 101, n)
    sleep = np.clip(np.round(r.normal(7, 1.4, n)), 4, 10).astype(int)
    prev = r.integers(48, 101, n)
    tut = np.clip(r.poisson(1.6, n), 0, 8)
    ex = r.integers(0, 7, n)
    guard = _cat(r, n, LMH, [0.2, 0.5, 0.3]); lib = _cat(r, n, LMH, [0.2, 0.5, 0.3]); motiv = _cat(r, n, LMH, [0.3, 0.5, 0.2])
    club = _cat(r, n, ["Yes", "No"], [0.6, 0.4]); net = _cat(r, n, ["Yes", "No"], [0.92, 0.08])
    inc = _cat(r, n, LMH, [0.4, 0.4, 0.2]); lect = _cat(r, n, ["Low", "Medium", "High", "Not Recorded"], [0.1, 0.6, 0.28, 0.02])
    sch = _cat(r, n, ["Public Secondary", "Private Secondary"], [0.69, 0.31])
    peer = _cat(r, n, ["Negative", "Neutral", "Positive"], [0.21, 0.39, 0.40])
    lsup = _cat(r, n, ["Yes", "No"], [0.1, 0.9])
    ged = _cat(r, n, ["Secondary", "College", "Postgraduate", "Not Recorded"], [0.49, 0.30, 0.19, 0.02])
    com = _cat(r, n, ["Near", "Moderate", "Far", "Not Recorded"], [0.59, 0.30, 0.10, 0.01])
    eff = lambda a, d: pd.Series(a).map(d).fillna(0).values  # noqa: E731
    score = (29.5 + 0.30 * att + 0.36 * hours + 0.05 * prev + 0.55 * tut + 0.08 * ex
             + eff(guard, {"Low": -1.0, "High": 1.0}) + eff(lib, {"Low": -1.0, "High": 1.0}) + eff(motiv, {"Low": -0.5, "High": 0.5})
             + eff(inc, {"Low": -0.5, "High": 0.5}) + eff(peer, {"Negative": -0.5, "Positive": 0.5}) + eff(lsup, {"Yes": -0.9})
             + eff(com, {"Far": -0.9}) + eff(net, {"No": -0.8}) + r.normal(0, 3.1, n))
    hi = r.random(n) < 0.015                                       # high-achiever cluster, above the cloud
    score = np.where(hi, r.integers(84, 100, n), score)
    score = np.clip(np.round(score), 0, 100).astype(int)
    entry = 2024 - (level // 100 - 1)
    out = pd.DataFrame({
        "Matric_No": [f"KU/{e}/{DEPT_CODE[depts[c][1]]}/{s:04d}" for e, c, s in zip(entry, ch, range(1, n + 1))],
        "Faculty": [depts[i][0] for i in ch], "Department": [depts[i][1] for i in ch], "Level": level, "Gender": gender,
        "Study_Hours_Wk": hours, "Lecture_Attendance_Pct": att, "Sleep_Hrs_Night": sleep, "Last_Sem_Score": prev,
        "Tutorials_Attended": tut, "Exercise_Hrs_Wk": ex, "Guardian_Support": guard, "Library_Access": lib, "Club_Member": club,
        "Motivation": motiv, "Home_Internet": net, "Household_Income": inc, "Lecturer_Rating": lect, "Secondary_School": sch,
        "Peer_Effect": peer, "Learning_Support": lsup, "Guardian_Education": ged, "Commute_Band": com, "Final_Score": score})
    out["Grade"] = pd.cut(out.Final_Score, [-1, 44, 49, 59, 69, 100], labels=["F", "D", "C", "B", "A"]).astype(str)
    out["Attendance_Band"] = pd.cut(out.Lecture_Attendance_Pct, [0, 69, 79, 89, 100],
                                    labels=["<70%", "70-79%", "80-89%", "90-100%"]).astype(str)
    out["At_Risk"] = np.where((out.Final_Score < 62) | (out.Lecture_Attendance_Pct < 65), "Yes", "No")
    return out


STATES = ["Kwara", "Lagos", "Oyo", "Ogun", "Osun", "Kogi", "Niger", "FCT", "Ekiti", "Ondo", "Edo", "Kano", "Kaduna", "Anambra",
          "Enugu", "Delta", "Rivers", "Benue", "Plateau", "Imo", "Abia", "Akwa Ibom", "Cross River", "Nasarawa", "Sokoto",
          "Katsina", "Bauchi", "Borno", "Adamawa", "Gombe", "Taraba", "Yobe", "Zamfara", "Kebbi", "Jigawa", "Bayelsa", "Ebonyi"]


def gen_welfare(n=9600):
    """Student welfare survey with a PHQ-9 depression screen: special-category data."""
    r = np.random.default_rng(SEED + 404)
    sw = np.array([30, 12, 10, 8, 7, 6, 6, 5, 4, 4, 3, 3, 3] + [1.5] * (len(STATES) - 13), float)
    progs = [d for lst in FACULTIES.values() for d in lst]
    press = _cat(r, n, [1, 2, 3, 4, 5], [0.12, 0.18, 0.30, 0.24, 0.16]).astype(int)
    money = _cat(r, n, [1, 2, 3, 4, 5], [0.14, 0.18, 0.27, 0.22, 0.19]).astype(int)
    sat = _cat(r, n, [1, 2, 3, 4, 5], [0.10, 0.17, 0.30, 0.27, 0.16]).astype(int)
    sleep = _cat(r, n, ["Less than 5 hours", "5-6 hours", "7-8 hours", "More than 8 hours"], [0.24, 0.30, 0.32, 0.14])
    diet = _cat(r, n, ["Healthy", "Moderate", "Unhealthy"], [0.28, 0.40, 0.32])
    fam = _cat(r, n, ["Yes", "No"], [0.18, 0.82])
    z = (-3.6 + 0.48 * press + 0.36 * money - 0.25 * sat + pd.Series(sleep).map({"Less than 5 hours": 0.6, "5-6 hours": 0.25,
                                                                                    "7-8 hours": 0, "More than 8 hours": 0.1}).values
         + 0.35 * (diet == "Unhealthy") + 0.4 * (fam == "Yes"))
    dep = r.random(n) < 1 / (1 + np.exp(-z))
    sh = np.where(r.random(n) < np.where(dep, 0.22, 0.04), "Yes", "No")
    out = pd.DataFrame({"Respondent_ID": [f"KWS-{i:05d}" for i in range(1, n + 1)],
                        "Gender": np.where(r.random(n) < 0.53, "Female", "Male"),
                        "Age": np.clip(np.round(r.gamma(4, 1.3, n) + 16), 16, 34).astype(int),
                        "State_of_Origin": np.array(STATES)[r.choice(len(STATES), n, p=sw / sw.sum())],
                        "Programme": np.array(progs)[r.integers(0, len(progs), n)],
                        "Study_Pressure_1to5": press,
                        "CGPA": np.clip(r.normal(3.35 - 0.12 * dep, 0.72, n), 1.0, 5.0).round(2),
                        "Course_Satisfaction_1to5": sat, "Sleep_Band": sleep, "Diet_Quality": diet,
                        "Daily_Study_Hrs": np.clip(np.round(r.gamma(3, 1.6, n)), 0, 12).astype(int),
                        "Money_Worry_1to5": money, "Family_MH_History": fam, "Self_Harm_Thoughts": sh,
                        "PHQ9_Positive": np.where(dep, "Yes", "No")})
    return out


def gen_digital(n=640):
    r = np.random.default_rng(SEED + 505)
    use = np.clip(r.normal(4.6, 1.5, n), 1.0, 9.0).round(1)
    out = pd.DataFrame({
        "Matric_No": [f"KDH-{i:04d}" for i in range(1, n + 1)],
        "Age": r.integers(17, 29, n), "Gender": np.where(r.random(n) < 0.52, "Female", "Male"),
        "Academic_Level": _cat(r, n, ["Undergraduate", "Postgraduate", "Pre-degree"], [0.72, 0.18, 0.10]),
        "Country": _cat(r, n, ["Nigeria", "Ghana", "Benin", "Cameroon", "Togo"], [0.86, 0.05, 0.04, 0.03, 0.02]),
        "Daily_Screen_Hrs": use,
        "Top_App": _cat(r, n, ["WhatsApp", "TikTok", "Instagram", "X", "YouTube", "Snapchat", "Facebook"],
                        [0.30, 0.22, 0.18, 0.10, 0.10, 0.05, 0.05]),
        "Hurts_Studies": np.where(r.random(n) < 1 / (1 + np.exp(-(use - 4.8) * 1.4)), "Yes", "No"),
        "Sleep_Hrs": np.clip(9.3 - 0.62 * use + r.normal(0, 0.55, n), 3.5, 9.8).round(1),
        "Wellbeing_Score": np.clip(np.round(9.6 - 0.75 * use + r.normal(0, 0.8, n)), 3, 9).astype(int),
        "Relationship_Status": _cat(r, n, ["Single", "In Relationship", "Complicated"], [0.55, 0.39, 0.06]),
        "Phone_Conflicts": np.clip(np.round(0.65 * use - 0.4 + r.normal(0, 0.8, n)), 0, 5).astype(int),
        "Dependence_Score": np.clip(np.round(1.1 * use + 1.2 + r.normal(0, 0.7, n)), 2, 9).astype(int)})
    return out


# ======================================================================================================
# Atlantic Pearl Hotel, Lekki
# ======================================================================================================
ROOM_RATE = {"Standard": 82000, "Standard Twin": 90000, "Business": 128000, "Deluxe": 118000, "Executive": 158000,
             "Family Suite": 205000, "Presidential Suite": 420000}
CHANNELS = ["Online Travel Agent", "Walk-in / Travel Agent", "Corporate", "Complimentary", "Airline Crew"]


def gen_hotel(n=34000):
    r = np.random.default_rng(SEED + 606)
    days = pd.date_range("2023-07-01", "2024-12-31", freq="D")
    m = days.month
    wt = np.select([m == 12, m.isin([7, 8]), m.isin([1, 2]), m == 4], [1.45, 1.15, 0.80, 1.1], 1.0) * \
        np.where(days.year == 2024, 1.12, 1.0) * np.array([0.9, 0.95, 1.0, 1.05, 1.15, 1.2, 0.85])[days.dayofweek]
    arr = pd.to_datetime(r.choice(days.values, n, p=wt / wt.sum()))
    yr = arr.year.values
    chan = _cat(r, n, CHANNELS, [0.635, 0.285, 0.058, 0.012, 0.010])
    # room mix: Deluxe rooms reopened after refurbishment in 2024; Standard Twin rooms mostly converted
    room_p = {2023: [0.58, 0.20, 0.05, 0.06, 0.07, 0.03, 0.01], 2024: [0.55, 0.06, 0.05, 0.24, 0.06, 0.03, 0.01]}
    rooms = np.array(list(ROOM_RATE))
    room = np.where(yr == 2023, rooms[np.searchsorted(np.cumsum(room_p[2023]), r.random(n)).clip(0, 6)],
                    rooms[np.searchsorted(np.cumsum(room_p[2024]), r.random(n)).clip(0, 6)])
    lead_mu = pd.Series(chan).map({"Online Travel Agent": 88, "Walk-in / Travel Agent": 72, "Corporate": 16,
                                   "Complimentary": 9, "Airline Crew": 4}).values
    lead = np.minimum(np.round(r.exponential(lead_mu)), 443).astype(int)
    rep = np.where(r.random(n) < np.where(chan == "Corporate", 0.30, np.where(chan == "Complimentary", 0.45, 0.02)), "Yes", "No")
    req = np.clip(r.poisson(np.where(chan == "Online Travel Agent", 0.75, 0.35)), 0, 5)
    wkend = r.choice([0, 1, 2], n, p=[0.36, 0.34, 0.30])
    wk = np.clip(r.poisson(np.where(chan == "Corporate", 2.0, 2.6)), 0, 12)
    zero = (wkend + wk) == 0
    wk = np.where(zero, 1, wk)
    nights = wkend + wk
    season = np.select([m == 12, m.isin([7, 8])], [1.25, 1.08], 1.0)
    rate = (pd.Series(room).map(ROOM_RATE).values * np.where(yr == 2024, 1.16, 1.0) * season[np.searchsorted(days.values, arr.values)]
            * pd.Series(chan).map({"Corporate": 0.88, "Airline Crew": 0.95, "Walk-in / Travel Agent": 1.0,
                                   "Online Travel Agent": 1.06, "Complimentary": 0.0}).values * r.lognormal(0, 0.1, n))
    rate = (np.round(rate / 100) * 100).astype(float)
    # cancellation: lead time, channel, requests, returning guests, year; Deluxe a little lower, Twin a little higher
    lt = np.log1p(lead)
    z = (-4.75 + 0.86 * lt + pd.Series(chan).map({"Online Travel Agent": 0.35, "Walk-in / Travel Agent": 0.05, "Corporate": -0.9,
                                                  "Airline Crew": 0.2, "Complimentary": -9}).values
         - 0.45 * req - 2.4 * (rep == "Yes") + 1.05 * (yr == 2024)
         + pd.Series(room).map({"Deluxe": -0.15, "Standard Twin": 0.15}).fillna(0).values)
    canc = r.random(n) < 1 / (1 + np.exp(-z))
    out = pd.DataFrame({
        "Check_In": arr, "Adults": r.choice([1, 2, 3], n, p=[0.27, 0.66, 0.07]),
        "Children": r.choice([0, 1, 2, 3], n, p=[0.90, 0.06, 0.035, 0.005]),
        "Weekend_Nights": wkend, "Week_Nights": wk, "Nights": nights,
        "Board_Basis": _cat(r, n, ["Room Only", "Bed & Breakfast", "Half Board", "Full Board"], [0.24, 0.66, 0.09, 0.01]),
        "Car_Park": np.where(r.random(n) < 0.04, "Yes", "No"), "Room_Category": room, "Days_Ahead": lead,
        "Sales_Channel": chan, "Returning_Guest": rep,
        "Prior_Cancellations": np.where(rep == "Yes", r.poisson(0.15, n), 0),
        "Prior_Stays": np.where(rep == "Yes", 1 + r.poisson(2.5, n), 0),
        "Room_Rate_NGN": rate, "Guest_Requests": req,
        "Reservation_Status": np.where(canc, "Cancelled", "Stayed")})
    out["Booked_On"] = (out.Check_In - pd.to_timedelta(out.Days_Ahead, unit="D")).dt.date
    out["Reservation_Value_NGN"] = out.Room_Rate_NGN * out.Nights
    out = out.sort_values(["Check_In", "Booked_On"]).reset_index(drop=True)
    out.insert(0, "Reservation_ID", [f"APH-{i:05d}" for i in range(1, len(out) + 1)])
    out["Check_In"] = out.Check_In.dt.date
    cols = ["Reservation_ID", "Booked_On", "Check_In", "Adults", "Children", "Weekend_Nights", "Week_Nights", "Nights",
            "Board_Basis", "Car_Park", "Room_Category", "Days_Ahead", "Sales_Channel", "Returning_Guest", "Prior_Cancellations",
            "Prior_Stays", "Room_Rate_NGN", "Guest_Requests", "Reservation_Status", "Reservation_Value_NGN"]
    return out[cols]
