from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.13 Ethics PII PHI'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Bootcamp datasets")
    return T(tag, **k)


TOPICS = [
    O("4.13.1",
      what="Analysts describe **real people** with numbers: patients, students, customers, staff. Careless handling can "
           "expose them, and careless analysis can treat them **unfairly**. Ethics is part of doing the job well.",
      like="Like a doctor's oath: skill matters, but so does doing no harm.",
      how=["Ask: whose data is this?", "Ask: could this harm or expose anyone?", "Ask: is the result fair to every group?"],
      tip="Ask these questions before the analysis, not after.",
      trap="'It's just numbers' — every row is a person.",
      case="The bootcamp datasets all describe people.",
      steps=["List each dataset.", "Name who is in it.", "Name the risk if mishandled."],
      table=[["Dataset", "Who is in it", "Risk if mishandled"], ["Lifecare admissions", "Patients", "Health details exposed"],
             ["Crestview wellbeing survey", "27,870 students", "Mental-health answers exposed"],
             ["SuperMart Online customers", "2,500 customers", "Phone and email leaked"],
             ["Salary survey", "1,781 respondents", "Pay of named people revealed"]],
      so="Every dataset in this course needs careful handling.",
      split=4.70),

    O("4.13.2",
      what="**Personal data (PII)** is anything that can identify a person. **Direct identifiers** do it alone (name, "
           "email, phone). **Quasi-identifiers** do it in combination (date of birth + city + gender).",
      like="Like a description of a suspect: no name, but 'tall, red car, lives on Allen Avenue' may be enough.",
      how=["Find direct identifiers: remove them.", "Find quasi-identifiers: band or generalise them.", "Analyse by an ID, not a name."],
      tip="Use customer_id in analysis; leave names in the database.",
      trap="Thinking data is anonymous because the name column is gone.",
      who="SuperMart Online",
      case="Classify the customers table.",
      steps=["Read each column.", "Direct, quasi or neither?", "Write a handling rule."],
      table=[["Columns", "Type", "Handling rule"], ["first_name, last_name, email, phone_number", "Direct identifier", "Never leave the database"],
             ["date_of_birth, city", "Quasi-identifier", "Band age; use state or zone"],
             ["customer_id", "Pseudonymous key", "Use for analysis"], ["membership_tier, total_spent", "Neither alone", "OK in aggregates"]],
      so="Remove direct identifiers; generalise quasi-identifiers.",
      split=4.20),

    O("4.13.3",
      what="**Health data (PHI)** is personal data about a person's health: conditions, medication, test results, "
           "screening values. It is a **special category** and needs the strictest handling.",
      like="Like a sealed medical envelope: only the right people open it.",
      how=["Identify health columns.", "Keep identifiers away from them.", "Report only aggregates, with a minimum cell size."],
      tip="A minimum of 5 per cell is a common rule for published health tables.",
      trap="Emailing an extract with Patient_Name and Medical_Condition together.",
      who="Lifecare Hospital",
      case="Classify Lifecare's health columns.",
      steps=["Find identifiers.", "Find health fields.", "Set the rule for each."],
      table=[["Columns", "Classification", "Rule"], ["Patient_Name, Patient_ID", "Direct identifier (PHI)", "Remove before any extract leaves"],
             ["Medical_Condition, Medication, Test_Result", "Health data", "Aggregate only; min 5 per cell"],
             ["HbA1c_Level, Blood_Glucose_mg_dL", "Health data", "Pseudonymised Screening_ID only"]],
      so="Health data never travels with names attached.",
      split=4.20),

    O("4.13.4",
      what="**Data minimisation**: use only the columns the question needs. Fewer columns means less to leak and "
           "less to get wrong.",
      like="Like packing for a day trip: take what you need, not the whole wardrobe.",
      how=["Write the question first.", "List the columns it needs.", "Drop everything else before you start."],
      tip="Power Query > Choose Columns makes this one step you can repeat.",
      trap="Exporting the full table 'just in case'.",
      who="Lifecare Hospital",
      case="Question: which departments have the highest readmission rate?",
      steps=["Count the columns available.", "Pick the ones needed.", "Drop the rest."],
      table=[["Measure", "Value"], ["Columns in the admissions extract", "22"],
             ["Columns needed", "3: Department, Readmit_30d, Admission_Date"],
             ["Columns dropped", "19, incl. Patient_Name, Doctor_Name, Billing_Amount"]],
      so="3 columns answer the question. The other 19 only add risk.",
      split=4.90),

    O("4.13.5",
      what="**Aggregation** reports counts instead of rows. **Suppression** hides small cells (for example fewer than "
           "5) that could point to one person. **Generalisation** (age bands, region instead of town) makes people "
           "harder to single out.",
      like="Like blurring faces in a crowd photo: you see the crowd, not the individuals.",
      how=["Aggregate: COUNTIFS by group.", "Suppress: IF(count<5, \"<5\", count).", "Generalise: band ages, drop small areas."],
      tip="Count how many people are unique on their combination of fields.",
      trap="Publishing a group of 3 with its rate.",
      who="Crestview Wellbeing Survey",
      case="Publish depression-screen results by programme from 27,870 responses.",
      steps=["Count respondents and positives by programme.", "Apply the <5 and <10 rules.", "Check uniqueness."],
      table=[["Check", "Result"], ["Pre-degree (largest)", "6,080 respondents; 70.8% positive"],
             ["Others (smallest)", "35 respondents; 60.0% positive"],
             ["Unique on Gender + Age + State + Programme", "3,652 people"],
             ["Unique after age bands, no State", "14 people"]],
      so="No programme needs suppressing, but exact fields make 3,652 people unique. Banding cuts it to 14.",
      split=4.90),

    O("4.13.6",
      what="Data collected for one **purpose** (care, admissions, a sale) needs a **legal basis** before you use it "
           "for another. In Nigeria this falls under the **Nigeria Data Protection Act 2023**.",
      like="Like borrowing a car for the school run: it doesn't mean you can take it on holiday.",
      how=["Ask: why was this collected?", "Ask: does my use fit that purpose?", "If not: check with the data owner or DPO."],
      tip="DPO means Data Protection Officer. Know who yours is.",
      trap="Using survey data for marketing because 'we already have it'.",
      case="Purpose and legal basis for each dataset.",
      steps=["Name the original purpose.", "Name the basis.", "Check new uses against it."],
      table=[["Dataset", "Typical basis", "Check before new use"], ["Lifecare", "Provision of care / public interest in health", "Ethics approval for research"],
             ["SuperMart", "Contract + legitimate interest", "Marketing needs consent"],
             ["Crestview survey", "Consent", "Only the stated purpose"]],
      so="New uses need checking. 'We have the data' is not a reason.",
      split=4.20),

    O("4.13.7",
      what="**Bias** enters through unrepresentative samples, missing groups, or headline averages that hide "
           "differences. A **fairness check** compares outcomes within comparable groups.",
      like="Like judging a race where some runners started further back.",
      how=["Compare the overall gap.", "Compare within each level.", "Report tiny groups as suppressed, not dropped."],
      tip="The largest within-group gap is often the real issue.",
      trap="Quoting only the headline average.",
      who="Salary Survey",
      case="Gender pay gap as a fairness test.",
      steps=["Raw gap, all levels.", "Gap within each seniority.", "Count small groups."],
      table=[["Measure", "Value", "Meaning"], ["Raw gap (all levels)", "9.0%", "What a headline quotes"],
             ["Mid level gap", "14.0%", "The largest: investigate"], ["Junior / Senior / Executive", "3.9% / 5.9% / 6.4%", "Smaller gaps"],
             ["'Other' gender respondents", "7", "Too few: suppress, don't drop"]],
      so="The headline 9.0% hides a 14.0% gap at Mid level.",
      split=4.70),

    O("4.13.8",
      what="Keep data **only as long as needed**, in **approved places**, **encrypted**, with **restricted** access. "
           "Delete training extracts when the course ends.",
      like="Like a bank vault: locked, logged, and emptied when the account closes.",
      how=["Store on encrypted, approved drives.", "Restrict shared links to named people.", "Delete extracts on schedule."],
      tip="Put a delete-by date in the file name or the log.",
      trap="Copies left on laptops, USB sticks and email.",
      case="Handling rules for the bootcamp data.",
      steps=["Where is it stored?", "Who can access it?", "When is it deleted?"],
      table=[["Control", "Practice"], ["Retention", "Training extracts deleted at the end of the bootcamp"],
             ["Storage", "Encrypted laptops; sensitive survey file restricted"],
             ["Sharing", "No personal data in email or public AI tools"], ["Links", "Restricted to named people"]],
      so="Most leaks come from forgotten copies. Fewer copies, fewer leaks.",
      split=4.90),

    O("4.13.9",
      what="**Governance** makes ethics routine: a named **data owner** approves extracts, the **analyst** documents "
           "minimisation, and the **DPO** reviews new uses.",
      like="Like a building's fire plan: named wardens, clear routes, regular drills.",
      how=["Know the data owner.", "Document what you took and why.", "Ask the DPO before new uses."],
      tip="Keep the minimisation note with the analysis files.",
      trap="Nobody knowing who approved an extract.",
      case="Who does what for a Lifecare readmission study.",
      steps=["Owner approves.", "Analyst documents.", "DPO reviews."],
      who="Lifecare Hospital",
      table=[["Role", "Action"], ["Data owner", "Approves the 3-column extract"],
             ["Analyst", "Documents that 19 columns were dropped"], ["DPO", "Reviews any new use of the data"],
             ["Everyone", "Follows the retention and sharing rules"]],
      so="Clear roles mean every extract can be traced and justified.",
      split=4.90),
]
