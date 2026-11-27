# Verified dataset facts for slide examples (generated)

## SuperMart Stores (SuperMart_Stores_Sales_2022_2025.csv, 9,994 rows; Directory 12 stores)
Store_ID                  Store_Name          City     State          Zone     Opened  Floor_Area_sqm     Store_Manager  Annual_Target_2025
    ST01               Lagos - Ikeja         Ikeja     Lagos    South-West 2015-03-02            2400    Folake Adeyemi           142000000
    ST02               Lagos - Lekki         Lekki     Lagos    South-West 2017-08-14            1850      Tunde Bakare            63000000
    ST03          Ibadan - Ring Road        Ibadan       Oyo    South-West 2016-05-09            1600      Kemi Oladipo            37000000
    ST04         Port Harcourt - GRA Port Harcourt    Rivers   South-South 2015-11-23            2100       Ebere Amadi            97000000
    ST05      Benin City - Ring Road    Benin City       Edo   South-South 2018-02-05            1400 Osagie Igbinedion            59000000
    ST06                  Uyo - Ewet           Uyo Akwa Ibom   South-South 2019-06-17            1150        Uduak Etim            62000000
    ST07                Abuja - Wuse         Abuja       FCT North-Central 2016-01-11            2200       Hauwa Bello            52000000
    ST08               Abuja - Garki         Abuja       FCT North-Central 2018-09-03            1500     Ibrahim Garba            39000000
    ST09              Jos - Rayfield           Jos   Plateau North-Central 2020-04-20            1050      Ladi Danjuma            61000000
    ST10 Enugu - Independence Layout         Enugu     Enugu    South-East 2017-03-13            1700    Chinedu Okafor            43000000
    ST11       Onitsha - Main Market       Onitsha   Anambra    South-East 2019-10-07            1300        Nkechi Eze            53000000
    ST12         Owerri - New Owerri        Owerri       Imo    South-East 2021-02-01            1100      Ikenna Nwosu            39000000
      lines  orders      sales    profit  margin
Year                                            
2022   1993     969  484247497  49543972  0.1023
2023   2102    1038  470532511  61618615  0.1310
2024   2587    1315  609205599  81795186  0.1343
2025   3312    1687  733215255  93439242  0.1274
FY2025 by store:
                             lines      sales    profit  margin  Annual_Target_2025  vs_target
Store_Name                                                                                    
Lagos - Ikeja                  663  146388345  29366453   0.201           142000000      1.031
Port Harcourt - GRA            352   93922995  24357072   0.259            97000000      0.968
Lagos - Lekki                  252   70819092  18561342   0.262            63000000      1.124
Jos - Rayfield                 210   66227429  18901912   0.285            61000000      1.086
Benin City - Ring Road         299   62116000   -713637  -0.011            59000000      1.053
Uyo - Ewet                     270   57043909   9587120   0.168            62000000      0.920
Onitsha - Main Market          199   47779842  -6220517  -0.130            53000000      0.902
Enugu - Independence Layout    185   43540305   5290482   0.122            43000000      1.013
Abuja - Wuse                   317   43421758  -8838505  -0.204            52000000      0.835
Abuja - Garki                  251   37448941  -2512563  -0.067            39000000      0.960
Ibadan - Ring Road             180   32920929  -4118850  -0.125            37000000      0.890
Owerri - New Owerri            134   31585710   9778933   0.310            39000000      0.810
FY2025 by category:
                 lines      sales    profit  margin
Category                                           
Furniture          686  215387269   3018378   0.014
Office Supplies   2002  246097175  39736618   0.161
Technology         624  271730811  50684246   0.187
FY2025 sub-category:
                  sales    profit
Sub_Category                     
Phones        105340516  12849317
Chairs         95554353   7643539
Binders        72788045   7669738
Storage        69677618   7402801
Copiers        62899388  25031789
Tables         60893542  -8140697
Accessories    59946232  15672357
Machines       43544675  -2869217
Appliances     42926932   7865264
Bookcases      30024280   -583627
Furnishings    28915094   4099163
Paper          27694718  12040841
Supplies       16049410   -955314
Art             8863068   2221967
Labels          3861216   1744615
Envelopes       3378574   1441757
Fasteners        857594    304949
FY2025 by zone: {'North-Central': 147098128, 'South-East': 122905857, 'South-South': 213082904, 'South-West': 250128366}
Segments FY25 lines: {'Consumer': 1668, 'Corporate': 980, 'Home Office': 664}
Ship modes FY25: {'Standard Class': 1897, 'Second Class': 657, 'First Class': 572, 'Same Day': 186}
FY25 monthly sales: {1: 43971374, 2: 20301133, 3: 58872353, 4: 36521536, 5: 44261110, 6: 52981726, 7: 45264416, 8: 63120888, 9: 87866652, 10: 77776923, 11: 118447825, 12: 83829319}
Discount values: [0.0, 0.1, 0.15, 0.2, 0.3, 0.32, 0.4, 0.45, 0.5, 0.6, 0.7, 0.8]
FY25 customers: 693  products: 1525  avg qty: 3.77  max line sales: 13999960  min: 444
Ship days FY25 mean: 3.91 {0: 173, 1: 164, 2: 431, 3: 388, 4: 866, 5: 670, 6: 409, 7: 211}
First 6 FY2025 rows:
 Row_ID       Order_ID Order_Date  Customer_Name     Store_Name        Category Sub_Category                                                                Product_Name  Sales  Quantity  Discount  Profit
   6683 SM-2025-107503 2025-01-01 Chijioke Obasi     Uyo - Ewet       Furniture  Furnishings                                          Linden 10" Round Wall Clock, Black  48896         4       0.2    8557
   6684 SM-2025-144463 2025-01-01 Gbenga Anyanwu  Lagos - Ikeja       Furniture  Furnishings                         Howard Miller 11-1/2" Diameter Brentwood Wall Clock 474430        11       0.0  199261
   6685 SM-2025-154466 2025-01-01   Safiya Obasi Jos - Rayfield Office Supplies      Binders                                     Wilson Jones Easy Flow II Sheet Lifters   3600         2       0.0    1728
   6686 SM-2025-151750 2025-01-01   Onyinye Etim   Abuja - Wuse Office Supplies      Storage                                               SAFCO Boltless Steel Shelving 454560         5       0.2 -107958
   6687 SM-2025-151750 2025-01-01   Onyinye Etim   Abuja - Wuse       Furniture  Furnishings Tenex Carpeted, Granite-Look or Clear Contemporary Contour Shape Chair Mats 141420         5       0.6 -187382
   6688 SM-2025-151750 2025-01-01   Onyinye Etim   Abuja - Wuse       Furniture       Chairs                                Office Star - Contemporary Task Swivel Chair 310744         4       0.3  -26635
Top 5 products FY25 by sales:
Product_Name
Canon imageCLASS 2200 Advanced Copier                 35699898
Martin Yale Chadless Opener Electric Letter Opener    11825902
GBC DocuBind TL300 Electric Binding System            10943278
Hewlett Packard LaserJet 3310 Copier                   9239846
Samsung Galaxy Mega 6.3                                9239780
Top customers FY25:
Customer_Name
Ikechukwu Ibe     14203278
Bala Onyeka       13723498
Chinedu Agbaje    10522550
Ngozi Yusuf        8459936
Chidi Adebayo      8167420
Lagos - Ikeja FY25 monthly: {1: 5163432, 2: 6864883, 3: 11331194, 4: 8784206, 5: 10157356, 6: 9086966, 7: 12923859, 8: 14534322, 9: 18866852, 10: 15776301, 11: 13580951, 12: 19318023}
Loss-making lines FY25: 620  lines with discount>0: 1722

## SuperMart Online (5 tables: transactions 11,000; customers 2,500; products 800; sellers 200; deliveries 10,430)
order_date range: 2020-01-20 2025-12-30  orders: 11000
order_status: {'Delivered': 8286, 'Completed': 1633, 'Cancelled': 570, 'Returned': 511}
payment_method: {'Paystack': 3133, 'Flutterwave': 2646, 'Bank Transfer': 2182, 'Cash on Delivery': 1928, 'USSD': 764, 'Card': 347}
payment_status: {'Paid': 10430, 'Failed': 299, 'Refunded': 271}
total_price sum all: 1,394,713,700  delivered+completed: 1,256,619,400  commission sum: 143,638,049
by year total_price: {2020: '32,292,900', 2021: '118,697,200', 2022: '210,443,850', 2023: '252,860,500', 2024: '348,433,250', 2025: '431,986,000'}
customers tier: {'Silver': 743, 'Gold': 737, 'Bronze': 512, 'Platinum': 508}
customers zone: {'South-West': 661, 'South-South': 506, 'North-Central': 431, 'South-East': 411, 'North-West': 327, 'North-East': 164}
top cities: {'Ado Ekiti': 99, 'Port Harcourt': 95, 'Benin City': 93, 'Makurdi': 91, 'Calabar': 90, 'Ikorodu': 89, 'Lagos': 89, 'Ilorin': 89}
product categories: {'Health': 138, 'Automotive': 124, 'Fashion': 121, 'Electronics': 115, 'Home': 109, 'Food': 102, 'Beauty': 91}
avg price: 77076  median: 21750.0
couriers: {'Gokada': 1340, 'FedEx Lagos': 1336, 'Max.ng': 1331, 'Peace Mass Transit': 1310, 'UPS Abuja': 1298, 'ABC Transport': 1287, 'DHL Nigeria': 1281, 'Red Star Express': 1247}
delivery_status: {'Delivered': 8286, 'Pending': 2144}
sellers account_status: {'Active': 154, 'Suspended': 46}
revenue by product category: {'Electronics': '717,407,700', 'Home': '368,892,800', 'Fashion': '120,395,900', 'Automotive': '94,870,550', 'Food': '53,537,400', 'Health': '27,882,300', 'Beauty': '11,727,050'}
sample transactions:
transaction_id         order_id customer_id product_id seller_id delivery_id  quantity  unit_price  total_price  commission_percentage  commission_amount order_status payment_method payment_status order_date
     TRX000001 ORD2024050500001   CUST02223  PROD00575  SELL0194   DEL000001         3      7200.0      21600.0                    7.5             1620.0    Delivered    Flutterwave           Paid 2024-05-05
     TRX000002 ORD2025052100002   CUST02179  PROD00018  SELL0019   DEL000002         1    240000.0     240000.0                   12.5            30000.0    Delivered    Flutterwave           Paid 2025-05-21
     TRX000003 ORD2023060300003   CUST01121  PROD00530  SELL0200   DEL000003         1     44000.0      44000.0                    7.5             3300.0    Delivered  Bank Transfer           Paid 2023-06-03
     TRX000004 ORD2025071500004   CUST00159  PROD00410  SELL0138   DEL000004         1     43700.0      43700.0                    5.0             2185.0    Delivered    Flutterwave           Paid 2025-07-15

## Crestview University (Crestview_Student_Performance.csv, 6,607 students)
faculty: {'Social Sciences': 2250, 'Science': 1943, 'Engineering': 1400, 'Arts & Education': 1014}
departments: {'Arts & Education': ['English', 'History', 'Education'], 'Engineering': ['Electrical Engineering', 'Mechanical Engineering', 'Civil Engineering'], 'Science': ['Microbiology', 'Biochemistry', 'Physics', 'Computer Science'], 'Social Sciences': ['Accounting', 'Political Science', 'Mass Communication', 'Economics']}
level: {100: 1994, 200: 1767, 300: 1508, 400: 1338}
grade: {'B': 4914, 'A': 1625, 'C': 68}  grade rule? exam min by grade: {'min': {'A': 70, 'B': 60, 'C': 55}, 'max': {'A': 100, 'B': 69, 'C': 59}}
at risk: {'No': 5706, 'Yes': 901}
attendance band: {'90-100%': 1746, '70-79%': 1681, '80-89%': 1607, '<70%': 1573}
       Hours_Studied_Per_Week  Attendance_Pct  Sleep_Hours  Previous_Score  Tutoring_Sessions  Exam_Score
count                 6607.00         6607.00      6607.00         6607.00            6607.00     6607.00
mean                    19.98           79.98         7.03           75.07               1.49       67.24
std                      5.99           11.55         1.47           14.40               1.23        3.89
min                      1.00           60.00         4.00           50.00               0.00       55.00
25%                     16.00           70.00         6.00           63.00               1.00       65.00
50%                     20.00           80.00         7.00           75.00               1.00       67.00
75%                     24.00           90.00         8.00           88.00               2.00       69.00
max                     44.00          100.00        10.00          100.00               8.00      100.00
mean exam by faculty: {'Arts & Education': 67.1, 'Engineering': 67.2, 'Science': 67.3, 'Social Sciences': 67.3}
at-risk rate by faculty: {'Arts & Education': 0.146, 'Engineering': 0.139, 'Science': 0.141, 'Social Sciences': 0.127}
corr attendance-exam: 0.581  hours-exam: 0.446  prev-exam: 0.175
gender: {'Male': 3814, 'Female': 2793}
sample:
      Student_ID          Faculty             Department  Level Gender  Hours_Studied_Per_Week  Attendance_Pct  Exam_Score Grade At_Risk
CU/2021/EEE/0001      Engineering Electrical Engineering    400   Male                      23              84          67     B      No
CU/2024/ENG/0002 Arts & Education                English    100 Female                      19              64          61     B     Yes
CU/2024/MCB/0003          Science           Microbiology    100   Male                      24              98          74     A      No
CU/2024/ACC/0004  Social Sciences             Accounting    100   Male                      29              89          71     A      No
CU/2023/POL/0005  Social Sciences      Political Science    200 Female                      19              92          70     A      No

## Lifecare General Hospital (Admissions 2019-2024: 54,860; wards 4; doctors 28; census 1,797 days)
Ward_ID          Ward  First_Room  Last_Room  Staffed_Beds
     W1  Emerald Ward         101        200           130
     W2 Sapphire Ward         201        300           130
     W3     Ruby Ward         301        400           130
     W4    Topaz Ward         401        500           130
admissions by year: {2019: 7280, 2020: 11157, 2021: 10795, 2022: 10891, 2023: 10918, 2024: 3819}
FY2023: n 10918  LOS mean 15.49  median 15.0  min/max 1 30  billing sum 13984047685.44  mean bill 1280825  readmit rate 0.1034
departments FY23: {'Endocrinology': 3682, 'Cardiology': 1828, 'Oncology': 1824, 'Orthopaedics & Rheumatology': 1806, 'Respiratory': 1778}
conditions FY23: {'Obesity': 1848, 'Diabetes': 1834, 'Hypertension': 1828, 'Cancer': 1824, 'Arthritis': 1806, 'Asthma': 1778}
admission type FY23: {'Urgent': 3663, 'Emergency': 3658, 'Elective': 3597}
HMO FY23: {'AXA Mansard Health': 2250, 'NHIA (National Health Insurance)': 2238, 'Hygeia HMO': 2215, 'Reliance HMO': 2148, 'Leadway Health': 2067}
ward FY23: {'Ruby Ward': 2798, 'Sapphire Ward': 2755, 'Topaz Ward': 2734, 'Emerald Ward': 2631}
test result FY23: {'Inconclusive': 3699, 'Abnormal': 3670, 'Normal': 3549}
medication FY23: {'Ibuprofen': 2250, 'Lipitor': 2209, 'Penicillin': 2199, 'Aspirin': 2165, 'Paracetamol': 2095}
age FY23: mean 51.3  min 13  max 89  75+: 1764
gender FY23: {'Female': 5482, 'Male': 5436}
blood types FY23: {'AB+': 1414, 'A-': 1397, 'AB-': 1377, 'O+': 1373, 'A+': 1371, 'B+': 1350, 'O-': 1336, 'B-': 1300}
mean bill by dept FY23: {'Cardiology': 1283004.0, 'Endocrinology': 1284843.0, 'Oncology': 1279622.0, 'Orthopaedics & Rheumatology': 1290056.0, 'Respiratory': 1262122.0}
readmit rate by dept FY23: {'Cardiology': 0.114, 'Endocrinology': 0.099, 'Oncology': 0.132, 'Orthopaedics & Rheumatology': 0.083, 'Respiratory': 0.092}
doctors:
Doctor_ID         Doctor_Name                  Department            Grade
   DOC001      Dr Funke Bello                  Cardiology       Consultant
   DOC002      Dr Amaka Dauda                  Cardiology       Consultant
   DOC003    Dr Tobiloba Igwe                  Cardiology       Consultant
   DOC004     Dr Chidinma Odu                  Cardiology Senior Registrar
   DOC005       Dr Ronke Etim                  Cardiology Senior Registrar
   DOC006     Dr Mustapha Ike               Endocrinology       Consultant
   DOC007      Dr Hadiza Uzor               Endocrinology       Consultant
   DOC008    Dr Adenike Okoye               Endocrinology       Consultant
   DOC009       Dr Bukola Uba               Endocrinology Senior Registrar
   DOC010     Dr Rotimi Okoro               Endocrinology Senior Registrar
   DOC011      Dr Segun Ekong               Endocrinology Senior Registrar
   DOC012      Dr Kayode Edet               Endocrinology Senior Registrar
   DOC013 Dr Temitope Oladipo               Endocrinology Senior Registrar
   DOC014    Dr Modupe Onwuka                    Oncology       Consultant
   DOC015    Dr Tochukwu Edet                    Oncology       Consultant
   DOC016   Dr Folake Adewale                    Oncology       Consultant
   DOC017   Dr Adesuwa Ukpabi                    Oncology Senior Registrar
   DOC018     Dr Efosa Yakubu                    Oncology Senior Registrar
   DOC019  Dr Omolara Okonkwo Orthopaedics & Rheumatology       Consultant
   DOC020    Dr Temitope Kanu Orthopaedics & Rheumatology       Consultant
   DOC021     Dr Ikenna Yusuf Orthopaedics & Rheumatology       Consultant
   DOC022    Dr Segun Nwankwo Orthopaedics & Rheumatology Senior Registrar
   DOC023      Dr Bello Akpan Orthopaedics & Rheumatology Senior Registrar
   DOC024      Dr Ijeoma Edet                 Respiratory       Consultant
   DOC025   Dr Eniola Okonkwo                 Respiratory       Consultant
   DOC026    Dr Ebere Adesina                 Respiratory       Consultant
   DOC027    Dr Haruna Inyang                 Respiratory Senior Registrar
   DOC028 Dr Adesuwa Suleiman                 Respiratory Senior Registrar
census 2023 days 365  mean occ 0.8953  max 0.9808  min 0.7885  staffed [520]
sample admissions:
 Admission_ID       Patient_Name  Age Gender Medical_Condition                  Department Admission_Type Admission_Date Discharge_Date  Length_of_Stay          Ward                              HMO  Billing_Amount Readmit_30d
ADM2023-40124     Ahmed Y. Oseni   23   Male           Obesity               Endocrinology       Elective     2023-01-01     2023-01-18              17     Ruby Ward               AXA Mansard Health       530732.54          No
ADM2023-40125      Musa C. Lawal   80   Male         Arthritis Orthopaedics & Rheumatology       Elective     2023-01-01     2023-01-17              16 Sapphire Ward NHIA (National Health Insurance)      1846591.57          No
ADM2023-40126 Blessing J. Inyang   25 Female      Hypertension                  Cardiology         Urgent     2023-01-01     2023-01-07               6     Ruby Ward               AXA Mansard Health       198978.76          No
ADM2023-40127  Grace N. Olaniyan   27 Female         Arthritis Orthopaedics & Rheumatology         Urgent     2023-01-01     2023-01-06               5  Emerald Ward                       Hygeia HMO      1670235.58          No

## EkoSuites Hotel Lagos (36,275 bookings 2023-2024)
status: {'Honoured': 24390, 'Cancelled': 11885}  cancel rate: 0.3276
room: {'Standard': 28130, 'Deluxe': 6057, 'Family Suite': 966, 'Standard Twin': 692, 'Executive': 265, 'Presidential Suite': 158, 'Business': 7}
meal: {'Bed & Breakfast': 27835, 'Room Only': 5130, 'Half Board': 3305, 'Full Board': 5}
segment: {'Online Travel Agent': 23214, 'Walk-in / Travel Agent': 10528, 'Corporate': 2017, 'Complimentary': 391, 'Airline Crew': 125}
avg rate: 103424  median: 99450.0  lead time mean: 85.2  median: 57.0  nights mean: 3.02  repeat guests: 930
honoured value sum: 7,051,104,670  by arrival year: {2023: '1,382,249,280', 2024: '5,668,855,390'}
cancel rate by lead band: {Interval(-1, 7, closed='right'): 0.089, Interval(7, 30, closed='right'): 0.189, Interval(30, 90, closed='right'): 0.253, Interval(90, 180, closed='right'): 0.449, Interval(180, 1000, closed='right'): 0.739}
cancel by segment: {'Airline Crew': 0.296, 'Complimentary': 0.0, 'Corporate': 0.109, 'Online Travel Agent': 0.365, 'Walk-in / Travel Agent': 0.299}

## ETEH Talent Salary Survey (1,781 respondents)
salary median 11000000.0  mean 11344180  min 2500000.0  max 25000000.0
job family: {'Software Engineering': 548, 'Marketing': 228, 'Engineering': 224, 'Data & Analytics': 167, 'Sales': 135, 'Human Resources': 133, 'Other Business': 78, 'Product': 75, 'Finance': 73, 'Operations': 52, 'Research': 32, 'Project Management': 22, 'Design': 14}
seniority: {'Senior': 766, 'Mid': 737, 'Junior': 156, 'Executive': 122}
education: {"Bachelor's": 764, "Master's": 568, 'PhD': 339, 'High School': 110}
gender: {'Male': 963, 'Female': 811, 'Other': 7}
median by education: {"Bachelor's": 8000000.0, 'High School': 3500000.0, "Master's": 12713100.0, 'PhD': 17000000.0}
median by seniority: {'Executive': 16500000.0, 'Junior': 4000000.0, 'Mid': 9450200.0, 'Senior': 13500000.0}
median by family: {'Data & Analytics': 15000000.0, 'Design': 5000000.0, 'Engineering': 13831700.0, 'Finance': 9500000.0, 'Human Resources': 9000000.0, 'Marketing': 9750000.0, 'Operations': 10050000.0, 'Other Business': 6000000.0, 'Product': 12000000.0, 'Project Management': 12000000.0, 'Research': 16000000.0, 'Sales': 4000000.0, 'Software Engineering': 12000000.0}
corr exp-salary: 0.819
data analyst rows: {'count': 57.0, 'mean': 11666667.0, 'std': 3051834.0, 'min': 5000000.0, '25%': 9500000.0, '50%': 11500000.0, '75%': 13500000.0, 'max': 19500000.0}

## Lifecare Diabetes Screening (100,000)
diabetes: {'No': 91500, 'Yes': 8500}  smoking: {'Not Recorded': 35816, 'Never': 35095, 'Former': 9352, 'Current': 9286, 'Not Current': 6447, 'Ever': 4004}
             Age        BMI  HbA1c_Level  Blood_Glucose_mg_dL
count  100000.00  100000.00    100000.00            100000.00
mean       41.89      27.32         5.53               138.06
std        22.52       6.64         1.07                40.71
min         0.08      10.01         3.50                80.00
25%        24.00      23.63         4.80               100.00
50%        43.00      27.32         5.80               140.00
75%        60.00      29.58         6.20               159.00
max        80.00      95.69         9.00               300.00
diabetes rate by HbA1c>=6.5: 0.25  hypertension yes: 7485

## Crestview Digital Wellbeing (705)
platform: {'Instagram': 249, 'TikTok': 154, 'Facebook': 123, 'WhatsApp': 54, 'Twitter': 30, 'LinkedIn': 21, 'WeChat': 15, 'Snapchat': 13, 'LINE': 12, 'KakaoTalk': 12, 'VKontakte': 12, 'YouTube': 10}  usage mean 4.92  sleep mean 6.87  corr usage-sleep -0.791
