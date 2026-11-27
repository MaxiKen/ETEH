"""Shared paths, constants and helpers for the ETEH bootcamp build pipeline.

Every script in _build/ is deterministic (fixed random seeds) so the whole pack
can be regenerated from the raw files in the repository root with:

    python3 _build/01_build_datasets.py
    python3 _build/02_build_section_examples.py
    python3 _build/03_build_assignments.py
    python3 _build/04_build_projects.py
    python3 _build/05_update_slides.py
"""
from __future__ import annotations

import hashlib
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT  # raw source files sit in the repository root
PACK = ROOT / "ETEH_Bootcamp_Pack"
SEC = PACK / "01_Section_Examples"
ASG = PACK / "02_Assignments"
IPR = PACK / "03_Instructor_Projects"
SPR = PACK / "04_Student_Projects"
CORE = SEC / "00_Core_Datasets"
W2D = SEC / "Week2_Data_Files"
W3D = SEC / "Week3_Data_Files"
STAGE = ROOT / "_build" / "_stage"   # intermediate parquet/pickle cache (git-ignored)

for p in (PACK, SEC, ASG, IPR, SPR, CORE, W2D, W3D, STAGE):
    p.mkdir(parents=True, exist_ok=True)

SEED = 20261127
FONT_DIR = Path(__import__("matplotlib").__file__).parent / "mpl-data" / "fonts" / "ttf"

BRAND = "ETEH Ltd  ·  Eaglet2EagleHub Limited"
NAVY = "#0B2545"
TEAL = "#13807A"
GOLD = "#E0A526"
LIGHT = "#EEF4F8"
RED = "#B23A48"
GREEN = "#2E7D32"


def h(text: str, mod: int) -> int:
    """Stable hash -> int in [0, mod). Used for deterministic assignments."""
    return int(hashlib.md5(str(text).encode("utf-8")).hexdigest(), 16) % mod


# --------------------------------------------------------------------------------------
# Nigerian reference data
# --------------------------------------------------------------------------------------
CITY_STATE = {
    "Aba": "Abia", "Abuja": "FCT", "Ado Ekiti": "Ekiti", "Akure": "Ondo", "Bauchi": "Bauchi",
    "Benin City": "Edo", "Calabar": "Cross River", "Enugu": "Enugu", "Ibadan": "Oyo", "Ife": "Osun",
    "Ikorodu": "Lagos", "Ilorin": "Kwara", "Jos": "Plateau", "Kaduna": "Kaduna", "Kano": "Kano",
    "Lagos": "Lagos", "Maiduguri": "Borno", "Makurdi": "Benue", "Minna": "Niger", "Ogbomoso": "Oyo",
    "Onitsha": "Anambra", "Oshogbo": "Osun", "Owerri": "Imo", "Port Harcourt": "Rivers",
    "Sokoto": "Sokoto", "Umuahia": "Abia", "Uyo": "Akwa Ibom", "Warri": "Delta", "Yenagoa": "Bayelsa",
    "Zaria": "Kaduna",
}
STATE_ZONE = {
    "Lagos": "South-West", "Oyo": "South-West", "Osun": "South-West", "Ondo": "South-West",
    "Ekiti": "South-West", "Ogun": "South-West",
    "Abia": "South-East", "Enugu": "South-East", "Anambra": "South-East", "Imo": "South-East",
    "Ebonyi": "South-East",
    "Rivers": "South-South", "Edo": "South-South", "Delta": "South-South", "Cross River": "South-South",
    "Akwa Ibom": "South-South", "Bayelsa": "South-South",
    "FCT": "North-Central", "Plateau": "North-Central", "Kwara": "North-Central", "Benue": "North-Central",
    "Niger": "North-Central", "Kogi": "North-Central", "Nasarawa": "North-Central",
    "Kano": "North-West", "Kaduna": "North-West", "Sokoto": "North-West", "Katsina": "North-West",
    "Bauchi": "North-East", "Borno": "North-East", "Adamawa": "North-East", "Gombe": "North-East",
}

FEMALE_FIRST = [
    "Adaeze", "Adaobi", "Adebola", "Adenike", "Adesuwa", "Aisha", "Amaka", "Amara", "Amina", "Anuli",
    "Asabe", "Ayomide", "Bilkisu", "Blessing", "Bukola", "Chiamaka", "Chidinma", "Chinwe", "Chioma",
    "Damilola", "Ebere", "Ejiro", "Esther", "Eniola", "Fatima", "Folake", "Funke", "Funmilayo", "Grace",
    "Hadiza", "Halima", "Hauwa", "Ifeoma", "Ijeoma", "Itohan", "Jumoke", "Kemi", "Khadija", "Ladi",
    "Lami", "Mariam", "Modupe", "Ngozi", "Nkechi", "Nkiru", "Nneka", "Obiageli", "Oluchi", "Omolara",
    "Onyinye", "Osas", "Rashida", "Ronke", "Safiya", "Salamatu", "Simisola", "Temitope", "Titilayo",
    "Uchenna", "Uduak", "Yetunde", "Zainab", "Zuwaira", "Abimbola", "Ifunanya",
]
MALE_FIRST = [
    "Abdul", "Abdullahi", "Abubakar", "Adamu", "Adebayo", "Ahmed", "Aliyu", "Babatunde", "Bala",
    "Bello", "Chibuzor", "Chidi", "Chinedu", "Chukwuemeka", "Damola", "Daniel", "Ebuka", "Efosa",
    "Ekene", "Emeka", "Emmanuel", "Femi", "Gbenga", "Godwin", "Haruna", "Ibrahim", "Idris", "Ifeanyi",
    "Ikenna", "Isah", "Jide", "Kayode", "Kelechi", "Kunle", "Lanre", "Musa", "Mustapha", "Nnamdi",
    "Obinna", "Ochuko", "Okey", "Olumide", "Osagie", "Rotimi", "Sani", "Segun", "Shehu", "Sunday",
    "Tayo", "Tochukwu", "Tunde", "Uche", "Umar", "Usman", "Victor", "Wale", "Yakubu", "Yusuf",
    "Zubairu", "Ademola", "Chijioke", "Ikechukwu", "Oluwaseun", "Tobiloba", "Uzoma",
]
SURNAMES = [
    "Abubakar", "Adebayo", "Adeleke", "Adesina", "Adeyemi", "Agbaje", "Ajayi", "Akande", "Akinola",
    "Aliyu", "Anyanwu", "Attah", "Azikiwe", "Bakare", "Balogun", "Bankole", "Bello", "Chukwu",
    "Danjuma", "Dauda", "Edet", "Effiong", "Egwu", "Ekong", "Ekwensi", "Eze", "Ezeani", "Fashola",
    "Garba", "Ibe", "Ibrahim", "Idowu", "Igwe", "Ike", "Iwu", "Jibril", "Kalu", "Kanu", "Lawal",
    "Madu", "Mbah", "Mohammed", "Musa", "Nnaji", "Nnamdi", "Nwachukwu", "Nwankwo", "Nwosu", "Obi",
    "Odili", "Odu", "Ogbonna", "Ogundipe", "Ogunlesi", "Ojo", "Okafor", "Okeke", "Okonkwo", "Okoro",
    "Okoye", "Oladipo", "Olaniyan", "Olawale", "Onuoha", "Onyeka", "Oyelaran", "Salami", "Sanni",
    "Shagari", "Suleiman", "Udoh", "Ugwu", "Umeh", "Usman", "Uzor", "Yakubu", "Yusuf", "Adewale",
    "Akpan", "Amadi", "Ayodele", "Babangida", "Chime", "Dike", "Ekpo", "Etim", "Ezenwa", "Ibekwe",
    "Inyang", "Isiaka", "Mustapha", "Nduka", "Obasi", "Obiora", "Ogbu", "Olatunji", "Onwuka",
    "Orji", "Osagie", "Oseni", "Raji", "Tijani", "Uba", "Ukpabi", "Williams", "Abiodun", "Okpara",
]
MIDDLE_INITIALS = list("ABCDEFGIJKLMNOPRSTUWY")


def naira(x: float) -> str:
    return f"₦{x:,.0f}"
