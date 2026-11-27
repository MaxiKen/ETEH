# Beginner revision of the five decks (v2): design and writing rules

Build: `python3 _build/06_revise_slides.py` (runs 05 first; v1 decks are kept in `_build/_stage/v1_decks`).
Preview PNGs: add `--preview 1.6.1 2.4.3` or `--preview V1` (→ `_build/_stage/preview/`).
Topic copy: one file per module, `_build/rev/t<vol>_<module>.py`, each with `TOPICS = [T(...), ...]` (see `_build/revlib.py`).

## Slide pair per topic
1. **Concept slide** (`tag · 1/2`):
   - IN SIMPLE TERMS (`what`, 15 pt).
   - THINK OF IT LIKE THIS (`like`, an everyday Nigerian analogy).
   - HOW TO DO IT (`how`, optional numbered steps).
   - REMEMBER (`tip`) and AVOID (`trap`). If these are omitted, the original deck's best-practice / watch-out text is used.
   - Cue at the bottom: "Worked example on the next slide »".
2. **Example slide** (`tag · 2/2`):
   - Green **WORKED EXAMPLE** badge with a tick (the end-of-group marker).
   - THE SITUATION (`who`, `case`).
   - STEP BY STEP (`steps`).
   - WHAT YOU SHOULD SEE (`table`).
   - WHAT THIS TELLS US (`so`).
   - TRY IT YOURSELF: the section workbook sheet by default, overridable with `tryit`.
   - Next-topic pointer.

## Writing rules
- Plain English for a complete beginner: short sentences, define every term the first time, no unexplained jargon.
- **Bold** with `**…**` only for the key term or the key click.
- Analogies come from everyday Nigerian life: market, danfo/BRT, NEPA bills, POS agents, jollof, ajo/esusu, WAEC/JAMB, owambe, mama-put, recharge cards.
- Worked examples use the bootcamp files. Every number must be true to the data (see FACTS.md, or check with pandas).
  - SuperMart Stores: `SuperMart_Stores_Sales_2022_2025.csv` (columns A Row_ID … R Sales, S Quantity, T Discount, U Profit; rows 2–9995) and `SuperMart_Stores_Directory.csv`.
  - SuperMart Online: 5 tables / `supermart_online.db`.
  - Crestview: `Crestview_Student_Performance.csv`. Grade: A ≥ 70, B 60–69, C 55–59. At risk: exam < 62 or attendance < 65%.
  - Lifecare: admissions, wards, doctors, census, `lifecare_hospital.db`.
  - Also: EkoSuites Hotel, the ETEH Talent Salary Survey, Diabetes Screening, and the Digital Wellbeing survey.
- Use ₦ for money. Use real Nigerian names, cities and products from the files.
- Approximate lengths: `what` ≤ 520 chars, `like` ≤ 230, `how` 3–4 items ≤ 90 chars, `case` ≤ 260, `steps` 3–4 items ≤ 95 chars, `table` ≤ 5 rows × ≤ 4 short columns, `so` ≤ 170.
  - The build measures text with Carlito (Calibri metrics) and prints OVERFLOW / CLIP / TABLE warnings, which must all be fixed.

## Progress
Tracked in `_build/rev/PROGRESS.md`.
