"""Week 4 edition-3: week summary lines and glossary."""

SUMMARY = {
    "4.1": "Summarise a column's centre, spread and shape, flag outliers and use the ToolPak's Descriptive Statistics.",
    "4.2": "Read distributions, run a Pareto analysis, rank with percentiles and standardise with z-scores.",
    "4.3": "Measure association with CORREL, build a matrix and avoid the causation traps.",
    "4.4": "Fit and read a simple regression: slope, intercept, R-squared, residuals and honest predictions.",
    "4.5": "Find trend and seasonality, forecast with moving averages and FORECAST.ETS, and measure the error.",
    "4.6": "Answer what-if questions with Goal Seek, data tables, scenarios and Solver, and document them.",
    "4.7": "Turn a finding into a decision with a headline, a three-act structure and honest evidence.",
    "4.8": "Write a structured report with an executive summary, clean tables and a checked PDF.",
    "4.9": "Build and deliver a deck with one message per slide and charts made for projection.",
    "4.10": "Audit an inherited model, protect a workbook and automate repetitive steps safely.",
    "4.11": "Use AI co-pilots with structured prompts and schema prompts, within the rules on data.",
    "4.12": "Verify AI output with independent checks, control totals, provenance and sign-off.",
    "4.13": "Handle personal and health data lawfully: minimise, anonymise, check bias and govern.",
    "4.14": "Benchmark stores fairly, verify a PDF against the ledger and write a decision memo.",
}

GLOSSARY = sorted([
    ("Coefficient of variation", "Standard deviation divided by the mean: spread without units."),
    ("Composite score", "An average of standardised measures used to rank items on several measures at once."),
    ("Interquartile range (IQR)", "Q3 minus Q1: the spread of the middle 50% of values."),
    ("Median", "The middle value after sorting; not moved by extremes."),
    ("Normal distribution", "The symmetric bell curve: about 68% within 1 SD and 95% within 2 SD of the mean."),
    ("Outlier", "A value far from the rest, often flagged with the 1.5 x IQR rule."),
    ("Pareto analysis", "Ranking items by size to find how few make up most of the total."),
    ("Percentile", "The value below which a given share of the data falls."),
    ("Skewed data", "Lopsided data with a long tail of large (or small) values; the mean is pulled away from the median."),
    ("Standard deviation", "The typical distance of values from their mean, in the data's units."),
    ("Z-score", "How many standard deviations a value is from the mean."),
], key=lambda kv: kv[0].lower())
