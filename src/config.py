"""
Configuration file — all sheet IDs, tab names, and business rules live here.
When your sheet structure changes, you ONLY edit this file.
"""

# ============================================================
# GOOGLE SHEET IDs (extracted from the URLs you shared)
# ============================================================

REVENUE_INDIA_SHEET = {
    "id": "1BRLYfnpYq2E9z5HMxUubjOxQRYVYP9pcA-VkycxOkP0",
    "name": "Client_Wise_Revenue_India",
    "worksheet": "Clientwise Billing Rev - India",
}

PAYROLL_SHEET = {
    "id": "1rE5q_gUd3nm6p1AG_0ow7UiMG8tikN8gCRHhQvrGnfg",
    "name": "Payroll_FY2025-26",
    "worksheet": "Payroll_Data",
}

OVERHEAD_SHEET = {
    "id": "1V0snKP9PPC2EDQbTJaBt1QRBnTiW9_qOTvNuovKx4bs",
    "name": "Overhead_FY2025_26",
    "worksheet": "Copy of Overheads",
}

REVENUE_ROW_SHEET = {
    "id": "1-WCwVHX7iZyzw9L9zpQZ57wV4_gNHj1NxKWzsu2Hpi8",
    "name": "Client_Wise_Revenue_ROW",
    "worksheet": "Clientwise Billing Rev - ROW",
}

OUTPUT_SHEET = {
    "id": "1r_3aT4MfS7DJvljT1kBLBqY4M7EYH3Rk_jeNE2KHhQI",
    "name": "Metrics Dashboard",
    "worksheet": "Metrics",
}

# ============================================================
# THE 6 REVENUE / COST CATEGORIES
# ============================================================
CATEGORIES = [
    "SaaS_India",
    "MS_India",
    "DSP_India",
    "SaaS_ROW",
    "MS_ROW",
    "DSP_ROW",
]

# ============================================================
# FY MONTHS (Apr 2025 - Mar 2026)
# ============================================================
FY_MONTHS = [
    "Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25",
    "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26",
]

# Mapping from revenue sheet month headers to a standard key
MONTH_MAP = {
    "Apr-25": "Apr-25", "May-25": "May-25", "Jun-25": "Jun-25",
    "Jul-25": "Jul-25", "Aug-25": "Aug-25", "Sep-25": "Sep-25",
    "Oct-25": "Oct-25", "Nov-25": "Nov-25", "Dec-25": "Dec-25",
    "Jan-26": "Jan-26", "Feb-26": "Feb-26", "Mar-26": "Mar-26",
}

# Payroll month mapping (Months column -> standard key)
PAYROLL_MONTH_MAP = {
    ("Apr", 2025): "Apr-25", ("May", 2025): "May-25", ("Jun", 2025): "Jun-25",
    ("Jul", 2025): "Jul-25", ("Aug", 2025): "Aug-25", ("Sep", 2025): "Sep-25",
    ("Oct", 2025): "Oct-25", ("Nov", 2025): "Nov-25", ("Dec", 2025): "Dec-25",
    ("Jan", 2026): "Jan-26", ("Feb", 2026): "Feb-26", ("Mar", 2026): "Mar-26",
}

# ============================================================
# TAGGING -> CATEGORY MAPPING (Revenue Sheets)
# ============================================================
TAGGING_TO_CATEGORY = {
    "SaaS - India": "SaaS_India",
    "Managed Service - India": "MS_India",
    "DSP - India": "DSP_India",
    "SaaS - ROW": "SaaS_ROW",
    "Managed Service - ROW": "MS_ROW",
    "DSP - ROW": "DSP_ROW",
}

# ============================================================
# PAYROLL ALLOCATION RULES
# ============================================================

# Rule 1: Anjum Ara — 70% DSP-ROW, 30% DSP-India (fixed split)
PAYROLL_FIXED_SPLIT = {
    "Anjum Ara": {
        "DSP_ROW": 0.70,
        "DSP_India": 0.30,
    },
}

# Rule 2: Revenue-based split into specific categories
# These employees' salaries are split proportionally to revenue in the listed categories
PAYROLL_REVENUE_SPLIT = {
    # Rule 2: Nishtha Chawla & Surya Ritolia → SaaS-India, MS-India
    "Nishtha Chawla": ["SaaS_India", "MS_India"],
    "Surya Ritolia": ["SaaS_India", "MS_India"],

    # Rule 3: US team → SaaS-ROW, MS-ROW, DSP-ROW
    "Anush Shetty": ["SaaS_ROW", "MS_ROW", "DSP_ROW"],
    "Luke Mcginnis": ["SaaS_ROW", "MS_ROW", "DSP_ROW"],
    "Lakshay Kapoor - US": ["SaaS_ROW", "MS_ROW", "DSP_ROW"],
    "Andrew Stevenson": ["SaaS_ROW", "MS_ROW", "DSP_ROW"],
    "Arya Chareonlarp": ["SaaS_ROW", "MS_ROW", "DSP_ROW"],

    # Rule 4: Rohit Aggarwal → all 6 categories
    "Rohit Aggarwal": ["SaaS_India", "MS_India", "DSP_India", "SaaS_ROW", "MS_ROW", "DSP_ROW"],

    # Rule 5: Renous Extenserve → SaaS-India, MS-India, DSP-India
    "Renous Extenserve Private Limited": ["SaaS_India", "MS_India", "DSP_India"],
}

# ============================================================
# PAYROLL DIRECT MAPPING (everyone else)
# Maps (Category_as_per_Meher, Region) to our 6 categories
# ============================================================
PAYROLL_DIRECT_MAP = {
    ("DSP", "India"): "DSP_India",
    ("DSP", "ROW"): "DSP_ROW",
    ("SaaS", "India"): "SaaS_India",
    ("SaaS", "ROW"): "SaaS_ROW",
    ("Managed Service", "India"): "MS_India",
    ("Managed Service", "ROW"): "MS_ROW",
    ("Neon", "India"): "MS_India",
    ("Neon", "ROW"): "MS_ROW",
    ("Hector", "India"): "SaaS_India",
    ("Hector", "ROW"): "SaaS_ROW",
    ("One Neon", "India"): "SaaS_India",  # fallback
    ("One Neon", "ROW"): "SaaS_ROW",      # fallback
}