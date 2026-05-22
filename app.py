import streamlit as st
import pandas as pd
import altair as alt
import re
from datetime import datetime, timedelta
from dateutil import parser

# ---------------------------
# Region Mapping Dictionary
# ---------------------------
DEFAULT_REGION_MAP = {
    "dsm": "Dar es Salaam","dar": "Dar es Salaam","znz": "Zanzibar","moro": "Morogoro",
    "mwz": "Mwanza","dom": "Dodoma","iringa": "Iringa","arusha": "Arusha","moshi": "Kilimanjaro",
    "pwani": "Pwani","tabora": "Tabora","mbeya": "Mbeya","tanga": "Tanga","kilimanjaro": "Kilimanjaro",
    "kigoma": "Kigoma","coast": "Pwani"
}

# ---------------------------
# Cleaning Functions (Spreadsheet)
# ---------------------------
def clean_name(name):
    if not name or str(name).strip() == "": return None
    return str(name).strip().title()

def clean_phone(phone):
    if not phone or str(phone).lower() in ["missing","invalid","none","null"]: return None
    digits = re.sub(r"\D","",str(phone))
    if digits.startswith("255"): return f"+{digits}"
    elif digits.startswith("0"): return f"+255{digits[1:]}"
    return f"+{digits}"

def clean_currency(value,user_currency="USD"):
    if not value or str(value).strip().lower() in ["-","null","none"]: return None
    cleaned = re.sub(r"[^\d.]","",str(value))
    try: return f"{float(cleaned)} {user_currency}"
    except: return None

def clean_date(date_str,today_date=None):
    if not date_str or str(date_str).strip().lower() in ["null","invalid date"]: return None
    today_date = today_date or datetime.today()
    ds = str(date_str).strip().lower()
    if ds=="today": return today_date.strftime("%Y-%m-%d")
    if ds=="yesterday": return (today_date-timedelta(days=1)).strftime("%Y-%m-%d")
    try: return parser.parse(date_str,dayfirst=True).strftime("%Y-%m-%d")
    except: return None

def clean_region(region,mapping_dict):
    if not region or str(region).strip().lower() in ["unknown","missing","null"]: return None
    region_clean = str(region).strip().lower()
    return mapping_dict.get(region_clean,region_clean.title())

def clean_record(record,user_currency="USD",mapping_dict=DEFAULT_REGION_MAP):
    return {
        "Name":clean_name(record.get("Name")),
        "Phone":clean_phone(record.get("Phone")),
        "Sales":clean_currency(record.get("Sales"),user_currency),
        "Date":clean_date(record.get("Date")),
        "Region":clean_region(record.get("Region"),mapping_dict)
    }

def deduplicate_records(records,user_currency="USD",mapping_dict=DEFAULT_REGION_MAP):
    cleaned_records=[clean_record(r,user_currency,mapping_dict) for r in records]
    deduped={}
    for rec in cleaned_records:
        key=(rec["Name"],rec["Phone"])
        if key not in deduped: deduped[key]=rec
        else:
            existing=deduped[key]
            if rec["Date"] and existing["Date"]:
                try:
                    rec_date=datetime.strptime(rec["Date"],"%Y-%m-%d")
                    existing_date=datetime.strptime(existing["Date"],"%Y-%m-%d")
                    if rec_date>existing_date: deduped[key]=rec; continue
                except: pass
            try:
                rec_sales=float(str(rec["Sales"]).split()[0]) if rec["Sales"] else 0
                existing_sales=float(str(existing["Sales"]).split()[0]) if existing["Sales"] else 0
                if rec_sales>existing_sales: deduped[key]=rec
            except: pass
    return list(deduped.values())

# ---------------------------
# Agricultural Data & Functions
# ---------------------------
# (all the crop calendars, spacing, yield, rotation, irrigation, fertigation, pest forecasts,
# chemical rotation, biological control, post-harvest, value addition, loan options, export guidelines
# plus calculate_yield and calculate_financials functions — from your second code block)

# ---------------------------
# Streamlit Multi-Tab App
# ---------------------------
st.set_page_config(page_title="Universal Cleaner", layout="wide")
st.title("🌍 Universal Cleaner")

tab1, tab2, tab3, tab4 = st.tabs([
    "Spreadsheet Cleaner",
    "Letters Cleaner",
    "Business Proposals",
    "Agricultural Manuals"
])

# Tab 1: Spreadsheet Cleaner
with tab1:
    st.header("📊 Clean Messy Spreadsheets")
    user_currency = st.selectbox("Select output currency:", ["USD","TZS","EUR","GBP","KES","UGX"])
    uploaded_mapping = st.file_uploader("Upload custom region mapping CSV (optional)", type=["csv"])
    mapping_dict = DEFAULT_REGION_MAP
    if uploaded_mapping:
        df_map = pd.read_csv(uploaded_mapping)
        mapping_dict = {row["abbreviation"].lower(): row["full_name"] for _, row in df_map.iterrows()}

    uploaded_file = st.file_uploader("Upload your spreadsheet (CSV or XLSX)", type=["csv","xlsx"])
    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        df.columns = [c.strip().title() for c in df.columns]
        records = df.to_dict(orient="records")
        cleaned = deduplicate_records(records,user_currency,mapping_dict)
        st.subheader("✅ Cleaned & Deduplicated Data")
        st.dataframe(pd.DataFrame(cleaned))
        cleaned_df = pd.DataFrame(cleaned)
        st.download_button("Download Cleaned Data", cleaned_df.to_csv(index=False).encode("utf-8"), "cleaned_data.csv", "text/csv")

# Tab 2: Letters Cleaner
with tab2:
    st.header("✉️ Formal & Informal Letters")
    text_input = st.text_area("Paste your letter text here:")
    if st.button("Clean Letter"):
        if text_input.strip():
            cleaned_text = " ".join(text_input.split())
            st.subheader("✅ Cleaned Letter")
            st.write(cleaned_text)

# Tab 3: Business Proposals
with tab3:
    st.header("📑 Business Proposals")
    proposal_input = st.text_area("Paste your business proposal here:")
    if st.button("Clean Proposal"):
        if proposal_input.strip():
            cleaned_proposal = " ".join(proposal_input.split())
            st.subheader("✅ Cleaned Proposal")
            st.write(cleaned_proposal)

# Tab 4: Agricultural Manuals
with tab4:
    st.header("🌱 Agricultural Planning & Manuals")

    crop_selection = st.multiselect("Select crops", list(CROP_CALENDAR.keys()))
    farm_location = st.text_input("Farm Location (Region/District)")
    planting_date = st.date_input("Planting Date")
    area_size = st.number_input("Area Size (acres)", min_value=0.1, value=1.0)
    market_price = st.number_input("Market Price (TZS per kg)", min_value=100, value=1500)
    spacing_choice = st.selectbox("Spacing Method", ["Row Planting","Diagonal Planting"])
    cost_per_acre = st.number_input("Cost per Acre (TZS)", min_value=1000, value=500000)

    if st.button("Generate Manual"):
        manual_sections = {
            "Crop Calendar": CROP_CALENDAR.get(crop_selection[0],""),
            "Plant Spacing": PLANT_SPACING.get(crop_selection[0],{}).get(spacing_choice,""),
            "Rotation Planning": ROTATION_GUIDELINES.get(crop_selection[0],""),
            "Irrigation Scheduling": IRRIGATION_GUIDELINES.get(crop_selection[0],""),
            "Fertigation Schedule": FERTIGATION_GUIDELINES.get(crop_selection[0],""),
            "Pest/Disease Forecast": PEST_FORECAST.get(crop_selection[0],""),
            "Chemical Rotation": CHEMICAL_ROTATION.get(crop_selection[0],""),
            "Biological Control": BIOLOGICAL_CONTROL.get(crop_selection[0],""),
            "Post-Harvest Handling": POST_HARVEST_GUIDELINES.get(crop_selection[0],""),
            "Value Addition": VALUE_ADDITION.get(crop_selection[0],""),
            "Loan/Credit Planning": LOAN_OPTIONS,
            "Export Market Integration": EXPORT_GUIDELINES.get(crop_selection[0],"")
        }
        st.subheader("✅ Structured Agricultural Manual")
        for section, content in manual_sections.items():
            st.markdown(f"### {section}")
            st.write(content if content else "Not provided")
