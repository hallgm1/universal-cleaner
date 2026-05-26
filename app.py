import streamlit as st
import pandas as pd
import re
from datetime import datetime, date, timedelta
from dateutil import parser as dateparser
import anthropic

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="🌍 Universal Document Cleaner", layout="wide")

CURRENCIES = {
    "USD":"$","EUR":"€","GBP":"£","JPY":"¥","CNY":"¥","INR":"₹",
    "CAD":"CA$","AUD":"A$","CHF":"Fr","KRW":"₩","BRL":"R$","MXN":"MX$",
    "SGD":"S$","HKD":"HK$","SEK":"kr","NOK":"kr","DKK":"kr","NZD":"NZ$",
    "ZAR":"R","AED":"AED","SAR":"SAR","THB":"฿","MYR":"RM","IDR":"Rp",
    "PHP":"₱","PLN":"zł","TRY":"₺","RUB":"₽","CZK":"Kč","HUF":"Ft","TZS":"Tzs.",
}

DEFAULT_REGION_MAP = {
    "dar":"Dar es Salaam","dsm":"Dar es Salaam","dar es salaam":"Dar es Salaam",
    "znz":"Zanzibar","zanzibar":"Zanzibar","mwz":"Mwanza","mwanza":"Mwanza",
    "dom":"Dodoma","dodoma":"Dodoma","moro":"Morogoro","morogoro":"Morogoro",
    "moshi":"Kilimanjaro","kilimanjaro":"Kilimanjaro","arusha":"Arusha",
    "mbeya":"Mbeya","tanga":"Tanga","iringa":"Iringa","tabora":"Tabora",
    "kigoma":"Kigoma","pwani":"Pwani","coast":"Pwani","unknown":"","tanzania":"",
}

NULL_VALS = {"null","nan","none","n/a","na","missing","missing_date","invalid",
             "invalid date","unknown","today","yesterday","-","","undefined","0"}
MONTH_MAP = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
             "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
HONORIFICS = re.compile(r"^(dr|eng|prof|mr|mrs|ms|rev|capt)\.\s*", re.I)

TODAY = date.today()
YESTERDAY = TODAY - timedelta(days=1)

# ══════════════════════════════════════════════════════════════════════════════
# CLEANING PRIMITIVES
# ══════════════════════════════════════════════════════════════════════════════
def parse_year(y):
    y = int(y)
    return y + 2000 if y < 100 else y

def clean_name(v):
    if not v or str(v).strip().lower() in NULL_VALS:
        return None
    v = HONORIFICS.sub("", str(v).strip())
    v = re.sub(r"\.\s*$", "", v).strip()
    return v.title()

def normalize_phone(v):
    if not v or str(v).strip().lower() in NULL_VALS or str(v).strip().lower() == "invalid":
        return None
    digits = re.sub(r"\D", "", str(v))
    if digits.startswith("255") and len(digits) == 12:
        return f"+{digits}"
    if digits.startswith("0") and len(digits) == 10:
        return f"+255{digits[1:]}"
    if len(digits) >= 9:
        return f"+{digits}"
    return str(v).strip()

def normalize_sales(v, currency_code="USD"):
    if not v or str(v).strip().lower() in NULL_VALS:
        return None
    v = re.sub(r"^tsh\s*", "", str(v), flags=re.I)
    v = re.sub(r"^tzs\.?\s*", "", v, flags=re.I)
    v = re.sub(r"[$€£¥₹₩]", "", v)
    v = re.sub(r"[,\s]", "", v)
    try:
        f = float(v)
        sym = CURRENCIES.get(currency_code, currency_code)
        return f"{sym}{f:,.2f}"
    except:
        return None

def normalize_date(v):
    if not v or str(v).strip().lower() in NULL_VALS:
        return None
    ds = str(v).strip().lower()
    if ds == "today":     return TODAY.strftime("%Y-%m-%d")
    if ds == "yesterday": return YESTERDAY.strftime("%Y-%m-%d")
    v = str(v).strip()
    m = None
    # YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
    if (m := re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", v)):
        return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    # DD/MM/YY or DD/MM/YYYY
    if (m := re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", v)):
        return f"{parse_year(m.group(3))}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    # DD-MM-YY
    if (m := re.match(r"^(\d{1,2})-(\d{1,2})-(\d{2,4})$", v)):
        return f"{parse_year(m.group(3))}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    # DD-Mon-YYYY
    if (m := re.match(r"^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$", v)):
        mo = MONTH_MAP.get(m.group(2).lower())
        if mo: return f"{parse_year(m.group(3))}-{str(mo).zfill(2)}-{m.group(1).zfill(2)}"
    # Mon DD YYYY
    if (m := re.match(r"^([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})$", v)):
        mo = MONTH_MAP.get(m.group(1).lower())
        if mo: return f"{m.group(3)}-{str(mo).zfill(2)}-{m.group(2).zfill(2)}"
    # Fallback: dateutil
    try:
        return dateparser.parse(v, dayfirst=True).strftime("%Y-%m-%d")
    except:
        return None

def normalize_region(v, region_map=None):
    if region_map is None:
        region_map = DEFAULT_REGION_MAP
    if not v or str(v).strip().lower() in NULL_VALS:
        return None
    key = str(v).strip().lower()
    return region_map.get(key, str(v).strip().title())

# ── Smart deduplication: keeps most-recent date, fallback to highest sales ────
def smart_deduplicate(df, currency_code="USD"):
    name_col  = next((c for c in df.columns if re.search(r"name|client",  c, re.I)), None)
    phone_col = next((c for c in df.columns if re.search(r"phone|contact|num", c, re.I)), None)
    date_col  = next((c for c in df.columns if re.search(r"date|trans",   c, re.I)), None)
    sales_col = next((c for c in df.columns if re.search(r"sales|total|amount|pay", c, re.I)), None)

    if not name_col or not phone_col:
        return df, pd.DataFrame(columns=df.columns)

    seen = {}
    keep_idx = []
    dup_idx  = []

    for idx, row in df.iterrows():
        key = (str(row.get(name_col, "") or "").lower().strip(),
               re.sub(r"\D", "", str(row.get(phone_col, "") or "")))
        if not key[0] or not key[1]:
            keep_idx.append(idx)
            continue
        if key not in seen:
            seen[key] = idx
            keep_idx.append(idx)
        else:
            existing_idx = seen[key]
            # Prefer more-recent date
            if date_col:
                try:
                    d_new = datetime.strptime(str(df.at[idx, date_col] or ""), "%Y-%m-%d")
                    d_old = datetime.strptime(str(df.at[existing_idx, date_col] or ""), "%Y-%m-%d")
                    if d_new > d_old:
                        dup_idx.append(existing_idx)
                        keep_idx.remove(existing_idx)
                        keep_idx.append(idx)
                        seen[key] = idx
                        continue
                except:
                    pass
            # Fallback: prefer higher sales
            if sales_col:
                try:
                    s_new = float(re.sub(r"[^0-9.]", "", str(df.at[idx, sales_col] or "")) or 0)
                    s_old = float(re.sub(r"[^0-9.]", "", str(df.at[existing_idx, sales_col] or "")) or 0)
                    if s_new > s_old:
                        dup_idx.append(existing_idx)
                        keep_idx.remove(existing_idx)
                        keep_idx.append(idx)
                        seen[key] = idx
                        continue
                except:
                    pass
            dup_idx.append(idx)

    return df.loc[keep_idx].reset_index(drop=True), df.loc[dup_idx].reset_index(drop=True)

# ── Unpack CSV-in-column-A (corrupted files) ──────────────────────────────────
def unpack_col_a(df):
    """If entire dataset is crammed into column A as CSV strings, unpack it."""
    if df.shape[1] != 1 and not (df.shape[1] <= 3 and df.iloc[:,1:].isna().all().all()):
        return df, False
    import csv, io
    raw_lines = df.iloc[:, 0].dropna().astype(str).tolist()
    if not raw_lines or "," not in raw_lines[0]:
        return df, False
    try:
        reader = csv.reader(raw_lines)
        rows = list(reader)
        headers = [h.strip() for h in rows[0]]
        data    = [[c.strip() for c in row] for row in rows[1:]]
        result  = pd.DataFrame(data, columns=headers)
        return result, True
    except:
        return df, False

# ── Claude API call ───────────────────────────────────────────────────────────
def call_claude(system_prompt, user_prompt, api_key, max_tokens=2000):
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role":"user","content":user_prompt}],
    )
    return msg.content[0].text.strip()

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
  .main-title { font-size:2.2rem; font-weight:900; letter-spacing:-0.02em; }
  .sub-title  { font-size:0.85rem; letter-spacing:0.15em; text-transform:uppercase; opacity:0.6; }
  .stat-box   { background:#f0f4ff; border-left:4px solid #1565c0; padding:12px 16px; border-radius:4px; }
  .warn-box   { background:#fff8f0; border-left:4px solid #e65100; padding:12px 16px; border-radius:4px; }
  .good-box   { background:#f0fff4; border-left:4px solid #2e7d32; padding:12px 16px; border-radius:4px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🌍 Universal Document Cleaner</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Spreadsheet · Letters · Business Proposals · Agricultural Manuals</div>', unsafe_allow_html=True)
st.markdown("---")

# ── Global sidebar settings ───────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Global Settings")
    currency_code = st.selectbox(
        "💱 Active Currency",
        list(CURRENCIES.keys()),
        index=list(CURRENCIES.keys()).index("TZS"),
        format_func=lambda c: f"{c}  ({CURRENCIES[c]})"
    )
    currency_sym = CURRENCIES[currency_code]
    st.caption(f"All financial values will use **{currency_code}** ({currency_sym})")

    st.markdown("---")
    api_key = st.text_input("🔑 Anthropic API Key (for AI features)", type="password", placeholder="sk-ant-...")
    st.caption("Required for AI typo fix, Letters, Proposals, and Agri Manual tabs.")

    st.markdown("---")
    st.markdown("**Custom Region Map**")
    map_file = st.file_uploader("Upload CSV (abbreviation, full_name)", type=["csv"], key="region_map")
    region_map = dict(DEFAULT_REGION_MAP)
    if map_file:
        try:
            df_map = pd.read_csv(map_file)
            for _, row in df_map.iterrows():
                region_map[str(row.iloc[0]).strip().lower()] = str(row.iloc[1]).strip()
            st.success(f"✓ Loaded {len(df_map)} custom regions")
        except Exception as e:
            st.error(f"Map error: {e}")

    st.markdown("---")
    st.caption("Universal Document Cleaner · v3.0")

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Spreadsheet Cleaner",
    "✉️ Letters Cleaner",
    "📑 Business Proposals",
    "🌱 Agricultural Manuals",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — SPREADSHEET CLEANER
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("📊 Spreadsheet Cleaner")
    st.caption("Handles corrupted files, CSV-in-column-A, 10+ date formats, phone normalization, smart deduplication.")

    uploaded = st.file_uploader("Upload spreadsheet (CSV or XLSX)", type=["csv","xlsx"], key="sheet_upload")

    if uploaded:
        # Load
        try:
            if uploaded.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded, dtype=str)
            else:
                df_raw = pd.read_excel(uploaded, dtype=str)
        except Exception as e:
            st.error(f"Could not read file: {e}")
            st.stop()

        # Unpack CSV-in-col-A
        df_raw, was_unpacked = unpack_col_a(df_raw)
        if was_unpacked:
            st.info("⚡ CSV-in-column-A structure detected and unpacked automatically.")

        # Normalize column headers
        df_raw.columns = [str(c).strip() for c in df_raw.columns]

        st.markdown(f"**Loaded:** {df_raw.shape[0]} rows · {df_raw.shape[1]} columns")
        with st.expander("👁 Raw data preview", expanded=True):
            st.dataframe(df_raw.head(10), use_container_width=True)

        st.markdown("---")
        st.markdown("### ⚙️ Per-Column Cleaning Rules")
        st.caption("Rules are auto-detected from column names. Toggle to adjust.")

        RULE_OPTIONS = [
            "Trim whitespace", "Title Case", "Strip honorific prefixes",
            "Normalize phone → +255XXXXXXXXX", "Normalize date → YYYY-MM-DD",
            "Parse sales/currency", "Canonicalize region", "Null strings → empty",
            "Normalize gender/ethnicity", "Redact PII",
        ]

        col_rules = {}
        ai_cols   = []

        cols_per_row = 3
        headers = list(df_raw.columns)
        for i in range(0, len(headers), cols_per_row):
            row_cols = st.columns(cols_per_row)
            for j, col_name in enumerate(headers[i:i+cols_per_row]):
                with row_cols[j]:
                    h = col_name.lower()
                    defaults = ["Trim whitespace", "Null strings → empty"]
                    if re.search(r"name|client",  h): defaults += ["Strip honorific prefixes","Title Case"]
                    if re.search(r"date|trans|hire|exit", h): defaults += ["Normalize date → YYYY-MM-DD"]
                    if re.search(r"salary|sales|pay|total|amount", h): defaults += ["Parse sales/currency"]
                    if re.search(r"phone|contact|num", h): defaults += ["Normalize phone → +255XXXXXXXXX"]
                    if re.search(r"region|zone|city|location", h): defaults += ["Canonicalize region"]
                    if re.search(r"gender|eth", h): defaults += ["Normalize gender/ethnicity"]

                    selected = st.multiselect(
                        f"**{col_name}**",
                        RULE_OPTIONS,
                        default=[d for d in defaults if d in RULE_OPTIONS],
                        key=f"col_{i}_{j}_{col_name}",
                    )
                    col_rules[col_name] = selected
                    if api_key:
                        if st.checkbox(f"🤖 AI typo fix", key=f"ai_{i}_{j}_{col_name}"):
                            ai_cols.append(col_name)

        st.markdown("---")
        run_dedup = st.checkbox("🔁 Smart deduplication (keeps most-recent date / highest sales)", value=True)

        if st.button("▶ CLEAN SPREADSHEET", type="primary"):
            df_work = df_raw.copy()

            with st.spinner("Cleaning…"):
                for col_name, rules in col_rules.items():
                    if col_name not in df_work.columns:
                        continue
                    new_vals = []
                    for val in df_work[col_name]:
                        v = str(val).strip() if pd.notna(val) else ""
                        # Null check first
                        if "Null strings → empty" in rules and v.lower() in NULL_VALS:
                            new_vals.append(None); continue
                        if "Trim whitespace" in rules:
                            v = re.sub(r"\s+", " ", v).strip()
                        if "Strip honorific prefixes" in rules:
                            v = HONORIFICS.sub("", v).rstrip(".").strip()
                        if "Normalize phone → +255XXXXXXXXX" in rules:
                            new_vals.append(normalize_phone(v)); continue
                        if "Parse sales/currency" in rules:
                            new_vals.append(normalize_sales(v, currency_code)); continue
                        if "Normalize date → YYYY-MM-DD" in rules:
                            new_vals.append(normalize_date(v)); continue
                        if "Canonicalize region" in rules:
                            new_vals.append(normalize_region(v, region_map)); continue
                        if "Title Case" in rules:
                            v = v.title()
                        if "Normalize gender/ethnicity" in rules:
                            gmap = {"f":"Female","female":"Female","m":"Male","male":"Male",
                                    "cau":"Caucasian","caucasian":"Caucasian","white":"Caucasian",
                                    "black":"Black","asian":"Asian","latino":"Latino/Hispanic","hispanic":"Latino/Hispanic"}
                            v = re.sub(r"\b(\w+)\b", lambda x: gmap.get(x.group().lower(), x.group()), v)
                        if "Redact PII" in rules:
                            v = re.sub(r"\b[A-Z][a-z]+ [A-Z][a-z]+\b", "[REDACTED]", v)
                            v = re.sub(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", "[EMAIL]", v)
                        new_vals.append(v if v else None)
                    df_work[col_name] = new_vals

                # AI typo fix
                if api_key and ai_cols:
                    for col_name in ai_cols:
                        if col_name not in df_work.columns: continue
                        try:
                            vals = df_work[col_name].fillna("").tolist()
                            import json
                            result = call_claude(
                                "You are a data cleaning assistant. Return ONLY a JSON array of cleaned strings, same length and order, no markdown.",
                                f'Fix typos in column "{col_name}":\n{json.dumps(vals)}',
                                api_key
                            )
                            fixed = json.loads(result.replace("```json","").replace("```","").strip())
                            df_work[col_name] = fixed
                        except Exception as e:
                            st.warning(f"AI fix failed for '{col_name}': {e}")

                # Smart deduplication
                df_dups = pd.DataFrame()
                if run_dedup:
                    df_work, df_dups = smart_deduplicate(df_work, currency_code)

            # Results
            st.markdown("---")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Original rows",    df_raw.shape[0])
            c2.metric("Unique rows",      df_work.shape[0])
            c3.metric("Duplicates removed", len(df_dups))
            c4.metric("Columns cleaned",  len([c for c,r in col_rules.items() if r]))

            st.markdown("### ✅ Cleaned Data")
            st.dataframe(df_work, use_container_width=True)
            st.download_button(
                "⬇ Download Cleaned CSV",
                df_work.to_csv(index=False).encode("utf-8"),
                "cleaned_data.csv", "text/csv",
            )

            if not df_dups.empty:
                with st.expander(f"🔁 {len(df_dups)} duplicate rows removed (best record kept)"):
                    st.dataframe(df_dups, use_container_width=True)
                    st.download_button(
                        "⬇ Download Duplicates Log",
                        df_dups.to_csv(index=False).encode("utf-8"),
                        "duplicates_removed.csv","text/csv",
                    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — LETTERS CLEANER
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("✉️ Letters Cleaner")
    st.caption("AI-powered tone rewriting, PII redaction, grammar and structure fixing.")

    col_l, col_r = st.columns(2)
    with col_l:
        tone = st.selectbox("Select tone", [
            "Formal / Legal", "Friendly", "Diplomatic", "Concise", "Persuasive"
        ])
        redact_pii = st.checkbox("Redact personal information (names, emails, phones)")
        letter_input = st.text_area("Paste letter text here", height=380,
            placeholder="Paste your letter — formal, informal, business, personal…")

        if st.button("✦ CLEAN LETTER", type="primary"):
            if not letter_input.strip():
                st.warning("Please paste a let
