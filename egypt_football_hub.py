"""
Egyptian Football Intelligence Hub
===================================
Requirements:
    pip install streamlit pandas plotly openpyxl

Run:
    streamlit run egypt_football_hub.py

Place your Excel files in the same folder:
    - Egyptian_Football_Pyramid.xlsx
    - Starters_team_breakdown.xlsx
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import os

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Assem's Football Brainchild",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* Georgia serif throughout */
  html, body, [class*="css"], h1, h2, h3, h4, p, div, label, span,
  .stMarkdown, .stRadio, .stSelectbox, .stMultiSelect, .stTextInput,
  [data-testid="stSidebar"] * {
    font-family: Georgia, 'Times New Roman', serif !important;
  }
  /* Metric cards */
  .metric-card {
    background: #f0f4ff;
    border: 1px solid #c7d7f9;
    border-radius: 8px;
    padding: 20px 24px;
    margin-bottom: 12px;
  }
  .metric-value { font-size: 2rem; font-weight: 700; color: #1d4ed8; font-family: Georgia, serif !important; }
  .metric-label { font-size: 0.78rem; color: #4b5fa6; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.06em; }
  hr { border-color: #e5e7eb; }
  /* Square checkbox-style radio buttons */
  [data-testid="stRadio"] label div[data-testid="stMarkdownContainer"] p { font-family: Georgia, serif !important; }
  div[role="radiogroup"] label div:first-child { border-radius: 3px !important; }
  /* Sidebar nav items */
  [data-testid="stSidebar"] [data-testid="stRadio"] label {
    padding: 7px 10px;
    border-radius: 4px;
    font-size: 0.88rem;
    color: #1e3a5f;
  }
  [data-testid="stSidebar"] [data-testid="stRadio"] label:hover { background: #eff6ff; }
  /* Headings */
  h1, h2, h3 { color: #1e3a5f !important; }
  /* Sidebar section labels above specific radio items */
  [data-testid="stSidebar"] [data-testid="stRadio"] label:nth-child(1)::before {
    content: "EGYPT ANALYSIS";
    display: block;
    font-size: 0.65rem;
    font-weight: 700;
    color: #9ca3af;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 10px 0 4px;
    font-family: Georgia, serif;
    pointer-events: none;
  }
  [data-testid="stSidebar"] [data-testid="stRadio"] label:nth-child(7)::before {
    content: "AFRICA ANALYSIS";
    display: block;
    font-size: 0.65rem;
    font-weight: 700;
    color: #9ca3af;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 14px 0 4px;
    border-top: 1px solid #e5e7eb;
    margin-top: 6px;
    font-family: Georgia, serif;
    pointer-events: none;
  }
</style>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────────────────────
PYRAMID_FILE   = "Egyptian_Football_Pyramid.xlsx"
STARTERS_FILE  = "Starters_team_breakdown.xlsx"
CAF_FILE         = "caf_index_v3.xlsx"
URBAN_POP_FILE   = "egypt_populations_final.csv"

PLOTLY_THEME = dict(
    template="plotly_white",
)

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#ffffff",
    plot_bgcolor="#f0f4ff",
    font_color="#1e3a5f",
)

def apply_theme(fig):
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def gov_normalize(s):
    """Normalise governorate name spelling variants."""
    if not isinstance(s, str):
        return str(s) if s else ""
    s = s.strip()
    mapping = {
        "Dakahleya": "Dakahlia", "Dakhaleya": "Dakahlia", "Dakahleya ": "Dakahlia",
        "Gharbeya": "Gharbia",  "Gharbeya ": "Gharbia",
        "Sharqia":  "Sharkia",  "Sharqeya": "Sharkia",  "Sharqeya ": "Sharkia",
        "Monofeya": "Menoufia", "Monofiya": "Menoufia",
        "Qalyoubeya": "Qalyubia", "Qalyoubeya ": "Qalyubia",
        "Ismailia ": "Ismailia", "Ismailia": "Ismailia",
        "Asyut": "Assiut",
        "Cairo ": "Cairo", "Giza ": "Giza",
        "Aswan ": "Aswan", "Minya ": "Minya",
        "Alexandria ": "Alexandria",
        "Beheira ": "Beheira",
        "Red Sea ": "Red Sea",
        "Matrouh ": "Matrouh",
        "North Sinai ": "North Sinai",
        "South Sinai": "South Sinai",
        "Damietta ": "Damietta",
    }
    return mapping.get(s, s)

def parse_excel_date(val):
    """Convert Excel serial date or string date to year (int) or None."""
    if pd.isna(val) or val == "" or val == "-":
        return None
    if isinstance(val, (int, float)):
        try:
            d = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
            return d.year
        except Exception:
            return None
    s = str(val).strip()
    for fmt in ["%d/%b/%Y", "%d/%m/%Y", "%Y-%m-%d"]:
        try:
            return pd.to_datetime(s, format=fmt).year
        except Exception:
            continue
    try:
        return pd.to_datetime(s).year
    except Exception:
        return None

def parse_birth_month(val):
    """Return month name from date string or Excel serial."""
    if pd.isna(val) or val == "" or val == "-":
        return None
    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    if isinstance(val, str) and val.strip() in months:
        return val.strip()
    if isinstance(val, (int, float)):
        try:
            d = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
            return d.strftime("%B")
        except Exception:
            return None
    s = str(val).strip()
    for fmt in ["%d/%b/%Y", "%d/%m/%Y"]:
        try:
            return pd.to_datetime(s, format=fmt).strftime("%B")
        except Exception:
            continue
    try:
        return pd.to_datetime(s).strftime("%B")
    except Exception:
        return None

@st.cache_data
def load_pyramid():
    xl = pd.ExcelFile(PYRAMID_FILE)
    clubs = xl.parse("Clubs")
    clubs.columns = [c.strip() for c in clubs.columns]
    clubs["Category"] = clubs["Category"].str.strip()
    clubs["Governorate"] = clubs["Governorate"].apply(gov_normalize)
    clubs["Official District"] = clubs.get("Official District", pd.Series(dtype=str)).fillna("")
    clubs["Level"] = pd.to_numeric(clubs["Level"], errors="coerce")
    clubs["Date Established"] = pd.to_numeric(clubs["Date Established"], errors="coerce")

    govs = xl.parse("Governorates")
    govs.columns = [c.strip() for c in govs.columns]
    # Drop total row
    govs = govs[govs["Governorate"].notna() & (govs["Governorate"] != "Total")]
    govs["Governorate"] = govs["Governorate"].apply(gov_normalize)
    numeric_cols = ["Population","Area (km2)","Density (km2)","GDP (PPP)",
                    "GDP per capita","Score","Football Index","Overperformance Index"]
    for c in numeric_cols:
        if c in govs.columns:
            govs[c] = pd.to_numeric(govs[c], errors="coerce")
    # Merge urban population data (for scatterplot tab only — not used in overperformance metric)
    try:
        upop = pd.read_csv("egypt_populations_final.csv")
        upop = upop[upop["Governorate"] != "Total"].copy()
        upop["Governorate"] = upop["Governorate"].str.strip().str.replace("]","",regex=False)
        name_fix = {"Daqahlia":"Dakahlia","Qalyoubeua":"Qalyubia","Monofia":"Menoufia",
                    "Beheira":"Beheira","Assyut":"Assiut","Sharqia":"Sharkia"}
        upop["Governorate"] = upop["Governorate"].replace(name_fix)
        upop["Urban_Pct"] = (upop["Urban_Total"] / upop["Total_Population"] * 100).round(1)
        upop["Rural_Pct"] = (upop["Rural_Total"] / upop["Total_Population"] * 100).round(1)
        govs["Governorate"] = govs["Governorate"].str.strip()
        govs = govs.merge(upop[["Governorate","Urban_Total","Rural_Total","Urban_Pct","Rural_Pct"]], on="Governorate", how="left")
    except Exception:
        pass

    # District sheet: row 0 = title, row 1 = blank, row 2 = headers, row 3+ = data
    dist_raw = xl.parse("Official By District", header=None)
    dist_raw.columns = ["District","Population","Governorate","Area_km2","Density"]
    districts = dist_raw.iloc[3:].reset_index(drop=True)
    districts = districts[districts["District"].notna() & (districts["District"] != "District")]
    for c in ["Population","Area_km2","Density"]:
        districts[c] = pd.to_numeric(districts[c], errors="coerce")
    districts = districts.dropna(subset=["Population"])
    districts["Governorate"] = districts["Governorate"].apply(gov_normalize)

    players = xl.parse(xl.sheet_names[3])  # International Players sheet
    players.columns = [str(c).strip() for c in players.columns]
    # Columns expected: Name, DOB, District, Governorate, something
    players = players.dropna(subset=[players.columns[0]])
    col_names = list(players.columns)
    players = players.rename(columns={
        col_names[0]: "Name",
        col_names[1]: "DOB",
        col_names[2]: "District",
        col_names[3]: "Governorate",
    })
    players["Governorate"] = players["Governorate"].apply(gov_normalize)
    players["Birth_Year"]  = players["DOB"].apply(parse_excel_date)
    players["Birth_Month"] = players["DOB"].apply(parse_birth_month)

    return clubs, govs, districts, players

@st.cache_data
def load_caf():
    if not os.path.exists(CAF_FILE):
        return None
    xl = pd.ExcelFile(CAF_FILE)
    ov  = xl.parse("Overperformance",      header=1)
    fi  = xl.parse("Football Index",        header=1)
    ei  = xl.parse("Economic Index",        header=1)
    ped = xl.parse("Tournament Pedigree",   header=1)
    # Merge overperformance + football component scores + economic index
    fi_cols = ["Country","Football Index"] + [c for c in fi.columns
               if c in ["FIFA Men","FIFA Women","Assoc Points","Infrastructure",
                        "Youth Score","Export Vol","Export Quality","Pedigree"]]
    df = ov.merge(fi[[c for c in fi_cols if c in fi.columns]],
                  on="Country", how="left", suffixes=("","_fi"))
    df = df.merge(ei[["Country","Economic Index"]], on="Country", how="left", suffixes=("","_ei"))
    df = df.rename(columns={"Country":"COUNTRY"})
    for c in df.columns:
        if df[c].dtype == object and c not in ["COUNTRY","Zone","Classification"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # Clean pedigree sheet
    ped = ped.dropna(subset=["Country"])
    ped = ped[ped["Country"].astype(str).str.strip() != "Country"]
    ped = ped[~ped["Country"].astype(str).str.startswith("Colour")]
    for c in ["Pedigree Score","AFCON Score","WC Score"]:
        if c in ped.columns:
            ped[c] = pd.to_numeric(ped[c], errors="coerce")
    return df, ped

@st.cache_data
def load_urban():
    if not os.path.exists(URBAN_POP_FILE):
        return None
    df = pd.read_csv(URBAN_POP_FILE)
    df = df[df["Governorate"] != "Total"].copy()
    # Fix known name issues
    df["Governorate"] = df["Governorate"].str.strip().str.replace("]","",regex=False)
    name_fix = {
        "Daqahlia":"Dakahlia","Qalyoubeua":"Qalyubia","Monofia":"Menoufia",
        "Beheira":"Beheira","Assyut":"Assiut","Sharqia":"Sharkia",
    }
    df["Governorate"] = df["Governorate"].replace(name_fix)
    df["Urban_Pct"] = (df["Urban_Total"] / df["Total_Population"] * 100).round(1)
    return df

@st.cache_data
def load_starters():
    xl = pd.ExcelFile(STARTERS_FILE)

    # Header is on row index 2, data starts at row 3
    raw = xl.parse("Season by Season", header=None)
    headers = [str(v).strip() if pd.notna(v) else f"col_{j}"
               for j, v in enumerate(raw.iloc[2])]
    ssn = raw.iloc[3:].copy()
    ssn.columns = headers
    ssn = ssn.reset_index(drop=True)

    # SEASON column contains numeric codes like 8687, 8788 etc - rename and format
    ssn = ssn.rename(columns={"SEASON": "Season"})

    # Filter out non-season rows
    ssn = ssn[ssn["Season"].apply(
        lambda x: pd.notna(x) and str(x).strip() not in
                  ["", "nan", "Average per ssn", "Total", "SEASON"]
    )]

    # Format season labels e.g. 8687 -> 86/87, 20002001 -> 00/01
    def fmt_season(s):
        s = str(s).strip().replace(".0","")
        if len(s) == 4:
            return s[:2] + "/" + s[2:]
        elif len(s) == 8:
            return s[2:4] + "/" + s[6:]
        return s
    ssn["Season"] = ssn["Season"].apply(fmt_season)

    # Drop empty placeholder columns
    ssn = ssn.drop(columns=[c for c in ssn.columns if c.startswith("col_")], errors="ignore")

    gov_cols = [c for c in ssn.columns if c not in ["Season", "TOTAL"]]
    for c in gov_cols + (["TOTAL"] if "TOTAL" in ssn.columns else []):
        ssn[c] = pd.to_numeric(ssn[c], errors="coerce").fillna(0)

    return ssn, gov_cols

# ── Sidebar navigation ─────────────────────────────────────────────────────────
EGYPT_PAGES = [
    "Overview",
    "Pyramid Map",
    "Governorate Intelligence",
    "District Density",
    "Player Origins",
    "Starter Pipeline",
]
AFRICA_PAGES = [
    "CAF Overperformance Index",
    "CAF Tournament Pedigree",
]
PAGES = EGYPT_PAGES + AFRICA_PAGES

# Build grouped page list with separators for display
ALL_PAGES_DISPLAY = (
    ["— Egypt Analysis —"] + EGYPT_PAGES +
    ["— Africa Analysis —"] + AFRICA_PAGES
)
NAVIGABLE = EGYPT_PAGES + AFRICA_PAGES

with st.sidebar:
    st.markdown("""
    <div style='padding:14px 0 18px; border-bottom:1px solid #e5e7eb; margin-bottom:14px;'>
      <div style='font-size:1.15rem; font-weight:700; color:#1e3a5f; font-family:Georgia,serif;'>
        Assem's Football Brainchild
      </div>
      <div style='font-size:0.72rem; color:#6b7280; margin-top:4px; font-family:Georgia,serif;
                  letter-spacing:0.04em; text-transform:uppercase;'>
        Egypt · Africa · Intelligence
      </div>
    </div>
    """, unsafe_allow_html=True)

    selected = st.radio("Navigation", NAVIGABLE,
                        format_func=lambda x: x,
                        label_visibility="collapsed")

    # Section headers injected as disabled separators via markdown
    # (rendered above the radio via CSS trick — just use captions)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style='font-size:0.68rem; color:#9ca3af; padding:8px 0;
                border-top:1px solid #e5e7eb; font-family:Georgia,serif;'>
      Data: 1986–2026 · Manually compiled
    </div>
    """, unsafe_allow_html=True)

page = selected


# ── Load data ──────────────────────────────────────────────────────────────────
missing = [f for f in [PYRAMID_FILE, STARTERS_FILE] if not os.path.exists(f)]
if missing:
    st.error(f"⚠️ Missing file(s): {', '.join(missing)}\n\nPlace them in the same folder as this script.")
    st.stop()

clubs, govs, districts, players = load_pyramid()
ssn_df, gov_cols = load_starters()
caf_result = load_caf()
caf_df, ped_df = caf_result if caf_result else (None, None)
urban_df = load_urban()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
if page == PAGES[0]:
    # Hero header
    st.markdown("""
    <div style='background:linear-gradient(135deg,#1e3a5f 0%,#1d4ed8 100%);
                border-radius:12px; padding:32px 36px; margin-bottom:28px;'>
      <div style='font-size:2rem; font-weight:700; color:#ffffff; font-family:Georgia,serif;
                  margin-bottom:8px;'>Assem's Football Brainchild</div>
      <div style='font-size:1rem; color:#bfdbfe; font-family:Georgia,serif; max-width:640px;
                  line-height:1.6;'>
        A comprehensive football intelligence model covering <b style="color:#fff;">Egypt</b> —
        its pyramid, governorates, player origins and talent pipeline — and
        <b style="color:#fff;">Africa</b> — a 54-nation index measuring football output
        against economic potential.
      </div>
      <div style='margin-top:20px; display:flex; gap:24px; flex-wrap:wrap;'>
        <div style='background:rgba(255,255,255,0.12); border-radius:8px; padding:10px 18px;'>
          <div style='font-size:0.7rem; color:#93c5fd; text-transform:uppercase;
                      letter-spacing:0.08em; font-family:Georgia,serif;'>Egypt scope</div>
          <div style='font-size:0.95rem; color:#fff; font-family:Georgia,serif; margin-top:3px;'>
            Clubs · Governorates · Districts · Players · Starters
          </div>
        </div>
        <div style='background:rgba(255,255,255,0.12); border-radius:8px; padding:10px 18px;'>
          <div style='font-size:0.7rem; color:#93c5fd; text-transform:uppercase;
                      letter-spacing:0.08em; font-family:Georgia,serif;'>Africa scope</div>
          <div style='font-size:0.95rem; color:#fff; font-family:Georgia,serif; margin-top:3px;'>
            54 CAF nations · Overperformance Index · Tournament Pedigree
          </div>
        </div>
        <div style='background:rgba(255,255,255,0.12); border-radius:8px; padding:10px 18px;'>
          <div style='font-size:0.7rem; color:#93c5fd; text-transform:uppercase;
                      letter-spacing:0.08em; font-family:Georgia,serif;'>Time span</div>
          <div style='font-size:0.95rem; color:#fff; font-family:Georgia,serif; margin-top:3px;'>
            1986 – 2026 · Manually compiled
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Egypt at a glance")
    c1, c2, c3, c4, c5 = st.columns(5)
    metrics = [
        (len(clubs[clubs["Category"]=="Men"]),        "Men's clubs, 3 tiers"),
        (len(clubs[clubs["Category"]=="Women"]),      "Women's Super League"),
        (len(clubs[clubs["Category"]=="Futsal"]),     "Futsal League clubs"),
        (len(players.dropna(subset=["Name"])),        "Internationals tracked"),
        (int(ssn_df["TOTAL"].sum()) if "TOTAL" in ssn_df.columns else 536,
                                                       "Starting berths 1986–2026"),
    ]
    for col, (val, label) in zip([c1,c2,c3,c4,c5], metrics):
        with col:
            st.markdown(f"""
            <div class='metric-card'>
              <div class='metric-value'>{val}</div>
              <div class='metric-label'>{label}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### Africa at a glance")
    if caf_df is not None:
        top_over = caf_df.loc[caf_df["Overperformance"].idxmax(), "COUNTRY"]
        top_foot = caf_df.loc[caf_df["Football Index"].idxmax(), "COUNTRY"]
        top_under = caf_df.loc[caf_df["Overperformance"].idxmin(), "COUNTRY"]
        n_over = int((caf_df["Overperformance"] > 0.08).sum())
        a1,a2,a3,a4 = st.columns(4)
        for col,(val,label) in zip([a1,a2,a3,a4],[
            (54,           "CAF nations indexed"),
            (top_foot,     "Top football index"),
            (top_over,     "Biggest overperformer"),
            (top_under,    "Biggest underperformer"),
        ]):
            with col:
                st.markdown(f"""
                <div class='metric-card'>
                  <div class='metric-value' style='font-size:1.4rem;'>{val}</div>
                  <div class='metric-label'>{label}</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("### Clubs by tier")
        tier_counts = clubs.groupby(["Level","Category"]).size().reset_index(name="Count")
        tier_counts["Tier"] = tier_counts.apply(
            lambda r: f"Tier {int(r['Level'])} – {r['Category']}" if r["Category"]=="Men"
                      else r["Category"], axis=1)
        fig = px.bar(tier_counts, x="Tier", y="Count", color="Category",
                     template="plotly_white", height=280,
                     color_discrete_map={"Men":"#58a6ff","Women":"#bc8cff","Futsal":"#56d364"})
        fig.update_layout(showlegend=True, margin=dict(t=10,b=10), xaxis_title="", yaxis_title="")
        st.plotly_chart(apply_theme(fig), use_container_width=True)

    with col_b:
        st.markdown("### Top governorates by club presence")
        gov_count = clubs.groupby("Governorate").size().reset_index(name="Clubs")
        gov_count = gov_count.sort_values("Clubs", ascending=True).tail(12)
        fig2 = px.bar(gov_count, y="Governorate", x="Clubs", orientation="h",
                      template="plotly_white", height=280,
                      color="Clubs", color_continuous_scale=["#dbeafe","#1d4ed8"])
        fig2.update_layout(margin=dict(t=10,b=10), xaxis_title="", yaxis_title="",
                           coloraxis_showscale=False)
        st.plotly_chart(apply_theme(fig2), use_container_width=True)

    st.markdown("### Key insight")
    st.markdown("""
    <div style='background:#eff6ff; border:1px solid #bfdbfe; border-left:3px solid #1a56db;
                border-radius:8px; padding:16px 20px; color:#1e3a5f;'>
      Cairo/Giza dominates club presence at all three tiers, yet <b style='color:#e3b341;'>Ismailia</b>
      consistently produces starting berths disproportionate to its population and GDP —
      the clearest overperformance signal in the dataset.
      <b style='color:#56d364;'>Dakahlia</b> (28 historical starters) and
      <b style='color:#56d364;'>Gharbia</b> (23) represent untapped talent corridors.
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — PYRAMID MAP
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Pyramid Map":
    st.markdown("## The Football Pyramid")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px;'>All clubs across the top three men's tiers, Women's Super League and Futsal League.</div>", unsafe_allow_html=True)

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        cat_filter = st.multiselect("Category", ["Men","Women","Futsal"],
                                    default=["Men","Women","Futsal"])
    with col_f2:
        level_opts = sorted(clubs["Level"].dropna().unique().tolist())
        level_filter = st.multiselect("Tier (Men only)", [int(x) for x in level_opts],
                                      default=[int(x) for x in level_opts])
    with col_f3:
        gov_opts = sorted(clubs["Governorate"].dropna().unique())
        gov_filter = st.multiselect("Governorate", gov_opts, default=list(gov_opts))

    filtered = clubs[
        (clubs["Category"].isin(cat_filter)) &
        (clubs["Governorate"].isin(gov_filter)) &
        ((clubs["Category"] != "Men") | (clubs["Level"].isin(level_filter)))
    ]

    # Map: bubble chart by governorate
    gov_bubble = filtered.groupby("Governorate").agg(
        Clubs=("Name","count"),
        Categories=("Category", lambda x: ", ".join(sorted(x.unique())))
    ).reset_index()

    # Approximate lat/lon for Egyptian governorates
    GOV_COORDS = {
        "Cairo": (30.06, 31.25), "Giza": (29.99, 31.17), "Alexandria": (31.20, 29.92),
        "Port Said": (31.26, 32.28), "Suez": (29.97, 32.54), "Ismailia": (30.60, 32.27),
        "Beheira": (30.85, 30.33), "Dakahlia": (31.04, 31.38), "Damietta": (31.42, 31.82),
        "Gharbia": (30.87, 31.03), "Kafr El Sheikh": (31.11, 30.94),
        "Menoufia": (30.50, 30.99), "Qalyubia": (30.33, 31.22), "Sharkia": (30.75, 31.87),
        "Assiut": (27.18, 31.18), "Aswan": (24.09, 32.90), "Beni Suef": (29.08, 31.10),
        "Fayoum": (29.31, 30.84), "Giza ": (29.99, 31.17),
        "Luxor": (25.69, 32.64), "Minya": (28.08, 30.75), "Qena": (26.16, 32.72),
        "Sohag": (26.56, 31.70), "Matrouh": (31.35, 27.24), "New Valley": (24.55, 27.17),
        "North Sinai": (30.28, 33.62), "South Sinai": (28.50, 33.75),
        "Red Sea": (27.23, 33.83),
    }
    gov_bubble["lat"] = gov_bubble["Governorate"].map(lambda g: GOV_COORDS.get(g.strip(), (None,None))[0])
    gov_bubble["lon"] = gov_bubble["Governorate"].map(lambda g: GOV_COORDS.get(g.strip(), (None,None))[1])
    gov_bubble = gov_bubble.dropna(subset=["lat","lon"])

    fig_map = px.scatter_mapbox(
        gov_bubble, lat="lat", lon="lon", size="Clubs",
        color="Clubs", hover_name="Governorate",
        hover_data={"Categories": True, "Clubs": True, "lat": False, "lon": False},
        color_continuous_scale=["#dbeafe","#3b82f6","#1d4ed8"],
        size_max=40, zoom=4.8,
        mapbox_style="carto-positron",
        height=480,
    )
    fig_map.update_layout(paper_bgcolor="#ffffff", margin=dict(t=0,b=0,l=0,r=0),
                          coloraxis_showscale=False)
    st.plotly_chart(apply_theme(fig_map), use_container_width=True)

    st.markdown(f"**{len(filtered)} clubs** shown · Use filters above to drill down")
    st.markdown("<hr>", unsafe_allow_html=True)
    display_cols = ["Name","Level","Category","Governorate","Official District","Date Established"]
    display_cols = [c for c in display_cols if c in filtered.columns]
    csv_clubs = filtered[display_cols].sort_values(["Level","Governorate","Name"]).to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download club list (CSV)", csv_clubs,
                       file_name="egypt_clubs.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — GOVERNORATE INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Governorate Intelligence":
    st.markdown("## Governorate Intelligence")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px;'>Football representation vs. economic and demographic weight — including urban population as an economic metric.</div>", unsafe_allow_html=True)

    col_m1, col_m2, col_m3 = st.columns(3)
    ov_col = "Overperformance Index"
    if ov_col in govs.columns:
        top_over  = govs.nlargest(1, ov_col)["Governorate"].values[0]
        top_score = govs.nlargest(1,"Score")["Governorate"].values[0] if "Score" in govs.columns else "—"
        gdp_low   = govs.nsmallest(1,"GDP per capita")["Governorate"].values[0] if "GDP per capita" in govs.columns else "—"
        col_m1.metric("Top overperformer", top_over)
        col_m2.metric("Highest football score", top_score)
        col_m3.metric("Lowest GDP per capita", gdp_low)

    tab1, tab2, tab3, tab4 = st.tabs(["Overperformance ranking", "Urban vs Rural & football", "Map", "Combined Football Output"])

    with tab1:
        if ov_col in govs.columns:
            hover_g = {ov_col:":.3f","Score":True,"GDP per capita":True}
            if "Urban_Pct" in govs.columns: hover_g["Urban_Pct"] = True
            ov = govs[["Governorate",ov_col]].dropna(subset=[ov_col])
            ov = ov.sort_values(ov_col, ascending=True)
            fig_ov = px.bar(ov, y="Governorate", x=ov_col, orientation="h",
                            color=ov_col,
                            color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                            template="plotly_white", height=550)
            fig_ov.update_layout(margin=dict(t=10,b=10), xaxis_title="Overperformance Index",
                                  yaxis_title="", coloraxis_showscale=False)
            fig_ov.add_vline(x=0, line_color="#9ca3af", line_dash="dash")
            st.plotly_chart(apply_theme(fig_ov), use_container_width=True)
            lbl = "Incorporates urban population, GDP per capita and total population as economic weight." if "Urban_Pct" in govs.columns else "Positive = overperforms relative to GDP & population."
            st.caption(lbl)

    with tab2:
        if "Urban_Pct" in govs.columns and "Score" in govs.columns:
            sc_df = govs.dropna(subset=["Urban_Pct","Score","Rural_Pct"]).copy()
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                st.markdown("#### Football score vs Urban population %")
                fig_urb = px.scatter(
                    sc_df, x="Urban_Pct", y="Score",
                    text="Governorate",
                    size="Population" if "Population" in sc_df.columns else None,
                    color=ov_col if ov_col in sc_df.columns else None,
                    hover_name="Governorate",
                    hover_data={"Urban_Pct":":.1f","Rural_Pct":":.1f","Score":True},
                    color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                    size_max=45,
                    labels={"Urban_Pct":"Urban population (%)","Score":"Football representation score"},
                    template="plotly_white", height=460,
                )
                fig_urb.update_traces(textposition="top center", textfont_size=8)
                fig_urb.update_layout(margin=dict(t=10,b=10), coloraxis_showscale=False)
                st.plotly_chart(apply_theme(fig_urb), use_container_width=True)
                st.caption("Bubble size = total population. Colour = overperformance. Ismailia and Dakahlia sitting high despite low urban % are the key outliers.")
            with col_u2:
                st.markdown("#### Football score vs Rural population %")
                fig_rur = px.scatter(
                    sc_df, x="Rural_Pct", y="Score",
                    text="Governorate",
                    size="Population" if "Population" in sc_df.columns else None,
                    color=ov_col if ov_col in sc_df.columns else None,
                    hover_name="Governorate",
                    hover_data={"Rural_Pct":":.1f","Urban_Pct":":.1f","Score":True},
                    color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                    size_max=45,
                    labels={"Rural_Pct":"Rural population (%)","Score":"Football representation score"},
                    template="plotly_white", height=460,
                )
                fig_rur.update_traces(textposition="top center", textfont_size=8)
                fig_rur.update_layout(margin=dict(t=10,b=10), coloraxis_showscale=False)
                st.plotly_chart(apply_theme(fig_rur), use_container_width=True)
                st.caption("High rural % with high football score = genuine grassroots talent corridors independent of urban infrastructure.")

    with tab3:
        GOV_COORDS = {
            "Cairo": (30.06, 31.25), "Giza": (29.99, 31.17), "Alexandria": (31.20, 29.92),
            "Port Said": (31.26, 32.28), "Suez": (29.97, 32.54), "Ismailia": (30.60, 32.27),
            "Beheira": (30.85, 30.33), "Dakahlia": (31.04, 31.38), "Damietta": (31.42, 31.82),
            "Gharbia": (30.87, 31.03), "Kafr El Sheikh": (31.11, 30.94),
            "Menoufia": (30.50, 30.99), "Qalyubia": (30.33, 31.22), "Sharkia": (30.75, 31.87),
            "Assiut": (27.18, 31.18), "Aswan": (24.09, 32.90), "Beni Suef": (29.08, 31.10),
            "Fayoum": (29.31, 30.84), "Luxor": (25.69, 32.64), "Minya": (28.08, 30.75),
            "Qena": (26.16, 32.72), "Sohag": (26.56, 31.70), "Matrouh": (31.35, 27.24),
            "New Valley": (24.55, 27.17), "North Sinai": (30.28, 33.62),
            "South Sinai": (28.50, 33.75), "Red Sea": (27.23, 33.83),
        }
        map_df = govs.copy()
        map_df["lat"] = map_df["Governorate"].map(lambda g: GOV_COORDS.get(g.strip(),(None,None))[0])
        map_df["lon"] = map_df["Governorate"].map(lambda g: GOV_COORDS.get(g.strip(),(None,None))[1])
        map_df = map_df.dropna(subset=["lat","lon","Overperformance Index"])

        fig_gmap = px.scatter_mapbox(
            map_df, lat="lat", lon="lon",
            color="Overperformance Index",
            size=map_df["Population"].fillna(1000000),
            hover_name="Governorate",
            hover_data={"Overperformance Index":":.3f","Score":True,
                        "GDP per capita":True,"lat":False,"lon":False},
            color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
            size_max=40, zoom=4.8,
            mapbox_style="carto-positron", height=480,
        )
        fig_gmap.update_layout(paper_bgcolor="#ffffff", margin=dict(t=0,b=0,l=0,r=0))
        st.plotly_chart(apply_theme(fig_gmap), use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    show_cols = [c for c in ["Governorate","Population","GDP per capita","Score",
                              "Football Index","Overperformance Index",
                              "Premier League","Div2","Div 3","Womens","Futsal"]
                 if c in govs.columns]
    with tab4:
        st.markdown("#### Combined Football Output — Clubs + NT Players per Governorate")
        st.markdown("<div style='color:#6b7280;font-family:Georgia,serif;font-size:0.9rem;margin-bottom:16px;'>Combines the football infrastructure score (clubs across all tiers) with the number of international players produced by each governorate. Both metrics normalised 0–1 then averaged.</div>", unsafe_allow_html=True)

        # Build combined output
        try:
            nt_counts = players.groupby("Governorate").size().reset_index(name="NT_Players")
            nt_counts["Governorate"] = nt_counts["Governorate"].str.strip()
            nt_fix = {"Dakahleya":"Dakahlia","Dakhaleya":"Dakahlia","Gharbeya":"Gharbia",
                      "Sharqeya":"Sharkia","Sharqia":"Sharkia","Monofeya":"Menoufia",
                      "Monofiya":"Menoufia","Qalyoubeya":"Qalyubia","Ismailia ":"Ismailia",
                      "Asyut":"Assiut","Cairo ":"Cairo","Giza ":"Giza"}
            nt_counts["Governorate"] = nt_counts["Governorate"].replace(nt_fix)

            comb = govs.merge(nt_counts, on="Governorate", how="left")
            comb["NT_Players"] = comb["NT_Players"].fillna(0)

            def minmax_c(s):
                mn,mx = s.min(), s.max()
                return (s-mn)/(mx-mn) if mx!=mn else pd.Series([0.5]*len(s), index=s.index)

            if "Score" in comb.columns:
                comb["Score_norm"] = minmax_c(comb["Score"].fillna(0))
                comb["NT_norm"]    = minmax_c(comb["NT_Players"])
                comb["Combined_Output"] = ((comb["Score_norm"] + comb["NT_norm"]) / 2)
                comb = comb.sort_values("Combined_Output", ascending=False).reset_index(drop=True)
                comb["Rank"] = range(1, len(comb)+1)

                # Bar chart
                fig_comb = px.bar(
                    comb.sort_values("Combined_Output", ascending=True),
                    y="Governorate", x="Combined_Output", orientation="h",
                    color="Combined_Output",
                    color_continuous_scale=["#dbeafe","#1d4ed8"],
                    hover_data={"Score":True, "NT_Players":True, "Combined_Output":":.4f"},
                    template="plotly_white", height=600,
                    labels={"Combined_Output":"Combined Football Output (0–1)"},
                )
                fig_comb.update_layout(margin=dict(t=10,b=10), xaxis_title="Combined Football Output (0–1)",
                                       yaxis_title="", coloraxis_showscale=False)
                st.plotly_chart(apply_theme(fig_comb), use_container_width=True)

                # Side by side: club score vs NT players
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    fig_nt = px.bar(
                        comb.sort_values("NT_Players", ascending=True),
                        y="Governorate", x="NT_Players", orientation="h",
                        color="NT_Players",
                        color_continuous_scale=["#fef3c7","#d97706"],
                        template="plotly_white", height=550,
                        labels={"NT_Players":"International players produced"},
                    )
                    fig_nt.update_layout(margin=dict(t=10,b=10), xaxis_title="NT players (1986–2026)",
                                         yaxis_title="", coloraxis_showscale=False)
                    st.plotly_chart(apply_theme(fig_nt), use_container_width=True)
                    st.caption("Raw international players per governorate. Dakahlia (23) and Ismailia (15) stand out relative to their infrastructure.")

                with col_c2:
                    fig_sc2 = px.scatter(
                        comb.dropna(subset=["Score","NT_Players"]),
                        x="Score", y="NT_Players",
                        text="Governorate",
                        color="Combined_Output",
                        color_continuous_scale=["#dbeafe","#1d4ed8"],
                        size="Combined_Output",
                        size_max=30,
                        hover_name="Governorate",
                        template="plotly_white", height=550,
                        labels={"Score":"Club infrastructure score","NT_Players":"NT players produced"},
                    )
                    fig_sc2.update_traces(textposition="top center", textfont_size=8)
                    fig_sc2.update_layout(margin=dict(t=10,b=10), coloraxis_showscale=False)
                    st.plotly_chart(apply_theme(fig_sc2), use_container_width=True)
                    st.caption("Top-right = strong on both. Sharkia (top-left) produces NT players with almost no formal infrastructure. Dakahlia similar.")

                # Key insight box
                top3 = comb.nlargest(3, "NT_Players")[["Governorate","NT_Players","Score"]].reset_index(drop=True)
                hidden = comb[(comb["NT_Players"] >= 5) & (comb["Score"] <= 5)][["Governorate","NT_Players","Score"]]
                st.markdown(f"""
                <div style='background:#f0f4ff;border:1px solid #c7d7f9;border-radius:8px;
                            padding:16px 20px;margin-top:8px;font-family:Georgia,serif;color:#1e3a5f;'>
                  <div style='font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;
                              color:#6b7280;margin-bottom:10px;'>Key findings</div>
                  <div style='font-size:0.9rem;line-height:1.7;'>
                    <b>Dakahlia</b> ranks 3rd overall but has almost no club infrastructure — 
                    23 NT players from a largely rural governorate with low GDP per capita. 
                    Pure grassroots talent corridor.<br>
                    <b>Sharkia</b> produced 10 NT players with only 1 club in the formal pyramid — 
                    the starkest infrastructure gap in the dataset.<br>
                    <b>Cairo</b> leads on both metrics but its dominance is partly structural — 
                    scouts and academies concentrate there, pulling talent from other governorates.
                  </div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("""
                <div style='margin-top:20px; padding: 0 4px; font-family:Georgia,serif;
                            font-size:0.82rem; color:#9ca3af; line-height:1.8;'>
                  <i><b>Regression analysis:</b> Urban population % explains 13.1% of variance in combined 
                  football output across Egyptian governorates (R²=0.131, p=0.064), just below the 
                  conventional significance threshold. GDP per capita explains only 0.7% (R²=0.007, p=0.68) 
                  and is not a meaningful predictor. Together, both variables explain 15.5% of variance — 
                  meaning roughly 84% of what drives football talent output across governorates is captured 
                  by neither urbanisation nor wealth. Informal pathways, historical club concentration, 
                  and scouting networks are likely the dominant forces.</i>
                </div>
                """, unsafe_allow_html=True)
        except Exception as e:
            st.warning(f"Could not compute combined output: {e}")

    csv_gov = govs[show_cols].sort_values("Overperformance Index", ascending=False).to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download governorate data (CSV)", csv_gov,
                       file_name="governorate_intelligence.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — DISTRICT DENSITY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "District Density":
    st.markdown("## District Density — Greater Cairo")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px;'>Population density by district reveals where the largest untapped talent pools sit.</div>", unsafe_allow_html=True)

    if districts.empty:
        st.warning("District data could not be parsed from the file.")
    else:
        col_d1, col_d2 = st.columns([2,1])
        with col_d1:
            top_n = st.slider("Show top N districts by density", 10, len(districts), 25)

        top_dist = districts.nlargest(top_n, "Density").sort_values("Density", ascending=True)

        fig_dist = px.bar(
            top_dist, y="District", x="Density", orientation="h",
            color="Density",
            color_continuous_scale=["#dbeafe","#3b82f6","#1d4ed8"],
            hover_data={"Population":True,"Governorate":True,"Area_km2":True},
            template="plotly_white", height=max(400, top_n * 22),
        )
        fig_dist.update_layout(margin=dict(t=10,b=10),
                                xaxis_title="Density (people/km²)",
                                yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(apply_theme(fig_dist), use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown("### Density vs Population")
            fig_dp = px.scatter(
                districts.dropna(subset=["Density","Population"]),
                x="Population", y="Density",
                hover_name="District",
                color="Governorate",
                template="plotly_white", height=350,
            )
            fig_dp.update_layout(margin=dict(t=10,b=10))
            st.plotly_chart(apply_theme(fig_dp), use_container_width=True)

        with col_e2:
            st.markdown("### Clubs in dense districts")
            # Cross-ref clubs with districts
            club_dist = clubs[clubs["Official District"].str.strip() != ""].copy()
            district_clubs = club_dist.groupby("Official District").size().reset_index(name="Clubs")
            district_clubs = district_clubs.rename(columns={"Official District":"District"})
            merged = districts.merge(district_clubs, on="District", how="left").fillna({"Clubs":0})
            merged["Clubs"] = merged["Clubs"].astype(int)
            merged["Unserved"] = merged["Clubs"] == 0
            top20 = merged.nlargest(20,"Density")[["District","Population","Density","Clubs","Unserved"]]
            n_unserved = int(merged.nlargest(20,"Density")["Unserved"].sum())
            fig_unserved = px.bar(
                top20.sort_values("Density", ascending=True),
                y="District", x="Density", orientation="h",
                color="Clubs",
                color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                hover_data={"Population":True,"Clubs":True},
                template="plotly_white", height=380,
            )
            fig_unserved.update_layout(margin=dict(t=10,b=10),
                                       xaxis_title="Density (people/km²)",
                                       yaxis_title="", coloraxis_showscale=False)
            st.plotly_chart(apply_theme(fig_unserved), use_container_width=True)
            st.caption(f"⚠️ {n_unserved} of the 20 densest districts have no mapped club. Red = no club present.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — PLAYER ORIGINS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Player Origins":
    st.markdown("## International Player Origins (1986–2026)")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px;'>Where Egypt's international players come from — and when they were born.</div>", unsafe_allow_html=True)

    p = players.dropna(subset=["Governorate"])
    p = p[~p["Governorate"].isin(["", "nan", "-"])]

    tab_p1, tab_p2, tab_p3 = st.tabs(["By governorate", "Birth month (RAE)", "Player list"])

    with tab_p1:
        gov_players = p.groupby("Governorate").size().reset_index(name="Players")
        gov_players = gov_players.sort_values("Players", ascending=True)

        fig_gp = px.bar(
            gov_players, y="Governorate", x="Players", orientation="h",
            color="Players",
            color_continuous_scale=["#dbeafe","#1d4ed8"],
            template="plotly_white", height=500,
        )
        fig_gp.update_layout(margin=dict(t=10,b=10), xaxis_title="Number of players",
                              yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(apply_theme(fig_gp), use_container_width=True)

        # Merge with pop for per-capita
        if "Population" in govs.columns:
            per_cap = gov_players.merge(
                govs[["Governorate","Population","GDP per capita"]], on="Governorate", how="left")
            per_cap = per_cap.dropna(subset=["Population"])
            per_cap["Players per million"] = (per_cap["Players"] / per_cap["Population"] * 1e6).round(1)
            per_cap = per_cap.sort_values("Players per million", ascending=True)
            st.markdown("#### Players per million population")
            fig_pc = px.bar(per_cap, y="Governorate", x="Players per million",
                            orientation="h", color="Players per million",
                            color_continuous_scale=["#fef9c3","#ca8a04"],
                            template="plotly_white", height=450)
            fig_pc.update_layout(margin=dict(t=10,b=10), xaxis_title="Players per million",
                                  yaxis_title="", coloraxis_showscale=False)
            st.plotly_chart(apply_theme(fig_pc), use_container_width=True)
            st.caption("Controls for population size. Ismailia and Dakahlia stand out strongly.")

    with tab_p2:
        month_order = ["January","February","March","April","May","June",
                       "July","August","September","October","November","December"]
        birth_df = p.dropna(subset=["Birth_Month"])
        birth_counts = birth_df["Birth_Month"].value_counts().reindex(month_order, fill_value=0).reset_index()
        birth_counts.columns = ["Month","Count"]
        birth_counts["Quarter"] = birth_counts["Month"].apply(
            lambda m: "Q1 (Jan–Mar)" if m in month_order[:3]
                      else "Q2 (Apr–Jun)" if m in month_order[3:6]
                      else "Q3 (Jul–Sep)" if m in month_order[6:9]
                      else "Q4 (Oct–Dec)")

        fig_bm = px.bar(
            birth_counts, x="Month", y="Count",
            color="Quarter",
            color_discrete_map={
                "Q1 (Jan–Mar)":"#238636","Q2 (Apr–Jun)":"#1f6feb",
                "Q3 (Jul–Sep)":"#9e6a03","Q4 (Oct–Dec)":"#da3633"},
            template="plotly_white", height=380,
        )
        fig_bm.update_layout(margin=dict(t=10,b=10), xaxis_title="", yaxis_title="Players")
        st.plotly_chart(apply_theme(fig_bm), use_container_width=True)

        q_counts = birth_counts.groupby("Quarter")["Count"].sum().reset_index()
        total = q_counts["Count"].sum()
        if total > 0:
            q_counts["Share"] = (q_counts["Count"] / total * 100).round(1)
            c1, c2, c3, c4 = st.columns(4)
            for col, (_, row) in zip([c1,c2,c3,c4], q_counts.iterrows()):
                col.metric(row["Quarter"], f'{int(row["Count"])} players', f'{row["Share"]}%')
        st.caption("**Relative Age Effect (RAE):** If Q1 and Q2 are overrepresented, players born early in the year have a systematic selection advantage — a known phenomenon in youth football worldwide.")

    with tab_p3:
        show_p = [c for c in ["Name","Governorate","District","Birth_Year","Birth_Month"] if c in p.columns]
        csv_players = p[show_p].to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download full player list (CSV)", csv_players,
                           file_name="egypt_international_players.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 6 — STARTER PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Starter Pipeline":
    st.markdown("## National Team Starter Pipeline (1986–2026)")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px;'>Which governorates supplied starting berths to Egypt's national team across 40 years of international football.</div>", unsafe_allow_html=True)

    if ssn_df.empty:
        st.warning("Season-by-season data could not be loaded.")
    else:
        # Available governorate columns
        avail_govs = [c for c in gov_cols if c in ssn_df.columns and ssn_df[c].sum() > 0]

        col_s1, col_s2 = st.columns([2,1])
        with col_s1:
            selected_govs = st.multiselect(
                "Compare governorates", avail_govs,
                default=avail_govs[:6] if len(avail_govs) >= 6 else avail_govs)
        with col_s2:
            chart_type = st.radio("Chart type", ["Line","Area","Bar"], horizontal=True)

        plot_df = ssn_df[["Season"] + [g for g in selected_govs if g in ssn_df.columns]].copy()
        plot_df = plot_df.melt(id_vars="Season", var_name="Governorate", value_name="Starting berths")
        plot_df["Starting berths"] = pd.to_numeric(plot_df["Starting berths"], errors="coerce").fillna(0)

        if chart_type == "Line":
            fig_pipe = px.line(plot_df, x="Season", y="Starting berths",
                               color="Governorate", markers=True,
                               template="plotly_white", height=420)
        elif chart_type == "Area":
            fig_pipe = px.area(plot_df, x="Season", y="Starting berths",
                               color="Governorate",
                               template="plotly_white", height=420)
        else:
            fig_pipe = px.bar(plot_df, x="Season", y="Starting berths",
                              color="Governorate", barmode="stack",
                              template="plotly_white", height=420)

        fig_pipe.update_layout(margin=dict(t=10,b=10), xaxis_title="",
                                xaxis_tickangle=-45, yaxis_title="Starting berths")
        st.plotly_chart(apply_theme(fig_pipe), use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

        # Totals summary
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("### All-time totals by governorate")
            if avail_govs:
                totals = ssn_df[avail_govs].sum().sort_values(ascending=True).reset_index()
                totals.columns = ["Governorate","Total starting berths"]
                fig_tot = px.bar(totals, y="Governorate", x="Total starting berths",
                                  orientation="h", color="Total starting berths",
                                  color_continuous_scale=["#dbeafe","#1d4ed8"],
                                  template="plotly_white", height=400)
                fig_tot.update_layout(margin=dict(t=10,b=10), coloraxis_showscale=False)
                st.plotly_chart(apply_theme(fig_tot), use_container_width=True)

        with col_t2:
            st.markdown("### Cairo/Giza dominance over time")
            if "Cairo/Giza" in ssn_df.columns and "TOTAL" in ssn_df.columns:
                dom_df = ssn_df[["Season","Cairo/Giza","TOTAL"]].copy()
                dom_df["Cairo/Giza"] = pd.to_numeric(dom_df["Cairo/Giza"], errors="coerce")
                dom_df["TOTAL"] = pd.to_numeric(dom_df["TOTAL"], errors="coerce")
                dom_df = dom_df.dropna()
                dom_df["Share %"] = (dom_df["Cairo/Giza"] / dom_df["TOTAL"] * 100).round(1)
                fig_dom = px.line(dom_df, x="Season", y="Share %",
                                   template="plotly_white", height=400)
                fig_dom.update_traces(line_color="#2563eb")
                fig_dom.add_hline(y=50, line_dash="dash", line_color="#9ca3af",
                                   annotation_text="50%")
                fig_dom.update_layout(margin=dict(t=10,b=10), xaxis_tickangle=-45,
                                       yaxis_title="% of starting berths from Cairo/Giza")
                st.plotly_chart(apply_theme(fig_dom), use_container_width=True)
                st.caption("Cairo/Giza's share has grown from ~33% in 1986 to 55–60% by the 2020s.")

        st.markdown("<hr>", unsafe_allow_html=True)
        csv_ssn = ssn_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download season-by-season data (CSV)", csv_ssn,
                           file_name="season_by_season.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# CAF OVERPERFORMANCE INDEX
# ══════════════════════════════════════════════════════════════════════════════
elif page == "CAF Overperformance Index":
    st.markdown("## CAF Overperformance Index")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px; font-family:Georgia,serif;'>Football output minus economic potential across all 54 CAF nations. Positive = punching above weight.</div>", unsafe_allow_html=True)

    if caf_df is None:
        st.error("caf_index_v3.xlsx not found. Place it in the same folder as this script.")
        st.stop()

    MICRO_STATES_HEADER = {"Seychelles","Mauritius","Cape Verde","Comoros","Djibouti",
                           "Sao Tome","Eswatini","Equatorial Guinea"}
    caf_filtered = caf_df[~caf_df["COUNTRY"].isin(MICRO_STATES_HEADER)]
    top_over  = caf_filtered.loc[caf_filtered["Overperformance"].idxmax(), "COUNTRY"]
    top_foot  = caf_df.loc[caf_df["Football Index"].idxmax(), "COUNTRY"]
    top_under = caf_filtered.loc[caf_filtered["Overperformance"].idxmin(), "COUNTRY"]
    n_over    = int((caf_filtered["Overperformance"] > 0.08).sum())

    # Most underachieving zone (filtered)
    if "Zone" in caf_df.columns:
        zone_ov = caf_filtered.groupby("Zone")["Overperformance"].mean()
        worst_zone = zone_ov.idxmin()
        worst_zone_val = zone_ov.min()
    else:
        worst_zone = "—"
        worst_zone_val = 0

    c1,c2,c3,c4,c5 = st.columns(5)
    for col,(val,label) in zip([c1,c2,c3,c4,c5],[
        (top_foot,   "Top football index"),
        (top_over,   "Biggest overperformer"),
        (top_under,  "Biggest underperformer"),
        (n_over,     "Nations punching above weight"),
        (f"{worst_zone} ({worst_zone_val:.3f})", "Most underachieving zone"),
    ]):
        with col:
            st.markdown(f"""<div class='metric-card'>
              <div class='metric-value' style='font-size:1.1rem;white-space:pre-line;'>{val}</div>
              <div class='metric-label'>{label}</div>
            </div>""", unsafe_allow_html=True)

    # Countries below 500k population - micro-states excluded from findings
    MICRO_STATES = {"Seychelles","Mauritius","Cape Verde","Comoros","Djibouti",
                    "Sao Tome","Eswatini","Equatorial Guinea"}

    st.markdown("<hr>", unsafe_allow_html=True)
    tab1, tab2, tab3 = st.tabs(["Overperformance ranking", "Football vs Economy", "By zone"])

    with tab1:
        df_ov = caf_df.sort_values("Overperformance", ascending=True).copy()

        fig_ov = px.bar(df_ov, y="COUNTRY", x="Overperformance", orientation="h",
                        color="Overperformance",
                        color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                        hover_data={"Football Index":":.3f","Economic Index":":.3f","Classification":True},
                        template="plotly_white", height=950)
        fig_ov.update_layout(margin=dict(t=10,b=10),
                             xaxis_title="Overperformance (Football Index - Economic Index)",
                             yaxis_title="", coloraxis_showscale=False)
        fig_ov.add_vline(x=0, line_dash="dash", line_color="#9ca3af")
        st.plotly_chart(apply_theme(fig_ov), use_container_width=True)
        st.caption("All 54 CAF nations shown. Green = overperforms. Red = underperforms.")

        # Findings box filtered to 500k+
        df_filtered = caf_df[~caf_df["COUNTRY"].isin(MICRO_STATES)].copy()
        top_over_f  = df_filtered.loc[df_filtered["Overperformance"].idxmax(), "COUNTRY"]
        top_under_f = df_filtered.loc[df_filtered["Overperformance"].idxmin(), "COUNTRY"]
        top_over_v  = df_filtered["Overperformance"].max()
        top_under_v = df_filtered["Overperformance"].min()
        n_over_f    = int((df_filtered["Overperformance"] > 0.08).sum())

        st.markdown(f"""
        <div style='background:#f0f4ff;border:1px solid #c7d7f9;border-radius:8px;
                    padding:16px 20px;margin-top:12px;font-family:Georgia,serif;color:#1e3a5f;'>
          <div style='font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;
                      color:#6b7280;margin-bottom:10px;'>
            Key findings — nations with population 500,000+ ({len(df_filtered)} of 54)
          </div>
          <div style='display:flex;gap:32px;flex-wrap:wrap;'>
            <div>
              <div style='font-size:1.2rem;font-weight:700;color:#16a34a;'>{top_over_f}</div>
              <div style='font-size:0.82rem;color:#6b7280;'>Biggest overperformer
                <span style='color:#16a34a;font-weight:600;'>&nbsp;+{top_over_v:.3f}</span>
              </div>
            </div>
            <div>
              <div style='font-size:1.2rem;font-weight:700;color:#dc2626;'>{top_under_f}</div>
              <div style='font-size:0.82rem;color:#6b7280;'>Biggest underperformer
                <span style='color:#dc2626;font-weight:600;'>&nbsp;{top_under_v:.3f}</span>
              </div>
            </div>
            <div>
              <div style='font-size:1.2rem;font-weight:700;color:#1d4ed8;'>{n_over_f}</div>
              <div style='font-size:0.82rem;color:#6b7280;'>Nations punching above weight</div>
            </div>
          </div>
          <div style='font-size:0.75rem;color:#9ca3af;margin-top:10px;'>
            Micro-states excluded from findings: {", ".join(sorted(MICRO_STATES))}
          </div>
        </div>
        """, unsafe_allow_html=True)

    with tab2:
        hover_cols = {"Overperformance":":.3f","Classification":True}
        if "Zone" in caf_df.columns: hover_cols["Zone"] = True
        fig_sc = px.scatter(caf_df.dropna(subset=["Economic Index","Football Index"]),
                            x="Economic Index", y="Football Index",
                            text="COUNTRY", color="Overperformance",
                            color_continuous_scale=["#ef4444","#d1d5db","#16a34a"],
                            size_max=20,
                            hover_data=hover_cols,
                            template="plotly_white", height=620)
        fig_sc.update_traces(textposition="top center", textfont_size=9,
                             marker=dict(size=10))
        fig_sc.add_shape(type="line", x0=0, y0=0, x1=1, y1=1,
                         line=dict(dash="dash", color="#9ca3af"))
        fig_sc.update_layout(margin=dict(t=10,b=10),
                             xaxis_title="Economic Index", yaxis_title="Football Index")
        st.plotly_chart(apply_theme(fig_sc), use_container_width=True)
        st.caption("Above the diagonal = overperforms. Below = underperforms.")

    with tab3:
        if "Zone" in caf_df.columns:
            zone_avg = caf_df.groupby("Zone").agg(
                Countries=("COUNTRY","count"),
                Avg_Football=("Football Index","mean"),
                Avg_Economic=("Economic Index","mean"),
                Avg_Overperformance=("Overperformance","mean"),
            ).round(3).reset_index().sort_values("Avg_Overperformance", ascending=False)
            fig_zone = px.bar(zone_avg, x="Zone", y=["Avg_Football","Avg_Economic"],
                              barmode="group",
                              color_discrete_map={"Avg_Football":"#1d4ed8","Avg_Economic":"#93c5fd"},
                              template="plotly_white", height=380,
                              labels={"value":"Average Index","variable":"Metric"})
            fig_zone.update_layout(margin=dict(t=10,b=10), xaxis_title="", yaxis_title="")
            st.plotly_chart(apply_theme(fig_zone), use_container_width=True)

            col_z1, col_z2 = st.columns([1,1])
            with col_z1:
                st.markdown("**Zone averages**")
                st.dataframe(zone_avg.reset_index(drop=True), use_container_width=True)
            with col_z2:
                st.markdown("""
                <div style='background:#f0f4ff;border:1px solid #c7d7f9;border-radius:8px;
                            padding:16px 18px;font-family:Georgia,serif;font-size:0.88rem;color:#1e3a5f;'>
                <b>Football Index (8 metrics, 12.5% each)</b><br>
                FIFA Men ranking · FIFA Women ranking · Association points<br>
                Infrastructure score · Youth tournaments<br>
                Player export volume · Player export quality · Tournament pedigree
                <br><br>
                <b>Economic Index (4 metrics, 25% each)</b><br>
                Population · GDP · GDP per capita · HDI
                <br><br>
                <b>Overperformance</b> = Football − Economic<br>
                All scores min-max normalised 0–1.
                </div>
                """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    csv_caf = caf_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CAF index data (CSV)", csv_caf,
                       file_name="caf_overperformance_index.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════════════════════════
# CAF TOURNAMENT PEDIGREE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "CAF Tournament Pedigree":
    st.markdown("## CAF Tournament Pedigree")
    st.markdown("<div style='color:#6b7280; margin-bottom:20px; font-family:Georgia,serif;'>AFCON performance 2017–2025 and World Cup qualification 2018 & 2022 across all 54 CAF nations.</div>", unsafe_allow_html=True)

    if ped_df is None:
        st.warning("Tournament pedigree data not available.")
        st.stop()

    # Pedigree bar chart
    ped_plot = ped_df[["Country","Pedigree Score","AFCON Score","WC Score"]].copy()
    for c in ["Pedigree Score","AFCON Score","WC Score"]:
        ped_plot[c] = pd.to_numeric(ped_plot[c], errors="coerce")
    ped_plot = ped_plot.dropna(subset=["Pedigree Score"])
    ped_plot = ped_plot.sort_values("Pedigree Score", ascending=True)
    fig_ped = px.bar(ped_plot, y="Country", x="Pedigree Score", orientation="h",
                     color="Pedigree Score",
                     color_continuous_scale=["#dbeafe","#1d4ed8"],
                     hover_data={"AFCON Score":":.3f","WC Score":":.3f"},
                     template="plotly_white", height=950)
    fig_ped.update_layout(margin=dict(t=10,b=10), xaxis_title="Pedigree Score (0–1)",
                          yaxis_title="", coloraxis_showscale=False)
    st.plotly_chart(apply_theme(fig_ped), use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    # Tournament results grid
    result_cols = [c for c in ped_df.columns if "AFCON" in str(c) or "WC" in str(c)]
    if result_cols:
        st.markdown("### Results breakdown")
        display_ped = ped_df[["Country","Pedigree Score","AFCON Score","WC Score"] + result_cols]
        display_ped = display_ped.sort_values("Pedigree Score", ascending=False).reset_index(drop=True)

        def colour_result(val):
            colors = {"Winner":"background-color:#ffd700;color:#000",
                      "Runner-up":"background-color:#c0c0c0;color:#000",
                      "3rd":"background-color:#cd7f32;color:#fff",
                      "4th":"background-color:#fef3c7;color:#000",
                      "QF":"background-color:#dbeafe;color:#000",
                      "R16+":"background-color:#dbeafe;color:#000",
                      "Group":"background-color:#f1f5f9;color:#000",
                      "DNQ":"background-color:#fee2e2;color:#000"}
            return colors.get(str(val).strip(), "")

        # Deduplicate columns before styling
        display_ped = display_ped.loc[:,~display_ped.columns.duplicated()]
        display_ped = display_ped.reset_index(drop=True)
        result_cols = [c for c in result_cols if c in display_ped.columns]
        st.dataframe(
            display_ped.style.map(colour_result, subset=result_cols),
            use_container_width=True, height=500
        )
    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("AFCON 2025: Morocco officially declared champion after CAF stripped Senegal (March 2026). Scoring: Winner=1.0 · Runner-up=0.75 · 3rd=0.55 · 4th=0.40 · QF=0.20 · Group=0.05 · DNQ=0. WC weighted 60%, AFCON 40%.")
