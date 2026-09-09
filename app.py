# ============================================================
# Tender & Project Controls Platform
# גרסה מלאה ומתוקנת:
# 1. בדיקת כדאיות מכרזים בתוך האתר
# 2. בקרת פרויקט וניהול לו״ז מקובץ Excel
# ============================================================

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from io import BytesIO
from collections import deque
import itertools
import html

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# הגדרות עמוד ועיצוב
# ============================================================

st.set_page_config(
    page_title="Tender & Project Controls Platform",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# הגנת סיסמה — האתר מכיל נתונים עסקיים אמיתיים (מחירי הצעות, שולי רווח)
# ולכן חייב להיות נעול מאחורי סיסמה ולא נגיש לכל מי שיש לו את הקישור.
# הסיסמה עצמה לא נמצאת בקוד (שנמצא בריפו ציבורי!) אלא ב-Streamlit Secrets:
# Settings -> Secrets, ולהוסיף שורה: APP_PASSWORD = "הסיסמה שתבחר"
# ============================================================

def check_password():
    def password_entered():
        correct_password = None
        try:
            correct_password = st.secrets.get("APP_PASSWORD")
        except Exception:
            correct_password = None

        if correct_password and st.session_state.get("password_input_field") == correct_password:
            st.session_state["password_correct"] = True
            st.session_state.pop("password_input_field", None)
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    try:
        configured_password = st.secrets.get("APP_PASSWORD")
    except Exception:
        configured_password = None

    st.markdown(
        """
        <div dir="rtl" style="max-width:420px;margin:80px auto;text-align:center;
        font-family:Arial, sans-serif;">
            <div style="font-size:40px;">🔒</div>
            <h2>גישה מוגבלת</h2>
            <p style="color:#666;">האתר מכיל נתוני מכרזים ופרויקטים עסקיים. יש להזין סיסמה כדי להמשיך.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if not configured_password:
        st.warning(
            "לא הוגדרה סיסמה למערכת עדיין. יש להוסיף secret בשם APP_PASSWORD "
            "בהגדרות האפליקציה ב-Streamlit Cloud (Settings → Secrets) כדי לאפשר כניסה."
        )
        return False

    _, center_col, _ = st.columns([1, 1, 1])

    with center_col:
        st.text_input(
            "סיסמה",
            type="password",
            key="password_input_field",
            on_change=password_entered
        )

        if st.session_state.get("password_correct") is False:
            st.error("סיסמה שגויה, נסה שוב.")

    return False


if not check_password():
    st.stop()


st.markdown(
    """
    <style>
    html, body {
        direction: rtl;
    }

    .block-container {
        direction: rtl;
        text-align: right;
        font-family: Arial, sans-serif;
        padding-top: 1.2rem;
        padding-bottom: 3rem;
    }

    /* רקע בהיר קבוע לתפריט הצד + צבע טקסט כהה קבוע לצידו (חובה יחד):
       בלי שורת הצבע הזו, הטקסט בתפריט הצד מקבל את צבע ברירת המחדל של
       העיצוב (בהיר, לתצוגה על רקע כהה) — ואז אצל מבקר שהדפדפן/המחשב שלו
       במצב כהה (Dark Mode) הטקסט יוצא כמעט בלתי-קריא: בהיר על רקע בהיר. */
    section[data-testid="stSidebar"] {
        direction: rtl;
        text-align: right;
        background-color: #f1f5f9;
    }

    section[data-testid="stSidebar"] * {
        color: #0f172a !important;
    }

    h1, h2, h3, h4, h5, h6, p, label {
        direction: rtl;
        text-align: right;
        font-family: Arial, sans-serif;
    }

    .hero-box {
        direction: rtl;
        text-align: right;
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 55%, #2563eb 100%);
        color: white;
        padding: 34px;
        border-radius: 22px;
        margin-bottom: 24px;
        box-shadow: 0 16px 36px rgba(15, 23, 42, 0.18);
    }

    .hero-title {
        direction: rtl;
        text-align: right;
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        direction: rtl;
        text-align: right;
        font-size: 16px;
        opacity: 0.94;
        line-height: 1.8;
    }

    .hero-badge {
        display: inline-block;
        direction: ltr;
        text-align: center;
        background-color: rgba(255,255,255,0.14);
        border: 1px solid rgba(255,255,255,0.24);
        padding: 6px 12px;
        border-radius: 999px;
        margin-left: 8px;
        margin-top: 14px;
        font-size: 13px;
        color: white;
    }

    .kpi-card {
        direction: rtl;
        text-align: right;
        background-color: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px;
        min-height: 120px;
        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.04);
    }

    .kpi-label {
        direction: rtl;
        text-align: right;
        color: #64748b;
        font-size: 14px;
        margin-bottom: 8px;
    }

    .kpi-value {
        direction: rtl;
        text-align: right;
        color: #0f172a;
        font-size: 28px;
        font-weight: 800;
    }

    .kpi-note {
        direction: rtl;
        text-align: right;
        color: #64748b;
        font-size: 12px;
        margin-top: 5px;
    }

    .status-card {
        direction: rtl;
        text-align: right;
        border-radius: 18px;
        padding: 18px;
        margin-top: 14px;
        margin-bottom: 18px;
        font-size: 16px;
        font-weight: 700;
    }

    .status-good {
        background-color: #ecfdf5;
        border: 1px solid #a7f3d0;
        color: #047857;
    }

    .status-medium {
        background-color: #fffbeb;
        border: 1px solid #fde68a;
        color: #92400e;
    }

    .status-high {
        background-color: #fff7ed;
        border: 1px solid #fed7aa;
        color: #9a3412;
    }

    .status-critical {
        background-color: #fef2f2;
        border: 1px solid #fecaca;
        color: #991b1b;
    }

    .summary-table {
        direction: rtl !important;
        text-align: right !important;
        width: 100%;
    }

    .summary-table table {
        width: 100%;
        border-collapse: collapse;
        direction: rtl !important;
        text-align: right !important;
        font-family: Arial, sans-serif;
        margin-bottom: 20px;
        background-color: white;
    }

    .summary-table th {
        background-color: #f3f4f6;
        padding: 10px;
        border: 1px solid #e5e7eb;
        text-align: right !important;
        font-weight: bold;
    }

    .summary-table td {
        padding: 10px;
        border: 1px solid #e5e7eb;
        text-align: right !important;
    }

    .recommendation-card {
        direction: rtl !important;
        text-align: right !important;
        unicode-bidi: plaintext;
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-right: 6px solid #2563eb;
        border-left: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 16px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
        min-height: 150px;
    }

    .recommendation-title {
        direction: rtl !important;
        text-align: right !important;
        unicode-bidi: plaintext;
        font-size: 17px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 8px;
    }

    .recommendation-meta {
        direction: rtl !important;
        text-align: right !important;
        unicode-bidi: plaintext;
        color: #64748b;
        font-size: 13px;
        margin-bottom: 10px;
        line-height: 1.7;
    }

    .recommendation-action {
        direction: rtl !important;
        text-align: right !important;
        unicode-bidi: plaintext;
        color: #111827;
        font-size: 14px;
        line-height: 1.8;
    }

    .risk-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-right: 4px;
    }

    .risk-high {
        background-color: #fee2e2;
        color: #b91c1c;
        border: 1px solid #fecaca;
    }

    .risk-medium {
        background-color: #fef3c7;
        color: #b45309;
        border: 1px solid #fde68a;
    }

    .risk-low {
        background-color: #dcfce7;
        color: #047857;
        border: 1px solid #bbf7d0;
    }

    .insight-box {
        direction: rtl;
        text-align: right;
        background-color: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 12px 16px;
        margin-bottom: 10px;
        line-height: 1.7;
        color: #111827;
    }

    div[data-testid="stDataFrame"] {
        direction: rtl;
    }

    .footer-note {
        direction: rtl;
        text-align: right;
        color: #64748b;
        font-size: 13px;
        line-height: 1.7;
        margin-top: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# פונקציות עזר כלליות
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def is_empty_cell(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass

    text = str(value).strip().lower()
    return text in ["", "none", "nan", "null"]


def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        if pd.isna(value):
            return default
        text = str(value).replace(",", "").replace("₪", "").strip()
        if text.lower() in ["", "none", "nan", "null"]:
            return default
        return float(text)
    except Exception:
        return default


def safe_percent(value, default=0.0):
    # כל שדות ה"%" באתר (סיכוי זכייה %, הסתברות %, אפקטיביות טיפול % וכו') מוזנים
    # תמיד כמספר גולמי בין 0 ל-100 (למשל 30 = 30%), לא כשבר עשרוני.
    # בעבר הפונקציה ניחשה לפי הגודל אם המספר כבר שבר עשרוני — וזה גרם לפרשנות
    # הפוכה בדיוק בערכים הקטנים והחשובים ביותר לניהול סיכונים (למשל "1" שאמור
    # להיות 1% היה מתפרש כ-100%). לכן תמיד מחלקים ב-100, בלי ניחוש.
    v = safe_float(value, default)
    return v / 100


def format_money(value):
    return f"₪{safe_float(value):,.0f}"


def format_days(value):
    return f"{safe_float(value):,.1f} ימים"


def format_percent(value):
    return f"{safe_float(value):.1%}"


def normalize_series(series):
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    min_val = s.min()
    max_val = s.max()

    if max_val == min_val:
        return pd.Series(np.ones(len(s)) * 0.5, index=s.index)

    return (s - min_val) / (max_val - min_val)


def update_chart_layout(fig):
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Arial", size=13),
        title_font=dict(size=20),
        margin=dict(l=30, r=30, t=70, b=40),
        hovermode="closest"
    )
    return fig


def display_summary_table(df):
    html_table = df.to_html(index=False, escape=False)
    st.markdown(
        f"""
        <div class="summary-table" dir="rtl">
            {html_table}
        </div>
        """,
        unsafe_allow_html=True
    )


def render_hero():
    st.markdown(
        """
        <div class="hero-box" dir="rtl">
            <div class="hero-title" dir="rtl">📊 פלטפורמת החלטות מכרזים ובקרת פרויקטים</div>
            <div class="hero-subtitle" dir="rtl">
                מערכת לניתוח כדאיות הגשה למכרזים, בקרת לו״ז, סיכונים, רזרבות,
                Monte Carlo, מדדי P50/P85/P90 ודשבורד ניהולי לקבלת החלטות.
            </div>
            <span class="hero-badge">Tender Go / No-Go</span>
            <span class="hero-badge">Project Controls</span>
            <span class="hero-badge">Monte Carlo</span>
            <span class="hero-badge">Risk Analytics</span>
            <span class="hero-badge">Financial Reserve</span>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_kpi_cards(kpi_data, columns=6):
    cols = st.columns(columns)
    for col, (label, value, note) in zip(cols, kpi_data):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card" dir="rtl">
                    <div class="kpi-label" dir="rtl">{html.escape(str(label))}</div>
                    <div class="kpi-value" dir="rtl">{html.escape(str(value))}</div>
                    <div class="kpi-note" dir="rtl">{html.escape(str(note))}</div>
                </div>
                """,
                unsafe_allow_html=True
            )


def render_insights(insights):
    for i, insight in enumerate(insights, start=1):
        st.markdown(
            f"""
            <div class="insight-box" dir="rtl">
                <b>{i}.</b> {html.escape(str(insight))}
            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# מודול מכרזים
# ============================================================

TENDER_BASE_COLUMNS = [
    "מזהה מכרז",
    "שם מכרז",
    "לקוח / מזמין",
    "תחום",
    "אומדן הכנסות",
    "עלות ישירה צפויה",
    "עלות עקיפה / תקורות",
    "עלות הכנת הצעה",
    "עלות אלטרנטיבית",
    "סיכוי זכייה %",
    "זמינות צוות",
    "סיכון חוזי",
    "חשיבות אסטרטגית",
    "עמידה בתנאי סף",
    "הערות"
]


def default_tender_df():
    return pd.DataFrame(
        [
            {
                "מזהה מכרז": "T1",
                "שם מכרז": "",
                "לקוח / מזמין": "",
                "תחום": "ניהול פרויקט",
                "אומדן הכנסות": 0,
                "עלות ישירה צפויה": 0,
                "עלות עקיפה / תקורות": 0,
                "עלות הכנת הצעה": 0,
                "עלות אלטרנטיבית": 0,
                "סיכוי זכייה %": 30,
                "זמינות צוות": "בינונית",
                "סיכון חוזי": "בינוני",
                "חשיבות אסטרטגית": "בינונית",
                "עמידה בתנאי סף": "כן",
                "הערות": ""
            }
        ]
    )


def default_risk_df():
    return pd.DataFrame(
        columns=[
            "מזהה מכרז",
            "תיאור סיכון",
            "קטגוריה",
            "הסתברות %",
            "השפעה כספית",
            "עלות טיפול",
            "פעיל?"
        ]
    )


def default_capacity_df():
    return pd.DataFrame(
        columns=[
            "מזהה מכרז",
            "תפקיד",
            "שעות נדרשות",
            "שעות זמינות",
            "עומס קיים בשעות"
        ]
    )


def default_cash_df():
    return pd.DataFrame(
        columns=[
            "מזהה מכרז",
            "תקופה",
            "תקבולים צפויים",
            "תשלומים צפויים"
        ]
    )


def default_contract_df():
    return pd.DataFrame(
        columns=[
            "מזהה מכרז",
            "סעיף חוזי",
            "ציון סיכון לפני טיפול 1-5",
            "אפקטיביות טיפול %",
            "משקל"
        ]
    )


def default_licensing_df():
    return pd.DataFrame(
        columns=[
            "מזהה מכרז",
            "שלב רישוי",
            "הסתברות מעבר חודשית %",
            "עלות עיכוב חודשית"
        ]
    )


def row_has_real_tender_data(row):
    tender_id = row.get("מזהה מכרז", None)
    tender_name = row.get("שם מכרז", None)
    client = row.get("לקוח / מזמין", None)

    revenue = safe_float(row.get("אומדן הכנסות", 0))
    direct_cost = safe_float(row.get("עלות ישירה צפויה", 0))
    indirect_cost = safe_float(row.get("עלות עקיפה / תקורות", 0))
    bid_cost = safe_float(row.get("עלות הכנת הצעה", 0))
    opportunity_cost = safe_float(row.get("עלות אלטרנטיבית", 0))

    has_text = not is_empty_cell(tender_id) or not is_empty_cell(tender_name) or not is_empty_cell(client)
    has_money = any(x > 0 for x in [revenue, direct_cost, indirect_cost, bid_cost, opportunity_cost])

    return has_text or has_money


def filter_valid_tender_rows(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=TENDER_BASE_COLUMNS)

    clean_df = df.copy()

    for col in TENDER_BASE_COLUMNS:
        if col not in clean_df.columns:
            clean_df[col] = ""

    clean_df = clean_df[TENDER_BASE_COLUMNS].copy()

    clean_df = clean_df[clean_df.apply(row_has_real_tender_data, axis=1)].copy()

    if clean_df.empty:
        return pd.DataFrame(columns=TENDER_BASE_COLUMNS)

    clean_df = clean_df.reset_index(drop=True)

    for idx, row in clean_df.iterrows():
        if is_empty_cell(row["מזהה מכרז"]):
            clean_df.at[idx, "מזהה מכרז"] = f"T{idx + 1}"

    clean_df["מזהה מכרז"] = clean_df["מזהה מכרז"].astype(str).str.strip()
    clean_df = clean_df[
        ~clean_df["מזהה מכרז"].str.lower().isin(["none", "nan", "null", ""])
    ].copy()

    return clean_df.reset_index(drop=True)


def filter_rows_by_tender_id(df):
    if df is None or df.empty:
        return pd.DataFrame()

    clean_df = df.copy()

    if "מזהה מכרז" not in clean_df.columns:
        return pd.DataFrame()

    clean_df = clean_df[~clean_df["מזהה מכרז"].apply(is_empty_cell)].copy()

    if clean_df.empty:
        return pd.DataFrame(columns=df.columns)

    clean_df["מזהה מכרז"] = clean_df["מזהה מכרז"].astype(str).str.strip()
    clean_df = clean_df[
        ~clean_df["מזהה מכרז"].str.lower().isin(["none", "nan", "null", ""])
    ].copy()

    return clean_df.reset_index(drop=True)


def level_to_score(value, high=100, medium=60, low=25, reverse=False):
    value = clean_text(value)

    if value in ["גבוהה", "גבוה", "High", "high"]:
        score = high
    elif value in ["בינונית", "בינוני", "Medium", "medium"]:
        score = medium
    elif value in ["נמוכה", "נמוך", "Low", "low"]:
        score = low
    else:
        score = medium

    if reverse:
        if score == high:
            return low
        if score == medium:
            return medium
        return high

    return score


def validate_tender_input(tender_df):
    errors = []
    warnings_list = []

    tender_df = filter_valid_tender_rows(tender_df)

    if tender_df.empty:
        errors.append("לא הוזן אף מכרז תקין. חובה למלא לפחות שורה אחת.")
        return errors, warnings_list

    duplicated = tender_df[tender_df["מזהה מכרז"].duplicated()]["מזהה מכרז"].tolist()
    if duplicated:
        errors.append("נמצאו מזהי מכרז כפולים: " + ", ".join(duplicated))

    for _, row in tender_df.iterrows():
        tid = clean_text(row["מזהה מכרז"])

        if is_empty_cell(row["שם מכרז"]):
            warnings_list.append(f"מכרז {tid}: חסר שם מכרז.")

        revenue = safe_float(row["אומדן הכנסות"])
        bid_cost = safe_float(row["עלות הכנת הצעה"])
        win_prob = safe_float(row["סיכוי זכייה %"])

        if revenue < 0:
            errors.append(f"מכרז {tid}: אומדן הכנסות לא יכול להיות שלילי.")

        if bid_cost < 0:
            errors.append(f"מכרז {tid}: עלות הכנת הצעה לא יכולה להיות שלילית.")

        if win_prob < 0 or win_prob > 100:
            errors.append(f"מכרז {tid}: סיכוי זכייה חייב להיות בין 0 ל־100.")

    return errors, warnings_list


def calculate_tender_base_results(tender_df):
    df = filter_valid_tender_rows(tender_df)

    if df.empty:
        return pd.DataFrame()

    numeric_cols = [
        "אומדן הכנסות",
        "עלות ישירה צפויה",
        "עלות עקיפה / תקורות",
        "עלות הכנת הצעה",
        "עלות אלטרנטיבית",
        "סיכוי זכייה %"
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["סיכוי זכייה"] = df["סיכוי זכייה %"].apply(lambda x: safe_percent(x))

    df["רווח צפוי"] = (
        df["אומדן הכנסות"]
        - df["עלות ישירה צפויה"]
        - df["עלות עקיפה / תקורות"]
    )

    df["שיעור רווח"] = np.where(
        df["אומדן הכנסות"] > 0,
        df["רווח צפוי"] / df["אומדן הכנסות"],
        0
    )

    df["ערך צפוי EV"] = (
        df["סיכוי זכייה"] * df["רווח צפוי"]
        - df["עלות הכנת הצעה"]
        - df["עלות אלטרנטיבית"]
    )

    df["ציון קיבולת"] = df["זמינות צוות"].apply(
        lambda x: level_to_score(x, high=100, medium=60, low=25)
    )

    df["ציון סיכון חוזי"] = df["סיכון חוזי"].apply(
        lambda x: level_to_score(x, high=100, medium=60, low=20, reverse=True)
    )

    df["ציון אסטרטגי"] = df["חשיבות אסטרטגית"].apply(
        lambda x: level_to_score(x, high=100, medium=60, low=20)
    )

    # ציון ה-EV מחושב כיחס EV מול אומדן ההכנסות (לא ביחס לשאר המכרזים בהרצה).
    # ערך צפוי ששווה ל-EV_REFERENCE_RATIO מההכנסה (ברירת מחדל: 15%) ומעלה מקבל
    # ציון מלא. זה חשוב כדי שהציון יהיה משמעותי גם כשבודקים מכרז אחד בלבד,
    # ולא יתנפח באופן מלאכותי רק כי מכרז אחר בהרצה גרוע יותר ממנו.
    EV_REFERENCE_RATIO = 0.15
    ev_ratio = np.where(
        df["אומדן הכנסות"] > 0,
        df["ערך צפוי EV"] / df["אומדן הכנסות"],
        np.where(df["ערך צפוי EV"] > 0, 1.0, -1.0)
    )
    df["ציון EV מנורמל"] = (np.clip(ev_ratio / EV_REFERENCE_RATIO, -1, 1) + 1) / 2
    df["ציון רווחיות"] = np.clip(df["שיעור רווח"] / 0.25, -1, 1)
    df["ציון רווחיות"] = ((df["ציון רווחיות"] + 1) / 2) * 100
    df["ציון זכייה"] = df["סיכוי זכייה"] * 100

    df["ציון מכרז משולב"] = (
        0.35 * (df["ציון EV מנורמל"] * 100)
        + 0.20 * df["ציון רווחיות"]
        + 0.20 * df["ציון זכייה"]
        + 0.10 * df["ציון קיבולת"]
        + 0.10 * df["ציון סיכון חוזי"]
        + 0.05 * df["ציון אסטרטגי"]
    )

    decisions = []
    reasons = []
    actions = []

    for _, row in df.iterrows():
        threshold = clean_text(row["עמידה בתנאי סף"])
        score = safe_float(row["ציון מכרז משולב"])
        ev = safe_float(row["ערך צפוי EV"])
        win_prob = safe_float(row["סיכוי זכייה"])
        bid_cost = safe_float(row["עלות הכנת הצעה"])
        capacity = clean_text(row["זמינות צוות"])
        contract_risk = clean_text(row["סיכון חוזי"])

        decision = "נדרש בירור נוסף"
        reason = "הציון המשולב נמצא בטווח ביניים."
        action = "לבצע בדיקת עומק לפני החלטה סופית."

        if threshold == "לא":
            decision = "לא מומלץ לגשת"
            reason = "אין עמידה בתנאי סף."
            action = "לא להתקדם אלא אם ניתן להשלים את תנאי הסף לפני מועד ההגשה."

        elif ev < -abs(bid_cost):
            decision = "לא מומלץ לגשת"
            reason = "הערך הצפוי שלילי ביחס לעלות ההצעה."
            action = "לבחון מחדש מחיר, סיכוי זכייה או סיבה אסטרטגית חריגה."

        elif win_prob < 0.15 and bid_cost > 0:
            decision = "לא מומלץ לגשת"
            reason = "סיכוי הזכייה נמוך ועלות הכנת ההצעה קיימת."
            action = "לא להשקיע בהצעה ללא יתרון תחרותי ברור."

        elif contract_risk in ["גבוהה", "גבוה"]:
            decision = "נדרש בירור נוסף"
            reason = "סיכון חוזי גבוה."
            action = "להעביר לבדיקה משפטית / הנהלה לפני החלטת Go."

        elif capacity in ["נמוכה", "נמוך"]:
            decision = "נדרש בירור נוסף"
            reason = "קיבולת צוות נמוכה."
            action = "לבדוק זמינות משאבים או תגבור לפני הגשה."

        elif score >= 70 and ev > 0:
            decision = "מומלץ לגשת"
            reason = "ציון משולב גבוה וערך צפוי חיובי."
            action = "להתקדם להכנת הצעה ולחדד תמחור וסיכונים."

        elif score < 50:
            decision = "לא מומלץ לגשת"
            reason = "ציון מכרז משולב נמוך."
            action = "לא לגשת בשלב זה, אלא אם קיימת חשיבות אסטרטגית חריגה."

        decisions.append(decision)
        reasons.append(reason)
        actions.append(action)

    df["החלטה"] = decisions
    df["סיבה מרכזית"] = reasons
    df["פעולה מומלצת"] = actions

    return df


def calculate_tender_risks(risk_df):
    df = filter_rows_by_tender_id(risk_df)

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    for col in ["הסתברות %", "השפעה כספית", "עלות טיפול"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["פעיל?"] = df["פעיל?"].astype(str).str.strip()
    df = df[df["פעיל?"] != "לא"].copy()

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df["הסתברות"] = df["הסתברות %"].apply(lambda x: safe_percent(x))
    df["עלות סיכון צפויה"] = df["הסתברות"] * df["השפעה כספית"] + df["עלות טיפול"]

    summary = (
        df.groupby("מזהה מכרז", dropna=False)["עלות סיכון צפויה"]
        .sum()
        .reset_index()
        .rename(columns={"עלות סיכון צפויה": "עלות סיכונים צפויה"})
    )

    return df, summary


def calculate_tender_capacity(capacity_df):
    df = filter_rows_by_tender_id(capacity_df)

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    for col in ["שעות נדרשות", "שעות זמינות", "עומס קיים בשעות"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["עומס כולל"] = df["שעות נדרשות"] + df["עומס קיים בשעות"]

    df["ניצולת"] = np.where(
        df["שעות זמינות"] > 0,
        df["עומס כולל"] / df["שעות זמינות"],
        0
    )

    df["סטטוס קיבולת מפורט"] = np.where(
        df["ניצולת"] > 1,
        "חריגה מקיבולת",
        np.where(df["ניצולת"] > 0.8, "עומס גבוה", "תקין")
    )

    summary = (
        df.groupby("מזהה מכרז", dropna=False)
        .agg(
            שעות_נדרשות=("שעות נדרשות", "sum"),
            שעות_זמינות=("שעות זמינות", "sum"),
            עומס_קיים=("עומס קיים בשעות", "sum"),
            ניצולת_מקסימלית=("ניצולת", "max")
        )
        .reset_index()
    )

    summary["סטטוס קיבולת מפורט"] = np.where(
        summary["ניצולת_מקסימלית"] > 1,
        "חריגה מקיבולת",
        np.where(summary["ניצולת_מקסימלית"] > 0.8, "עומס גבוה", "תקין")
    )

    return df, summary


def calculate_cash_flow(cash_df):
    df = filter_rows_by_tender_id(cash_df)

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df["תקבולים צפויים"] = pd.to_numeric(df["תקבולים צפויים"], errors="coerce").fillna(0)
    df["תשלומים צפויים"] = pd.to_numeric(df["תשלומים צפויים"], errors="coerce").fillna(0)

    df["תזרים נקי"] = df["תקבולים צפויים"] - df["תשלומים צפויים"]
    df = df.sort_values(["מזהה מכרז", "תקופה"]).copy()
    df["תזרים מצטבר"] = df.groupby("מזהה מכרז")["תזרים נקי"].cumsum()

    summary = (
        df.groupby("מזהה מכרז", dropna=False)
        .agg(
            תזרים_נקי_כולל=("תזרים נקי", "sum"),
            תזרים_מצטבר_מינימלי=("תזרים מצטבר", "min")
        )
        .reset_index()
    )

    summary["סטטוס תזרים"] = np.where(
        summary["תזרים_מצטבר_מינימלי"] < 0,
        "תזרים שלילי בתקופה מסוימת",
        "תקין"
    )

    return df, summary


def calculate_contract_risk(contract_df):
    df = filter_rows_by_tender_id(contract_df)

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df["ציון סיכון לפני טיפול 1-5"] = pd.to_numeric(
        df["ציון סיכון לפני טיפול 1-5"], errors="coerce"
    ).fillna(3)

    df["אפקטיביות טיפול %"] = pd.to_numeric(
        df["אפקטיביות טיפול %"], errors="coerce"
    ).fillna(0)

    df["משקל"] = pd.to_numeric(df["משקל"], errors="coerce").fillna(1)

    df["אפקטיביות טיפול"] = df["אפקטיביות טיפול %"].apply(lambda x: safe_percent(x))
    df["ציון לאחר טיפול"] = df["ציון סיכון לפני טיפול 1-5"] * (1 - df["אפקטיביות טיפול"])
    df["ציון משוקלל"] = df["ציון לאחר טיפול"] * df["משקל"]

    summary = (
        df.groupby("מזהה מכרז", dropna=False)
        .agg(
            סכום_ציון_משוקלל=("ציון משוקלל", "sum"),
            סכום_משקל=("משקל", "sum")
        )
        .reset_index()
    )

    summary["ציון סיכון חוזי מתקדם"] = np.where(
        summary["סכום_משקל"] > 0,
        summary["סכום_ציון_משוקלל"] / summary["סכום_משקל"],
        0
    )

    summary["סטטוס סיכון חוזי מתקדם"] = np.where(
        summary["ציון סיכון חוזי מתקדם"] >= 4,
        "גבוה",
        np.where(summary["ציון סיכון חוזי מתקדם"] >= 2.5, "בינוני", "נמוך")
    )

    return df, summary


def triangular_sample(min_val, mode_val, max_val, n):
    min_val = safe_float(min_val)
    mode_val = safe_float(mode_val)
    max_val = safe_float(max_val)

    if min_val > mode_val:
        min_val = mode_val

    if mode_val > max_val:
        max_val = mode_val

    if min_val == max_val:
        return np.full(n, min_val)

    return np.random.triangular(min_val, mode_val, max_val, size=n)


def run_tender_monte_carlo(tender_results, n_simulations=5000, uncertainty_factor=0.25, random_seed=42):
    np.random.seed(int(random_seed))

    rows = []
    simulations = []

    for _, row in tender_results.iterrows():
        tender_id = row["מזהה מכרז"]
        tender_name = row["שם מכרז"]

        profit = safe_float(row["רווח צפוי"])
        win_prob = safe_float(row["סיכוי זכייה"])
        proposal_cost = safe_float(row["עלות הכנת הצעה"])
        opportunity_cost = safe_float(row["עלות אלטרנטיבית"])
        risk_cost = safe_float(row.get("עלות סיכונים צפויה", 0))

        profit_min = profit * (1 - uncertainty_factor)
        profit_mode = profit
        profit_max = profit * (1 + uncertainty_factor)

        profit_samples = triangular_sample(profit_min, profit_mode, profit_max, n_simulations)
        win_samples = np.random.random(n_simulations) < win_prob

        value_samples = (
            win_samples * profit_samples
            - proposal_cost
            - opportunity_cost
            - risk_cost
        )

        loss_samples = -value_samples

        var_95 = np.percentile(loss_samples, 95)
        cvar_95 = loss_samples[loss_samples >= var_95].mean() if np.any(loss_samples >= var_95) else var_95

        rows.append({
            "מזהה מכרז": tender_id,
            "שם מכרז": tender_name,
            "MC ערך ממוצע": np.mean(value_samples),
            "MC P5": np.percentile(value_samples, 5),
            "MC P50": np.percentile(value_samples, 50),
            "MC P95": np.percentile(value_samples, 95),
            "MC הסתברות הפסד": np.mean(value_samples < 0),
            "VaR 95": var_95,
            "CVaR 95": cvar_95
        })

        simulations.append(
            pd.DataFrame(
                {
                    "מזהה מכרז": tender_id,
                    "שם מכרז": tender_name,
                    "ערך סימולציה": value_samples,
                    "הפסד סימולציה": loss_samples
                }
            )
        )

    summary_df = pd.DataFrame(rows)
    all_simulations = pd.concat(simulations, ignore_index=True) if simulations else pd.DataFrame()

    return summary_df, all_simulations


def calculate_sensitivity(tender_results):
    rows = []

    for _, row in tender_results.iterrows():
        tender_id = row["מזהה מכרז"]
        tender_name = row["שם מכרז"]

        base_profit = safe_float(row["רווח צפוי"])
        base_prob = safe_float(row["סיכוי זכייה"])
        proposal_cost = safe_float(row["עלות הכנת הצעה"])
        opportunity_cost = safe_float(row["עלות אלטרנטיבית"])
        risk_cost = safe_float(row.get("עלות סיכונים צפויה", 0))

        scenarios = {
            "בסיס": (base_profit, base_prob),
            "סיכוי זכייה -20%": (base_profit, max(base_prob * 0.8, 0)),
            "סיכוי זכייה +20%": (base_profit, min(base_prob * 1.2, 1)),
            "רווח -20%": (base_profit * 0.8, base_prob),
            "רווח +20%": (base_profit * 1.2, base_prob),
        }

        for scenario, (profit, prob) in scenarios.items():
            ev = prob * profit - proposal_cost - opportunity_cost - risk_cost

            rows.append({
                "מזהה מכרז": tender_id,
                "שם מכרז": tender_name,
                "תרחיש": scenario,
                "ערך צפוי EV": ev
            })

    return pd.DataFrame(rows)


def optimize_tender_portfolio(tender_results, max_proposal_budget, max_tenders):
    df = tender_results.copy()
    df = df[df["החלטה"] != "לא מומלץ לגשת"].copy()

    if df.empty:
        return pd.DataFrame(), {"message": "אין מכרזים מתאימים לאופטימיזציה."}

    if len(df) > 15:
        df = df.sort_values("ערך צפוי EV", ascending=False).head(15).copy()

    best_value = -1e18
    best_combo = []

    records = df.to_dict("records")

    for r in range(1, min(int(max_tenders), len(records)) + 1):
        for combo in itertools.combinations(records, r):
            total_cost = sum(safe_float(x["עלות הכנת הצעה"]) for x in combo)
            total_ev = sum(safe_float(x["ערך צפוי EV"]) for x in combo)

            if total_cost <= max_proposal_budget and total_ev > best_value:
                best_value = total_ev
                best_combo = combo

    if not best_combo:
        return pd.DataFrame(), {"message": "לא נמצא שילוב מכרזים שעומד במגבלות."}

    result_df = pd.DataFrame(best_combo)

    summary = {
        "מספר מכרזים נבחרים": len(result_df),
        "עלות הכנת הצעות כוללת": result_df["עלות הכנת הצעה"].sum(),
        "ערך צפוי כולל": result_df["ערך צפוי EV"].sum(),
        "רווח צפוי כולל": result_df["רווח צפוי"].sum()
    }

    return result_df, summary


def calculate_licensing_duration_risk(licensing_df, n_simulations=5000, random_seed=42):
    # הערה מתודולוגית: זה לא מודל Markov (אין כאן כמה מצבים ומטריצת מעברים
    # ביניהם) — זהו אומדן משך המתנה עד לאישור, מבוסס על הנחה שבכל חודש יש
    # הסתברות קבועה p לקבל את האישור (כמו הטלת מטבע חוזרת עד "עץ" ראשון).
    # ההתפלגות הזו נקראת התפלגות גיאומטרית, והתוחלת שלה היא 1/p — אבל במקום
    # להסתפק בממוצע נקודתי, דוגמים אותה בפועל באלפי סימולציות (כמו בשאר
    # הכלי) כדי לקבל גם טווח P50/P90, לא רק מספר בודד.
    df = filter_rows_by_tender_id(licensing_df)

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    rng = np.random.default_rng(int(random_seed))

    df["הסתברות מעבר חודשית %"] = pd.to_numeric(
        df["הסתברות מעבר חודשית %"], errors="coerce"
    ).fillna(50)

    df["עלות עיכוב חודשית"] = pd.to_numeric(
        df["עלות עיכוב חודשית"], errors="coerce"
    ).fillna(0)

    df["הסתברות מעבר"] = df["הסתברות מעבר חודשית %"].apply(lambda x: max(safe_percent(x), 0.01))
    df["משך צפוי בחודשים"] = 1 / df["הסתברות מעבר"]
    df["עלות עיכוב צפויה"] = df["משך צפוי בחודשים"] * df["עלות עיכוב חודשית"]

    summary_rows = []

    for tender_id, group in df.groupby("מזהה מכרז", dropna=False):
        total_months_samples = np.zeros(n_simulations)
        total_cost_samples = np.zeros(n_simulations)

        for _, row in group.iterrows():
            p = row["הסתברות מעבר"]
            monthly_cost = row["עלות עיכוב חודשית"]

            # np.random.geometric(p) מחזיר את מספר הניסיונות עד להצלחה הראשונה,
            # בדיוק המשמעות של "כמה חודשים עד שהשלב הזה יאושר".
            months_samples = rng.geometric(p, size=n_simulations)

            total_months_samples = total_months_samples + months_samples
            total_cost_samples = total_cost_samples + months_samples * monthly_cost

        summary_rows.append({
            "מזהה מכרז": tender_id,
            "משך_רישוי_צפוי": np.mean(total_months_samples),
            "משך רישוי P50 (חודשים)": np.percentile(total_months_samples, 50),
            "משך רישוי P90 (חודשים)": np.percentile(total_months_samples, 90),
            "עלות_עיכוב_רישוי_צפויה": np.mean(total_cost_samples),
            "עלות עיכוב רישוי P90": np.percentile(total_cost_samples, 90),
        })

    summary = pd.DataFrame(summary_rows)

    return df, summary


def build_tender_results_excel(output):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        output["tender_results"].to_excel(writer, sheet_name="תוצאות מכרזים", index=False)

        if not output.get("risk_detail", pd.DataFrame()).empty:
            output["risk_detail"].to_excel(writer, sheet_name="סיכונים", index=False)

        if not output.get("capacity_detail", pd.DataFrame()).empty:
            output["capacity_detail"].to_excel(writer, sheet_name="קיבולת", index=False)

        if not output.get("cash_detail", pd.DataFrame()).empty:
            output["cash_detail"].to_excel(writer, sheet_name="תזרים", index=False)

        if not output.get("contract_detail", pd.DataFrame()).empty:
            output["contract_detail"].to_excel(writer, sheet_name="סיכון חוזי", index=False)

        if not output.get("licensing_detail", pd.DataFrame()).empty:
            output["licensing_detail"].to_excel(writer, sheet_name="רישוי - זמני היתרים", index=False)

        if not output.get("mc_summary", pd.DataFrame()).empty:
            output["mc_summary"].to_excel(writer, sheet_name="Monte Carlo", index=False)

        if not output.get("sensitivity", pd.DataFrame()).empty:
            output["sensitivity"].to_excel(writer, sheet_name="רגישות", index=False)

        if not output.get("portfolio_selected", pd.DataFrame()).empty:
            output["portfolio_selected"].to_excel(writer, sheet_name="פורטפוליו נבחר", index=False)

        pd.DataFrame({"תובנות": output.get("insights", [])}).to_excel(writer, sheet_name="תובנות", index=False)

    buffer.seek(0)
    return buffer


def render_tender_cards(results):
    total = len(results)
    go_count = (results["החלטה"] == "מומלץ לגשת").sum()
    no_go_count = (results["החלטה"] == "לא מומלץ לגשת").sum()
    review_count = (results["החלטה"] == "נדרש בירור נוסף").sum()

    total_ev = results["ערך צפוי EV"].sum()
    total_profit = results["רווח צפוי"].sum()

    best_row = results.sort_values("ציון מכרז משולב", ascending=False).iloc[0]
    best_name = clean_text(best_row["שם מכרז"]) or clean_text(best_row["מזהה מכרז"])

    kpi_data = [
        ("מכרזים שנבדקו", f"{total}", "מספר שורות תקינות"),
        ("מומלץ לגשת", f"{go_count}", "Go"),
        ("לא מומלץ", f"{no_go_count}", "No-Go"),
        ("דורש בירור", f"{review_count}", "Conditional"),
        ("ערך צפוי כולל", format_money(total_ev), "EV כולל"),
        ("המכרז המוביל", str(best_name), f"ציון {best_row['ציון מכרז משולב']:.1f}")
    ]

    render_kpi_cards(kpi_data, columns=6)


def create_tender_charts(results, mc_summary=None, sensitivity=None):
    figs = {}

    decision_counts = results["החלטה"].value_counts().reset_index()
    decision_counts.columns = ["החלטה", "מספר מכרזים"]

    figs["התפלגות החלטות"] = update_chart_layout(
        px.bar(
            decision_counts,
            x="החלטה",
            y="מספר מכרזים",
            title="התפלגות החלטות Go / No-Go",
            text="מספר מכרזים"
        )
    )

    top = results.sort_values("ציון מכרז משולב", ascending=False).copy()

    figs["ציון מכרז משולב"] = update_chart_layout(
        px.bar(
            top,
            x="ציון מכרז משולב",
            y="שם מכרז",
            orientation="h",
            title="ציון מכרז משולב לפי מכרז",
            hover_data=["מזהה מכרז", "החלטה", "ערך צפוי EV"]
        ).update_layout(yaxis={"categoryorder": "total ascending"})
    )

    figs["ערך צפוי"] = update_chart_layout(
        px.bar(
            top,
            x="ערך צפוי EV",
            y="שם מכרז",
            orientation="h",
            title="ערך צפוי EV לפי מכרז",
            hover_data=["מזהה מכרז", "סיכוי זכייה", "רווח צפוי"]
        ).update_layout(yaxis={"categoryorder": "total ascending"})
    )

    figs["רווח צפוי"] = update_chart_layout(
        px.bar(
            top,
            x="רווח צפוי",
            y="שם מכרז",
            orientation="h",
            title="רווח צפוי לפי מכרז",
            hover_data=["מזהה מכרז", "שיעור רווח"]
        ).update_layout(yaxis={"categoryorder": "total ascending"})
    )

    figs["עלות הצעה מול EV"] = update_chart_layout(
        px.scatter(
            results,
            x="עלות הכנת הצעה",
            y="ערך צפוי EV",
            size=np.maximum(results["אומדן הכנסות"], 1),
            hover_name="שם מכרז",
            color="החלטה",
            title="עלות הכנת הצעה מול ערך צפוי"
        )
    )

    if mc_summary is not None and not mc_summary.empty:
        figs["Monte Carlo - הסתברות הפסד"] = update_chart_layout(
            px.bar(
                mc_summary.sort_values("MC הסתברות הפסד", ascending=False),
                x="MC הסתברות הפסד",
                y="שם מכרז",
                orientation="h",
                title="הסתברות הפסד לפי סימולציית Monte Carlo"
            ).update_layout(xaxis_tickformat=".0%", yaxis={"categoryorder": "total ascending"})
        )

    if sensitivity is not None and not sensitivity.empty:
        figs["ניתוח רגישות"] = update_chart_layout(
            px.line(
                sensitivity,
                x="תרחיש",
                y="ערך צפוי EV",
                color="שם מכרז",
                title="ניתוח רגישות לערך צפוי לפי מכרז",
                markers=True
            )
        )

    return figs


def tender_card_html(row):
    decision = clean_text(row["החלטה"])

    if decision == "מומלץ לגשת":
        badge_class = "risk-low"
    elif decision == "נדרש בירור נוסף":
        badge_class = "risk-medium"
    else:
        badge_class = "risk-high"

    name = clean_text(row["שם מכרז"]) or clean_text(row["מזהה מכרז"])

    return f"""
    <div class="recommendation-card" dir="rtl">
        <div class="recommendation-title" dir="rtl">
            {html.escape(name)}
        </div>
        <div class="recommendation-meta" dir="rtl">
            מזהה: {html.escape(str(row["מזהה מכרז"]))} |
            לקוח: {html.escape(str(row["לקוח / מזמין"]))} |
            החלטה:
            <span class="risk-badge {badge_class}">{html.escape(decision)}</span> |
            ציון: {safe_float(row["ציון מכרז משולב"]):.1f}
        </div>
        <div class="recommendation-action" dir="rtl">
            <b>ערך צפוי:</b> {format_money(row["ערך צפוי EV"])} |
            <b>רווח צפוי:</b> {format_money(row["רווח צפוי"])} |
            <b>סיכוי זכייה:</b> {format_percent(row["סיכוי זכייה"])}<br>
            <b>סיבה מרכזית:</b> {html.escape(str(row["סיבה מרכזית"]))}<br>
            <b>פעולה מומלצת:</b> {html.escape(str(row["פעולה מומלצת"]))}
        </div>
    </div>
    """


def render_tender_decision_cards(results, max_cards=6):
    top = results.sort_values("ציון מכרז משולב", ascending=False).head(max_cards).reset_index(drop=True)

    for i in range(0, len(top), 2):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(tender_card_html(top.iloc[i]), unsafe_allow_html=True)

        if i + 1 < len(top):
            with col2:
                st.markdown(tender_card_html(top.iloc[i + 1]), unsafe_allow_html=True)


def render_tender_module():
    st.header("בדיקת כדאיות הגשה למכרזים")
    st.write(
        "במסלול זה ממלאים מכרזים ישירות באתר, בוחרים אילו בדיקות להפעיל, "
        "והמערכת מחשבת החלטות Go / No-Go, ערך צפוי, רווחיות, סיכונים, קיבולת, תזרים ומדדי סיכון מתקדמים."
    )

    if "tender_input_df" not in st.session_state:
        st.session_state["tender_input_df"] = default_tender_df()

    if "risk_input_df" not in st.session_state:
        st.session_state["risk_input_df"] = default_risk_df()

    if "capacity_input_df" not in st.session_state:
        st.session_state["capacity_input_df"] = default_capacity_df()

    if "cash_input_df" not in st.session_state:
        st.session_state["cash_input_df"] = default_cash_df()

    if "contract_input_df" not in st.session_state:
        st.session_state["contract_input_df"] = default_contract_df()

    if "licensing_input_df" not in st.session_state:
        st.session_state["licensing_input_df"] = default_licensing_df()

    st.info(
        "חשוב: מלא את הנתונים בתוך הטבלה ואז לחץ על כפתור החישוב. "
        "הערכים נשמרים רק לאחר לחיצה על כפתור החישוב, כדי למנוע מחיקה בזמן ההקלדה."
    )

    with st.form("tender_main_form", clear_on_submit=False):
        st.subheader("שלב 1 — מילוי מכרזים")

        tender_df = st.data_editor(
            st.session_state["tender_input_df"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="tender_input_editor",
            column_config={
                "תחום": st.column_config.SelectboxColumn(
                    "תחום",
                    options=["ניהול פרויקט", "ניהול תכנון", "פיקוח", "PMO", "בינוי", "תשתיות", "ייעוץ", "אחר"]
                ),
                "זמינות צוות": st.column_config.SelectboxColumn(
                    "זמינות צוות",
                    options=["גבוהה", "בינונית", "נמוכה"]
                ),
                "סיכון חוזי": st.column_config.SelectboxColumn(
                    "סיכון חוזי",
                    options=["נמוך", "בינוני", "גבוה"]
                ),
                "חשיבות אסטרטגית": st.column_config.SelectboxColumn(
                    "חשיבות אסטרטגית",
                    options=["גבוהה", "בינונית", "נמוכה"]
                ),
                "עמידה בתנאי סף": st.column_config.SelectboxColumn(
                    "עמידה בתנאי סף",
                    options=["כן", "לא"]
                ),
                "סיכוי זכייה %": st.column_config.NumberColumn("סיכוי זכייה %", min_value=0, max_value=100, step=1),
                "אומדן הכנסות": st.column_config.NumberColumn("אומדן הכנסות", min_value=0, step=1000),
                "עלות ישירה צפויה": st.column_config.NumberColumn("עלות ישירה צפויה", min_value=0, step=1000),
                "עלות עקיפה / תקורות": st.column_config.NumberColumn("עלות עקיפה / תקורות", min_value=0, step=1000),
                "עלות הכנת הצעה": st.column_config.NumberColumn("עלות הכנת הצעה", min_value=0, step=500),
                "עלות אלטרנטיבית": st.column_config.NumberColumn("עלות אלטרנטיבית", min_value=0, step=500),
            }
        )

        st.subheader("שלב 2 — בחירת בדיקות להפעלה")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.checkbox("בדיקת Go / No-Go בסיסית", value=True, disabled=True)
            st.checkbox("בדיקת רווחיות ותמחור", value=True, disabled=True)
            st.checkbox("בדיקת ערך צפוי Expected Value", value=True, disabled=True)
            check_risks = st.checkbox("בדיקת סיכונים", value=True)

        with c2:
            check_capacity = st.checkbox("בדיקת קיבולת צוות", value=True)
            check_cashflow = st.checkbox("בדיקת תזרים", value=False)
            check_contract = st.checkbox("בדיקת סיכון חוזי מתקדם", value=True)
            check_mc = st.checkbox("סימולציית Monte Carlo למכרזים", value=True)

        with c3:
            check_var = st.checkbox("VaR / CVaR / CFaR", value=True)
            check_sensitivity = st.checkbox("ניתוח רגישות", value=True)
            check_portfolio = st.checkbox("אופטימיזציית פורטפוליו מכרזים", value=False)
            check_licensing = st.checkbox(
                "אומדן משכי רישוי והיתרים",
                value=False,
                help="אומדן זמן המתנה לאישור/היתר לפי הסתברות מעבר חודשית, כולל טווח P50/P90 מסימולציה."
            )

        st.subheader("שלב 3 — נתונים משלימים לבדיקות מתקדמות")

        with st.expander("נתוני סיכונים"):
            st.caption("כל שורה היא סיכון הקשור למכרז מסוים. אם אין סיכונים, אפשר להשאיר ריק.")
            risk_df = st.data_editor(
                st.session_state["risk_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tender_risk_editor",
                column_config={
                    "פעיל?": st.column_config.SelectboxColumn("פעיל?", options=["כן", "לא"]),
                    "הסתברות %": st.column_config.NumberColumn("הסתברות %", min_value=0, max_value=100, step=1),
                    "השפעה כספית": st.column_config.NumberColumn("השפעה כספית", min_value=0, step=1000),
                    "עלות טיפול": st.column_config.NumberColumn("עלות טיפול", min_value=0, step=500),
                }
            )

        with st.expander("נתוני קיבולת צוות"):
            st.caption("בדיקה זו בודקת האם המכרזים יוצרים עומס יתר על בעלי תפקידים.")
            capacity_df = st.data_editor(
                st.session_state["capacity_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tender_capacity_editor",
                column_config={
                    "תפקיד": st.column_config.SelectboxColumn(
                        "תפקיד",
                        options=["מנהל פרויקט", "מנהל תכנון", "מפקח", "כלכלן", "יועץ משפטי", "אחר"]
                    ),
                    "שעות נדרשות": st.column_config.NumberColumn("שעות נדרשות", min_value=0, step=1),
                    "שעות זמינות": st.column_config.NumberColumn("שעות זמינות", min_value=0, step=1),
                    "עומס קיים בשעות": st.column_config.NumberColumn("עומס קיים בשעות", min_value=0, step=1),
                }
            )

        with st.expander("נתוני תזרים"):
            st.caption("בדיקה זו מזהה תזרים שלילי בתקופות שונות.")
            cash_df = st.data_editor(
                st.session_state["cash_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tender_cash_editor",
                column_config={
                    "תקבולים צפויים": st.column_config.NumberColumn("תקבולים צפויים", min_value=0, step=1000),
                    "תשלומים צפויים": st.column_config.NumberColumn("תשלומים צפויים", min_value=0, step=1000),
                }
            )

        with st.expander("נתוני סיכון חוזי מתקדם"):
            st.caption("הציון הוא 1-5. ככל שהציון גבוה יותר, הסיכון גבוה יותר.")
            contract_df = st.data_editor(
                st.session_state["contract_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tender_contract_editor",
                column_config={
                    "ציון סיכון לפני טיפול 1-5": st.column_config.NumberColumn(
                        "ציון סיכון לפני טיפול 1-5",
                        min_value=1,
                        max_value=5,
                        step=1
                    ),
                    "אפקטיביות טיפול %": st.column_config.NumberColumn(
                        "אפקטיביות טיפול %",
                        min_value=0,
                        max_value=100,
                        step=1
                    ),
                    "משקל": st.column_config.NumberColumn("משקל", min_value=0, step=1),
                }
            )

        with st.expander("נתוני רישוי והיתרים"):
            st.caption("בדיקה זו מעריכה משך ועלות עיכוב בשלבי רישוי או אישור.")
            licensing_df = st.data_editor(
                st.session_state["licensing_input_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tender_markov_editor",
                column_config={
                    "הסתברות מעבר חודשית %": st.column_config.NumberColumn(
                        "הסתברות מעבר חודשית %",
                        min_value=1,
                        max_value=100,
                        step=1
                    ),
                    "עלות עיכוב חודשית": st.column_config.NumberColumn(
                        "עלות עיכוב חודשית",
                        min_value=0,
                        step=1000
                    ),
                }
            )

        st.subheader("שלב 4 — הגדרות מתקדמות")

        c4, c5, c6, c7 = st.columns(4)

        with c4:
            n_simulations = st.number_input(
                "מספר סימולציות Monte Carlo",
                min_value=1000,
                max_value=50000,
                value=5000,
                step=1000
            )

        with c5:
            uncertainty_factor = st.slider(
                "רמת אי־ודאות ברווח",
                min_value=0.05,
                max_value=0.75,
                value=0.25,
                step=0.05
            )

        with c6:
            random_seed = st.number_input(
                "Seed אקראיות",
                min_value=1,
                value=42,
                step=1
            )

        with c7:
            max_tenders = st.number_input(
                "מספר מכרזים מקסימלי בפורטפוליו",
                min_value=1,
                value=3,
                step=1
            )

        max_proposal_budget = st.number_input(
            "מגבלת תקציב להכנת הצעות עבור אופטימיזציית פורטפוליו",
            min_value=0,
            value=50000,
            step=5000
        )

        run_tender = st.form_submit_button(
            "חשב כדאיות מכרזים",
            use_container_width=True
        )

    if not run_tender:
        st.info("מלא את הנתונים בטבלה ולחץ על 'חשב כדאיות מכרזים'.")
        return

    st.session_state["tender_input_df"] = tender_df
    st.session_state["risk_input_df"] = risk_df
    st.session_state["capacity_input_df"] = capacity_df
    st.session_state["cash_input_df"] = cash_df
    st.session_state["contract_input_df"] = contract_df
    st.session_state["licensing_input_df"] = licensing_df

    errors, warnings_list = validate_tender_input(tender_df)

    if errors:
        st.error("נמצאו בעיות שמונעות חישוב.")
        for error in errors:
            st.write(f"- {error}")
        return

    if warnings_list:
        st.warning("נמצאו אזהרות בקלט. החישוב יבוצע, אך כדאי לבדוק:")
        for warning in warnings_list:
            st.write(f"- {warning}")

    tender_results = calculate_tender_base_results(tender_df)

    if tender_results.empty:
        st.error("לא נמצאו מכרזים תקינים לחישוב.")
        return

    risk_detail = pd.DataFrame()
    risk_summary = pd.DataFrame()
    capacity_detail = pd.DataFrame()
    capacity_summary = pd.DataFrame()
    cash_detail = pd.DataFrame()
    cash_summary = pd.DataFrame()
    contract_detail = pd.DataFrame()
    contract_summary = pd.DataFrame()
    licensing_detail = pd.DataFrame()
    licensing_summary = pd.DataFrame()
    mc_summary = pd.DataFrame()
    mc_sims = pd.DataFrame()
    sensitivity = pd.DataFrame()
    portfolio_selected = pd.DataFrame()
    portfolio_summary = {}

    if check_risks:
        risk_detail, risk_summary = calculate_tender_risks(risk_df)
        if not risk_summary.empty:
            tender_results = tender_results.merge(risk_summary, on="מזהה מכרז", how="left")
            tender_results["עלות סיכונים צפויה"] = tender_results["עלות סיכונים צפויה"].fillna(0)
            tender_results["ערך צפוי EV"] = tender_results["ערך צפוי EV"] - tender_results["עלות סיכונים צפויה"]
        else:
            tender_results["עלות סיכונים צפויה"] = 0

    if check_capacity:
        capacity_detail, capacity_summary = calculate_tender_capacity(capacity_df)
        if not capacity_summary.empty:
            tender_results = tender_results.merge(capacity_summary, on="מזהה מכרז", how="left")

    if check_cashflow:
        cash_detail, cash_summary = calculate_cash_flow(cash_df)
        if not cash_summary.empty:
            tender_results = tender_results.merge(cash_summary, on="מזהה מכרז", how="left")

    if check_contract:
        contract_detail, contract_summary = calculate_contract_risk(contract_df)
        if not contract_summary.empty:
            tender_results = tender_results.merge(contract_summary, on="מזהה מכרז", how="left")

    if check_licensing:
        licensing_detail, licensing_summary = calculate_licensing_duration_risk(
            licensing_df,
            n_simulations=int(n_simulations),
            random_seed=int(random_seed)
        )
        if not licensing_summary.empty:
            tender_results = tender_results.merge(licensing_summary, on="מזהה מכרז", how="left")
            tender_results["עלות_עיכוב_רישוי_צפויה"] = tender_results["עלות_עיכוב_רישוי_צפויה"].fillna(0)
            tender_results["ערך צפוי EV"] = tender_results["ערך צפוי EV"] - tender_results["עלות_עיכוב_רישוי_צפויה"]
        else:
            tender_results["עלות_עיכוב_רישוי_צפויה"] = 0

    if check_mc or check_var:
        mc_summary, mc_sims = run_tender_monte_carlo(
            tender_results=tender_results,
            n_simulations=int(n_simulations),
            uncertainty_factor=float(uncertainty_factor),
            random_seed=int(random_seed)
        )
        if not mc_summary.empty:
            tender_results = tender_results.merge(mc_summary, on=["מזהה מכרז", "שם מכרז"], how="left")

    if check_sensitivity:
        sensitivity = calculate_sensitivity(tender_results)

    if check_portfolio:
        portfolio_selected, portfolio_summary = optimize_tender_portfolio(
            tender_results=tender_results,
            max_proposal_budget=max_proposal_budget,
            max_tenders=int(max_tenders)
        )

    insights = [
        f"נבדקו {len(tender_results)} מכרזים תקינים.",
        f"מספר מכרזים שמומלץ לגשת אליהם: {(tender_results['החלטה'] == 'מומלץ לגשת').sum()}.",
        f"מספר מכרזים שלא מומלץ לגשת אליהם: {(tender_results['החלטה'] == 'לא מומלץ לגשת').sum()}.",
        f"מספר מכרזים שדורשים בירור נוסף: {(tender_results['החלטה'] == 'נדרש בירור נוסף').sum()}.",
        f"הערך הצפוי הכולל של המכרזים הוא {format_money(tender_results['ערך צפוי EV'].sum())}.",
        f"עלות הכנת ההצעות הכוללת היא {format_money(tender_results['עלות הכנת הצעה'].sum())}."
    ]

    best_row = tender_results.sort_values("ציון מכרז משולב", ascending=False).iloc[0]
    insights.append(
        f"המכרז בעל הציון המשולב הגבוה ביותר הוא '{clean_text(best_row['שם מכרז']) or clean_text(best_row['מזהה מכרז'])}' "
        f"עם ציון {best_row['ציון מכרז משולב']:.1f}."
    )

    output = {
        "tender_results": tender_results,
        "risk_detail": risk_detail,
        "capacity_detail": capacity_detail,
        "cash_detail": cash_detail,
        "contract_detail": contract_detail,
        "licensing_detail": licensing_detail,
        "mc_summary": mc_summary,
        "sensitivity": sensitivity,
        "portfolio_selected": portfolio_selected,
        "portfolio_summary": portfolio_summary,
        "insights": insights
    }

    st.success("חישוב המכרזים הסתיים בהצלחה.")

    render_tender_cards(tender_results)

    tabs = st.tabs([
        "סיכום מכרזים",
        "גרפים",
        "טבלת תוצאות",
        "בדיקות מתקדמות",
        "הורדת דוח"
    ])

    with tabs[0]:
        st.subheader("תובנות ניהוליות")
        render_insights(insights)

        st.subheader("כרטיסי החלטה למכרזים")
        render_tender_decision_cards(tender_results, max_cards=6)

    with tabs[1]:
        figs = create_tender_charts(tender_results, mc_summary=mc_summary, sensitivity=sensitivity)

        col1, col2 = st.columns(2)

        for i, (title, fig) in enumerate(figs.items()):
            if i % 2 == 0:
                with col1:
                    st.markdown(f"### {title}")
                    st.plotly_chart(fig, use_container_width=True)
            else:
                with col2:
                    st.markdown(f"### {title}")
                    st.plotly_chart(fig, use_container_width=True)

    with tabs[2]:
        st.subheader("טבלת תוצאות מכרזים")
        st.dataframe(tender_results, use_container_width=True, hide_index=True)

    with tabs[3]:
        if not risk_detail.empty:
            st.markdown("### סיכונים")
            st.dataframe(risk_detail, use_container_width=True, hide_index=True)

        if not capacity_detail.empty:
            st.markdown("### קיבולת צוות")
            st.dataframe(capacity_detail, use_container_width=True, hide_index=True)

        if not cash_detail.empty:
            st.markdown("### תזרים")
            st.dataframe(cash_detail, use_container_width=True, hide_index=True)

        if not contract_detail.empty:
            st.markdown("### סיכון חוזי")
            st.dataframe(contract_detail, use_container_width=True, hide_index=True)

        if not licensing_detail.empty:
            st.markdown("### רישוי והיתרים")
            st.dataframe(licensing_detail, use_container_width=True, hide_index=True)

        if not mc_summary.empty:
            st.markdown("### Monte Carlo / VaR / CVaR")
            st.dataframe(mc_summary, use_container_width=True, hide_index=True)

        if not sensitivity.empty:
            st.markdown("### ניתוח רגישות")
            st.dataframe(sensitivity, use_container_width=True, hide_index=True)

        if not portfolio_selected.empty:
            st.markdown("### פורטפוליו מכרזים מומלץ")
            st.dataframe(portfolio_selected, use_container_width=True, hide_index=True)
            st.json(portfolio_summary)

        if (
            risk_detail.empty
            and capacity_detail.empty
            and cash_detail.empty
            and contract_detail.empty
            and licensing_detail.empty
            and mc_summary.empty
            and sensitivity.empty
            and portfolio_selected.empty
        ):
            st.info("לא הוזנו נתונים מתקדמים או שלא הופעלו בדיקות מתקדמות.")

    with tabs[4]:
        excel_buffer = build_tender_results_excel(output)

        st.download_button(
            label="הורד דוח בדיקת מכרזים Excel",
            data=excel_buffer,
            file_name="תוצאות_בדיקת_מכרזים.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# ============================================================
# מודול בקרת פרויקט
# ============================================================

SHEET_TASKS = "משימות"
SHEET_LINKS = "קשרים"
SHEET_FINANCE = "פרמטרים פיננסיים"
SHEET_RISKS = "סיכונים"
SHEET_SCENARIOS = "תרחישים"

TASK_COLUMNS_HE = {
    "מזהה פעילות": "task_id",
    "שם פעילות": "task_name",
    "שלב": "stage",
    "סוג פעילות": "task_type",
    "תאריך התחלה מוקדם": "earliest_start_date",
    "משך אופטימי בימים": "optimistic_days",
    "משך סביר בימים": "most_likely_days",
    "משך פסימי בימים": "pessimistic_days",
    "עלות בסיסית": "base_cost",
    "אחראי": "owner",
    "קטגוריה": "category"
}

LINK_COLUMNS_HE = {
    "מזהה פעילות קודמת": "predecessor_id",
    "מזהה פעילות עוקבת": "successor_id",
    "סוג קשר": "relationship_type",
    "Lag בימים": "lag_days",
    "הערה": "note"
}

RISK_COLUMNS_HE = {
    "מזהה סיכון": "risk_id",
    "תיאור סיכון": "risk_description",
    "פעילות קשורה": "related_task_id",
    "הסתברות": "probability",
    "השפעה מינימלית בימים": "min_delay_days",
    "השפעה סבירה בימים": "most_likely_delay_days",
    "השפעה מקסימלית בימים": "max_delay_days",
    "השפעה כספית": "cost_impact",
    "פעיל?": "is_active"
}

FINANCE_COLUMNS_HE = {
    "פרמטר": "parameter",
    "ערך": "value",
    "יחידה": "unit",
    "הסבר": "description"
}

SCENARIO_COLUMNS_HE = {
    "שם תרחיש": "scenario_name",
    "מספר סימולציות": "n_simulations",
    "משך יעד בימים": "target_duration",
    "רמת ביטחון": "confidence_level",
    "הערה": "note"
}

REQUIRED_SHEETS = [
    SHEET_TASKS,
    SHEET_LINKS,
    SHEET_FINANCE,
    SHEET_RISKS,
    SHEET_SCENARIOS
]


def read_sheet_hebrew(file_path, sheet_name, columns_map):
    df_original = pd.read_excel(file_path, sheet_name=sheet_name)
    df_original.columns = df_original.columns.astype(str).str.strip()

    missing_columns = [col for col in columns_map.keys() if col not in df_original.columns]

    if missing_columns:
        raise ValueError(
            f"בגיליון '{sheet_name}' חסרות העמודות הבאות: "
            + ", ".join(missing_columns)
        )

    df = df_original.rename(columns=columns_map)
    return df_original, df


def load_all_inputs(file_path):
    xls = pd.ExcelFile(file_path)
    existing_sheets = xls.sheet_names

    missing_sheets = [s for s in REQUIRED_SHEETS if s not in existing_sheets]

    if missing_sheets:
        raise ValueError("חסרים בקובץ הגיליונות הבאים: " + ", ".join(missing_sheets))

    tasks_original, tasks = read_sheet_hebrew(file_path, SHEET_TASKS, TASK_COLUMNS_HE)
    links_original, links = read_sheet_hebrew(file_path, SHEET_LINKS, LINK_COLUMNS_HE)
    finance_original, finance = read_sheet_hebrew(file_path, SHEET_FINANCE, FINANCE_COLUMNS_HE)
    risks_original, risks = read_sheet_hebrew(file_path, SHEET_RISKS, RISK_COLUMNS_HE)
    scenarios_original, scenarios = read_sheet_hebrew(file_path, SHEET_SCENARIOS, SCENARIO_COLUMNS_HE)

    return {
        "tasks_original": tasks_original,
        "links_original": links_original,
        "finance_original": finance_original,
        "risks_original": risks_original,
        "scenarios_original": scenarios_original,
        "tasks": tasks,
        "links": links,
        "finance": finance,
        "risks": risks,
        "scenarios": scenarios
    }


def validate_tasks(tasks):
    df = tasks.copy()

    if df["task_id"].isna().any():
        raise ValueError("בגיליון משימות יש שורות עם מזהה פעילות ריק.")

    df["task_id"] = df["task_id"].astype(str).str.strip()
    df["task_name"] = df["task_name"].astype(str).str.strip()
    df["stage"] = df["stage"].astype(str).str.strip()
    df["task_type"] = df["task_type"].astype(str).str.strip()
    df["owner"] = df["owner"].astype(str).str.strip()
    df["category"] = df["category"].astype(str).str.strip()

    duplicated = df[df["task_id"].duplicated()]["task_id"].tolist()

    if duplicated:
        raise ValueError("נמצאו מזהי פעילות כפולים: " + ", ".join(duplicated))

    duration_cols = ["optimistic_days", "most_likely_days", "pessimistic_days"]

    for col in duration_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

        if df[col].isna().any():
            raise ValueError(f"בגיליון משימות, בעמודה {col}, יש ערכים ריקים או לא מספריים.")

        if (df[col] < 0).any():
            raise ValueError(f"בגיליון משימות, בעמודה {col}, יש ערכים שליליים.")

    invalid_duration = df[
        (df["optimistic_days"] > df["most_likely_days"]) |
        (df["most_likely_days"] > df["pessimistic_days"])
    ]

    if not invalid_duration.empty:
        ids = invalid_duration["task_id"].tolist()
        raise ValueError("נמצאו פעילויות שבהן סדר הזמנים אינו תקין: " + ", ".join(ids))

    df["base_cost"] = pd.to_numeric(df["base_cost"], errors="coerce").fillna(0)

    if (df["base_cost"] < 0).any():
        raise ValueError("בגיליון משימות נמצאו עלויות בסיסיות שליליות.")

    return df


def validate_links(links, tasks):
    df = links.copy()

    if df.empty:
        return df

    df["predecessor_id"] = df["predecessor_id"].astype(str).str.strip()
    df["successor_id"] = df["successor_id"].astype(str).str.strip()
    df["relationship_type"] = df["relationship_type"].astype(str).str.strip().str.upper()
    df["lag_days"] = pd.to_numeric(df["lag_days"], errors="coerce").fillna(0)

    allowed_relationships = {"FS", "SS", "FF", "SF"}
    invalid_rel = df[~df["relationship_type"].isin(allowed_relationships)]

    if not invalid_rel.empty:
        raise ValueError("נמצאו סוגי קשר לא תקינים. מותר רק: FS, SS, FF, SF")

    task_ids = set(tasks["task_id"])

    missing_pred = df[~df["predecessor_id"].isin(task_ids)]["predecessor_id"].tolist()
    missing_succ = df[~df["successor_id"].isin(task_ids)]["successor_id"].tolist()

    if missing_pred:
        raise ValueError("פעילויות קודמות שלא קיימות במשימות: " + ", ".join(missing_pred))

    if missing_succ:
        raise ValueError("פעילויות עוקבות שלא קיימות במשימות: " + ", ".join(missing_succ))

    self_links = df[df["predecessor_id"] == df["successor_id"]]

    if not self_links.empty:
        raise ValueError("נמצאו קשרים שבהם פעילות תלויה בעצמה.")

    return df


def validate_risks(risks, tasks):
    df = risks.copy()

    if df.empty:
        return df

    df["risk_id"] = df["risk_id"].astype(str).str.strip()
    df["related_task_id"] = df["related_task_id"].astype(str).str.strip()
    df["is_active"] = df["is_active"].astype(str).str.strip()

    task_ids = set(tasks["task_id"])
    missing_related_tasks = df[~df["related_task_id"].isin(task_ids)]["related_task_id"].tolist()

    if missing_related_tasks:
        raise ValueError("בגיליון סיכונים יש פעילויות קשורות שלא קיימות: " + ", ".join(missing_related_tasks))

    numeric_cols = [
        "probability",
        "min_delay_days",
        "most_likely_delay_days",
        "max_delay_days",
        "cost_impact"
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    invalid_prob = df[(df["probability"] < 0) | (df["probability"] > 1)]

    if not invalid_prob.empty:
        raise ValueError("בגיליון סיכונים יש הסתברויות שאינן בין 0 ל-1.")

    invalid_delay = df[
        (df["min_delay_days"] > df["most_likely_delay_days"]) |
        (df["most_likely_delay_days"] > df["max_delay_days"])
    ]

    if not invalid_delay.empty:
        raise ValueError("בגיליון סיכונים יש סדר השפעת ימים לא תקין.")

    return df


def read_financial_parameters(finance):
    df = finance.copy()

    df["parameter"] = df["parameter"].astype(str).str.strip()
    df["value"] = pd.to_numeric(df["value"], errors="coerce")

    params = dict(zip(df["parameter"], df["value"]))

    required = [
        "תקציב בסיסי",
        "ריבית שנתית",
        "עלות תקורה יומית",
        "מדד תשומות שנתי",
        "רמת ביטחון לרזרבה",
        "משך יעד בימים",
        "מספר סימולציות",
        "Seed אקראיות"
    ]

    missing = [p for p in required if p not in params]

    if missing:
        raise ValueError("חסרים פרמטרים פיננסיים: " + ", ".join(missing))

    return params


def choose_scenario(scenarios, scenario_name="בסיס"):
    df = scenarios.copy()

    df["scenario_name"] = df["scenario_name"].astype(str).str.strip()
    df["n_simulations"] = pd.to_numeric(df["n_simulations"], errors="coerce")
    df["target_duration"] = pd.to_numeric(df["target_duration"], errors="coerce")
    df["confidence_level"] = pd.to_numeric(df["confidence_level"], errors="coerce")

    if scenario_name not in df["scenario_name"].values:
        raise ValueError(f"לא נמצא תרחיש בשם {scenario_name} בגיליון תרחישים.")

    row = df[df["scenario_name"] == scenario_name].iloc[0]

    return {
        "scenario_name": row["scenario_name"],
        "n_simulations": int(row["n_simulations"]),
        "target_duration": float(row["target_duration"]),
        "confidence_level": float(row["confidence_level"])
    }


def build_graph(tasks, links):
    task_ids = tasks["task_id"].tolist()

    successors = {task_id: [] for task_id in task_ids}
    predecessors = {task_id: [] for task_id in task_ids}
    links_by_successor = {task_id: [] for task_id in task_ids}

    for _, row in links.iterrows():
        pred = row["predecessor_id"]
        succ = row["successor_id"]
        rel = row["relationship_type"]
        lag = row["lag_days"]

        successors[pred].append(succ)
        predecessors[succ].append(pred)

        links_by_successor[succ].append({
            "predecessor_id": pred,
            "successor_id": succ,
            "relationship_type": rel,
            "lag_days": lag
        })

    return successors, predecessors, links_by_successor


def topological_sort(task_ids, predecessors, successors):
    in_degree = {task_id: len(predecessors[task_id]) for task_id in task_ids}
    queue = deque([task_id for task_id in task_ids if in_degree[task_id] == 0])
    order = []

    while queue:
        current = queue.popleft()
        order.append(current)

        for succ in successors[current]:
            in_degree[succ] -= 1

            if in_degree[succ] == 0:
                queue.append(succ)

    if len(order) != len(task_ids):
        raise ValueError("נמצאה תלות מעגלית בקשרים. יש לבדוק את גיליון קשרים.")

    return order


def sample_pert(a, m, b, n, lamb=4):
    if a == m == b:
        return np.full(n, a)

    if not (a <= m <= b):
        raise ValueError("חייב להתקיים: optimistic <= most_likely <= pessimistic")

    if a == b:
        return np.full(n, a)

    alpha = 1 + lamb * ((m - a) / (b - a))
    beta = 1 + lamb * ((b - m) / (b - a))

    samples_01 = np.random.beta(alpha, beta, n)
    return a + samples_01 * (b - a)


def run_monte_carlo(tasks, links, risks, n_simulations=10000, random_seed=42):
    np.random.seed(random_seed)

    task_ids = tasks["task_id"].tolist()
    successors, predecessors, links_by_successor = build_graph(tasks, links)
    calc_order = topological_sort(task_ids, predecessors, successors)

    duration_samples = {}
    base_duration_samples = {}

    for _, row in tasks.iterrows():
        task_id = row["task_id"]

        samples = sample_pert(
            row["optimistic_days"],
            row["most_likely_days"],
            row["pessimistic_days"],
            n_simulations,
            lamb=4
        )

        base_duration_samples[task_id] = samples.copy()
        duration_samples[task_id] = samples.copy()

    risk_cost_samples = np.zeros(n_simulations)
    risk_delay_by_task = {task_id: np.zeros(n_simulations) for task_id in task_ids}

    active_risks = risks[risks["is_active"] == "כן"].copy() if not risks.empty else pd.DataFrame()

    for _, risk in active_risks.iterrows():
        related_task = risk["related_task_id"]
        p = risk["probability"]

        occurs = np.random.random(n_simulations) < p

        delay_samples = sample_pert(
            risk["min_delay_days"],
            risk["most_likely_delay_days"],
            risk["max_delay_days"],
            n_simulations,
            lamb=4
        )

        applied_delay = occurs * delay_samples
        applied_cost = occurs * risk["cost_impact"]

        duration_samples[related_task] = duration_samples[related_task] + applied_delay
        risk_delay_by_task[related_task] = risk_delay_by_task[related_task] + applied_delay
        risk_cost_samples = risk_cost_samples + applied_cost

    project_durations = np.zeros(n_simulations)

    start_results = {task_id: np.zeros(n_simulations) for task_id in task_ids}
    finish_results = {task_id: np.zeros(n_simulations) for task_id in task_ids}
    critical_count = {task_id: 0 for task_id in task_ids}

    for sim in range(n_simulations):
        start = {}
        finish = {}
        controlling_predecessor = {}

        for task_id in calc_order:
            duration = duration_samples[task_id][sim]
            constraints = []

            if len(links_by_successor[task_id]) == 0:
                constraints.append((0, None))
            else:
                for link in links_by_successor[task_id]:
                    pred = link["predecessor_id"]
                    rel = link["relationship_type"]
                    lag = link["lag_days"]

                    if rel == "FS":
                        required_start = finish[pred] + lag
                    elif rel == "SS":
                        required_start = start[pred] + lag
                    elif rel == "FF":
                        required_start = finish[pred] + lag - duration
                    elif rel == "SF":
                        required_start = start[pred] + lag - duration
                    else:
                        raise ValueError(f"סוג קשר לא מוכר: {rel}")

                    constraints.append((required_start, pred))

            selected_start, selected_pred = max(constraints, key=lambda x: x[0])
            selected_start = max(0, selected_start)

            start[task_id] = selected_start
            finish[task_id] = selected_start + duration
            controlling_predecessor[task_id] = selected_pred

            start_results[task_id][sim] = start[task_id]
            finish_results[task_id][sim] = finish[task_id]

        project_duration = max(finish.values())
        project_durations[sim] = project_duration

        terminal_task = max(finish, key=finish.get)
        current = terminal_task

        while current is not None:
            critical_count[current] += 1
            current = controlling_predecessor[current]

    criticality_index = {
        task_id: critical_count[task_id] / n_simulations
        for task_id in task_ids
    }

    return {
        "project_durations": project_durations,
        "base_duration_samples": base_duration_samples,
        "duration_samples": duration_samples,
        "risk_delay_by_task": risk_delay_by_task,
        "risk_cost_samples": risk_cost_samples,
        "start_results": start_results,
        "finish_results": finish_results,
        "criticality_index": criticality_index,
        "calculation_order": calc_order,
        "successors": successors,
        "predecessors": predecessors,
        "links_by_successor": links_by_successor
    }


def calculate_schedule_summary(results, target_duration):
    durations = results["project_durations"]

    return {
        "משך ממוצע": np.mean(durations),
        "סטיית תקן": np.std(durations),
        "P50": np.percentile(durations, 50),
        "P85": np.percentile(durations, 85),
        "P90": np.percentile(durations, 90),
        "משך מינימלי": np.min(durations),
        "משך מקסימלי": np.max(durations),
        "משך יעד": target_duration,
        "הסתברות איחור": np.mean(durations > target_duration)
    }


def calculate_financial_results(results, finance_params, target_duration, confidence_level):
    durations = results["project_durations"]
    risk_costs = results["risk_cost_samples"]

    base_budget = finance_params["תקציב בסיסי"]
    annual_interest = finance_params["ריבית שנתית"]
    overhead_daily = finance_params["עלות תקורה יומית"]
    annual_index = finance_params["מדד תשומות שנתי"]

    delay_days = np.maximum(durations - target_duration, 0)

    financing_cost_total = base_budget * annual_interest * durations / 365
    financing_cost_target = base_budget * annual_interest * target_duration / 365
    financing_cost_over_target = np.maximum(financing_cost_total - financing_cost_target, 0)

    overhead_delay_cost = overhead_daily * delay_days

    indexation_cost_total = base_budget * annual_index * durations / 365
    indexation_cost_target = base_budget * annual_index * target_duration / 365
    indexation_cost_over_target = np.maximum(indexation_cost_total - indexation_cost_target, 0)

    total_delay_cost = (
        financing_cost_over_target
        + overhead_delay_cost
        + indexation_cost_over_target
        + risk_costs
    )

    confidence_percentile = confidence_level * 100
    reserve_cost = np.percentile(total_delay_cost, confidence_percentile)

    reserve_days = max(
        np.percentile(durations, confidence_percentile) - target_duration,
        0
    )

    financial_summary = {
        "תקציב בסיסי": base_budget,
        "ריבית שנתית": annual_interest,
        "עלות תקורה יומית": overhead_daily,
        "מדד תשומות שנתי": annual_index,
        "רמת ביטחון לרזרבה": confidence_level,
        "עלות עיכוב ממוצעת": np.mean(total_delay_cost),
        "עלות עיכוב P50": np.percentile(total_delay_cost, 50),
        "עלות עיכוב P85": np.percentile(total_delay_cost, 85),
        "עלות עיכוב P90": np.percentile(total_delay_cost, 90),
        "רזרבה מומלצת לפי רמת הביטחון": reserve_cost,
        "רזרבת ימים לפי רמת הביטחון": reserve_days
    }

    financial_arrays = {
        "delay_days": delay_days,
        "financing_cost_over_target": financing_cost_over_target,
        "overhead_delay_cost": overhead_delay_cost,
        "indexation_cost_over_target": indexation_cost_over_target,
        "risk_costs": risk_costs,
        "total_delay_cost": total_delay_cost
    }

    return financial_summary, financial_arrays


def build_task_risk_table(tasks, results):
    rows = []

    for _, row in tasks.iterrows():
        task_id = row["task_id"]
        samples = results["duration_samples"][task_id]
        risk_delay = results["risk_delay_by_task"][task_id]

        p50 = np.percentile(samples, 50)
        p85 = np.percentile(samples, 85)
        p90 = np.percentile(samples, 90)

        rows.append({
            "מזהה פעילות": task_id,
            "שם פעילות": row["task_name"],
            "שלב": row["stage"],
            "סוג פעילות": row["task_type"],
            "אחראי": row["owner"],
            "קטגוריה": row["category"],
            "משך אופטימי": row["optimistic_days"],
            "משך סביר": row["most_likely_days"],
            "משך פסימי": row["pessimistic_days"],
            "משך ממוצע לאחר סיכונים": np.mean(samples),
            "תוספת ממוצעת מסיכונים": np.mean(risk_delay),
            "סטיית תקן": np.std(samples),
            "P50 משך פעילות": p50,
            "P85 משך פעילות": p85,
            "P90 משך פעילות": p90,
            "פער P90 מול משך סביר": p90 - row["most_likely_days"],
            "מדד קריטיות": results["criticality_index"][task_id],
            "עלות בסיסית": row["base_cost"]
        })

    table = pd.DataFrame(rows)

    table["נרמול קריטיות"] = normalize_series(table["מדד קריטיות"])
    table["נרמול פער P90"] = normalize_series(table["פער P90 מול משך סביר"])
    table["נרמול סיכונים"] = normalize_series(table["תוספת ממוצעת מסיכונים"])
    table["נרמול סטיית תקן"] = normalize_series(table["סטיית תקן"])

    table["ציון סיכון משולב"] = (
        0.40 * table["נרמול קריטיות"]
        + 0.30 * table["נרמול פער P90"]
        + 0.20 * table["נרמול סיכונים"]
        + 0.10 * table["נרמול סטיית תקן"]
    )

    table["רמת סיכון"] = pd.cut(
        table["ציון סיכון משולב"],
        bins=[-0.01, 0.33, 0.66, 1.01],
        labels=["נמוכה", "בינונית", "גבוהה"]
    )

    table = table.sort_values(
        by=["ציון סיכון משולב", "מדד קריטיות", "פער P90 מול משך סביר"],
        ascending=False
    )

    return table


def identify_main_risk_reason(row):
    components = {
        "קריטיות גבוהה בנתיב הפרויקט": row["נרמול קריטיות"],
        "פער גדול בין P90 לבין משך סביר": row["נרמול פער P90"],
        "השפעה גבוהה של סיכונים חיצוניים": row["נרמול סיכונים"],
        "תנודתיות גבוהה במשך הפעילות": row["נרמול סטיית תקן"]
    }

    return max(components, key=components.get)


def recommend_action(row):
    task_name = str(row["שם פעילות"])
    stage = str(row["שלב"])
    risk_level = str(row["רמת סיכון"])
    main_reason = identify_main_risk_reason(row)

    if "רכש" in task_name or "ציוד" in task_name or "חומרים" in task_name:
        return "להקדים רכש, לבדוק זמינות ספקים, להגדיר ספק חלופי ולנעול מועדי אספקה חוזיים."

    if "רשות" in task_name or "היתר" in task_name or "אישור" in task_name:
        return "לקיים פגישת תיאום מוקדמת עם הרשות, להכין רשימת דרישות מלאה ולעקוב שבועית אחר סטטוס האישור."

    if "תכנון" in stage or "יועצים" in task_name or "תיאום" in task_name:
        return "לקבוע ישיבת תיאום יועצים קבועה, להגדיר תוצרים לפי תאריך ולבצע בקרת חסמים שבועית."

    if "מערכות" in task_name or "חשמל" in task_name or "אינסטלציה" in task_name:
        return "לבצע תיאום מערכות מוקדם, לבדוק התנגשויות, לאשר Shop Drawings מוקדם ולבצע בקרת ממשקים."

    if "גמר" in task_name or "גמרים" in task_name:
        return "לבנות תוכנית האצה לגמרים, לחלק לאזורי עבודה, ולבצע בקרת קבלני משנה בתדירות גבוהה."

    if "מסירה" in task_name or "בדיקות" in task_name or "חשבון" in task_name:
        return "להכין תוכנית מסירה מוקדמת, לפתוח רשימת ליקויים מתגלגלת ולהגדיר אחריות לסגירת כל סעיף."

    if "עפר" in task_name or "תשתיות" in task_name or "פיתוח" in task_name:
        return "לוודא סקר תשתיות מוקדם, להכין תוכנית ניהול סיכוני שטח ולבדוק זמינות ציוד וקבלני משנה."

    if risk_level == "גבוהה":
        return "להכין תוכנית פעולה מיידית: אחראי, תאריך יעד, פעולת מניעה ופעולת גיבוי."

    if main_reason == "פער גדול בין P90 לבין משך סביר":
        return "לבחון מחדש את אומדן המשך, להוסיף רזרבת זמן ולבדוק אפשרות להאצת הפעילות."

    if main_reason == "קריטיות גבוהה בנתיב הפרויקט":
        return "להכניס את הפעילות למעקב שבועי בישיבת בקרת פרויקט ולבדוק השפעה על הנתיב הקריטי."

    if main_reason == "השפעה גבוהה של סיכונים חיצוניים":
        return "לבחון פעולות הפחתת סיכון, ביטוח, ספק חלופי או שינוי שיטת ביצוע."

    return "להמשיך מעקב שוטף ולבדוק האם נדרש עדכון אומדן או רזרבה."


def build_action_recommendations_table(task_risk_table, top_n=12):
    top = task_risk_table.sort_values("ציון סיכון משולב", ascending=False).head(top_n).copy()

    rows = []

    for _, row in top.iterrows():
        main_reason = identify_main_risk_reason(row)
        action = recommend_action(row)

        rows.append({
            "מזהה פעילות": row["מזהה פעילות"],
            "שם פעילות": row["שם פעילות"],
            "שלב": row["שלב"],
            "אחראי": row["אחראי"],
            "רמת סיכון": row["רמת סיכון"],
            "ציון סיכון משולב": row["ציון סיכון משולב"],
            "מדד קריטיות": row["מדד קריטיות"],
            "פער P90 מול משך סביר": row["פער P90 מול משך סביר"],
            "סיבה מרכזית": main_reason,
            "פעולה מומלצת": action,
            "עדיפות טיפול": "מיידית" if row["רמת סיכון"] == "גבוהה" else "מעקב קרוב"
        })

    return pd.DataFrame(rows)


def create_management_insights(schedule_summary, financial_summary, task_risk_table, recommendations_table):
    insights = []

    insights.append(f"משך P50 של הפרויקט הוא כ-{schedule_summary['P50']:.1f} ימים.")
    insights.append(f"משך P85 הוא כ-{schedule_summary['P85']:.1f} ימים, ומשך P90 הוא כ-{schedule_summary['P90']:.1f} ימים.")
    insights.append(f"הסתברות האיחור מול משך היעד היא {schedule_summary['הסתברות איחור']:.1%}.")
    insights.append(f"הרזרבה הכספית המומלצת לפי רמת הביטחון היא כ-{format_money(financial_summary['רזרבה מומלצת לפי רמת הביטחון'])}.")
    insights.append(f"רזרבת הימים המומלצת היא כ-{financial_summary['רזרבת ימים לפי רמת הביטחון']:.1f} ימים.")

    top_task = task_risk_table.iloc[0]
    insights.append(f"הפעילות בעלת ציון הסיכון המשולב הגבוה ביותר היא '{top_task['שם פעילות']}', ברמת סיכון {top_task['רמת סיכון']}.")

    top_recommendation = recommendations_table.iloc[0]
    insights.append(f"הפעולה המומלצת הראשונה היא עבור '{top_recommendation['שם פעילות']}': {top_recommendation['פעולה מומלצת']}")

    return insights


def create_duration_distribution_chart(results, schedule_summary):
    durations = results["project_durations"]

    fig = px.histogram(
        x=durations,
        nbins=50,
        title="התפלגות משך הפרויקט לפי סימולציית Monte Carlo",
        labels={"x": "משך פרויקט בימים", "y": "מספר סימולציות"}
    )

    for label in ["P50", "P85", "P90", "משך יעד"]:
        if label in schedule_summary:
            fig.add_vline(
                x=schedule_summary[label],
                line_dash="dash",
                annotation_text=label
            )

    fig.update_layout(
        xaxis_title="משך פרויקט בימים",
        yaxis_title="מספר סימולציות"
    )

    return update_chart_layout(fig)


def create_s_curve_chart(results, schedule_summary):
    durations = np.sort(results["project_durations"])
    probabilities = np.arange(1, len(durations) + 1) / len(durations)

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=durations,
            y=probabilities,
            mode="lines",
            name="S-Curve"
        )
    )

    for label in ["P50", "P85", "P90", "משך יעד"]:
        if label in schedule_summary:
            fig.add_vline(
                x=schedule_summary[label],
                line_dash="dash",
                annotation_text=label
            )

    fig.update_layout(
        title="S-Curve הסתברותית לסיום הפרויקט",
        xaxis_title="משך פרויקט בימים",
        yaxis_title="הסתברות מצטברת לסיום",
        yaxis_tickformat=".0%"
    )

    return update_chart_layout(fig)


def create_percentile_summary_chart(schedule_summary):
    data = pd.DataFrame({
        "מדד": ["משך יעד", "P50", "P85", "P90"],
        "משך בימים": [
            schedule_summary["משך יעד"],
            schedule_summary["P50"],
            schedule_summary["P85"],
            schedule_summary["P90"]
        ]
    })

    fig = px.bar(
        data,
        x="מדד",
        y="משך בימים",
        title="השוואת משך יעד מול P50 / P85 / P90",
        text="משך בימים"
    )

    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig.update_layout(xaxis_title="מדד", yaxis_title="משך בימים")

    return update_chart_layout(fig)


def create_criticality_chart(task_risk_table, top_n=10):
    top = task_risk_table.sort_values("מדד קריטיות", ascending=False).head(top_n).copy()

    fig = px.bar(
        top,
        x="מדד קריטיות",
        y="שם פעילות",
        orientation="h",
        title=f"{top_n} הפעילויות הקריטיות ביותר לפי Criticality Index",
        hover_data=[
            "מזהה פעילות",
            "שלב",
            "אחראי",
            "פער P90 מול משך סביר",
            "ציון סיכון משולב"
        ]
    )

    fig.update_layout(
        xaxis_tickformat=".0%",
        yaxis={"categoryorder": "total ascending"}
    )

    return update_chart_layout(fig)


def create_combined_risk_score_chart(task_risk_table, top_n=10):
    top = task_risk_table.sort_values("ציון סיכון משולב", ascending=False).head(top_n).copy()

    fig = px.bar(
        top,
        x="ציון סיכון משולב",
        y="שם פעילות",
        orientation="h",
        title=f"{top_n} הפעילויות המסוכנות ביותר לפי ציון סיכון משולב",
        hover_data=[
            "מזהה פעילות",
            "שלב",
            "אחראי",
            "רמת סיכון",
            "מדד קריטיות",
            "פער P90 מול משך סביר",
            "תוספת ממוצעת מסיכונים"
        ]
    )

    fig.update_layout(yaxis={"categoryorder": "total ascending"})

    return update_chart_layout(fig)


def create_p90_gap_chart(task_risk_table, top_n=10):
    top = task_risk_table.sort_values("פער P90 מול משך סביר", ascending=False).head(top_n)

    fig = px.bar(
        top,
        x="פער P90 מול משך סביר",
        y="שם פעילות",
        orientation="h",
        title=f"{top_n} הפעילויות עם פער P90 הגבוה ביותר מול המשך הסביר",
        hover_data=[
            "מזהה פעילות",
            "שלב",
            "אחראי",
            "P90 משך פעילות"
        ]
    )

    fig.update_layout(yaxis={"categoryorder": "total ascending"})

    return update_chart_layout(fig)


def create_financial_distribution_chart(financial_arrays, financial_summary):
    costs = financial_arrays["total_delay_cost"]

    fig = px.histogram(
        x=costs,
        nbins=50,
        title="התפלגות עלות העיכוב והרזרבות לפי סימולציה",
        labels={"x": "עלות עיכוב ורזרבות ₪", "y": "מספר סימולציות"}
    )

    fig.add_vline(x=financial_summary["עלות עיכוב P50"], line_dash="dash", annotation_text="עלות P50")
    fig.add_vline(x=financial_summary["עלות עיכוב P85"], line_dash="dash", annotation_text="עלות P85")
    fig.add_vline(x=financial_summary["עלות עיכוב P90"], line_dash="dash", annotation_text="עלות P90")

    fig.update_layout(xaxis_title="עלות ₪", yaxis_title="מספר סימולציות")

    return update_chart_layout(fig)


def create_financial_components_chart(financial_arrays):
    data = pd.DataFrame({
        "רכיב": ["מימון", "תקורות", "מדד תשומות", "סיכונים כספיים"],
        "עלות ממוצעת": [
            np.mean(financial_arrays["financing_cost_over_target"]),
            np.mean(financial_arrays["overhead_delay_cost"]),
            np.mean(financial_arrays["indexation_cost_over_target"]),
            np.mean(financial_arrays["risk_costs"])
        ]
    })

    fig = px.bar(
        data,
        x="רכיב",
        y="עלות ממוצעת",
        title="פירוק עלות עיכוב ממוצעת לפי רכיבים",
        text="עלות ממוצעת"
    )

    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig.update_layout(yaxis_title="עלות ממוצעת ₪")

    return update_chart_layout(fig)


def create_risk_level_bar_chart(task_risk_table):
    data = (
        task_risk_table
        .groupby("רמת סיכון", observed=True)
        .size()
        .reset_index(name="מספר פעילויות")
    )

    fig = px.bar(
        data,
        x="רמת סיכון",
        y="מספר פעילויות",
        title="התפלגות פעילויות לפי רמת סיכון",
        text="מספר פעילויות"
    )

    fig.update_traces(textposition="outside")
    fig.update_layout(xaxis_title="רמת סיכון", yaxis_title="מספר פעילויות")

    return update_chart_layout(fig)


def build_results_excel(output):
    buffer = BytesIO()

    schedule_summary_df = pd.DataFrame(
        list(output["schedule_summary"].items()),
        columns=["מדד", "ערך"]
    )

    financial_summary_df = pd.DataFrame(
        list(output["financial_summary"].items()),
        columns=["מדד", "ערך"]
    )

    insights_df = pd.DataFrame({"תובנות ניהוליות": output["insights"]})

    simulation_df = pd.DataFrame({
        "משך פרויקט בימים": output["simulation_results"]["project_durations"],
        "ימי איחור": output["financial_arrays"]["delay_days"],
        "עלות מימון עודפת": output["financial_arrays"]["financing_cost_over_target"],
        "עלות תקורות עיכוב": output["financial_arrays"]["overhead_delay_cost"],
        "עלות מדד תשומות עודפת": output["financial_arrays"]["indexation_cost_over_target"],
        "עלות סיכונים כספית": output["financial_arrays"]["risk_costs"],
        "עלות עיכוב כוללת": output["financial_arrays"]["total_delay_cost"]
    })

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        output["tasks"].to_excel(writer, sheet_name="משימות מעובדות", index=False)
        output["links"].to_excel(writer, sheet_name="קשרים מעובדים", index=False)
        output["risks"].to_excel(writer, sheet_name="סיכונים מעובדים", index=False)
        schedule_summary_df.to_excel(writer, sheet_name="סיכום לוז", index=False)
        financial_summary_df.to_excel(writer, sheet_name="סיכום פיננסי", index=False)
        output["task_risk_table"].to_excel(writer, sheet_name="טבלת סיכונים", index=False)
        output["recommendations_table"].to_excel(writer, sheet_name="המלצות פעולה", index=False)
        insights_df.to_excel(writer, sheet_name="תובנות", index=False)
        simulation_df.to_excel(writer, sheet_name="תוצאות סימולציה", index=False)

    buffer.seek(0)
    return buffer


def run_project_control_core(tasks_raw, links_raw, risks_raw, finance_params, scenario, inputs=None):
    tasks = validate_tasks(tasks_raw)
    links = validate_links(links_raw, tasks)
    risks = validate_risks(risks_raw, tasks)

    n_simulations = scenario["n_simulations"]
    target_duration = scenario["target_duration"]
    confidence_level = scenario["confidence_level"]
    random_seed = int(finance_params["Seed אקראיות"])

    results = run_monte_carlo(
        tasks=tasks,
        links=links,
        risks=risks,
        n_simulations=n_simulations,
        random_seed=random_seed
    )

    schedule_summary = calculate_schedule_summary(results, target_duration)

    financial_summary, financial_arrays = calculate_financial_results(
        results=results,
        finance_params=finance_params,
        target_duration=target_duration,
        confidence_level=confidence_level
    )

    task_risk_table = build_task_risk_table(tasks, results)

    recommendations_table = build_action_recommendations_table(
        task_risk_table,
        top_n=12
    )

    insights = create_management_insights(
        schedule_summary,
        financial_summary,
        task_risk_table,
        recommendations_table
    )

    figures = {
        "התפלגות משך הפרויקט": create_duration_distribution_chart(results, schedule_summary),
        "S-Curve": create_s_curve_chart(results, schedule_summary),
        "השוואת אחוזונים": create_percentile_summary_chart(schedule_summary),
        "פעילויות קריטיות": create_criticality_chart(task_risk_table, top_n=10),
        "ציון סיכון משולב": create_combined_risk_score_chart(task_risk_table, top_n=10),
        "פער P90": create_p90_gap_chart(task_risk_table, top_n=10),
        "התפלגות עלויות": create_financial_distribution_chart(financial_arrays, financial_summary),
        "פירוק עלויות": create_financial_components_chart(financial_arrays),
        "התפלגות רמות סיכון": create_risk_level_bar_chart(task_risk_table)
    }

    return {
        "inputs": inputs,
        "tasks": tasks,
        "links": links,
        "risks": risks,
        "finance_params": finance_params,
        "scenario": scenario,
        "simulation_results": results,
        "schedule_summary": schedule_summary,
        "financial_summary": financial_summary,
        "financial_arrays": financial_arrays,
        "task_risk_table": task_risk_table,
        "recommendations_table": recommendations_table,
        "insights": insights,
        "figures": figures
    }


def run_full_project_control_analysis(file_path, scenario_name="בסיס"):
    inputs = load_all_inputs(file_path)

    finance_params = read_financial_parameters(inputs["finance"])
    scenario = choose_scenario(inputs["scenarios"], scenario_name=scenario_name)

    return run_project_control_core(
        tasks_raw=inputs["tasks"],
        links_raw=inputs["links"],
        risks_raw=inputs["risks"],
        finance_params=finance_params,
        scenario=scenario,
        inputs=inputs
    )


def project_risk_level_text(prob_delay):
    if prob_delay >= 0.70:
        return "קריטי", "status-critical"
    if prob_delay >= 0.40:
        return "גבוה", "status-high"
    if prob_delay >= 0.20:
        return "בינוני", "status-medium"
    return "נמוך", "status-good"


def render_project_status_card(schedule_summary, financial_summary):
    prob_delay = schedule_summary["הסתברות איחור"]
    risk_text, risk_class = project_risk_level_text(prob_delay)

    target = schedule_summary["משך יעד"]
    p50 = schedule_summary["P50"]
    p85 = schedule_summary["P85"]
    p90 = schedule_summary["P90"]
    reserve_days = financial_summary["רזרבת ימים לפי רמת הביטחון"]
    reserve_money = financial_summary["רזרבה מומלצת לפי רמת הביטחון"]

    st.markdown(
        f"""
        <div class="status-card {risk_class}" dir="rtl">
            סטטוס פרויקט: סיכון לו״ז {risk_text}
            <br>
            <span style="font-size:13px;font-weight:400;">
                משך יעד: {target:.1f} ימים |
                P50: {p50:.1f} ימים |
                P85: {p85:.1f} ימים |
                P90: {p90:.1f} ימים |
                רזרבת ימים מומלצת: {reserve_days:.1f} |
                רזרבה כספית מומלצת: {format_money(reserve_money)}
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )


def project_recommendation_card_html(row):
    risk_level = str(row["רמת סיכון"])

    if "גבוה" in risk_level:
        risk_class = "risk-high"
    elif "בינ" in risk_level:
        risk_class = "risk-medium"
    else:
        risk_class = "risk-low"

    return f"""
    <div class="recommendation-card" dir="rtl">
        <div class="recommendation-title" dir="rtl">
            {html.escape(str(row["שם פעילות"]))}
        </div>
        <div class="recommendation-meta" dir="rtl">
            מזהה: {html.escape(str(row["מזהה פעילות"]))} |
            שלב: {html.escape(str(row["שלב"]))} |
            אחראי: {html.escape(str(row["אחראי"]))} |
            רמת סיכון:
            <span class="risk-badge {risk_class}">{html.escape(risk_level)}</span> |
            עדיפות: {html.escape(str(row["עדיפות טיפול"]))}
        </div>
        <div class="recommendation-action" dir="rtl">
            <b>סיבה מרכזית:</b> {html.escape(str(row["סיבה מרכזית"]))}<br>
            <b>פעולה מומלצת:</b> {html.escape(str(row["פעולה מומלצת"]))}
        </div>
    </div>
    """


def render_project_recommendation_cards(recommendations_table, max_cards=6):
    top = recommendations_table.head(max_cards).reset_index(drop=True)

    for i in range(0, len(top), 2):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(project_recommendation_card_html(top.iloc[i]), unsafe_allow_html=True)

        if i + 1 < len(top):
            with col2:
                st.markdown(project_recommendation_card_html(top.iloc[i + 1]), unsafe_allow_html=True)



# ------------------------------------------------------------
# הזנה ידנית של נתוני הפרויקט (בלי קובץ Excel)
# ------------------------------------------------------------

def default_manual_task_df():
    return pd.DataFrame(
        [
            {
                "שם פעילות": "תכנון ראשוני",
                "משך משוער (ימים)": 10,
                "טווח אופטימי (ימים)": 0,
                "טווח פסימי (ימים)": 0,
                "עלות (₪)": 0
            }
        ]
    )


def default_manual_link_df():
    return pd.DataFrame(
        columns=[
            "פעילות קודמת (שם)",
            "פעילות עוקבת (שם)",
            "סוג קשר",
            "השהיה (ימים)"
        ]
    )


def default_manual_risk_df():
    return pd.DataFrame(
        columns=[
            "תיאור סיכון",
            "פעילות קשורה (שם)",
            "הסתברות %",
            "השפעת עיכוב משוערת (ימים)",
            "השפעה כספית (₪)",
            "פעיל?"
        ]
    )


UNCERTAINTY_LEVELS = {
    "אומדן ראשוני / תכנון מוקדם (טווח רחב)": (0.35, 0.60),
    "אומדן בינוני / תכנון מפורט (ברירת מחדל)": (0.20, 0.30),
    "אומדן מבוסס / אחרי מכרז או חוזה (טווח צר)": (0.10, 0.15),
}

DEFAULT_UNCERTAINTY_LEVEL = "אומדן בינוני / תכנון מפורט (ברירת מחדל)"


def build_manual_tasks_df(ui_df, uncertainty_level=DEFAULT_UNCERTAINTY_LEVEL):
    optimistic_pct, pessimistic_pct = UNCERTAINTY_LEVELS.get(
        uncertainty_level, UNCERTAINTY_LEVELS[DEFAULT_UNCERTAINTY_LEVEL]
    )

    df = ui_df.copy()
    df["שם פעילות"] = df["שם פעילות"].apply(clean_text)
    df = df[df["שם פעילות"] != ""].reset_index(drop=True)

    if df.empty:
        raise ValueError("צריך למלא לפחות פעילות אחת עם שם, בשלב 1.")

    duplicated = df[df["שם פעילות"].duplicated()]["שם פעילות"].tolist()

    if duplicated:
        raise ValueError(
            "יש לך יותר משם פעילות אחת עם אותו שם: " + ", ".join(duplicated) +
            ". צריך לתת לכל פעילות שם ייחודי."
        )

    rows = []

    for _, row in df.iterrows():
        duration = safe_float(row.get("משך משוער (ימים)", 0))

        if duration <= 0:
            raise ValueError(f"לפעילות '{row['שם פעילות']}' צריך למלא משך משוער גדול מ-0.")

        optimistic = safe_float(row.get("טווח אופטימי (ימים)", 0))
        pessimistic = safe_float(row.get("טווח פסימי (ימים)", 0))

        if optimistic <= 0:
            optimistic = round(duration * (1 - optimistic_pct), 1)

        if pessimistic <= 0:
            pessimistic = round(duration * (1 + pessimistic_pct), 1)

        optimistic = min(optimistic, duration)
        pessimistic = max(pessimistic, duration)

        rows.append({
            "task_id": row["שם פעילות"],
            "task_name": row["שם פעילות"],
            "stage": "",
            "task_type": "משימה",
            "owner": "",
            "category": "",
            "optimistic_days": optimistic,
            "most_likely_days": duration,
            "pessimistic_days": pessimistic,
            "base_cost": safe_float(row.get("עלות (₪)", 0))
        })

    return pd.DataFrame(rows)


def build_manual_links_df(ui_df, task_names, sequential_default):
    df = ui_df.copy()
    df["פעילות קודמת (שם)"] = df["פעילות קודמת (שם)"].apply(clean_text)
    df["פעילות עוקבת (שם)"] = df["פעילות עוקבת (שם)"].apply(clean_text)
    df = df[(df["פעילות קודמת (שם)"] != "") & (df["פעילות עוקבת (שם)"] != "")]

    if df.empty:
        if sequential_default and len(task_names) > 1:
            rows = []

            for i in range(len(task_names) - 1):
                rows.append({
                    "predecessor_id": task_names[i],
                    "successor_id": task_names[i + 1],
                    "relationship_type": "FS",
                    "lag_days": 0,
                    "note": ""
                })

            return pd.DataFrame(rows)

        return pd.DataFrame(
            columns=["predecessor_id", "successor_id", "relationship_type", "lag_days", "note"]
        )

    rows = []

    for _, row in df.iterrows():
        rel = clean_text(row.get("סוג קשר", "FS")).upper() or "FS"

        rows.append({
            "predecessor_id": row["פעילות קודמת (שם)"],
            "successor_id": row["פעילות עוקבת (שם)"],
            "relationship_type": rel,
            "lag_days": safe_float(row.get("השהיה (ימים)", 0)),
            "note": ""
        })

    return pd.DataFrame(rows)


def build_manual_risks_df(ui_df):
    df = ui_df.copy()
    df["פעילות קשורה (שם)"] = df["פעילות קשורה (שם)"].apply(clean_text)
    df = df[df["פעילות קשורה (שם)"] != ""].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(
            columns=[
                "risk_id", "risk_description", "related_task_id", "probability",
                "min_delay_days", "most_likely_delay_days", "max_delay_days",
                "cost_impact", "is_active"
            ]
        )

    rows = []

    for i, row in df.iterrows():
        delay = safe_float(row.get("השפעת עיכוב משוערת (ימים)", 0))
        active = clean_text(row.get("פעיל?", "כן")) or "כן"

        rows.append({
            "risk_id": f"R{i + 1}",
            "risk_description": clean_text(row.get("תיאור סיכון", "")),
            "related_task_id": row["פעילות קשורה (שם)"],
            "probability": safe_percent(row.get("הסתברות %", 0)),
            "min_delay_days": round(delay * 0.5, 1),
            "most_likely_delay_days": delay,
            "max_delay_days": round(delay * 1.5, 1),
            "cost_impact": safe_float(row.get("השפעה כספית (₪)", 0)),
            "is_active": active
        })

    return pd.DataFrame(rows)


def is_untouched_default_manual_data(tasks_ui_df, links_ui_df, risks_ui_df, target_duration, confidence_level):
    """
    בודק אם המשתמש הריץ ניתוח מבלי לשנות כלום מנתוני ברירת המחדל/הדוגמה —
    כדי להזהיר שהתוצאות לא משקפות פרויקט אמיתי.
    """
    try:
        tasks_match = tasks_ui_df.reset_index(drop=True).equals(
            default_manual_task_df().reset_index(drop=True)
        )
    except Exception:
        tasks_match = False

    try:
        links_empty = links_ui_df.dropna(how="all").empty
    except Exception:
        links_empty = True

    try:
        risks_empty = risks_ui_df.dropna(how="all").empty
    except Exception:
        risks_empty = True

    try:
        targets_match = (float(target_duration) == 30.0) and (float(confidence_level) == 0.85)
    except Exception:
        targets_match = False

    return tasks_match and links_empty and risks_empty and targets_match


def render_project_controls_manual():
    st.write(
        "ממלאים את נתוני הפרויקט ישירות כאן באתר — בלי צורך בקובץ Excel. "
        "המערכת מריצה סימולציית Monte Carlo ומציגה P50/P85/P90, הסתברות איחור, רזרבות, גרפים והמלצות פעולה."
    )

    if "manual_task_df" not in st.session_state:
        st.session_state["manual_task_df"] = default_manual_task_df()

    if "manual_link_df" not in st.session_state:
        st.session_state["manual_link_df"] = default_manual_link_df()

    if "manual_risk_df" not in st.session_state:
        st.session_state["manual_risk_df"] = default_manual_risk_df()

    st.info(
        "חשוב: מלא את הנתונים בטבלאות ואז לחץ על כפתור החישוב בתחתית הטופס. "
        "הערכים נשמרים רק לאחר לחיצה על הכפתור, כדי למנוע מחיקה בזמן ההקלדה."
    )

    with st.form("manual_project_form", clear_on_submit=False):
        st.subheader("שלב 1 — פעילויות הפרויקט")
        st.caption(
            "כל שורה היא פעילות. חובה למלא שם ומשך משוער. "
            "טווח אופטימי/פסימי ועלות הם אופציונליים — אם משאירים 0, המערכת משלימה אוטומטית לפי רמת "
            "האי-ודאות שתבחר למטה (בהתאם לעקרון אומדן שלוש הנקודות / PERT מתוך PMBOK)."
        )

        uncertainty_level = st.selectbox(
            "רמת אי-ודאות באומדן (למילוי אוטומטי של טווח אופטימי/פסימי כשלא מזינים אותם ידנית)",
            options=list(UNCERTAINTY_LEVELS.keys()),
            index=list(UNCERTAINTY_LEVELS.keys()).index(DEFAULT_UNCERTAINTY_LEVEL),
            help=(
                "ככל שהאומדן מוקדם יותר בפרויקט, כך נהוג לתת לו טווח רחב יותר בין אופטימי לפסימי "
                "(בדומה לעקרון סיווגי דיוק אומדן מקובלים בענף הבנייה והתשתיות). "
                "הטווח א-סימטרי בכוונה — בפועל עיכובים גדולים שכיחים יותר מהקדמות גדולות."
            )
        )

        tasks_ui_df = st.data_editor(
            st.session_state["manual_task_df"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="manual_task_editor",
            column_config={
                "משך משוער (ימים)": st.column_config.NumberColumn("משך משוער (ימים)", min_value=0, step=1),
                "טווח אופטימי (ימים)": st.column_config.NumberColumn(
                    "טווח אופטימי (ימים)", min_value=0, step=1, help="השאר 0 כדי שהמערכת תחשב אוטומטית"
                ),
                "טווח פסימי (ימים)": st.column_config.NumberColumn(
                    "טווח פסימי (ימים)", min_value=0, step=1, help="השאר 0 כדי שהמערכת תחשב אוטומטית"
                ),
                "עלות (₪)": st.column_config.NumberColumn("עלות (₪)", min_value=0, step=1000),
            }
        )

        st.subheader("שלב 2 — סדר הפעילויות")

        sequential_default = st.checkbox(
            "הפעילויות מבוצעות ברצף, בדיוק לפי הסדר שהזנתי למעלה (מומלץ)",
            value=True,
            help="אם תבטל את הסימון ולא תגדיר קשרים משלך בטבלה למטה, הפעילויות ייחשבו כמבוצעות במקביל."
        )

        with st.expander("קשרים מותאמים בין פעילויות (אופציונלי — למי שרוצה שליטה מדויקת)"):
            st.caption(
                "למלא רק אם הסדר לא פשוט רציף — למשל פעילות שמתחילה יחד עם קודמתה. "
                "בוחרים שם פעילות בדיוק כפי שהוקלד בשלב 1. אם משאירים ריק, ייעשה שימוש בסימון למעלה."
            )
            links_ui_df = st.data_editor(
                st.session_state["manual_link_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="manual_link_editor",
                column_config={
                    "סוג קשר": st.column_config.SelectboxColumn("סוג קשר", options=["FS", "SS", "FF", "SF"]),
                    "השהיה (ימים)": st.column_config.NumberColumn("השהיה (ימים)", step=1),
                }
            )

        st.subheader("שלב 3 — יעד ורמת ביטחון")

        col1, col2 = st.columns(2)

        with col1:
            target_duration = st.number_input(
                "משך יעד לפרויקט (ימים) — מולו נבדקת הסתברות לאיחור",
                min_value=1,
                value=30,
                step=1
            )

        with col2:
            confidence_level = st.slider(
                "רמת ביטחון לחישוב רזרבה",
                min_value=0.5,
                max_value=0.99,
                value=0.85,
                step=0.01,
                help="0.85 = רזרבה שתספיק ב-85% מהמקרים (P85)"
            )

        with st.expander("נתונים כספיים (אופציונלי — אם לא ממלאים, מדלגים על ניתוח כספי)"):
            fin_col1, fin_col2 = st.columns(2)

            with fin_col1:
                base_budget = st.number_input("תקציב בסיסי (₪)", min_value=0, value=0, step=10000)
                overhead_daily = st.number_input("עלות תקורה ליום עיכוב (₪)", min_value=0, value=0, step=100)

            with fin_col2:
                annual_interest = st.number_input(
                    "ריבית מימון שנתית (שיעור עשרוני, למשל 0.07 = 7%)",
                    min_value=0.0, value=0.0, step=0.01, format="%.3f"
                )
                annual_index = st.number_input(
                    "מדד תשומות שנתי (שיעור עשרוני, למשל 0.035 = 3.5%)",
                    min_value=0.0, value=0.0, step=0.005, format="%.3f"
                )

        st.subheader("שלב 4 — סיכונים (אופציונלי)")

        with st.expander("סיכונים בפרויקט"):
            st.caption(
                "כל שורה היא סיכון שמקושר לפעילות אחת. אם אין סיכונים לרישום, אפשר להשאיר ריק."
            )
            risks_ui_df = st.data_editor(
                st.session_state["manual_risk_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="manual_risk_editor",
                column_config={
                    "הסתברות %": st.column_config.NumberColumn("הסתברות %", min_value=0, max_value=100, step=1),
                    "השפעה כספית (₪)": st.column_config.NumberColumn("השפעה כספית (₪)", min_value=0, step=1000),
                    "פעיל?": st.column_config.SelectboxColumn("פעיל?", options=["כן", "לא"]),
                }
            )

        st.subheader("שלב 5 — הגדרות מתקדמות (אופציונלי)")

        with st.expander("הגדרות מתקדמות"):
            adv_col1, adv_col2 = st.columns(2)

            with adv_col1:
                n_simulations = st.number_input(
                    "מספר סימולציות Monte Carlo",
                    min_value=1000,
                    max_value=50000,
                    value=10000,
                    step=1000,
                    help="כמה תרחישים אקראיים להריץ. יותר = תוצאה מדויקת יותר אך חישוב איטי יותר. ברירת המחדל מתאימה לרוב המקרים."
                )

            with adv_col2:
                random_seed = st.number_input(
                    "Seed אקראיות",
                    min_value=1,
                    value=42,
                    step=1,
                    help="מספר שקובע את נקודת ההתחלה של הגרלת התרחישים. אותו Seed עם אותם נתונים ייתן תמיד את אותה תוצאה — שימושי לשחזור חישוב קודם. אין צורך לשנות אותו."
                )

        confirm_demo_data = st.checkbox(
            "ידוע לי שלא ערכתי את שורת הדוגמה בשלב 1, ואני רוצה להריץ בכל זאת ניתוח הדגמה בלבד "
            "(לא ניתוח של פרויקט אמיתי)"
        )

        run_manual = st.form_submit_button(
            "הרץ ניתוח פרויקט",
            use_container_width=True
        )

    if not run_manual:
        st.info("מלא את הנתונים למעלה ולחץ על 'הרץ ניתוח פרויקט'.")
        return

    st.session_state["manual_task_df"] = tasks_ui_df
    st.session_state["manual_link_df"] = links_ui_df
    st.session_state["manual_risk_df"] = risks_ui_df

    if is_untouched_default_manual_data(
        tasks_ui_df, links_ui_df, risks_ui_df, target_duration, confidence_level
    ) and not confirm_demo_data:
        st.warning(
            "⚠️ **שים לב — אלו נתוני הדוגמה, לא נתוני הפרויקט שלך.**\n\n"
            "הטבלה בשלב 1 עדיין מכילה רק את השורה המובנית לדוגמה ('תכנון ראשוני', 10 ימים), "
            "ומשך היעד ורמת הביטחון עדיין בערכי ברירת המחדל (30 ימים, 85%). "
            "אם תריץ עכשיו, התוצאות (KPI-ים, אחוזונים, המלצות) יתארו את הדוגמה הזו בלבד — "
            "ולא ישקפו שום דבר אמיתי על הפרויקט שלך, גם אם הן ייראו מלאות ומקצועיות.\n\n"
            "**כדי לקבל ניתוח אמיתי:** ערוך את הטבלה בשלב 1 עם הפעילויות האמיתיות של הפרויקט שלך "
            "(שמות, משכים, ואם ידוע — טווחים אופטימי/פסימי ועלויות), ועדכן את משך היעד ורמת הביטחון בשלב 3 לפי הפרויקט. "
            "לאחר מכן לחץ שוב על 'הרץ ניתוח פרויקט'.\n\n"
            "אם רצית רק לראות הדגמה של איך הכלי עובד — סמן את התיבה מעל הכפתור ולחץ שוב."
        )
        return

    try:
        tasks_raw = build_manual_tasks_df(tasks_ui_df, uncertainty_level=uncertainty_level)
        task_names = tasks_raw["task_id"].tolist()
        links_raw = build_manual_links_df(links_ui_df, task_names, sequential_default)
        risks_raw = build_manual_risks_df(risks_ui_df)

        finance_params = {
            "תקציב בסיסי": base_budget,
            "ריבית שנתית": annual_interest,
            "עלות תקורה יומית": overhead_daily,
            "מדד תשומות שנתי": annual_index,
            "רמת ביטחון לרזרבה": confidence_level,
            "משך יעד בימים": target_duration,
            "מספר סימולציות": n_simulations,
            "Seed אקראיות": random_seed
        }

        scenario = {
            "scenario_name": "הזנה ידנית",
            "n_simulations": int(n_simulations),
            "target_duration": float(target_duration),
            "confidence_level": float(confidence_level)
        }

    except Exception as e:
        st.error("יש בעיה בנתונים שהוזנו:")
        st.exception(e)
        return

    with st.spinner("מריץ סימולציית Monte Carlo ומחשב תוצאות..."):
        try:
            output = run_project_control_core(
                tasks_raw=tasks_raw,
                links_raw=links_raw,
                risks_raw=risks_raw,
                finance_params=finance_params,
                scenario=scenario
            )

        except Exception as e:
            st.error("אירעה שגיאה במהלך הניתוח.")
            st.exception(e)
            return

    if confirm_demo_data:
        st.warning(
            "⚠️ התוצאות שלמטה מבוססות על נתוני הדוגמה בלבד (לא נתוני פרויקט אמיתיים) — "
            "זו הרצת הדגמה שביקשת."
        )

    render_project_control_results(output)


# ------------------------------------------------------------
# הזנה מקובץ Excel (מסלול משני, למי שכבר יש לו קובץ מוכן)
# ------------------------------------------------------------

def render_project_controls_file_upload():
    st.write(
        "מסלול זה מיועד למי שכבר יש לו קובץ Excel בתבנית של המערכת. "
        "המערכת מריצה סימולציית Monte Carlo ומציגה P50/P85/P90, הסתברות איחור, רזרבות, גרפים והמלצות פעולה."
    )

    uploaded_file = st.file_uploader(
        "העלה קובץ Excel של הפרויקט",
        type=["xlsx"],
        key="project_file_uploader"
    )

    if uploaded_file is None:
        st.info("העלה קובץ Excel כדי להתחיל.")
        return

    try:
        preview_inputs = load_all_inputs(uploaded_file)
        scenarios_df = preview_inputs["scenarios"].copy()
        scenarios_df["scenario_name"] = scenarios_df["scenario_name"].astype(str).str.strip()
        scenario_options = scenarios_df["scenario_name"].tolist()

    except Exception as e:
        st.error("הקובץ לא נקרא בהצלחה.")
        st.exception(e)
        return

    col_a, col_b = st.columns([1, 2])

    with col_a:
        selected_scenario = st.selectbox(
            "בחר תרחיש להרצה",
            scenario_options,
            index=0
        )

    with col_b:
        st.write("")
        st.write("")
        run_button = st.button(
            "הרץ ניתוח פרויקט",
            use_container_width=True
        )

    if not run_button:
        st.warning("בחר תרחיש ולחץ על כפתור הרצת הניתוח.")
        return

    with st.spinner("מריץ סימולציית Monte Carlo ומחשב תוצאות..."):
        try:
            output = run_full_project_control_analysis(
                uploaded_file,
                scenario_name=selected_scenario
            )

        except Exception as e:
            st.error("אירעה שגיאה במהלך הניתוח.")
            st.exception(e)
            return

    render_project_control_results(output)


def render_project_controls_module():
    st.header("בקרת פרויקט וניהול לו״ז")

    render_project_controls_manual()

    st.divider()

    with st.expander("יש לך כבר קובץ Excel מוכן? אפשר להעלות אותו במקום למלא ידנית"):
        render_project_controls_file_upload()


def render_project_control_results(output):
    schedule_summary = output["schedule_summary"]
    financial_summary = output["financial_summary"]

    st.success("הניתוח הסתיים בהצלחה.")

    st.subheader("דשבורד מנהלים — מדדי מפתח")

    kpi_data = [
        ("P50", format_days(schedule_summary["P50"]), "משך חציוני צפוי"),
        ("P85", format_days(schedule_summary["P85"]), "משך ברמת ביטחון גבוהה"),
        ("P90", format_days(schedule_summary["P90"]), "תרחיש שמרני יותר"),
        ("הסתברות איחור", format_percent(schedule_summary["הסתברות איחור"]), "מול משך היעד"),
        ("רזרבת ימים", format_days(financial_summary["רזרבת ימים לפי רמת הביטחון"]), "לפי רמת הביטחון"),
        ("רזרבה כספית", format_money(financial_summary["רזרבה מומלצת לפי רמת הביטחון"]), "לפי רמת הביטחון"),
    ]

    render_kpi_cards(kpi_data, columns=6)
    render_project_status_card(schedule_summary, financial_summary)

    tabs = st.tabs([
        "סיכום מנהלים",
        "גרפים",
        "טבלת סיכונים",
        "המלצות פעולה",
        "נתוני קלט",
        "הורדת דוח"
    ])

    with tabs[0]:
        st.subheader("סיכום מנהלים")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### סיכום לו״ז הסתברותי")

            schedule_df = pd.DataFrame(
                [
                    ["משך ממוצע", format_days(schedule_summary["משך ממוצע"])],
                    ["סטיית תקן", format_days(schedule_summary["סטיית תקן"])],
                    ["P50", format_days(schedule_summary["P50"])],
                    ["P85", format_days(schedule_summary["P85"])],
                    ["P90", format_days(schedule_summary["P90"])],
                    ["משך יעד", format_days(schedule_summary["משך יעד"])],
                    ["הסתברות איחור", format_percent(schedule_summary["הסתברות איחור"])]
                ],
                columns=["מדד", "ערך"]
            )

            display_summary_table(schedule_df)

        with col2:
            st.markdown("### סיכום פיננסי")

            financial_df = pd.DataFrame(
                [
                    ["תקציב בסיסי", format_money(financial_summary["תקציב בסיסי"])],
                    ["ריבית שנתית", format_percent(financial_summary["ריבית שנתית"])],
                    ["עלות תקורה יומית", format_money(financial_summary["עלות תקורה יומית"])],
                    ["מדד תשומות שנתי", format_percent(financial_summary["מדד תשומות שנתי"])],
                    ["עלות עיכוב ממוצעת", format_money(financial_summary["עלות עיכוב ממוצעת"])],
                    ["עלות עיכוב P50", format_money(financial_summary["עלות עיכוב P50"])],
                    ["עלות עיכוב P85", format_money(financial_summary["עלות עיכוב P85"])],
                    ["עלות עיכוב P90", format_money(financial_summary["עלות עיכוב P90"])],
                    ["רזרבה מומלצת", format_money(financial_summary["רזרבה מומלצת לפי רמת הביטחון"])],
                    ["רזרבת ימים", format_days(financial_summary["רזרבת ימים לפי רמת הביטחון"])]
                ],
                columns=["מדד", "ערך"]
            )

            display_summary_table(financial_df)

        st.divider()

        st.subheader("תובנות ניהוליות")
        render_insights(output["insights"])

        st.divider()

        st.subheader("פעולות מומלצות ראשונות")
        render_project_recommendation_cards(output["recommendations_table"], max_cards=4)

    with tabs[1]:
        st.subheader("גרפים וניתוח ויזואלי")

        chart_col1, chart_col2 = st.columns(2)
        chart_items = list(output["figures"].items())

        for idx, (title, fig) in enumerate(chart_items):
            if idx % 2 == 0:
                with chart_col1:
                    st.markdown(f"### {title}")
                    st.plotly_chart(fig, use_container_width=True)
            else:
                with chart_col2:
                    st.markdown(f"### {title}")
                    st.plotly_chart(fig, use_container_width=True)

    with tabs[2]:
        st.subheader("טבלת סיכונים לפי פעילות")
        st.dataframe(
            output["task_risk_table"],
            use_container_width=True,
            hide_index=True
        )

    with tabs[3]:
        st.subheader("המלצות פעולה לפי פעילות")

        render_project_recommendation_cards(output["recommendations_table"], max_cards=8)

        st.markdown("### טבלת המלצות מלאה")

        st.dataframe(
            output["recommendations_table"],
            use_container_width=True,
            hide_index=True
        )

    with tabs[4]:
        st.subheader("נתוני קלט מעובדים")

        input_tab1, input_tab2, input_tab3 = st.tabs([
            "משימות",
            "קשרים",
            "סיכונים"
        ])

        with input_tab1:
            st.dataframe(
                output["tasks"],
                use_container_width=True,
                hide_index=True
            )

        with input_tab2:
            st.dataframe(
                output["links"],
                use_container_width=True,
                hide_index=True
            )

        with input_tab3:
            st.dataframe(
                output["risks"],
                use_container_width=True,
                hide_index=True
            )

    with tabs[5]:
        st.subheader("הורדת קובץ תוצאות")

        excel_buffer = build_results_excel(output)

        st.download_button(
            label="הורד קובץ תוצאות Excel",
            data=excel_buffer,
            file_name="תוצאות_פלטפורמת_בקרת_פרויקט.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown(
            """
            <div class="footer-note" dir="rtl">
            הערה: המודל מבוסס על הנתונים שהוזנו בקובץ האקסל. בפרויקט אמיתי יש לוודא
            שהאומדנים, הסיכונים, הקשרים והפרמטרים הפיננסיים נבדקו ואושרו על ידי מנהל הפרויקט.
            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# ממשק ראשי
# ============================================================

render_hero()

with st.sidebar:
    st.header("ניווט")
    st.write("בחר את סוג הניתוח שתרצה לבצע.")
    st.divider()
    st.caption("המערכת כוללת שני מסלולים: בדיקת כדאיות מכרזים ובקרת פרויקטים.")

analysis_type = st.radio(
    "בחר סוג ניתוח",
    [
        "בדיקת כדאיות מכרזים",
        "בקרת פרויקט וניהול לו״ז"
    ],
    horizontal=True
)

st.divider()

if analysis_type == "בדיקת כדאיות מכרזים":
    render_tender_module()
else:
    render_project_controls_module()
