# ============================================================
# Project Controls & Optimization Platform
# Streamlit Web App
# גרסה מאוחדת, מקצועית ומתוקנת לשלב פיתוח מקומי
# ============================================================

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from io import BytesIO
from collections import deque

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# 1. הגדרות עמוד ועיצוב
# ============================================================

st.set_page_config(
    page_title="Project Controls Platform",
    page_icon="📊",
    layout="wide"
)

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

    section[data-testid="stSidebar"] {
        direction: rtl;
        text-align: right;
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
        font-size: 30px;
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
        padding: 20px;
        margin-top: 10px;
        margin-bottom: 18px;
        font-size: 16px;
        font-weight: 700;
    }

    .status-small {
        direction: rtl;
        text-align: right;
        font-size: 13px;
        font-weight: 400;
        margin-top: 6px;
        line-height: 1.7;
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

    .summary-table tr:nth-child(even) {
        background-color: #fafafa;
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
        min-height: 170px;
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

    .stTabs [data-baseweb="tab-list"] {
        direction: rtl;
        gap: 4px;
    }

    .stTabs [data-baseweb="tab"] {
        font-family: Arial, sans-serif;
        font-size: 15px;
        padding: 10px 16px;
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
# 2. שמות גיליונות ועמודות
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


# ============================================================
# 3. פונקציות עזר
# ============================================================

def format_money(value):
    return f"₪{value:,.0f}"


def format_days(value):
    return f"{value:,.1f} ימים"


def format_percent(value):
    return f"{value:.1%}"


def normalize_series(series):
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    min_val = s.min()
    max_val = s.max()

    if max_val == min_val:
        return s * 0

    return (s - min_val) / (max_val - min_val)


def display_summary_table(df):
    html = df.to_html(index=False, escape=False)
    st.markdown(
        f"""
        <div class="summary-table" dir="rtl">
            {html}
        </div>
        """,
        unsafe_allow_html=True
    )


def risk_level_text(prob_delay):
    if prob_delay >= 0.70:
        return "קריטי", "status-critical"
    if prob_delay >= 0.40:
        return "גבוה", "status-high"
    if prob_delay >= 0.20:
        return "בינוני", "status-medium"
    return "נמוך", "status-good"


def risk_badge_class(risk_level):
    risk_level = str(risk_level)
    if "גבוה" in risk_level:
        return "risk-high"
    if "בינ" in risk_level:
        return "risk-medium"
    return "risk-low"


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


# ============================================================
# 4. בדיקות תקינות
# ============================================================

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


# ============================================================
# 5. גרף קשרים וסימולציה
# ============================================================

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

    active_risks = risks[risks["is_active"] == "כן"].copy()

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


# ============================================================
# 6. חישובי סיכום
# ============================================================

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


# ============================================================
# 7. המלצות פעולה
# ============================================================

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


# ============================================================
# 8. גרפים
# ============================================================

def update_chart_layout(fig):
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Arial", size=13),
        title_font=dict(size=20),
        margin=dict(l=30, r=30, t=70, b=40),
        hovermode="closest"
    )

    return fig


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


# ============================================================
# 9. ייצוא לאקסל
# ============================================================

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


# ============================================================
# 10. הפונקציה הראשית
# ============================================================

def run_full_project_control_analysis(file_path, scenario_name="בסיס"):
    inputs = load_all_inputs(file_path)

    tasks = validate_tasks(inputs["tasks"])
    links = validate_links(inputs["links"], tasks)
    risks = validate_risks(inputs["risks"], tasks)
    finance_params = read_financial_parameters(inputs["finance"])

    scenario = choose_scenario(inputs["scenarios"], scenario_name=scenario_name)

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


# ============================================================
# 11. רכיבי ממשק
# ============================================================

def render_hero():
    st.markdown(
        """
        <div class="hero-box" dir="rtl">
            <div class="hero-title" dir="rtl">📊 פלטפורמת בקרת פרויקטים ואופטימיזציה</div>
            <div class="hero-subtitle" dir="rtl">
                מערכת לניתוח לו״ז, סיכונים ורזרבות בפרויקטי בנייה ותשתיות באמצעות
                Monte Carlo, התפלגות PERT, מדדי P50/P85/P90 ומודל פיננסי לרזרבות.
            </div>
            <span class="hero-badge">Project Controls</span>
            <span class="hero-badge">Monte Carlo</span>
            <span class="hero-badge">Risk Analytics</span>
            <span class="hero-badge">Financial Reserve</span>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_kpi_cards(schedule_summary, financial_summary):
    kpi_data = [
        ("P50", format_days(schedule_summary["P50"]), "משך חציוני צפוי"),
        ("P85", format_days(schedule_summary["P85"]), "משך ברמת ביטחון גבוהה"),
        ("P90", format_days(schedule_summary["P90"]), "תרחיש שמרני יותר"),
        ("הסתברות איחור", format_percent(schedule_summary["הסתברות איחור"]), "מול משך היעד"),
        ("רזרבת ימים", format_days(financial_summary["רזרבת ימים לפי רמת הביטחון"]), "לפי רמת הביטחון"),
        ("רזרבה כספית", format_money(financial_summary["רזרבה מומלצת לפי רמת הביטחון"]), "לפי רמת הביטחון"),
    ]

    cols = st.columns(6)

    for col, (label, value, note) in zip(cols, kpi_data):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card" dir="rtl">
                    <div class="kpi-label" dir="rtl">{label}</div>
                    <div class="kpi-value" dir="rtl">{value}</div>
                    <div class="kpi-note" dir="rtl">{note}</div>
                </div>
                """,
                unsafe_allow_html=True
            )


def render_status_card(schedule_summary, financial_summary):
    prob_delay = schedule_summary["הסתברות איחור"]
    risk_text, risk_class = risk_level_text(prob_delay)

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
            <div class="status-small" dir="rtl">
                משך יעד: {target:.1f} ימים |
                P50: {p50:.1f} ימים |
                P85: {p85:.1f} ימים |
                P90: {p90:.1f} ימים |
                רזרבת ימים מומלצת: {reserve_days:.1f} |
                רזרבה כספית מומלצת: {format_money(reserve_money)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_insights(insights):
    for i, insight in enumerate(insights, start=1):
        st.markdown(
            f"""
            <div class="insight-box" dir="rtl">
                <b>{i}.</b> {insight}
            </div>
            """,
            unsafe_allow_html=True
        )


def recommendation_card_html(row):
    risk_class = risk_badge_class(row["רמת סיכון"])

    return f"""
    <div class="recommendation-card" dir="rtl">
        <div class="recommendation-title" dir="rtl">
            {row["שם פעילות"]}
        </div>
        <div class="recommendation-meta" dir="rtl">
            מזהה: {row["מזהה פעילות"]} |
            שלב: {row["שלב"]} |
            אחראי: {row["אחראי"]} |
            רמת סיכון:
            <span class="risk-badge {risk_class}">{row["רמת סיכון"]}</span> |
            עדיפות: {row["עדיפות טיפול"]}
        </div>
        <div class="recommendation-action" dir="rtl">
            <b>סיבה מרכזית:</b> {row["סיבה מרכזית"]}<br>
            <b>פעולה מומלצת:</b> {row["פעולה מומלצת"]}
        </div>
    </div>
    """


def render_recommendation_cards(recommendations_table, max_cards=6, two_columns=True):
    top = recommendations_table.head(max_cards).reset_index(drop=True)

    if not two_columns:
        for _, row in top.iterrows():
            st.markdown(recommendation_card_html(row), unsafe_allow_html=True)
        return

    for i in range(0, len(top), 2):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(recommendation_card_html(top.iloc[i]), unsafe_allow_html=True)

        if i + 1 < len(top):
            with col2:
                st.markdown(recommendation_card_html(top.iloc[i + 1]), unsafe_allow_html=True)


# ============================================================
# 12. ממשק המשתמש הראשי
# ============================================================

render_hero()

with st.sidebar:
    st.header("⚙️ הגדרות הרצה")
    st.write("1. העלה קובץ Excel")
    st.write("2. בחר תרחיש")
    st.write("3. לחץ על הרצת ניתוח")
    st.divider()
    st.caption("הקובץ חייב לכלול את הגיליונות: משימות, קשרים, פרמטרים פיננסיים, סיכונים, תרחישים.")

uploaded_file = st.file_uploader(
    "העלה קובץ Excel של הפרויקט",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("העלה קובץ Excel כדי להתחיל.")
    st.stop()

try:
    preview_inputs = load_all_inputs(uploaded_file)
    scenarios_df = preview_inputs["scenarios"].copy()
    scenarios_df["scenario_name"] = scenarios_df["scenario_name"].astype(str).str.strip()
    scenario_options = scenarios_df["scenario_name"].tolist()

except Exception as e:
    st.error("הקובץ לא נקרא בהצלחה.")
    st.exception(e)
    st.stop()

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
        "🚀 הרץ ניתוח פרויקט",
        use_container_width=True
    )

if not run_button:
    st.warning("בחר תרחיש ולחץ על כפתור הרצת הניתוח.")
    st.stop()

with st.spinner("מריץ סימולציית Monte Carlo ומחשב תוצאות..."):
    try:
        output = run_full_project_control_analysis(
            uploaded_file,
            scenario_name=selected_scenario
        )

    except Exception as e:
        st.error("אירעה שגיאה במהלך הניתוח.")
        st.exception(e)
        st.stop()

schedule_summary = output["schedule_summary"]
financial_summary = output["financial_summary"]

st.success("הניתוח הסתיים בהצלחה.")

st.subheader("דשבורד מנהלים — מדדי מפתח")
render_kpi_cards(schedule_summary, financial_summary)
render_status_card(schedule_summary, financial_summary)


# ============================================================
# 13. טאבים
# ============================================================

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
    render_recommendation_cards(output["recommendations_table"], max_cards=4, two_columns=True)


with tabs[1]:
    st.subheader("גרפים וניתוח ויזואלי")
    st.caption("הגרפים מציגים את התפלגות משך הפרויקט, רמות ביטחון, קריטיות, סיכון משולב ועלויות.")

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
    st.caption("הטבלה מדרגת פעילויות לפי קריטיות, פער P90, תוספת מסיכונים, סטיית תקן וציון סיכון משולב.")

    st.dataframe(
        output["task_risk_table"],
        use_container_width=True,
        hide_index=True
    )


with tabs[3]:
    st.subheader("המלצות פעולה לפי פעילות")
    st.caption("המערכת מייצרת המלצות פעולה ראשוניות לפי מאפייני הפעילות ומדדי הסיכון.")

    render_recommendation_cards(output["recommendations_table"], max_cards=8, two_columns=True)

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

    st.markdown(
        """
        הקובץ כולל את כל תוצאות הניתוח: סיכום לו״ז, סיכום פיננסי,
        טבלת סיכונים, המלצות פעולה, תובנות ותוצאות סימולציה.
        """
    )

    excel_buffer = build_results_excel(output)

    st.download_button(
        label="⬇️ הורד קובץ תוצאות Excel",
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