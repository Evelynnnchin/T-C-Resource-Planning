import csv
import hashlib
import io
import re
from typing import List, Tuple

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Resource Planning Load & Hiring Tool", layout="wide")

MONTHS = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}

UNASSIGNED_NAMES = {"", "TBC", "TBD", "NA", "N/A", "NONE", "-", "NIL", "UNASSIGNED"}
PROJECT_HEADERS = {"JRL", "CRL", "RTS"}


# -----------------------------
# Helper functions
# -----------------------------
def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    if text.lower() == "nan":
        return ""
    return text


def to_number(value) -> float:
    text = clean_text(value)
    if text == "":
        return 0.0

    text = text.replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)

    if text in {"", ".", "-", "-."}:
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def to_optional_number(value):
    text = clean_text(value)
    if text == "":
        return np.nan

    number = to_number(text)
    return number if not np.isnan(number) else np.nan


def parse_year(value):
    text = clean_text(value)
    match = re.search(r"(19|20)\d{2}", text)
    if match:
        return int(match.group(0))
    return None


def parse_month(value):
    text = clean_text(value).upper()[:3]
    return MONTHS.get(text)


def parse_requirement(requirement: str) -> Tuple[float, float, float]:
    """
    Extract pax, hours/day and days from strings like:
    - 1 pax/8h/5 days
    - 4 pax/8h/2days
    - 3pax/8h/5 days (2 shifts)
    - 2 pax/8h/16days/train
    """
    req = clean_text(requirement).lower()

    pax = np.nan
    hours = np.nan
    days = np.nan

    pax_match = re.search(r"(\d+(?:\.\d+)?)\s*pax", req)
    if pax_match:
        pax = float(pax_match.group(1))

    hours_match = re.search(r"(\d+(?:\.\d+)?)\s*h", req)
    if hours_match:
        hours = float(hours_match.group(1))

    days_match = re.search(r"(\d+(?:\.\d+)?)\s*days?", req)
    if days_match:
        days = float(days_match.group(1))

    return pax, hours, days


def is_assigned_name(name: str) -> bool:
    value = clean_text(name).upper()

    if value in UNASSIGNED_NAMES:
        return False

    if value.startswith("TBC"):
        return False

    return True


def classify_role(role: str) -> str:
    text = clean_text(role).lower()

    if text == "":
        return "Uncategorised"

    if "subcon" in text:
        return "Subcon"

    if "backup" in text:
        return "Backup"

    if "atc/ats" in text or ("atc" in text and "ats" in text):
        return "ATC/ATS"

    if "ats" in text:
        return "ATS"

    if "atc" in text:
        return "ATC"

    if "sig" in text:
        return "SIG"

    if "comms" in text or "comm" in text:
        return "Comms"

    if "train" in text:
        return "Train"

    if "csf" in text:
        return "CSF"

    if "coordinator" in text:
        return "Coordinator"

    if "manager" in text:
        return "Manager"

    if "design" in text:
        return "Design"

    if "rcs" in text or "network" in text:
        return "RCS/Network"

    if "engineer" in text:
        return "Engineer"

    return clean_text(role).split()[0]


def read_uploaded_table(uploaded_file) -> List[List[str]]:
    filename = uploaded_file.name.lower()

    if filename.endswith((".xlsx", ".xls")):
        raw_df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna("")
        rows = raw_df.astype(str).values.tolist()
    else:
        raw = uploaded_file.getvalue().decode("utf-8-sig", errors="replace")

        # Your pasted table is usually tab-separated.
        delimiter = "\t" if raw.count("\t") >= raw.count(",") else ","
        rows = list(csv.reader(io.StringIO(raw), delimiter=delimiter))

    max_cols = max((len(r) for r in rows), default=0)
    fixed_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    return fixed_rows


def find_header_row(rows: List[List[str]]) -> int:
    """
    Finds the row with Requirement / Pax / Hours/day / Days / JAN FEB MAR...
    """
    best_row = 0
    best_score = -1

    for idx, row in enumerate(rows[:30]):
        lower_values = [clean_text(x).lower() for x in row]
        month_count = sum(1 for x in row if parse_month(x))

        score = month_count

        if "requirement" in lower_values:
            score += 20

        if "pax" in lower_values:
            score += 5

        if "hours/day" in lower_values or "hours" in lower_values:
            score += 5

        if "days" in lower_values:
            score += 5

        if score > best_score:
            best_score = score
            best_row = idx

    if best_score < 5:
        raise ValueError(
            "Could not detect the header row. Please upload the table with the year row and month row intact."
        )

    return best_row


def detect_columns(rows: List[List[str]], header_idx: int):
    header = rows[header_idx]
    header_lower = [clean_text(x).lower() for x in header]

    def find_col(possible_names, default=None):
        for name in possible_names:
            if name in header_lower:
                return header_lower.index(name)
        return default

    # Based on your format:
    # Col A = SN / Project
    # Col B = Activity or Role
    # Col C = Name
    # Col D = Requirement
    sn_col = 0
    role_col = 1
    name_col = 2

    req_col = find_col(["requirement"], 3)
    pax_col = find_col(["pax"], None)
    hours_col = find_col(["hours/day", "hours per day", "hours"], None)
    days_col = find_col(["days", "day"], None)

    month_cols = [idx for idx, value in enumerate(header) if parse_month(value)]

    if not month_cols:
        raise ValueError("No month columns were detected. Make sure the month row contains JAN, FEB, MAR, etc.")

    year_row = rows[header_idx - 1] if header_idx > 0 else [""] * len(header)

    current_year = None
    month_map = []

    for idx in range(len(header)):
        detected_year = parse_year(year_row[idx]) if idx < len(year_row) else None

        if detected_year:
            current_year = detected_year

        if idx in month_cols:
            month_num = parse_month(header[idx])
            year = current_year

            if year is None:
                label = f"Unknown-{clean_text(header[idx])}-{idx}"
                month_dt = pd.NaT
            else:
                label = pd.Timestamp(year, month_num, 1).strftime("%Y-%b")
                month_dt = pd.Timestamp(year, month_num, 1)

            month_map.append(
                {
                    "col": idx,
                    "label": label,
                    "date": month_dt,
                }
            )

    return sn_col, role_col, name_col, req_col, pax_col, hours_col, days_col, month_map


def is_project_header(sn: str, role: str, name: str, requirement: str, has_load: bool) -> bool:
    sn_clean = clean_text(sn).upper()

    return (
        sn_clean in PROJECT_HEADERS
        and clean_text(requirement) == ""
        and not has_load
    )


def is_summary_start(sn: str, role: str) -> bool:
    return clean_text(sn).lower() == "total" or clean_text(role).lower() == "total"


def parse_resource_table(rows: List[List[str]]):
    header_idx = find_header_row(rows)

    (
        sn_col,
        role_col,
        name_col,
        req_col,
        pax_col,
        hours_col,
        days_col,
        month_map,
    ) = detect_columns(rows, header_idx)

    current_project = ""
    current_activity = ""
    current_activity_no = ""
    current_requirement = ""
    current_pax = np.nan
    current_hours = np.nan
    current_days = np.nan

    summary_started = False

    detail_records = []
    long_records = []

    data_rows = rows[header_idx + 1:]

    for source_row_no, row in enumerate(data_rows, start=header_idx + 2):
        sn = clean_text(row[sn_col]) if sn_col < len(row) else ""
        role_or_activity = clean_text(row[role_col]) if role_col < len(row) else ""
        name = clean_text(row[name_col]) if name_col < len(row) else ""
        requirement = clean_text(row[req_col]) if req_col < len(row) else ""

        month_values = {
            m["label"]: to_number(row[m["col"]])
            for m in month_map
            if m["col"] < len(row)
        }

        has_load = any(abs(v) > 1e-12 for v in month_values.values())

        # Stop before existing manual summary section.
        if is_summary_start(sn, role_or_activity):
            summary_started = True
            continue

        if summary_started:
            continue

        # Skip completely blank rows.
        if not any([sn, role_or_activity, name, requirement, has_load]):
            continue

        # Project rows like JRL / CRL / RTS.
        if is_project_header(sn, role_or_activity, name, requirement, has_load):
            current_project = clean_text(sn).upper()
            current_activity = ""
            current_activity_no = ""
            current_requirement = ""
            current_pax = np.nan
            current_hours = np.nan
            current_days = np.nan
            continue

        # Skip display header row like: JRL | Name
        if clean_text(name).lower() == "name" and not requirement and not has_load:
            continue

        req_pax, req_hours, req_days = parse_requirement(requirement)

        pax_value = to_optional_number(row[pax_col]) if pax_col is not None and pax_col < len(row) else req_pax
        hours_value = to_optional_number(row[hours_col]) if hours_col is not None and hours_col < len(row) else req_hours
        days_value = to_optional_number(row[days_col]) if days_col is not None and days_col < len(row) else req_days

        if np.isnan(pax_value):
            pax_value = req_pax

        if np.isnan(hours_value):
            hours_value = req_hours

        if np.isnan(days_value):
            days_value = req_days

        # Important:
        # Rows with requirement in Column D and no name = activity/subtotal rows.
        # The person rows below inherit this activity requirement.
        # But rows with requirement and a name, e.g. T&C Manager | Eric Tan | 1 pax/8h/5 days,
        # are treated as actual load rows.
        is_activity_row = (
            (bool(requirement) and not name)
            or (role_or_activity and not name and not has_load)
        )

        if is_activity_row:
            current_activity = role_or_activity if role_or_activity else current_activity
            current_activity_no = sn if sn else current_activity_no
            current_requirement = requirement
            current_pax = pax_value
            current_hours = hours_value
            current_days = days_value
            continue

        # Resource row.
        if not role_or_activity and not has_load:
            continue

        assigned = is_assigned_name(name)
        display_name = name if name else "Unassigned"
        role_group = classify_role(role_or_activity)

        row_activity = current_activity if current_activity else role_or_activity
        row_requirement = requirement if requirement else current_requirement
        row_pax = pax_value if requirement else current_pax
        row_hours = hours_value if requirement else current_hours
        row_days = days_value if requirement else current_days

        base_record = {
            "Source Row": source_row_no,
            "Project": current_project,
            "Activity No": current_activity_no,
            "Activity": row_activity,
            "Role": role_or_activity,
            "Role Group": role_group,
            "Name": display_name,
            "Assigned Name?": assigned,
            "Activity Requirement": row_requirement,
            "Activity Pax": row_pax,
            "Activity Hours/day": row_hours,
            "Activity Days": row_days,
        }

        detail_records.append({**base_record, **month_values})

        for m in month_map:
            load_value = month_values.get(m["label"], 0.0)

            if abs(load_value) > 1e-12:
                long_records.append(
                    {
                        **base_record,
                        "Month": m["label"],
                        "Month Date": m["date"],
                        "Load": load_value,
                    }
                )

    detail_df = pd.DataFrame(detail_records)
    long_df = pd.DataFrame(long_records)
    months_df = pd.DataFrame(month_map).rename(columns={"label": "Month", "date": "Month Date"})

    return detail_df, long_df, months_df


def make_person_load(long_df: pd.DataFrame, capacity: float) -> pd.DataFrame:
    if long_df.empty:
        return pd.DataFrame()

    assigned = long_df[long_df["Assigned Name?"]].copy()

    if assigned.empty:
        return pd.DataFrame()

    person_load = (
        assigned.groupby(["Name", "Month", "Month Date"], as_index=False)
        .agg(
            Load=("Load", "sum"),
            Projects=("Project", lambda x: ", ".join(sorted(set(clean_text(v) for v in x if clean_text(v))))),
            Roles=("Role", lambda x: ", ".join(sorted(set(clean_text(v) for v in x if clean_text(v))))),
            Activities=("Activity", lambda x: " | ".join(sorted(set(clean_text(v) for v in x if clean_text(v)))[:5])),
        )
        .sort_values(["Month Date", "Load"], ascending=[True, False])
    )

    person_load["Overload"] = (person_load["Load"] - capacity).clip(lower=0)
    person_load["Extra Pax to Cover Overload"] = np.ceil(person_load["Overload"]).astype(int)
    person_load["Status"] = np.where(person_load["Load"] > capacity, "Overloaded", "OK")

    return person_load


def make_role_demand(long_df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if long_df.empty:
        return pd.DataFrame()

    return (
        long_df.groupby([group_col, "Month", "Month Date"], as_index=False)
        .agg(Load=("Load", "sum"))
        .sort_values([group_col, "Month Date"])
    )


def initial_headcount(long_df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if long_df.empty:
        return pd.DataFrame(columns=[group_col, "Available Pax"])

    groups = sorted(long_df[group_col].dropna().astype(str).unique())

    assigned = long_df[long_df["Assigned Name?"]].copy()

    if assigned.empty:
        counts = pd.Series(dtype=int)
    else:
        counts = assigned.drop_duplicates([group_col, "Name"]).groupby(group_col)["Name"].nunique()

    output = pd.DataFrame({group_col: groups})
    output["Available Pax"] = output[group_col].map(counts).fillna(0).astype(int)

    return output


def make_role_gap(role_demand: pd.DataFrame, headcount_df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if role_demand.empty:
        return pd.DataFrame()

    headcount = headcount_df.copy()
    headcount["Available Pax"] = pd.to_numeric(headcount["Available Pax"], errors="coerce").fillna(0)

    gap = role_demand.merge(headcount, on=group_col, how="left")
    gap["Available Pax"] = gap["Available Pax"].fillna(0)

    gap["Shortage Load"] = (gap["Load"] - gap["Available Pax"]).clip(lower=0)
    gap["Extra Pax Needed"] = np.ceil(gap["Shortage Load"]).astype(int)
    gap["Status"] = np.where(gap["Extra Pax Needed"] > 0, "Need Hire/Reassign", "OK")

    return gap.sort_values(["Month Date", group_col])


def pivot_load(df: pd.DataFrame, index_col: str, value_col: str = "Load") -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    pivot = df.pivot_table(
        index=index_col,
        columns="Month",
        values=value_col,
        aggfunc="sum",
        fill_value=0,
    )

    month_order = df.sort_values("Month Date")["Month"].drop_duplicates().tolist()
    pivot = pivot.reindex(columns=month_order)

    return pivot.reset_index()


def style_overload_table(df: pd.DataFrame, capacity: float):
    month_cols = [c for c in df.columns if re.match(r"^\d{4}-[A-Za-z]{3}$", str(c))]

    def highlight(value):
        try:
            return "background-color: #ffd6d6; font-weight: bold" if float(value) > capacity else ""
        except Exception:
            return ""

    return df.style.applymap(highlight, subset=month_cols).format(precision=2)


def create_excel_export(
    detail_df: pd.DataFrame,
    person_load: pd.DataFrame,
    role_gap: pd.DataFrame,
    monthly_summary: pd.DataFrame,
    hiring_summary: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Cleaned_Input")
        person_load.to_excel(writer, index=False, sheet_name="Person_Load")
        person_load[person_load["Overload"] > 0].to_excel(writer, index=False, sheet_name="Person_Overload")
        role_gap.to_excel(writer, index=False, sheet_name="Role_Hiring_Gap")
        hiring_summary.to_excel(writer, index=False, sheet_name="Hiring_Summary")
        monthly_summary.to_excel(writer, index=False, sheet_name="Monthly_Summary")

        workbook = writer.book

        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1,
            }
        )

        overload_format = workbook.add_format(
            {
                "bg_color": "#FFD6D6",
            }
        )

        number_format = workbook.add_format(
            {
                "num_format": "0.00",
            }
        )

        sheets = {
            "Cleaned_Input": detail_df,
            "Person_Load": person_load,
            "Person_Overload": person_load[person_load["Overload"] > 0],
            "Role_Hiring_Gap": role_gap,
            "Hiring_Summary": hiring_summary,
            "Monthly_Summary": monthly_summary,
        }

        for sheet_name, df in sheets.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)

            for col_num, value in enumerate(df.columns):
                worksheet.write(0, col_num, value, header_format)
                width = min(max(len(str(value)) + 2, 10), 35)
                worksheet.set_column(col_num, col_num, width)

            for col_num, col_name in enumerate(df.columns):
                if (
                    col_name in {"Load", "Overload", "Shortage Load", "Available Pax"}
                    or re.match(r"^\d{4}-[A-Za-z]{3}$", str(col_name))
                ):
                    worksheet.set_column(col_num, col_num, 12, number_format)

            if sheet_name in {"Person_Load", "Person_Overload"} and "Overload" in df.columns:
                over_col = df.columns.get_loc("Overload")
                worksheet.conditional_format(
                    1,
                    over_col,
                    max(len(df), 1),
                    over_col,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": 0,
                        "format": overload_format,
                    },
                )

            if sheet_name == "Role_Hiring_Gap" and "Extra Pax Needed" in df.columns:
                extra_col = df.columns.get_loc("Extra Pax Needed")
                worksheet.conditional_format(
                    1,
                    extra_col,
                    max(len(df), 1),
                    extra_col,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": 0,
                        "format": overload_format,
                    },
                )

    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("Resource Planning Load & Hiring Tool")

st.caption(
    "Upload your resource planning table. "
    "The app treats 1.0 load as one full person-month by default."
)

with st.sidebar:
    st.header("Settings")

    capacity = st.number_input(
        "Monthly capacity per person",
        min_value=0.1,
        max_value=5.0,
        value=1.0,
        step=0.1,
    )

    group_col = st.radio(
        "Hiring gap grouping",
        ["Role Group", "Role"],
        index=0,
    )

    show_zero_rows = st.checkbox(
        "Show zero-load resource rows",
        value=False,
    )

uploaded_file = st.file_uploader(
    "Upload resource planning table",
    type=["txt", "csv", "xlsx", "xls"],
)

if uploaded_file is None:
    st.info(
        "Upload the pasted table as .txt/.csv or Excel. "
        "The format should follow your current table with Year row, Month row, "
        "Requirement, Pax, Hours/day, Days, and monthly load columns."
    )
    st.stop()

try:
    rows = read_uploaded_table(uploaded_file)
    detail_df, long_df, months_df = parse_resource_table(rows)
except Exception as exc:
    st.error(f"Could not parse the uploaded table: {exc}")
    st.stop()

if detail_df.empty or long_df.empty:
    st.warning(
        "The table was read, but no resource-load rows were detected. "
        "Check that the detail rows are below the activity requirement rows "
        "and that monthly cells contain numeric loads."
    )
    st.stop()

if not show_zero_rows:
    month_cols = months_df["Month"].tolist()
    available_month_cols = [c for c in month_cols if c in detail_df.columns]
    detail_display = detail_df[detail_df[available_month_cols].sum(axis=1) > 0].copy()
else:
    detail_display = detail_df.copy()

person_load = make_person_load(long_df, capacity)
role_demand = make_role_demand(long_df, group_col)
base_headcount = initial_headcount(long_df, group_col)

signature_source = f"{uploaded_file.name}-{len(detail_df)}-{len(long_df)}-{group_col}"
signature = hashlib.md5(signature_source.encode("utf-8")).hexdigest()[:10]
headcount_key = f"headcount_{signature}"

if headcount_key not in st.session_state:
    st.session_state[headcount_key] = base_headcount

st.sidebar.markdown("---")
st.sidebar.subheader("Available Headcount")
st.sidebar.caption("Edit this to test different hiring scenarios.")

edited_headcount = st.sidebar.data_editor(
    st.session_state[headcount_key],
    num_rows="dynamic",
    use_container_width=True,
    key=f"editor_{headcount_key}",
)

role_gap = make_role_gap(role_demand, edited_headcount, group_col)
hiring_summary = role_gap[role_gap["Extra Pax Needed"] > 0].copy()

monthly_total = (
    long_df.groupby(["Month", "Month Date"], as_index=False)
    .agg(Total_Load=("Load", "sum"))
)

monthly_extra = (
    role_gap.groupby(["Month", "Month Date"], as_index=False)
    .agg(Total_Extra_Pax_Needed=("Extra Pax Needed", "sum"))
)

monthly_summary = (
    monthly_total.merge(monthly_extra, on=["Month", "Month Date"], how="outer")
    .fillna(0)
    .sort_values("Month Date")
)

person_overload = (
    person_load[person_load["Overload"] > 0].copy()
    if not person_load.empty
    else pd.DataFrame()
)

peak_total_load = monthly_summary["Total_Load"].max() if not monthly_summary.empty else 0
peak_month = (
    monthly_summary.loc[monthly_summary["Total_Load"].idxmax(), "Month"]
    if not monthly_summary.empty
    else "-"
)

max_person_load = person_load["Load"].max() if not person_load.empty else 0
worst_person = (
    person_load.loc[person_load["Load"].idxmax(), "Name"]
    if not person_load.empty
    else "-"
)

peak_extra = (
    monthly_summary["Total_Extra_Pax_Needed"].max()
    if not monthly_summary.empty
    else 0
)

peak_extra_month = (
    monthly_summary.loc[monthly_summary["Total_Extra_Pax_Needed"].idxmax(), "Month"]
    if not monthly_summary.empty
    else "-"
)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

kpi1.metric("Peak total load", f"{peak_total_load:.1f}", peak_month)
kpi2.metric("Worst person load", f"{max_person_load:.1f}", worst_person)
kpi3.metric("Overloaded person-months", f"{len(person_overload):,}")
kpi4.metric("Peak extra pax needed", f"{peak_extra:.0f}", peak_extra_month)

tab_dashboard, tab_clean, tab_people, tab_hiring, tab_export = st.tabs(
    [
        "Dashboard",
        "Cleaned Rows",
        "Person Overload",
        "Hiring Gap",
        "Export",
    ]
)

with tab_dashboard:
    st.subheader("Monthly Demand")

    chart_df = monthly_summary.copy()

    if len(chart_df) > 0:
        chart_df = chart_df.sort_values("Month Date")
        chart_df["Month Text"] = chart_df["Month"].astype(str)

        st.write("Total monthly load")
        total_chart = chart_df.set_index("Month Text")[["Total_Load"]]
        st.line_chart(total_chart, use_container_width=True)

        st.write("Total extra pax needed by month")
        extra_chart = chart_df.set_index("Month Text")[["Total_Extra_Pax_Needed"]]
        st.bar_chart(extra_chart, use_container_width=True)

    st.subheader("Immediate Hiring/Reassignment Alerts")

    if hiring_summary.empty:
        st.success("No hiring gap based on the current available headcount settings.")
    else:
        st.dataframe(
            hiring_summary[
                [
                    group_col,
                    "Month",
                    "Load",
                    "Available Pax",
                    "Shortage Load",
                    "Extra Pax Needed",
                    "Status",
                ]
            ],
            use_container_width=True,
        )

with tab_clean:
    st.subheader("Cleaned Resource Rows")

    st.caption(
        "Rows with Column D requirement and no name are treated as activity/subtotal rows. "
        "Rows below them are treated as resource rows and inherit the activity requirement."
    )

    st.dataframe(
        detail_display,
        use_container_width=True,
        height=600,
    )

with tab_people:
    st.subheader("Person Load")

    if person_load.empty:
        st.warning(
            "No assigned names were found. "
            "TBC/blank rows are counted for role demand, but not for individual overload."
        )
    else:
        person_pivot = pivot_load(person_load, "Name")

        st.write("Cells above capacity are highlighted.")
        st.dataframe(
            style_overload_table(person_pivot, capacity),
            use_container_width=True,
            height=500,
        )

        st.subheader("Overload Details")

        if person_overload.empty:
            st.success("No named person exceeds the monthly capacity.")
        else:
            st.dataframe(
                person_overload,
                use_container_width=True,
                height=500,
            )

with tab_hiring:
    st.subheader(f"Hiring Gap by {group_col}")

    st.caption(
        "Demand is total monthly load. "
        "Available pax comes from the editable headcount table in the sidebar. "
        "Extra pax = CEILING(MAX(Demand - Available, 0))."
    )

    role_pivot = pivot_load(role_gap, group_col, value_col="Extra Pax Needed")

    st.write("Extra pax needed by month")
    st.dataframe(
        role_pivot,
        use_container_width=True,
        height=500,
    )

    st.write("Detailed gap calculation")
    st.dataframe(
        role_gap,
        use_container_width=True,
        height=500,
    )

with tab_export:
    st.subheader("Download Results")

    export_bytes = create_excel_export(
        detail_df,
        person_load,
        role_gap,
        monthly_summary,
        hiring_summary,
    )

    st.download_button(
        label="Download analysis as Excel",
        data=export_bytes,
        file_name="resource_planning_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("### Output sheets")
    st.write(
        "Cleaned_Input, Person_Load, Person_Overload, "
        "Role_Hiring_Gap, Hiring_Summary, Monthly_Summary"
    )
