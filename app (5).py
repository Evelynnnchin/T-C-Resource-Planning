import streamlit as st
import pandas as pd
from streamlit_echarts import st_echarts
import re
import io
from datetime import datetime
from zoneinfo import ZoneInfo
import textwrap

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
    REPORTLAB_AVAILABLE = True
except ModuleNotFoundError:
    REPORTLAB_AVAILABLE = False


# =========================================================
# 1. Page Configuration
# =========================================================
st.set_page_config(page_title="Editable T&C Org Chart", layout="wide")
st.title("🏢 Editable T&C Organizational Chart")


# =========================================================
# 2. Expected Upload Format
# =========================================================
REQUIRED_COLS = [
    "Name / Team Name",
    "Type",
    "Job Title",
    "Reports To",
    "Color Group",
]

OPTIONAL_COLS = ["Time Period"]
ALL_COLS = REQUIRED_COLS + OPTIONAL_COLS


def normalise_header(value):
    value = str(value).strip().lower()
    value = value.replace("colour", "color")
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


HEADER_ALIASES = {
    "nameteamname": "Name / Team Name",
    "name": "Name / Team Name",
    "teamname": "Name / Team Name",
    "nameteam": "Name / Team Name",

    "type": "Type",

    "jobtitle": "Job Title",
    "role": "Job Title",
    "position": "Job Title",
    "title": "Job Title",

    "reportsto": "Reports To",
    "supervisor": "Reports To",
    "manager": "Reports To",
    "parent": "Reports To",

    "colorgroup": "Color Group",
    "colourgroup": "Color Group",
    "group": "Color Group",
    "project": "Color Group",
    "department": "Color Group",

    "timeperiod": "Time Period",
    "period": "Time Period",
    "duration": "Time Period",
}


def standardise_uploaded_table(df):
    """
    Converts uploaded Excel/CSV into:
    Name / Team Name | Type | Job Title | Reports To | Color Group | Time Period
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}
    for col in df.columns:
        key = normalise_header(col)
        if key in HEADER_ALIASES:
            rename_map[col] = HEADER_ALIASES[key]

    df = df.rename(columns=rename_map)

    if "Name / Team Name" not in df.columns:
        raise ValueError(
            "Missing required column: Name / Team Name. "
            "Your Excel must have a name/team name column."
        )

    if "Type" not in df.columns:
        df["Type"] = "Person"

    if "Job Title" not in df.columns:
        df["Job Title"] = ""

    if "Reports To" not in df.columns:
        df["Reports To"] = "None"

    if "Color Group" not in df.columns:
        df["Color Group"] = "None"

    if "Time Period" not in df.columns:
        df["Time Period"] = ""

    df = df[ALL_COLS].copy()

    for col in ALL_COLS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Name / Team Name"] = df["Name / Team Name"].replace(
        {"nan": "", "NaN": "", "None": ""}
    )
    df = df[df["Name / Team Name"] != ""].copy()

    df["Type"] = df["Type"].replace({"": "Person"})
    df["Type"] = df["Type"].apply(
        lambda x: "Team Box"
        if str(x).strip().lower() in [
            "team",
            "team box",
            "teambox",
            "group",
            "project group",
        ]
        else "Person"
    )

    df["Reports To"] = df["Reports To"].replace(
        {"": "None", "nan": "None", "NaN": "None"}
    )

    df["Color Group"] = df["Color Group"].replace(
        {"": "None", "nan": "None", "NaN": "None"}
    )

    return df


def read_uploaded_org_file(uploaded_file, sheet_name=None):
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(file_bytes))
        return standardise_uploaded_table(raw)

    if file_name.endswith((".xlsx", ".xlsm", ".xls")):
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name or 0)
        return standardise_uploaded_table(raw)

    raise ValueError("Please upload an Excel file or CSV file.")


def get_excel_sheet_names(uploaded_file):
    file_name = uploaded_file.name.lower()

    if not file_name.endswith((".xlsx", ".xlsm", ".xls")):
        return []

    xls = pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()))
    return xls.sheet_names


# =========================================================
# 3. Default Data
# =========================================================
def get_default_data():
    return pd.DataFrame({
        "Name / Team Name": [
            "Ramli, Mohammed Helmi",
            "T&C Coordinator",

            "DTL",
            "Chin, Raymond",
            "Araguete Gamero, Eduardo Andres",
            "Uthaiyasuriyan, Mohan",
            "Bin Abdul Shukor, Ahmad Syafiq",
            "Tan, Sam Teng Boon",
            "Cher, Yee Hern Malcolm",
            "BIN AHMAD, MUHAMAD ZULKHAIRI",
            "Mohamed Haleem, Mohamed Irfan",
            "Bin Mawasi, Muhammad Khairi",
            "Almonte, Rhyle Manuel",

            "JRL",
            "Tan, Zhong Han",
            "Bin Powzan, Muhammad Faridzuan",
            "Sidik, Diyana",
            "Muhammad Zaki Bin Ismail",
            "Muhammad Sufian Bin Moksin",
            "Tan Yih Chyuan",
            "Akmal",

            "CRL",
            "Udeaja, Chukwudi Augustine",

            "RTS",
            "Khew, Aceline",
            "Jack",
            "Vincent",
            "Binte Samsudin, Khairunnisa",
            "Nazmi",
            "Ku, Teerapat Kian Xiong",

            "Train",
            "Fabro, Richter",
            "Bin Sabari, Irwan",

            "CSF",
            "Eldho, Basil",
        ],
        "Type": [
            "Person",
            "Team Box",

            "Team Box",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",

            "Team Box",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",

            "Team Box",
            "Person",

            "Team Box",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",
            "Person",

            "Team Box",
            "Person",
            "Person",

            "Team Box",
            "Person",
        ],
        "Job Title": [
            "T&C Manager",
            "T&C Coordinator",

            "Project Group",
            "Senior System Design Engineer",
            "ATS Design Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",
            "T&C Engineer",

            "Project Group",
            "ATS System Design Engineer",
            "Signalling T&C Engineer",
            "T&C Coordinator",
            "T&C Engineer",
            "T&C Engineer",
            "Role TBC",
            "T&C Engineer",

            "Project Group",
            "OSIT Manager",

            "Project Group",
            "Signalling T&C Engineer",
            "Role TBC",
            "Role TBC",
            "Senior RAMS Engineer",
            "Role TBC",
            "T&C Coordinator",

            "Project Group",
            "Trainborne T&C Engineer",
            "System Engineer",

            "Project Group",
            "Signalling T&C Engineer",
        ],
        "Reports To": [
            "None",
            "Ramli, Mohammed Helmi",

            "T&C Coordinator",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",

            "T&C Coordinator",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",

            "T&C Coordinator",
            "CRL",

            "T&C Coordinator",
            "RTS",
            "RTS",
            "RTS",
            "RTS",
            "RTS",
            "RTS",

            "T&C Coordinator",
            "Train",
            "Train",

            "T&C Coordinator",
            "CSF",
        ],
        "Color Group": [
            "Management",
            "Management",

            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",
            "DTL",

            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",
            "JRL",

            "CRL",
            "CRL",

            "RTS",
            "RTS",
            "RTS",
            "RTS",
            "RTS",
            "RTS",
            "RTS",

            "Train",
            "Train",
            "Train",

            "CSF",
            "CSF",
        ],
        "Time Period": [""] * 35,
    })


if "org_data" not in st.session_state:
    st.session_state.org_data = get_default_data()

if "loaded_upload_signature" not in st.session_state:
    st.session_state.loaded_upload_signature = None


# =========================================================
# 4. Helper Functions
# =========================================================
def prepare_clean_df(df):
    clean = df.copy()

    for col in ALL_COLS:
        if col not in clean.columns:
            clean[col] = ""

    clean = clean.dropna(subset=["Name / Team Name"]).copy()
    clean["Name / Team Name"] = clean["Name / Team Name"].astype(str).str.strip()
    clean = clean[clean["Name / Team Name"] != ""].copy()

    clean["Reports To"] = clean["Reports To"].fillna("None").astype(str).str.strip()
    clean["Reports To"] = clean["Reports To"].replace(
        {"": "None", "nan": "None", "NaN": "None"}
    )

    clean["Job Title"] = clean["Job Title"].fillna("").astype(str).str.strip()

    clean["Color Group"] = clean["Color Group"].fillna("None").astype(str).str.strip()
    clean["Color Group"] = clean["Color Group"].replace(
        {"": "None", "nan": "None", "NaN": "None"}
    )

    clean["Type"] = clean["Type"].fillna("Person").astype(str).str.strip()
    clean["Type"] = clean["Type"].replace({"": "Person"})

    clean["Time Period"] = clean["Time Period"].fillna("").astype(str).str.strip()

    clean["Role Group"] = clean["Job Title"].apply(
        lambda x: re.sub(r"\s*\d+$", "", str(x)).strip()
    )

    return clean


def detect_loops(df):
    names = set(df["Name / Team Name"].astype(str).str.strip())
    reports_to = {}

    for _, row in df.iterrows():
        name = str(row["Name / Team Name"]).strip()
        manager = str(row["Reports To"]).strip()

        if name and manager in names and manager != "None":
            reports_to[name] = manager

    problem_names = []

    for name in names:
        seen = set()
        current = name

        while current in reports_to:
            if current in seen:
                problem_names.append(name)
                break

            seen.add(current)
            current = reports_to[current]

    return sorted(set(problem_names))


def find_missing_managers(df):
    names = set(df["Name / Team Name"].astype(str).str.strip())
    missing = []

    for manager in df["Reports To"].astype(str).str.strip().unique():
        if manager and manager.lower() != "none" and manager not in names:
            missing.append(manager)

    return sorted(set(missing))


def dataframe_to_excel_bytes(df):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Org Chart Import", index=False)

    buffer.seek(0)
    return buffer.getvalue()


def hex_to_reportlab_colour(hex_code, fallback="#ced4da"):
    if not REPORTLAB_AVAILABLE:
        return None

    value = str(hex_code).strip()

    if not value.startswith("#"):
        value = fallback

    try:
        return colors.HexColor(value)
    except Exception:
        return colors.HexColor(fallback)


# =========================================================
# 5. Upload Excel / CSV
# =========================================================
st.markdown("### 📤 Upload Excel File")
st.write(
    "Upload an Excel or CSV file with these columns: "
    "`Name / Team Name`, `Type`, `Job Title`, `Reports To`, `Color Group`. "
    "`Time Period` is optional."
)

uploaded_file = st.file_uploader(
    "Upload org chart Excel or CSV",
    type=["xlsx", "xlsm", "xls", "csv"],
    key="org_chart_upload",
)

selected_sheet = None

if uploaded_file is not None:
    try:
        sheet_names = get_excel_sheet_names(uploaded_file)

        if sheet_names:
            selected_sheet = st.selectbox(
                "Select Excel sheet",
                sheet_names,
                index=0,
                key="selected_excel_sheet",
            )

        upload_signature = (
            uploaded_file.name,
            uploaded_file.size,
            selected_sheet or "csv",
        )

        if st.session_state.loaded_upload_signature != upload_signature:
            loaded_df = read_uploaded_org_file(uploaded_file, selected_sheet)
            st.session_state.org_data = loaded_df
            st.session_state.loaded_upload_signature = upload_signature
            st.success("Uploaded file loaded. The org chart will generate from this table.")

        if st.button("Reload uploaded file", use_container_width=True):
            loaded_df = read_uploaded_org_file(uploaded_file, selected_sheet)
            st.session_state.org_data = loaded_df
            st.session_state.loaded_upload_signature = upload_signature
            st.success("Reloaded uploaded file.")

        st.caption("Tip: row order controls the left-to-right project order and vertical people order.")

    except Exception as e:
        st.error(f"Could not load uploaded file: {e}")


# =========================================================
# 6. Sidebar Options
# =========================================================
clean_df = prepare_clean_df(st.session_state.org_data)

st.sidebar.header("🔍 View Options")

filter_type = st.sidebar.radio(
    "Select a view:",
    [
        "Show All (No Filters)",
        "Highlight by Name",
        "Highlight by Role Group",
        "Highlight by Color Group",
    ],
)

selected_person = "All"
selected_role_group = "All"
selected_dept = "All"

if filter_type == "Highlight by Name":
    selected_person = st.sidebar.selectbox(
        "Select Name:",
        sorted(clean_df["Name / Team Name"].unique().tolist()),
    )

elif filter_type == "Highlight by Role Group":
    valid_roles = sorted([r for r in clean_df["Role Group"].unique().tolist() if r != ""])
    selected_role_group = st.sidebar.selectbox("Select Role Group:", valid_roles)
    role_count = len(clean_df[clean_df["Role Group"] == selected_role_group])
    st.sidebar.success(f"👥 Total {selected_role_group} count: **{role_count}**")

elif filter_type == "Highlight by Color Group":
    selected_dept = st.sidebar.selectbox(
        "Select Color Group:",
        sorted(clean_df["Color Group"].unique().tolist()),
    )

filter_active = filter_type != "Show All (No Filters)"


# =========================================================
# 7. Sidebar Colors
# =========================================================
st.sidebar.header("🎨 Customise Team Colors")

default_palette = {
    "Management": "#0081a7",
    "Group": "#00afb9",
    "Project Based": "#00afb9",
    "Shared Pool": "#00afb9",

    "CRL": "#00afb9",
    "JRL": "#00afb9",
    "RTS": "#00afb9",
    "DTL": "#00afb9",
    "Train": "#00afb9",
    "CSF": "#00afb9",

    "None": "#ced4da",

    "CRL / OSIT": "#00afb9",
    "JRL Mainline": "#00afb9",
    "ATC": "#00afb9",
    "ATC / ATS": "#00afb9",
    "ATS": "#00afb9",
    "Comms / DCS / RCS / Network": "#00afb9",
    "Signalling": "#00afb9",
    "Subcon": "#ffb703",
}

color_map = {}

with st.sidebar.expander("Click to change colors"):
    for dept in sorted(clean_df["Color Group"].unique()):
        default_c = default_palette.get(dept, "#ced4da")
        color_map[dept] = st.color_picker(str(dept), default_c)


# =========================================================
# 8. Sidebar Layout Settings
# =========================================================
st.sidebar.header("📐 Chart Settings")

with st.sidebar.expander("Layout Settings"):
    chart_width = st.slider("Chart Width", 1000, 8000, 2200, 100)
    chart_height = st.slider("Chart Height", 800, 8000, 3000, 100)

with st.sidebar.expander("Font and Box Settings", expanded=True):
    name_font_size = st.slider("Name Font Size", 8, 30, 12, 1)
    role_font_size = st.slider("Role / Job Title Font Size", 6, 24, 10, 1)
    time_font_size = st.slider("Time Period Font Size", 6, 22, 9, 1)

    node_width = st.slider("Box Width", 120, 400, 180, 10)
    node_height = st.slider("Box Height", 50, 200, 75, 10)

    horizontal_gap = st.slider("PDF Horizontal Gap", 30, 220, 70, 5)
    vertical_gap = st.slider("PDF Vertical Gap", 40, 240, 90, 5)


# =========================================================
# 9. Data Editor
# =========================================================
st.markdown("### ✏️ Edit Data Directly")
st.write("You can still edit after upload. The chart updates from this table.")

clean_df = prepare_clean_df(st.session_state.org_data)

all_possible_managers = clean_df["Name / Team Name"].tolist() + ["None"]

dynamic_color_groups = sorted(
    set(default_palette.keys()).union(set(clean_df["Color Group"].astype(str).tolist()))
)

edited_df = st.data_editor(
    st.session_state.org_data,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    height=380,
    column_config={
        "Name / Team Name": st.column_config.TextColumn(
            "Name / Team Name",
            required=True,
        ),
        "Type": st.column_config.SelectboxColumn(
            "Type",
            help="Is this a real person or a structural team box?",
            options=["Person", "Team Box"],
            required=True,
        ),
        "Job Title": st.column_config.TextColumn(
            "Job Title",
        ),
        "Reports To": st.column_config.SelectboxColumn(
            "Reports To",
            help="Select the Person or Team this row sits under.",
            options=all_possible_managers,
            required=True,
        ),
        "Color Group": st.column_config.SelectboxColumn(
            "Color Group",
            help="Select which color group to apply to the box.",
            options=dynamic_color_groups,
        ),
        "Time Period": st.column_config.TextColumn(
            "Time Period",
        ),
    },
)

st.session_state.org_data = edited_df
clean_df = prepare_clean_df(st.session_state.org_data)

col_download_1, col_download_2 = st.columns(2)

with col_download_1:
    try:
        st.download_button(
            "Download Current Table as Excel",
            data=dataframe_to_excel_bytes(clean_df[ALL_COLS]),
            file_name="org_chart_import_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception:
        pass

with col_download_2:
    st.download_button(
        "Download Current Table as CSV",
        data=clean_df[ALL_COLS].to_csv(index=False).encode("utf-8-sig"),
        file_name="org_chart_import_table.csv",
        mime="text/csv",
        use_container_width=True,
    )

st.markdown("***")


# =========================================================
# 10. Build Tree Data
# =========================================================
def get_node_display_data(current_name, df):
    person_data = df[df["Name / Team Name"] == current_name]

    if person_data.empty:
        return None

    person = person_data.iloc[0]

    role = person.get("Job Title", "")
    role_group = person.get("Role Group", "")
    time_period = person.get("Time Period", "")
    dept = person.get("Color Group", "N/A")
    entry_type = person.get("Type", "Person")

    is_match = True

    if filter_active:
        if filter_type == "Highlight by Name" and current_name != selected_person:
            is_match = False
        elif filter_type == "Highlight by Role Group" and role_group != selected_role_group:
            is_match = False
        elif filter_type == "Highlight by Color Group" and dept != selected_dept:
            is_match = False

    return {
        "name": current_name,
        "role": str(role),
        "role_group": str(role_group),
        "time_period": str(time_period),
        "dept": str(dept),
        "entry_type": str(entry_type),
        "is_match": is_match,
    }


def build_tree(current_name, df, visited=None, real_supervisor=None):
    """
    Top-down chart, but bottom-level people are stacked vertically.

    Example:
    T&C Manager
       ↓
    T&C Coordinator
       ↓
    DTL     JRL     CRL
     ↓       ↓       ↓
    person  person  person
     ↓       ↓
    person  person
    """
    if visited is None:
        visited = set()

    if current_name in visited:
        return None

    visited.add(current_name)

    node_data = get_node_display_data(current_name, df)

    if node_data is None:
        return None

    role = node_data["role"]
    time_period = node_data["time_period"]
    dept = node_data["dept"]
    is_match = node_data["is_match"]

    person = df[df["Name / Team Name"] == current_name].iloc[0]
    actual_supervisor = person.get("Reports To", "None")
    display_supervisor = real_supervisor if real_supervisor else actual_supervisor

    real_direct_reports = df[df["Reports To"] == current_name]["Name / Team Name"].tolist()
    reports_str = ", ".join(real_direct_reports) if real_direct_reports else "None"

    node_color = color_map.get(dept, "#ced4da")

    if filter_active and not is_match:
        item_style = {
            "color": "#f8f9fa",
            "borderColor": "#ced4da",
            "borderWidth": 1,
        }

        display_text = f"{{name_faded|{current_name}}}"

        if str(role).strip() != "":
            display_text += f"\n{{role_faded|{role}}}"

        if time_period and str(time_period).strip() != "":
            display_text += f"\n{{time_faded|{time_period}}}"

    else:
        item_style = {
            "color": node_color,
            "borderColor": node_color,
            "borderWidth": 1,
        }

        display_text = f"{{name_active|{current_name}}}"

        if str(role).strip() != "":
            display_text += f"\n{{role_active|{role}}}"

        if time_period and str(time_period).strip() != "":
            display_text += f"\n{{time_active|{time_period}}}"

    tooltip_value = (
        f"<b>Name / Team:</b> {current_name}<br/>"
        f"<b>Role:</b> {role if str(role).strip() != '' else 'N/A'}<br/>"
        f"<b>Color Group:</b> {dept}<br/>"
        f"<b>Time Period:</b> {time_period if str(time_period).strip() != '' else 'N/A'}<br/>"
        f"<b>Reports To:</b> {display_supervisor}<br/>"
        f"<b>Direct Reports ({len(real_direct_reports)}):</b> {reports_str}"
    )

    node = {
        "name": display_text,
        "value": tooltip_value,
        "itemStyle": item_style,
        "raw": node_data,
        "children": [],
    }

    # =====================================================
    # STACKING LOGIC
    # If all direct reports are bottom-level people,
    # turn them into a vertical chain instead of a wide row.
    # =====================================================
    if len(real_direct_reports) > 0:
        all_children_are_bottom_level = True

        for report in real_direct_reports:
            child_has_reports = not df[df["Reports To"] == report].empty

            if child_has_reports:
                all_children_are_bottom_level = False
                break

        if all_children_are_bottom_level:
            current_chain_link = node

            for report in real_direct_reports:
                child_node = build_tree(
                    report,
                    df,
                    visited,
                    real_supervisor=current_name,
                )

                if child_node:
                    current_chain_link["children"].append(child_node)
                    current_chain_link = child_node

        else:
            for report in real_direct_reports:
                child_node = build_tree(report, df, visited)

                if child_node:
                    node["children"].append(child_node)

    return node


def get_root_names(df):
    top_level_matches = df[df["Reports To"].astype(str).str.lower() == "none"]

    if not top_level_matches.empty:
        return top_level_matches["Name / Team Name"].tolist()

    return [df.iloc[0]["Name / Team Name"]]


def build_chart_tree(df):
    roots = get_root_names(df)

    if len(roots) == 1:
        return build_tree(roots[0], df)

    children = []
    visited = set()

    for root in roots:
        child = build_tree(root, df, visited)

        if child:
            children.append(child)

    return {
        "name": "{name_active|Org Chart}",
        "value": "<b>Name / Team:</b> Org Chart<br/><b>Role:</b> Master Root",
        "itemStyle": {
            "color": color_map.get("Management", "#0081a7"),
            "borderColor": color_map.get("Management", "#0081a7"),
            "borderWidth": 1,
        },
        "raw": {
            "name": "Org Chart",
            "role": "Master Root",
            "role_group": "Master Root",
            "time_period": "",
            "dept": "Management",
            "entry_type": "Team Box",
            "is_match": True,
        },
        "children": children,
    }


# =========================================================
# 11. PDF Export Functions
# =========================================================
def assign_tree_positions(root):
    """
    Top-down layout.
    Because bottom-level people are chained in build_tree(),
    each project branch becomes narrow and stacked.
    """
    positions = {}
    nodes = {}
    edges = []
    next_leaf_x = [0]
    max_depth = [0]

    def walk(node, depth, path):
        node_id = path
        nodes[node_id] = node
        max_depth[0] = max(max_depth[0], depth)

        children = node.get("children", [])

        if children:
            child_x_values = []

            for idx, child in enumerate(children):
                child_id = f"{path}.{idx}"
                edges.append((node_id, child_id))
                walk(child, depth + 1, child_id)
                child_x_values.append(positions[child_id][0])

            x = (min(child_x_values) + max(child_x_values)) / 2

        else:
            x = next_leaf_x[0]
            next_leaf_x[0] += 1

        positions[node_id] = (x, depth)

    walk(root, 0, "0")

    leaf_count = max(next_leaf_x[0], 1)

    return positions, nodes, edges, leaf_count, max_depth[0]


def draw_centered_wrapped_text(c, text, x, top_y, width, font_name, font_size, max_lines):
    if not text or str(text).strip() == "":
        return

    safe_text = str(text).replace("&amp;", "&")
    approx_char_width = max(font_size * 0.52, 1)
    wrap_chars = max(int(width / approx_char_width), 8)

    lines = []

    for part in safe_text.split("\n"):
        wrapped = textwrap.wrap(part, width=wrap_chars) or [""]
        lines.extend(wrapped)

    if len(lines) > max_lines:
        lines = lines[:max_lines]

        if len(lines[-1]) > 3:
            lines[-1] = lines[-1][:-3] + "..."

    line_height = font_size + 3

    for idx, line in enumerate(lines):
        line_width = stringWidth(line, font_name, font_size)
        c.drawString(x + (width - line_width) / 2, top_y - (idx * line_height), line)


def make_full_org_chart_pdf(tree_data, chart_title="T&C Organizational Chart"):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is not installed. Add reportlab to requirements.txt.")

    positions, nodes, edges, leaf_count, max_depth = assign_tree_positions(tree_data)

    box_w = node_width
    box_h = node_height
    x_gap = horizontal_gap
    y_gap = vertical_gap

    margin_x = 45
    margin_y = 45
    title_space = 55

    chart_w = (leaf_count - 1) * (box_w + x_gap) + box_w
    chart_h = (max_depth + 1) * box_h + max_depth * y_gap

    page_w = max(chart_w + margin_x * 2, 900)
    page_h = max(chart_h + margin_y * 2 + title_space, 650)

    max_page_size = 14400
    page_w = min(page_w, max_page_size)
    page_h = min(page_h, max_page_size)

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_w, page_h))

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(margin_x, page_h - margin_y, chart_title)

    c.setFont("Helvetica", 9)
    exported_time = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%d %b %Y, %I:%M %p SGT")
    c.drawString(margin_x, page_h - margin_y - 16, f"Exported: {exported_time}")

    chart_top_y = page_h - margin_y - title_space

    def node_box_xy(node_id):
        x_unit, depth = positions[node_id]
        left = margin_x + x_unit * (box_w + x_gap)
        top = chart_top_y - depth * (box_h + y_gap)
        bottom = top - box_h
        return left, top, bottom

    c.setStrokeColor(colors.HexColor("#6c757d"))
    c.setLineWidth(1)

    for parent_id, child_id in edges:
        p_left, p_top, p_bottom = node_box_xy(parent_id)
        c_left, c_top, c_bottom = node_box_xy(child_id)

        parent_x = p_left + box_w / 2
        child_x = c_left + box_w / 2
        parent_y = p_bottom
        child_y = c_top

        mid_y = (parent_y + child_y) / 2

        c.line(parent_x, parent_y, parent_x, mid_y)
        c.line(parent_x, mid_y, child_x, mid_y)
        c.line(child_x, mid_y, child_x, child_y)

    pdf_name_size = max(name_font_size * 0.8, 6)
    pdf_role_size = max(role_font_size * 0.8, 5)
    pdf_time_size = max(time_font_size * 0.8, 5)

    name_y_offset = max(12, pdf_name_size + 5)
    role_y_offset = name_y_offset + max(16, pdf_role_size + 8)
    time_y_offset_from_bottom = max(10, pdf_time_size + 4)

    for node_id, node in nodes.items():
        raw = node.get("raw", {})
        name = raw.get("name", "")
        role = raw.get("role", "")
        time_period = raw.get("time_period", "")
        dept = raw.get("dept", "")
        is_match = raw.get("is_match", True)

        node_color = color_map.get(dept, "#ced4da")

        left, top, bottom = node_box_xy(node_id)

        if filter_active and not is_match:
            fill = colors.HexColor("#f8f9fa")
            stroke = colors.HexColor("#ced4da")
            name_colour = colors.HexColor("#6c757d")
            role_colour = colors.HexColor("#adb5bd")
            time_colour = colors.HexColor("#adb5bd")

        else:
            fill = hex_to_reportlab_colour(node_color, "#ced4da")
            stroke = fill
            name_colour = colors.white
            role_colour = colors.HexColor("#f8f9fa")
            time_colour = colors.HexColor("#e9ecef")

        c.setFillColor(fill)
        c.setStrokeColor(stroke)
        c.roundRect(left, bottom, box_w, box_h, radius=6, fill=1, stroke=1)

        c.setFont("Helvetica-Bold", pdf_name_size)
        c.setFillColor(name_colour)
        draw_centered_wrapped_text(
            c,
            name,
            left + 8,
            top - name_y_offset,
            box_w - 16,
            "Helvetica-Bold",
            pdf_name_size,
            max_lines=2,
        )

        if str(role).strip() != "":
            c.setFont("Helvetica", pdf_role_size)
            c.setFillColor(role_colour)
            draw_centered_wrapped_text(
                c,
                role,
                left + 8,
                top - role_y_offset,
                box_w - 16,
                "Helvetica",
                pdf_role_size,
                max_lines=1,
            )

        if str(time_period).strip() != "":
            c.setFont("Helvetica", pdf_time_size)
            c.setFillColor(time_colour)
            draw_centered_wrapped_text(
                c,
                time_period,
                left + 8,
                bottom + time_y_offset_from_bottom,
                box_w - 16,
                "Helvetica",
                pdf_time_size,
                max_lines=1,
            )

    c.showPage()
    c.save()
    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# 12. Build and Render Chart
# =========================================================
if not clean_df.empty:
    duplicate_names = clean_df[clean_df["Name / Team Name"].duplicated()]["Name / Team Name"].unique().tolist()

    if duplicate_names:
        st.warning(
            "Duplicate names found. The org chart works best when each Name / Team Name is unique: "
            + ", ".join(duplicate_names)
        )

    missing_managers = find_missing_managers(clean_df)

    if missing_managers:
        st.warning(
            "These 'Reports To' names are not found in the Name / Team Name column, so their rows may not appear: "
            + ", ".join(missing_managers)
        )

    loops = detect_loops(clean_df)

    if loops:
        st.error(
            "There is a circular reporting line. Please fix these rows before viewing/exporting: "
            + ", ".join(loops)
        )

    else:
        tree_data = build_chart_tree(clean_df)

        options = {
            "tooltip": {
                "trigger": "item",
                "triggerOn": "click",
                "formatter": "{c}",
                "backgroundColor": "rgba(255, 255, 255, 0.95)",
                "borderColor": "#ccc",
                "borderWidth": 1,
                "textStyle": {"color": "#333"},
            },
            "toolbox": {
                "show": True,
                "right": "30px",
                "top": "15px",
                "feature": {
                    "saveAsImage": {
                        "name": "TC_Org_Chart",
                        "title": "Download as PNG",
                        "pixelRatio": 3,
                    }
                },
            },
            "series": [
                {
                    "type": "tree",
                    "data": [tree_data],
                    "orient": "TB",
                    "top": "5%",
                    "left": "2%",
                    "bottom": "5%",
                    "right": "2%",
                    "symbol": "rect",
                    "symbolSize": [node_width, node_height],
                    "edgeShape": "polyline",
                    "roam": True,
                    "initialTreeDepth": -1,
                    "expandAndCollapse": True,
                    "animationDuration": 550,
                    "animationDurationUpdate": 750,
                    "label": {
                        "position": "insideLeft",
                        "offset": [10, 0],
                        "rich": {
                            "name_active": {
                                "fontSize": name_font_size,
                                "fontWeight": "bold",
                                "color": "#ffffff",
                                "lineHeight": name_font_size + 6,
                            },
                            "role_active": {
                                "fontSize": role_font_size,
                                "color": "#f8f9fa",
                                "lineHeight": role_font_size + 5,
                            },
                            "time_active": {
                                "fontSize": time_font_size,
                                "color": "#e9ecef",
                                "lineHeight": time_font_size + 4,
                            },
                            "name_faded": {
                                "fontSize": name_font_size,
                                "fontWeight": "bold",
                                "color": "#6c757d",
                                "lineHeight": name_font_size + 6,
                            },
                            "role_faded": {
                                "fontSize": role_font_size,
                                "color": "#adb5bd",
                                "lineHeight": role_font_size + 5,
                            },
                            "time_faded": {
                                "fontSize": time_font_size,
                                "color": "#ced4da",
                                "lineHeight": time_font_size + 4,
                            },
                        },
                    },
                }
            ],
        }

        st_echarts(
            options=options,
            height=f"{chart_height}px",
            width=f"{chart_width}px",
        )

        st.markdown("### 📥 Download Full-Scale Chart")

        if REPORTLAB_AVAILABLE:
            try:
                pdf_bytes = make_full_org_chart_pdf(tree_data)
                file_stamp = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y%m%d_%H%M")

                st.download_button(
                    label="Download Whole Org Chart as PDF",
                    data=pdf_bytes,
                    file_name=f"TC_Org_Chart_Full_{file_stamp}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

                st.caption(
                    "This PDF is generated from the full org chart data, so it is not limited to the visible screen area."
                )

            except Exception as e:
                st.error(f"Unable to generate PDF: {e}")

        else:
            st.error(
                "PDF download needs the reportlab package. Add `reportlab` to requirements.txt, then reboot the Streamlit app."
            )

else:
    st.warning("The table is empty. Please upload a file or add at least one person to generate the chart.")
