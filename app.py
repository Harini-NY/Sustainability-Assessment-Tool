import streamlit as st 
import pandas as pd
import json
import sqlite3
from ppt_generator import generate_ppt
# Load Excel file
@st.cache_data
def load_data():
    parameter_df = pd.read_excel("Data/site_config.xlsx", sheet_name="Parameter")
    systems_df = pd.read_excel("Data/site_config.xlsx", sheet_name="Systems")
    return parameter_df, systems_df

parameter_df, systems_df = load_data()

# Saved assesments 


def get_connection():
    return sqlite3.connect("Data/assessments.db")

def create_table():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Assessments (
        ProjectTitle TEXT PRIMARY KEY,
        Site TEXT,
        AssessmentData TEXT,
        LastUpdatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()

create_table()
conn = get_connection()
cursor = conn.cursor()

try:
    cursor.execute("""
        ALTER TABLE Assessments
        ADD COLUMN Islocked INTEGER DEFAULT 0
    """)
except:
    pass

conn.commit()
conn.close()

def save_assessment(project_title, site, data):
    conn = get_connection()
    cursor = conn.cursor()

    data_json = json.dumps(data)

    cursor.execute("""
INSERT INTO Assessments
(ProjectTitle, Site, AssessmentData, LastUpdatedAt)
VALUES (?, ?, ?, CURRENT_TIMESTAMP)

ON CONFLICT(ProjectTitle)
DO UPDATE SET
    Site = excluded.Site,
    AssessmentData = excluded.AssessmentData,
    LastUpdatedAt = CURRENT_TIMESTAMP
""", (project_title, site, data_json))

    conn.commit()
    conn.close()


def load_assessment(project_title):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT Site, AssessmentData
    FROM Assessments
    WHERE ProjectTitle = ?
""", (project_title,))

    row = cursor.fetchone()
    cursor.close()
    conn.close()

    if row:
        site = row [0]
        data = json.loads(row[1])
        return site,data
    return None, None 

def get_saved_assessments(site):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ProjectTitle
        FROM Assessments
        WHERE Site = ?
        ORDER BY ProjectTitle
    """, (site,))
    assessments = [row[0] for row in cursor.fetchall()]
    conn.close()
    return assessments 

def delete_assessment(project_title):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        DELETE FROM Assessments
        WHERE PRojectTitle = ?
    """, (project_title,))
    conn.commit()
    conn.close()

def update_lock_status(project_title, locked):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE Assessments
        SET Islocked = ?
        WHERE ProjectTitle = ?
    """, (locked, project_title))

    conn.commit()
    conn.close()


def get_lock_status(project_title):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Islocked
        FROM Assessments
        WHERE ProjectTitle = ?
    """, (project_title,))

    row = cursor.fetchone()

    conn.close()

    if row:
        return bool(row[0])

    return False



parameter_df.columns = parameter_df.columns.str.strip()
systems_df.columns = systems_df.columns.str.strip()

parameter_df["Site"] = parameter_df["Site"].str.strip()
systems_df["Site"] = systems_df["Site"].str.strip()

parameter_results = []
systems_results = []

st.markdown("""
<style>
details {
    background-color: white;
    border-radius: 8px;
    border: 1px solid #D0D7DE;
}

details summary {
    background-color: #A7C7E7 !important;
    color: black !important;
    padding: 12px;
    border-radius: 8px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)
# Title # Assessment Management
def save_current_assessment():
    site = st.session_state.get("site")
    year = st.session_state.get("project_year")
    if st.session_state.get("selected_assessment") == "New Project":
        project_suffix = st.session_state.get("project_suffix","").strip()
        if not project_suffix:
            st.session_state["save_error"] = True 
            return 
    assessment_data = {}
    for key, value in st.session_state.items():
        if key not in [
            "project_title",
            "selected_assessment",
            "loaded_assessment",
            "site",
            "save_error",
        ]:
            assessment_data[key] = value 

    save_assessment(
        project_title,
        site,
        assessment_data
    )
    st.session_state["is_duplicate"] = False
    st.session_state["loaded_assessment"] = project_title 
    st.session_state["save_success"] = True 

def auto_save_current_assessment():
    if st.session_state.get("skip_auto_save_once",False):
        st.session_state["skip_auto_save_once"] = False
        return
    if st.session_state.get("new_project_reset", False):
        return 
    project_title = st.session_state.get("loaded_assessment")
    #st.write("LOADED ASSESSMENT:", st.session_state.get("loaded_assessment"))
    if not project_title:
        return 
    if get_lock_status(project_title):
        return
    if st.session_state.get("is_duplicate", False):
        return 
    if not st.session_state.get("loaded_assessment"):
        return 
    if get_lock_status(project_title):
        return
    
    saved_site, saved_data = load_assessment(project_title)
    #st.write("DATABASE SITE:", saved_data)
    #st.write("DATABASE DATA:", saved_data)
    if saved_site is None:
        return
    
    
    assessment_data = {}
    excluded_keys = [
        "project_title",
        "selected_assessment",
        "loaded_assessment",
        "site",
        "save_error",
        "save_success",
        "delete_success",
        "is_duplicate",
        "duplicate_requested",
        "duplicate_site"    
    ]
    for key, value in st.session_state.items():
        if key not in excluded_keys:
            assessment_data[key] = value
    if assessment_data == saved_data:
        return
    
    save_assessment(
        project_title,
        saved_site,
        assessment_data
    )
    st.write(
    "loaded=", st.session_state.get("loaded_assessment"),
    "current=", st.session_state.get("project_title")
)


def prepare_duplicate(project_title):
    saved_site, saved_data = load_assessment(project_title)
    for key, value in saved_data.items():
        if key in[
            "selected_assessment",
            "site",
            "loaded_assessment",
            "project_title"
        ]:
            continue
        st.session_state[key] = value
    st.session_state["duplicate_site"] = saved_site
    st.session_state["project_year"] = saved_data.get("project_year", 2026)
    original_suffix = saved_data.get("project_suffix", "").strip()
    if original_suffix:
        st.session_state["project_suffix"] = original_suffix + "Copy"
    else:
        st.session_state["project_suffix"] = "Copy"
    st.session_state["is_duplicate"] = True
    st.session_state["duplicate_requested"] = True
    st.session_state.pop("loaded_assessment", None)
    st.session_state.pop("is_locked", None)

def reset_new_project():
    for key in list(st.session_state.keys()):

        # Checkboxes
        if key.endswith("_na"):
            st.session_state[key] = False

        # Dropdowns
        elif key == "tool_completed":
            st.session_state[key] = "NA"

        elif key == "design_reviewed":
            st.session_state[key] = "NA"

        # Text fields
        elif key == "key_takeaway":
            st.session_state[key] = ""

        elif "mitigation" in key:
            st.session_state[key] = ""

        elif key == "project_suffix":
            st.session_state[key] = ""

        # Numeric inputs
        elif any(x in key for x in [
            "baseline",
            "load",
            "tariff",
            "redundancy",
            "new_load"
        ]):
            st.session_state[key] = 0

    st.session_state.pop("loaded_assessment", None)
    st.session_state["new_project_reset"] = True



st.title("Sustainability Assessment Tool")
if st.session_state.get("duplicate_requested", False):
    st.session_state["site"] = st.session_state["duplicate_site"]
    st.session_state["selected_assessment"] = "New Project"
    st.session_state.pop("duplicate_requested", None)

#Site Dropdown
site_options = ["-- Select Site --"] + list(parameter_df["Site"].unique())
site = st.selectbox("Select Site",site_options, key="site")
if site == "-- Select Site --":
    saved_assessments = []
else:
    saved_assessments = get_saved_assessments(site)
assessment_options = ["New Project"] + saved_assessments

selected_assessment = st.selectbox("Assessment", assessment_options, key="selected_assessment")

LOCKED = (
    get_lock_status(selected_assessment)
    if selected_assessment != "New Project"
    else False
)
if st.session_state.get("reload_project", False):
    saved_site, saved_data = load_assessment(selected_assessment)
    if saved_data is not None:
        for key, value in saved_data.items():
            if key in [ 
                "selected_assessment",
                "site",
                "loaded_assessment",
                "project_title"
            ]:
                continue 
            st.session_state[key] = value 
        st.session_state["loaded_assessment"] = save_assessment
    st.session_state["reload_project"] = False

if selected_assessment == "New Project":

    if (
        st.session_state.get("loaded_assessment")
        and not st.session_state.get("is_duplicate", False)
        and st.session_state.get("selected_assessment_previous") != "New Project"
    ):
        reset_new_project()
        st.rerun()

    year = st.selectbox(
        "Year",
        [2026,2027,2028,2029,2030,2031,2032,2033,2034,2035],
        key ="project_year"
    )
    project_suffix = st.text_input(
        "Project Name",
        key="project_suffix"
    )

    import re

    numbers = []

    for p in saved_assessments:
        match = re.search(rf"{site}{year}\((\d+)\)", p)
        if match:
            numbers.append(int(match.group(1)))

    project_number = 1
    while project_number in numbers:
        project_number += 1

    project_title = (f"{site}{year}({project_number}) - {project_suffix}")
    st.session_state["project_title"] = project_title
    st.session_state["selected_assessment_previous"] = selected_assessment
else:
    project_title = selected_assessment 
    st.session_state["project_title"] = project_title
    if st.session_state.get("loaded_assessment") != project_title:
        saved_site, saved_data = load_assessment(project_title)
        if saved_data is not None:
            for key, value in saved_data.items():
                if key in [
                    "selected_assessment",
                    "site",
                    "loaded_assessment",
                    "project_title"
                ]:
                    continue
                st.session_state[key] = value

            st.session_state["loaded_assessment"] = project_title
            st.session_state["new_project_reset"] = False
            st.session_state["is_duplicate"] = False
            st.rerun()

col1,col2,col3,col4= st.columns(4)

with col1:
    st.button("Save Assessment", type="primary", on_click= save_current_assessment)

with col2:
    if selected_assessment != "New Project":
        if st.button("Delete Assessment"):
            delete_assessment(selected_assessment)

            st.session_state["delete_success"] = True
            st.session_state.pop("loaded_assessment", None)

            st.rerun()
with col3:
    if selected_assessment != "New Project" :
        if st.button("Duplicate Assessment"):
            prepare_duplicate(selected_assessment)
            st.rerun()
with col4:
    if selected_assessment != "New Project":

        if LOCKED:

            if st.button("🔒 Lock"):
                update_lock_status(
                    selected_assessment,
                    0
                )
                st.session_state["skip_auto_save_once"] = True
                st.session_state["reload_project"] = True 
                st.rerun()

        else:

            if st.button("🔓 Unlock"):
                update_lock_status(
                    selected_assessment,
                    1
                )
                st.session_state["skip_auto_save_once"] = True
                st.session_state["reload_project"] = True
                st.rerun()

if st.session_state.get("save_success"):
    st.success(f"Assessment saved successfully!")
    st.session_state.pop("save_success", None)
if st.session_state.get("delete_success"):
    st.success("Assessment deleted successfully!")
    st.session_state.pop("delete_success", None)
if st.session_state.get("save_error"):
    st.error("Please enter a Project Title before saving.")
    st.session_state.pop("save_error", None)


impact_info = {
    "+": (
        "Positive Impact",
        "#CDECCF",
        "Projected load is reduced from the site baseline."
    ),
    "=": (
        "No Significant Impact",
        "#FFE28A",
        "Project has no additional load impact on the site."
    ),
    "-": (
        "Negative Impact",
        "#FFC48C",
        "Load from project increases the site baseline"
    ),
    "- -": (
        "Double Negative Impact",
        "#FFB3B3",
        "Load from project increases site baseline and exceeds the tariff/redundancy limit"
    ),
    "N/A": (
        "Inconclusive Impact",
        "#F3F4F6",
        "Parameter has no effect on impact for this project"
    )
}
#Tabs
tab1,tab2,tab3 = st.tabs(["Parameter", "Systems","Assessment Summary"])
with tab1:
    if site == "-- Select Site --":
        st.info("Please select a site.")
        st.stop()
    site_data = parameter_df[parameter_df["Site"] == site]
    st.header("Parameter")

    for index, row in site_data.iterrows():
        if row["Parameter"] == "Carbon Footprint":
            with st.expander("Carbon Footprint (tCO2e)"):
                st.subheader("Scope 1: Natural Gas Emissions")
                selected_unit = st.text_input("Unit", value=row["Unit"],key=f"scope1_unit_{index}", disabled = LOCKED)
                col1, col2, col3 = st.columns(3)
                with col1:
                    c1,c2 = st.columns([5,2])
                    with c2:
                        st.write("")
                        st.write("")
                        scope1_baseline_na = st.checkbox("N/A",key=f"scope1_baseline_na_{index}", disabled = LOCKED)
                    with c1:
                        scope1_baseline= st.number_input("Site Baseline",disabled= scope1_baseline_na or LOCKED, step=1.0, key=f"scope1_baseline_{index}")
                        if scope1_baseline_na:
                            scope1_baseline = None
                    
                with col2:
                    c1,c2 = st.columns([5,2])
                    with c2:
                        st.write("")
                        st.write("")
                        scope1_load_na = st.checkbox("N/A", key=f"scope1_load_na_{index}", disabled=LOCKED)
                    with c1:
                        scope1_load = st.number_input("Load",disabled= scope1_load_na or LOCKED, step=1.0, key=f"scope1_load_{index}")
                        if scope1_load_na:
                            scope1_load = None

                with col3: 
                    c1,c2 = st.columns([5,2])
                    with c2:
                        st.write("")
                        st.write("")
                        scope1_tariff_na = st.checkbox("N/A", key=f"scope1_tariff_na_{index}", disabled=LOCKED)
                    with c1:
                        scope1_tariff = st.number_input("Tariff Limit",disabled= scope1_tariff_na or LOCKED, step=1.0, key=f"scope1_tariff_{index}")
                        if scope1_tariff_na:
                            scope1_tariff = None 

                st.caption("Choose N/A if the parameter is not applicable")
                
                if (scope1_baseline is None or scope1_load is None):
                    scope1_impact = "N/A"
                    scope1_projected_load = "N/A"
                elif scope1_tariff is None:
                    scope1_projected_load = scope1_baseline + scope1_load
                    if scope1_load == 0:
                        scope1_impact = "="  
                    if scope1_load < 0:
                        scope1_impact = "+"
                    if scope1_load > 0:
                        scope1_impact = "-"
                    
                else:
                    scope1_projected_load = scope1_baseline + scope1_load
                                    
                    if scope1_load == 0:
                        scope1_impact = "="
                    elif scope1_load < 0:
                        scope1_impact = "+"
                    elif scope1_projected_load > scope1_tariff:
                        scope1_impact = "- -"
                    else:
                        scope1_impact = "-"
                                
                   
                st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #87CEEB;padding: 5px;border-radius: 8px;}</style>""", unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Projected Load", scope1_projected_load)
                with col2:
                    impact_title, impact_colour, impact_desc = impact_info.get(scope1_impact,("Unknown", "#E9ECEF", ""))
                    
                    st.markdown(
                        f"""
                    <div style="background-color:{impact_colour}; padding:10px 10px; border-radius:10px; line-height:1.2;">
                        <p style="margin:2px 0;font-size:20px;"><b>{scope1_impact}</b></p>
                        <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                        <p style="margin:0;font-size:13px;">{impact_desc}</p>
                    </div>
                    """,
                        unsafe_allow_html=True
                    )
                
                #Default values 
                scope1_mitigation = "N/A"
                scope1_new_load = "N/A"
                scope1_new_impact = "N/A"
                
                # Mitigation
                if scope1_impact in ["- -", "-", "="]:
                    scope1_mitigation = st.text_area("Mitigation",placeholder="Keep within 20 words.",key=f"scope1_mitigation_{index}", disabled=LOCKED)
                    word_count = len(scope1_mitigation.split())
                    if word_count >20:
                        st.error("Mitigation cannot exceed 20 words.")
                    scope1_new_load = st.number_input("New Load", disabled = LOCKED, value=0.0,step=1.0,key=f"scope1_new_load_{index}")
                            
                # New Impact Calculation 
                    if (scope1_baseline is None or scope1_load is None):
                        scope1_new_projected_load = "N/A" 
                        scope1_new_impact = "N/A"
                    elif scope1_tariff is None :
                        scope1_new_projected_load = scope1_baseline + scope1_new_load
                        if scope1_new_load == 0:
                            scope1_new_impact = "="
                        elif scope1_new_load < 0:
                            scope1_new_impact = "+"
                        elif scope1_new_load > 0:
                            scope1_new_impact = "-"          
                    else:
                        scope1_new_projected_load = scope1_baseline + scope1_new_load
                        if scope1_new_load == 0:
                            scope1_new_impact = "="
                        elif scope1_new_load < 0: 
                            scope1_new_impact = "+"
                        elif scope1_new_projected_load > scope1_tariff:
                            scope1_new_impact = "- -"
                        else: 
                            scope1_new_impact = "-"
                
                    st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #e6a0b8; padding: 5px; order-radius: 8px;}</style>""", unsafe_allow_html=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Projected New Load", scope1_new_projected_load)
                    with col2:
                        impact_title, impact_colour, impact_desc = impact_info.get(scope1_new_impact,("Unknown", "#6c757d", ""))
                        st.markdown(
                            f"""
                            <div style="background-color:{impact_colour}; padding:10px 10px; border-radius:10px; line-height:1.2;">
                                <p style="margin:2px 0;font-size:20px;"><b>{scope1_new_impact}</b></p>
                                <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                                <p style="margin:0;font-size:13px;">{impact_desc}</p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                st.divider()
                st.subheader("Scope 2: Electricity Emissions")

                selected_unit = st.text_input("Unit", value=row["Unit"],key=f"scope2_unit_{index}", disabled= LOCKED)

                col1, col2, col3 = st.columns(3)
                with col1:
                    c1,c2 = st.columns([5,2])
                    with c2:
                        st.write("")
                        st.write("")
                        scope2_baseline_na = st.checkbox("N/A",key=f"scope2_baseline_na_{index}", disabled = LOCKED)
                    with c1:
                        scope2_baseline = st.number_input("Site Baseline", disabled=scope2_baseline_na or LOCKED, step=1.0, key=f"scope2_baseline_{index}")
                        if scope2_baseline_na:
                            scope2_baseline = None
            
                with col2:
                    c1,c2 = st.columns([5,2])

                    with c2:
                        st.write("")
                        st.write("")
                        scope2_load_na = st.checkbox("N/A", key=f"scope2_load_na_{index}", disabled = LOCKED)
                    with c1:
                        scope2_load = st.number_input("Load", disabled= scope2_load_na or LOCKED, step=1.0, key=f"scope2_load_{index}")
                        if scope2_load_na:
                            scope2_load = None

                with col3:
                    c1,c2 = st.columns([5,2])

                    with c2:
                        st.write("")
                        st.write("")
                        scope2_tariff_na = st.checkbox("N/A",key=f"scope2_tariff_na_{index}", disabled = LOCKED)

                    with c1:
                        scope2_tariff = st.number_input("Tariff Limit", disabled= scope2_tariff_na or LOCKED, step=1.0, key=f"scope2_tariff_{index}")
                        if scope2_tariff_na:
                            scope2_tariff = None 

                st.caption("Choose N/A if the parameter is not appliacble")

                if (scope2_baseline is None or scope2_load is None):
                    scope2_impact = "N/A"
                    scope2_projected_load = "N/A"
                elif scope2_tariff is None:
                    scope2_projected_load = scope2_baseline + scope2_load
                    if scope2_load == 0:
                        scope2_impact = "="
                    elif scope2_load < 0:
                        scope2_impact = "+"
                    else:
                        scope2_impact = "- -"
                else:
                    scope2_projected_load = scope2_baseline + scope2_load
                                    
                    if scope2_load == 0:
                        scope2_impact = "="
                    elif scope2_load < 0:
                        scope2_impact = "+"
                    elif scope2_projected_load > scope2_tariff:
                        scope2_impact = "- -"
                    else:
                        scope2_impact = "-"
                                
                   
                st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #87CEEB;padding: 5px;border-radius: 8px;}</style>""", unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Projected Load", scope2_projected_load)
                with col2:
                    impact_title, impact_colour, impact_desc = impact_info.get(scope2_impact, ("Unknown", "#6c757d", ""))
                    st.markdown(
                        f"""
                        <div style="background-color:{impact_colour}; padding:10px 10px; border-radius:10px; line-height:1.2;">
                            <p style="margin:2px 0;font-size:20px;"><b>{scope2_impact}</b></p>
                            <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                            <p style="margin:0;font-size:13px;">{impact_desc}</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                
                #Default values 
                scope2_mitigation = "N/A"
                scope2_new_load = "N/A"
                scope2_new_impact = "N/A"
                
                # Mitigation
                if scope2_impact in ["- -", "-", "="]:
                    scope2_mitigation = st.text_area("Mitigation", placeholder ="Keep within 20 words", key=f"scope2_mitigation_{index}", disabled = LOCKED)
                    word_count = len(scope2_mitigation.split()) 
                    if word_count > 20:
                        st.error("Mitigation cannot exceed 20 words.")
                    scope2_new_load = st.number_input("New Load",value=0.0,step=1.0,key=f"scope2_new_load_{index}", disabled = LOCKED)
                            
                # New Impact Calculation 
                    if (scope2_baseline is None or scope2_load is None):
                        scope2_new_projected_load = "N/A" 
                        scope2_new_impact = "N/A"
                    elif scope2_tariff is None:
                        scope2_new_projected_load = scope2_baseline + scope2_new_load
                        if scope2_new_load == 0:
                            scope2_new_impact = "+"
                        elif scope2_new_load < 0:
                            scope2_new_impact = "+"
                        else:
                            scope2_new_impact = "-"
                        
                    else:
                        scope2_new_projected_load = scope2_baseline + scope2_new_load
                        if scope2_new_load == 0:
                            scope2_new_impact = "="
                        elif scope2_new_load < 0: 
                            scope2_new_impact = "+"
                        elif scope2_new_projected_load > scope2_tariff:
                            scope2_new_impact = "- -"
                        else: 
                            scope2_new_impact = "-"
                
                    st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #e6a0b8; padding: 5px; order-radius: 8px;}</style>""", unsafe_allow_html=True)

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Projected New Load", scope2_new_projected_load)
                    with col2:
                        impact_title, impact_colour, impact_desc = impact_info.get(scope2_new_impact,("Unknown", "#E9ECEF", ""))
                        st.markdown(
                            f"""
                            <div style="background-color:{impact_colour}; padding:10px 10px; border-radius:10px; line-height:1.2;">
                                <p style="margin:2px 0;font-size:20px;"><b>{scope2_new_impact}</b></p>
                                <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                                <p style="margin:0;font-size:13px;">{impact_desc}</p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                parameter_results.append({
                    "Parameter": "Carbon Footprint Scope 1",
                    "Unit": "tCO2e",
                    "Baseline": scope1_baseline,
                    "Projected Load": scope1_projected_load,
                    "Tariff": scope1_tariff,
                    "Impact": scope1_impact,
                    "Mitigation": scope1_mitigation,
                    "New Load": scope1_new_load,
                    "New Impact": scope1_new_impact
                })
                
                parameter_results.append({
                    "Parameter": "Carbon Footprint Scope 2",
                    "Unit": "tCO2e",
                    "Baseline": scope2_baseline,
                    "Projected Load": scope2_projected_load,
                    "Tariff": scope2_tariff,
                    "Impact": scope2_impact,
                    "Mitigation": scope2_mitigation,
                    "New Load": scope2_new_load,
                    "New Impact": scope2_new_impact
                })



        else:
            with st.expander(f"{row['Parameter']} ({row['Unit']})"):
                selected_unit = st.text_input("Unit", value=row["Unit"],key=f"parameter_unit_{index}", disabled = LOCKED)
                
                # Inputs
                col1, col2, col3 = st.columns(3)
                with col1:
                    c1, c2 = st.columns([5,2])

                    with c2:
                        st.write("")
                        st.write("")
                        parameter_baseline_na = st.checkbox("N/A",key=f"parameter_baseline_na_{index}", disabled = LOCKED)

                    with c1:
                        parameter_baseline = st.number_input("Site Baseline",disabled=parameter_baseline_na or LOCKED, step=1.0, key=f"baseline_{index}")

                    if parameter_baseline_na:
                        parameter_baseline = None
            

                with col2:
                    c1, c2 = st.columns([5,2])

                    with c2:
                        st.write("")
                        st.write("")
                        parameter_load_na = st.checkbox("N/A",key=f"parameter_load_na_{index}", disabled= LOCKED)

                    with c1:
                        parameter_load = st.number_input("Load",disabled=parameter_load_na or LOCKED, step=1.0, key=f"load_{index}")
                    if parameter_load_na:
                        parameter_load = None 
                with col3:
                    c1,c2 = st.columns([5,2])
                    with c2:
                        st.write("")
                        st.write("")
                        tariff_na = st.checkbox("N/A", key=f"tariff_na_{index}", disabled = LOCKED)
                    with c1:
                        tariff = st.number_input("Tariff Limit", disabled=tariff_na or LOCKED, step=1.0, key=f"tariff_{index}")
                    if tariff_na:
                        tariff = None
                st.caption("Choose N/A if the parameter is not applicable")


                if parameter_baseline is None or parameter_load is None:
                    parameter_impact = "N/A"
                    parameter_projected_load = "N/A"
                elif tariff is None:
                    parameter_projected_load = parameter_baseline + parameter_load
                    if parameter_load == 0:
                        parameter_impact = "="
                    elif parameter_load < 0:
                        parameter_impact = "+"
                    else:
                        parameter_impact = "-"
                else:
                    parameter_projected_load = parameter_baseline + parameter_load
                    
                    if parameter_load == 0:
                        parameter_impact = "="
                    elif parameter_load < 0:
                        parameter_impact = "+"
                    elif parameter_projected_load > tariff:
                        parameter_impact = "- -"
                    else:
                        parameter_impact = "-"
                
   
                st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #87CEEB;padding: 5px;border-radius: 8px;}</style>""", unsafe_allow_html=True)

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Projected Load", parameter_projected_load)
                with col2:
                    impact_title, impact_colour, impact_desc = impact_info.get(parameter_impact,("Unknown", "#E9ECEF", ""))
                    st.markdown(
                        f"""
                        <div style="
                            background-color:{impact_colour};
                            padding:10px 10px;
                            border-radius:10px;
                            line-height:1.2;
                        ">
                            <p style="margin:2px 0;font-size:20px;"><b>{parameter_impact}</b></p>
                            <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                            <p style="margin:0;font-size:13px;">{impact_desc}</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

    #Default values 
                parameter_mitigation = "N/A"
                parameter_new_load = "N/A"
                parameter_new_impact = "N/A"

    # Mitigation
                if parameter_impact in ["- -", "-", "="]:
                    parameter_mitigation = st.text_area("Mitigation", placeholder= "Keep within 20 words",key=f"parameter_mitigation_{index}", disabled = LOCKED)
                    word_count = len(parameter_mitigation.split())
                    if word_count >20:
                        st.error("Mitigation cannot exceed 20 words.")
                    parameter_new_load = st.number_input("New Load",value=0.0,step=1.0,key=f"parameter_new_load_{index}", disabled = LOCKED)
            
            # New Impact Calculation 
                    if parameter_baseline is None or parameter_load is None:
                        parameter_new_projected_load = "N/A" 
                        parameter_new_impact = "N/A"
                    elif tariff is None: 
                        parameter_new_projected_load = parameter_baseline + parameter_new_load
                        if parameter_new_load == 0:
                            parameter_new_impact = "="
                        elif parameter_new_load < 0:
                            parameter_new_impact = "+"
                        elif parameter_new_load > 0:
                            parameter_new_impact = "-"
                    else:
                        parameter_new_projected_load = (parameter_baseline + parameter_new_load)

                        if parameter_new_load == 0:
                            parameter_new_impact = "="
                        elif parameter_new_load < 0: 
                            parameter_new_impact = "+"
                        elif parameter_new_projected_load > tariff:
                            parameter_new_impact = "- -"
                        else: 
                            parameter_new_impact = "-"

                    st.markdown("""<style>[data-testid="stMetric"] {background-color: #ffffff;border-left: 5px solid #e6a0b8; padding: 5px; order-radius: 8px;}</style>""", unsafe_allow_html=True)
    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Projected New Load", parameter_new_projected_load)
                    with col2:
                        impact_title, impact_colour, impact_desc = impact_info.get(parameter_new_impact,("Unknown", "#E9ECEF", ""))
                        st.markdown(
                            f"""
                            <div style="
                                background-color:{impact_colour};
                                padding:10px 10px;
                                border-radius:10px;
                                line-height:1.2;
                            ">
                                <p style="margin:2px 0;font-size:24px;"><b>{parameter_new_impact}</b></p>
                                <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                                <p style="margin:0;font-size:13px;">{impact_desc}</p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                parameter_results.append({
                    "Parameter": row["Parameter"],
                    "Unit": selected_unit,
                    "Baseline": parameter_baseline,
                    "Projected Load": parameter_projected_load,
                    "Tariff": tariff,
                    "Impact": parameter_impact,
                    "Mitigation": parameter_mitigation,
                    "New Load": parameter_new_load,
                    "New Impact": parameter_new_impact
                })
    
with tab2:

    # Retrieve units of the site
    systems_data = systems_df[systems_df["Site"] == site]

    st.header("Systems")

    for index, row in systems_data.iterrows():

        with st.expander(f"{row['System']} ({row['Units']})"):
            selected_unit = st.text_input("Unit", value=row["Units"],key=f"systems_units_{index}", disabled = LOCKED)

            # Inputs
            col1, col2, col3 = st.columns(3)
            with col1:
                c1,c2 = st.columns([5,2])
                with c2:
                    st.write("")
                    st.write("")
                    systems_baseline_na = st.checkbox("N/A",key=f"systems_baseline_na_{index}", disabled = LOCKED)

                with c1:
                    systems_baseline = st.number_input("Site Baseline", disabled=systems_baseline_na or LOCKED, step=1.0, key=f"system_baseline_{index}")

                if systems_baseline_na:
                    systems_baseline = None

            with col2:
                c1,c2 = st.columns([5,2])
                with c2:
                    st.write("")
                    st.write("")
                    systems_load_na = st.checkbox("N/A",key=f"systems_load_na_{index}", disabled = LOCKED)
                with c1:
                    systems_load = st.number_input("Site Load",disabled=systems_load_na or LOCKED, step=1.0, key=f"system_load_{index}")
                if systems_load_na:
                    systems_load = None

            with col3:
                c1,c2 = st.columns([5,2])
                with c2:
                    st.write("")
                    st.write("")
                    redundancy_na = st.checkbox("N/A",key=f"system_redundancy_na_{index}", disabled = LOCKED)
                with c1:
                    redundancy = st.number_input("Redundancy Limit",disabled = redundancy_na or LOCKED, step=1.0, key=f"system_redundancy_{index}")
                if redundancy_na:
                    redundancy = None

            # Impact Calculation
            if systems_baseline is None or systems_load is None :
                systems_impact = "N/A"
                systems_projected = "N/A"
            elif redundancy is None :
                systems_projected = systems_baseline + systems_load
                if systems_load == 0:
                    systems_impact = "="
                elif systems_load < 0:
                    systems_impact = "+"
                else:
                    systems_impact = "-"

            else:
                systems_projected = systems_baseline + systems_load

                if systems_load == 0:
                    systems_impact = "="

                elif systems_load < 0:
                    systems_impact = "+"

                elif systems_projected > redundancy:
                    systems_impact = "- -"

                else:
                    systems_impact = "-"

            # Display values
            st.markdown("""
            <style>
            [data-testid="stMetric"] {
                background-color: #ffffff;
                border-left: 5px solid #87CEEB;
                padding: 5px;
                border-radius: 8px;
            }
            </style>
            """, unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                st.metric("Projected Load", systems_projected)

            with col2:
                impact_title, impact_colour, impact_desc = impact_info.get(systems_impact,("Unknown", "#E9ECEF", ""))
                st.markdown(
                    f"""
                    <div style="
                        background-color:{impact_colour};
                        padding:10px 10px;
                        border-radius:10px;
                        line-height:1.2;
                    ">
                        <p style="margin:2px 0;font-size:24px;"><b>{systems_impact}</b></p>
                        <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                        <p style="margin:0;font-size:13px;">{impact_desc}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            # Default values
            systems_mitigation = "N/A"
            systems_new_load = "N/A"
            systems_new_impact = "N/A"

            # Mitigation
            if systems_impact in ["- -", "-", "="]:

                systems_mitigation = st.text_area("Mitigation",placeholder="Keep within 20 words", key=f"system_mitigation_{index}", disabled = LOCKED)
                word_count = len(systems_mitigation.split())
                if word_count > 20:
                    st.error("Mitigation cannot exceed 20 words.")

                systems_new_load = st.number_input("New Load",value=0.0,step=1.0,key=f"system_new_load_{index}", disabled = LOCKED)

                # New Impact Calculation
                if systems_baseline is None or systems_load is None:
                    systems_new_projected = "N/A"
                    systems_new_impact = "N/A"
                elif redundancy is None:
                    systems_new_projected = systems_baseline + systems_new_load
                    if systems_new_load == 0:
                        systems_new_impact = "="
                    elif systems_new_load < 0:
                        systems_new_impact = "+"
                    else:
                        systems_new_impact = "-"


                else:
                    systems_new_projected = (systems_baseline + systems_new_load)

                    if systems_new_load == 0:
                        systems_new_impact = "="

                    elif systems_new_load < 0:
                        systems_new_impact = "+"

                    else:
                        if systems_new_projected > redundancy:
                            systems_new_impact = "- -"
                        else:
                            systems_new_impact = "-"

                col1, col2 = st.columns(2)

                with col1:
                    st.metric("Projected New Load",systems_new_projected)

                with col2:
                    impact_title, impact_colour, impact_desc = impact_info.get(systems_new_impact,("Unknown", "#E9ECEF", ""))
                    st.markdown(
                        f"""
                        <div style="
                            background-color:{impact_colour};
                            padding:10px 10px;
                            border-radius:10px;
                            line-height:1.2;
                        ">
                            <p style="margin:2px 0;font-size:24px;"><b>{systems_new_impact}</b></p>
                            <p style="margin:2px 0;font-size:14px;"><b>{impact_title}</b></p>
                            <p style="margin:0;font-size:13px;">{impact_desc}</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
            systems_results.append({
                "System": row["System"],
                "Unit": selected_unit,
                "Baseline": systems_baseline,
                "Projected Load": systems_projected,
                "Redundancy": redundancy,
                "Impact": systems_impact,
                "Mitigation": systems_mitigation,
                "New Load": systems_new_load,
                "New Impact": systems_new_impact})

with tab3:
  st.header("Assessment Summary")

  col1,col2 = st.columns(2)
  with col1:
    tool_completed = st.selectbox("Energy Assessment Tool completed ?", ["NA","Yes", "No"], key="tool_completed", disabled = LOCKED)
  with col2:
    design_reviewed = st.selectbox("Is Sustainable Design AES 00.005STD reviewed?", ["NA","Yes","No"], key="design_reviewed", disabled = LOCKED)

  key_takeaway = st.text_area("Key Takeaway", height = 150, key="key_takeaway", disabled = LOCKED)


  if st.button("Generate PPT"):
    ppt_file = generate_ppt(parameter_results,systems_results,tool_completed,design_reviewed,key_takeaway,site,project_title)
    with open(ppt_file, "rb") as file:
        st.download_button("Download PPT",file,
            file_name=f"{project_title}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

auto_save_current_assessment()