import streamlit as st
import pandas as pd
import sqlite3
import json
from pathlib import Path
from datetime import date, datetime
import html
import base64

st.set_page_config(page_title="Semester 7 ML Project Evaluation", page_icon="📊", layout="wide")
DB_PATH = Path(__file__).with_name("ml_project_evaluation.db")
ROSTER_PATH = Path(__file__).with_name("student_roster.csv")
ASSET_PATH = Path(__file__).with_name("assets")
VIVA_PATH = Path(__file__).with_name("ML_Project_Title_Wise_Viva_Questions_22_Groups.xlsx")
VIVA_CSV_PATH = Path(__file__).with_name("ML_Project_Title_Wise_Viva_Questions_22_Groups.csv")

# Local project illustrations are mapped by group number so the same project
# always uses the same image, regardless of card order or filtering.
_PROJECT_IMAGE_FILES = {
    1: "group_01_travelvista_booking_management_system.jpg",
    2: "group_02_pricemyphone.jpg",
    3: "group_03_resume_screening_system.jpg",
    4: "group_04_smart_demand_products_prediction.jpg",
    5: "group_05_job_seekers_prediction.jpg",
    6: "group_06_food_delivery_time_prediction.jpg",
    7: "group_07_hotel_booking_cancellation.jpg",
    8: "group_08_machine_learning.jpg",
    9: "group_09_house_rent_prediction.jpg",
    10: "group_10_house_price_prediction.jpg",
    11: "group_11_email_phishing_detection.jpg",
    12: "group_12_hostel_energysense.jpg",
    13: "group_13_flight_price_prediction.jpg",
    14: "group_14_telco_churn_prediction_system.jpg",
    15: "group_15_cancer_awareness_and_prediction_system.jpg",
    16: "group_16_examination_preparation_performance_analyzer.jpg",
    17: "group_17_student_stress_level_classification_and_psych.jpg",
    18: "group_18_prediction.jpg",
    19: "group_19_trip_buddy.jpg",
    20: "group_20_machine_learning.jpg",
    21: "group_21_machine_learning.jpg",
    22: "group_22_machine_learning.jpg",
}

def project_image_data(group_no):
    """Return the matching local group illustration as an embeddable data URI."""
    try:
        key = int(float(group_no))
    except (TypeError, ValueError):
        key = -1
    image_file = ASSET_PATH / _PROJECT_IMAGE_FILES.get(key, "group_08_machine_learning.jpg")
    try:
        encoded = base64.b64encode(image_file.read_bytes()).decode("ascii")
        return f"data:image/jpeg;base64,{encoded}"
    except OSError:
        return ""

COMPONENTS = [
    {"name":"Phase 1 – Problem Definition & EDA","max":15,"criteria":[
        ("Problem statement, objectives and real-world relevance",3),("Dataset selection, source and description",2),
        ("Pandas loading, inspection and data understanding",3),("Data cleaning: missing values, duplicates, outliers and transformations",4),
        ("EDA visualizations and meaningful insights",3)],
     "checks":["Problem statement and objectives are clearly defined","Dataset source and attributes are documented","Data is loaded and inspected using Pandas","Missing values and duplicates are checked and handled","Outliers are investigated and treatment is justified","Feature types and transformations are explained","Relevant Matplotlib/Seaborn visualizations are included","EDA findings are interpreted, not merely plotted","Phase 1 report and notebook are submitted"]},
    {"name":"Phase 2 – Evaluation Metrics","max":15,"criteria":[
        ("Correct identification of classification or regression",3),("Appropriate train-test split and prevention of data leakage",3),
        ("Baseline model implementation",3),("Appropriate metrics calculated and justified",4),
        ("Confusion matrix generation and interpretation, where applicable",2)],
     "checks":["Problem type and target variable are correctly identified","Training and testing data are separated appropriately","Preprocessing avoids data leakage","A baseline model is implemented","Classification metrics are used where appropriate","ROC-AUC is used only when suitable","Confusion matrix is interpreted for classification","Metric selection is justified","Regression tasks use suitable metrics such as MAE, RMSE or R²"]},
    {"name":"Phase 3 – Supervised Models","max":25,"criteria":[
        ("At least three suitable supervised algorithms implemented",5),("Correct training, testing and preprocessing",6),
        ("Basic hyperparameter tuning",5),("Comparative performance analysis",4),
        ("Best model selection and interpretation",5)],
     "checks":["At least three suitable supervised algorithms are implemented","Algorithms are appropriate for the task","Training and testing run successfully","Scaling is applied where required","Hyperparameters are tuned at a basic level","Validation strategy is appropriate","Models are evaluated on comparable data","A model comparison table is provided","Best model is selected using evidence","Limitations and improvements are discussed"]},
    {"name":"Phase 4 – Unsupervised Learning","max":15,"criteria":[
        ("At least one suitable technique implemented",4),("Preprocessing and parameter selection",4),
        ("Visualization of clusters or patterns",3),("Interpretation and meaningful insights",2),("Report and code submission",2)],
     "checks":["At least one technique: K-Means, DBSCAN or Apriori","Features are prepared appropriately","Parameters or thresholds are explained","Clusters or patterns are presented clearly","Results are interpreted in context","Findings and limitations are documented","Phase 4 report and code are submitted"]},
    {"name":"Final Project Report","max":10,"criteria":[
        ("Structure, abstract, introduction and objectives",2),("Dataset, methodology and implementation details",2),
        ("Results, graphs and interpretation",2),("Conclusion, limitations and future scope",2),("References, formatting and originality",2)],
     "checks":["Report follows the prescribed 15–25 page guideline","All four phases are documented","Figures and tables have meaningful labels","Code and results match the notebook","References and external sources are acknowledged","Report is organized and readable"]},
    {"name":"Presentation","max":10,"criteria":[
        ("Organization and clarity of slides",2),("Problem, dataset and methodology explained",2),
        ("Demonstration of models and results",2),("Interpretation and conclusions",2),("Communication and time management",2)],
     "checks":["Slides follow a logical sequence","Presentation covers the full ML pipeline","Graphs and results are explained","Demonstration or supporting evidence is prepared","Presentation fits the allotted time"]},
    {"name":"Viva / Demonstration","max":10,"criteria":[
        ("Understanding of problem and dataset",2),("Preprocessing and EDA decisions",2),
        ("Algorithms, training and metrics",2),("Code execution or project demonstration",2),
        ("Results, limitations and individual contribution",2)],
     "checks":["Student explains the problem independently","Student explains preprocessing decisions","Student explains algorithm choices","Student understands evaluation results","Student explains their own contribution"]},
]
SUBMISSION_ITEMS = ["Phase 1 report and EDA notebook","Phase 2 report and metric calculations","Phase 3 report and supervised model comparison","Phase 4 report and unsupervised analysis","Complete executable Jupyter Notebook","Final report (15–25 pages)","Presentation file","Dataset/source documented","Project demonstration completed","Individual viva completed"]
TOTAL_MARKS = sum(c["max"] for c in COMPONENTS)

def default_evaluation():
    return {
        "scores": {c["name"]:{crit:0 for crit,_ in c["criteria"]} for c in COMPONENTS},
        "checklist": {c["name"]:{item:False for item in c["checks"]} for c in COMPONENTS},
        "remarks": {c["name"]:"" for c in COMPONENTS},
        "submission_checks": {item:False for item in SUBMISSION_ITEMS},
        "dataset_source":"","overall_strengths":"","overall_improvements":"",
        "recommendation":"Accepted with minor corrections","last_updated":date.today().isoformat()
    }

def db():
    conn=sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory=sqlite3.Row
    return conn

def adjust_mark_widget(widget_key, delta, max_mark):
    """Safely adjust a marks widget from its button callback."""
    current = int(st.session_state.get(widget_key, 0))
    st.session_state[widget_key] = max(0, min(int(max_mark), current + int(delta)))


def render_phase_cards(state_key, include_final=True):
    """Render all evaluation components in one compact, single-line navigation row."""
    labels = ["Problem & EDA", "Methodology", "Implementation", "Unsupervised", "Report", "Presentation", "Viva"]
    phase_count = len(COMPONENTS)
    total_cards = phase_count + (1 if include_final else 0)
    current = int(st.session_state.get(state_key, 0))
    st.markdown("<div class='phase-nav-label'>Evaluation components</div>", unsafe_allow_html=True)
    # Keep every component on one row. Short single-line labels prevent the tiles from
    # growing vertically; the columns distribute evenly across the available page width.
    cols = st.columns(total_cards, gap="small")
    for idx, col in enumerate(cols):
        with col:
            active = idx == current
            if idx < phase_count:
                maximum = COMPONENTS[idx]["max"]
                label = f"{idx + 1} · {labels[idx]}\n{maximum} marks"
            else:
                label = "✓ Final Review\nOverall"
            if st.button(label, key=f"{state_key}_card_{idx}", type="primary" if active else "secondary", use_container_width=True):
                st.session_state[state_key] = idx
                st.rerun()
    return min(max(current, 0), total_cards - 1)



def get_group_phase_viva_questions(group_no, phase_index):
    """Read the bundled project-specific question bank for one group and phase."""
    if not VIVA_CSV_PATH.exists():
        return []
    try:
        questions_df = pd.read_csv(VIVA_CSV_PATH, encoding="utf-8-sig")
        questions_df["Group"] = pd.to_numeric(questions_df["Group"], errors="coerce")
        questions_df["Question No."] = pd.to_numeric(questions_df["Question No."], errors="coerce")
        phase_prefix = f"Phase {int(phase_index) + 1}"
        subset = questions_df[(questions_df["Group"] == int(group_no)) & questions_df["Phase"].astype(str).str.startswith(phase_prefix)]
        subset = subset.sort_values("Question No.")
        return [(int(row["Question No."]), str(row["Viva Question"])) for _, row in subset.iterrows()]
    except Exception:
        return []


def project_viva_questions(project_title, phase_index):
    """Create phase-specific viva prompts grounded in the student's recorded project title."""
    title = (str(project_title).strip() if project_title is not None else "")
    if not title or title.lower() == "nan":
        title = "your machine learning project"
    title_context = f'"{title}"'
    # Extract meaningful title terms to help the evaluator probe the actual project domain.
    cleaned = title.replace("(", " ").replace(")", " ").replace("/", " ").replace("-", " ")
    words = [w.strip(" ,.:;") for w in cleaned.split() if len(w.strip(" ,.:;")) > 2]
    domain_phrase = " ".join(words[:5]) if words else title
    questions_by_phase = [
        [
            f"What exact real-world problem does {title_context} solve, and who would use the solution?",
            f"What are the input data, expected output, and success criteria for {title_context}?",
            f"Why is machine learning appropriate for {title_context} instead of a rule-based or conventional software solution?",
            f"What assumptions and limitations did you define for {title_context} before starting the analysis?"
        ],
        [
            f"Which dataset did you select for {title_context}, and why does it represent the problem?",
            f"What are the target variable and the most important attributes for {title_context}?",
            f"How did you check missing values, duplicates, outliers, and possible data leakage in this dataset?",
            f"What patterns did your EDA reveal about {domain_phrase}, and how did those patterns affect your next steps?"
        ],
        [
            f"Which supervised algorithms did you compare for {title_context}, and why were they suitable?",
            f"How did you split the data for {title_context}, and how did you prevent training/test data leakage?",
            f"Which evaluation metric best matches the real-world objective of {title_context}, and why?",
            f"Which model performed best for {title_context}; can you explain the result using your actual metric values?"
        ],
        [
            f"What unsupervised technique did you use in {title_context}, and what was the purpose of applying it?",
            f"How did you choose the features and any parameters for the clustering or pattern-discovery step?",
            f"How did you interpret the discovered clusters or patterns in the context of {domain_phrase}?",
            f"How would you check whether the discovered groups or patterns are meaningful and not just noise?"
        ],
        [
            f"Where does the report explain the complete workflow used in {title_context}?",
            f"Which table, chart, or result provides the strongest evidence for your main conclusion?",
            f"Can you trace one reported result in {title_context} back to the code and notebook cell that produced it?",
            f"What limitations and future improvements did you document for {title_context}?"
        ],
        [
            f"Explain {title_context} in one minute: problem, data, method, and main result.",
            f"Which slide best supports the main claim made by {title_context}, and what does its evidence show?",
            f"If a stakeholder challenged the model result in {title_context}, how would you justify your conclusion?",
            f"What is the most important limitation of {title_context}, and how would you address it next?"
        ],
        [
            f"Walk through one complete example from input to output in {title_context}.",
            f"Why did you choose the final model or method for {title_context}, and what alternative did you reject?",
            f"Show how you would explain one prediction, metric, cluster, or pattern produced by {title_context}.",
            f"Which part of {title_context} did you implement personally, and how can you demonstrate it in the code?"
        ],
        [
            f"What are the three strongest pieces of evidence that {title_context} meets its stated objectives?",
            f"Can you reproduce the key result of {title_context} from the submitted notebook and explain the output?",
            f"What important limitation remains in {title_context}, and what would you improve before deployment?",
            f"How would you explain the value and risks of {title_context} to a non-technical stakeholder?"
        ],
    ]
    idx = max(0, min(int(phase_index), len(questions_by_phase)-1))
    return questions_by_phase[idx]

def init_db():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS students (
            roll_no TEXT PRIMARY KEY, student_name TEXT NOT NULL, group_no INTEGER,
            ml_project_title TEXT NOT NULL DEFAULT '')""")
        con.execute("""CREATE TABLE IF NOT EXISTS evaluations (
            roll_no TEXT PRIMARY KEY REFERENCES students(roll_no), evaluation_json TEXT NOT NULL,
            updated_at TEXT NOT NULL)""")
        con.execute("""CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, occurred_at TEXT NOT NULL,
            activity_type TEXT NOT NULL, description TEXT NOT NULL, details TEXT DEFAULT '')""")
        # Seed from the supplied class roster only when the database has no students.
        count=con.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        if count==0 and ROSTER_PATH.exists():
            roster=pd.read_csv(ROSTER_PATH,dtype=str).fillna("")
            for _,r in roster.iterrows():
                roll=str(r["roll_no"]).strip()
                con.execute("INSERT OR IGNORE INTO students VALUES (?,?,?,?)",
                            (roll,str(r["student_name"]).strip(),int(r["group_no"]),str(r["ml_project_title"]).strip()))
                con.execute("INSERT OR IGNORE INTO evaluations VALUES (?,?,?)",
                            (roll,json.dumps(default_evaluation()),date.today().isoformat()))
        else:
            # Add newly listed roster entries without overwriting existing marks.
            if ROSTER_PATH.exists():
                roster=pd.read_csv(ROSTER_PATH,dtype=str).fillna("")
                for _,r in roster.iterrows():
                    roll=str(r["roll_no"]).strip()
                    con.execute("INSERT OR IGNORE INTO students VALUES (?,?,?,?)",
                                (roll,str(r["student_name"]).strip(),int(r["group_no"]),str(r["ml_project_title"]).strip()))
                    con.execute("INSERT OR IGNORE INTO evaluations VALUES (?,?,?)",
                                (roll,json.dumps(default_evaluation()),date.today().isoformat()))

def get_students():
    with db() as con:
        return pd.read_sql_query("""SELECT s.roll_no,s.student_name,s.group_no,s.ml_project_title,e.updated_at
            FROM students s LEFT JOIN evaluations e ON s.roll_no=e.roll_no ORDER BY CAST(s.roll_no AS INTEGER)""",con)

def get_evaluation(roll):
    with db() as con:
        row=con.execute("SELECT evaluation_json FROM evaluations WHERE roll_no=?",(roll,)).fetchone()
    if not row:
        return default_evaluation()
    saved=json.loads(row["evaluation_json"])
    defaults=default_evaluation()
    # Backfill newly introduced checklist/remarks fields for older seeded or restored records.
    for section in ("scores","checklist","remarks"):
        for comp in COMPONENTS:
            cname=comp["name"]
            saved.setdefault(section,{})
            saved[section].setdefault(cname,defaults[section][cname])
            if isinstance(defaults[section][cname],dict):
                for key,value in defaults[section][cname].items():
                    saved[section][cname].setdefault(key,value)
    for key,value in defaults.items():
        saved.setdefault(key,value)
    for item,value in defaults["submission_checks"].items():
        saved["submission_checks"].setdefault(item,value)
    return saved

def save_evaluation(roll,ev):
    ev["last_updated"]=date.today().isoformat()
    with db() as con:
        con.execute("INSERT OR REPLACE INTO evaluations(roll_no,evaluation_json,updated_at) VALUES (?,?,?)",
                    (roll,json.dumps(ev,ensure_ascii=False),ev["last_updated"]))

def marks_total(ev):
    return sum(sum(float(v) for v in ev["scores"][c["name"]].values()) for c in COMPONENTS)


def generate_overall_feedback(ev):
    """Generate strengths and improvement areas from the marks recorded across all phases."""
    component_results = []
    criterion_results = []
    awarded_total = 0.0
    maximum_total = 0
    any_marks = False

    for component in COMPONENTS:
        component_name = component["name"]
        component_max = component["max"]
        score_map = ev.get("scores", {}).get(component_name, {})
        component_score = sum(float(score_map.get(criterion, 0) or 0) for criterion, _ in component["criteria"])
        awarded_total += component_score
        maximum_total += component_max
        any_marks = any_marks or component_score > 0
        component_pct = (component_score / component_max * 100) if component_max else 0
        component_results.append((component_name, component_score, component_max, component_pct))

        for criterion, criterion_max in component["criteria"]:
            score = float(score_map.get(criterion, 0) or 0)
            criterion_results.append((component_name, criterion, score, criterion_max, score / criterion_max if criterion_max else 0))

    if not any_marks:
        message = "No marks have been recorded yet. Enter and save marks across the evaluation phases to generate an overall summary."
        return message, message

    strengths = []
    improvements = []
    ranked_components = sorted(component_results, key=lambda item: item[3], reverse=True)

    strong_components = [item for item in component_results if item[3] >= 75]
    if strong_components:
        names = ", ".join(name for name, score, maximum, pct in sorted(strong_components, key=lambda item: item[3], reverse=True))
        strengths.append(f"Strong-performing components: {names}.")
    best = ranked_components[0]
    if best[1] > 0 and not strong_components:
        strengths.append(f"Current strongest component: {best[0]}.")

    # Show strong criteria as a simple checklist; deliberately omit all marks and percentages.
    strong_criteria = [item for item in criterion_results if item[4] >= 0.8 and item[2] > 0]
    for component_name, criterion, score, maximum, ratio in strong_criteria[:8]:
        strengths.append(f"{criterion} — {component_name}.")

    low_components = [item for item in component_results if item[3] < 60]
    if low_components:
        names = ", ".join(f"{name} ({score:g}/{maximum}, {pct:.0f}%)" for name, score, maximum, pct in sorted(low_components, key=lambda item: item[3]))
        improvements.append(f"Components needing attention (below 60%): {names}.")

    low_criteria = [item for item in criterion_results if item[4] < 0.5]
    for component_name, criterion, score, maximum, ratio in low_criteria[:6]:
        improvements.append(f"Review '{criterion}' in {component_name} ({score:g}/{maximum}); strengthen the implementation or supporting explanation.")

    if awarded_total < maximum_total * 0.6:
        improvements.insert(0, f"Overall score is {awarded_total:g}/{maximum_total} ({awarded_total / maximum_total * 100:.0f}%). Prioritize the lowest-scoring criteria before final submission.")

    if not strengths:
        strengths.append("Marks have been entered, but no component has reached the strong-performance threshold yet. Continue recording marks as the project is assessed.")
    if not improvements:
        improvements.append("No component is currently below the improvement threshold. Review the detailed criteria and viva responses to identify any finer improvements.")

    return "\n".join(f"• {item}" for item in strengths), "\n".join(f"• {item}" for item in improvements)

init_db()

def log_activity(activity_type, description, details=""):
    """Persist an audit-friendly activity entry in the same SQLite database."""
    try:
        with db() as con:
            con.execute("INSERT INTO activity_log(occurred_at, activity_type, description, details) VALUES (?,?,?,?)",
                        (datetime.now().astimezone().isoformat(timespec="seconds"), str(activity_type), str(description), str(details)))
    except Exception:
        pass

def get_recent_activity(limit=12):
    with db() as con:
        return con.execute("SELECT occurred_at, activity_type, description, details FROM activity_log ORDER BY id DESC LIMIT ?", (int(limit),)).fetchall()

def open_project_group(group_label):
    # Streamlit callbacks run before the next script rerun, so navigation state
    # can safely be changed here rather than after the radio widget is created.
    st.session_state["selected_evaluation_group"] = str(group_label)
    log_activity("Group opened", f"Opened Group {group_label} details", "Project Group Details")
    st.session_state["Navigation"] = "Project Group Details"

def back_to_groups():
    # Change the navigation widget's value from its callback, not from the
    # body of the script after the widget has been instantiated.
    st.session_state["Navigation"] = "Student Evaluation"

# Global sidebar design: polished navigation panel shared by every page.
st.markdown("""
<style>
/* Explicit, visible scrollbars on the main page and sidebar. Streamlit's main
   viewport is a dedicated scroll region so long forms never get clipped. */
html, body { min-height:100% !important; height:100% !important; overflow:hidden !important; overflow-x:hidden !important; }
[data-testid="stAppViewContainer"] {
    height:100vh !important; min-height:100vh !important; max-height:100vh !important;
    overflow:hidden !important; display:flex !important;
}
[data-testid="stMain"] {
    height:100vh !important; min-height:0 !important; max-height:100vh !important;
    flex:1 1 auto !important; overflow-y:scroll !important; overflow-x:hidden !important;
    scrollbar-gutter:stable; overscroll-behavior-y:contain;
}
[data-testid="stMainBlockContainer"], [data-testid="stAppViewBlockContainer"] {
    min-height:auto !important; height:auto !important; max-height:none !important;
    overflow:visible !important; box-sizing:border-box !important; width:100% !important; max-width:100% !important;
}
[data-testid="stMainBlockContainer"] {
    padding-bottom:clamp(2rem,5vh,4rem) !important; scrollbar-gutter:stable;
}
/* Keep scrollbar tracks visible and easy to use on Chromium/Windows and Firefox. */
[data-testid="stMain"], [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
    scrollbar-width:auto !important; scrollbar-color:#8fa4c2 #edf2f8 !important;
}
[data-testid="stMain"]::-webkit-scrollbar,
[data-testid="stSidebar"]::-webkit-scrollbar,
[data-testid="stSidebar"] > div:first-child::-webkit-scrollbar { width:12px !important; }
[data-testid="stMain"]::-webkit-scrollbar-track,
[data-testid="stSidebar"]::-webkit-scrollbar-track,
[data-testid="stSidebar"] > div:first-child::-webkit-scrollbar-track { background:#edf2f8 !important; }
[data-testid="stMain"]::-webkit-scrollbar-thumb,
[data-testid="stSidebar"]::-webkit-scrollbar-thumb,
[data-testid="stSidebar"] > div:first-child::-webkit-scrollbar-thumb { background:#8fa4c2 !important; border:3px solid #edf2f8 !important; border-radius:10px !important; }
[data-testid="stSidebar"] { height:100vh !important; max-height:100vh !important; overflow-y:scroll !important; overflow-x:hidden !important; overscroll-behavior:contain; }
[data-testid="stSidebar"] > div:first-child { min-height:100%; height:auto !important; max-height:none !important; overflow-y:visible !important; }
/* Keep long labels, headings, and cards within their available width. */
[data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"],
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] { min-width:0; overflow-wrap:anywhere; }
[data-testid="stHorizontalBlock"] { min-width:0; }
@media (max-width: 1000px) {
    [data-testid="stMainBlockContainer"] { padding-left:1rem !important; padding-right:1rem !important; }
    .group-hero { grid-template-columns:64px minmax(0,1fr) !important; }
    .group-hero-icon { grid-row:1; }
    .group-hero-title { grid-column:2; }
    .group-stat { grid-column:span 1; }
    .group-list-hero { align-items:flex-start; flex-wrap:wrap; }
    .group-list-stats { width:100%; flex-wrap:wrap; }
}
@media (max-width: 700px) {
    [data-testid="stMainBlockContainer"] { padding-top:2rem !important; padding-left:.75rem !important; padding-right:.75rem !important; }
    [data-testid="stHorizontalBlock"] { flex-wrap:wrap !important; gap:.65rem !important; }
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        min-width:100% !important; width:100% !important; flex:1 1 100% !important;
    }
    .group-hero { grid-template-columns:minmax(0,1fr) !important; gap:10px !important; }
    .group-hero-icon, .group-hero-title, .group-stat { grid-column:1 !important; }
    .group-hero-icon { grid-row:auto !important; }
    .group-list-hero { padding:16px !important; }
    .group-list-stats { flex-direction:column !important; }
    .group-list-stat { width:100% !important; }
    .group-tile { min-height:unset !important; }
    .criterion-row { flex-wrap:wrap; gap:.6rem; }
    .final-review-head { flex-wrap:wrap; }
    .final-review-panel { min-width:0; }
}
:root { --ml-navy:#14233f; --ml-teal:#087f8c; --ml-muted:#718096; }
/* Use the available canvas and start page content near the top. */
[data-testid="stMainBlockContainer"] {
    max-width: 100% !important;
    width: 100% !important;
    padding-top: 1.25rem !important;
    padding-left: clamp(1rem, 2.4vw, 2.5rem) !important;
    padding-right: clamp(1rem, 2.4vw, 2.5rem) !important;
}
[data-testid="stAppViewBlockContainer"] { padding-top: .35rem !important; overflow: visible !important; }
[data-testid="stMain"] { overflow-y: scroll !important; overflow-x: hidden !important; }
[data-testid="stMainBlockContainer"] > div:first-child { overflow: visible !important; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #183653 0%, #102944 100%) !important;
    border-right: 1px solid #0b2038;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 1rem; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { margin-bottom: .35rem; }
.sidebar-brand {
    display:flex; align-items:center; gap:.7rem; padding:.8rem .65rem 1rem;
    border-bottom:1px solid rgba(255,255,255,.16); margin-bottom:1.05rem;
}
.sidebar-brand-mark {
    width:42px; height:42px; border-radius:13px; display:flex; align-items:center;
    justify-content:center; color:#fff; font-size:1.2rem; font-weight:850;
    background:linear-gradient(145deg,#0f766e,#2563eb); box-shadow:0 5px 12px rgba(37,99,235,.18);
}
.sidebar-brand-title { color:#ffffff; font-size:.94rem; font-weight:850; line-height:1.25; }
.sidebar-brand-subtitle { color:#b8c9dc; font-size:.68rem; margin-top:.18rem; }
.sidebar-section-label {
    color:#9fb4ca; text-transform:uppercase; letter-spacing:.12em; font-size:.65rem;
    font-weight:800; padding:.1rem .75rem .45rem;
}
[data-testid="stSidebar"] [data-testid="stRadio"] > label { display:none; }
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] { gap:.3rem; }
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] {
    padding:.62rem .72rem; border:1px solid transparent; border-radius:11px;
    transition:background .15s ease,border-color .15s ease;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"]:hover {
    background:rgba(255,255,255,.09); border-color:rgba(255,255,255,.12);
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"][aria-checked="true"] {
    background:linear-gradient(100deg,#2563eb,#1d4ed8); border-color:#3b82f6;
    box-shadow:inset 3px 0 0 #ffffff;
}
/* Force high-contrast navigation text across Streamlit versions. The label may
   be rendered as p, span, or nested div depending on the Streamlit release. */
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"],
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] *,
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] p,
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] span,
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] div {
    color:#f8fbff !important;
    font-size:.95rem !important;
    font-weight:650 !important;
    opacity:1 !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"][aria-checked="true"],
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"][aria-checked="true"] * {
    color:#ffffff !important;
    font-weight:800 !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"] svg {
    color:#f8fbff !important;
    fill:#f8fbff !important;
}
/* Final high-specificity override: Streamlit versions may render radio labels
   outside the [role="radio"] node. Keep every Workspace option white. */
[data-testid="stSidebar"] [data-testid="stRadio"] label,
[data-testid="stSidebar"] [data-testid="stRadio"] label *,
[data-testid="stSidebar"] [data-testid="stRadio"] label p,
[data-testid="stSidebar"] [data-testid="stRadio"] label span,
[data-testid="stSidebar"] [data-testid="stRadio"] label div,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"],
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] *,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] p,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] span,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] div {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
    text-shadow:none !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label p,
[data-testid="stSidebar"] [data-testid="stRadio"] label span,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] p,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] span {
    font-size:.95rem !important;
    font-weight:650 !important;
}
.sidebar-footer {
    margin-top:1.4rem; padding:.8rem .75rem; border:1px solid rgba(255,255,255,.18);
    background:rgba(255,255,255,.07); border-radius:13px;
}
.sidebar-footer-title { color:#ffffff; font-size:.76rem; font-weight:800; }
.sidebar-footer-copy { color:#b8c9dc; font-size:.68rem; line-height:1.45; margin-top:.2rem; }

/* Phase evaluation workspace: polished phase selector and criterion rows. */
.phase-nav-label { color:#718096; font-size:.76rem; font-weight:800; letter-spacing:.1em; text-transform:uppercase; margin:.2rem 0 .65rem; }
/* Evaluation component selectors stay compact and single-line across the row. */
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] button { white-space:nowrap; }
div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] button p { white-space:nowrap; font-size:clamp(.58rem, .72vw, .82rem); line-height:1.15; }
.phase-header { display:flex; align-items:center; justify-content:space-between; gap:1.25rem; padding:1.1rem 1.25rem; border:1px solid #dbe7f7; border-radius:18px; background:linear-gradient(115deg,#ffffff 0%,#f4f8ff 100%); margin:.7rem 0 1rem; }
.phase-title { color:#102a56; font-size:1.55rem; font-weight:850; line-height:1.25; }
.phase-subtitle { color:#718096; font-size:.88rem; margin-top:.35rem; }
.phase-score-card { min-width:245px; border:1px solid #dbe7f7; border-radius:14px; padding:.8rem 1rem; background:#fff; }
.phase-score-label { color:#64748b; font-size:.78rem; font-weight:700; }
.phase-score-line { display:flex; align-items:center; gap:.75rem; margin-top:.25rem; }
.phase-score-number { color:#2563eb; font-size:1.65rem; font-weight:900; white-space:nowrap; }
.phase-score-number span { color:#172b4d; font-size:1.1rem; }
.phase-progress { height:8px; background:#e2eaf5; border-radius:20px; overflow:hidden; flex:1; min-width:50px; }
.phase-progress-fill { height:100%; background:linear-gradient(90deg,#2563eb,#3984ff); border-radius:20px; }
.criterion-row { display:flex; align-items:center; gap:1rem; padding:.8rem .95rem; border:1px solid #dce7f7; border-radius:13px; background:#fff; margin:.45rem 0; min-height:75px; }
.criterion-number { display:flex; align-items:center; justify-content:center; width:42px; height:42px; min-width:42px; border-radius:12px; background:#eef4ff; color:#23416e; font-weight:850; font-size:1rem; }
.criterion-main { flex:1; min-width:0; }
.criterion-title { color:#102a56; font-size:.93rem; font-weight:750; line-height:1.35; }
.criterion-max { display:inline-block; margin-top:.35rem; border-radius:99px; padding:.18rem .6rem; color:#2563eb; background:#eaf1ff; font-size:.74rem; font-weight:750; }
[data-testid="stRadio"]:has(.phase-selector) { width:100%; }
[class*="st-key-phase_selector_container"] [role="radiogroup"] { display:grid; grid-template-columns:repeat(8,minmax(0,1fr)); gap:.5rem; width:100%; }
[class*="st-key-phase_selector_container"] [role="radio"] { display:flex; justify-content:center; align-items:center; text-align:center; padding:.85rem .45rem; min-height:58px; border:1px solid #e1eaf6; border-radius:13px; background:#f5f8fd; color:#19365f; font-weight:750; box-shadow:0 1px 2px rgba(18,42,80,.03); }
[class*="st-key-phase_selector_container"] [role="radio"][aria-checked="true"] { background:linear-gradient(120deg,#2563eb,#1d4ed8); border-color:#2563eb; box-shadow:0 6px 15px rgba(37,99,235,.18); }
[class*="st-key-phase_selector_container"] [role="radio"][aria-checked="true"] p { color:#fff !important; }
[class*="st-key-phase_selector_container"] [role="radio"] p { font-size:.82rem; margin:0; font-weight:750; }
@media (max-width: 1000px) { [class*="st-key-phase_selector_container"] [role="radiogroup"] { grid-template-columns:repeat(4,minmax(0,1fr)); } .phase-header { align-items:flex-start; flex-direction:column; } .phase-score-card { width:100%; } }

/* Responsive phase cards shared by individual and group evaluation. */
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button,
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button {
  min-height: 132px; height: auto; white-space: pre-line; overflow-wrap: anywhere;
  line-height: 1.45; font-size: .9rem; font-weight: 700; padding: .75rem .45rem;
  border: 1px solid #dce6f3; border-radius: 13px; box-shadow: 0 2px 7px rgba(20,45,80,.04);
}
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button[kind="primary"],
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button[kind="primary"] {
  border: 2px solid #2563eb; background: #eef5ff; color: #1455c0;
}
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button p,
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button p { white-space: pre-line; line-height: 1.4; }
@media(max-width:850px) {
  [class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button,
  [class*="st-key-phase_nav_index_"] [data-testid="stButton"] button { min-height: 112px; font-size: .82rem; padding: .55rem .3rem; }
}

/* Evaluation workspace inspired by the requested faculty portal reference. */
.eval-project-banner{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(160px,.7fr) minmax(230px,1fr);gap:18px;align-items:center;padding:18px 20px;margin:4px 0 16px;border:1px solid #d9e7fb;border-radius:16px;background:linear-gradient(110deg,#f2f7ff,#ffffff 58%,#f1f8ff);box-sizing:border-box;width:100%;}
.eval-project-main{display:flex;align-items:center;gap:14px;min-width:0}.eval-project-icon{display:flex;align-items:center;justify-content:center;width:58px;height:58px;min-width:58px;border-radius:15px;background:linear-gradient(145deg,#347eff,#1d5fe8);color:white;font-size:26px}.eval-project-copy{min-width:0}.eval-project-title{font-size:clamp(17px,1.5vw,23px);font-weight:850;line-height:1.3;color:#10264d;overflow-wrap:anywhere}.eval-project-subtitle{font-size:14px;color:#536987;margin-top:5px;line-height:1.45;overflow-wrap:anywhere}.eval-project-students{border-left:1px solid #c7d5e8;padding-left:18px;min-width:0}.eval-mini-label{font-size:13px;color:#536987;font-weight:650;margin-bottom:8px}.eval-avatar-row{display:flex;gap:6px;flex-wrap:wrap}.eval-avatar{width:35px;height:35px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#e3ebf6;color:#294766;font-size:12px;font-weight:800;border:1px solid #d7e1ef}.eval-progress-wrap{min-width:0}.eval-progress-head{display:flex;justify-content:space-between;gap:10px;color:#3d5476;font-size:13px;margin-bottom:9px}.eval-progress-track{height:10px;background:#dfe6f0;border-radius:99px;overflow:hidden}.eval-progress-fill{height:100%;background:linear-gradient(90deg,#2563eb,#3b82f6);border-radius:99px}.eval-progress-caption{font-size:12px;color:#516987;margin-top:6px}
.eval-student-line{display:flex;flex-wrap:wrap;gap:8px 16px;align-items:center;padding:0 2px 10px;color:#536987;font-size:13px}.eval-student-line strong{color:#18365d}
.phase-score-card{min-width:280px;max-width:420px;flex:0 1 420px;border:1px solid #dce8f7;border-radius:15px;padding:1rem 1.1rem;background:linear-gradient(135deg,#f7fbff,#eef6ff)}
.criterion-title{font-size:1rem;}.criterion-max{font-size:.8rem}.criterion-number{width:44px;height:44px;min-width:44px;font-size:1.05rem}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radio"]{padding:.72rem .78rem;border-radius:10px;}
@media(max-width:1050px){.eval-project-banner{grid-template-columns:minmax(0,1fr) minmax(0,1fr)}.eval-project-progress,.eval-progress-wrap{grid-column:1/-1}.eval-project-students{border-left:0;padding-left:0}}
@media(max-width:650px){.eval-project-banner{grid-template-columns:1fr;padding:14px}.eval-project-students,.eval-progress-wrap{grid-column:1}.phase-title{font-size:1.25rem}.phase-score-card{min-width:0;max-width:none;width:100%}.criterion-title{font-size:.94rem}}
.group-breadcrumb{display:block;box-sizing:border-box;width:100%;min-height:42px;padding:10px 0 14px;border-bottom:1px solid #e6edf6;color:#536987;font-size:15px;line-height:1.5;margin:0 0 16px;overflow:visible;overflow-wrap:anywhere}
.group-breadcrumb span{color:#8292aa}
.group-hero{display:grid;grid-template-columns:76px minmax(190px,1.4fr) repeat(3,minmax(0,1fr));gap:14px;align-items:center;margin:18px 0 24px;padding:4px 0;width:100%;box-sizing:border-box;overflow:visible;min-height:100px}
.group-hero-icon{height:76px;width:76px;border-radius:18px;background:linear-gradient(135deg,#e9f2ff,#dceaff);display:flex;align-items:center;justify-content:center;color:#2563eb;font-size:38px}
.group-hero-title{min-width:0}
.group-hero-title h1{font-size:clamp(21px,2vw,30px);font-weight:800;color:#102b63;line-height:1.18;margin:0 0 10px;overflow-wrap:anywhere}
.group-hero-title p{font-size:15px;color:#26436d;margin:0;line-height:1.45;overflow-wrap:anywhere;word-break:normal}
.group-hero-title p strong{color:#657894}
.group-stat{display:flex;align-items:center;gap:10px;min-width:0;min-height:88px;padding:12px;border:1px solid #e4ebf4;border-radius:16px;background:linear-gradient(145deg,#fff,#f8fbff);box-shadow:0 4px 14px rgba(22,55,105,.04);box-sizing:border-box;overflow:hidden}
.group-stat-icon{width:44px;height:44px;min-width:44px;flex-shrink:0;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:700}
.group-stat-icon.blue{background:#e7f0ff;color:#2563eb}.group-stat-icon.purple{background:#f0eaff;color:#7044d6}.group-stat-icon.green{background:#e3f6ef;color:#15966d}
.group-stat > div:last-child{min-width:0;flex:1}
.group-stat b{display:block;font-size:clamp(18px,1.5vw,25px);color:#102b63;line-height:1.2;white-space:normal;overflow-wrap:anywhere;word-break:normal}
.group-stat span{display:block;color:#73839d;font-size:13px;margin-top:5px;line-height:1.3;overflow-wrap:anywhere}
@media(max-width:1150px){.group-hero{grid-template-columns:64px minmax(0,1fr) repeat(3,minmax(120px,.75fr));gap:10px}.group-hero-icon{width:64px;height:64px}.group-stat{padding:10px;gap:8px}.group-stat-icon{width:38px;height:38px;min-width:38px;font-size:19px}}
@media(max-width:900px){.group-hero{grid-template-columns:64px minmax(0,1fr);}.group-hero-icon{grid-row:1}.group-hero-title{grid-column:2}.group-stat{grid-column:span 2;min-height:76px}.group-stat:nth-of-type(3){grid-column:span 1}.group-stat:nth-of-type(4){grid-column:span 1}.group-stat:nth-of-type(5){grid-column:span 2}}
@media(max-width:560px){.group-hero{grid-template-columns:1fr;gap:10px}.group-hero-icon,.group-hero-title,.group-stat{grid-column:1!important}.group-stat{min-height:70px}}
.partner-card{height:340px;box-sizing:border-box;border:1px solid #dce8ff;border-radius:20px;padding:24px 22px 20px;background:linear-gradient(155deg,#fff 0%,#fbfdff 100%);box-shadow:0 8px 24px rgba(28,73,140,.07);margin-bottom:10px;display:flex;flex-direction:column;justify-content:space-between}.partner-top{display:flex;align-items:center;gap:18px;min-height:100px;padding-bottom:18px;border-bottom:1px solid #e8eef7}.partner-avatar{height:88px;width:88px;flex-shrink:0;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:28px;font-weight:800;box-shadow:inset 0 0 0 1px rgba(255,255,255,.65)}.partner-identity{min-width:0;flex:1}.partner-identity h3{font-size:19px;line-height:1.35;color:#0c285b;font-weight:800;margin:0 0 10px;overflow-wrap:anywhere}.partner-identity p{font-size:15px;color:#71829d;margin:0;line-height:1.7}.group-pill{display:inline-block;background:#e8f1ff;color:#1263ed;border-radius:20px;padding:4px 11px;font-weight:750;white-space:nowrap}
.partner-marks{background:#f6f9fe;border:1px solid #edf2fa;border-radius:14px;padding:18px;margin:18px 0 16px}.partner-marks-line{display:flex;align-items:center;justify-content:space-between;color:#526581;font-size:17px;font-weight:650;margin-bottom:12px}.partner-marks-line strong{font-size:28px;color:#0964ed}.partner-marks-line em{font-size:18px;color:#687a95;font-style:normal;font-weight:550}.partner-progress{height:12px;background:#e1e7ef;border-radius:20px;overflow:hidden;margin-bottom:13px}.partner-progress span{height:100%;display:block;background:linear-gradient(90deg,#2563eb,#397cf5);border-radius:20px}.partner-marks small{font-size:15px;color:#687b98}.partner-card + div [data-testid='stButton'] button{height:58px;min-height:58px;border-radius:11px;font-size:17px;font-weight:750;background:linear-gradient(100deg,#1263ed,#0b5ce5);border:1px solid #1263ed;color:#fff;box-shadow:0 5px 12px rgba(18,99,237,.14)}.partner-card + div [data-testid='stButton'] button:hover{background:#084fca;color:#fff;border-color:#084fca}
@media(max-width:1200px){.group-hero{grid-template-columns:75px minmax(220px,1fr) repeat(3,minmax(120px,.7fr));gap:10px}.group-stat{padding:12px;gap:9px}.group-stat-icon{height:42px;width:42px}.group-stat b{font-size:20px}}
@media(max-width:900px){.group-hero{grid-template-columns:70px 1fr}.group-stat{grid-column:span 1}.partner-card{padding:16px}.partner-avatar{width:58px;height:58px;font-size:21px}}

/* Group selector inspired by the supplied dashboard reference. */
.group-list-hero{display:flex;align-items:center;justify-content:space-between;gap:22px;padding:22px 24px;margin:10px 0 22px;min-height:112px;overflow:visible;border:1px solid #e0ebfb;border-radius:20px;background:linear-gradient(115deg,#f3f8ff 0%,#fff 65%,#f4f9ff 100%);box-shadow:0 8px 24px rgba(37,99,235,.055)}
.group-list-brand{display:flex;align-items:center;gap:16px;min-width:0}
.group-list-mark{width:76px;height:76px;min-width:76px;border-radius:17px;background:linear-gradient(145deg,#347eff,#2160e8);display:flex;align-items:center;justify-content:center;color:#fff;font-size:32px;box-shadow:0 8px 20px rgba(37,99,235,.18)}
.group-list-title{display:block;font-size:clamp(22px,2.1vw,32px);font-weight:850;color:#10265c;line-height:1.28;margin:0 0 7px;padding-top:2px;overflow:visible;overflow-wrap:anywhere}
.group-list-copy{font-size:14px;color:#587098;line-height:1.45;margin:0}
.group-list-stats{display:flex;gap:12px;flex-shrink:0}
.group-list-stat{display:flex;align-items:center;gap:12px;min-width:170px;padding:12px 16px;border:1px solid #deebfc;border-radius:15px;background:rgba(255,255,255,.88)}
.group-list-stat-icon{width:54px;height:54px;min-width:54px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#e8f1ff;color:#1765ed;font-size:24px}
.group-list-stat span{display:block;font-size:12px;color:#5d7193;font-weight:650}.group-list-stat b{display:block;font-size:25px;color:#1765ed;line-height:1.2;margin-top:3px}
.group-list-tools{display:flex;align-items:end;justify-content:space-between;gap:16px;margin:0 0 16px}
.group-list-section h2{font-size:23px;color:#10265c;font-weight:850;margin:0 0 4px}.group-list-section p{font-size:14px;color:#7084a5;margin:0}
.group-tile{height:100%;min-height:232px;padding:17px 17px 14px;border:1px solid #e0eafa;border-radius:17px;background:linear-gradient(145deg,#fff 0%,#fbfdff 100%);box-shadow:0 5px 15px rgba(28,60,112,.055);margin-bottom:10px;}
div[class*="st-key-group_card_"]{border:1px solid #e0eafa;border-radius:17px;background:linear-gradient(145deg,#fff 0%,#fbfdff 100%);box-shadow:0 5px 15px rgba(28,60,112,.055);padding:16px 17px 14px;margin-bottom:10px;min-height:280px!important;height:280px!important;max-height:none!important;box-sizing:border-box;display:flex;flex-direction:column;overflow:visible;}
div[class*="st-key-group_card_"] > div{min-width:0;}
div[class*="st-key-group_card_"] [data-testid="stVerticalBlock"]{height:100%;min-height:0;display:flex!important;flex-direction:column!important;align-items:stretch!important;}
/* Let the card's content area take remaining space, then pin each action button to the same bottom baseline. */
div[class*="st-key-group_card_"] [data-testid="stElementContainer"]:has(.group-avatars){flex:1 1 auto!important;min-height:0;}
div[class*="st-key-group_card_"] [data-testid="stElementContainer"]:has(button){margin-top:auto!important;flex:0 0 auto!important;width:100%;}
div[class*="st-key-group_card_"] [data-testid="stElementContainer"]:has(button) button{width:100%;}
div[class*="st-key-group_card_"] .group-tile-title{overflow-wrap:anywhere;word-break:normal;}
div[class*="st-key-group_card_"]:hover{box-shadow:0 9px 24px rgba(28,60,112,.10)}
[data-testid="stHorizontalBlock"]:has(div[class*="st-key-group_card_"]){align-items:stretch!important;}
[data-testid="stHorizontalBlock"]:has(div[class*="st-key-group_card_"]) > [data-testid="column"]{display:flex!important;align-items:stretch!important;}
[data-testid="stHorizontalBlock"]:has(div[class*="st-key-group_card_"]) > [data-testid="column"] > div{width:100%;}
.group-tile-top{display:flex;gap:13px;align-items:flex-start;min-width:0}
.group-tile-icon{width:66px;height:66px;min-width:66px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:29px}
.group-tile-main{flex:1;min-width:0;padding-top:3px}
.group-tile-number{font-size:14px;font-weight:750;color:#536b93;margin-bottom:5px}
.group-tile-title{font-size:17px;font-weight:850;color:#071f59;line-height:1.35;overflow-wrap:anywhere;min-height:24px}
.group-tile-subtitle{font-size:13px;color:#526a91;line-height:1.35;margin-top:2px;overflow-wrap:anywhere}
.group-count-pill{display:inline-flex;align-items:center;gap:5px;background:#edf4ff;color:#1765ed;padding:6px 10px;border-radius:30px;font-size:12px;font-weight:750;white-space:nowrap;margin-left:auto}
.group-avatars{display:flex;align-items:center;padding-left:4px;margin:14px 0 15px 80px;min-height:37px}
.group-avatar-mini{width:39px;height:39px;border:2px solid #fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;margin-left:-4px}
.group-tile .stButton{margin-top:0}
.group-tile button{border-radius:10px!important;min-height:43px!important;font-weight:800!important}
.group-tile button[kind="secondary"]{border-color:#76a6ff!important;color:#1262ee!important;background:#fff!important}
.group-tile button[kind="primary"]{background:linear-gradient(100deg,#2563eb,#1d4ed8)!important;border-color:#2563eb!important}
@media(max-width:1100px){.group-list-hero{align-items:flex-start;flex-direction:column}.group-list-stats{width:100%}.group-list-stat{flex:1}div[class*="st-key-group_card_"]{height:280px!important;min-height:280px!important}.group-count-pill{font-size:11px;padding:5px 8px}}
@media(max-width:650px){.group-list-stats{flex-direction:column}.group-list-stat{width:100%}.group-list-tools{align-items:stretch;flex-direction:column}.group-list-mark{width:58px;height:58px;min-width:58px}.group-avatars{margin-left:0}div[class*="st-key-group_card_"]{height:auto!important;min-height:245px!important;max-height:none!important}}
.group-list-toolbar-label{font-size:20px;font-weight:800;color:#10265c;margin:8px 0 10px}
.group-sort-label{font-size:13px;color:#344c73;padding:8px 0 6px;text-align:right}
.group-tile-icon{width:72px;height:72px;min-width:72px;border-radius:17px;display:flex;align-items:center;justify-content:center;font-size:30px}
div[class*="st-key-group_card_"]{height:176px!important;min-height:176px!important;max-height:none!important;padding:17px 15px 12px;border-radius:16px;background:linear-gradient(145deg,#fff,#fbfdff);}
div[class*="st-key-group_card_"] [data-testid="stVerticalBlock"]{gap:0.35rem!important}
.group-tile-top{gap:12px;align-items:flex-start}
.group-tile-number{display:inline-flex;background:#eaf2ff;color:#1765ed;border-radius:18px;padding:4px 10px;font-size:12px;margin:0 0 7px}
.group-tile-title{font-size:16px;line-height:1.25;min-height:0;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.group-tile-subtitle{font-size:13px;margin-top:5px;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}
.group-count-pill{font-size:11px;padding:6px 9px;flex-shrink:0}
.group-card-divider{height:1px;background:#edf2fa;margin:12px 0 0 84px}
.group-avatars{margin:0 0 0 84px;min-height:28px;flex:1}
.group-avatar-mini{width:31px;height:31px;font-size:10px}
div[class*="st-key-group_card_"] [data-testid="stElementContainer"]:has(button){margin-top:auto!important}
div[class*="st-key-group_card_"] button{border:0!important;background:transparent!important;color:#0964f7!important;box-shadow:none!important;min-height:30px!important;height:30px!important;padding:0 4px!important;text-align:right!important;font-weight:750!important}
div[class*="st-key-group_card_"] button:hover{background:#eef5ff!important}
.group-pagination-summary{color:#7284a3;font-weight:600;padding:9px 0}
@media(max-width:900px){div[class*="st-key-group_card_"]{height:190px!important;min-height:190px!important;padding:14px!important}.group-tile-icon{width:56px;height:56px;min-width:56px}.group-card-divider,.group-avatars{margin-left:0}}
@media(max-width:650px){.group-list-hero{padding:16px!important}.group-list-title{font-size:23px}.group-list-copy{font-size:13px}.group-list-stat{min-width:0}.group-tile-title{font-size:15px}div[class*="st-key-group_card_"]{height:auto!important;min-height:180px!important}}


/* Final Review dashboard: summary metrics, assessment panels and recommendation cards. */
.final-review-hero{border:1px solid #dce9fb;border-radius:18px;padding:18px 20px 12px;background:linear-gradient(120deg,#fff 0%,#f4f8ff 100%);box-shadow:0 5px 18px rgba(28,67,130,.05);margin:0 0 14px}
.final-review-head{display:flex;align-items:center;gap:14px;margin-bottom:16px}
.final-review-icon{width:58px;height:58px;min-width:58px;border-radius:15px;background:linear-gradient(145deg,#3478ff,#2563eb);display:flex;align-items:center;justify-content:center;color:#fff;font-size:27px}
.final-review-title{font-size:clamp(23px,2.2vw,32px);font-weight:850;color:#10265c;line-height:1.15}
.final-review-subtitle{font-size:14px;color:#5b7197;margin-top:5px;line-height:1.4}
.final-review-metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.final-review-metric{display:flex;align-items:center;gap:12px;min-width:0;padding:13px 15px;border:1px solid #e2ebf8;border-radius:13px;background:rgba(246,249,255,.9)}
.final-review-metric-icon{width:48px;height:48px;min-width:48px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:800}
.final-review-metric-icon.blue{background:#e4efff;color:#2671ef}.final-review-metric-icon.green{background:#dcf9ee;color:#09976a}.final-review-metric-icon.purple{background:#eee8ff;color:#7344dc}.final-review-metric-icon.teal{background:#ddfaf4;color:#0c9b78}
.final-review-metric-label{font-size:12px;color:#536b99;line-height:1.35}.final-review-metric-value{font-size:23px;font-weight:850;color:#10265c;line-height:1.25;overflow-wrap:anywhere}.final-review-banner{display:flex;align-items:center;gap:13px;padding:13px 17px;margin:14px 0;border:1px solid #d7e7ff;border-radius:13px;background:linear-gradient(100deg,#f4f8ff,#fff)}
.final-review-banner-icon{width:48px;height:48px;min-width:48px;border-radius:13px;background:#e2efff;display:flex;align-items:center;justify-content:center;font-size:25px}
.final-review-banner-title{font-weight:850;color:#10265c;font-size:16px}.final-review-banner-copy{font-size:13px;color:#526b9b;margin-top:3px;line-height:1.45}
.final-review-panel{height:100%;border:1px solid #e0eafa;border-radius:15px;padding:16px;background:linear-gradient(180deg,#fff,#f9fbff)}.final-review-panel.weak{border-color:#f5dddd;background:linear-gradient(180deg,#fff,#fffafa)}
.final-review-panel-head{display:flex;align-items:center;gap:11px;margin-bottom:12px}.final-review-panel-icon{width:46px;height:46px;min-width:46px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:23px;background:#ddf8e9;color:#0b9861}.final-review-panel.weak .final-review-panel-icon{background:#ffe4e5;color:#e33445}
.final-review-panel-title{font-size:19px;font-weight:850;color:#10265c;line-height:1.2}.final-review-panel-subtitle{font-size:12px;color:#63789e;margin-top:3px}.final-review-count{margin-left:auto;padding:6px 10px;border-radius:30px;background:#ddf8e9;color:#07834f;font-size:12px;font-weight:800;white-space:nowrap}.final-review-panel.weak .final-review-count{background:#ffe4e5;color:#d72c3d}
.final-review-item{display:flex;align-items:center;gap:10px;padding:9px 10px;border:1px solid #e7f2ec;border-radius:10px;background:#fbfffc;margin:7px 0;min-width:0}.final-review-item.weak{border-color:#fde5e5;background:#fffdfd}.final-review-number{display:flex;align-items:center;justify-content:center;width:37px;height:37px;min-width:37px;border-radius:50%;background:#e0faeb;color:#098a51;font-weight:850}.final-review-item.weak .final-review-number{background:#ffe4e5;color:#d92e3f}.final-review-item-text{min-width:0;flex:1;color:#10265c;font-size:13px;line-height:1.4;overflow-wrap:anywhere}.final-review-item-sub{font-size:11px;color:#647aa1;margin-top:2px}.final-review-tag{padding:5px 9px;border-radius:22px;background:#e4efff;color:#1769df;font-size:11px;font-weight:750;white-space:nowrap}
.final-review-recommendation{border:1px solid #dce8fa;border-radius:15px;background:#fff;padding:17px;margin:14px 0}.final-review-recommendation-title{font-size:20px;font-weight:850;color:#10265c}.final-review-recommendation-subtitle{font-size:13px;color:#60769d;margin:4px 0 12px}
@media(max-width:1000px){.final-review-metrics{grid-template-columns:repeat(2,minmax(0,1fr))}.final-review-panel-title{font-size:16px}}
@media(max-width:560px){.final-review-metrics{grid-template-columns:1fr}.final-review-head{align-items:flex-start}.final-review-icon{width:46px;height:46px;min-width:46px}.final-review-tag{display:none}}

/* Readable, professional evaluation component tiles; all eight remain in one row. */
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button,
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button {
  min-height: 104px !important; height: 104px !important; padding: 12px 8px !important;
  border-radius: 14px !important; border: 1px solid #d7e5fa !important;
  background: linear-gradient(180deg,#ffffff 0%,#f7faff 100%) !important;
  color: #18345f !important; box-shadow: 0 3px 10px rgba(24,61,112,.07) !important;
  white-space: pre-line !important; overflow-wrap: anywhere;
}
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button p,
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button p {
  font-size: clamp(12px, .88vw, 15px) !important; font-weight: 750 !important;
  line-height: 1.45 !important; white-space: pre-line !important; margin: 0 !important;
}
[class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button[kind="primary"],
[class*="st-key-phase_nav_index_"] [data-testid="stButton"] button[kind="primary"] {
  border: 2px solid #1769f6 !important; background: linear-gradient(180deg,#eef5ff,#e4efff) !important;
  color: #075ce5 !important; box-shadow: 0 5px 14px rgba(37,99,235,.14) !important;
}
@media(max-width:850px) {
  [class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button,
  [class*="st-key-phase_nav_index_"] [data-testid="stButton"] button { min-height: 94px !important; height: 94px !important; padding: 8px 4px !important; }
  [class*="st-key-bulk_phase_index_"] [data-testid="stButton"] button p,
  [class*="st-key-phase_nav_index_"] [data-testid="stButton"] button p { font-size: 12px !important; }
}

</style>
""", unsafe_allow_html=True)
# Static login for this local faculty evaluation portal prototype.
# For production use, replace these demo credentials with hashed passwords and a proper identity provider.
LOGIN_USERS = {
    "nehal": {"password": "nehalsheth123", "display_name": "Nehal"},
    "drashti": {"password": "drashtishah123", "display_name": "Drashti"},
    "om": {"password": "ompatel123", "display_name": "Om"},
}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "current_user" not in st.session_state:
    st.session_state.current_user = ""

if not st.session_state.authenticated:
    st.markdown("""
    <style>
    [data-testid="stSidebar"] {display:none;}
    [data-testid="collapsedControl"] {display:none;}
    .block-container {max-width:1180px;padding-top:3rem;padding-bottom:3rem;}
    .login-brand-panel {border-radius:24px;padding:3rem 2.5rem;background:linear-gradient(145deg,#102d4b,#174b70 70%,#087e8b);color:#fff;min-height:470px;display:flex;flex-direction:column;justify-content:center;box-shadow:0 18px 50px rgba(12,44,75,.16);}
    .login-brand-mark {width:64px;height:64px;border-radius:18px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.25);display:flex;align-items:center;justify-content:center;font-weight:900;font-size:23px;margin-bottom:26px;color:#fff;}
    .login-brand-title {font-size:clamp(30px,3vw,44px);font-weight:850;color:#fff;line-height:1.12;letter-spacing:-.02em;}
    .login-brand-copy {font-size:16px;line-height:1.65;color:#e1effb;margin-top:15px;max-width:470px;}
    .login-brand-tag {font-size:12px;letter-spacing:.15em;text-transform:uppercase;font-weight:800;color:#a9f2ec;margin-bottom:12px;}
    .login-form-title {font-size:28px;font-weight:850;color:#132b4d;line-height:1.2;}
    .login-form-copy {font-size:14px;color:#61738b;margin:8px 0 24px;line-height:1.5;}
    .login-security-note {font-size:12px;color:#75859a;margin-top:18px;line-height:1.5;}
    [data-testid="stForm"] {border:0;padding:0;}
    [data-testid="stTextInput"] label {font-weight:700;color:#233d5b;}
    [data-testid="stTextInput"] input {min-height:46px;border-radius:10px;}
    [data-testid="stFormSubmitButton"] button {min-height:48px;border-radius:10px;font-weight:800;background:#2563eb;border-color:#2563eb;color:white;width:100%;}
    @media(max-width:800px){.block-container{padding-top:1.2rem}.login-brand-panel{min-height:auto;padding:2rem}}
    </style>
    """, unsafe_allow_html=True)
    left, right = st.columns([1.05, 0.95], gap="large")
    with left:
        st.markdown("""
        <div class="login-brand-panel">
          <div class="login-brand-mark">ML</div>
          <div class="login-brand-tag">Semester 7 · Faculty Workspace</div>
          <div class="login-brand-title">Machine Learning<br>Project Evaluation</div>
          <div class="login-brand-copy">A central workspace to evaluate student projects, record phase-wise marks, review evidence, and prepare final recommendations.</div>
        </div>
        """, unsafe_allow_html=True)
    with right:
        with st.container():
            st.markdown('<div class="login-form-title">Welcome back</div><div class="login-form-copy">Sign in with your portal credentials to continue to the evaluation dashboard.</div>', unsafe_allow_html=True)
            with st.form("portal_login_form", clear_on_submit=False):
                username = st.text_input("Username", placeholder="Enter username", autocomplete="username")
                password = st.text_input("Password", type="password", placeholder="Enter password", autocomplete="current-password")
                submitted = st.form_submit_button("Sign in", use_container_width=True)
            if submitted:
                key = username.strip().lower()
                account = LOGIN_USERS.get(key)
                if account and password == account["password"]:
                    st.session_state.authenticated = True
                    st.session_state.current_user = account["display_name"]
                    log_activity("Login", f"{account['display_name']} signed in", "Authentication")
                    st.session_state["Navigation"] = "Dashboard"
                    st.rerun()
                else:
                    st.error("Incorrect username or password. Please try again.")
            st.markdown('<div class="login-security-note">Authorized users only · Your login remains active for this browser session.</div>', unsafe_allow_html=True)
    st.stop()

with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
      <div class="sidebar-brand-mark">ML</div>
      <div><div class="sidebar-brand-title">Project Evaluation</div>
      <div class="sidebar-brand-subtitle">SEMESTER 7 · MACHINE LEARNING</div></div>
    </div>
    <div class="sidebar-section-label">Workspace</div>
    """, unsafe_allow_html=True)
    page=st.radio("Navigation",["Dashboard","Student Data Upload","Student Evaluation","Project Group Details","Records & Export"], key="Navigation", label_visibility="collapsed")
    st.markdown(f"<div style='color:#ffffff;font-size:13px;padding:14px 4px 4px'>Signed in as <b>{html.escape(st.session_state.current_user)}</b></div>", unsafe_allow_html=True)
    def logout_callback():
        # Widget callbacks run before the next rerun, so reset navigation safely
        # before the sidebar radio widget is instantiated.
        log_activity("Logout", f"{st.session_state.get('current_user', 'User')} signed out", "Authentication")
        st.session_state.authenticated = False
        st.session_state.current_user = ""
        st.session_state["Navigation"] = "Dashboard"

    st.button("Log out", key="portal_logout", use_container_width=True, on_click=logout_callback)
    previous_page = st.session_state.get("_activity_last_page")
    if previous_page != page:
        log_activity("Page visit", f"Opened {page}", "Navigation")
        st.session_state["_activity_last_page"] = page
students_df=get_students()

# Shared page heading: consistent branding across every authenticated workspace page.
# The heading is intentionally allowed to wrap instead of clipping on narrow screens.
_PAGE_HEADERS = {
    "Dashboard": ("Evaluation Dashboard", "Welcome back! Review the current status of your Machine Learning project evaluations."),
    "Student Data Upload": ("Student Data Upload", "Manage the student roster, update project titles, and move students between groups."),
    "Student Evaluation": ("Group-wise Project Evaluation", "Select a project group, review all project partners, and evaluate students across the project phases."),
    "Project Group Details": ("Project Group Details", "Review group members, project information, and group-wise evaluation progress."),
    "Records & Export": ("Records & Export", "Review saved evaluation records, recent website activity, and export your data."),
}
_header_title, _header_subtitle = _PAGE_HEADERS.get(page, (page, "Machine Learning Project Evaluation · Semester 7"))
st.markdown("""<style>
.stNumberInput input{text-align:center !important;font-weight:750 !important;font-size:16px !important;color:#152b55 !important}
.stNumberInput button{min-height:36px !important}

/* High-contrast decrement/increment controls: plain text glyphs avoid emoji rendering issues. */
div[class*="st-key-minus_"] [data-testid="stButton"] > button {
  min-width: 36px !important; min-height: 40px !important; height: 40px !important;
  padding: 0 !important; border: 1px solid #dc2626 !important; border-radius: 10px !important;
  background: #fee2e2 !important; color: #b91c1c !important; font-size: 25px !important;
  font-weight: 900 !important; line-height: 1 !important; box-shadow: none !important;
}
div[class*="st-key-minus_"] [data-testid="stButton"] > button:hover { background: #fecaca !important; color: #991b1b !important; border-color: #b91c1c !important; }
div[class*="st-key-plus_"] [data-testid="stButton"] > button {
  min-width: 36px !important; min-height: 40px !important; height: 40px !important;
  padding: 0 !important; border: 1px solid #16a34a !important; border-radius: 10px !important;
  background: #dcfce7 !important; color: #15803d !important; font-size: 25px !important;
  font-weight: 900 !important; line-height: 1 !important; box-shadow: none !important;
}
div[class*="st-key-plus_"] [data-testid="stButton"] > button:hover { background: #bbf7d0 !important; color: #166534 !important; border-color: #15803d !important; }
div[class*="st-key-minus_"] [data-testid="stButton"] > button:disabled,
div[class*="st-key-plus_"] [data-testid="stButton"] > button:disabled { opacity: .65 !important; color: #475569 !important; }
div[class*="st-key-minus_"] [data-testid="stButton"] p,
div[class*="st-key-plus_"] [data-testid="stButton"] p { font-size: 25px !important; font-weight: 900 !important; line-height: 1 !important; margin: 0 !important; }

.page-heading-card{display:flex;align-items:center;gap:16px;width:100%;box-sizing:border-box;min-width:0;overflow:visible;padding:18px 22px;margin:0 0 20px;border:1px solid #dce9fb;border-radius:20px;background:linear-gradient(115deg,#f2f7ff 0%,#ffffff 66%,#f0fbf8 100%);box-shadow:0 6px 22px rgba(35,86,160,.055)}
.page-heading-icon{display:flex;align-items:center;justify-content:center;flex:0 0 76px;width:76px;height:76px;border-radius:18px;background:linear-gradient(145deg,#3478ff,#205be8);color:#fff;font-size:34px;box-shadow:0 8px 18px rgba(37,99,235,.18)}
.page-heading-copy{flex:1 1 auto;min-width:0}
.page-heading-title{margin:0;color:#10265c;font-size:clamp(23px,2.2vw,32px);font-weight:850;line-height:1.22;letter-spacing:-.025em;overflow-wrap:anywhere;word-break:normal;white-space:normal}
.page-heading-subtitle{margin:7px 0 0;color:#526b91;font-size:15px;line-height:1.5;overflow-wrap:anywhere;white-space:normal}
@media(max-width:650px){.page-heading-card{align-items:flex-start;gap:12px;padding:14px;margin-bottom:15px;border-radius:15px}.page-heading-icon{flex-basis:52px;width:52px;height:52px;border-radius:13px;font-size:25px}.page-heading-title{font-size:22px}.page-heading-subtitle{font-size:13px;line-height:1.45}}
</style>""", unsafe_allow_html=True)
st.markdown(
    "<section class='page-heading-card' aria-label='Page heading'>"
    "<div class='page-heading-icon'>♟</div><div class='page-heading-copy'>"
    f"<h1 class='page-heading-title'>{html.escape(_header_title)}</h1>"
    f"<p class='page-heading-subtitle'>{html.escape(_header_subtitle)}</p>"
    "</div></section>",
    unsafe_allow_html=True,
)

if page=="Dashboard":
    # Dashboard-only visual theme: clean white cards, compact typography, and a navy/teal accent.
    st.markdown("""
    <style>
    .block-container {padding-top: 1.35rem; padding-bottom: 2rem; max-width: 1500px;}
    [data-testid="stSidebar"] {background: #f7f9fc; border-right: 1px solid #e7ebf2;}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {font-size: .93rem;}
    .dash-hero {display:flex; justify-content:space-between; align-items:center; gap:1rem; padding: 1.15rem 1.35rem; border:1px solid #e7ebf2; border-radius:18px; background:linear-gradient(115deg,#ffffff 0%,#f5f9ff 70%,#effcf9 100%); margin-bottom:1.15rem;}
    .dash-kicker {font-size:.82rem; text-transform:uppercase; letter-spacing:.12em; color:#0f8b8d; font-weight:800; margin-bottom:.35rem;}
    .dash-title {font-size:2rem; font-weight:800; color:#17233b; line-height:1.15; margin:0;}
    .dash-subtitle {font-size:1rem; color:#68758a; margin-top:.45rem;}
    .dash-pill {display:inline-block; background:#e6f8f5; color:#087f78; border:1px solid #c6eee7; padding:.42rem .7rem; border-radius:999px; font-size:.85rem; font-weight:700; white-space:nowrap;}
    .stat-card {background:#fff; border:1px solid #e8edf4; border-radius:15px; padding:1rem .95rem .9rem; box-shadow:0 3px 12px rgba(27,45,74,.045); min-height:132px; position:relative; overflow:hidden; margin-bottom:.8rem;}
    .stat-card:after {content:''; position:absolute; right:-18px; top:-20px; width:75px; height:75px; border-radius:50%; background:var(--tint); opacity:.55;}
    .stat-top {display:flex; align-items:center; gap:.6rem; color:#52627a; font-size:.92rem; font-weight:650; position:relative; z-index:1;}
    .stat-icon {display:inline-flex; align-items:center; justify-content:center; width:35px; height:35px; border-radius:10px; background:var(--tint); color:var(--accent); font-size:1.12rem;}
    .stat-value {font-size:2.1rem; font-weight:800; color:#17233b; line-height:1.3; margin-top:.55rem; position:relative; z-index:1;}
    .stat-note {font-size:.84rem; color:#66758b; margin-top:.15rem; position:relative; z-index:1;}
    .panel {background:#fff; border:1px solid #e8edf4; border-radius:16px; padding:1rem 1.1rem; box-shadow:0 3px 12px rgba(27,45,74,.035); height:100%;}
    .panel-title {font-size:1.2rem; font-weight:800; color:#17233b; margin-bottom:.15rem;}
    .panel-subtitle {font-size:.9rem; color:#66758b; margin-bottom:.75rem;}
    .activity-row {display:flex; gap:.65rem; padding:.65rem 0; border-bottom:1px solid #f0f2f6; align-items:flex-start;}
    .activity-row:last-child {border-bottom:none;}
    .activity-icon {width:31px; height:31px; display:flex; align-items:center; justify-content:center; border-radius:10px; background:#edf4ff; color:#2563eb; flex-shrink:0;}
    .activity-main {font-size:.92rem; color:#334155; font-weight:650;}
    .activity-detail {font-size:.82rem; color:#8a95a6; margin-top:.15rem;}
    .quick-action {background:#fff; border:1px solid #e8edf4; border-radius:12px; padding:.8rem .75rem; min-height:76px;}
    .quick-action-title {font-size:.95rem; font-weight:750; color:#24324a;}
    .quick-action-sub {font-size:.82rem; color:#8a95a6; margin-top:.2rem;}
    div[data-testid="stButton"] > button {border-radius:10px; border-color:#dce4ee; font-weight:650;}
    div[data-testid="stButton"] > button[kind="primary"] {background:#087f8c; border-color:#087f8c;}
    </style>
    """, unsafe_allow_html=True)

    total_students = len(students_df)
    evaluated_count = pending_count = 0
    total_marks = 0.0
    component_totals = {component["name"]: 0.0 for component in COMPONENTS}
    evaluations = []
    for _, student in students_df.iterrows():
        evaluation = get_evaluation(str(student["roll_no"]))
        score = marks_total(evaluation)
        total_marks += score
        evaluated_count += int(score > 0)
        pending_count += int(score <= 0)
        evaluations.append((student, evaluation, score))
        for component in COMPONENTS:
            component_totals[component["name"]] += sum(float(v) for v in evaluation["scores"][component["name"]].values())
    average_marks = total_marks / total_students if total_students else 0
    evaluation_rate = evaluated_count / total_students * 100 if total_students else 0
    project_titles_count = int(students_df["ml_project_title"].fillna("").astype(str).str.strip().ne("").sum()) if total_students else 0
    submission_complete_count = sum(all(ev.get("submission_checks", {}).get(item, False) for item in SUBMISSION_ITEMS) for _, ev, _ in evaluations)

    st.markdown(f'''<div class="dash-hero"><div><div class="dash-kicker">Semester 7 · Faculty workspace</div><div class="dash-title">Evaluation Dashboard</div><div class="dash-subtitle">Welcome back! Here’s the current status of your Machine Learning project evaluations.</div></div><div class="dash-pill">● Live data · SQLite connected</div></div>''', unsafe_allow_html=True)

    def stat_card(icon, label, value, note, accent, tint):
        st.markdown(f'''<div class="stat-card" style="--accent:{accent};--tint:{tint};"><div class="stat-top"><span class="stat-icon">{icon}</span><span>{label}</span></div><div class="stat-value">{value}</div><div class="stat-note">{note}</div></div>''', unsafe_allow_html=True)

    # Detailed insight tiles replace the old component bar chart.
    group_count = int(students_df["group_no"].fillna("").astype(str).str.strip().replace("nan", "").replace("None", "").ne("").sum()) if total_students else 0
    unique_groups = int(students_df["group_no"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if total_students else 0
    missing_titles = max(total_students - project_titles_count, 0)
    submission_pending = max(total_students - submission_complete_count, 0)
    scores_list = [score for _, _, score in evaluations]
    highest_score = max(scores_list, default=0)
    lowest_positive = min((score for score in scores_list if score > 0), default=0)
    full_marks_count = sum(score >= TOTAL_MARKS for score in scores_list)
    partial_count = sum(0 < score < TOTAL_MARKS for score in scores_list)
    no_marks_count = sum(score == 0 for score in scores_list)
    students_with_dataset = sum(bool(str(ev.get("dataset_source", "")).strip()) for _, ev, _ in evaluations)
    checklist_items_total = len(SUBMISSION_ITEMS) * total_students
    checklist_items_done = sum(sum(bool(ev.get("submission_checks", {}).get(item, False)) for item in SUBMISSION_ITEMS) for _, ev, _ in evaluations)
    checklist_rate = checklist_items_done / checklist_items_total * 100 if checklist_items_total else 0
    average_component = {}
    for component in COMPONENTS:
        max_component = component["max"]
        average_component[component["name"]] = (component_totals[component["name"]] / total_students) if total_students else 0

    st.markdown('<div class="panel-title">Evaluation Insights</div><div class="panel-subtitle">Key indicators calculated from the student roster and saved individual evaluations</div>', unsafe_allow_html=True)
    insights = [
        ("♙", "Total Students", f"{total_students}", "Students in roster", "#2563eb", "#eaf1ff"),
        ("✓", "Students Evaluated", f"{evaluated_count}", "At least one mark recorded", "#0f9f78", "#e5f8f1"),
        ("◷", "Evaluation Pending", f"{pending_count}", "No marks recorded yet", "#d97706", "#fff4df"),
        ("%", "Evaluation Coverage", f"{evaluation_rate:.1f}%", "Students with marks", "#7c3aed", "#f2eaff"),
        ("▦", "Project Groups", f"{unique_groups}", "Unique group numbers", "#0891b2", "#e4f8fb"),
        ("▤", "Project Titles", f"{project_titles_count}", "Titles recorded", "#4f46e5", "#ecebff"),
        ("!", "Titles Missing", f"{missing_titles}", "Students needing a title", "#dc2626", "#fff0f0"),
        ("★", "Class Average", f"{average_marks:.1f}/{TOTAL_MARKS}", "Average marks per student", "#9333ea", "#f4eaff"),
        ("↑", "Highest Total", f"{highest_score:.0f}/{TOTAL_MARKS}", "Highest current student score", "#15803d", "#e8f8ed"),
        ("↓", "Lowest Positive", f"{lowest_positive:.0f}" if lowest_positive else "—", "Lowest score above zero", "#c2410c", "#fff1e8"),
        ("◎", "Full Marks", f"{full_marks_count}", "Students at maximum score", "#0f766e", "#e1f7f3"),
        ("◐", "Partially Evaluated", f"{partial_count}", "Some marks recorded", "#b45309", "#fff6df"),
        ("○", "Not Started", f"{no_marks_count}", "Students with zero marks", "#64748b", "#eef2f7"),
        ("✓", "Complete Submissions", f"{submission_complete_count}", "All final checklist items done", "#047857", "#e3f8ee"),
        ("…", "Submission Pending", f"{submission_pending}", "Checklist not fully complete", "#be123c", "#fff0f4"),
        ("☷", "Checklist Completion", f"{checklist_rate:.1f}%", "All student checklist items", "#0369a1", "#e8f5ff"),
        ("⌁", "Dataset Sources", f"{students_with_dataset}", "Students with source recorded", "#6d28d9", "#f2eaff"),
        ("Σ", "Marks Awarded", f"{total_marks:.0f}", "Combined marks across class", "#334155", "#edf2f7"),
    ]
    tile_cols = st.columns(4, gap="medium")
    for idx, (icon, label, value, note, accent, tint) in enumerate(insights):
        with tile_cols[idx % 4]:
            stat_card(icon, label, value, note, accent, tint)
    st.write("")

    st.markdown('<div class="panel-title">Recent Activity Across the Website</div><div class="panel-subtitle">Latest page visits, roster changes, group management, evaluations and exports recorded in the activity log</div>', unsafe_allow_html=True)
    recent_activity = get_recent_activity(12)
    if not recent_activity:
        st.info("No activity has been recorded yet. Navigation and saved operations will appear here automatically.")
    else:
        activity_icons = {"Page visit": "↗", "Student import": "⇧", "Project title updated": "✎", "Student moved": "⇄", "Evaluation saved": "✓", "Export": "⇩", "Group opened": "▦"}
        activity_cols = st.columns(2, gap="medium")
        for idx, item in enumerate(recent_activity):
            occurred = str(item["occurred_at"] or "")
            try:
                occurred_display = datetime.fromisoformat(occurred).astimezone().strftime("%d %b %Y · %I:%M %p")
            except Exception:
                occurred_display = occurred
            with activity_cols[idx % 2]:
                icon = activity_icons.get(item["activity_type"], "•")
                description = html.escape(str(item["description"] or "Activity"))
                details = html.escape(str(item["details"] or ""))
                st.markdown(
                    f'<div class="panel" style="padding:.85rem 1rem;margin-bottom:.65rem;display:flex;gap:.75rem;align-items:flex-start">'
                    f'<div class="activity-icon">{icon}</div><div style="min-width:0;flex:1">'
                    f'<div class="activity-main">{description}</div>'
                    f'<div class="activity-detail">{html.escape(occurred_display)} · {html.escape(str(item["activity_type"]))}</div>'
                    + (f'<div class="activity-detail">{details}</div>' if details else '')
                    + '</div></div>', unsafe_allow_html=True
                )

    st.write("")
    st.markdown('<div class="panel-title">Quick Actions</div><div class="panel-subtitle">Jump directly to common faculty tasks</div>', unsafe_allow_html=True)
    quick = st.columns(3, gap="medium")
    with quick[0]:
        st.markdown('<div class="quick-action"><div class="quick-action-title">＋ Student Data Upload</div><div class="quick-action-sub">Add or update the class roster</div></div>', unsafe_allow_html=True)
        st.button("Open Student Upload →", key="dash_upload", use_container_width=True, on_click=lambda: st.session_state.update(Navigation="Student Data Upload"))
    with quick[1]:
        st.markdown('<div class="quick-action"><div class="quick-action-title">▦ Group Evaluation</div><div class="quick-action-sub">Choose a project group and evaluate</div></div>', unsafe_allow_html=True)
        st.button("Open Evaluations →", key="dash_evaluation", use_container_width=True, type="primary", on_click=lambda: st.session_state.update(Navigation="Student Evaluation"))
    with quick[2]:
        st.markdown('<div class="quick-action"><div class="quick-action-title">⇩ Records & Export</div><div class="quick-action-sub">Review records and download reports</div></div>', unsafe_allow_html=True)
        st.button("View Records →", key="dash_records", use_container_width=True, on_click=lambda: st.session_state.update(Navigation="Records & Export"))

    st.caption("Dashboard figures are calculated from the current student roster and saved individual evaluations. A student is counted as evaluated when at least one mark has been recorded.")
elif page=="Student Data Upload":
    st.subheader("Student Data & Project Group Management")
    st.caption("View the complete student roster, edit a project title for an entire group, or move a student between groups. Every assigned group must contain 1–3 students.")

    # Import/add students section
    with st.expander("＋ Upload or add student records", expanded=False):
        st.info(f"{len(students_df)} students are currently stored. Only roll number, student name, group number and ML project title are managed here.")
        st.write("Upload CSV/XLSX to add or update records. Required column: `student_name`. Optional: `roll_no`, `group_no`, `ml_project_title`. ASP.NET fields are ignored.")
        template=pd.DataFrame([{"roll_no":"58","student_name":"STUDENT FULL NAME","group_no":"23","ml_project_title":"Machine Learning Project Title"}])
        st.download_button("Download upload template CSV",template.to_csv(index=False).encode("utf-8-sig"),"ml_student_upload_template.csv","text/csv",key="student_upload_template")
        uploaded=st.file_uploader("Upload additional student data (CSV or Excel)",type=["csv","xlsx"],key="student_upload_file")
        if uploaded:
            try:
                df=pd.read_excel(uploaded,dtype=str).fillna("") if uploaded.name.lower().endswith(".xlsx") else pd.read_csv(uploaded,dtype=str).fillna("")
                df.columns=[str(c).strip().lower().replace(" ","_") for c in df.columns]
                st.dataframe(df.head(15),use_container_width=True,hide_index=True)
                if "student_name" not in df.columns:
                    st.error("A student_name column is required.")
                elif st.button("Import / update student records",type="primary",key="import_student_records"):
                    clean=[]; skipped=0
                    for _,r in df.iterrows():
                        name=str(r.get("student_name"," ")).strip()
                        roll=str(r.get("roll_no","")).strip()
                        group=str(r.get("group_no","")).strip()
                        title=str(r.get("ml_project_title","")).strip()
                        if not name or not roll:
                            skipped+=1; continue
                        clean.append((roll,name,int(group) if group.isdigit() else None,title))
                    # Validate the final group sizes before writing any imported changes.
                    with db() as con:
                        existing_rows=con.execute("SELECT roll_no,group_no FROM students").fetchall()
                        final={row[0]:row[1] for row in existing_rows}
                        existing_groups={row[1] for row in existing_rows if row[1] is not None}
                        for roll,name,group,title in clean: final[roll]=group
                        counts={}
                        for group in final.values():
                            if group is not None: counts[group]=counts.get(group,0)+1
                        oversized={g:n for g,n in counts.items() if n>3}
                        emptied=[g for g in existing_groups if counts.get(g,0)<1]
                        if oversized or emptied:
                            if oversized:
                                st.error("Import cancelled: these groups would exceed the maximum of 3 students: " + ", ".join(f"Group {g} ({n})" for g,n in sorted(oversized.items())))
                            if emptied:
                                st.error("Import cancelled: these existing groups would have no students left: " + ", ".join(f"Group {g}" for g in sorted(emptied)))
                        else:
                            added=updated=0
                            for roll,name,group,title in clean:
                                exists=con.execute("SELECT 1 FROM students WHERE roll_no=?",(roll,)).fetchone()
                                con.execute("""INSERT INTO students(roll_no,student_name,group_no,ml_project_title) VALUES(?,?,?,?)
                                    ON CONFLICT(roll_no) DO UPDATE SET student_name=excluded.student_name,
                                    group_no=excluded.group_no, ml_project_title=excluded.ml_project_title""",(roll,name,group,title))
                                con.execute("INSERT OR IGNORE INTO evaluations VALUES (?,?,?)",(roll,json.dumps(default_evaluation()),date.today().isoformat()))
                                if exists: updated+=1
                                else: added+=1
                            log_activity("Student import", "Imported / updated student records", f"File: {uploaded.name} · Added: {added} · Updated: {updated} · Skipped: {skipped}")
                            st.success(f"Import finished: {added} added, {updated} updated, {skipped} skipped.")
                            st.rerun()
            except Exception as e:
                st.error(f"Could not import file: {e}")

    # Reload the roster after any edits; show every stored student and detail in a searchable table.
    roster=get_students().copy()
    if roster.empty:
        st.warning("No students found. Upload a roster to begin.")
    else:
        roster["group_sort"] = pd.to_numeric(roster["group_no"], errors="coerce")
        roster = roster.sort_values(["group_sort", "roll_no"], na_position="last")
        st.markdown("### Complete student roster")
        st.caption(f"{len(roster)} student records · sorted by group and roll number")
        show_roster=roster[["roll_no","student_name","group_no","ml_project_title","updated_at"]].copy()
        show_roster.columns=["Roll No.","Student Name","Group No.","Project Title","Last Evaluation Update"]
        st.dataframe(show_roster,use_container_width=True,hide_index=True,height=min(520, max(180, 38*(len(show_roster)+1))))

        # Group-wide project title editor: save the title consistently to every partner in that group.
        st.markdown("### Edit project titles")
        st.caption("A group's project title is shared by all partners. Saving here updates the title for every student in that group.")
        groups=roster[roster["group_no"].notna()].groupby("group_no",sort=True)
        for group_no, members in groups:
            members=members.copy()
            existing_titles=[str(v).strip() for v in members["ml_project_title"].tolist() if str(v).strip() and str(v).lower()!="nan"]
            current_title=existing_titles[0] if existing_titles else ""
            with st.expander(f"Group {int(group_no) if str(group_no).replace('.','',1).isdigit() else group_no} · {len(members)} student(s) · {current_title or 'Project title not set'}", expanded=False):
                title_key=f"group_title_edit_{group_no}"
                edited_title=st.text_input("Project title",value=current_title,key=title_key,placeholder="Enter the group's ML project title")
                if st.button("Save project title for this group",key=f"save_group_title_{group_no}",type="primary"):
                    with db() as con:
                        con.execute("UPDATE students SET ml_project_title=? WHERE group_no=?",(edited_title.strip(),group_no))
                    log_activity("Project title updated", f"Updated project title for Group {group_no}", edited_title.strip())
                    st.success(f"Project title updated for Group {group_no}.")
                    st.rerun()

        # Move students safely between existing groups while preserving min 1 / max 3 membership.
        st.markdown("### Move a student to another group")
        st.caption("A move is allowed only when the destination has fewer than 3 students and the source group will retain at least 1 student. The destination group's title will be applied to the moved student.")
        current_groups=sorted([int(g) for g in roster["group_no"].dropna().unique() if str(g).replace('.','',1).isdigit()])
        group_counts={g:int((roster["group_no"]==g).sum()) for g in current_groups}
        moveable=roster[roster["group_no"].notna()]
        for _, student in moveable.iterrows():
            roll=str(student["roll_no"]); source=int(student["group_no"])
            dest_options=[g for g in current_groups if g!=source and group_counts.get(g,0)<3]
            with st.container(border=True):
                cols=st.columns([2.2,1,2.2,1])
                cols[0].markdown(f"**{student['student_name']}**  ")
                cols[0].caption(f"Roll {roll} · Group {source}")
                if not dest_options or group_counts.get(source,0)<=1:
                    cols[1].caption("Move unavailable")
                    cols[2].caption("Source group must retain at least 1 student; destination must have room.")
                    cols[3].button("Move",key=f"move_disabled_{roll}",disabled=True,use_container_width=True)
                else:
                    target=cols[1].selectbox("Destination",dest_options,key=f"move_target_{roll}",label_visibility="collapsed",format_func=lambda g:f"Group {g} ({group_counts.get(g,0)}/3)")
                    cols[2].caption(f"Project: {student['ml_project_title'] or 'Not set'}")
                    if cols[3].button("Move student",key=f"move_student_{roll}",use_container_width=True):
                        if group_counts.get(target,0)>=3:
                            st.error(f"Group {target} is already at the maximum of 3 students.")
                        elif group_counts.get(source,0)<=1:
                            st.error(f"Group {source} must retain at least 1 student.")
                        else:
                            with db() as con:
                                title_row=con.execute("SELECT ml_project_title FROM students WHERE group_no=? AND roll_no<>? AND TRIM(ml_project_title)<>'' ORDER BY roll_no LIMIT 1",(target,roll)).fetchone()
                                target_title=title_row[0] if title_row else ""
                                con.execute("UPDATE students SET group_no=?, ml_project_title=? WHERE roll_no=?",(target,target_title,roll))
                            log_activity("Student moved", f"Moved {student['student_name']} (Roll {roll}) from Group {source} to Group {target}", f"Project title: {target_title or 'Not set'}")
                            st.success(f"Moved {student['student_name']} to Group {target}.")
                            st.rerun()

        st.info("Group size rule: each assigned group must contain between 1 and 3 students. Student evaluation records and marks remain linked to the student's roll number when they move groups.")

elif page=="Student Evaluation":
    st.markdown("""<style>
    div[class*="st-key-group_card_"] {height:258px!important;min-height:258px!important;max-height:none!important;padding:15px 16px 11px!important;border-radius:19px!important;border:1px solid var(--tile-border,#dce9ff)!important;background:linear-gradient(145deg,var(--tile-bg,#ffffff) 0%,#ffffff 100%)!important;box-shadow:0 7px 22px rgba(30,64,120,.055)!important;box-sizing:border-box;display:flex;flex-direction:column;overflow:visible}
    div[class*="st-key-group_card_"] [data-testid="stVerticalBlock"]{height:100%;gap:.35rem!important;display:flex!important;flex-direction:column!important}
    div[class*="st-key-group_card_"] [data-testid="stElementContainer"]:has(button){margin-top:auto!important;flex:0 0 auto!important;width:100%}
    .project-card-badges{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:5px}
    .project-card-group{display:inline-flex;align-items:center;padding:6px 12px;border-radius:24px;background:var(--tile-pill,#e8f1ff);color:var(--tile-accent,#1765ed);font-size:13px;font-weight:800;white-space:nowrap}
    .project-card-count{display:inline-flex;align-items:center;gap:5px;padding:6px 10px;border-radius:24px;background:#eaf2ff;color:#1263ed;font-size:12px;font-weight:800;white-space:nowrap}
    .project-card-body{display:grid;grid-template-columns:minmax(95px,42%) minmax(0,1fr);align-items:center;gap:14px;min-width:0;flex:1;min-height:0}
    .project-card-image{width:100%;height:118px;object-fit:contain;border-radius:15px;display:block;background:var(--tile-pill,#f3f7ff)}
    .project-card-copy{min-width:0;overflow-wrap:anywhere}
    .project-card-title{font-size:17px;line-height:1.22;font-weight:850;color:#081f5b;overflow-wrap:anywhere;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
    .project-card-subtitle{font-size:13px;line-height:1.42;color:#526b91;margin-top:7px;overflow-wrap:anywhere;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
    div[class*="st-key-group_card_"] .stButton button{display:flex!important;align-items:center!important;justify-content:space-between!important;width:100%!important;min-height:35px!important;height:35px!important;border:0!important;border-top:1px solid #edf2fa!important;border-radius:0!important;background:transparent!important;color:var(--tile-accent,#0964f7)!important;box-shadow:none!important;padding:7px 3px 0!important;font-weight:800!important;text-align:left!important}
    div[class*="st-key-group_card_"] .stButton button::after{content:'→';display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;min-width:32px;border-radius:50%;background:var(--tile-pill,#e8f1ff);font-size:23px;line-height:1;color:var(--tile-accent,#0964f7);margin-left:10px}
    div[class*="st-key-group_card_"] .stButton button:hover{background:transparent!important;opacity:.82}
    @media(max-width:1100px){div[class*="st-key-group_card_"]{height:258px!important;min-height:258px!important}.project-card-image{height:112px}}
    @media(max-width:650px){div[class*="st-key-group_card_"]{height:auto!important;min-height:224px!important;padding:13px!important}.project-card-body{grid-template-columns:38% minmax(0,1fr);gap:10px}.project-card-image{height:104px}.project-card-title{font-size:15px}.project-card-subtitle{font-size:12px}.project-card-count{font-size:10px;padding:5px 7px}}
    </style>""", unsafe_allow_html=True)
    if students_df.empty:
        st.warning("No students found. Add them under Student Data Upload.")
    else:
        group_df = students_df.copy()
        group_df["group_sort"] = pd.to_numeric(group_df["group_no"], errors="coerce")
        group_df = group_df.sort_values(["group_sort", "roll_no"], na_position="last")
        grouped = list(group_df.groupby("group_no", dropna=False, sort=False))

        if "selected_evaluation_group" not in st.session_state:
            st.session_state["selected_evaluation_group"] = None

        # Header and live class-level group statistics.
        group_count = len(grouped)
        student_count = len(group_df)
        st.markdown(
            "<div class='group-list-hero'>"
            "<div class='group-list-brand'><div class='group-list-mark'>♟</div><div>"
            "<div class='group-list-title'>Group-wise Project Evaluation</div>"
            "<p class='group-list-copy'>Select a project group first. Review all project partners, then evaluate each student individually across the project phases.</p>"
            "</div></div><div class='group-list-stats'>"
            f"<div class='group-list-stat'><div class='group-list-stat-icon'>♧</div><div><span>Total Groups</span><b>{group_count}</b></div></div>"
            f"<div class='group-list-stat'><div class='group-list-stat-icon'>♟</div><div><span>Total Students</span><b>{student_count}</b></div></div>"
            "</div></div>", unsafe_allow_html=True
        )

        # Compact search/filter/sort toolbar inspired by the supplied reference.
        st.markdown("<div class='group-list-toolbar-label'>Project groups</div>", unsafe_allow_html=True)
        search_col, filter_col, sort_label_col, sort_col = st.columns([2.2, 1.05, .35, .9], gap="small", vertical_alignment="bottom")
        with search_col:
            search_text = st.text_input("Search groups", placeholder="⌕  Search by group name or project title...", key="group_search_text", label_visibility="collapsed")
        with filter_col:
            group_filter = st.selectbox("Filter groups", ["All Groups", "Title needs confirmation", "Has evaluation marks", "No marks recorded"], key="group_filter_choice", label_visibility="collapsed")
        with sort_label_col:
            st.markdown("<div class='group-sort-label'>Sort by</div>", unsafe_allow_html=True)
        with sort_col:
            sort_choice = st.selectbox("Sort by", ["Group Number", "Project Title", "Student Count"], key="group_sort_choice", label_visibility="collapsed")

        group_cards = []
        palette = [("#e7f0ff", "#1765ed", "#dce9ff"), ("#fbe7f7", "#b91455", "#f6dff0"), ("#e3f8ec", "#078b7c", "#d8f2e8"), ("#f1e9ff", "#793de2", "#e8ddff"), ("#fff2d9", "#c77700", "#fbe8cb"), ("#ffe8ec", "#d91e48", "#f9dce5"), ("#def8f7", "#079ca8", "#d3f4f5"), ("#efe8ff", "#7438e8", "#e6ddff"), ("#e3f8ec", "#168d51", "#d9f3e3")]
        # One-sentence project descriptions shown beneath every project title.
        # Group-specific entries take priority so the copy stays aligned with the roster's actual titles.
        subtitle_map = {
            "travelvista (booking management system)": "Helps users plan trips by organizing travel bookings and destination details.",
            "travelvista": "Helps users plan trips by organizing travel bookings and destination details.",
            "pricemyphone": "Predicts a mobile phone's price from its specifications and market features.",
            "resume screening system": "Screens resumes and ranks candidates against job requirements using machine learning.",
            "smart demand products prediction": "Forecasts product demand from historical sales and market patterns.",
            "job seekers prediction": "Analyzes candidate and job-related data to estimate suitable job opportunities.",
            "food delivery time prediction": "Estimates delivery time using order, distance, and delivery conditions.",
            "hotel booking cancellation": "Predicts whether a hotel reservation is likely to be cancelled.",
            "house rent prediction": "Estimates rental prices from property features and location details.",
            "house price prediction": "Estimates property prices using location, size, and housing attributes.",
            "email phishing detection": "Identifies suspicious emails that may contain phishing attempts.",
            "hostel energysense": "Analyzes hostel energy usage to understand and forecast consumption patterns.",
            "flight price prediction": "Estimates flight fares using route, timing, and other travel features.",
            "telco churn prediction system": "Predicts which telecom customers may stop using the service.",
            "cancer awareness and prediction system": "Presents cancer awareness information and supports risk prediction from relevant data.",
            "examination preparation & performance analyzer": "Tracks preparation and analyzes performance to highlight areas for improvement.",
            "student stress level classification and psychological factor": "Classifies student stress levels using relevant student and lifestyle factors.",
            "prediction": "Uses machine learning data patterns to generate predictions for the selected project domain.",
            "trip buddy": "Supports travel planning with trip ideas, destination details, and recommendations.",
            "ipl match prediction": "Predicts IPL match outcomes using team, player, and match statistics.",
            "ipl auction predictor": "Estimates player value and supports IPL auction decisions using historical statistics.",
            "project title to be confirmed": "A general machine learning prediction project that learns patterns from data.",
            "machine learning project": "A machine learning system that learns from data to make useful predictions.",
        }
        for idx, (group_no, members) in enumerate(grouped):
            title_values = [str(t).strip() for t in members["ml_project_title"].tolist() if str(t).strip() and str(t).lower() != "nan"]
            unique_titles = list(dict.fromkeys(title_values))
            project_title = unique_titles[0] if unique_titles else "Project title to be confirmed"
            if len(unique_titles) > 1:
                project_title = "Multiple project titles — please verify"
            label_group = "Not assigned" if pd.isna(group_no) or str(group_no).strip() == "" else str(group_no)
            search_blob = f"{label_group} {project_title} " + " ".join(members["student_name"].astype(str).tolist())
            member_marks = [marks_total(get_evaluation(str(r["roll_no"]))) for _, r in members.iterrows()]
            has_marks = any(mark > 0 for mark in member_marks)
            needs_title = (not title_values) or "confirm" in project_title.lower() or len(unique_titles) > 1
            if search_text and search_text.strip().lower() not in search_blob.lower():
                continue
            if group_filter == "Title needs confirmation" and not needs_title:
                continue
            if group_filter == "Has evaluation marks" and not has_marks:
                continue
            if group_filter == "No marks recorded" and has_marks:
                continue
            bg, fg, icon = palette[idx % len(palette)]
            image_uri = project_image_data(group_no)
            subtitle = subtitle_map.get(project_title.lower().strip(), "Uses machine learning to analyze project data and produce a useful prediction or insight.")
            group_cards.append({"idx":idx,"group":label_group,"title":project_title,"members":members,"bg":bg,"fg":fg,"border":palette[idx % len(palette)][2],"image":image_uri,"subtitle":subtitle})

        if sort_choice == "Project Title":
            group_cards.sort(key=lambda g: g["title"].lower())
        elif sort_choice == "Student Count":
            group_cards.sort(key=lambda g: (-len(g["members"]), str(g["group"])))
        else:
            group_cards.sort(key=lambda g: (999999 if not str(g["group"]).isdigit() else int(g["group"]), str(g["group"])))

        page_size = 9
        page_count = max(1, (len(group_cards) + page_size - 1) // page_size)
        if st.session_state.get("group_cards_page", 1) > page_count:
            st.session_state["group_cards_page"] = 1
        page_number = st.session_state.get("group_cards_page", 1)
        if not group_cards:
            st.info("No project groups match the current search or filter.")
        else:
            start_idx = (page_number - 1) * page_size
            visible_cards = group_cards[start_idx:start_idx + page_size]
            for row_start in range(0, len(visible_cards), 3):
                tile_cols = st.columns(3, gap="medium")
                for col_idx, card in enumerate(visible_cards[row_start:row_start + 3]):
                    with tile_cols[col_idx]:
                        label_group = card["group"]
                        with st.container(key=f"group_card_{label_group}"):
                            st.markdown(
                                f"<style>div[class*=\"st-key-group_card_{html.escape(label_group)}\"]{{--tile-accent:{card['fg']};--tile-pill:{card['bg']};--tile-border:{card['border']};background:linear-gradient(145deg,{card['bg']}55,#ffffff 82%)!important;border-color:{card['border']}!important}} div[class*=\"st-key-group_card_{html.escape(label_group)}\"] .stButton button::after{{--tile-accent:{card['fg']};--tile-pill:{card['bg']}}}</style>"
                                f"<div class='project-card-badges' style='--tile-accent:{card['fg']};--tile-pill:{card['bg']};--tile-border:{card['border']}'>"
                                f"<span class='project-card-group'>Group {html.escape(label_group)}</span>"
                                f"<span class='project-card-count'>♟ {len(card['members'])} Students</span></div>"
                                f"<div class='project-card-body' style='--tile-accent:{card['fg']};--tile-pill:{card['bg']};--tile-border:{card['border']}'>"
                                f"<img class='project-card-image' src='{card['image']}' alt='{html.escape(card['title'])} project illustration'>"
                                f"<div class='project-card-copy'><div class='project-card-title'>{html.escape(card['title'])}</div>"
                                f"<div class='project-card-subtitle'>{html.escape(card['subtitle'])}</div></div></div>", unsafe_allow_html=True
                            )
                            st.button("View Group", key=f"open_group_{label_group}", use_container_width=True,
                                      type="primary" if st.session_state["selected_evaluation_group"] == str(label_group) else "secondary",
                                      on_click=open_project_group, args=(label_group,))
            range_start = start_idx + 1
            range_end = min(start_idx + page_size, len(group_cards))
            page_left, page_right = st.columns([1, 1], vertical_alignment="center")
            with page_left:
                st.markdown(f"<div class='group-pagination-summary'>Showing {range_start} – {range_end} of {len(group_cards)} groups</div>", unsafe_allow_html=True)
            with page_right:
                nav_cols = st.columns([1, 1, 1, 1, 1], gap="small")
                with nav_cols[0]:
                    if st.button("‹", key="group_page_prev", disabled=page_number <= 1, use_container_width=True):
                        st.session_state["group_cards_page"] = page_number - 1
                        st.rerun()
                for pnum in range(1, min(page_count, 3) + 1):
                    with nav_cols[pnum]:
                        if st.button(str(pnum), key=f"group_page_{pnum}", type="primary" if pnum == page_number else "secondary", use_container_width=True):
                            st.session_state["group_cards_page"] = pnum
                            st.rerun()
                with nav_cols[4]:
                    if st.button("›", key="group_page_next", disabled=page_number >= page_count, use_container_width=True):
                        st.session_state["group_cards_page"] = page_number + 1
                        st.rerun()

elif page=="Project Group Details":
    if students_df.empty:
        st.warning("No students found. Add them under Student Data Upload.")
    else:
        group_df = students_df.copy()
        group_df["group_sort"] = pd.to_numeric(group_df["group_no"], errors="coerce")
        group_df = group_df.sort_values(["group_sort", "roll_no"], na_position="last")
        selected_group = st.session_state.get("selected_evaluation_group")
        st.markdown("<div class='group-breadcrumb'>⌂ &nbsp; <span>Students</span> &nbsp; / &nbsp; <strong>Group " + html.escape(str(selected_group or '—')) + "</strong></div>", unsafe_allow_html=True)
        st.button("← Back to all project groups", key="back_to_groups", on_click=back_to_groups)
        if selected_group is None:
            st.info("Choose a group from Student Evaluation to open its project details.")
        else:
            if selected_group == "Not assigned":
                selected_members = group_df[group_df["group_no"].isna() | (group_df["group_no"].astype(str).str.strip() == "")]
            else:
                selected_members = group_df[group_df["group_no"].astype(str) == selected_group]

            if selected_members.empty:
                st.info("This group no longer exists in the roster. Select another group.")
                st.session_state["selected_evaluation_group"] = None
            else:
                titles = [str(t).strip() for t in selected_members["ml_project_title"].tolist() if str(t).strip() and str(t).lower() != "nan"]
                project_title = titles[0] if titles else "Project title to be confirmed"
                group_label = "Not assigned" if selected_group == "Not assigned" else f"Group {selected_group}"
                total_students = len(selected_members)
                awarded_total = sum(marks_total(get_evaluation(str(r["roll_no"]))) for _, r in selected_members.iterrows())
                max_group_marks = total_students * TOTAL_MARKS
                completion = round((awarded_total / max_group_marks) * 100) if max_group_marks else 0
                # Shared title banner and student/project summary.
                partner_initials = []
                partner_palette = [("#e6f0ff", "#2463cf"), ("#f0eaff", "#6338c5"), ("#e2f7ef", "#138b68")]
                for pidx, (_, prow) in enumerate(selected_members.iterrows()):
                    p_name = str(prow["student_name"])
                    p_initials = "".join(part[0] for part in p_name.split()[:2]).upper() or "ST"
                    p_bg, p_fg = partner_palette[pidx % len(partner_palette)]
                    partner_initials.append(f"<span class='partners-mini-avatar' style='background:{p_bg};color:{p_fg}'>{html.escape(p_initials)}</span>")
                st.markdown("""
                <style>
                .partners-page-heading{display:flex;align-items:center;gap:20px;padding:25px 28px;margin:0 0 24px;border:1px solid #dce8ff;border-radius:20px;background:linear-gradient(110deg,#ffffff 0%,#f5f9ff 100%);box-shadow:0 7px 22px rgba(27,72,140,.06)}
                .partners-page-icon{width:86px;height:86px;flex:0 0 86px;display:flex;align-items:center;justify-content:center;border-radius:18px;background:linear-gradient(145deg,#4f91ff,#075be8);color:#fff;font-size:39px;box-shadow:0 8px 18px rgba(25,100,238,.2)}
                .partners-page-copy{flex:1;min-width:0}
                .partners-page-copy h1{font-size:clamp(1.55rem,2vw,2.25rem);line-height:1.2;color:#102c61;margin:0 0 8px;font-weight:850;overflow-wrap:anywhere}
                .partners-page-copy p{font-size:clamp(1rem,1.2vw,1.25rem);line-height:1.45;color:#6c7f9e;margin:0;overflow-wrap:anywhere}
                .partners-page-students{min-width:185px;padding:8px 18px;border-left:1px solid #dce6f4}
                .partners-page-students b{display:block;font-size:17px;color:#102c61;margin-bottom:10px}
                .partners-mini-row{display:flex;align-items:center;gap:8px}
                .partners-mini-avatar{width:54px;height:54px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:19px;font-weight:800;border:2px solid white;box-shadow:0 2px 8px rgba(18,45,85,.07)}
                .partners-page-project{min-width:250px;max-width:330px;border-left:1px solid #dce6f4;padding:8px 0 8px 24px}
                .partners-page-project span{display:block;color:#71829d;font-size:15px;margin-bottom:6px}
                .partners-page-project strong{display:block;color:#102c61;font-size:17px;line-height:1.45;overflow-wrap:anywhere}
                .partners-section-title{font-size:clamp(1.4rem,1.8vw,1.85rem);font-weight:850;color:#102c61;margin:8px 0 16px}
                @media(max-width:1000px){.partners-page-heading{flex-wrap:wrap}.partners-page-copy{min-width:55%}.partners-page-students,.partners-page-project{border-left:0;padding:10px 0;min-width:100%;max-width:none}.partners-page-icon{width:68px;height:68px;flex-basis:68px}}
                @media(max-width:700px){.partner-card{height:auto;min-height:330px}.partner-avatar{width:68px;height:68px;font-size:22px}.partner-identity h3{font-size:17px}}
                </style>
                """, unsafe_allow_html=True)
                st.markdown(
                    "<div class='partners-page-heading'>"
                    "<div class='partners-page-icon'>♧</div>"
                    "<div class='partners-page-copy'><h1>" + html.escape(group_label + " · Project Partners") + "</h1>"
                    "<p>Review each student's marks and continue their individual or group-wise evaluation.</p></div>"
                    "<div class='partners-page-students'><b>♧ &nbsp; Students (" + str(total_students) + ")</b>"
                    "<div class='partners-mini-row'>" + "".join(partner_initials) + "</div></div>"
                    "<div class='partners-page-project'><span>Project Title</span><strong>" + html.escape(project_title) + "</strong></div>"
                    "</div>", unsafe_allow_html=True
                )
                st.markdown("<div class='partners-section-title'>Project partners</div>", unsafe_allow_html=True)
                partner_cols = st.columns(3, gap="medium")
                active_key = f"active_group_student_{selected_group}"
                active_student = st.session_state.get(active_key)
                initials_colors = [("#e6f0ff", "#2463cf"), ("#f0eaff", "#6338c5"), ("#e2f7ef", "#138b68")]
                for idx, (_, partner) in enumerate(selected_members.iterrows()):
                    roll = str(partner["roll_no"])
                    partner_ev = get_evaluation(roll)
                    partner_total = marks_total(partner_ev)
                    pct = round(partner_total / TOTAL_MARKS * 100) if TOTAL_MARKS else 0
                    name = str(partner["student_name"])
                    initials = "".join([part[0] for part in name.split()[:2]]).upper() or "ST"
                    bg, fg = initials_colors[idx % len(initials_colors)]
                    with partner_cols[idx % 3]:
                        st.markdown(
                            "<div class='partner-card'>"
                            f"<div class='partner-top'><div class='partner-avatar' style='background:{bg};color:{fg}'>{html.escape(initials)}</div>"
                            f"<div class='partner-identity'><h3>{html.escape(name)}</h3><p>Roll No.: {html.escape(roll)} &nbsp;·&nbsp; <span class='group-pill'>{html.escape(group_label)}</span></p></div></div>"
                            f"<div class='partner-marks'><div class='partner-marks-line'><span>Marks</span><strong>{partner_total:.0f}<em> / {TOTAL_MARKS}</em></strong></div>"
                            f"<div class='partner-progress'><span style='width:{pct}%'></span></div><small>{pct}% completed</small></div></div>",
                            unsafe_allow_html=True
                        )
                selected = active_student
                # Evaluation is intentionally group-only: criteria are entered once with a
                # separate score column for every partner, preventing individual-mode drift.
                evaluation_mode = "Evaluate all group partners together"
                if evaluation_mode == "Evaluate all group partners together":
                    st.markdown("### Group evaluation · enter marks for all partners")
                    st.caption("Each criterion appears once. Enter marks in the student columns; each student's marks are saved to their own evaluation record.")
                    bulk_i = render_phase_cards(f"bulk_phase_index_{selected_group}", include_final=True)
                    if bulk_i < len(COMPONENTS):
                        viva_row_cols = st.columns([2.2, 1], gap="small", vertical_alignment="center")
                        with viva_row_cols[0]:
                            st.caption(f"Group {selected_group} · {COMPONENTS[bulk_i]['name']} · Project-specific viva questions")
                        with viva_row_cols[1]:
                            if st.button(f"↗ Viva Questions · Phase {bulk_i + 1}", key=f"bulk_viva_link_{selected_group}_{bulk_i}", use_container_width=True):
                                @st.dialog(f"Group {selected_group} · Phase {bulk_i + 1} Viva Questions", width="large")
                                def show_group_phase_viva_dialog():
                                    group_title_values = selected_members["ml_project_title"].dropna().astype(str)
                                    group_project_title = group_title_values.iloc[0] if not group_title_values.empty else "Machine Learning Project"
                                    st.markdown(f"**Project:** {html.escape(group_project_title)}")
                                    st.caption(f"Group {selected_group} · {COMPONENTS[bulk_i]['name']}")
                                    group_phase_questions = get_group_phase_viva_questions(selected_group, bulk_i)
                                    if group_phase_questions:
                                        for qno, question in group_phase_questions:
                                            st.markdown(f"**Q{qno}.** {question}")
                                    else:
                                        for qno, question in enumerate(project_viva_questions(group_project_title, bulk_i), 1):
                                            st.markdown(f"**Q{qno}.** {question}")
                                show_group_phase_viva_dialog()
                    if bulk_i >= len(COMPONENTS):
                        st.markdown("### Final Review · all group students")
                        st.caption("This review combines every partner's saved evaluation. Strengths and improvement areas are shown student-by-student and as a group; no individual evaluation mode is available.")
                        group_feedback = []
                        all_strengths, all_weaknesses = [], []
                        for _, review_student in selected_members.iterrows():
                            review_roll = str(review_student["roll_no"])
                            review_name = str(review_student["student_name"])
                            review_ev = get_evaluation(review_roll)
                            review_strengths, review_weaknesses = generate_overall_feedback(review_ev)
                            review_total = marks_total(review_ev)
                            review_pct = round(review_total / TOTAL_MARKS * 100) if TOTAL_MARKS else 0
                            def _feedback_lines(value):
                                return [line.strip().lstrip("•-* ").strip() for line in str(value).splitlines() if line.strip()]
                            strength_lines = _feedback_lines(review_strengths)
                            weakness_lines = _feedback_lines(review_weaknesses)
                            all_strengths.extend(f"{review_name}: {item}" for item in strength_lines if "No marks have been recorded" not in item)
                            all_weaknesses.extend(f"{review_name}: {item}" for item in weakness_lines if "No marks have been recorded" not in item)
                            group_feedback.append((review_name, review_roll, review_total, review_pct, strength_lines, weakness_lines))
                        summary_cols = st.columns(3, gap="medium")
                        summary_cols[0].metric("Group students", len(group_feedback))
                        summary_cols[1].metric("Group marks awarded", f"{sum(item[2] for item in group_feedback):.0f} / {len(group_feedback) * TOTAL_MARKS}")
                        group_average = round(sum(item[3] for item in group_feedback) / max(len(group_feedback), 1))
                        summary_cols[2].metric("Average performance", f"{group_average}%")
                        strengths_col, weaknesses_col = st.columns(2, gap="medium")
                        with strengths_col:
                            st.markdown("#### ✅ Group strengths")
                            if all_strengths:
                                for item in all_strengths:
                                    st.markdown(f"- {item}")
                            else:
                                st.info("No strengths can be identified until marks are entered and saved.")
                        with weaknesses_col:
                            st.markdown("#### ⚠️ Group weaknesses / areas to improve")
                            if all_weaknesses:
                                for item in all_weaknesses:
                                    st.markdown(f"- {item}")
                            else:
                                st.info("No improvement areas are currently identified.")
                        st.markdown("#### Student-wise review")
                        for review_name, review_roll, review_total, review_pct, strength_lines, weakness_lines in group_feedback:
                            with st.container(border=True):
                                st.markdown(f"**{review_name}** · Roll {review_roll} · **{review_total:.0f}/{TOTAL_MARKS} ({review_pct}%)**")
                                detail_cols = st.columns(2, gap="medium")
                                with detail_cols[0]:
                                    st.markdown("**Strengths**")
                                    for item in strength_lines:
                                        st.markdown(f"- {item}")
                                with detail_cols[1]:
                                    st.markdown("**Areas for improvement**")
                                    for item in weakness_lines:
                                        st.markdown(f"- {item}")
                        st.info("Final Review is a read-only group summary. Each student's phase marks remain stored in that student's own record.")
                    else:
                        bulk_comp = COMPONENTS[bulk_i]
                        bulk_rows = []
                        bulk_evaluations = {}
                        for _, bulk_student in selected_members.iterrows():
                            bulk_roll = str(bulk_student["roll_no"])
                            bulk_evaluations[bulk_roll] = get_evaluation(bulk_roll)
                            bulk_rows.append((bulk_roll, str(bulk_student["student_name"])))
                        # Header and per-student summary, constrained to the current project group (max 3 partners).
                        header_cols = st.columns([2.6] + [1] * len(bulk_rows), gap="small")
                        header_cols[0].markdown("**Assessment criterion**")
                        for col, (bulk_roll, bulk_name) in zip(header_cols[1:], bulk_rows):
                            col.markdown(f"**{html.escape(bulk_name)}**")
                            col.caption(f"Roll {html.escape(bulk_roll)}")
                        st.divider()
                        for bulk_j, (bulk_criterion, bulk_maxmark) in enumerate(bulk_comp["criteria"]):
                            criterion_cols = st.columns([2.6] + [1] * len(bulk_rows), gap="small", vertical_alignment="center")
                            with criterion_cols[0]:
                                st.markdown(f"**{bulk_j+1}. {html.escape(bulk_criterion)}**")
                                st.caption(f"Maximum: {bulk_maxmark} marks per student")
                            for student_col, (bulk_roll, bulk_name) in zip(criterion_cols[1:], bulk_rows):
                                score_key = f"bulk_score_{selected_group}_{bulk_roll}_{bulk_i}_{bulk_j}"
                                old_score = int(bulk_evaluations[bulk_roll]["scores"][bulk_comp["name"]].get(bulk_criterion, 0))
                                if score_key not in st.session_state:
                                    st.session_state[score_key] = old_score
                                with student_col:
                                    # Compact score control: red decrement on the left, centered value, green increment on the right.
                                    control_cols = st.columns([0.72, 1.25, 0.72], gap="small", vertical_alignment="center")
                                    with control_cols[0]:
                                        st.button(
                                            "−", key=f"minus_{score_key}", use_container_width=True,
                                            disabled=int(st.session_state[score_key]) <= 0,
                                            on_click=adjust_mark_widget, args=(score_key, -1, bulk_maxmark)
                                        )
                                    with control_cols[1]:
                                        score_value = st.number_input(
                                            f"{bulk_name} · {bulk_criterion}", min_value=0, max_value=bulk_maxmark, step=1,
                                            key=score_key, label_visibility="collapsed"
                                        )
                                    with control_cols[2]:
                                        st.button(
                                            "+", key=f"plus_{score_key}", use_container_width=True,
                                            disabled=int(st.session_state[score_key]) >= bulk_maxmark,
                                            on_click=adjust_mark_widget, args=(score_key, 1, bulk_maxmark)
                                        )
                                bulk_evaluations[bulk_roll]["scores"][bulk_comp["name"]][bulk_criterion] = int(score_value)
                            st.markdown("<div class='bulk-criterion-divider'></div>", unsafe_allow_html=True)
                        st.markdown("#### Current phase totals")
                        total_cols = st.columns(len(bulk_rows), gap="small")
                        for col, (bulk_roll, bulk_name) in zip(total_cols, bulk_rows):
                            phase_total = sum(int(bulk_evaluations[bulk_roll]["scores"][bulk_comp["name"]].get(c, 0)) for c, _ in bulk_comp["criteria"])
                            with col:
                                st.metric(bulk_name, f"{phase_total} / {bulk_comp['max']}")
                        if st.button("Save marks for all group partners", key=f"save_bulk_{selected_group}_{bulk_i}", type="primary", use_container_width=True):
                            for bulk_roll, bulk_name in bulk_rows:
                                save_evaluation(bulk_roll, bulk_evaluations[bulk_roll])
                                log_activity(
                                    "Group evaluation saved",
                                    f"Saved {bulk_comp['name']} marks for {bulk_name} (Roll {bulk_roll})",
                                    f"Group {selected_group} · {bulk_comp['name']} · {sum(int(bulk_evaluations[bulk_roll]['scores'][bulk_comp['name']].get(c, 0)) for c, _ in bulk_comp['criteria'])}/{bulk_comp['max']}"
                                )
                            st.success(f"Saved {bulk_comp['name']} marks for all {len(bulk_rows)} group partner(s).")
                            st.rerun()
                elif selected is None:
                    st.info("Select **Evaluate Student** on a project partner card, or choose **Evaluate all group partners together** to enter marks for the whole group in one view.")
                elif selected not in selected_members["roll_no"].astype(str).tolist():
                    st.session_state.pop(active_key, None)
                    st.info("Select a student from this group to continue.")
                else:
                    st.divider()
                    st.markdown("### Individual student evaluation")
                    st.caption("Marks, evidence checklists, and remarks are stored separately for each student.")

                    row = students_df[students_df["roll_no"].astype(str) == selected].iloc[0]
                    ev = get_evaluation(selected)

                    project_title = str(row.get("ml_project_title", "") or "Project title to be confirmed")
                    group_label = str(row.get("group_no", "—"))
                    group_members = students_df[students_df["group_no"].astype(str) == group_label]
                    group_total = len(group_members)
                    group_scores = [marks_total(get_evaluation(str(member["roll_no"]))) for _, member in group_members.iterrows()]
                    group_max = max(group_total * TOTAL_MARKS, 1)
                    group_pct = round(sum(group_scores) / group_max * 100)
                    avatar_html = "".join("<span class='eval-avatar'>" + html.escape("".join(part[0] for part in str(member["student_name"]).split()[:2]).upper() or "ST") + "</span>" for _, member in group_members.iterrows())
                    st.markdown(
                        "<div class='eval-project-banner'><div class='eval-project-main'><div class='eval-project-icon'>▤</div><div class='eval-project-copy'>"
                        + "<div class='eval-project-title'>" + html.escape(project_title) + "</div>"
                        + "<div class='eval-project-subtitle'>Group " + html.escape(group_label) + " · Semester 7 · Machine Learning Project</div></div></div>"
                        + "<div class='eval-project-students'><div class='eval-mini-label'>Project partners (" + str(group_total) + ")</div><div class='eval-avatar-row'>" + avatar_html + "</div></div>"
                        + "<div class='eval-progress-wrap'><div class='eval-progress-head'><span>Group evaluation progress</span><strong>" + str(group_pct) + "%</strong></div>"
                        + "<div class='eval-progress-track'><div class='eval-progress-fill' style='width:" + str(group_pct) + "%'></div></div>"
                        + "<div class='eval-progress-caption'>" + str(sum(score > 0 for score in group_scores)) + " / " + str(group_total) + " students with marks recorded</div></div></div>",
                        unsafe_allow_html=True
                    )
                    st.markdown("<div class='eval-student-line'><span>Currently evaluating: <strong>" + html.escape(str(row['student_name'])) + "</strong></span><span>Roll No.: <strong>" + html.escape(selected) + "</strong></span><span>Individual marks are saved separately for every student.</span></div>", unsafe_allow_html=True)

                    i = render_phase_cards(f"phase_nav_index_{selected}", include_final=True)
                    if i < len(COMPONENTS):
                        comp = COMPONENTS[i]
                        subtotal = sum(int(ev["scores"][comp["name"]].get(criterion, 0)) for criterion, _ in comp["criteria"])
                        percentage = round((subtotal / comp["max"]) * 100) if comp["max"] else 0
                        st.markdown(
                            f'<div class="phase-header"><div><div class="phase-title">{html.escape(comp["name"])}</div>'
                            f'<div class="phase-subtitle">Maximum: {comp["max"]} marks · Assess each criterion below</div></div>'
                            f'<div class="phase-score-card"><div class="phase-score-label">MARKS AWARDED</div>'
                            f'<div class="phase-score-line"><div class="phase-score-number">{subtotal} <span>/ {comp["max"]}</span></div>'
                            f'<div class="phase-progress"><div class="phase-progress-fill" style="width:{percentage}%"></div></div>'
                            f'<strong style="color:#526987;font-size:.85rem">{percentage}%</strong></div></div></div>',
                            unsafe_allow_html=True
                        )
                        evidence_link_col, viva_link_col = st.columns([1, 1], gap="small")
                        with evidence_link_col:
                            if st.button("↗ Evidence List", key=f"evidence_link_{selected}_{i}", type="secondary"):
                                @st.dialog("Evidence List", width="large")
                                def show_evidence_dialog():
                                    st.markdown(f"### {comp['name']}")
                                    st.caption("Use this list to verify the student's project artefacts and implementation evidence.")
                                    for evidence_item in comp["checks"]:
                                        st.markdown(f"- {evidence_item}")
                                show_evidence_dialog()
                        with viva_link_col:
                            if st.button("↗ Viva Questions", key=f"viva_link_{selected}_{i}", type="secondary"):
                                @st.dialog("Project-specific Viva Questions", width="large")
                                def show_viva_dialog():
                                    project_title = row["ml_project_title"] or "Project title to be confirmed"
                                    st.markdown(f"**Project:** {project_title}")
                                    st.caption(f"Suggested questions for {comp['name']}. Use them as prompts and ask follow-up questions based on the student's answers.")
                                    for qno, question in enumerate(project_viva_questions(row["ml_project_title"], i), 1):
                                        st.markdown(f"**Q{qno}.** {question}")
                                show_viva_dialog()
                        for j, (criterion, maxmark) in enumerate(comp["criteria"]):
                            score_key = f"score_{selected}_{i}_{j}"
                            old = int(ev["scores"][comp["name"]].get(criterion, 0))
                            if score_key not in st.session_state:
                                st.session_state[score_key] = old
                            row_cols = st.columns([0.65, 5.6, 0.55, 0.8, 0.55], gap="small", vertical_alignment="center")
                            with row_cols[0]:
                                st.markdown(f'<div class="criterion-number">{j+1}</div>', unsafe_allow_html=True)
                            with row_cols[1]:
                                st.markdown(
                                    f'<div class="criterion-title">{html.escape(criterion)}</div>'
                                    f'<div class="criterion-max">Out of {maxmark}</div>',
                                    unsafe_allow_html=True
                                )
                            with row_cols[2]:
                                st.button("−", key=f"minus_{selected}_{i}_{j}", use_container_width=True,
                                          disabled=int(st.session_state[score_key]) <= 0,
                                          on_click=adjust_mark_widget, args=(score_key, -1, maxmark))
                            with row_cols[3]:
                                val = st.number_input(
                                    f"Marks for criterion {j+1}", min_value=0, max_value=maxmark, step=1,
                                    key=score_key, label_visibility="collapsed"
                                )
                            with row_cols[4]:
                                st.button("+", key=f"plus_{selected}_{i}_{j}", use_container_width=True,
                                          disabled=int(st.session_state[score_key]) >= maxmark,
                                          on_click=adjust_mark_widget, args=(score_key, 1, maxmark))
                            ev["scores"][comp["name"]][criterion] = int(val)

                        subtotal = sum(int(ev["scores"][comp["name"]].get(criterion, 0)) for criterion, _ in comp["criteria"])
                        remaining = comp["max"] - subtotal
                        sum_cols = st.columns([1, 1, 2])
                        sum_cols[0].metric("Component score", f"{subtotal} / {comp['max']}")
                        sum_cols[1].metric("Marks remaining", remaining)
                        sum_cols[2].progress(min(subtotal / comp["max"], 1.0) if comp["max"] else 0.0, text=f"{round(subtotal / comp['max'] * 100) if comp['max'] else 0}% completed")
                    else:
                        i = 7

                    if i == len(COMPONENTS):
                        generated_strengths, generated_weaknesses = generate_overall_feedback(ev)
                        total_awarded = marks_total(ev)
                        performance_pct = round((total_awarded / TOTAL_MARKS) * 100) if TOTAL_MARKS else 0
                        completed_components = sum(
                            1 for component in COMPONENTS
                            if sum(float(ev.get("scores", {}).get(component["name"], {}).get(criterion, 0) or 0)
                                   for criterion, _ in component["criteria"]) > 0
                        )
                        completion_pct = round(completed_components / len(COMPONENTS) * 100) if COMPONENTS else 0
                        status_label = "Ready for Review" if completed_components == len(COMPONENTS) else "In Progress"
                        if performance_pct >= 75:
                            performance_label = "Excellent"
                        elif performance_pct >= 60:
                            performance_label = "Good"
                        elif performance_pct >= 40:
                            performance_label = "Developing"
                        else:
                            performance_label = "Needs Support"

                        st.markdown(
                            '<div class="final-review-hero">'
                            '<div class="final-review-head"><div class="final-review-icon">▣</div>'
                            '<div style="min-width:0;flex:1"><div class="final-review-title">Final Review &amp; Submission</div>'
                            '<div class="final-review-subtitle">Verify deliverables, review overall assessment and record the final recommendation.</div></div></div>'
                            '<div class="final-review-metrics">'
                            f'<div class="final-review-metric"><div class="final-review-metric-icon blue">🏆</div><div><div class="final-review-metric-label">Total Marks</div><div class="final-review-metric-value">{total_awarded:.0f} / {TOTAL_MARKS}</div></div></div>'
                            f'<div class="final-review-metric"><div class="final-review-metric-icon green">◉</div><div><div class="final-review-metric-label">Overall Performance</div><div class="final-review-metric-value">{performance_pct}%</div><div class="final-review-metric-label">{html.escape(performance_label)}</div></div></div>'
                            f'<div class="final-review-metric"><div class="final-review-metric-icon teal">▥</div><div><div class="final-review-metric-label">Phases Completed</div><div class="final-review-metric-value">{completed_components} / {len(COMPONENTS)}</div></div></div>'
                            f'<div class="final-review-metric"><div class="final-review-metric-icon blue">▣</div><div><div class="final-review-metric-label">Submission Status</div><div class="final-review-metric-value" style="font-size:17px">{html.escape(status_label)}</div></div></div>'
                            '</div></div>', unsafe_allow_html=True
                        )
                        evidence_link_col, viva_link_col = st.columns([1, 1], gap="small")
                        with evidence_link_col:
                            if st.button("↗ Evidence List", key=f"evidence_link_{selected}_7", type="secondary", use_container_width=True):
                                @st.dialog("Final Review Evidence List", width="large")
                                def show_final_evidence_dialog():
                                    st.caption("Verify the final project deliverables before assigning the final recommendation.")
                                    for evidence_item in SUBMISSION_ITEMS:
                                        st.markdown(f"- {evidence_item}")
                                show_final_evidence_dialog()
                        with viva_link_col:
                            if st.button("↗ Viva Questions", key=f"viva_link_{selected}_7", type="secondary", use_container_width=True):
                                @st.dialog("Project-specific Final Review Viva Questions", width="large")
                                def show_final_viva_dialog():
                                    project_title = row["ml_project_title"] or "Project title to be confirmed"
                                    st.markdown(f"**Project:** {project_title}")
                                    st.caption("Suggested questions for final review and demonstration.")
                                    for qno, question in enumerate(project_viva_questions(row["ml_project_title"], 7), 1):
                                        st.markdown(f"**Q{qno}.** {question}")
                                show_final_viva_dialog()

                        st.markdown(
                            '<div class="final-review-banner"><div class="final-review-banner-icon">💡</div>'
                            '<div><div class="final-review-banner-title">Automatically generated overall assessment</div>'
                            '<div class="final-review-banner-copy">Generated from the marks recorded across all evaluation phases. Save marks in the relevant phases first; this summary updates from the stored evaluation data.</div></div></div>',
                            unsafe_allow_html=True
                        )

                        def feedback_items(text):
                            items = []
                            for line in str(text).splitlines():
                                cleaned = line.strip().lstrip("•-* ").strip()
                                if cleaned:
                                    items.append(cleaned)
                            return items

                        strength_items = feedback_items(generated_strengths)
                        weakness_items = feedback_items(generated_weaknesses)
                        strengths_html = []
                        for item_no, item_text in enumerate(strength_items, 1):
                            # Strength cards intentionally show descriptive evidence only, never marks or percentages.
                            clean_text = html.escape(item_text)
                            phase_name = next((component["name"] for component in COMPONENTS if component["name"] in item_text), "Overall")
                            display_text = clean_text.replace(html.escape(phase_name), "").strip(" —:-") if phase_name != "Overall" else clean_text
                            strengths_html.append(
                                '<div class="final-review-item"><div class="final-review-number">✓</div>'
                                f'<div class="final-review-item-text">{display_text}'
                                f'<div class="final-review-item-sub">{html.escape(phase_name)}</div></div>'
                                f'<div class="final-review-tag">{html.escape(phase_name.split(" – ")[0] if phase_name != "Overall" else "Overall")}</div></div>'
                            )
                        weaknesses_html = []
                        for item_no, item_text in enumerate(weakness_items, 1):
                            clean_text = html.escape(item_text)
                            phase_name = next((component["name"] for component in COMPONENTS if component["name"] in item_text), "Review")
                            weaknesses_html.append(
                                '<div class="final-review-item weak"><div class="final-review-number">!</div>'
                                f'<div class="final-review-item-text">{clean_text}'
                                f'<div class="final-review-item-sub">{html.escape(phase_name)}</div></div>'
                                f'<div class="final-review-tag">{html.escape(phase_name.split(" – ")[0] if phase_name != "Review" else "Review")}</div></div>'
                            )
                        strengths_body = "".join(strengths_html) or '<div class="final-review-item"><div class="final-review-item-text">No strong components identified yet. Save marks across the phases to generate the checklist.</div></div>'
                        weaknesses_body = "".join(weaknesses_html) or '<div class="final-review-item weak"><div class="final-review-item-text">No improvement areas currently identified.</div></div>'
                        left, right = st.columns(2, gap="medium")
                        with left:
                            st.markdown(
                                '<div class="final-review-panel"><div class="final-review-panel-head">'
                                '<div class="final-review-panel-icon">▥</div><div style="min-width:0"><div class="final-review-panel-title">Overall Strengths</div>'
                                '<div class="final-review-panel-subtitle">Key strong areas identified from phase-wise evaluation</div></div>'
                                f'<div class="final-review-count">{len(strength_items)} strengths</div></div>{strengths_body}</div>',
                                unsafe_allow_html=True
                            )
                        with right:
                            st.markdown(
                                '<div class="final-review-panel weak"><div class="final-review-panel-head">'
                                '<div class="final-review-panel-icon">◎</div><div style="min-width:0"><div class="final-review-panel-title">Weaknesses / Areas for Improvement</div>'
                                '<div class="final-review-panel-subtitle">Aspects that need more focus before final submission</div></div>'
                                f'<div class="final-review-count">{len(weakness_items)} areas</div></div>{weaknesses_body}</div>',
                                unsafe_allow_html=True
                            )

                        # Keep generated text in the existing fields for compatibility with exports.
                        ev["overall_strengths"] = generated_strengths
                        ev["overall_improvements"] = generated_weaknesses
                        st.markdown('<div class="final-review-recommendation"><div class="final-review-recommendation-title">▣ &nbsp; Final Recommendation</div><div class="final-review-recommendation-subtitle">Select the final recommendation based on the overall evaluation.</div></div>', unsafe_allow_html=True)
                        options = ["Accepted", "Accepted with minor corrections", "Resubmission required"]
                        current = ev.get("recommendation", "Accepted with minor corrections")
                        ev["recommendation"] = st.radio(
                            "Final recommendation", options,
                            index=options.index(current) if current in options else 1,
                            key=f"recommendation_{selected}", horizontal=True
                        )

                    st.divider()
                    st.metric("Final Total", f"{marks_total(ev):.0f} / {TOTAL_MARKS}")
                    if st.button("Save this student's evaluation", type="primary", key=f"save_evaluation_{selected}"):
                        save_evaluation(selected, ev)
                        log_activity("Evaluation saved", f"Saved evaluation for {row['student_name']} (Roll {selected})", f"Group {row.get('group_no', '—')} · Total {marks_total(ev):.0f}/{TOTAL_MARKS} · Recommendation: {ev.get('recommendation', '')}")
                        st.success(f"Evaluation saved for {row['student_name']} (Roll {selected}).")
                        st.rerun()


else:
    st.subheader("Records & Export")
    st.write("Review every student's saved marks and download a class marksheet in Excel or PDF format.")
    st.info("The evaluation is out of 100. The internal marksheet score is calculated as half of the evaluation total, out of 50.")
    summary=[]
    details=[]
    for _,r in students_df.iterrows():
        roll=str(r["roll_no"]); ev=get_evaluation(roll)
        total = float(marks_total(ev))
        half_total = total / 2
        rec={"Group":r["group_no"],"Roll No.":roll,"Student Name":r["student_name"],"ML Project Title":r["ml_project_title"],"Marks (Out of 100)":round(total,2),"Marks (Out of 50)":round(half_total,2),"Recommendation":ev.get("recommendation",""),"Last Updated":ev.get("last_updated","")}
        for comp in COMPONENTS:
            rec[comp["name"]]=sum(ev["scores"][comp["name"]].values())
            for criterion,maxmark in comp["criteria"]:
                details.append({"Group":r["group_no"],"Roll No.":roll,"Student Name":r["student_name"],"ML Project Title":r["ml_project_title"],"Component":comp["name"],"Criterion":criterion,"Awarded Marks":ev["scores"][comp["name"]].get(criterion,0),"Maximum Marks":maxmark,"Faculty Remarks":ev["remarks"].get(comp["name"],"")})
        summary.append(rec)
    summary_df=pd.DataFrame(summary)
    details_df=pd.DataFrame(details)

    st.markdown("### Student marksheet")
    st.caption("Marks out of 50 are automatically calculated as Total Marks ÷ 2. Records update from the saved evaluation data.")
    marksheet_df = summary_df[["Group", "Roll No.", "Student Name", "ML Project Title", "Marks (Out of 100)", "Marks (Out of 50)"]].copy()
    st.dataframe(
        marksheet_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Group": st.column_config.NumberColumn("Group", format="%d"),
            "Roll No.": st.column_config.TextColumn("Roll No."),
            "Student Name": st.column_config.TextColumn("Student Name", width="medium"),
            "ML Project Title": st.column_config.TextColumn("Project Title", width="large"),
            "Marks (Out of 100)": st.column_config.NumberColumn("Marks / 100", format="%.2f"),
            "Marks (Out of 50)": st.column_config.NumberColumn("Marks / 50", format="%.2f"),
        },
    )

    from io import BytesIO
    excel_bytes = b""
    excel_error = None
    try:
        # Excel export needs openpyxl. Missing optional dependencies must not crash the page.
        import openpyxl  # noqa: F401
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            marksheet_df.to_excel(writer, index=False, sheet_name="Class Marksheet")
            summary_df.to_excel(writer, index=False, sheet_name="Full Summary")
            details_df.to_excel(writer, index=False, sheet_name="Criterion Details")
            # Make each worksheet easier to read when opened in Excel.
            for worksheet in writer.book.worksheets:
                worksheet.freeze_panes = "A2"
                worksheet.auto_filter.ref = worksheet.dimensions
                for column_cells in worksheet.columns:
                    max_length = max((len(str(cell.value or "")) for cell in column_cells), default=10)
                    worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 42)
        excel_bytes = excel_buffer.getvalue()
    except ImportError as exc:
        excel_error = exc
        st.warning("Excel export needs the optional package `openpyxl`. Install it in the same Python environment running Streamlit: `python -m pip install openpyxl`. PDF and CSV exports remain available.")

    pdf_buffer = BytesIO()
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name="MarksheetTitle", parent=styles["Title"], alignment=TA_CENTER, textColor=colors.HexColor("#153768"), fontSize=17, leading=21))
        styles.add(ParagraphStyle(name="CellSmall", parent=styles["BodyText"], fontSize=7, leading=9, wordWrap="CJK"))
        story = [Paragraph("Semester 7 · Machine Learning Project Marksheet", styles["MarksheetTitle"]), Spacer(1, 4*mm), Paragraph("Evaluation total is out of 100; the marksheet score is half of the total, out of 50.", styles["BodyText"]), Spacer(1, 4*mm)]
        pdf_cols = ["Group", "Roll No.", "Student Name", "ML Project Title", "Marks (Out of 100)", "Marks (Out of 50)"]
        table_data = [[Paragraph(f"<b>{c}</b>", styles["CellSmall"]) for c in ["Group", "Roll No.", "Student Name", "Project Title", "Marks / 100", "Marks / 50"]]]
        for _, row in marksheet_df.iterrows():
            table_data.append([Paragraph(str(row[c]).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), styles["CellSmall"]) for c in pdf_cols])
        table = Table(table_data, repeatRows=1, colWidths=[18*mm, 22*mm, 43*mm, 92*mm, 27*mm, 27*mm])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#155EEF")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#D8E2F0")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F4F8FF")]),
            ("ALIGN", (0,0), (1,-1), "CENTER"),
            ("ALIGN", (4,1), (5,-1), "CENTER"),
            ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(table)
        doc.build(story)
        pdf_bytes = pdf_buffer.getvalue()
    except Exception as exc:
        pdf_bytes = b""
        st.warning(f"PDF export could not be prepared: {exc}. Install reportlab with `python -m pip install reportlab`.")

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        if excel_bytes:
            if st.download_button("Download marksheet (Excel)", excel_bytes, "semester7_ml_marksheet.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="export_marksheet_excel", use_container_width=True):
                log_activity("Export", "Downloaded class marksheet Excel", f"{len(marksheet_df)} student records")
        else:
            st.button("Download marksheet (Excel)", disabled=True, key="export_marksheet_excel_disabled", use_container_width=True, help="Install openpyxl to enable Excel export.")
    with col2:
        if pdf_bytes and st.download_button("Download marksheet (PDF)", pdf_bytes, "semester7_ml_marksheet.pdf", "application/pdf", key="export_marksheet_pdf", use_container_width=True):
            log_activity("Export", "Downloaded class marksheet PDF", f"{len(marksheet_df)} student records")

    with st.expander("Additional exports", expanded=False):
        if st.download_button("Download full class summary (CSV)",summary_df.to_csv(index=False).encode("utf-8-sig"),"semester7_ml_class_summary.csv","text/csv", key="export_class_summary"):
            log_activity("Export", "Downloaded class summary CSV", f"{len(summary_df)} student records")
        if st.download_button("Download detailed evaluations (CSV)",details_df.to_csv(index=False).encode("utf-8-sig"),"semester7_ml_detailed_evaluations.csv","text/csv", key="export_detailed_evaluations"):
            log_activity("Export", "Downloaded detailed evaluation CSV", f"{len(details_df)} criterion records")
        with db() as con:
            db_bytes=Path(DB_PATH).read_bytes() if Path(DB_PATH).exists() else b""
        if st.download_button("Download SQLite database",db_bytes,"ml_project_evaluation.db","application/x-sqlite3", key="export_database"):
            log_activity("Export", "Downloaded SQLite database backup", "ml_project_evaluation.db")
    st.warning("Streamlit Community Cloud's local filesystem may reset after redeployment or restart. For permanent production records, use hosted PostgreSQL/Supabase or download the database and marksheet backups regularly.")

st.divider()
st.caption("Mark scheme: Phase 1 15 · Phase 2 15 · Phase 3 25 · Phase 4 15 · Report 10 · Presentation 10 · Viva 10 = 100.")
