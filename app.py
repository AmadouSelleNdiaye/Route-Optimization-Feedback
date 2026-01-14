import os
import json
import base64
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

# ------------------------- Config ------------------------- #
IDC_FILE = "IDC NAME.xlsx"
STATIONS_FILE = "Station adresses ASN.xlsx"

# ✅ Banner image ONLY for title block (local file)
BANNER_FILE = "MicrosoftFormTheme.jpg"

SUBMISSIONS_DIR = "submissions"
ATTACHMENTS_DIR = os.path.join(SUBMISSIONS_DIR, "attachments")

# ✅ Excel export (single file updated every submission)
EXCEL_EXPORT_FILE = os.path.join(SUBMISSIONS_DIR, "route_optimization_feedback.xlsx")

# ✅ From screenshots (added)
IDC_LIAISON_OPTIONS = [
    "I don't know",
    "Adam M.",
    "Azhar N.",
    "Bhugesh Y.",
    "Emmanuelle L.",
    "Jeffrey L.",
    "Meshwa K.",
    "Farsheed F.",
    "Safiétou D.",
    "Sam A.",
    "Will M.",
]

VEHICLE_TYPE_OPTIONS = [
    "I don't know",
    "Gas 120 cuft",
    "Gas 280 cuft",
    "Cargo Bikes",
    "Ford E-transit EV",
    "Esprinter EV",
    "Brightdrop EV",
]

# ------------------------- Helpers ------------------------- #
@st.cache_data(show_spinner=False)
def load_idc_list(path: str) -> list[str]:
    df = pd.read_excel(path)
    col = "COMPANY_NAME"
    if col not in df.columns:
        raise ValueError(f"IDC file must contain '{col}'")
    return sorted(df[col].astype(str).str.strip().replace("nan", "").unique().tolist())


@st.cache_data(show_spinner=False)
def load_station_list(path: str) -> list[str]:
    df = pd.read_excel(path)
    col = "Station Tag"
    if col not in df.columns:
        raise ValueError(f"Stations file must contain '{col}'")
    return sorted(df[col].astype(str).str.strip().replace("nan", "").unique().tolist())


def ensure_dirs():
    os.makedirs(SUBMISSIONS_DIR, exist_ok=True)
    os.makedirs(ATTACHMENTS_DIR, exist_ok=True)


def safe_filename(name: str) -> str:
    keep = []
    for ch in name:
        if ch.isalnum() or ch in ("-", "_", ".", " "):
            keep.append(ch)
    cleaned = "".join(keep).strip().replace(" ", "_")
    return cleaned[:150] if cleaned else "file"


def save_submission(payload: dict, uploaded_files: list) -> str:
    ensure_dirs()

    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    submission_id = f"rof_{ts}"

    attachment_paths = []
    for f in uploaded_files or []:
        if f is None:
            continue
        fname = safe_filename(f.name)
        out_path = os.path.join(ATTACHMENTS_DIR, f"{submission_id}__{fname}")
        with open(out_path, "wb") as w:
            w.write(f.getbuffer())
        attachment_paths.append(out_path)

    payload["submission_id"] = submission_id
    payload["attachments"] = attachment_paths
    payload["submitted_at_utc"] = datetime.now(timezone.utc).isoformat()

    out_json = os.path.join(SUBMISSIONS_DIR, f"{submission_id}.json")
    with open(out_json, "w", encoding="utf-8") as w:
        json.dump(payload, w, ensure_ascii=False, indent=2)

    return out_json


def append_submission_to_excel(payload: dict, excel_path: str) -> None:
    """
    Append one submission (payload dict) to a single Excel file.
    Creates the file if it doesn't exist. Keeps all columns that ever appeared.
    """
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    new_row = pd.DataFrame([payload])

    if os.path.exists(excel_path):
        try:
            existing = pd.read_excel(excel_path)
        except Exception:
            existing = pd.DataFrame()

        all_cols = list(dict.fromkeys(list(existing.columns) + list(new_row.columns)))
        existing = existing.reindex(columns=all_cols)
        new_row = new_row.reindex(columns=all_cols)

        updated = pd.concat([existing, new_row], ignore_index=True)
    else:
        updated = new_row

    updated.to_excel(excel_path, index=False)


def is_digits_only(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    return bool(s) and s.isdigit()


def banner_as_data_uri(path: str) -> str:
    if not os.path.exists(path):
        return ""
    ext = os.path.splitext(path)[1].lower()
    mime = "image/jpeg" if ext in [".jpg", ".jpeg"] else "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return f"data:{mime};base64,{b64}"


# ------------------------- Page config ------------------------- #
st.set_page_config(
    page_title="Route Optimization Feedback",
    page_icon="🧭",
    layout="centered",
)

logo_icon = "logo_intelcom_2024.png"
logo_image = "logo_intelcom_2024.png"
st.logo(icon_image=logo_icon, image=logo_image)

banner_uri = banner_as_data_uri(BANNER_FILE)

# ------------------------- CSS ------------------------- #
st.markdown(
    """
<style>
.block-container { max-width: 80%; padding-top: 2rem; }
.card { border-radius: 18px; overflow: hidden; box-shadow: 0 14px 36px rgba(0,0,0,0.18); margin-bottom: 20px; }
.title-banner { height: 190px; background-size: cover; background-position: center; position: relative; }
.title-overlay { height: 100%; background: rgba(0,0,0,0.45); padding: 34px 42px; display: flex; flex-direction: column; justify-content: center; }
.forms-title { font-size: 34px; font-weight: 700; margin: 0; color:#ffffff !important; }
.forms-subtitle { margin-top: 10px; font-size: 15px; color: rgba(255,255,255,0.9) !important; }

div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea,
div[data-baseweb="select"] > div { background-color: #f5f5f5 !important; }

footer { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)

# ------------------------- Load lists ------------------------- #
try:
    idc_list = load_idc_list(IDC_FILE)
    station_list = load_station_list(STATIONS_FILE)
except Exception as e:
    st.error(
        "Unable to load the Excel lists.\n\n"
        f"- Expected files next to `app.py`: `{IDC_FILE}` and `{STATIONS_FILE}`\n"
        f"- Error: {e}"
    )
    st.stop()

# ------------------------- TITLE BLOCK ------------------------- #
if banner_uri:
    banner_style = f"background-image: url('{banner_uri}');"
else:
    banner_style = "background: #2b2b2b;"

st.markdown(
    f"""
<div class="card">
  <div class="title-banner" style="{banner_style}">
    <div class="title-overlay">
      <h2 class="forms-title">Route Optimization Feedback</h2>
      <div class="forms-subtitle">Share your feedback with us</div>
    </div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ==========================================================
# ✅ DYNAMIC DEPENDENT DROPDOWN INSIDE THE "FORM-LIKE" AREA
# ==========================================================
MAIN_ISSUE_CATEGORIES = [
    "",
    "Routing",
    "Address / GPS location",
    "Access",
    "Road conditions",
    "Parking",
    "Business / customer issue",
    "Problems related to the app",
]

SUBCATS = {
    "Routing": [
        "Unnecessary detour",
        "Loop / backtracking",
        "Stop order doesn’t make sense",
    ],
    "Address / GPS location": [
        "Incorrect address",
        "Incorrect GPS pin",
        "Entrance/door is not at the right location",
    ],
    "Access": [
        "Gated community / security",
        "Access code",
        "Drop-off location hard to find",
    ],
    "Road conditions": [
        "Traffic",
        "Construction",
        "Weather",
    ],
    "Parking": [
        "Hard to park",
        "Truck restrictions",
    ],
    "Business / customer issue": [
        "Business closed",
        "Missing or inadequate instructions",
    ],
    "Problems related to the app": [
        "Connection problems (uploading POD)",
        "Offline map issues",
    ],
}

def _on_main_issue_change():
    st.session_state["sub_issue"] = ""

def _reset_form():
    keys = [
        "driver_last_name", "driver_first_name", "driver_email_required",
        "idc_liaison", "idc_select", "idc_new",
        "station_select", "station_new",
        "route_number", "route_date",
        "vehicle_type",
        "issue_applies_to",
        "stop_number", "stop_address",
        "severity", "route_satisfaction",
        "time_lost", "parcel_tracking_id",
        "main_issue", "sub_issue",
        "what_happened", "what_should", "suggestion",
        "attachments", "agree",
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]

# ------------------------- "FORM" (WITHOUT st.form) ------------------------- #
with st.container():

    # =========================
    # 1) Identification
    # =========================
    st.subheader("1) Identification")

    c1, c2 = st.columns(2)
    with c1:
        driver_last_name = st.text_input("Driver last name *", placeholder="Last name", key="driver_last_name")
        driver_first_name = st.text_input("Driver first name *", placeholder="First name", key="driver_first_name")
        idc_liaison = st.selectbox(
            "Who is your IDC Liaison? *",
            options=IDC_LIAISON_OPTIONS,
            index=0,
            key="idc_liaison",
        )

    with c2:
        driver_email_required = st.text_input("Driver email *", placeholder="name@company.com", key="driver_email_required")

        idc_options = ["", "I don't know"] + idc_list + ["➕ Add a new IDC"]
        selected_idc = st.selectbox("IDC *", options=idc_options, index=0, key="idc_select")
        if selected_idc == "➕ Add a new IDC":
            idc_name = st.text_input("Enter new IDC name *", placeholder="e.g., Amazon", key="idc_new")
        else:
            idc_name = selected_idc

    station_options = [""] + station_list + ["➕ Add a new Station"]
    selected_station = st.selectbox("Station *", options=station_options, index=0, key="station_select")
    if selected_station == "➕ Add a new Station":
        station = st.text_input("Enter new Station code * (e.g., MONT)", placeholder="MONT", key="station_new")
    else:
        station = selected_station

    col3, col4 = st.columns(2)
    with col3:
        route_number = st.text_input(
            "Route number * (digits only, e.g., 1235)",
            placeholder="1235",
            help="Numbers only. Example: 1235",
            key="route_number",
        )
        if route_number.strip() and (not route_number.strip().isdigit()):
            st.warning("Route number must contain digits only (e.g., 1235).")

    with col4:
        route_date = st.date_input("Route date *", key="route_date")

    vehicle_type = st.selectbox(
        "What type of Vehicle do you have? *",
        options=VEHICLE_TYPE_OPTIONS,
        index=0,
        key="vehicle_type",
    )

    # =========================
    # 2) Problem scope
    # =========================
    st.subheader("2) Problem scope")

    issue_applies_to = st.radio(
        "This issue applies to: *",
        options=["Entire route", "Specific stop"],
        horizontal=True,
        index=0,
        key="issue_applies_to",
    )

    stop_number = ""
    stop_address = ""
    if issue_applies_to == "Specific stop":
        s1, s2 = st.columns(2)
        with s1:
            stop_number = st.text_input("Stop number *", placeholder="e.g., 12", key="stop_number")
        with s2:
            stop_address = st.text_input("Stop address *", placeholder="e.g., 123 Main St, City", key="stop_address")

    severity = st.radio(
        "Severity *",
        options=["Low", "Medium", "High", "Critical"],
        horizontal=True,
        index=1,
        key="severity",
    )

    route_satisfaction = st.select_slider(
        "Route satisfaction / rating (optional)",
        options=["0", "1", "2", "3", "4", "5"],
        value="0",
        help="1 = very bad, 5 = excellent. Choose 0 if not applicable.",
        key="route_satisfaction",
    )

    time_lost = st.radio(
        "Estimated time lost *",
        options=["0–15 min", "15–30 min", "30–60 min", "60+ min"],
        horizontal=True,
        index=0,
        key="time_lost",
    )

    parcel_tracking_id = st.text_input(
        "Parcel Tracking ID",
        placeholder="Example: INTELXXX0001111",
        help="Required if specific stop / Optional if entire route",
        key="parcel_tracking_id",
    )

    # =========================
    # 3) Issue identification (dynamic)
    # =========================
    st.subheader("3) Issue identification")

    main_issue = st.selectbox(
        "Main issue category *",
        options=MAIN_ISSUE_CATEGORIES,
        index=0,
        key="main_issue",
        on_change=_on_main_issue_change,
    )

    available_subcats = SUBCATS.get(main_issue, [])

    sub_issue = st.selectbox(
        "Sub-category *",
        options=([""] + available_subcats) if main_issue else [""],
        index=0,
        disabled=(not main_issue),
        key="sub_issue",
        help="Select a main issue category first." if not main_issue else None,
    )

    what_happened = st.text_area(
        "What happened? (optional)",
        placeholder="Describe what you observed.",
        height=110,
        key="what_happened",
    )
    # =========================
    # 4) Details
    # =========================
    st.subheader("4) Details")



    what_should = st.text_area(
        "What should have happened ideally? (optional)",
        placeholder="Describe the expected behavior.",
        height=90,
        key="what_should",
    )

    suggestion = st.text_area(
        "Suggestion (optional)",
        placeholder="Any suggestion to improve the route/experience.",
        height=90,
        key="suggestion",
    )

    # =========================
    # 5) Attachments
    # =========================
    st.subheader("5) Attachments")

    attachments = st.file_uploader(
        "Attach screenshots / files (optional)",
        accept_multiple_files=True,
        type=["png", "jpg", "jpeg", "pdf", "xlsx", "csv", "txt"],
        key="attachments",
    )

    agree = st.checkbox(
        "I confirm this information is accurate to the best of my knowledge *",
        key="agree",
    )

    submitted = st.button("Submit feedback", type="primary")

    if submitted:
        errors = []

        # Clean
        driver_last_name_clean = (driver_last_name or "").strip()
        driver_first_name_clean = (driver_first_name or "").strip()
        driver_email_required_clean = (driver_email_required or "").strip()
        idc_name_clean = (idc_name or "").strip()
        station_clean = (station or "").strip()
        route_number_clean = (route_number or "").strip()
        idc_liaison_clean = (idc_liaison or "").strip()
        vehicle_type_clean = (vehicle_type or "").strip()
        parcel_tracking_id_clean = (parcel_tracking_id or "").strip()

        stop_number_clean = (stop_number or "").strip()
        stop_address_clean = (stop_address or "").strip()

        # Validation (Identification)
        if not driver_last_name_clean:
            errors.append("Driver last name is required.")
        if not driver_first_name_clean:
            errors.append("Driver first name is required.")
        if not driver_email_required_clean:
            errors.append("Driver email is required.")
        if not idc_name_clean:
            errors.append("IDC is required.")
        if not station_clean:
            errors.append("Station is required.")
        if not route_number_clean:
            errors.append("Route number is required.")
        elif not is_digits_only(route_number_clean):
            errors.append("Route number must contain digits only (e.g., 1235).")

        if not idc_liaison_clean:
            errors.append("IDC Liaison is required.")
        if not vehicle_type_clean:
            errors.append("Vehicle type is required.")

        # Validation (Problem scope)
        if issue_applies_to == "Specific stop":
            if not stop_number_clean:
                errors.append("Stop number is required when 'Specific stop' is selected.")
            if not stop_address_clean:
                errors.append("Stop address is required when 'Specific stop' is selected.")
            if not parcel_tracking_id_clean:
                errors.append("Parcel Tracking ID is required when 'Specific stop' is selected.")

        if not time_lost:
            errors.append("Estimated time lost is required.")

        # Validation (Issue identification)
        if not main_issue:
            errors.append("Main issue category is required.")
        if main_issue and not sub_issue:
            errors.append("Sub-category is required.")

        if not agree:
            errors.append("You must confirm the accuracy checkbox.")

        if errors:
            st.error("Please fix the following:\n- " + "\n- ".join(errors))
            st.stop()

        payload = {
            # Identification
            "driver_last_name": driver_last_name_clean,
            "driver_first_name": driver_first_name_clean,
            "driver_email": driver_email_required_clean,
            "idc": idc_name_clean,
            "idc_liaison": idc_liaison_clean,
            "station": station_clean,
            "route_number": route_number_clean,
            "route_date": route_date.isoformat(),
            "vehicle_type": vehicle_type_clean,
            "parcel_tracking_id": parcel_tracking_id_clean if parcel_tracking_id_clean else None,

            # Problem scope
            "issue_applies_to": issue_applies_to,
            "stop_number": stop_number_clean if issue_applies_to == "Specific stop" else None,
            "stop_address": stop_address_clean if issue_applies_to == "Specific stop" else None,
            "severity": severity,
            "route_satisfaction": route_satisfaction,
            "estimated_time_lost": time_lost,

            # Issue identification
            "main_issue_category": main_issue,
            "sub_category": sub_issue,

            # Details
            "what_happened": (what_happened or "").strip(),
            "what_should_have_happened": (what_should or "").strip(),
            "suggestion": (suggestion or "").strip(),
        }

        # ✅ Save JSON + attachments (existing behavior)
        out_path = save_submission(payload, attachments)

        # ✅ Append to Excel (single file updated each submission)
        append_submission_to_excel(payload, EXCEL_EXPORT_FILE)

        st.success("✅ Feedback submitted successfully!")
        st.caption(f"Saved JSON to: {out_path}")
        st.caption(f"Appended to Excel: {EXCEL_EXPORT_FILE}")
        st.json(payload)

        # ✅ Clear everything like clear_on_submit=True
        _reset_form()
        st.rerun()
