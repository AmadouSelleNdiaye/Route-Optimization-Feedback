import os
import json
import base64
from io import BytesIO
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
import requests

# ------------------------- Config ------------------------- #
IDC_FILE = "IDC NAME.xlsx"  # (plus utilisé après modif, tu peux supprimer si tu veux)
STATIONS_FILE = "Station adresses ASN.xlsx"
BANNER_FILE = "MicrosoftFormTheme.jpg"

# Local storage (JSON) — optionnel
SUBMISSIONS_DIR = "submissions"
ATTACHMENTS_DIR = os.path.join(SUBMISSIONS_DIR, "attachments")

# ------------------------- SharePoint / Graph Config ------------------------- #
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _get_secret_or_env(key: str, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.getenv(key, default)


# TENANT_ID = _get_secret_or_env("TENANT_ID")
# CLIENT_ID = _get_secret_or_env("CLIENT_ID")
# CLIENT_SECRET = _get_secret_or_env("CLIENT_SECRET")
#
# SP_HOSTNAME = _get_secret_or_env("SP_HOSTNAME")
# SP_SITE_PATH = _get_secret_or_env("SP_SITE_PATH")
# SP_EXCEL_PATH = _get_secret_or_env("SP_EXCEL_PATH")

TENANT_ID = "f73a06e7-4ed5-443b-b33c-3f6cd3b6ee9a"
CLIENT_ID = "774c15c8-b0ae-4d4d-ab35-05479ce84c94"
CLIENT_SECRET = "gnU8Q~DtHyDlc8yybbKVHzWG.chwQ8xHrYEijbym"

SP_HOSTNAME = "gestionintelcom.sharepoint.com"
SP_SITE_PATH = "/sites/Optimisation243"
SP_EXCEL_PATH = "/General/route_optimization_feedback.xlsx"
# ✅ Dossier SharePoint où stocker les images (optionnel)
SP_ATTACHMENTS_FOLDER = _get_secret_or_env("SP_ATTACHMENTS_FOLDER")

# ------------------------- Form options ------------------------- #
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
    "Routing": ["Unnecessary detour", "Loop / backtracking", "Stop order doesn’t make sense"],
    "Address / GPS location": ["Incorrect address", "Incorrect GPS pin", "Entrance/door is not at the right location"],
    "Access": ["Gated community / security", "Access code", "Drop-off location hard to find"],
    "Road conditions": ["Traffic", "Construction", "Weather"],
    "Parking": ["Hard to park", "Truck restrictions"],
    "Business / customer issue": ["Business closed", "Missing or inadequate instructions"],
    "Problems related to the app": ["Connection problems (uploading POD)", "Offline map issues"],
}

# ==========================================================
# Helpers
# ==========================================================
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


def make_submission_id() -> str:
    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    return f"rof_{ts}"


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


def mime_from_filename(name: str) -> str:
    ext = os.path.splitext(name.lower())[1]
    if ext == ".png":
        return "image/png"
    if ext in (".jpg", ".jpeg"):
        return "image/jpeg"
    return "application/octet-stream"


def _clean_for_name(x: str, max_len: int = 60) -> str:
    x = (x or "").strip().lower().replace(" ", "_")
    keep = []
    for ch in x:
        if ch.isalnum() or ch in ("_", "-", ".", "@"):
            keep.append(ch)
    return "".join(keep)[:max_len] if keep else "unknown"


# ✅ MODIF: nom de fichier basé sur driver_id + idc_id (plus de nom/prenom/email)
def build_sp_attachment_name(submission_id: str, driver_id: str, idc_id: str, original_filename: str) -> str:
    base = f"{submission_id}__driver_{_clean_for_name(driver_id)}__idc_{_clean_for_name(idc_id)}"
    orig = safe_filename(original_filename)
    return f"{base}__{orig}"


def guess_sp_attachments_folder() -> str:
    base = os.path.dirname(SP_EXCEL_PATH or "").replace("\\", "/")
    if not base.startswith("/"):
        base = "/" + base
    return f"{base}/ROF_Attachments"


# ==========================================================
# ✅ Microsoft Graph: Auth + Excel remote-only
# ==========================================================
def graph_is_configured() -> bool:
    required = [TENANT_ID, CLIENT_ID, CLIENT_SECRET, SP_HOSTNAME, SP_SITE_PATH, SP_EXCEL_PATH]
    return all(bool(x) for x in required)


def graph_get_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def graph_get_site_id(token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{SP_HOSTNAME}:{SP_SITE_PATH}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()["id"]


def graph_download_excel_bytes(token: str, site_id: str, sp_file_path: str) -> bytes | None:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:{sp_file_path}:/content"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code == 404:
        return None
    if not r.ok:
        raise RuntimeError(f"SharePoint download failed: {r.status_code} - {r.text}")
    return r.content


def graph_upload_excel_bytes(token: str, site_id: str, sp_file_path: str, content: bytes) -> None:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:{sp_file_path}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    r = requests.put(url, headers=headers, data=content, timeout=120)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"SharePoint upload failed: {r.status_code} - {r.text}")


def append_payload_to_remote_excel(payload: dict) -> None:
    if not graph_is_configured():
        raise RuntimeError("Graph is not configured (missing TENANT_ID/CLIENT_ID/CLIENT_SECRET/SP_* settings).")

    token = graph_get_token()
    site_id = graph_get_site_id(token)

    existing_bytes = graph_download_excel_bytes(token, site_id, SP_EXCEL_PATH)

    if existing_bytes:
        try:
            existing_df = pd.read_excel(BytesIO(existing_bytes))
        except Exception:
            existing_df = pd.DataFrame()
    else:
        existing_df = pd.DataFrame()

    new_row = pd.DataFrame([payload])
    all_cols = list(dict.fromkeys(list(existing_df.columns) + list(new_row.columns)))
    existing_df = existing_df.reindex(columns=all_cols)
    new_row = new_row.reindex(columns=all_cols)
    updated_df = pd.concat([existing_df, new_row], ignore_index=True)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        updated_df.to_excel(writer, index=False, sheet_name="Submissions")
    out.seek(0)

    graph_upload_excel_bytes(token, site_id, SP_EXCEL_PATH, out.getvalue())


# ==========================================================
# ✅ Upload images to SharePoint
# ==========================================================
def graph_upload_file_bytes(token: str, site_id: str, sp_file_path: str, content: bytes, content_type: str) -> None:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:{sp_file_path}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": content_type}
    r = requests.put(url, headers=headers, data=content, timeout=120)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"SharePoint upload failed: {r.status_code} - {r.text}")


# ✅ MODIF: signature = (images, submission_id, driver_id, idc_id)
def upload_images_to_sharepoint(images: list, submission_id: str, driver_id: str, idc_id: str) -> list[str]:
    if not images:
        return []

    if not graph_is_configured():
        raise RuntimeError("Graph is not configured; cannot upload images to SharePoint.")

    token = graph_get_token()
    site_id = graph_get_site_id(token)

    folder = (SP_ATTACHMENTS_FOLDER or guess_sp_attachments_folder()).replace("\\", "/")
    if not folder.startswith("/"):
        folder = "/" + folder

    sp_paths = []
    for f in images:
        if f is None:
            continue

        ext = os.path.splitext(f.name.lower())[1]
        if ext not in (".png", ".jpg", ".jpeg"):
            continue

        sp_name = build_sp_attachment_name(submission_id, driver_id, idc_id, f.name)
        sp_path = f"{folder}/{sp_name}".replace("//", "/")

        content = f.getvalue()
        ctype = mime_from_filename(f.name)
        graph_upload_file_bytes(token, site_id, sp_path, content, ctype)

        sp_paths.append(sp_path)

    return sp_paths


# ==========================================================
# Local helpers (lists)
# ==========================================================
@st.cache_data(show_spinner=False)
def load_station_list(path: str) -> list[str]:
    df = pd.read_excel(path)
    col = "Station Tag"
    if col not in df.columns:
        raise ValueError(f"Stations file must contain '{col}'")
    return sorted(df[col].astype(str).str.strip().replace("nan", "").unique().tolist())


def save_submission_json(payload: dict, submission_id: str) -> str:
    ensure_dirs()
    payload["submission_id"] = submission_id
    payload["submitted_at_utc"] = datetime.now(timezone.utc).isoformat()
    out_json = os.path.join(SUBMISSIONS_DIR, f"{submission_id}.json")
    with open(out_json, "w", encoding="utf-8") as w:
        json.dump(payload, w, ensure_ascii=False, indent=2)
    return out_json


def _on_main_issue_change():
    st.session_state["sub_issue"] = ""


# ✅ MODIF: reset keys (remove old identity/idc keys, add new ones)
def _reset_form():
    keys = [
        "driver_id",
        "idc_id",
        "idc_liaison",
        "station_select", "station_new",
        "route_number", "route_date",
        "vehicle_type",
        "issue_applies_to",
        "stop_number",
        "severity", "route_satisfaction",
        "time_lost", "parcel_tracking_id",
        "main_issue", "sub_issue",
        "what_happened", "what_should", "suggestion",
        "attachments", "agree",
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]


# ------------------------- Page config ------------------------- #
st.set_page_config(page_title="Route Optimization Feedback", page_icon="🧭", layout="centered")
st.logo(icon_image="logo_intelcom_2024.png", image="logo_intelcom_2024.png")

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
    station_list = load_station_list(STATIONS_FILE)
except Exception as e:
    st.error(
        "Unable to load the Excel lists.\n\n"
        f"- Expected file next to `app.py`: `{STATIONS_FILE}`\n"
        f"- Error: {e}"
    )
    st.stop()

# ------------------------- Title block ------------------------- #
banner_style = f"background-image: url('{banner_uri}');" if banner_uri else "background: #2b2b2b;"
st.markdown(
    f"""
<div class="card">
  <div class="title-banner" style="{banner_style}">
    <div class="title-overlay">
      <h1 class="forms-title">Route Optimization Feedback</h1>
      <div class="forms-subtitle">Share your feedback with us</div>
    </div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ------------------------- FORM UI ------------------------- #
st.subheader("1) Identification")

c1, c2 = st.columns(2)
with c1:
    # ✅ NEW
    driver_id = st.text_input("Driver ID *", placeholder="Enter your driver ID", key="driver_id")
    idc_liaison = st.selectbox("Who is your IDC Liaison? *", options=IDC_LIAISON_OPTIONS, index=0, key="idc_liaison")

with c2:
    # ✅ NEW
    idc_id = st.text_input("IDC ID *", placeholder="Enter your IDC ID", key="idc_id")

station_options = [""] + station_list + ["➕ Add a new Station"]
selected_station = st.selectbox("Station *", options=station_options, index=0, key="station_select")
station = st.text_input("Enter new Station code * (e.g., MONT)", placeholder="MONT", key="station_new") if selected_station == "➕ Add a new Station" else selected_station

col3, col4 = st.columns(2)
with col3:
    route_number = st.text_input("Route number * (digits only, e.g., 1235)", placeholder="1235", key="route_number")
    if route_number.strip() and (not route_number.strip().isdigit()):
        st.warning("Route number must contain digits only (e.g., 1235).")
with col4:
    route_date = st.date_input("Route date *", key="route_date")

vehicle_type = st.selectbox("What type of Vehicle do you have? *", options=VEHICLE_TYPE_OPTIONS, index=0, key="vehicle_type")

st.subheader("2) Problem scope")
issue_applies_to = st.radio(
    "This issue applies to: *",
    options=["Entire route", "Specific stop"],
    horizontal=True,
    index=0,
    key="issue_applies_to",
)

stop_number = ""
if issue_applies_to == "Specific stop":
    stop_number = st.text_input("Stop number *", placeholder="e.g., 12", key="stop_number")

severity = st.radio("Severity *", options=["Low", "Medium", "High", "Critical"], horizontal=True, index=1, key="severity")
route_satisfaction = st.select_slider("Route satisfaction / rating (optional)", options=["0", "1", "2", "3", "4", "5"], value="0", key="route_satisfaction")
time_lost = st.radio("Estimated time lost *", options=["0–15 min", "15–30 min", "30–60 min", "60+ min"], horizontal=True, index=0, key="time_lost")

parcel_tracking_id = st.text_input("Parcel Tracking ID", placeholder="Example: INTELXXX0001111", key="parcel_tracking_id")

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
)

st.subheader("4) Details")
what_happened = st.text_area("What happened? (optional)", height=110, key="what_happened")
what_should = st.text_area("What should have happened ideally? (optional)", height=90, key="what_should")
suggestion = st.text_area("Suggestion (optional)", height=90, key="suggestion")

st.subheader("5) Attachments")
attachments = st.file_uploader(
    "Attach images only (png, jpg, jpeg)",
    accept_multiple_files=True,
    type=["png", "jpg", "jpeg"],
    key="attachments",
)

agree = st.checkbox("I confirm this information is accurate to the best of my knowledge *", key="agree")
submitted = st.button("Submit feedback", type="primary")

# ------------------------- Submit handler ------------------------- #
if submitted:
    errors = []

    # ✅ NEW fields
    driver_id_clean = (driver_id or "").strip()
    idc_id_clean = (idc_id or "").strip()

    station_clean = (station or "").strip()
    route_number_clean = (route_number or "").strip()
    idc_liaison_clean = (idc_liaison or "").strip()
    vehicle_type_clean = (vehicle_type or "").strip()
    parcel_tracking_id_clean = (parcel_tracking_id or "").strip()
    stop_number_clean = (stop_number or "").strip()

    # ✅ Validation (Identification) - NEW
    if not driver_id_clean:
        errors.append("Driver ID is required.")
    if not idc_id_clean:
        errors.append("IDC ID is required.")

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

    if issue_applies_to == "Specific stop":
        if not stop_number_clean:
            errors.append("Stop number is required when 'Specific stop' is selected.")
        if not parcel_tracking_id_clean:
            errors.append("Parcel Tracking ID is required when 'Specific stop' is selected.")

    if not time_lost:
        errors.append("Estimated time lost is required.")
    if not main_issue:
        errors.append("Main issue category is required.")
    if main_issue and not sub_issue:
        errors.append("Sub-category is required.")
    if not agree:
        errors.append("You must confirm the accuracy checkbox.")

    for f in attachments or []:
        ext = os.path.splitext(f.name.lower())[1]
        if ext not in (".png", ".jpg", ".jpeg"):
            errors.append(f"Only images are allowed. Invalid file: {f.name}")

    if errors:
        st.error("Please fix the following:\n- " + "\n- ".join(errors))
        st.stop()

    submission_id = make_submission_id()

    payload = {
        # ✅ NEW identification fields
        "driver_id": driver_id_clean,
        "idc_id": idc_id_clean,

        "idc_liaison": idc_liaison_clean,
        "station": station_clean,
        "route_number": route_number_clean,
        "route_date": route_date.isoformat(),
        "vehicle_type": vehicle_type_clean,
        "parcel_tracking_id": parcel_tracking_id_clean if parcel_tracking_id_clean else None,

        "issue_applies_to": issue_applies_to,
        "stop_number": stop_number_clean if issue_applies_to == "Specific stop" else None,

        "severity": severity,
        "route_satisfaction": route_satisfaction,
        "estimated_time_lost": time_lost,

        "main_issue_category": main_issue,
        "sub_category": sub_issue,

        "what_happened": (what_happened or "").strip(),
        "what_should_have_happened": (what_should or "").strip(),
        "suggestion": (suggestion or "").strip(),
    }

    # 1) Upload images to SharePoint (now uses driver_id/idc_id in file name)
    sp_files_ok = False
    sp_files_err = None
    sp_attachment_paths = []

    try:
        sp_attachment_paths = upload_images_to_sharepoint(
            images=attachments,
            submission_id=submission_id,
            driver_id=driver_id_clean,
            idc_id=idc_id_clean,
        )
        sp_files_ok = True
    except Exception as e:
        sp_files_err = str(e)

    payload["sp_attachments"] = sp_attachment_paths

    # 2) Save JSON locally (optional)
    out_json = save_submission_json(payload, submission_id=submission_id)

    # 3) Update SharePoint Excel (same function as before)
    sp_excel_ok = False
    sp_excel_err = None
    try:
        append_payload_to_remote_excel(payload)
        sp_excel_ok = True
    except Exception as e:
        sp_excel_err = str(e)

    if sp_files_ok and sp_excel_ok:
        st.success("✅ Feedback submitted! Excel updated and images uploaded to SharePoint.")
        st.caption(f"Excel: {SP_EXCEL_PATH}")
        st.caption(f"Images folder: {(SP_ATTACHMENTS_FOLDER or guess_sp_attachments_folder())}")
    elif (not sp_files_ok) and sp_excel_ok:
        st.warning("⚠️ Excel updated, but image upload failed.")
        st.caption(f"Images error: {sp_files_err}")
    elif sp_files_ok and (not sp_excel_ok):
        st.warning("⚠️ Images uploaded, but Excel update failed.")
        st.caption(f"Excel error: {sp_excel_err}")
    else:
        st.warning("⚠️ SharePoint update failed (Excel + images). Saved JSON locally.")
        st.caption(f"JSON: {out_json}")
        st.caption(f"Images error: {sp_files_err}")
        st.caption(f"Excel error: {sp_excel_err}")

    _reset_form()
    st.rerun()
