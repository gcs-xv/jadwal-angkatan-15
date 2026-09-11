import io
import json
import random
import re
import csv
import calendar
import uuid
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from difflib import get_close_matches
from urllib.parse import urlparse

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

try:
    from supabase import create_client
except ImportError:  # Lets local development open before `supabase` is installed.
    create_client = None

st.set_page_config(page_title="Jadwal Angkatan 15", page_icon="✦", layout="wide")
ROSTER = {"Aliyah":"F", "Soma":"M", "Syamsul":"M", "Ferrel":"M", "Kezia":"F", "Alam":"M", "Bagus":"M", "Retno":"F", "Rachel":"F", "Irpan":"M", "Farez":"M"}
NAMES = list(ROSTER)
DAY_NAMES = ("Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu")
MONTHS = ("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
WEEKDAY_WORDS = {"senin", "selasa", "rabu", "kamis", "jumat", "sabtu", "minggu"}
DOCTOR_WORDS = {"drg", "dr", "sp", "mmf", "mf", "subsp", "comf", "tr", "tm", "mars", "ph", "d", *WEEKDAY_WORDS}


def default_cohort_config():
    """Neutral source rows. Admin names them after seeing the pasted roster."""
    return pd.DataFrame([
        {"Kolom": "kelompok_1", "Label angkatan": "Kelompok 1", "Residen per hari": 1, "Aktif": True},
        {"Kolom": "kelompok_2", "Label angkatan": "Kelompok 2", "Residen per hari": 1, "Aktif": True},
        {"Kolom": "kelompok_3", "Label angkatan": "Kelompok 3", "Residen per hari": 1, "Aktif": True},
        {"Kolom": "kelompok_4", "Label angkatan": "Kelompok 4", "Residen per hari": 2, "Aktif": True},
        {"Kolom": "kelompok_5", "Label angkatan": "Kelompok 5", "Residen per hari": 3, "Aktif": True},
        {"Kolom": "kelompok_6", "Label angkatan": "Kelompok 6", "Residen per hari": 4, "Aktif": True},
        {"Kolom": "kelompok_7", "Label angkatan": "Kelompok 7", "Residen per hari": 5, "Aktif": True},
        {"Kolom": "kelompok_8", "Label angkatan": "Kelompok 8", "Residen per hari": 5, "Aktif": True},
    ])


def roster_words(value):
    return re.findall(r"[A-Za-zÀ-ÿ]+(?:[-'][A-Za-zÀ-ÿ]+)?", str(value))


def clean_sentence(value):
    return re.sub(r"\s+", " ", str(value)).strip(" ,;")


def split_roster_blocks(text):
    """Split a Word/Google Docs table copy by its repeated weekday header."""
    header = re.compile(r"(?m)^(?=(?:(?:Senin|Selasa|Rabu|Kamis|Jumat|Sabtu|Minggu)(?:\s+|$))+$)")
    points = [match.start() for match in header.finditer(text)]
    if not points:
        return [text]
    return [text[start:end] for start, end in zip(points, points[1:] + [len(text)])]


def parse_tabular_roster(text, config):
    """Read a TSV/CSV table copied directly from a spreadsheet or CSV preview."""
    delimiter = "\t" if "\t" in text else ";"
    table = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    active = config.copy()
    active["Kolom"] = active["Kolom"].astype(str).str.strip().str.lower()
    active = active[(active["Aktif"] == True) & active["Kolom"].ne("")]
    definitions = list(active["Kolom"])
    rows, warnings = [], []
    for line_number, cells in enumerate(table, start=1):
        cells = [cell.strip() for cell in cells]
        if not any(cells) or (cells and cells[0].lower() == "month"):
            continue
        if len(cells) < 3 + len(definitions):
            warnings.append(f"Baris {line_number} tidak dipetakan karena hanya memiliki {len(cells)} kolom; konfigurasi membutuhkan minimal {3 + len(definitions)}.")
            continue
        try:
            parsed_date = datetime.strptime(cells[1], "%Y-%m-%d").date()
        except ValueError:
            try:
                parsed_date = datetime.strptime(cells[1], "%d/%m/%Y").date()
            except ValueError:
                warnings.append(f"Baris {line_number} tidak dipetakan karena tanggal `{cells[1]}` tidak dikenali.")
                continue
        item = {"Tanggal": parsed_date.isoformat(), "DPJP": cells[2]}
        for index, code in enumerate(definitions, start=3):
            item[code] = re.sub(r"\s*\|\s*", ", ", cells[index]).strip()
        rows.append(item)
    return pd.DataFrame(rows), list(dict.fromkeys(warnings)), []


def parse_pasted_roster(text, config):
    """Parse a visual table pasted as text without requiring a CSV upload.

    Word often wraps cells at different points. The parser first reads the weekly
    date block and then reconstructs each cohort from its expected member count.
    It also learns cohort membership from clean blocks, allowing it to recover
    rows whose columns were interleaved during copying.
    """
    non_empty_lines = [line for line in text.splitlines() if line.strip()]
    if len(non_empty_lines) >= 2 and ("\t" in non_empty_lines[0] or non_empty_lines[0].count(";") >= 3):
        return parse_tabular_roster(text, config)

    active = config.copy()
    active["Kolom"] = active["Kolom"].astype(str).str.strip().str.lower()
    active["Residen per hari"] = pd.to_numeric(active["Residen per hari"], errors="coerce").fillna(0).astype(int)
    active = active[(active["Aktif"] == True) & active["Kolom"].ne("") & (active["Residen per hari"] > 0)]
    definitions = [(row["Kolom"], int(row["Residen per hari"])) for _, row in active.iterrows()]
    if not definitions:
        return pd.DataFrame(), ["Tambahkan minimal satu angkatan aktif dengan jumlah residen per hari lebih dari nol."], []

    blocks, warnings, skipped = [], [], []
    for block_number, block in enumerate(split_roster_blocks(text), start=1):
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        dates = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", block)
        if not dates:
            continue
        column_count = len(dates)
        date_line = max(index for index, line in enumerate(lines) if re.search(r"\d{2}/\d{2}/\d{4}", line))
        operational_start = None
        for index, line in enumerate(lines[date_line + 1:], start=date_line + 1):
            words = roster_words(line)
            if len(words) == column_count and all(word.lower() not in DOCTOR_WORDS for word in words):
                operational_start = index
                break
        if operational_start is None:
            skipped.extend(dates)
            warnings.append(f"Blok {block_number}: ditemukan {column_count} tanggal, tetapi baris Pilot tidak memiliki {column_count} nama. Blok tidak dipetakan agar tidak salah.")
            continue

        stream = roster_words(" ".join(lines[operational_start:]))
        required = column_count * sum(size for _, size in definitions)
        if len(stream) < required:
            skipped.extend(dates)
            warnings.append(f"Blok {block_number}: roster hanya memiliki {len(stream)} nama, sedangkan konfigurasi membutuhkan {required}. Blok tidak dipetakan.")
            continue

        doctors_text = " ".join(lines[date_line + 1:operational_start])
        doctor_parts = [clean_sentence(item) for item in re.split(r"(?i)(?=(?:dr\.\s*)?drg\.)", doctors_text) if re.search(r"(?i)(?:dr\.\s*)?drg\.", item)]
        provisional, position = {}, 0
        for code, size in definitions:
            provisional[code] = stream[position:position + column_count * size]
            position += column_count * size
        blocks.append({
            "dates": dates,
            "count": column_count,
            "tail": stream,
            "provisional": provisional,
            "doctors": doctor_parts,
        })

    # Establish cohort identity from the first positional pass. This corrects the
    # common Word-copy issue where the A14/A15 visual cells appear interleaved.
    votes = defaultdict(Counter)
    display_names = {}
    for block in blocks:
        for code, people in block["provisional"].items():
            for person in people:
                key = person.lower()
                votes[key][code] += 1
                display_names.setdefault(key, person)
    owner = {person: counts.most_common(1)[0][0] for person, counts in votes.items()}
    known_names = list(display_names)

    rows = []
    for block in blocks:
        count = block["count"]
        corrected_tail = []
        for person in block["tail"]:
            key = person.lower()
            if key not in owner:
                close = get_close_matches(key, known_names, n=1, cutoff=.83)
                if close:
                    warnings.append(f"Nama `{person}` dibaca sebagai `{display_names[close[0]]}` karena ejaan sangat mirip.")
                    person, key = display_names[close[0]], close[0]
            corrected_tail.append(person)
        groups = {}
        for code, size in definitions:
            recovered = [person for person in corrected_tail if owner.get(person.lower()) == code]
            expected = count * size
            groups[code] = recovered if len(recovered) == expected else block["provisional"][code]
            if len(recovered) not in (0, expected):
                warnings.append(f"{block['dates'][0]}: {code} terbaca {len(recovered)}/{expected} nama setelah pemulihan; gunakan hasil preview untuk koreksi.")
        for index, raw_date in enumerate(block["dates"]):
            item = {
                "Tanggal": datetime.strptime(raw_date, "%d/%m/%Y").date().isoformat(),
                "DPJP": block["doctors"][index] if len(block["doctors"]) == count else "",
            }
            for code, size in definitions:
                people = groups[code][index * size:(index + 1) * size]
                item[code] = ", ".join(people)
            rows.append(item)
        if len(block["doctors"]) != count:
            warnings.append(f"{block['dates'][0]}: DPJP tidak dapat dipisahkan otomatis. Isi kolom DPJP pada tabel preview bila diperlukan.")

    if skipped:
        pretty = ", ".join(datetime.strptime(value, "%d/%m/%Y").strftime("%d %b") for value in skipped)
        warnings.append(f"Tanggal yang belum masuk preview: {pretty}.")
    return pd.DataFrame(rows), list(dict.fromkeys(warnings)), skipped


def app_secret(name):
    try:
        return st.secrets[name]
    except (KeyError, FileNotFoundError):
        return None


@st.cache_resource(show_spinner=False)
def get_supabase_client(url, secret_key):
    return create_client(url, secret_key)


def database():
    url, secret_key = app_secret("SUPABASE_URL"), app_secret("SUPABASE_SECRET_KEY")
    if create_client is None:
        return None, "Library Supabase belum terpasang. Jalankan ulang deployment setelah requirements diperbarui."
    if not url or not secret_key:
        return None, "Supabase belum dikonfigurasi di Streamlit Secrets."
    parsed_url = urlparse(str(url))
    if parsed_url.scheme != "https" or not parsed_url.hostname or not parsed_url.hostname.endswith(".supabase.co") or parsed_url.path not in ("", "/"):
        return None, "SUPABASE_URL harus berupa Project URL API seperti `https://abcdefgh.supabase.co`, bukan URL dashboard Supabase."
    try:
        return get_supabase_client(url, secret_key), None
    except Exception as error:
        return None, f"Koneksi Supabase tidak dapat dibuat: {error}"


def json_records(frame):
    """Convert a DataFrame to JSON-safe records for a jsonb column."""
    return json.loads(frame.to_json(orient="records"))


def load_monthly_roster(month_key):
    client, error = database()
    if error:
        return None, error
    try:
        response = client.table("monthly_rosters").select("roster, cohorts, updated_at").eq("roster_month", month_key).execute()
        data = response.data or []
        return (data[0] if data else None), None
    except Exception as error:
        return None, f"Roster bulan ini belum dapat dibaca: {error}"


def save_monthly_roster(month_key, roster, config):
    client, error = database()
    if error:
        return error
    try:
        client.table("monthly_rosters").upsert({
            "roster_month": month_key,
            "roster": json_records(roster),
            "cohorts": json_records(config),
            "updated_at": datetime.utcnow().isoformat(),
        }, on_conflict="roster_month").execute()
        return None
    except Exception as error:
        message = str(error)
        if "JSON could not be generated" in message or "404" in message:
            return "Supabase mengembalikan 404. Periksa SUPABASE_URL: gunakan Project URL API `https://<project-ref>.supabase.co`, bukan tautan dashboard."
        return f"Roster belum tersimpan: {message}"


def load_daily_assignment(assignment_date):
    client, error = database()
    if error:
        return None, error
    try:
        response = client.table("daily_assignments").select("assignment, assignment_text, updated_at").eq("assignment_date", assignment_date).execute()
        data = response.data or []
        return (data[0] if data else None), None
    except Exception as error:
        message = str(error)
        if "daily_assignments" in message or "404" in message:
            return None, "Tabel pembagian belum dibuat di Supabase. Jalankan SQL `daily_assignments` yang disediakan bersama update ini."
        return None, f"Pembagian belum dapat dibaca: {message}"


def save_daily_assignment(month_key, assignment_date, assignment, assignment_text):
    client, error = database()
    if error:
        return error
    try:
        client.table("daily_assignments").upsert({
            "assignment_date": assignment_date,
            "roster_month": month_key,
            "assignment": assignment,
            "assignment_text": assignment_text,
            "updated_at": datetime.utcnow().isoformat(),
        }, on_conflict="assignment_date").execute()
        return None
    except Exception as error:
        return f"Pembagian belum tersimpan: {error}"


def reset_month_state(month_key):
    stored, error = load_monthly_roster(month_key)
    st.session_state.active_roster_month = month_key
    st.session_state.roster_warnings = []
    st.session_state.roster_skipped = []
    st.session_state.pasted_roster = ""
    st.session_state.parsed_roster = pd.DataFrame(stored["roster"]) if stored and stored.get("roster") else None
    st.session_state.cohort_config = pd.DataFrame(stored["cohorts"]) if stored and stored.get("cohorts") else default_cohort_config()
    for key in ("cohort_config_editor", "parsed_roster_editor", "pasted_roster_input"):
        st.session_state.pop(key, None)
    return stored, error


def admin_access():
    if st.session_state.get("roster_admin", False):
        left, right = st.columns([4, 1])
        left.success("Mode admin aktif. Perubahan roster akan disimpan untuk semua pengguna.")
        if right.button("Keluar admin", use_container_width=True, key="admin_logout"):
            st.session_state.roster_admin = False
            st.rerun()
        return True
    with st.expander("Akses admin"):
        password = st.text_input("Password admin", type="password", key="admin_password")
        if st.button("Masuk sebagai admin", key="admin_login"):
            expected = app_secret("ADMIN_PASSWORD")
            if expected and password == expected:
                st.session_state.roster_admin = True
                st.rerun()
            elif not expected:
                st.error("ADMIN_PASSWORD belum diisi di Streamlit Secrets.")
            else:
                st.error("Password admin salah.")
    return False


def split_residents(value):
    return [name.strip() for name in re.split(r"[,|\n]", str(value or "")) if name.strip()]


def active_daily_roster(row, config):
    result = {}
    for _, cohort in config.iterrows():
        code = str(cohort.get("Kolom", "")).strip().lower()
        if code and bool(cohort.get("Aktif", True)):
            result[code] = split_residents(row.get(code, ""))
    return result


def stable_shuffle(values, assignment_date, salt):
    values = list(values)
    random.Random(f"{assignment_date}:{salt}").shuffle(values)
    return values


def distribute_patients(residents, count, assignment_date, salt):
    teams = [[] for _ in range(count)]
    if not residents or not count:
        return teams
    for index, name in enumerate(stable_shuffle(residents, assignment_date, salt)):
        teams[index % count].append(name)
    return teams


def distribute_roles(residents, patient_count, roles, assignment_date, salt):
    result = [{role: [] for role in roles} for _ in range(patient_count)]
    if not residents or not patient_count:
        return result
    slots = [(patient_index, role) for patient_index in range(patient_count) for role in roles]
    slots = stable_shuffle(slots, assignment_date, f"{salt}:slots")
    people = stable_shuffle(residents, assignment_date, f"{salt}:people")
    if len(people) >= len(slots):
        slots = (slots * (len(people) // len(slots))) + slots[:len(people) % len(slots)]
    for index, slot in enumerate(slots):
        patient_index, role = slot
        result[patient_index][role].append(people[index % len(people)])
    return result


def parse_patient_lines(value, post_op=False):
    patients = []
    for line in str(value or "").splitlines():
        line = line.strip()
        if not line:
            continue
        name, separator, meta = line.partition("|")
        patients.append({"name": name.strip(), "meta": meta.strip() if separator and post_op else ""})
    return patients


def patients_from_table(frame, post_op=False):
    patients = []
    for _, row in frame.iterrows():
        name = str(row.get("Pasien", "")).strip()
        if name:
            patients.append({"name": name, "meta": str(row.get("POD awal", "")).strip() if post_op else ""})
    return patients


def patients_from_entries(entries, post_op=False):
    patients = []
    for entry in entries:
        name = str(entry.get("name", "")).strip()
        case = str(entry.get("case", "")).strip()
        if not name:
            continue
        display = f"{name} ({case})" if case else name
        patients.append({"name": display, "meta": str(entry.get("pod", "")).strip() if post_op else ""})
    return patients


def empty_patient_entry(post_op=False):
    return {"id": uuid.uuid4().hex, "name": "", "case": "", "pod": "POD 0" if post_op else ""}


def pod_labels(meta):
    meta = re.sub(r"\s+", " ", (meta or "").strip())
    if meta:
        numeric = re.search(r"(?i)POD\s*(\d+)\s*$", meta)
        roman = re.search(r"(?i)POD\s*(I|II|III|IV|V|VI|VII|VIII|IX|X)\s*$", meta)
        if numeric:
            current = int(numeric.group(1))
            roman_values = ("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X")
            following = roman_values[current] if current < len(roman_values) else str(current + 1)
            return (f"POD {current}", f"POD {following}")
        if roman:
            roman_values = ("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X")
            current = roman.group(1).upper()
            index = roman_values.index(current)
            following = roman_values[index + 1] if index + 1 < len(roman_values) else str(index + 2)
            return (f"POD {current}", f"POD {following}")
        return (meta, "POD I")
    return ("POD I", "POD II")


def cohort_label_map(config):
    return {str(row["Kolom"]).lower(): str(row["Label angkatan"]) for _, row in config.iterrows()}


def sort_names_by_cohort(names, roster, cohort_codes):
    order = {code: index for index, code in enumerate(cohort_codes)}
    membership = {name: code for code, members in roster.items() for name in members}
    return sorted(dict.fromkeys(names), key=lambda name: (order.get(membership.get(name, ""), 999), name.lower()))


def build_daily_assignment(roster, assignment_date, post_ops, pre_ops, igds, pilot, copilot, erm, review):
    """Generic version of the legacy cohort-by-cohort assignment theorem."""
    # Pilot coordinates the day and is not assigned again to Post-op/Pre-op/IGD.
    roster = {code: [name for name in members if name != pilot] for code, members in roster.items()}
    codes = list(roster)
    post_assignment = []
    if post_ops:
        # Each cohort is allocated independently to every patient-POD slot.
        # This prevents a small cohort from disappearing from some patients and
        # preserves fairness inside the cohort rather than across cohorts.
        teams = [[[], []] for _ in post_ops]
        for code in codes:
            slots = [(patient_index, pod_index) for patient_index in range(len(post_ops)) for pod_index in range(2)]
            members = stable_shuffle(roster[code], assignment_date, f"post:{code}:people")
            if not members:
                continue
            target_slots = max(len(slots), len(members))
            expanded_slots = (slots * (target_slots // len(slots))) + slots[:target_slots % len(slots)]
            expanded_slots = stable_shuffle(expanded_slots, assignment_date, f"post:{code}:slots")
            for index, slot in enumerate(expanded_slots):
                patient_index, pod_index = slot
                teams[patient_index][pod_index].append(members[index % len(members)])
        for index, patient in enumerate(post_ops):
            labels = pod_labels(patient.get("meta"))
            first_team = sort_names_by_cohort(teams[index][0], roster, codes)
            second_team = sort_names_by_cohort(teams[index][1], roster, codes)
            post_assignment.append({"name": patient["name"], "pod_lines": [{"label": labels[0], "team": first_team}, {"label": labels[1], "team": second_team}]})

    def build_role_section(patients, final_role, salt):
        roles = ["soap", "rm", "erm", final_role]
        section = [{"name": patient["name"], "soap": [], "rm_erm": [], final_role: []} for patient in patients]
        for code in codes:
            allocation = distribute_roles(roster[code], len(patients), roles, assignment_date, f"{salt}:{code}")
            for index, roles_for_patient in enumerate(allocation):
                for role, names in roles_for_patient.items():
                    target_role = "rm_erm" if role in ("rm", "erm") else role
                    section[index][target_role].extend(names)
        for entry in section:
            for role in ("soap", "rm_erm", final_role):
                entry[role] = sort_names_by_cohort(entry[role], roster, codes)
        return section

    pre_assignment = build_role_section(pre_ops, "tsr", "pre")
    igd_assignment = build_role_section(igds, "er", "igd")
    return {
        "date": assignment_date,
        "day_name": DAY_NAMES[datetime.strptime(assignment_date, "%Y-%m-%d").weekday()],
        "pilot": pilot,
        "copilot": copilot,
        "erm_manual": erm,
        "review_manual": review,
        "post_op": post_assignment,
        "pre_op": pre_assignment,
        "igd": igd_assignment,
    }


def assignment_text(assignment, labels):
    current = datetime.strptime(assignment["date"], "%Y-%m-%d")
    lines = [f"Pembagian tugas jaga {assignment['day_name']}, {current:%d/%m/%Y}", "", f"Pilot : {assignment['pilot']}", f"Co Pilot : {assignment['copilot']}", ""]
    if assignment["post_op"]:
        lines.append(f"*{len(assignment['post_op'])} Post Op*")
        for index, patient in enumerate(assignment["post_op"], start=1):
            lines.append(f"{index}. {patient['name']}")
            for pod in patient["pod_lines"]:
                lines.append(f"   {pod['label']} : {', '.join(pod['team'])}")
        lines.append("")
    for title, section, last_role in (("PRE-OP", assignment["pre_op"], "TSR"), ("IGD", assignment["igd"], "ER")):
        if section:
            lines.append(title)
            for index, patient in enumerate(section, start=1):
                lines += [f"{index}. {patient['name']}", f"   SOAP : {', '.join(patient['soap'])}", f"   RM/ERM : {', '.join(patient['rm_erm'])}", f"   {last_role} : {', '.join(patient[last_role.lower()])}"]
            lines.append("")
    lines += [f"ERM : {assignment['erm_manual']}", f"Review : {assignment['review_manual']}"]
    return "\n".join(lines)


def date_button_grid(available_dates):
    available = {str(value) for value in available_dates}
    reference = datetime.strptime(sorted(available)[0], "%Y-%m-%d").date()
    current = st.session_state.get("selected_assignment_date")
    if current not in available:
        current = sorted(available)[0]
        st.session_state.selected_assignment_date = current
    st.markdown("<div class='panel'><b>Pilih tanggal pembagian</b><br><span style='color:#60717d'>Klik tanggal dengan roster yang tersedia.</span></div>", unsafe_allow_html=True)
    for column, name in zip(st.columns(7), DAY_NAMES):
        column.caption(name[:3])
    for week in calendar.monthcalendar(reference.year, reference.month):
        columns = st.columns(7)
        for weekday, day_number in enumerate(week):
            if not day_number:
                columns[weekday].write("")
                continue
            iso_date = date(reference.year, reference.month, day_number).isoformat()
            if iso_date in available:
                active = iso_date == current
                if columns[weekday].button(str(day_number), key=f"assignment_day_{iso_date}", type="primary" if active else "secondary", use_container_width=True):
                    st.session_state.selected_assignment_date = iso_date
                    st.rerun()
            else:
                columns[weekday].button(str(day_number), key=f"missing_day_{iso_date}", disabled=True, use_container_width=True)
    return st.session_state.selected_assignment_date


def render_assignment_workspace(parsed, config, month_key, is_admin):
    available_dates = parsed["Tanggal"].dropna().astype(str).tolist()
    if not available_dates:
        return
    selected_date = date_button_grid(available_dates)
    row = parsed.loc[parsed["Tanggal"].astype(str) == selected_date].iloc[0]
    roster = active_daily_roster(row, config)
    labels = cohort_label_map(config)
    all_names = [name for code in roster for name in roster[code]]
    all_names = list(dict.fromkeys(all_names))
    readable_date = datetime.strptime(selected_date, "%Y-%m-%d").strftime("%d %B %Y")
    st.markdown(f"## Pembagian Tanggal {readable_date}")
    st.caption(f"DPJP: {row.get('DPJP', '-') or '-'} · {len(all_names)} residen tersedia")

    saved, database_error = load_daily_assignment(selected_date)
    if database_error:
        st.warning(database_error)
    if saved:
        st.markdown("<div class='panel'><b>Pembagian tersimpan</b><br><span style='color:#60717d'>Versi ini dapat diakses kembali setiap kali tanggal tersebut dibuka.</span></div>", unsafe_allow_html=True)
        st.code(saved.get("assignment_text", ""), language=None)
    elif not is_admin:
        st.info("Belum ada pembagian tersimpan untuk tanggal ini. Isi pasien di bawah untuk membuat pembagian pertama.")
    st.markdown("<div class='panel'><b>Buat atau bagi ulang</b><br><span style='color:#60717d'>Algoritme membagi setiap angkatan secara proporsional pada Post-op, Pre-op, dan IGD.</span></div>", unsafe_allow_html=True)
    default_person = all_names[0] if all_names else ""
    post_state = f"post_patient_entries_{selected_date}"
    pre_state = f"pre_patient_entries_{selected_date}"
    igd_state = f"igd_patient_entries_{selected_date}"
    st.session_state.setdefault(post_state, [empty_patient_entry(post_op=True)])
    st.session_state.setdefault(pre_state, [empty_patient_entry()])
    st.session_state.setdefault(igd_state, [empty_patient_entry()])
    post_entries, pre_entries, igd_entries = st.session_state[post_state], st.session_state[pre_state], st.session_state[igd_state]
    with st.form(f"assignment_setup_form_{selected_date}", border=False):
        one, two, three, four = st.columns(4)
        with one:
            pilot = st.selectbox("Pilot", ["", *all_names], index=1 if default_person else 0, key=f"pilot_{selected_date}")
        with two:
            copilot = st.selectbox("Co-pilot", ["", *all_names], index=2 if len(all_names) > 1 else 0, key=f"copilot_{selected_date}")
        with three:
            erm = st.selectbox("ERM", ["", *all_names], key=f"erm_{selected_date}")
        with four:
            review = st.selectbox("Review", ["", *all_names], key=f"review_{selected_date}")
        post_col, pre_col, igd_col = st.columns(3)
        current_post, current_pre, current_igd = [], [], []
        remove_post = remove_pre = remove_igd = None
        with post_col:
            st.caption("Post-op — tambah pasien dengan tombol +")
            post_table = st.data_editor(pd.DataFrame([{"Pasien": "", "POD awal": "POD 0"}]), num_rows="dynamic", hide_index=True, use_container_width=True, height=180, key=f"post_table_{selected_date}")
            st.markdown("#### Post-op")
            st.caption("Nama pasien, kasus, dan POD awal")
            for index, entry in enumerate(post_entries):
                row_id = entry["id"]
                st.caption(f"Pasien {index + 1}")
                name_col, case_col = st.columns([1, 1])
                name = name_col.text_input("Nama pasien", value=entry["name"], key=f"post_name_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Nama pasien")
                case = case_col.text_input("Kasus", value=entry["case"], key=f"post_case_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Kasus")
                pod_col, delete_col = st.columns([3, 1])
                pod = pod_col.selectbox("POD awal", ["POD 0", "POD I", "POD II", "POD III"], index=["POD 0", "POD I", "POD II", "POD III"].index(entry.get("pod", "POD 0")) if entry.get("pod", "POD 0") in ["POD 0", "POD I", "POD II", "POD III"] else 0, key=f"post_pod_{selected_date}_{row_id}", label_visibility="collapsed")
                if delete_col.form_submit_button("Hapus", key=f"remove_post_{selected_date}_{row_id}", use_container_width=True):
                    remove_post = index
                current_post.append({"id": row_id, "name": name, "case": case, "pod": pod})
            add_post = st.form_submit_button("+ Tambah pasien Post-op", key=f"add_post_{selected_date}", use_container_width=True)
        with pre_col:
            st.caption("Pre-op — tambah pasien dengan tombol +")
            pre_table = st.data_editor(pd.DataFrame([{"Pasien": ""}]), num_rows="dynamic", hide_index=True, use_container_width=True, height=180, key=f"pre_table_{selected_date}")
            st.markdown("#### Pre-op")
            st.caption("Nama pasien dan kasus")
            for index, entry in enumerate(pre_entries):
                row_id = entry["id"]
                st.caption(f"Pasien {index + 1}")
                name_col, case_col, delete_col = st.columns([1, 1, .55])
                name = name_col.text_input("Nama pasien", value=entry["name"], key=f"pre_name_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Nama pasien")
                case = case_col.text_input("Kasus", value=entry["case"], key=f"pre_case_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Kasus")
                if delete_col.form_submit_button("Hapus", key=f"remove_pre_{selected_date}_{row_id}", use_container_width=True):
                    remove_pre = index
                current_pre.append({"id": row_id, "name": name, "case": case})
            add_pre = st.form_submit_button("+ Tambah pasien Pre-op", key=f"add_pre_{selected_date}", use_container_width=True)
        with igd_col:
            st.caption("IGD — tambah pasien dengan tombol +")
            igd_table = st.data_editor(pd.DataFrame([{"Pasien": ""}]), num_rows="dynamic", hide_index=True, use_container_width=True, height=180, key=f"igd_table_{selected_date}")
            st.markdown("#### IGD")
            st.caption("Nama pasien dan kasus")
            for index, entry in enumerate(igd_entries):
                row_id = entry["id"]
                st.caption(f"Pasien {index + 1}")
                name_col, case_col, delete_col = st.columns([1, 1, .55])
                name = name_col.text_input("Nama pasien", value=entry["name"], key=f"igd_name_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Nama pasien")
                case = case_col.text_input("Kasus", value=entry["case"], key=f"igd_case_{selected_date}_{row_id}", label_visibility="collapsed", placeholder="Kasus")
                if delete_col.form_submit_button("Hapus", key=f"remove_igd_{selected_date}_{row_id}", use_container_width=True):
                    remove_igd = index
                current_igd.append({"id": row_id, "name": name, "case": case})
            add_igd = st.form_submit_button("+ Tambah pasien IGD", key=f"add_igd_{selected_date}", use_container_width=True)
        generate = st.form_submit_button("Buat pembagian otomatis", type="primary", use_container_width=True)
    if remove_post is not None:
        st.session_state[post_state] = [entry for index, entry in enumerate(current_post) if index != remove_post]
        st.rerun()
    if remove_pre is not None:
        st.session_state[pre_state] = [entry for index, entry in enumerate(current_pre) if index != remove_pre]
        st.rerun()
    if remove_igd is not None:
        st.session_state[igd_state] = [entry for index, entry in enumerate(current_igd) if index != remove_igd]
        st.rerun()
    if add_post:
        st.session_state[post_state] = current_post + [empty_patient_entry(post_op=True)]
        st.rerun()
    if add_pre:
        st.session_state[pre_state] = current_pre + [empty_patient_entry()]
        st.rerun()
    if add_igd:
        st.session_state[igd_state] = current_igd + [empty_patient_entry()]
        st.rerun()
    if generate:
        st.session_state[post_state], st.session_state[pre_state], st.session_state[igd_state] = current_post, current_pre, current_igd
        assignment = build_daily_assignment(
            roster, selected_date, patients_from_table(post_table, post_op=True), patients_from_table(pre_table), patients_from_table(igd_table), pilot, copilot, erm, review,
            roster, selected_date, patients_from_entries(current_post, post_op=True), patients_from_entries(current_pre), patients_from_entries(current_igd), pilot, copilot, erm, review,
        )
        st.session_state.assignment_draft = assignment
        st.session_state.assignment_draft_date = selected_date

    draft = st.session_state.get("assignment_draft") if st.session_state.get("assignment_draft_date") == selected_date else None
    assignment_to_edit = draft or (saved or {}).get("assignment")
    if assignment_to_edit:
        initial_text = assignment_text(assignment_to_edit, labels)
        st.markdown("<div class='panel'><b>Pratinjau pembagian</b><br><span style='color:#60717d'>Ubah teks bila ada pembagian manual, lalu simpan. Teks tersimpan menjadi pembagian resmi untuk tanggal ini.</span></div>", unsafe_allow_html=True)
        st.markdown("<div class='panel'><b>Ubah pembagian</b><br><span style='color:#60717d'>Semua teks di bawah dapat diubah. Tombol simpan akan menimpa pembagian tanggal ini di Supabase.</span></div>", unsafe_allow_html=True)
        with st.form(f"assignment_save_form_{selected_date}", border=False):
            manual_text = st.text_area("Pembagian tanggal terpilih", value=(saved or {}).get("assignment_text", initial_text) if not draft else initial_text, height=440, key=f"assignment_text_{selected_date}")
            save_assignment = st.form_submit_button("Simpan pembagian tanggal ini", type="primary", use_container_width=True)
            save_assignment = st.form_submit_button("Paksa simpan perubahan manual", type="primary", use_container_width=True)
        if save_assignment:
            error = save_daily_assignment(month_key, selected_date, assignment_to_edit, manual_text)
            if error:
                st.error(error)
            else:
                st.session_state.assignment_draft = None
                st.success(f"Pembagian {readable_date} tersimpan dan dapat dibuka kembali.")
                st.rerun()


def render_roster_intake():
    st.markdown("<div class='masthead'><div class='service-line'>DEPARTEMEN BEDAH MULUT & MAKSILOFASIAL</div><h1>Pembagian Jaga</h1><p>Roster disimpan per bulan. Data paste dibaca sebagai kelompok angkatan murni; Pilot dan Co-pilot dipilih kemudian saat pembagian.</p></div>", unsafe_allow_html=True)
    today = date.today()
    month_col, year_col = st.columns([2, 1])
    with month_col:
        month_number = st.selectbox("Bulan roster", list(range(1, 13)), index=today.month - 1, format_func=lambda value: MONTHS[value - 1], key="roster_month_number")
    with year_col:
        year = int(st.number_input("Tahun roster", min_value=2024, max_value=2035, value=today.year, key="roster_year"))
    month_key = date(year, month_number, 1).isoformat()
    if st.session_state.get("active_roster_month") != month_key:
        stored, database_error = reset_month_state(month_key)
        if database_error:
            st.warning(database_error)
        elif stored:
            st.success(f"Roster {MONTHS[month_number - 1]} {year} dimuat. Terakhir diperbarui {stored.get('updated_at', '-') }.")

    is_admin = admin_access()
    config = st.session_state.cohort_config
    parsed = st.session_state.get("parsed_roster")
    is_admin = bool(st.session_state.get("roster_admin", False))

    # Pembagian adalah aktivitas utama. Tampilkan sebelum pengelolaan roster.
    if parsed is not None and not parsed.empty:
        render_assignment_workspace(parsed, config, month_key, is_admin)
        st.divider()

    is_admin = admin_access()

    if is_admin:
        st.markdown("<div class='panel'><b>Konfigurasi angkatan</b><br><span style='color:#60717d'>Awalnya data bernama Kelompok 1–8. Ubah labelnya menjadi angkatan yang benar, serta tambahkan atau kurangi baris bila format sumber berubah.</span></div>", unsafe_allow_html=True)
        with st.form("cohort_config_form", border=False):
            edited_config = st.data_editor(
                config,
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                column_config={"Residen per hari": st.column_config.NumberColumn(min_value=1, step=1), "Aktif": st.column_config.CheckboxColumn()},
                key="cohort_config_editor",
            )
            apply_config = st.form_submit_button("Terapkan konfigurasi angkatan")
        if apply_config:
            st.session_state.cohort_config = edited_config
            config = edited_config
            st.session_state.pop("parsed_roster_editor", None)
            st.success("Konfigurasi angkatan diterapkan. Paste atau petakan ulang roster bila jumlah kelompok berubah.")
        st.markdown("<div class='panel'><b>Tempel atau perbarui roster</b><br><span style='color:#60717d'>Paste hanya sekali untuk bulan ini. Data lama akan diganti saat admin menekan Simpan roster.</span></div>", unsafe_allow_html=True)
        pasted = st.text_area("Roster yang ditempel", value=st.session_state.get("pasted_roster", ""), height=250, placeholder="Tempel seluruh tabel roster di sini…", key="pasted_roster_input")
        if st.button("Petakan roster", type="primary", key="map_roster"):
            st.session_state.pasted_roster = pasted
            parsed, warnings, skipped = parse_pasted_roster(pasted, config)
            st.session_state.parsed_roster = parsed
            st.session_state.roster_warnings = warnings
            st.session_state.roster_skipped = skipped
            st.session_state.pop("parsed_roster_editor", None)
        parsed = st.session_state.get("parsed_roster")

    if parsed is None:
        st.info("Belum ada roster tersimpan untuk bulan ini. Admin perlu menempel dan menyimpan roster terlebih dahulu.")
        return
    for warning in st.session_state.get("roster_warnings", []):
        st.warning(warning)
    if parsed.empty:
        st.error("Belum ada tanggal yang dapat dipetakan. Cek struktur paste dan konfigurasi angkatan.")
        return

    labels = dict(zip(config["Kolom"].astype(str).str.lower(), config["Label angkatan"].astype(str)))
    shown = parsed.rename(columns=labels)
    st.markdown("<div class='panel'><b>Roster bulan terpilih</b><br><span style='color:#60717d'>Pastikan tabel ini benar sebelum dipakai untuk pembagian jaga.</span></div>", unsafe_allow_html=True)
    if is_admin:
        with st.form("roster_editor_form", border=False):
            edited = st.data_editor(shown, num_rows="dynamic", hide_index=True, use_container_width=True, height=520, key="parsed_roster_editor")
            apply_col, save_col = st.columns(2)
            apply_edits = apply_col.form_submit_button("Terapkan koreksi tabel", use_container_width=True)
            save_roster = save_col.form_submit_button("Simpan roster bulan ini", type="primary", use_container_width=True)
        reverse_labels = {label: code for code, label in labels.items()}
        corrected = edited.rename(columns=reverse_labels)
        if apply_edits or save_roster:
            st.session_state.parsed_roster = corrected
        if save_roster:
            error = save_monthly_roster(month_key, corrected, config)
            if error:
                st.error(error)
            else:
                st.success(f"Roster {MONTHS[month_number - 1]} {year} tersimpan permanen.")
    else:
        st.dataframe(shown, hide_index=True, use_container_width=True, height=520)
    metrics = st.columns(3)
    metrics[0].metric("Tanggal terbaca", len(shown))
    metrics[1].metric("Angkatan aktif", len(labels))
    metrics[2].metric("Butuh koreksi", len(st.session_state.get("roster_skipped", [])))
    st.caption("Roster ini adalah sumber pembagian Post-op, Pre-op, dan IGD. Pilot dan Co-pilot akan dipilih dari kelompok yang tersedia saat pembagian, bukan dibaca dari paste. Pengguna biasa tidak dapat mengubah roster.")
    st.divider()
    render_assignment_workspace(st.session_state.parsed_roster, config, month_key, is_admin)


def init_state():
    st.session_state.setdefault("unavailable", {name: set() for name in NAMES})
    st.session_state.setdefault("forbidden", set())
    st.session_state.setdefault("schedule", None)
    st.session_state.setdefault("summary", None)
    st.session_state.setdefault("notices", [])


def longest_streak(day_sets, name):
    best = current = 0
    for team in day_sets:
        current = current + 1 if name in team else 0
        best = max(best, current)
    return best


def targets_for(slots, rng):
    base, extra = divmod(slots, len(NAMES))
    order = NAMES[:]
    rng.shuffle(order)
    return {name: base + int(index < extra) for index, name in enumerate(order)}


def sunday_targets(days, quota, totals, rng):
    """Hard Sunday fairness: each person gets either floor or ceil share."""
    slots = sum(day.weekday() == 6 for day in days) * quota
    base, extra = divmod(slots, len(NAMES))
    order = NAMES[:]
    rng.shuffle(order)
    order.sort(key=lambda name: totals[name], reverse=True)
    result = {name: base for name in NAMES}
    for name in order[:extra]:
        result[name] += 1
    return result if all(result[name] <= totals[name] for name in NAMES) else None


def forbidden_with(picked, candidate, forbidden):
    return any(tuple(sorted((candidate, other))) in forbidden for other in picked)


def make_fair_jaga(days, quota, unavailable, forbidden, rng):
    if quota == 0:
        return [set() for _ in days], [], {name: 0 for name in NAMES}
    total_slots = len(days) * quota
    max_without_three = (len(days) // 3) * 2 + min(len(days) % 3, 2)
    enforce_streak = total_slots <= len(NAMES) * max_without_three
    notices = [] if enforce_streak else ["Batas dua Jaga berturut-turut dilonggarkan karena kuota terlalu tinggi; fairness total dan Minggu tetap dikunci."]
    for _ in range(1500):
        totals = targets_for(total_slots, rng)
        sunday = sunday_targets(days, quota, totals, rng)
        if sunday is None:
            continue
        remaining, sunday_left = Counter(totals), Counter(sunday)
        teams, failed = [], False
        for current in days:
            key, is_sunday = current.isoformat(), current.weekday() == 6
            yesterday = teams[-1] if teams else set()
            two_days_ago = teams[-2] if len(teams) > 1 else set()

            def eligible(name):
                if remaining[name] <= 0 or key in unavailable[name]:
                    return False
                if enforce_streak and name in yesterday and name in two_days_ago:
                    return False
                # Preserve Sunday share, so a Sunday target cannot be consumed on a weekday.
                return sunday_left[name] > 0 if is_sunday else remaining[name] > sunday_left[name]

            picked = []
            for _slot in range(quota):
                candidates = [name for name in NAMES if name not in picked and eligible(name) and not forbidden_with(picked, name, forbidden)]
                if not candidates:
                    failed = True
                    break

                def score(name):
                    gentle_pair_bias = 0.0
                    if {"Ferrel", "Alam"} & set(picked) and name in {"Ferrel", "Alam"} and rng.random() < .42:
                        gentle_pair_bias = .35
                    return (remaining[name], sunday_left[name] if is_sunday else 0, int(name not in yesterday), int(name not in two_days_ago), gentle_pair_bias, rng.random())
                picked.append(max(candidates, key=score))
            if failed:
                break
            for name in picked:
                remaining[name] -= 1
                if is_sunday:
                    sunday_left[name] -= 1
            teams.append(set(picked))
        if not failed and not any(remaining.values()) and not any(sunday_left.values()):
            return teams, notices, sunday
    return None, ["Tidak ada jadwal yang memenuhi fairness total, fairness Minggu, dan semua request. Ubah request atau kuota."], None


def choose_role(candidates, quota, counts, rng):
    if quota == 0:
        return []
    if len(candidates) < quota:
        return []
    return sorted(candidates, key=lambda name: (counts[name], rng.random()))[:quota]


def make_schedule(days, quotas, doru, unavailable, forbidden, seed):
    rng = random.Random(seed)
    teams, notices, _ = make_fair_jaga(days, quotas["Jaga"], unavailable, forbidden, rng)
    if teams is None:
        return None, None, notices
    counts = {role: Counter() for role in quotas}
    sunday, rows = Counter(), []
    for index, current in enumerate(days):
        jaga = sorted(teams[index]); used = set(jaga)
        review = choose_role([name for name in NAMES if name not in used and name not in doru], quotas["Review"], counts["Review"], rng)
        if len(review) != quotas["Review"]:
            return None, None, [f"{current:%d %b}: kuota Review tidak dapat dipenuhi."]
        used.update(review)
        erm = choose_role([name for name in NAMES if name not in used and name not in doru], quotas["ERM"], counts["ERM"], rng)
        if len(erm) != quotas["ERM"]:
            return None, None, [f"{current:%d %b}: kuota ERM tidak dapat dipenuhi."]
        for role, people in (("Jaga", jaga), ("Review", review), ("ERM", erm)):
            counts[role].update(people)
        if current.weekday() == 6:
            sunday.update(jaga)
        rows.append({"date":current.isoformat(), "Tanggal":current.strftime("%a, %d %b %Y"), "jaga":", ".join(jaga), "review":", ".join(review), "erm":", ".join(erm), "Jaga":", ".join(jaga), "Review":", ".join(review), "ERM":", ".join(erm)})
    total = [counts["Jaga"][name] for name in NAMES]
    sunday_counts = [sunday[name] for name in NAMES]
    streaks = {name: longest_streak(teams, name) for name in NAMES}
    if max(total) - min(total) > 1 or max(sunday_counts) - min(sunday_counts) > 1:
        return None, None, ["Validasi fairness gagal; tidak ada jadwal diekspor."]
    if max(streaks.values()) > 2 and not notices:
        return None, None, ["Validasi Jaga berturut-turut gagal; tidak ada jadwal diekspor."]
    summary = pd.DataFrame({"Nama":NAMES, "Jaga":total, "Jaga Minggu":sunday_counts, "Jaga Berturut Terpanjang":[streaks[name] for name in NAMES], "Review":[counts["Review"][name] for name in NAMES], "ERM":[counts["ERM"][name] for name in NAMES]})
    return pd.DataFrame(rows), summary, notices


def shade(cell, color):
    props = cell._tc.get_or_add_tcPr(); node = props.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd"); props.append(node)
    node.set(qn("w:fill"), color)


def bordered(cell):
    props = cell._tc.get_or_add_tcPr(); borders = props.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders"); props.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}"); borders.append(node)
        node.set(qn("w:val"), "single"); node.set(qn("w:sz"), "6"); node.set(qn("w:color"), "1F2937")


def write_cell(cell, text, size=8, bold=False):
    cell.text = ""; p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text); run.font.name = "Times New Roman"; run.font.size = Pt(size); run.bold = bold
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER; bordered(cell)


def weeks_between(start, end):
    current = start - timedelta(days=start.weekday()); finish = end + timedelta(days=6-end.weekday()); result = []
    while current <= finish:
        result.append([current + timedelta(days=offset) if start <= current + timedelta(days=offset) <= end else None for offset in range(7)])
        current += timedelta(days=7)
    return result


def schedule_docx(schedule, summary, start, end):
    doc = Document(); section = doc.sections[0]; section.orientation = WD_ORIENT.LANDSCAPE; section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(1.5); section.right_margin = Cm(.6); section.top_margin = Cm(1.4); section.bottom_margin = Cm(.6)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; r = p.add_run("JADWAL JAGA REVIEW DAN ERM ANGKATAN 15"); r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(16)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; label = f"{MONTHS[start.month-1].upper()} {start.year}" if (start.month,start.year)==(end.month,end.year) else f"{start:%d/%m/%Y} sampai {end:%d/%m/%Y}"; r = p.add_run(label); r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(12)
    assignments = {date.fromisoformat(row.date):row for row in schedule.itertuples()}
    for week_index, week in enumerate(weeks_between(start,end)):
        if week_index and week_index % 2 == 0: doc.add_page_break()
        table = doc.add_table(rows=5, cols=7); table.autofit = False
        for column, current in enumerate(week):
            color = "E7B008" if column != 6 else "E05252"; shade(table.cell(0,column),color); write_cell(table.cell(0,column),DAY_NAMES[column],10,True); shade(table.cell(1,column),color); write_cell(table.cell(1,column),current.strftime("%d/%m/%Y") if current else "",9,True)
            for row_index, role in enumerate(("jaga","review","erm"),2):
                cell = table.cell(row_index,column)
                if current:
                    if row_index == 2: shade(cell,"F3F4F6")
                    write_cell(cell,f"{role.title()}\n{getattr(assignments[current],role)}",7.4)
                else: write_cell(cell,"",8)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
    heading = doc.add_heading("Rekap Fairness Jaga",level=2)
    for run in heading.runs: run.font.name="Times New Roman"; run.font.color.rgb=RGBColor(0,0,0)
    note=doc.add_paragraph("Total Jaga dan Jaga Minggu masing-masing berselisih maksimal satu antaranggota."); note.runs[0].font.name="Times New Roman"; note.runs[0].font.size=Pt(10)
    recap=doc.add_table(rows=1,cols=len(summary.columns))
    for column,field in enumerate(summary.columns): shade(recap.cell(0,column),"E7B008"); write_cell(recap.cell(0,column),field,8,True)
    for row in summary.sort_values("Nama").itertuples(index=False):
        cells=recap.add_row().cells
        for column,value in enumerate(row): write_cell(cells[column],str(value),8)
    output=io.BytesIO(); doc.save(output); return output.getvalue()


init_state()
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { --ink:#102b38; --muted:#5e727d; --line:#d5e1e2; --teal:#08756d; --teal-dark:#075a55; --mint:#e8f4f1; --paper:#f5f8f7; --gold:#b98a2d; }
.stApp { background:var(--paper); color:var(--ink); font-family:'DM Sans','Helvetica Neue',Arial,sans-serif; }
.block-container { max-width:1240px; padding-top:2.2rem; padding-bottom:4.5rem; }
h1,h2,h3,[data-testid='stMetricLabel'] { font-family:'Manrope','Helvetica Neue',Arial,sans-serif; color:var(--ink); letter-spacing:-.035em; }
.masthead { position:relative; background:#fff; border:1px solid var(--line); border-radius:16px; padding:1.55rem 1.65rem 1.6rem; margin:0 0 1.55rem; overflow:hidden; }
.masthead:before { content:'OMFS'; position:absolute; right:1.65rem; top:1.35rem; color:#dcebea; font-family:'Manrope',sans-serif; font-size:2.2rem; font-weight:800; letter-spacing:-.08em; }
.masthead:after { content:''; position:absolute; left:1.65rem; bottom:0; width:112px; height:3px; background:var(--teal); }
.masthead h1 { position:relative; margin:.25rem 0 .3rem; font-size:2.2rem; font-weight:800; }
.masthead p { color:var(--muted); max-width:700px; margin:0; font-size:.96rem; }
.service-line { color:var(--teal); font-family:'Manrope',sans-serif; font-size:.72rem; letter-spacing:.14em; font-weight:800; }
.panel { background:#fff; border:1px solid var(--line); border-radius:12px; padding:1.08rem 1.2rem; margin:.8rem 0 1rem; box-shadow:0 1px 1px rgba(16,43,56,.02); }
.panel b { font-family:'Manrope',sans-serif; font-size:.95rem; }
[data-testid='stMetric'] { background:#fff; border:1px solid var(--line); border-radius:12px; padding:1rem; box-shadow:none; }
[data-testid='stMetricValue'] { color:var(--teal-dark); font-family:'Manrope',sans-serif; }
div.stButton > button { border-radius:8px; font-family:'DM Sans',sans-serif; font-weight:700; min-height:2.55rem; box-shadow:none; border-color:#b9ccce; }
div.stButton > button[kind='primary'] { background:var(--teal); border-color:var(--teal); }
div.stButton > button[kind='primary']:hover { background:var(--teal-dark); border-color:var(--teal-dark); }
div[data-testid='stTabs'] button { font-family:'Manrope',sans-serif; font-size:.88rem; }
div[data-testid='stExpander'], div[data-testid='stForm'] { background:#fff; border:1px solid var(--line); border-radius:12px; }
div[data-testid='stDataFrame'] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
div[role='radiogroup'] { background:#e7efee; border-radius:10px; padding:4px; width:fit-content; }
div[data-baseweb='select'] > div, div[data-baseweb='input'] > div { border-radius:8px; }
textarea { font-family:'DM Sans','Helvetica Neue',Arial,sans-serif !important; line-height:1.55 !important; }
[data-testid='stForm'] { padding:1.1rem 1.15rem .4rem; }
button[kind='secondary'] { background:#fff; }
</style>""", unsafe_allow_html=True)
module = st.radio("Modul", ["Penjadwalan Jaga, Review, ERM", "Pembagian Jaga"], horizontal=True, label_visibility="collapsed", key="module")
if module == "Pembagian Jaga":
    render_roster_intake()
    st.stop()
st.markdown("<div class='masthead'><div class='service-line'>DEPARTEMEN BEDAH MULUT & MAKSILOFASIAL • ANGKATAN 15</div><h1>Clinical Duty Roster</h1><p>Susun Jaga, Review, dan ERM dengan distribusi yang tervalidasi. Fairness total dan hari Minggu dikunci sebelum jadwal dapat diekspor.</p></div>", unsafe_allow_html=True)

today=date.today(); mode=st.radio("Jenis jadwal",["Bulanan","Rentang tanggal"],horizontal=True,label_visibility="collapsed",key="schedule_mode")
if mode=="Bulanan":
    left,right=st.columns([2,1])
    with left: month=st.selectbox("Bulan",list(range(1,13)),index=today.month-1,format_func=lambda item:MONTHS[item-1])
    with right: year=int(st.number_input("Tahun",2024,2035,today.year))
    start=date(year,month,1); next_month=date(year+(month==12),1 if month==12 else month+1,1); end=next_month-timedelta(days=1)
else:
    left,right=st.columns(2)
    with left: start=st.date_input("Mulai",value=today.replace(day=1))
    with right: end=st.date_input("Selesai",value=today)

st.markdown("<div class='panel'><b>Komposisi layanan harian</b><br><span style='color:#60717d'>Tetapkan kebutuhan staf untuk periode ini. Tidak ada Kares atau peran Bantu.</span></div>",unsafe_allow_html=True)
c1,c2,c3,c4=st.columns([1,1,1,1.7])
with c1:q_jaga=int(st.number_input("Jaga",0,len(NAMES),4))
with c2:q_review=int(st.number_input("Review",0,len(NAMES),3))
with c3:q_erm=int(st.number_input("ERM",0,len(NAMES),2))
with c4:seed=int(st.number_input("Variasi jadwal",1,999999,1501,help="Ganti angka untuk alternatif yang tetap fair."))

with st.expander("Aturan tim & request khusus"):
    doru=st.multiselect("Doru — tepat 2 orang (tetap Jaga, tidak masuk Review/ERM)",NAMES,default=["Ferrel","Alam"],max_selections=2,key="doru")
    left,right=st.columns(2)
    with left:
        st.markdown("#### Tidak tersedia untuk Jaga")
        request_name=st.selectbox("Nama",NAMES,key="request_name"); request_date=st.date_input("Tanggal",value=start,min_value=start,max_value=end,key="request_date")
        if st.button("Tambahkan request",use_container_width=True): st.session_state.unavailable[request_name].add(request_date.isoformat()); st.rerun()
        requests=[(name,value) for name in NAMES for value in sorted(st.session_state.unavailable[name])]
        if requests:
            st.dataframe(pd.DataFrame(requests,columns=["Nama","Tidak tersedia"]),hide_index=True,use_container_width=True)
            selected=st.selectbox("Hapus request",[f"{name} — {value}" for name,value in requests])
            if st.button("Hapus request terpilih"):
                name,value=selected.split(" — ");st.session_state.unavailable[name].discard(value);st.rerun()
        else: st.caption("Belum ada request khusus.")
    with right:
        st.markdown("#### Larangan pasangan Jaga")
        pair_a=st.selectbox("Nama pertama",NAMES,key="pair_a");pair_b=st.selectbox("Nama kedua",NAMES,index=1,key="pair_b")
        if st.button("Tambah larangan pasangan",use_container_width=True):
            if pair_a==pair_b: st.error("Pilih dua nama berbeda.")
            else: st.session_state.forbidden.add(tuple(sorted((pair_a,pair_b))));st.rerun()
        if st.session_state.forbidden: st.dataframe(pd.DataFrame(sorted(st.session_state.forbidden),columns=["Nama pertama","Nama kedua"]),hide_index=True,use_container_width=True)

total_roles=q_jaga+q_review+q_erm; valid=start<=end and len(doru)==2 and total_roles<=len(NAMES)
if start>end: st.error("Tanggal selesai harus sesudah tanggal mulai.")
elif len(doru)!=2: st.error("Pilih tepat dua Doru.")
elif total_roles>len(NAMES): st.error("Total kuota harian melebihi 11 anggota; satu orang tidak boleh memegang dua peran sehari.")
g1,g2,g3=st.columns([1.3,1,3])
with g1: generate=st.button("Buat jadwal fair",type="primary",use_container_width=True,disabled=not valid,key="generate")
with g2: shuffle=st.button("Acak ulang",use_container_width=True,disabled=not valid,key="shuffle")
with g3: st.caption("Aturan keras: satu peran/hari • Doru hanya Jaga • total Jaga dan Jaga Minggu selisih maksimal satu.")
if generate or shuffle:
    days=[start+timedelta(days=index) for index in range((end-start).days+1)]; quotas={"Jaga":q_jaga,"Review":q_review,"ERM":q_erm}
    with st.spinner("Mencari susunan yang fair untuk semua orang..."):
        schedule,summary,notices=make_schedule(days,quotas,set(doru),st.session_state.unavailable,st.session_state.forbidden,seed+int(shuffle))
    if schedule is None: st.error(" ".join(notices))
    else:
        st.session_state.schedule=schedule;st.session_state.summary=summary;st.session_state.notices=notices;st.session_state.period=(start,end)

if st.session_state.schedule is not None:
    schedule,summary=st.session_state.schedule,st.session_state.summary
    st.markdown("## Hasil jadwal")
    m1,m2,m3,m4=st.columns(4);m1.metric("Hari terjadwal",len(schedule));m2.metric("Rentang total Jaga",f"{summary['Jaga'].min()}–{summary['Jaga'].max()}");m3.metric("Rentang Jaga Minggu",f"{summary['Jaga Minggu'].min()}–{summary['Jaga Minggu'].max()}");m4.metric("Berturut terpanjang",int(summary['Jaga Berturut Terpanjang'].max()))
    tab1,tab2,tab3=st.tabs(["Jadwal","Fairness","Word"])
    with tab1:st.dataframe(schedule[["Tanggal","Jaga","Review","ERM"]],hide_index=True,use_container_width=True,height=520)
    with tab2:
        st.dataframe(summary.sort_values("Nama"),hide_index=True,use_container_width=True)
        st.success("Lolos validasi: total Jaga dan Jaga Minggu masing-masing selisih maksimal satu antaranggota.")
        for notice in st.session_state.notices:st.warning(notice)
    with tab3:
        word_file=schedule_docx(schedule,summary,*st.session_state.period)
        st.download_button("Unduh jadwal Word",word_file,"jadwal-angkatan-15.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
