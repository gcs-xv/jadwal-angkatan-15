import io
import random
from collections import Counter
from datetime import date, timedelta

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

st.set_page_config(page_title="Jadwal Angkatan 15", page_icon="🗓️", layout="wide")

ROSTER = {
    "Aliyah": "F", "Soma": "M", "Syamsul": "M", "Ferrel": "M",
    "Kezia": "F", "Alam": "M", "Bagus": "M", "Retno": "F",
    "Rachel": "F", "Irpan": "M", "Farez": "M",
}
NAMES = list(ROSTER)
DAY_NAMES = ("Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu")
MONTH_NAMES = ("", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER")


def parse_unavailable(text):
    result = {name: set() for name in NAMES}
    lookup = {name.lower(): name for name in NAMES}
    for line in text.splitlines():
        if ":" not in line:
            continue
        raw_name, raw_dates = line.split(":", 1)
        name = lookup.get(raw_name.strip().lower())
        if name:
            result[name].update(value.strip() for value in raw_dates.split(",") if value.strip())
    return result


def longest_streak(day_sets, name):
    best = current = 0
    for people in day_sets:
        current = current + 1 if name in people else 0
        best = max(best, current)
    return best


def choose_role(candidates, need, counts, rng):
    candidates = list(candidates)
    if len(candidates) < need:
        return []
    return sorted(candidates, key=lambda person: (counts[person], rng.random()))[:need]


def make_schedule(days, quotas, doru, female_required, unavailable, seed):
    """Build a certified schedule. Jaga has equal counts (difference <= 1)."""
    rng = random.Random(seed)
    counts = {role: Counter() for role in quotas}
    sunday = Counter()
    rows, warnings, jaga_days = [], [], []
    offset = rng.randrange(len(NAMES))

    for index, day in enumerate(days):
        key = day.isoformat()
        wanted = [NAMES[(offset + index * quotas["Jaga"] + slot) % len(NAMES)] for slot in range(quotas["Jaga"])]
        jaga = [person for person in wanted if key not in unavailable[person]]
        while len(jaga) < quotas["Jaga"]:
            previous = jaga_days[-1] if jaga_days else set()
            candidates = [person for person in NAMES if person not in jaga and key not in unavailable[person]]
            candidates.sort(key=lambda person: (person in previous, counts["Jaga"][person], rng.random()))
            if not candidates:
                return None, None, [f"{day:%d %b}: kuota Jaga tidak dapat dipenuhi."], None
            jaga.append(candidates[0])
        if female_required and not any(ROSTER[person] == "F" for person in jaga):
            previous = jaga_days[-1] if jaga_days else set()
            women = [person for person in NAMES if ROSTER[person] == "F" and person not in jaga and key not in unavailable[person]]
            if not women:
                return None, None, [f"{day:%d %b}: tidak ada perempuan tersedia untuk Jaga."], None
            replacement = min(women, key=lambda person: (person in previous, counts["Jaga"][person], rng.random()))
            removable = max(jaga, key=lambda person: (ROSTER[person] == "F", counts["Jaga"][person], person in previous))
            jaga[jaga.index(removable)] = replacement

        used = set(jaga)
        review_pool = [person for person in NAMES if person not in used and person not in doru and key not in unavailable[person]]
        review = choose_role(review_pool, quotas["Review"], counts["Review"], rng)
        if len(review) != quotas["Review"]:
            return None, None, [f"{day:%d %b}: kuota Review tidak dapat dipenuhi."], None
        used.update(review)
        erm_pool = [person for person in NAMES if person not in used and person not in doru and key not in unavailable[person]]
        erm = choose_role(erm_pool, quotas["ERM"], counts["ERM"], rng)
        if len(erm) != quotas["ERM"]:
            return None, None, [f"{day:%d %b}: kuota ERM tidak dapat dipenuhi."], None

        for role, people in (("Jaga", jaga), ("Review", review), ("ERM", erm)):
            counts[role].update(people)
        if day.weekday() == 6:
            sunday.update(jaga)
        jaga_days.append(set(jaga))
        rows.append({"date": day.isoformat(), "Tanggal": day.strftime("%a, %d %b %Y"), "jaga": ", ".join(jaga), "review": ", ".join(review), "erm": ", ".join(erm), "Jaga": ", ".join(jaga), "Review": ", ".join(review), "ERM": ", ".join(erm)})

    jaga_values = [counts["Jaga"][name] for name in NAMES]
    longest = {name: longest_streak(jaga_days, name) for name in NAMES}
    if max(jaga_values) - min(jaga_values) > 1:
        warnings.append("Ketersediaan yang dipilih membuat fairness Jaga (selisih maksimal satu) tidak dapat dijamin.")
    if max(longest.values()) > 2:
        warnings.append("Ketersediaan yang dipilih memaksa tiga Jaga berturut-turut atau lebih.")
    if warnings:
        return None, None, warnings, None
    summary = pd.DataFrame({"Nama": NAMES, "Jaga": jaga_values, "Jaga Minggu": [sunday[name] for name in NAMES], "Jaga Berturut Terpanjang": [longest[name] for name in NAMES], "Review": [counts["Review"][name] for name in NAMES], "ERM": [counts["ERM"][name] for name in NAMES]})
    return pd.DataFrame(rows), summary, [], jaga_days


def shade(cell, color):
    props = cell._tc.get_or_add_tcPr()
    node = props.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd"); props.append(node)
    node.set(qn("w:fill"), color)


def border(cell):
    props = cell._tc.get_or_add_tcPr()
    borders = props.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders"); props.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}"); borders.append(node)
        node.set(qn("w:val"), "single"); node.set(qn("w:sz"), "6"); node.set(qn("w:color"), "000000")


def write_cell(cell, text, size=8, bold=False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"; run.font.size = Pt(size); run.bold = bold
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    border(cell)


def weeks_between(start, end):
    current = start - timedelta(days=start.weekday())
    finish = end + timedelta(days=6 - end.weekday())
    weeks = []
    while current <= finish:
        weeks.append([current + timedelta(days=delta) if start <= current + timedelta(days=delta) <= end else None for delta in range(7)])
        current += timedelta(days=7)
    return weeks


def schedule_docx(schedule, summary, start, end):
    doc = Document(); section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(1.5); section.right_margin = Cm(0.6); section.top_margin = Cm(1.4); section.bottom_margin = Cm(0.6)
    title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("JADWAL JAGA REVIEW DAN ERM ANGKATAN 15"); run.bold = True; run.font.name = "Times New Roman"; run.font.size = Pt(16)
    subtitle = doc.add_paragraph(); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"{MONTH_NAMES[start.month]} {start.year}" if (start.month, start.year) == (end.month, end.year) else f"{start:%d/%m/%Y} sampai {end:%d/%m/%Y}"); run.bold = True; run.font.name = "Times New Roman"; run.font.size = Pt(12)
    assignments = {date.fromisoformat(row.date): row for row in schedule.itertuples()}
    for week_index, week in enumerate(weeks_between(start, end)):
        if week_index and week_index % 2 == 0:
            doc.add_page_break()
        table = doc.add_table(rows=5, cols=7); table.autofit = False
        for column, current in enumerate(week):
            header_color = "FF0000" if column == 6 else "FFF200"
            shade(table.cell(0, column), header_color); write_cell(table.cell(0, column), DAY_NAMES[column], 10, True)
            shade(table.cell(1, column), header_color); write_cell(table.cell(1, column), current.strftime("%d/%m/%Y") if current else "", 9, True)
            for row_index, role in enumerate(("jaga", "review", "erm"), start=2):
                cell = table.cell(row_index, column)
                if current:
                    if row_index == 2: shade(cell, "F2F2F2")
                    write_cell(cell, f"{role.title()}\n{getattr(assignments[current], role)}", 7.4)
                else:
                    write_cell(cell, "", 8)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
    heading = doc.add_heading("Rekap Fairness Jaga", level=2)
    for run in heading.runs: run.font.name = "Times New Roman"; run.font.color.rgb = RGBColor(0, 0, 0)
    note = doc.add_paragraph("Selisih total Jaga antaranggota maksimal satu. Tidak ada anggota dengan tiga Jaga berturut-turut.")
    note.runs[0].font.name = "Times New Roman"; note.runs[0].font.size = Pt(10)
    recap = doc.add_table(rows=1, cols=len(summary.columns))
    for column, field in enumerate(summary.columns):
        shade(recap.cell(0, column), "FFF200"); write_cell(recap.cell(0, column), field, 8, True)
    for row in summary.sort_values("Nama").itertuples(index=False):
        cells = recap.add_row().cells
        for column, value in enumerate(row): write_cell(cells[column], str(value), 8)
    output = io.BytesIO(); doc.save(output); return output.getvalue()


st.title("🗓️ Jadwal Angkatan 15")
st.caption("Generator Jaga, Review, dan ERM. Tidak ada Kares atau peran bantu.")
with st.sidebar:
    st.header("Pengaturan")
    start = st.date_input("Tanggal mulai", value=date.today().replace(day=1))
    total_days = st.number_input("Jumlah hari", 1, 60, 30)
    jaga = st.number_input("Orang Jaga per hari", 1, len(NAMES), 4)
    review = st.number_input("Orang Review per hari", 0, len(NAMES), 3)
    erm = st.number_input("Orang ERM per hari", 0, len(NAMES), 2)
    doru = st.multiselect("Doru (pilih tepat 2)", NAMES, default=["Ferrel", "Alam"])
    female = st.checkbox("Jaga wajib ada minimal satu perempuan", value=True)
    seed = st.number_input("Variasi jadwal", 1, 999999, 1501)

st.info("Aturan absolut: setiap orang hanya satu peran per hari; Review dan ERM tidak boleh rangkap; dua Doru tidak boleh Review/ERM tetapi tetap ikut Jaga. Fairness hanya dihitung dari Jaga: selisih total maksimal satu, dengan maksimum dua Jaga berturut-turut.")
unavailable_text = st.text_area("Tidak tersedia (opsional)", placeholder="Aliyah: 2026-09-12, 2026-09-15", help="Format satu baris per orang: Nama: YYYY-MM-DD, YYYY-MM-DD")
if len(doru) != 2:
    st.warning("Pilih tepat dua orang Doru sebelum membuat jadwal.")
if st.button("Buat jadwal", type="primary", disabled=len(doru) != 2):
    quotas = {"Jaga": int(jaga), "Review": int(review), "ERM": int(erm)}
    if sum(quotas.values()) > len(NAMES):
        st.error("Kuota harian melebihi jumlah anggota karena peran tidak boleh rangkap.")
    else:
        days = [start + timedelta(days=index) for index in range(int(total_days))]
        schedule, summary, warnings, _ = make_schedule(days, quotas, set(doru), female, parse_unavailable(unavailable_text), int(seed))
        if warnings:
            st.error("\n".join(warnings))
            st.caption("Ubah ketersediaan, kuota, atau tanggal. Jadwal tidak diekspor bila aturan fairness Jaga tidak terpenuhi.")
        else:
            st.session_state["schedule"] = schedule
            st.session_state["summary"] = summary
            st.session_state["start"] = start
            st.session_state["end"] = days[-1]

if "schedule" in st.session_state:
    schedule = st.session_state["schedule"]
    summary = st.session_state["summary"]
    st.subheader("Jadwal")
    st.dataframe(schedule[["Tanggal", "Jaga", "Review", "ERM"]], use_container_width=True, hide_index=True)
    st.subheader("Ringkasan fairness Jaga")
    st.dataframe(summary.sort_values("Nama"), use_container_width=True, hide_index=True)
    word_file = schedule_docx(schedule, summary, st.session_state["start"], st.session_state["end"])
    st.download_button("Unduh jadwal Word", word_file, "jadwal-angkatan-15.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.success("Jadwal tervalidasi: semua anggota mendapat porsi Jaga yang setara (selisih maksimal satu) dan tidak ada tiga Jaga berturut-turut.")
