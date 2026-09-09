import random
from collections import Counter
from datetime import date
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Jadwal Angkatan 15", page_icon="🗓️", layout="wide")
ROSTER = {"Aliyah":"F", "Soma":"M", "Syamsul":"M", "Ferrel":"M", "Kezia":"F", "Alam":"M", "Bagus":"M", "Retno":"F", "Rachel":"F", "Irpan":"M", "Farez":"M"}
NAMES = list(ROSTER)

def parse_unavailable(text):
    result = {n:set() for n in NAMES}; lookup = {n.lower():n for n in NAMES}
    for line in text.splitlines():
        if ":" not in line: continue
        raw, dates = line.split(":", 1); name = lookup.get(raw.strip().lower())
        if name: result[name].update(x.strip() for x in dates.split(",") if x.strip())
    return result

def parse_pairs(text):
    result = set(); lookup = {n.lower():n for n in NAMES}
    for item in text.split(","):
        if "-" not in item: continue
        a, b = (x.strip() for x in item.split("-", 1)); a, b = lookup.get(a.lower()), lookup.get(b.lower())
        if a and b and a != b: result.add(frozenset((a,b)))
    return result

def conflict(name, group, pairs):
    return any(frozenset((name, other)) in pairs for other in group)

def make_schedule(days, quotas, doru, female_required, unavailable, forbidden, seed):
    rng = random.Random(seed); counts = {r:Counter() for r in quotas}; sunday = Counter()
    rows, warnings, paired_days = [], [], 0
    pair_cap = int(len(days) * 0.75)  # internal preference: frequent, never mandatory
    for day in days:
        key, used = day.isoformat(), set(); assigned = {r:[] for r in quotas}
        def pool(role, group):
            people = [n for n in NAMES if n not in used and key not in unavailable[n]]
            if role in ("Review", "ERM"): people = [n for n in people if n not in doru]
            return [n for n in people if not conflict(n, group, forbidden)]
        for role in ("Jaga", "Review", "ERM"):
            group, need = assigned[role], quotas[role]
            if role == "Jaga" and female_required and need:
                women = [n for n in pool(role, group) if ROSTER[n] == "F"]
                if women:
                    person = min(women, key=lambda n:(counts[role][n], rng.random())); group.append(person); used.add(person)
                else: warnings.append(f"{day:%d %b}: tidak ada perempuan untuk Jaga.")
            if role == "Jaga" and need-len(group) >= 2 and paired_days < pair_cap:
                pref = ("Ferrel", "Alam")
                if rng.random() < .82 and all(n in pool(role, group) for n in pref):
                    group.extend(pref); used.update(pref); paired_days += 1
            while len(group) < need:
                people = pool(role, group)
                if not people:
                    warnings.append(f"{day:%d %b}: kuota {role} hanya {len(group)}/{need}."); break
                def score(n):
                    return counts[role][n]*30 + (sunday[n]*12 if role=="Jaga" and day.weekday()==6 else 0) + rng.random()
                person = min(people, key=score); group.append(person); used.add(person)
            for person in group:
                counts[role][person] += 1
                if role == "Jaga" and day.weekday() == 6: sunday[person] += 1
        rows.append({"Tanggal":day.strftime("%a, %d %b %Y"), "Jaga":", ".join(assigned["Jaga"]) or "—", "Review":", ".join(assigned["Review"]) or "—", "ERM":", ".join(assigned["ERM"]) or "—"})
    return pd.DataFrame(rows), counts, warnings


st.title("🗓️ Jadwal Angkatan 15")
st.caption("Generator Jaga, Review, dan ERM — tanpa Kares atau peran bantu.")
with st.sidebar:
    st.header("Pengaturan")
    start = st.date_input("Tanggal mulai", value=date.today())
    total_days = st.number_input("Jumlah hari", 1, 60, 14)
    jaga = st.number_input("Orang Jaga per hari", 1, len(NAMES), 4)
    review = st.number_input("Orang Review per hari", 0, len(NAMES), 3)
    erm = st.number_input("Orang ERM per hari", 0, len(NAMES), 2)
    doru = st.multiselect("Doru (pilih tepat 2)", NAMES, default=["Ferrel", "Alam"])
    female = st.checkbox("Jaga wajib ada minimal satu perempuan", value=True)
    seed = st.number_input("Variasi jadwal", 1, 999999, 1501)

st.info("Aturan absolut: satu orang hanya satu peran per hari; Review dan ERM tidak boleh rangkap; Doru boleh Jaga, tetapi tidak boleh Review/ERM. Fairness dihitung dari beban Jaga, termasuk Jaga hari Minggu.")
left, right = st.columns(2)
with left:
    unavailable_text = st.text_area("Tidak tersedia (satu baris per orang)", placeholder="Aliyah: 2026-09-12, 2026-09-15", help="Format: Nama: YYYY-MM-DD, YYYY-MM-DD")
with right:
    forbidden_text = st.text_area("Pasangan yang tidak boleh satu peran", placeholder="Aliyah-Soma, Kezia-Bagus", help="Pasangan ini tidak akan ditempatkan pada kelompok peran yang sama.")

if len(doru) != 2:
    st.warning("Pilih tepat dua orang Doru sebelum membuat jadwal.")
if st.button("Buat jadwal", type="primary", disabled=len(doru) != 2):
    quotas = {"Jaga":int(jaga), "Review":int(review), "ERM":int(erm)}
    if sum(quotas.values()) > len(NAMES):
        st.error("Kuota harian melebihi jumlah anggota karena peran tidak boleh dirangkap.")
    else:
        days = list(pd.date_range(start, periods=int(total_days), freq="D").date)
        table, counts, warnings = make_schedule(days, quotas, set(doru), female, parse_unavailable(unavailable_text), parse_pairs(forbidden_text), int(seed))
        st.session_state["table"] = table
        st.session_state["counts"] = counts
        st.session_state["warnings"] = warnings


if "table" in st.session_state:
    table = st.session_state["table"]
    st.subheader("Jadwal")
    st.dataframe(table, use_container_width=True, hide_index=True)
    st.download_button("Unduh CSV", table.to_csv(index=False).encode("utf-8-sig"), "jadwal-angkatan-15.csv", "text/csv")
    st.subheader("Ringkasan fairness Jaga")
    jaga_counts = st.session_state["counts"]["Jaga"]
    summary = pd.DataFrame({"Nama":NAMES, "Total Jaga":[jaga_counts[n] for n in NAMES]}).sort_values(["Total Jaga", "Nama"])
    st.dataframe(summary, use_container_width=True, hide_index=True)
    if st.session_state["warnings"]:
        st.warning("\n".join(dict.fromkeys(st.session_state["warnings"])))
    else:
        st.success("Semua kuota dan aturan yang dipilih berhasil dipenuhi.")
