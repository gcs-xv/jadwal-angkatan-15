
st.markdown("""<style>
.stApp{background:radial-gradient(circle at 8% 0%,#1c315d 0,#0b1222 34rem,#f6f8fc 34rem)} .block-container{max-width:1240px;padding-top:2.2rem;padding-bottom:4rem} h1,h2{letter-spacing:-.035em}.hero{color:white;padding:.4rem 0 2rem}.hero p{color:#c9d7f2;max-width:710px}.eyebrow{color:#7dd3fc;letter-spacing:.14em;font-weight:700;font-size:.72rem}.panel{background:#fff;border:1px solid #e3e8f1;border-radius:18px;padding:17px 20px;box-shadow:0 8px 24px rgba(22,34,55,.06);margin:.5rem 0 1rem}[data-testid="stMetric"]{background:#fff;border:1px solid #e3e8f1;border-radius:16px;padding:13px 15px;box-shadow:0 8px 22px rgba(22,34,55,.06)}div.stButton>button{border-radius:10px;font-weight:650;min-height:2.65rem}
</style>""",unsafe_allow_html=True)
st.markdown("<div class='hero'><div class='eyebrow'>ANGKATAN 15 • PLANNING STUDIO</div><h1>Jadwal yang fair, tanpa drama.</h1><p>Atur periode, pilih kuota, lalu sistem membagi Jaga secara merata—termasuk hari Minggu. Request khusus cukup pilih nama dan tanggal.</p></div>",unsafe_allow_html=True)
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { --ink:#152a38; --muted:#60717d; --line:#dce5e7; --teal:#0b766e; --teal-dark:#075b55; --mint:#e9f5f2; --paper:#f7faf9; }
.stApp { background:var(--paper); color:var(--ink); font-family:'DM Sans','Helvetica Neue',Arial,sans-serif; }
.block-container { max-width:1180px; padding-top:2.5rem; padding-bottom:4rem; }
h1,h2,h3,[data-testid='stMetricLabel'] { font-family:'Manrope','Helvetica Neue',Arial,sans-serif; color:var(--ink); letter-spacing:-.035em; }
.masthead { border-bottom:1px solid var(--line); padding:0 0 1.85rem; margin-bottom:1.55rem; }
.masthead h1 { margin:.25rem 0 .3rem; font-size:2rem; font-weight:800; }
.masthead p { color:var(--muted); max-width:700px; margin:0; font-size:.96rem; }
.service-line { color:var(--teal); font-family:'Manrope',sans-serif; font-size:.72rem; letter-spacing:.12em; font-weight:800; }
.panel { background:#fff; border:1px solid var(--line); border-radius:10px; padding:1.1rem 1.2rem; margin:.8rem 0 1rem; }
.panel b { font-family:'Manrope',sans-serif; font-size:.95rem; }
[data-testid='stMetric'] { background:#fff; border:1px solid var(--line); border-radius:10px; padding:1rem; box-shadow:none; }
[data-testid='stMetricValue'] { color:var(--teal-dark); font-family:'Manrope',sans-serif; }
div.stButton > button { border-radius:7px; font-family:'DM Sans',sans-serif; font-weight:700; min-height:2.55rem; box-shadow:none; }
div.stButton > button[kind='primary'] { background:var(--teal); border-color:var(--teal); }
div.stButton > button[kind='primary']:hover { background:var(--teal-dark); border-color:var(--teal-dark); }
div[data-testid='stTabs'] button { font-family:'Manrope',sans-serif; font-size:.88rem; }
div[data-testid='stExpander'] { background:#fff; border:1px solid var(--line); border-radius:10px; }
div[data-testid='stDataFrame'] { border:1px solid var(--line); border-radius:10px; overflow:hidden; }
</style>""", unsafe_allow_html=True)
st.markdown("<div class='masthead'><div class='service-line'>DEPARTEMEN BEDAH MULUT & MAKSILOFASIAL • ANGKATAN 15</div><h1>Clinical Duty Roster</h1><p>Susun Jaga, Review, dan ERM dengan distribusi yang tervalidasi. Fairness total dan hari Minggu dikunci sebelum jadwal dapat diekspor.</p></div>", unsafe_allow_html=True)


st.markdown("<div class='panel'><b>Komposisi per hari</b><br><span style='color:#64748b'>Semua peran dapat diubah. Tidak ada Kares atau Bantu.</span></div>",unsafe_allow_html=True)
st.markdown("<div class='panel'><b>Komposisi layanan harian</b><br><span style='color:#60717d'>Tetapkan kebutuhan staf untuk periode ini. Tidak ada Kares atau peran Bantu.</span></div>",unsafe_allow_html=True)
c1,c2,c3,c4=st.columns([1,1,1,1.7])
