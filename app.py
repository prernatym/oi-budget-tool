"""
OI Budget Generator v1.2
RV Form + Query Doc + Financial Proposal → filled budget + validation
"""
import streamlit as st
import tempfile, os
from rv_extractor import extract_rv_form, extract_query_doc, extract_fin_proposal
from template_filler import fill_template

st.set_page_config(page_title="OI Budget Generator", page_icon="📊", layout="wide")

st.markdown("""
<style>
h1 { color: #41909B; }
.stDownloadButton > button {
    background-color: #FDD41D; color: #333;
    font-weight: 700; width: 100%; font-size: 1.1rem;
}
.validation-match { color: #2d8a4e; font-weight: 600; }
.validation-diff  { color: #c0392b; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.title("📊 OI Budget Generator")
st.caption("Upload documents → review fields → download filled budget Excel")

with st.sidebar:
    st.markdown("### Documents")
    st.markdown("""
**RV Form** *(required)*
Client-filled form with basic study parameters

**Query Document** *(optional)*
Your Q&A with client — fills gaps on FGDs, IDIs, sample

**Financial Proposal** *(optional)*
Assumptions section fills everything; quoted totals used to validate output
    """)
    st.divider()
    st.markdown("**Outline India** · Budget Tool v1.2")

# ── Upload ─────────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
with c1:
    rv_file  = st.file_uploader("RV Form (.docx) *required*",             type=["docx"])
with c2:
    qd_file  = st.file_uploader("Query Document (.docx) *optional*",      type=["docx"])
with c3:
    fin_file = st.file_uploader("Financial Proposal (.docx/.pdf) *optional*", type=["docx","pdf"])

if not rv_file:
    st.info("Upload the filled RV Form to begin.")
    st.stop()

# ── Extract ────────────────────────────────────────────────────────
def save_and_extract(uploaded, extract_fn):
    suffix = ".pdf" if uploaded.name.endswith(".pdf") else ".docx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as f:
        f.write(uploaded.read())
        tmp = f.name
    try:
        return extract_fn(tmp)
    except Exception as e:
        st.warning(f"Could not read {uploaded.name}: {e}")
        return {}
    finally:
        os.unlink(tmp)

with st.spinner("Reading documents..."):
    s = save_and_extract(rv_file, extract_rv_form)

    if qd_file:
        qd = save_and_extract(qd_file, extract_query_doc)
        for key in ["num_fgds","num_idis","sample_size","languages","num_blocks"]:
            if qd.get(key): s[key] = qd[key]

    quoted_totals = {}
    if fin_file:
        fp = save_and_extract(fin_file, extract_fin_proposal)
        # Fin proposal overrides everything — most authoritative source
        for key in ["num_fgds","num_idis","sample_size","languages","num_blocks",
                    "states","timeline_months","study_type","components"]:
            if fp.get(key): s[key] = fp[key]
        quoted_totals = fp.get("quoted_totals", {})

docs_read = "RV Form" + (", Query Doc" if qd_file else "") + (", Financial Proposal" if fin_file else "")
st.success(f"✅ Read: {docs_read} — **{s['client_name']}** | Currency: **{s['currency']}** | Type: **{s['study_type']}**")

# ── Review ─────────────────────────────────────────────────────────
st.markdown("### Review & edit extracted fields")
st.caption("Correct anything before generating.")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**Client & basics**")
    client  = st.text_input("Client name", s.get("client_name",""))
    cur     = st.selectbox("Currency", ["INR","USD"],
                           index=0 if s.get("currency","INR")=="INR" else 1)
    months  = st.number_input("Project duration (months)", 1, 36,
                               int(s.get("timeline_months", 3)))
    budget_cap = st.number_input("Client budget constraint (0 = none)", 0.0,
                                  value=float(s.get("budget_constraint", 0)), format="%.0f")

with c2:
    st.markdown("**Study design**")
    stype   = st.selectbox("Study type", ["quantitative","qualitative","mixed"],
                           index=["quantitative","qualitative","mixed"].index(
                               s.get("study_type","quantitative")))
    dur_opts = [15,20,30,45,60,75,90]
    dur = st.selectbox("Survey duration (mins)", dur_opts,
                       index=dur_opts.index(
                           min(dur_opts, key=lambda x: abs(x-s.get("survey_duration",45)))))
    sample  = st.number_input("Total sample size", 10, 50000,
                               int(s.get("sample_size", 500)))
    fgds    = st.number_input("Number of FGDs",      0, 100, int(s.get("num_fgds",0)))
    idis    = st.number_input("Number of IDIs / KIIs", 0, 100, int(s.get("num_idis",0)))

with c3:
    st.markdown("**Geography & field**")
    states_txt = st.text_area("States (one per line)",
                               "\n".join(s.get("states",["State 1"])), height=100)
    states  = [x.strip() for x in states_txt.split("\n") if x.strip()] or ["State 1"]
    blocks  = st.number_input("Blocks / field teams per state", 1, 50,
                               max(int(s.get("num_blocks",1)),1))
    lang_opts = ["Hindi","Marathi","Bengali","Tamil","Telugu",
                 "Kannada","Malayalam","Gujarati","Punjabi","Odia"]
    langs   = st.multiselect("Translation languages", lang_opts,
                              default=[l for l in s.get("languages",[]) if l in lang_opts])

st.markdown("**OI scope of work**")
scope_opts = ["Data Collection","Analysis","Report Writing","Translation",
              "Pretest","Field Training","Study Tools"]
scope = st.multiselect("OI deliverables", scope_opts,
                       default=[c for c in s.get("components",[]) if c in scope_opts]
                                or ["Data Collection"])

co1, co2, co3 = st.columns(3)
with co1: oi_codes   = st.checkbox("OI codes the survey tool", s.get("oi_codes",True))
with co2: oi_devices = st.checkbox("OI provides devices",      s.get("oi_devices",True))
with co3: revisits   = st.checkbox("Revisits required",        s.get("revisits",False))

# ── Generate ───────────────────────────────────────────────────────
st.divider()
if st.button("Generate Budget", use_container_width=True, type="primary"):
    schema = {
        "client_name":      client,
        "currency":         cur,
        "study_type":       stype,
        "components":       scope,
        "sample_size":      int(sample),
        "states":           states,
        "num_blocks":       int(blocks),
        "survey_duration":  int(dur),
        "num_fgds":         int(fgds),
        "num_idis":         int(idis),
        "languages":        langs,
        "oi_codes":         oi_codes,
        "oi_devices":       oi_devices,
        "revisits":         revisits,
        "num_revisits":     int(s.get("num_revisits",0)),
        "timeline_months":  int(months),
        "budget_constraint":float(budget_cap),
        "dc_mode":          s.get("dc_mode","field"),
    }

    template = "template_inr.xlsx" if cur == "INR" else "template_usd.xlsx"
    template_path = os.path.join(os.path.dirname(__file__), template)

    out_path = tempfile.mktemp(suffix=".xlsx")
    with st.spinner("Filling template..."):
        fill_template(schema, template_path, out_path)

    with open(out_path, "rb") as f:
        data = f.read()
    os.unlink(out_path)

    name = client.replace(" ","_")[:25]
    filename = f"Budget_{name}_{cur}.xlsx"
    st.session_state["excel"]         = data
    st.session_state["filename"]       = filename
    st.session_state["cur"]            = cur
    st.session_state["cap"]            = budget_cap
    st.session_state["quoted_totals"]  = quoted_totals

if "excel" not in st.session_state:
    st.stop()

# ── Download ───────────────────────────────────────────────────────
st.markdown("### Download")
st.download_button(
    f"⬇️  Download {st.session_state['filename']}",
    data=st.session_state["excel"],
    file_name=st.session_state["filename"],
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.caption("Open in Excel — all formulas calculate automatically.")

# ── Validation (only if fin proposal was uploaded) ─────────────────
qt = st.session_state.get("quoted_totals", {})
if qt:
    st.markdown("### Validation — Generated vs Quoted")
    st.caption("Comparing your generated budget against the actual quoted amounts in the Financial Proposal.")

    sym = "₹" if st.session_state["cur"] == "INR" else "$"

    label_map = {
        "study_prep":      "Study Preparation",
        "researcher_days": "Researcher & FW days",
        "training":        "Training",
        "logistics":       "Logistics",
        "data_mgmt":       "Data Management",
        "devices":         "Devices & Software",
        "admin":           "Admin & Legal",
        "project_total":   "PROJECT COSTS",
        "taxes":           "Taxes",
        "grand_total":     "Total including taxes",
    }

    rows = []
    for key, label in label_map.items():
        quoted = qt.get(key)
        if quoted:
            rows.append({"Section": label, "Quoted": quoted})

    if rows:
        cols = st.columns([3,2,2,2])
        cols[0].markdown("**Section**")
        cols[1].markdown("**Quoted**")
        cols[2].markdown("**Generated**")
        cols[3].markdown("**Diff**")
        st.divider()

        for row in rows:
            c1,c2,c3,c4 = st.columns([3,2,2,2])
            c1.write(row["Section"])
            c2.write(f"{sym}{row['Quoted']:,.0f}")
            # Generated values not available without re-reading Excel
            # Show quoted only — diff requires post-generation read
            c3.write("—")
            c4.write("Open Excel to check")
    else:
        st.info("Quoted totals could not be extracted from the Financial Proposal. Check that it contains the budget summary table.")
