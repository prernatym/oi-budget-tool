"""
OI Budget Generator
Upload RV Form → review extracted fields → download filled budget template
"""
import streamlit as st
import tempfile, os, shutil
from rv_extractor import extract_rv_form
from template_filler import fill_template

st.set_page_config(page_title="OI Budget Generator", page_icon="📊", layout="wide")

st.markdown("""
<style>
h1 { color: #41909B; }
.stDownloadButton > button {
    background-color: #FDD41D; color: #333;
    font-weight: 700; width: 100%; font-size: 1.1rem;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 OI Budget Generator")
st.caption("Upload a filled RV Form — review the extracted fields — download the filled budget template")

# ── Sidebar ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Steps")
    st.markdown("1. Upload filled RV Form (.docx)\n2. Review and correct extracted fields\n3. Generate and download budget")
    st.divider()
    st.markdown("**Templates used**")
    st.markdown("- INR: Budget template (Indian)\n- USD: Budget template (non-Indian)")

# ── Upload ─────────────────────────────────────────────────────────
uploaded = st.file_uploader("Upload filled RV Form (.docx)", type=["docx"])
if not uploaded:
    st.info("Upload a filled RV Form to begin.")
    st.stop()

# Extract
with st.spinner("Reading form..."):
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        f.write(uploaded.read())
        tmp = f.name
    try:
        s = extract_rv_form(tmp)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()
    finally:
        os.unlink(tmp)

st.success(f"✅ Read: **{s['client_name']}** | Detected currency: **{s['currency']}** | Study type: **{s['study_type']}**")

# ── Review form ────────────────────────────────────────────────────
st.markdown("### Review extracted fields")
st.caption("Correct anything the auto-extraction got wrong before generating.")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**Client & basics**")
    client  = st.text_input("Client name", s["client_name"])
    cur     = st.selectbox("Currency", ["INR", "USD"], index=0 if s["currency"] == "INR" else 1)
    months  = st.number_input("Project duration (months)", 1, 36, int(s["timeline_months"]))
    budget_cap = st.number_input("Client budget (0 = none)", 0.0, value=float(s["budget_constraint"]), format="%.0f")

with c2:
    st.markdown("**Study design**")
    stype   = st.selectbox("Study type", ["quantitative", "qualitative", "mixed"],
                           index=["quantitative","qualitative","mixed"].index(s["study_type"]))
    dur_opts = [15, 20, 30, 45, 60, 75, 90]
    dur     = st.selectbox("Survey duration (mins)", dur_opts,
                           index=dur_opts.index(min(dur_opts, key=lambda x: abs(x - s["survey_duration"]))))
    sample  = st.number_input("Total sample size (HH surveys)", 10, 50000, int(s["sample_size"]))
    fgds    = st.number_input("Number of FGDs", 0, 100, int(s["num_fgds"]))
    idis    = st.number_input("Number of IDIs / KIIs", 0, 100, int(s["num_idis"]))

with c3:
    st.markdown("**Geography & field**")
    states_txt = st.text_area("States (one per line)", "\n".join(s["states"]), height=100)
    states  = [x.strip() for x in states_txt.split("\n") if x.strip()] or ["State 1"]
    blocks  = st.number_input("Blocks / field teams per state", 1, 50, max(int(s["num_blocks"]), 1))
    lang_opts = ["Hindi","Marathi","Bengali","Tamil","Telugu","Kannada","Malayalam","Gujarati","Punjabi","Odia"]
    langs   = st.multiselect("Translation languages", lang_opts,
                             default=[l for l in s["languages"] if l in lang_opts])

st.markdown("**Scope of work (what OI is responsible for)**")
scope_opts = ["Data Collection", "Analysis", "Report Writing", "Translation", "Pretest",
              "Field Training", "Study Tools"]
scope = st.multiselect("OI deliverables", scope_opts,
                       default=[c for c in s["components"] if c in scope_opts] or ["Data Collection"])

co1, co2, co3 = st.columns(3)
with co1: oi_codes   = st.checkbox("OI codes the survey tool", s["oi_codes"])
with co2: oi_devices = st.checkbox("OI provides devices",      s["oi_devices"])
with co3: revisits   = st.checkbox("Revisits required",        s["revisits"])

# ── Generate ───────────────────────────────────────────────────────
st.divider()
if st.button("Generate Budget", use_container_width=True, type="primary"):
    schema = {
        "client_name":       client,
        "currency":          cur,
        "study_type":        stype,
        "components":        scope,
        "sample_size":       int(sample),
        "states":            states,
        "num_blocks":        int(blocks),
        "survey_duration":   int(dur),
        "num_fgds":          int(fgds),
        "num_idis":          int(idis),
        "languages":         langs,
        "oi_codes":          oi_codes,
        "oi_devices":        oi_devices,
        "revisits":          revisits,
        "num_revisits":      int(s["num_revisits"]),
        "back_check":        s["back_check"],
        "timeline_months":   int(months),
        "budget_constraint": float(budget_cap),
        "dc_mode":           s["dc_mode"],
    }

    template = "template_inr.xlsx" if cur == "INR" else "template_usd.xlsx"
    template_path = os.path.join(os.path.dirname(__file__), template)

    out_path = tempfile.mktemp(suffix=".xlsx")
    with st.spinner("Filling template..."):
        fill_template(schema, template_path, out_path)

    with open(out_path, "rb") as f:
        data = f.read()
    os.unlink(out_path)

    name = client.replace(" ", "_")[:25]
    filename = f"Budget_{name}_{cur}.xlsx"

    st.session_state["excel"] = data
    st.session_state["filename"] = filename
    st.session_state["cur"] = cur
    st.session_state["budget_cap"] = budget_cap

if "excel" in st.session_state:
    st.markdown("### Download")
    if st.session_state.get("budget_cap", 0) > 0:
        st.info(f"📋 Client budget constraint: {st.session_state['cur']} {st.session_state['budget_cap']:,.0f} — open the Excel to check the total against this.")
    st.download_button(
        f"⬇️  Download {st.session_state['filename']}",
        data=st.session_state["excel"],
        file_name=st.session_state["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("The downloaded file is your actual budget template with input cells filled. Open in Excel — all formulas will calculate automatically.")
