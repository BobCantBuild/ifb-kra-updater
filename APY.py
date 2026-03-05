import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(
    page_title="KRA Auto-Updater | IFB",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
* { font-family: 'Inter', sans-serif !important; }
.stApp { background: #060B18; }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stSidebar"] { display: none; }
h1,h2,h3,p,label,span,div { color: #E2E8F0; }
.upload-card {
    background: linear-gradient(145deg, #0F1729, #131E35);
    border: 1px solid #1E3A5F; border-radius: 16px; padding: 22px 20px; margin-bottom: 4px;
}
.card-title { font-size:13px; font-weight:600; color:#64B5F6 !important;
    letter-spacing:0.8px; text-transform:uppercase; margin-bottom:10px; }
.file-ok { display:inline-flex; align-items:center; gap:8px;
    background:rgba(34,197,94,0.12); border:1px solid rgba(34,197,94,0.3);
    border-radius:8px; padding:6px 12px; font-size:12px;
    color:#22C55E !important; margin-top:8px; font-weight:500; }
div[data-testid="stButton"] > button {
    background:linear-gradient(135deg,#2563EB,#7C3AED) !important;
    color:white !important; border:none !important; border-radius:12px !important;
    padding:14px 40px !important; font-size:16px !important; font-weight:600 !important;
    width:100% !important; box-shadow:0 4px 20px rgba(37,99,235,0.35) !important; }
div[data-testid="stButton"] > button:disabled {
    background:linear-gradient(135deg,#1e3a5f,#2d1b5e) !important; color:#4B5563 !important; }
.step-container { display:flex; flex-direction:column; gap:10px; padding:20px;
    background:#0D1526; border-radius:14px; border:1px solid #1E3A5F; margin:16px 0; }
.step-row { display:flex; align-items:center; gap:14px; padding:10px 14px;
    border-radius:8px; font-size:14px; font-weight:500; }
.step-row.done    { background:rgba(34,197,94,0.08);  color:#4ADE80 !important; }
.step-row.running { background:rgba(59,130,246,0.12); color:#60A5FA !important; }
.step-row.pending { color:#4B5563 !important; }
.step-icon { font-size:16px; min-width:22px; }
.metric-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin:20px 0; }
.metric-tile { background:linear-gradient(145deg,#0F1729,#131E35);
    border:1px solid #1E3A5F; border-radius:14px; padding:20px; text-align:center; }
.metric-val { font-size:32px; font-weight:700; color:#60A5FA !important; }
.metric-lbl { font-size:12px; color:#64748B !important; margin-top:4px; font-weight:500; }
[data-testid="stDownloadButton"] > button {
    background:linear-gradient(135deg,#059669,#0D9488) !important;
    color:white !important; border:none !important; border-radius:12px !important;
    padding:14px 40px !important; font-size:16px !important; font-weight:600 !important;
    width:100% !important; box-shadow:0 4px 20px rgba(5,150,105,0.35) !important; margin-top:8px; }
.log-area { background:#060B18; border:1px solid #1E3A5F; border-radius:10px;
    padding:16px 20px; font-family:monospace !important; font-size:12px; color:#475569;
    max-height:320px; overflow-y:auto; line-height:1.8; }
.log-area .ok   { color:#22C55E; } .log-area .warn { color:#F59E0B; }
.log-area .err  { color:#EF4444; } .log-area .info { color:#60A5FA; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:linear-gradient(135deg,#0F1729 0%,#131E35 50%,#0F1729 100%);
     border:1px solid #1E3A5F; border-radius:20px; padding:36px 40px 28px; margin-bottom:28px;">
  <div style="display:flex;align-items:center;gap:18px;margin-bottom:10px;">
    <div style="background:linear-gradient(135deg,#2563EB,#7C3AED);border-radius:14px;
         padding:12px;font-size:26px;line-height:1;">📊</div>
    <div>
      <h1 style="margin:0;font-size:26px;font-weight:700;">KRA Auto-Updater</h1>
      <p style="margin:4px 0 0;color:#64748B;font-size:14px;">IFB Service · Cochin Cluster · Nandu S Kumar</p>
    </div>
  </div>
  <p style="color:#94A3B8;font-size:14px;margin:0;">Upload monthly reports and KRA template — fills all tabs automatically.</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HARDCODED FRANCHISE CODE MAP  (KRA display name → 7-digit franchise code)
# Update this if franchises are added/removed from Nandu's KRA
# ══════════════════════════════════════════════════════════════════════════════
KRA_CODES = {
    "ELECTRO CARE":             "3011917",
    "Hitech Service":           "3001887",
    "J.K Enterprises":          "3010062",
    "CENTURY AIRCONDITIONING":  None,      # inactive / no code in param
    "KRISMA TECH ASSOCIATES":   "3012041",
    "M.S SERVICES":             "3007010",
    "M/s TMR Systems":          "3000821",
    "MFBS INDIA CUSTOMER CARE": "3004150",
    "PROMPTCARE":               "3002573",
    "NATH SERVICE":             "3011679",
    "ULTRA CARE":               "3004835",
}

# Alias map: what each KRA franchise is called INSIDE specific KRA tabs
# (Exchange, Essential Budget, Accesories Budget use short names)
KRA_ALIASES = {
    "Hitech Service":           ["HITECH", "HITECH SERVICE"],
    "J.K Enterprises":          ["JK", "J.K ENTERPRISES", "JK ENTERPRISES"],
    "KRISMA TECH ASSOCIATES":   ["KRISMA TECH", "KRISMA TECH ASSOCIATES"],
    "M.S SERVICES":             ["MS", "MS SERVICES", "M.S SERVICES"],
    "M/s TMR Systems":          ["TMR", "TMR SYSTEMS"],
    "MFBS INDIA CUSTOMER CARE": ["MFBS", "MFBS INDIA"],
    "PROMPTCARE":               ["PROMPT CARE", "PROMPTCARE"],
    "NATH SERVICE":             ["NATH", "NATH SERVICE"],
    "ULTRA CARE":               ["ULTRA CARE"],
    "ELECTRO CARE":             ["ELECTRO CARE"],
    "CENTURY AIRCONDITIONING":  ["CENTURY", "CENTURY AIRCONDITIONING"],
}

# ── CONSTANTS ──────────────────────────────────────────────────────────────────
MONTHS = ["April","May","June","July","August","September",
          "October","November","December","January","February","March"]

def fy_idx(m):            return MONTHS.index(m)
def kra_col(m):           return 4 + fy_idx(m)
def sub_col_letter(m):    return chr(ord('D') + fy_idx(m))
def cluster_month_col(m): return 4 + fy_idx(m)
def mc_closed_col(m):     return 6 + fy_idx(m) * 2

# ── FILE UPLOADS ───────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin-bottom:12px;">① Upload Input Files</p>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="upload-card"><p class="card-title">📁 KL Cluster Report</p>', unsafe_allow_html=True)
    kl_file = st.file_uploader("kl", type=["xlsx"], key="kl", label_visibility="collapsed")
    if kl_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{kl_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">MC Reg · Abv2Days · Social · Apni Dhukhan · ESS · ACC</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="upload-card"><p class="card-title">📁 Parameter Dashboard</p>', unsafe_allow_html=True)
    param_file = st.file_uploader("param", type=["xlsx"], key="param", label_visibility="collapsed")
    if param_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{param_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">INS · SER · CSS · NR · MC Hit · Rep calls · SA Prod · AMC</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="upload-card"><p class="card-title">📁 KRA Template</p>', unsafe_allow_html=True)
    kra_file = st.file_uploader("kra", type=["xlsx"], key="kra", label_visibility="collapsed")
    if kra_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{kra_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">KRA NANDU template (14 tabs)</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ── MONTH SELECTOR ─────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 10px;">② Select Month</p>', unsafe_allow_html=True)
mc1, mc2 = st.columns([1,3])
with mc1:
    selected_month = st.selectbox("Month", MONTHS, index=9, label_visibility="collapsed")
with mc2:
    fi  = fy_idx(selected_month); kc = kra_col(selected_month)
    scl = sub_col_letter(selected_month); mcc = mc_closed_col(selected_month)
    cmc = cluster_month_col(selected_month)
    st.markdown(f"""
    <div style="background:#0F1729;border:1px solid #1E3A5F;border-radius:12px;
         padding:14px 20px;display:flex;gap:28px;align-items:center;flex-wrap:wrap;">
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">Month</div>
           <div style="font-size:18px;font-weight:700;color:#60A5FA;">{selected_month}</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">FY Index</div>
           <div style="font-size:18px;font-weight:700;color:#A78BFA;">{fi+1} / 12</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">Cluster Col</div>
           <div style="font-size:18px;font-weight:700;color:#F472B6;">col {cmc}</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">MC Closed Col</div>
           <div style="font-size:18px;font-weight:700;color:#FB923C;">col {mcc}</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">KRA Write Col</div>
           <div style="font-size:18px;font-weight:700;color:#34D399;">{kc} → {scl}</div></div>
    </div>""", unsafe_allow_html=True)

# ── PROCESS BUTTON ─────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 10px;">③ Process</p>', unsafe_allow_html=True)
all_ready = kl_file and param_file and kra_file
if not all_ready:
    missing = [x for x,f in [("KL Cluster",kl_file),("Parameter",param_file),("KRA Template",kra_file)] if not f]
    st.markdown(f'<div style="background:rgba(251,146,60,0.08);border:1px solid rgba(251,146,60,0.3);border-radius:10px;padding:12px 18px;color:#FB923C;font-size:13px;margin-bottom:12px;">⚠️ Still needed: {" · ".join(missing)}</div>', unsafe_allow_html=True)
btn = st.button("⚡  Generate Updated KRA", disabled=not all_ready)

# ── HELPERS ────────────────────────────────────────────────────────────────────
def safe_int(v):
    try: return int(float(v)) if v not in (None,"") and str(v).strip() not in ('nan','None','—','#N/A','#REF!') else 0
    except: return 0

def safe_float(v):
    try: return float(v) if v not in (None,"") and str(v).strip() not in ('nan','None','—','#N/A','#REF!') else 0.0
    except: return 0.0

def find_sheet(wb, keywords):
    kws = [k.lower() for k in keywords]
    for name in wb.sheetnames:
        if all(k in name.lower() for k in kws): return wb[name]
    for name in wb.sheetnames:
        if any(k in name.lower() for k in kws): return wb[name]
    return None

def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or "").lower())

def aliases_for(kra_name):
    """Return all possible aliases for a KRA franchise name (normalised)"""
    base = [kra_name] + KRA_ALIASES.get(kra_name, [])
    return [norm(a) for a in base]

def kra_row_map(ws, kra_names):
    """
    Map each KRA franchise name → its starting row in this KRA sheet.
    Uses full aliases so short names like HITECH, TMR, JK all match correctly.
    Scans col B for franchise names; matches by alias list.
    """
    result = {}
    # pre-build alias → kra_name lookup
    alias_to_kname = {}
    for kname in kra_names:
        for a in aliases_for(kname):
            alias_to_kname[a] = kname

    for r in range(1, ws.max_row + 1):
        b = str(ws.cell(r, 2).value or "").strip()
        if not b: continue
        bn = norm(b)
        # exact alias match first
        if bn in alias_to_kname:
            kname = alias_to_kname[bn]
            if kname not in result:
                result[kname] = r
            continue
        # partial match: any alias is substring of bn or vice versa
        for alias, kname in alias_to_kname.items():
            if kname in result: continue
            if len(alias) >= 4 and (alias in bn or bn in alias):
                result[kname] = r
                break
        # Overall row
        if any(x in b.lower() for x in ["overall","nandu","total"]):
            result.setdefault("OVERALL", r)

    return result

def wr(ws, r, c, v):
    if ws and r and c:
        try: ws.cell(row=r, column=c, value=v)
        except: pass

def code_lookup(lk, kra_name, fn=safe_int):
    """Look up by franchise code from KRA_CODES map"""
    code = KRA_CODES.get(kra_name)
    if code:
        return fn(lk.get(code, 0))
    return fn(0)

def build_code_lookup(ws, code_col, val_col, skip_rows=1):
    """Build {7-digit-code: value}"""
    lkp = {}
    if not ws: return lkp
    for r in range(1 + skip_rows, ws.max_row + 1):
        code = str(ws.cell(r, code_col).value or "").strip()
        if re.match(r'^\d{7}$', code):
            lkp[code] = ws.cell(r, val_col).value
    return lkp

# ── MAIN PROCESSING ────────────────────────────────────────────────────────────
if btn and all_ready:
    logs = []
    def log(msg, t="info"):
        icon = {"ok":"✅","warn":"⚠️","err":"❌","info":"→"}.get(t,"→")
        logs.append(f'<span class="{t}">{icon} {msg}</span>')

    steps_ph = st.empty()
    STEPS = [
        ("📂","Reading source files"),
        ("🔍","Detecting franchises & codes"),
        ("📊","Building all lookups"),
        ("🗺️","Building KRA row maps"),
        ("✍️","Writing data to KRA tabs"),
        ("🔗","Wiring KRA Sheet dashboard"),
        ("💾","Saving & finalizing"),
    ]
    def render_steps(cur):
        html = '<div class="step-container">'
        for i,(icon,label) in enumerate(STEPS):
            if i < cur:  cls,tick = "done","✓"
            elif i==cur: cls,tick = "running","◉"
            else:        cls,tick = "pending","○"
            html += f'<div class="step-row {cls}"><span class="step-icon">{tick}</span><span>{icon} {label}</span></div>'
        return html + '</div>'

    try:
        # ── STEP 0: Read files ──────────────────────────────────────────────
        steps_ph.markdown(render_steps(0), unsafe_allow_html=True)
        fi  = fy_idx(selected_month); kc = kra_col(selected_month)
        scl = sub_col_letter(selected_month)
        cmc = cluster_month_col(selected_month)
        mcc = mc_closed_col(selected_month)

        kl_wb    = load_workbook(BytesIO(kl_file.read()),    data_only=True)
        param_wb = load_workbook(BytesIO(param_file.read()), data_only=True)
        kra_wb   = load_workbook(BytesIO(kra_file.read()),   keep_vba=False)
        log(f"Loaded | {selected_month} | cluster_col:{cmc} mc_col:{mcc} kra_col:{kc}({scl})", "ok")

        # ── STEP 1: Get KRA franchise list from Call Load tab ───────────────
        steps_ph.markdown(render_steps(1), unsafe_allow_html=True)
        ws_cl_kra = find_sheet(kra_wb, ["call load"])
        kra_franchises = []
        if ws_cl_kra:
            for r in range(1, ws_cl_kra.max_row + 1):
                b = str(ws_cl_kra.cell(r,2).value or "").strip()
                c = str(ws_cl_kra.cell(r,3).value or "").strip()
                if b and c.lower() == "installation" and b.lower() not in ("franchisee",""):
                    kra_franchises.append(b)

        # Validate codes — warn for any KRA franchise not in KRA_CODES
        for kname in kra_franchises:
            if kname not in KRA_CODES:
                log(f"'{kname}' not in KRA_CODES map — add it for correct lookup", "warn")
            elif KRA_CODES[kname] is None:
                log(f"'{kname}' has no franchise code (inactive)", "warn")

        all_names = kra_franchises + ["OVERALL"]
        log(f"{len(kra_franchises)} franchises detected: {kra_franchises}", "ok")

        # ── STEP 2: Build all lookups ───────────────────────────────────────
        steps_ph.markdown(render_steps(2), unsafe_allow_html=True)

        # Param sheets
        ws_ins    = find_sheet(param_wb, ["ins"])
        ws_nr     = find_sheet(param_wb, ["nr"])
        ws_css    = find_sheet(param_wb, ["css"])
        ws_mc_hit = find_sheet(param_wb, ["mc hit"])
        ws_rep    = find_sheet(param_wb, ["rep call","rep calls"])
        ws_sa_p   = find_sheet(param_wb, ["sa prod"])
        ws_amc    = find_sheet(param_wb, ["amc"])

        # Cluster sheets
        ws_mc_reg  = find_sheet(kl_wb, ["mc reg"])
        ws_abv2    = find_sheet(kl_wb, ["abv 2"])
        ws_repeat  = find_sheet(kl_wb, ["repeat call"])
        ws_sa_att  = find_sheet(kl_wb, ["sa attend"])
        ws_social  = find_sheet(kl_wb, ["social"])
        ws_ess_cl  = find_sheet(kl_wb, ["ess bdg"])
        ws_acc_cl  = find_sheet(kl_wb, ["acc bdg"])
        ws_dhukhan = find_sheet(kl_wb, ["apni"])

        # ── All lookups keyed by 7-digit franchise code ──────────────────────
        # INS&SER: code=col1, ins_closed=col3, ins_6hrs=col4, ser_closed=col5, ser_24hrs=col6
        lk_ins_closed = build_code_lookup(ws_ins, 1, 3, skip_rows=2)
        lk_ins_6hrs   = build_code_lookup(ws_ins, 1, 4, skip_rows=2)
        lk_ser_closed = build_code_lookup(ws_ins, 1, 5, skip_rows=2)
        lk_ser_24hrs  = build_code_lookup(ws_ins, 1, 6, skip_rows=2)

        # NR: code=col1, nr_total=col9, nr_closed=col10
        lk_nr_total  = build_code_lookup(ws_nr, 1, 9,  skip_rows=1)
        lk_nr_closed = build_code_lookup(ws_nr, 1, 10, skip_rows=1)

        # CSS: code=col1, css_ok=col3, css_not_ok=col4, css_happy=col5
        lk_css_ok     = build_code_lookup(ws_css, 1, 3, skip_rows=1)
        lk_css_not_ok = build_code_lookup(ws_css, 1, 4, skip_rows=1)
        lk_css_happy  = build_code_lookup(ws_css, 1, 5, skip_rows=1)

        # MC Hit: code=col1, mc_reg=col3, mc_closed=col4
        lk_mc_hit_reg    = build_code_lookup(ws_mc_hit, 1, 3, skip_rows=1)
        lk_mc_hit_closed = build_code_lookup(ws_mc_hit, 1, 4, skip_rows=1)

        # Rep calls: code=col2, rep_total=col6, rep_ticket=col7
        lk_rep_total  = build_code_lookup(ws_rep, 2, 6, skip_rows=1)
        lk_rep_ticket = build_code_lookup(ws_rep, 2, 7, skip_rows=1)

        # SA Prod: code=col2, no_of_sa=col11
        lk_sa_count = build_code_lookup(ws_sa_p, 2, 11, skip_rows=1)

        # AMC: code=col3, amc_tgt=col5, amc_ach=col8, wty_ach=col11, crm_val=col14
        lk_amc_tgt   = build_code_lookup(ws_amc, 3, 5,  skip_rows=2)
        lk_amc_ach   = build_code_lookup(ws_amc, 3, 8,  skip_rows=2)
        lk_wty_ach   = build_code_lookup(ws_amc, 3, 11, skip_rows=2)
        lk_amc_crm_v = build_code_lookup(ws_amc, 3, 14, skip_rows=2)

        # Cluster: code=col1, month col = cmc
        lk_mc_closed = build_code_lookup(ws_mc_reg,  1, mcc, skip_rows=2)
        lk_abv2      = build_code_lookup(ws_abv2,    1, cmc, skip_rows=1)
        lk_repeat_r  = build_code_lookup(ws_repeat,  1, cmc, skip_rows=1)
        lk_sa_att_r  = build_code_lookup(ws_sa_att,  1, cmc, skip_rows=1)
        lk_social    = build_code_lookup(ws_social,  1, cmc, skip_rows=1)
        lk_ess_tgt   = build_code_lookup(ws_ess_cl,  1, 22,  skip_rows=1)
        lk_ess_ach   = build_code_lookup(ws_ess_cl,  1, 23,  skip_rows=1)
        lk_acc_tgt   = build_code_lookup(ws_acc_cl,  1, 22,  skip_rows=1)
        lk_acc_ach   = build_code_lookup(ws_acc_cl,  1, 23,  skip_rows=1)
        lk_dhukhan   = build_code_lookup(ws_dhukhan, 1, cmc, skip_rows=1)

        log("All lookups built (keyed by franchise code)", "ok")

        # ── Build data dict per KRA franchise ────────────────────────────────
        data = {}
        for kname in all_names:
            def g(lk, fn=safe_int):   return code_lookup(lk, kname, fn)
            ins_c  = g(lk_ins_closed)
            ser_c  = g(lk_ser_closed)
            mc_c   = g(lk_mc_hit_closed)
            sa_tot = g(lk_sa_count)
            sa_rat = g(lk_sa_att_r, safe_float)
            data[kname] = {
                "ins_closed":   ins_c,
                "ins_6hrs":     g(lk_ins_6hrs),
                "ser_closed":   ser_c,
                "ser_24hrs":    g(lk_ser_24hrs),
                "mc_hit_closed":mc_c,
                "mc_hit_reg":   g(lk_mc_hit_reg),
                # >2 days: RowA=SER closed, RowB=avg pending
                "ser_closed_2d":ser_c,
                "avg_pend":     round(g(lk_abv2, safe_float), 2),
                # Repeat: RowA=service tickets, RowB=repeat count
                "rep_ticket":   g(lk_rep_ticket),
                "rep_total":    g(lk_rep_total),
                # CSS
                "css_ok":       g(lk_css_ok),
                "css_not_ok":   g(lk_css_not_ok),
                "css_happy":    g(lk_css_happy),
                # NR
                "nr_closed":    g(lk_nr_closed),
                "nr_total":     g(lk_nr_total),
                # Social: RowA=INS+SER+MC, RowB=social media calls
                "total_calls":  ins_c + ser_c + mc_c,
                "social":       g(lk_social),
                # SA Attendance
                "sa_total":     sa_tot,
                "sa_25days":    round(sa_rat * sa_tot),
                # AMC
                "amc_tgt":      g(lk_amc_tgt, safe_float),
                "amc_ach_no":   g(lk_amc_ach) + g(lk_wty_ach),
                "amc_crm_val":  g(lk_amc_crm_v, safe_float),
                # ESS / ACC
                "ess_tgt":      g(lk_ess_tgt, safe_float),
                "ess_ach":      g(lk_ess_ach, safe_float),
                "acc_tgt":      g(lk_acc_tgt, safe_float),
                "acc_ach":      g(lk_acc_ach, safe_float),
                # Exchange = Apni Dhukhan
                "exchange":     g(lk_dhukhan),
            }

        # Log data snapshot for key franchises
        for kname in kra_franchises[:3]:
            d = data[kname]
            log(f"  {kname}: INS={d['ins_closed']} SER={d['ser_closed']} MC={d['mc_hit_closed']} Total={d['total_calls']} Avg2d={d['avg_pend']} Exch={d['exchange']}", "info")

        log(f"Data built for {len(data)} entries", "ok")

        # ── STEP 3: Build KRA row maps ───────────────────────────────────────
        steps_ph.markdown(render_steps(3), unsafe_allow_html=True)

        ws_map = {
            "Call Load":         find_sheet(kra_wb, ["call load"]),
            "Installation":      find_sheet(kra_wb, ["installation"]),
            "Service":           find_sheet(kra_wb, ["service"]),
            ">2 days Pending":   find_sheet(kra_wb, [">2"]),
            "CSS":               find_sheet(kra_wb, ["css"]),
            "Negative Response": find_sheet(kra_wb, ["negative"]),
            "Social M Calls":    find_sheet(kra_wb, ["social"]),
            "Repeat Calls":      find_sheet(kra_wb, ["repeat"]),
            "MC Calls":          find_sheet(kra_wb, ["mc call"]),
            "SA Attendance":     find_sheet(kra_wb, ["sa attend"]),
            "AMC Achievement":   find_sheet(kra_wb, ["amc achiev"]),
            "Essential Budget":  find_sheet(kra_wb, ["essential"]),
            "Accesories Budget": find_sheet(kra_wb, ["accesories","accessories"]),
            "Exchange":          find_sheet(kra_wb, ["exchange"]),
        }
        missing_ws = [k for k,v in ws_map.items() if not v]
        if missing_ws: log(f"Missing KRA sheets: {missing_ws}", "warn")

        kra_rmaps = {tab: kra_row_map(ws, all_names) for tab,ws in ws_map.items() if ws}

        # Log match quality
        for tab, rmap in kra_rmaps.items():
            matched   = [k for k in kra_franchises if k in rmap]
            unmatched = [k for k in kra_franchises if k not in rmap]
            if unmatched:
                log(f"'{tab}': unmatched → {unmatched}", "warn")
            else:
                log(f"'{tab}': all {len(matched)} franchises matched ✓", "ok")

        log(f"Row maps built for {len(kra_rmaps)} tabs", "ok")

        # ── STEP 4: Write data ───────────────────────────────────────────────
        steps_ph.markdown(render_steps(4), unsafe_allow_html=True)

        updated = []
        for kname in all_names:
            d = data[kname]
            for tab, ws in ws_map.items():
                if not ws: continue
                sr = kra_rmaps.get(tab, {}).get(kname)
                if not sr: continue
                col = kc

                if tab == "Call Load":
                    wr(ws, sr,   col, d["ins_closed"])
                    wr(ws, sr+1, col, d["ser_closed"])
                    wr(ws, sr+2, col, d["mc_hit_closed"])

                elif tab == "Installation":
                    wr(ws, sr,   col, d["ins_closed"])
                    wr(ws, sr+1, col, d["ins_6hrs"])

                elif tab == "Service":
                    wr(ws, sr,   col, d["ser_closed"])
                    wr(ws, sr+1, col, d["ser_24hrs"])

                elif tab == ">2 days Pending":
                    wr(ws, sr,   col, d["ser_closed_2d"])
                    wr(ws, sr+1, col, d["avg_pend"])
                    # Row C = % formula — do NOT touch

                elif tab == "Repeat Calls":
                    wr(ws, sr,   col, d["rep_ticket"])
                    wr(ws, sr+1, col, d["rep_total"])
                    # Row C = % formula — do NOT touch

                elif tab == "SA Attendance":
                    wr(ws, sr,   col, d["sa_total"])
                    wr(ws, sr+1, col, d["sa_25days"])
                    # Row C = % formula — do NOT touch

                elif tab == "CSS":
                    wr(ws, sr,   col, d["css_happy"])
                    wr(ws, sr+1, col, d["css_ok"])
                    wr(ws, sr+2, col, d["css_not_ok"])

                elif tab == "Negative Response":
                    wr(ws, sr,   col, d["nr_closed"])
                    wr(ws, sr+1, col, d["nr_total"])

                elif tab == "Social M Calls":
                    # Row A = INS + SER + MC (total calls closed)
                    # Row B = Social Media calls from cluster
                    # Row C = % formula — do NOT touch
                    wr(ws, sr,   col, d["total_calls"])
                    wr(ws, sr+1, col, d["social"])

                elif tab == "MC Calls":
                    wr(ws, sr,   col, d["mc_hit_closed"])
                    wr(ws, sr+1, col, d["mc_hit_reg"])

                elif tab == "AMC Achievement":
                    wr(ws, sr,   col, d["amc_tgt"])
                    wr(ws, sr+1, col, d["amc_ach_no"])
                    wr(ws, sr+4, col, d["amc_crm_val"])

                elif tab == "Essential Budget":
                    wr(ws, sr,   col, d["ess_tgt"])
                    wr(ws, sr+1, col, d["ess_ach"])

                elif tab == "Accesories Budget":
                    wr(ws, sr,   col, d["acc_tgt"])
                    wr(ws, sr+1, col, d["acc_ach"])

                elif tab == "Exchange":
                    # Single row — Apni Dhukhan Jan value
                    wr(ws, sr, col, d["exchange"])

                if tab not in updated: updated.append(tab)

        log(f"Written to {len(updated)} tabs", "ok")

        # ── STEP 5: Wire KRA Sheet dashboard ────────────────────────────────
        steps_ph.markdown(render_steps(5), unsafe_allow_html=True)
        ws_dash = find_sheet(kra_wb, ["kra sheet"])
        if ws_dash:
            c = scl
            formulas = {
                9:  f"='Call Load'!{c}59",    10: f"='Call Load'!{c}60",
                11: f"='Call Load'!{c}61",    12: f"='Call Load'!{c}62",
                13: f"=Installation!{c}39",   14: f"=Service!{c}39",
                15: f"='>2 days Pending'!{c}39",
                16: f"=CSS!{c}50",
                17: f"='Negative Response'!{c}39",
                18: f"='Social M Calls'!{c}38",
                20: f"='Repeat Calls'!{c}39",
                21: f"='MC Calls'!{c}39",
                24: f"='AMC Achievement'!{c}72",
                25: f"='AMC Achievement'!{c}75",
                26: f"='AMC HIT Rate'!{c}78",
                27: f"='AMC HIT Rate'!{c}81",
                28: f"='Essential Budget'!{c}45",
                29: f"='Accesories Budget'!{c}48",
                30: f"='Spare Cosnumption'!{c}104",
            }
            dash_col = kra_col(selected_month)
            for row_num, formula in formulas.items():
                ws_dash.cell(row=row_num, column=dash_col).value = formula
            log(f"Dashboard wired: {len(formulas)} rows → col {c}", "ok")
        else:
            log("KRA Sheet tab not found", "warn")

        # ── STEP 6: Save ─────────────────────────────────────────────────────
        steps_ph.markdown(render_steps(6), unsafe_allow_html=True)
        out = BytesIO()
        kra_wb.save(out); out.seek(0)
        log("Saved ✓", "ok")
        steps_ph.markdown(render_steps(7), unsafe_allow_html=True)

        # ── RESULTS ──────────────────────────────────────────────────────────
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,rgba(34,197,94,0.1),rgba(5,150,105,0.05));
             border:1px solid rgba(34,197,94,0.3);border-radius:16px;padding:20px 24px;
             margin:20px 0;display:flex;align-items:center;gap:14px;">
          <div style="font-size:28px;">🎉</div>
          <div>
            <div style="font-size:17px;font-weight:700;color:#4ADE80;">KRA Successfully Updated!</div>
            <div style="font-size:13px;color:#6EE7B7;margin-top:3px;">
              {selected_month} 2026 &nbsp;·&nbsp; {len(kra_franchises)} franchises &nbsp;·&nbsp; {len(updated)} tabs written
            </div>
          </div>
        </div>""", unsafe_allow_html=True)

        ov = data.get("OVERALL",{})
        st.markdown(f"""
        <div class="metric-grid">
          <div class="metric-tile"><div class="metric-val">{len(kra_franchises)}</div><div class="metric-lbl">Franchises</div></div>
          <div class="metric-tile"><div class="metric-val">{len(updated)}</div><div class="metric-lbl">Tabs Updated</div></div>
          <div class="metric-tile"><div class="metric-val">{ov.get('ins_closed',0)}</div><div class="metric-lbl">INS Closed</div></div>
          <div class="metric-tile"><div class="metric-val">{ov.get('ser_closed',0)}</div><div class="metric-lbl">SER Closed</div></div>
        </div>""", unsafe_allow_html=True)

        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">Franchise Data Preview</p>', unsafe_allow_html=True)
        rows_list = []
        for kname in kra_franchises:
            d = data.get(kname,{})
            code = KRA_CODES.get(kname,"—")
            rows_list.append({
                "Franchise": kname, "Code": code,
                "INS": d.get("ins_closed",0),     "INS 6H": d.get("ins_6hrs",0),
                "SER": d.get("ser_closed",0),     "SER 24H": d.get("ser_24hrs",0),
                "MC":  d.get("mc_hit_closed",0),  "Total Calls": d.get("total_calls",0),
                "Avg >2d": d.get("avg_pend",0),   "Repeat": d.get("rep_total",0),
                "SA Tot": d.get("sa_total",0),    "SA≥25d": d.get("sa_25days",0),
                "CSS😊": d.get("css_happy",0),    "Social": d.get("social",0),
                "Exchange": d.get("exchange",0),  "ESS Ach": d.get("ess_ach",0),
            })
        st.dataframe(pd.DataFrame(rows_list), use_container_width=True, height=380)

        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">Processing Log</p>', unsafe_allow_html=True)
        st.markdown(f'<div class="log-area">{"<br>".join(logs)}</div>', unsafe_allow_html=True)

        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">④ Download</p>', unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️  Download KRA — {selected_month} 2026",
            data=out, file_name=f"KRA-{selected_month}-2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        import traceback
        steps_ph.empty()
        st.error(f"❌ {e}")
        st.code(traceback.format_exc())

if not btn and not all_ready:
    st.markdown("""
    <div style="background:linear-gradient(145deg,#0F1729,#131E35);
         border:1px dashed #1E3A5F;border-radius:20px;padding:60px 40px;text-align:center;margin-top:8px;">
      <div style="font-size:48px;margin-bottom:16px;">📋</div>
      <h3 style="color:#4B5563;font-weight:600;margin:0 0 10px;">Ready to Begin</h3>
      <p style="color:#374151;font-size:14px;max-width:400px;margin:0 auto;line-height:1.7;">Upload 3 files, select month, click Generate</p>
    </div>""", unsafe_allow_html=True)
