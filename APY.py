import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import time

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
    border: 1px solid #1E3A5F; border-radius: 16px;
    padding: 22px 20px; margin-bottom: 4px; transition: border-color 0.3s;
}
.upload-card:hover { border-color: #3B82F6; }
.card-title {
    font-size: 13px; font-weight: 600; color: #64B5F6 !important;
    letter-spacing: 0.8px; text-transform: uppercase; margin-bottom: 10px;
}
.file-ok {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(34,197,94,0.12); border: 1px solid rgba(34,197,94,0.3);
    border-radius: 8px; padding: 6px 12px;
    font-size: 12px; color: #22C55E !important; margin-top: 8px; font-weight: 500;
}
div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #2563EB, #7C3AED) !important;
    color: white !important; border: none !important; border-radius: 12px !important;
    padding: 14px 40px !important; font-size: 16px !important; font-weight: 600 !important;
    width: 100% !important; letter-spacing: 0.3px;
    box-shadow: 0 4px 20px rgba(37,99,235,0.35) !important; transition: all 0.2s !important;
}
div[data-testid="stButton"] > button:hover {
    transform: translateY(-1px); box-shadow: 0 6px 28px rgba(37,99,235,0.5) !important;
}
div[data-testid="stButton"] > button:disabled {
    background: linear-gradient(135deg, #1e3a5f, #2d1b5e) !important;
    color: #4B5563 !important; transform: none !important; box-shadow: none !important;
}
.step-container {
    display: flex; flex-direction: column; gap: 10px; padding: 20px;
    background: #0D1526; border-radius: 14px; border: 1px solid #1E3A5F; margin: 16px 0;
}
.step-row {
    display: flex; align-items: center; gap: 14px; padding: 10px 14px;
    border-radius: 8px; font-size: 14px; font-weight: 500;
}
.step-row.done    { background: rgba(34,197,94,0.08);  color: #4ADE80 !important; }
.step-row.running { background: rgba(59,130,246,0.12); color: #60A5FA !important; }
.step-row.pending { color: #4B5563 !important; }
.step-icon { font-size: 16px; min-width: 22px; }
.metric-grid {
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin: 20px 0;
}
.metric-tile {
    background: linear-gradient(145deg, #0F1729, #131E35);
    border: 1px solid #1E3A5F; border-radius: 14px; padding: 20px; text-align: center;
}
.metric-val { font-size: 32px; font-weight: 700; color: #60A5FA !important; }
.metric-lbl { font-size: 12px; color: #64748B !important; margin-top: 4px; font-weight: 500; letter-spacing: 0.5px; }
[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
[data-testid="stDataFrame"] table { background: #0F1729; }
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #059669, #0D9488) !important;
    color: white !important; border: none !important; border-radius: 12px !important;
    padding: 14px 40px !important; font-size: 16px !important; font-weight: 600 !important;
    width: 100% !important; box-shadow: 0 4px 20px rgba(5,150,105,0.35) !important;
    letter-spacing: 0.3px; margin-top: 8px;
}
[data-testid="stSelectbox"] > div {
    background: #0F1729 !important; border-color: #1E3A5F !important; border-radius: 10px !important;
}
.stAlert { border-radius: 10px !important; }
.log-area {
    background: #060B18; border: 1px solid #1E3A5F; border-radius: 10px;
    padding: 16px 20px; font-family: monospace !important;
    font-size: 12px; color: #475569; max-height: 280px; overflow-y: auto; line-height: 1.8;
}
.log-area .ok   { color: #22C55E; }
.log-area .warn { color: #F59E0B; }
.log-area .err  { color: #EF4444; }
.log-area .info { color: #60A5FA; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:linear-gradient(135deg,#0F1729 0%,#131E35 50%,#0F1729 100%);
     border:1px solid #1E3A5F; border-radius:20px; padding:36px 40px 28px; margin-bottom:28px;
     position:relative; overflow:hidden;">
  <div style="position:absolute;top:-40px;right:-40px;width:180px;height:180px;
       background:radial-gradient(circle,rgba(59,130,246,0.15),transparent 70%);border-radius:50%;"></div>
  <div style="display:flex;align-items:center;gap:18px;margin-bottom:10px;">
    <div style="background:linear-gradient(135deg,#2563EB,#7C3AED);border-radius:14px;
         padding:12px;font-size:26px;line-height:1;">📊</div>
    <div>
      <h1 style="margin:0;font-size:26px;font-weight:700;letter-spacing:-0.5px;">KRA Auto-Updater</h1>
      <p style="margin:4px 0 0;color:#64748B;font-size:14px;">IFB Service · Cochin Cluster · Nandu S Kumar</p>
    </div>
  </div>
  <p style="color:#94A3B8;font-size:14px;margin:0;max-width:560px;line-height:1.6;">
    Upload your monthly reports and the KRA template — fills all KRA tabs automatically.
  </p>
</div>
""", unsafe_allow_html=True)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

def fy_idx(m):         return (MONTHS.index(m) - 3) % 12
def kra_col(m):        return 4 + fy_idx(m)
def sub_col_letter(m): return chr(ord('D') + fy_idx(m))
def kra_dash_col(m):   return 4 + fy_idx(m)
def og_cols(m):
    base = 4 + fy_idx(m) * 3
    return base, base+1, base+2

# ── FILE UPLOADS ──────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin-bottom:12px;">① Upload Input Files</p>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="upload-card"><p class="card-title">📁 KL Cluster Report</p>', unsafe_allow_html=True)
    kl_file = st.file_uploader("kl", type=["xlsx"], key="kl", label_visibility="collapsed")
    if kl_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{kl_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">OG Calls · MC · Abv2Days · Repeat · SA · ESS · ACC · Social · Apni Dhukhan</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c2:
    st.markdown('<div class="upload-card"><p class="card-title">📁 Parameter Dashboard</p>', unsafe_allow_html=True)
    param_file = st.file_uploader("param", type=["xlsx"], key="param", label_visibility="collapsed")
    if param_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{param_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">INS·SER·CSS·NR·MC Hit·SA Prod·AMC</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with c3:
    st.markdown('<div class="upload-card"><p class="card-title">📁 KRA Template</p>', unsafe_allow_html=True)
    kra_file = st.file_uploader("kra", type=["xlsx"], key="kra", label_visibility="collapsed")
    if kra_file: st.markdown(f'<div class="file-ok">✓ &nbsp;{kra_file.name}</div>', unsafe_allow_html=True)
    else: st.markdown('<p style="color:#374151;font-size:12px;margin-top:4px;">KRA NANDU template</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ── MONTH SELECTOR ────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 10px;">② Select Month</p>', unsafe_allow_html=True)

mc1, mc2 = st.columns([1, 3])
with mc1:
    selected_month = st.selectbox("Month", MONTHS, index=0, label_visibility="collapsed")
with mc2:
    fi  = fy_idx(selected_month)
    _, oc, _ = og_cols(selected_month)
    kc  = kra_col(selected_month)
    scl = sub_col_letter(selected_month)
    st.markdown(f"""
    <div style="background:#0F1729;border:1px solid #1E3A5F;border-radius:12px;
         padding:14px 20px;display:flex;gap:32px;align-items:center;">
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">FY Index</div>
           <div style="font-size:20px;font-weight:700;color:#60A5FA;">Month {fi+1} of 12</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">OG Closed Col</div>
           <div style="font-size:20px;font-weight:700;color:#A78BFA;">Col {oc} (0-idx)</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">KRA Write Col</div>
           <div style="font-size:20px;font-weight:700;color:#34D399;">Col {kc} (1-base)</div></div>
      <div><div style="font-size:11px;color:#475569;text-transform:uppercase;">Sub-sheet Col</div>
           <div style="font-size:20px;font-weight:700;color:#FB923C;">{scl} → Apr–Mar ✓</div></div>
    </div>
    """, unsafe_allow_html=True)

# ── PROCESS BUTTON ────────────────────────────────────────────────────────────
st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 10px;">③ Process</p>', unsafe_allow_html=True)

all_ready = kl_file and param_file and kra_file
if not all_ready:
    missing = []
    if not kl_file:    missing.append("KL Cluster Report")
    if not param_file: missing.append("Parameter Dashboard")
    if not kra_file:   missing.append("KRA Template")
    st.markdown(f'<div style="background:rgba(251,146,60,0.08);border:1px solid rgba(251,146,60,0.3);border-radius:10px;padding:12px 18px;color:#FB923C;font-size:13px;margin-bottom:12px;">⚠️ Still needed: {" · ".join(missing)}</div>', unsafe_allow_html=True)

btn = st.button("⚡  Generate Updated KRA", disabled=not all_ready)

# ── HELPERS ───────────────────────────────────────────────────────────────────
def safe_int(v):
    try: return int(float(v)) if v is not None and str(v) not in ('nan','None','—') else 0
    except: return 0

def safe_float(v):
    try: return float(v) if v is not None and str(v) not in ('nan','None','—') else 0.0
    except: return 0.0

def build_lookup(ws, code_col=0, val_col=12, skip_rows=1):
    """Fast openpyxl-based lookup: {7digit_code: value}"""
    lkp = {}
    for i, row in enumerate(ws.iter_rows(min_row=1+skip_rows, values_only=True)):
        code = str(row[code_col]).strip() if row[code_col] else ""
        if re.match(r'^\d{7}$', code):
            try: lkp[code] = row[val_col]
            except: lkp[code] = 0
    return lkp

def find_sheet(wb, keywords):
    for name in wb.sheetnames:
        if all(k.lower() in name.lower() for k in keywords): return wb[name]
    for name in wb.sheetnames:
        if any(k.lower() in name.lower() for k in keywords): return wb[name]
    return None

def kra_name_map(ws, franchises):
    lkp = {n.lower().replace(" ","").replace(".","").replace("/","")[:10]: c
           for c,(n,_) in franchises.items()}
    m = {}
    for r in range(1, 300):
        cv = str(ws.cell(row=r, column=2).value or "").strip()
        k  = cv.lower().replace(" ","").replace(".","").replace("/","")[:10]
        if k in lkp and k and lkp[k] not in m:
            m[lkp[k]] = r
    for r in range(1, 300):
        v = str(ws.cell(row=r, column=2).value or "").strip().lower()
        if any(x in v for x in ["overall","nandu","total","cluster","manu"]):
            m.setdefault("OVERALL", r)
    return m

def wr(ws, r, c, v):
    if ws and r and c:
        try: ws.cell(row=r, column=c, value=v)
        except: pass

# ── MAIN PROCESSING ───────────────────────────────────────────────────────────
if btn and all_ready:
    logs = []
    def log(msg, t="info"):
        icon = {"ok":"✅","warn":"⚠️","err":"❌","info":"→"}.get(t,"→")
        logs.append(f'<span class="{t}">{icon} {msg}</span>')

    steps_ph = st.empty()
    STEPS = [
        ("📂", "Reading source files"),
        ("🔍", "Detecting franchises & building lookups"),
        ("📊", "Extracting parameter data"),
        ("🗺️", "Building KRA row maps"),
        ("✍️", "Writing data to KRA tabs"),
        ("🔗", "Wiring KRA Sheet dashboard"),
        ("💾", "Saving & finalizing"),
    ]

    def render_steps(current):
        html = '<div class="step-container">'
        for i,(icon,label) in enumerate(STEPS):
            if i < current:  cls,tick = "done","✓"
            elif i==current: cls,tick = "running","◉"
            else:            cls,tick = "pending","○"
            html += f'<div class="step-row {cls}"><span class="step-icon">{tick}</span><span>{icon} {label}</span></div>'
        html += '</div>'
        return html

    try:
        # ── STEP 0: Read ALL files with openpyxl (fast, no pandas read_excel) ──
        steps_ph.markdown(render_steps(0), unsafe_allow_html=True)

        fi  = fy_idx(selected_month)
        kc  = kra_col(selected_month)
        scl = sub_col_letter(selected_month)
        _, oc, op = og_cols(selected_month)

        # Month column indices (0-based)
        MONTH_COL     = 3 + fi          # April=3,May=4...Jan=12,Feb=13,Mar=14
        MC_CLOSED_COL = 2 + fi * 2 + 1  # MC Reg&Closed: April_CLOSED=3,May=5...Jan=23

        log(f"Month:{selected_month} FY:{fi} month_col:{MONTH_COL} mc_closed_col:{MC_CLOSED_COL} kra_col:{kc}", "info")

        kl_wb    = load_workbook(BytesIO(kl_file.read()),    data_only=True)
        param_wb = load_workbook(BytesIO(param_file.read()), data_only=True)
        kra_wb   = load_workbook(BytesIO(kra_file.read()),   keep_vba=False)

        log("Workbooks loaded", "ok")

        # ── STEP 1: Detect franchises + build all lookups at once ──────────
        steps_ph.markdown(render_steps(1), unsafe_allow_html=True)

        # Detect franchises from OG Calls (cluster)
        ws_og = find_sheet(kl_wb, ["og call","og calls"])
        if not ws_og: ws_og = find_sheet(kl_wb, ["og"])
        franchises = {}
        for row in ws_og.iter_rows(min_row=2, values_only=True):
            code = str(row[0]).strip() if row[0] else ""
            if re.match(r'^\d{7}$', code):
                franchises[code] = (str(row[1]).strip(), None)

        log(f"{len(franchises)} franchises from OG Calls", "ok")

        # ── ALL cluster lookups in one pass each ───────────────────────────
        # OG Calls: closed=oc, pending=op
        lk_og_closed  = build_lookup(ws_og,                                    code_col=0, val_col=oc,            skip_rows=1)
        lk_og_pending = build_lookup(ws_og,                                    code_col=0, val_col=op,            skip_rows=1)
        lk_mc_closed  = build_lookup(find_sheet(kl_wb, ["mc reg"]),            code_col=0, val_col=MC_CLOSED_COL, skip_rows=2)
        lk_abv2       = build_lookup(find_sheet(kl_wb, ["abv 2"]),             code_col=0, val_col=MONTH_COL,     skip_rows=1)
        lk_repeat     = build_lookup(find_sheet(kl_wb, ["repeat call"]),       code_col=0, val_col=MONTH_COL,     skip_rows=1)
        lk_sa_att     = build_lookup(find_sheet(kl_wb, ["sa attend"]),         code_col=0, val_col=MONTH_COL,     skip_rows=1)
        lk_social     = build_lookup(find_sheet(kl_wb, ["social"]),            code_col=0, val_col=MONTH_COL,     skip_rows=1)
        lk_ess_tgt    = build_lookup(find_sheet(kl_wb, ["ess bdg"]),           code_col=0, val_col=21,            skip_rows=1)
        lk_ess_ach    = build_lookup(find_sheet(kl_wb, ["ess bdg"]),           code_col=0, val_col=22,            skip_rows=1)
        lk_acc_tgt    = build_lookup(find_sheet(kl_wb, ["acc bdg"]),           code_col=0, val_col=21,            skip_rows=1)
        lk_acc_ach    = build_lookup(find_sheet(kl_wb, ["acc bdg"]),           code_col=0, val_col=22,            skip_rows=1)
        lk_dhukhan    = build_lookup(find_sheet(kl_wb, ["apni"]),              code_col=0, val_col=MONTH_COL,     skip_rows=1)

        log("Cluster lookups built", "ok")

        # ── ALL param lookups in one pass each ─────────────────────────────
        ws_ins = find_sheet(param_wb, ["ins"])
        ws_nr  = find_sheet(param_wb, ["nr"])
        ws_css = find_sheet(param_wb, ["css"])
        ws_rep = find_sheet(param_wb, ["rep"])
        ws_mc  = find_sheet(param_wb, ["mc hit","mc"])
        ws_sa  = find_sheet(param_wb, ["sa prod"])
        ws_amc = find_sheet(param_wb, ["amc"])

        # Build param lookups: match by franchise code (col0 or col1) → value
        def build_param_lookup(ws, code_col, val_col):
            lkp = {}
            if not ws: return lkp
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) <= max(code_col, val_col): continue
                code = str(row[code_col]).strip() if row[code_col] else ""
                if re.match(r'^\d{7}$', code):
                    try: lkp[code] = row[val_col]
                    except: lkp[code] = 0
            return lkp

        # INS & SER: code=col0, ins_closed=col2, ins_6hrs=col3, ser_closed=col4, ser_24hrs=col5
        lk_ins_closed = build_param_lookup(ws_ins, 0, 2)
        lk_ins_6hrs   = build_param_lookup(ws_ins, 0, 3)
        lk_ser_closed = build_param_lookup(ws_ins, 0, 4)
        lk_ser_24hrs  = build_param_lookup(ws_ins, 0, 5)

        # NR: code=col0, nr_calls=col9, nr_neg=col8
        lk_nr_calls = build_param_lookup(ws_nr, 0, 9)
        lk_nr_neg   = build_param_lookup(ws_nr, 0, 8)

        # CSS: code=col0, css_ok=col2, css_not_ok=col3, css_happy=col4
        lk_css_ok     = build_param_lookup(ws_css, 0, 2)
        lk_css_not_ok = build_param_lookup(ws_css, 0, 3)
        lk_css_happy  = build_param_lookup(ws_css, 0, 4)

        # Rep calls: code=col0, rep_closed=col6, rep_count=col4
        lk_rep_closed = build_param_lookup(ws_rep, 0, 6)
        lk_rep_count  = build_param_lookup(ws_rep, 0, 4)

        # MC Hit: code=col0, mc_reg=col2, mc_closed=col3
        lk_mc_reg    = build_param_lookup(ws_mc, 0, 2)
        lk_mc_p_closed = build_param_lookup(ws_mc, 0, 3)

        # SA Prod: code=col1, no_of_sa=col10
        lk_sa_count = build_param_lookup(ws_sa, 1, 10)

        # AMC: code=col0, amc_tgt=col4, amc_nos=col7+col10, amc_val=col13
        lk_amc_tgt = build_param_lookup(ws_amc, 0, 4)
        lk_amc_7   = build_param_lookup(ws_amc, 0, 7)
        lk_amc_10  = build_param_lookup(ws_amc, 0, 10)
        lk_amc_val = build_param_lookup(ws_amc, 0, 13)

        log("Param lookups built", "ok")

        # ── STEP 2: Build data dict (pure dict lookups — instant) ──────────
        steps_ph.markdown(render_steps(2), unsafe_allow_html=True)

        all_codes = list(franchises.keys()) + ["OVERALL"]
        data = {}
        for code in all_codes:
            mc_c   = safe_int(lk_mc_closed.get(code, 0))
            sa_tot = safe_int(lk_sa_count.get(code, 0))
            data[code] = {
                "ins_closed":  safe_int(lk_ins_closed.get(code, 0)),
                "ins_6hrs":    safe_int(lk_ins_6hrs.get(code, 0)),
                "ser_closed":  safe_int(lk_ser_closed.get(code, 0)),
                "ser_24hrs":   safe_int(lk_ser_24hrs.get(code, 0)),
                "og_closed":   safe_int(lk_og_closed.get(code, 0)),
                "og_pending":  safe_int(lk_og_pending.get(code, 0)),
                "mc_closed":   mc_c,
                "avg_pend":    round(safe_float(lk_abv2.get(code, 0)), 2),
                "rep_count":   round(safe_float(lk_repeat.get(code, 0)) * mc_c),
                "sa_total":    sa_tot,
                "sa_25days":   round(safe_float(lk_sa_att.get(code, 0)) * sa_tot),
                "css_happy":   safe_int(lk_css_happy.get(code, 0)),
                "css_ok":      safe_int(lk_css_ok.get(code, 0)),
                "css_not_ok":  safe_int(lk_css_not_ok.get(code, 0)),
                "nr_calls":    safe_int(lk_nr_calls.get(code, 0)),
                "nr_neg":      safe_int(lk_nr_neg.get(code, 0)),
                "social":      safe_int(lk_social.get(code, 0)),
                "rep_closed":  safe_int(lk_rep_closed.get(code, 0)),
                "mc_reg":      safe_int(lk_mc_reg.get(code, 0)),
                "amc_target":  safe_float(lk_amc_tgt.get(code, 0)),
                "amc_nos":     safe_int(lk_amc_7.get(code, 0)) + safe_int(lk_amc_10.get(code, 0)),
                "amc_val":     safe_float(lk_amc_val.get(code, 0)),
                "ess_tgt":     safe_float(lk_ess_tgt.get(code, 0)),
                "ess_ach":     safe_float(lk_ess_ach.get(code, 0)),
                "acc_tgt":     safe_float(lk_acc_tgt.get(code, 0)),
                "acc_ach":     safe_float(lk_acc_ach.get(code, 0)),
                "exchange":    safe_int(lk_dhukhan.get(code, 0)),
            }

        log(f"Data built for {len(data)} entries", "ok")

        # ── STEP 3: Build KRA row maps ──────────────────────────────────────
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

        kra_rmaps = {tab: kra_name_map(ws, franchises)
                     for tab, ws in ws_map.items() if ws}
        log(f"KRA row maps built for {len(kra_rmaps)} tabs", "ok")

        # ── STEP 4: Write data ──────────────────────────────────────────────
        steps_ph.markdown(render_steps(4), unsafe_allow_html=True)

        updated, skipped = [], {}
        for code in all_codes:
            d = data[code]
            for tab, ws in ws_map.items():
                if not ws: continue
                sr = kra_rmaps.get(tab, {}).get(code)
                if not sr:
                    skipped.setdefault(tab, 0); skipped[tab] += 1
                    continue
                col = kc

                if   tab == "Call Load":
                    wr(ws, sr,   col, d["ins_closed"])
                    wr(ws, sr+1, col, d["ser_closed"])
                    wr(ws, sr+2, col, d["mc_closed"])
                elif tab == "Installation":
                    wr(ws, sr,   col, d["ins_closed"])
                    wr(ws, sr+1, col, d["ins_6hrs"])
                elif tab == "Service":
                    wr(ws, sr,   col, d["ser_closed"])
                    wr(ws, sr+1, col, d["ser_24hrs"])
                elif tab == ">2 days Pending":
                    wr(ws, sr,   col, d["mc_closed"])
                    wr(ws, sr+1, col, d["avg_pend"])
                elif tab == "Repeat Calls":
                    wr(ws, sr,   col, d["mc_closed"])
                    wr(ws, sr+1, col, d["rep_count"])
                elif tab == "SA Attendance":
                    wr(ws, sr,   col, d["sa_total"])
                    wr(ws, sr+1, col, d["sa_25days"])
                elif tab == "CSS":
                    wr(ws, sr,   col, d["css_happy"])
                    wr(ws, sr+1, col, d["css_ok"])
                    wr(ws, sr+2, col, d["css_not_ok"])
                elif tab == "Negative Response":
                    wr(ws, sr,   col, d["nr_calls"])
                    wr(ws, sr+1, col, d["nr_neg"])
                elif tab == "Social M Calls":
                    wr(ws, sr,   col, d["og_closed"])
                    wr(ws, sr+1, col, d["social"])
                elif tab == "MC Calls":
                    wr(ws, sr,   col, d["mc_closed"])
                    wr(ws, sr+1, col, d["mc_reg"])
                elif tab == "AMC Achievement":
                    wr(ws, sr,   col, d["amc_target"])
                    wr(ws, sr+1, col, d["amc_nos"])
                    wr(ws, sr+4, col, d["amc_val"])
                elif tab == "Essential Budget":
                    wr(ws, sr,   col, d["ess_tgt"])
                    wr(ws, sr+1, col, d["ess_ach"])
                elif tab == "Accesories Budget":
                    wr(ws, sr,   col, d["acc_tgt"])
                    wr(ws, sr+1, col, d["acc_ach"])
                elif tab == "Exchange":
                    wr(ws, sr,   col, d["exchange"])

                if tab not in updated: updated.append(tab)

        for tab, cnt in skipped.items():
            log(f"{cnt} unmatched in '{tab}'", "warn")
        log(f"Written to {len(updated)} tabs", "ok")

        # ── STEP 5: KRA Sheet dashboard formulas ───────────────────────────
        steps_ph.markdown(render_steps(5), unsafe_allow_html=True)

        ws_dash = find_sheet(kra_wb, ["kra sheet"])
        if ws_dash:
            dash_col = kra_dash_col(selected_month)
            c = scl
            formulas = {
                9:  f"='Call Load'!{c}59",
                10: f"='Call Load'!{c}60",
                11: f"='Call Load'!{c}61",
                12: f"='Call Load'!{c}62",
                13: f"=Installation!{c}39",
                14: f"=Service!{c}39",
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
            for row_num, formula in formulas.items():
                ws_dash.cell(row=row_num, column=dash_col).value = formula
            log(f"Dashboard wired: {len(formulas)} rows → col {c}", "ok")
        else:
            log("KRA Sheet not found", "warn")

        # ── STEP 6: Save ───────────────────────────────────────────────────
        steps_ph.markdown(render_steps(6), unsafe_allow_html=True)

        out = BytesIO()
        kra_wb.save(out)
        out.seek(0)
        log("Saved ✓", "ok")
        steps_ph.markdown(render_steps(7), unsafe_allow_html=True)

        # ── RESULTS ────────────────────────────────────────────────────────
        ov = data.get("OVERALL", {})
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,rgba(34,197,94,0.1),rgba(5,150,105,0.05));
             border:1px solid rgba(34,197,94,0.3);border-radius:16px;padding:20px 24px;
             margin:20px 0;display:flex;align-items:center;gap:14px;">
          <div style="font-size:28px;">🎉</div>
          <div>
            <div style="font-size:17px;font-weight:700;color:#4ADE80;">KRA Successfully Updated!</div>
            <div style="font-size:13px;color:#6EE7B7;margin-top:3px;">
              {selected_month} 2026 &nbsp;·&nbsp; {len(franchises)} franchises &nbsp;·&nbsp;
              {len(updated)} tabs written &nbsp;·&nbsp; KRA Sheet dashboard ✓
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="metric-grid">
          <div class="metric-tile"><div class="metric-val">{len(franchises)}</div>
            <div class="metric-lbl">Franchises Mapped</div></div>
          <div class="metric-tile"><div class="metric-val">{len(updated)}</div>
            <div class="metric-lbl">KRA Tabs Updated</div></div>
          <div class="metric-tile"><div class="metric-val">{ov.get('ins_closed',0)}</div>
            <div class="metric-lbl">Total INS Closed</div></div>
          <div class="metric-tile"><div class="metric-val">{ov.get('ser_closed',0)}</div>
            <div class="metric-lbl">Total SER Closed</div></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">Franchise Data Preview</p>', unsafe_allow_html=True)
        rows_list = []
        for code,(name,_) in franchises.items():
            d = data.get(code,{})
            rows_list.append({
                "Code": code, "Franchise": name,
                "INS":      d.get("ins_closed",0),
                "INS 6H":   d.get("ins_6hrs",0),
                "SER":      d.get("ser_closed",0),
                "SER 24H":  d.get("ser_24hrs",0),
                "MC Closed":d.get("mc_closed",0),
                "Avg Pend": d.get("avg_pend",0),
                "Repeat":   d.get("rep_count",0),
                "SA Total": d.get("sa_total",0),
                "SA ≥25d":  d.get("sa_25days",0),
                "CSS Happy":d.get("css_happy",0),
                "NR Neg":   d.get("nr_neg",0),
                "AMC Nos":  d.get("amc_nos",0),
                "Exchange": d.get("exchange",0),
                "ESS Ach":  d.get("ess_ach",0),
                "ACC Ach":  d.get("acc_ach",0),
            })
        st.dataframe(pd.DataFrame(rows_list), use_container_width=True, height=380)

        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">Processing Log</p>', unsafe_allow_html=True)
        st.markdown(f'<div class="log-area">{"<br>".join(logs)}</div>', unsafe_allow_html=True)

        clean_name = f"KRA-{selected_month}-2026.xlsx"
        st.markdown('<p style="font-size:13px;color:#64748B;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;margin:20px 0 8px;">④ Download</p>', unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️  Download  KRA — {selected_month} 2026",
            data=out, file_name=clean_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        import traceback
        steps_ph.empty()
        st.error(f"❌ {e}")
        st.code(traceback.format_exc())

# ── IDLE STATE ────────────────────────────────────────────────────────────────
if not btn and not all_ready:
    st.markdown("""
    <div style="background:linear-gradient(145deg,#0F1729,#131E35);
         border:1px dashed #1E3A5F;border-radius:20px;
         padding:60px 40px;text-align:center;margin-top:8px;">
      <div style="font-size:48px;margin-bottom:16px;">📋</div>
      <h3 style="color:#4B5563;font-weight:600;margin:0 0 10px;">Ready to Begin</h3>
      <p style="color:#374151;font-size:14px;max-width:400px;margin:0 auto;line-height:1.7;">
        Upload your 3 Excel files above, select the target month,
        and click <strong style="color:#60A5FA;">Generate Updated KRA</strong>
      </p>
      <div style="display:flex;justify-content:center;gap:24px;margin-top:28px;flex-wrap:wrap;">
        <div style="background:#0D1526;border:1px solid #1E3A5F;border-radius:10px;
             padding:12px 20px;font-size:13px;color:#475569;">📁 KL Cluster Report</div>
        <div style="font-size:18px;color:#1E3A5F;padding-top:10px;">+</div>
        <div style="background:#0D1526;border:1px solid #1E3A5F;border-radius:10px;
             padding:12px 20px;font-size:13px;color:#475569;">📁 Parameter Dashboard</div>
        <div style="font-size:18px;color:#1E3A5F;padding-top:10px;">+</div>
        <div style="background:#0D1526;border:1px solid #1E3A5F;border-radius:10px;
             padding:12px 20px;font-size:13px;color:#475569;">📁 KRA Template</div>
        <div style="font-size:18px;color:#1E3A5F;padding-top:10px;">→</div>
        <div style="background:linear-gradient(135deg,rgba(37,99,235,0.15),rgba(124,58,237,0.15));
             border:1px solid rgba(37,99,235,0.3);border-radius:10px;
             padding:12px 20px;font-size:13px;color:#60A5FA;font-weight:600;">✨ Updated KRA</div>
      </div>
    </div>
    """, unsafe_allow_html=True)
