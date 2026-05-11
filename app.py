import streamlit as st
import pandas as pd
import math
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="OPAL Order Tool", page_icon="💎", layout="wide")

st.markdown("""
<style>
/* ── hide default streamlit chrome ── */
#MainMenu, footer, header {visibility: hidden;}

/* ── app background ── */
.stApp {
    background: radial-gradient(ellipse at top left, #0e0e1a 0%, #08080f 60%);
}

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background: #0c0c18 !important;
    border-right: 1px solid rgba(137,212,245,0.15);
}
[data-testid="stSidebar"] .stMarkdown h2 {
    background: linear-gradient(135deg, #89d4f5, #c9a6ff, #f5a6d3);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 1.3rem;
    letter-spacing: 0.12em;
    font-weight: 700;
}

/* ── metric cards ── */
[data-testid="metric-container"] {
    background: rgba(137,212,245,0.04);
    border: 1px solid rgba(137,212,245,0.18);
    border-radius: 10px;
    padding: 14px 18px;
    box-shadow: 0 0 18px rgba(137,212,245,0.06);
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    color: #89d4f5 !important;
    font-size: 0.72rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e8e8f0 !important;
    font-size: 1.4rem;
    font-weight: 700;
}

/* ── tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #12121e;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
    border: 1px solid rgba(137,212,245,0.12);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 7px;
    color: #888 !important;
    font-weight: 500;
    letter-spacing: 0.04em;
    padding: 8px 20px;
}
.stTabs [aria-selected="true"] {
    background: rgba(137,212,245,0.1) !important;
    color: #89d4f5 !important;
    border-bottom: 2px solid #89d4f5 !important;
}

/* ── download buttons ── */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, rgba(137,212,245,0.15), rgba(201,166,255,0.15));
    border: 1px solid rgba(137,212,245,0.35);
    color: #89d4f5 !important;
    border-radius: 8px;
    font-weight: 600;
    letter-spacing: 0.05em;
    transition: all 0.2s;
}
[data-testid="stDownloadButton"] button:hover {
    background: linear-gradient(135deg, rgba(137,212,245,0.28), rgba(201,166,255,0.28));
    border-color: #89d4f5;
    box-shadow: 0 0 16px rgba(137,212,245,0.25);
}

/* ── expanders ── */
[data-testid="stExpander"] {
    border: 1px solid rgba(137,212,245,0.15) !important;
    border-radius: 10px !important;
    background: rgba(137,212,245,0.02);
}

/* ── multiselect ── */
[data-baseweb="select"] {
    background: #12121e;
}
[data-baseweb="tag"] {
    background: rgba(137,212,245,0.15) !important;
    border: 1px solid rgba(137,212,245,0.3) !important;
    color: #89d4f5 !important;
}

/* ── section headings ── */
.stMarkdown h3 {
    background: linear-gradient(135deg, #89d4f5, #c9a6ff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: 0.06em;
}

/* ── dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(137,212,245,0.12);
    border-radius: 8px;
    overflow: hidden;
}

/* ── number inputs & selectbox ── */
[data-testid="stNumberInput"] input,
[data-baseweb="select"] div {
    background: #12121e !important;
    border-color: rgba(137,212,245,0.2) !important;
}

/* ── dividers ── */
hr {
    border-color: rgba(137,212,245,0.12) !important;
}
</style>
""", unsafe_allow_html=True)

# ── styles ────────────────────────────────────────────────────
def _f(color='000000', bold=False, sz=10): return Font(color=color, bold=bold, size=sz)
def _p(color): return PatternFill('solid', fgColor=color)
def _a(h='left', v='center', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _b():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

HDR_FILL = _p('1F4E79'); CAT_FILL = _p('2E75B6')
HDR_FONT = _f('FFFFFF', bold=True, sz=10); CAT_FONT = _f('FFFFFF', bold=True, sz=10)
BLD_FONT = _f('000000', bold=True, sz=10); BLK_FONT = _f('000000', sz=10)
RED_FILL = _p('FFC7CE'); YEL_FILL = _p('FFEB9C'); GRN_FILL = _p('C6EFCE')
BORDER   = _b()

def hcell(ws, r, c, val, fill=None, font=None, align='center', wrap=True, colspan=0):
    cell = ws.cell(row=r, column=c, value=val)
    cell.fill = fill or HDR_FILL
    cell.font = font or HDR_FONT
    cell.alignment = _a(align, wrap=wrap)
    cell.border = BORDER
    if colspan > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+colspan-1)
    return cell

def vcell(ws, r, c, val, font=None, fill=None, fmt=None, align='left', bold=False):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = font or (_f('000000', bold=True) if bold else BLK_FONT)
    if fill: cell.fill = fill
    if fmt:  cell.number_format = fmt
    cell.alignment = _a(align)
    cell.border = BORDER
    return cell

# ── sidebar ───────────────────────────────────────────────────
with st.sidebar:
    try:
        st.image("opal_logo.png", use_container_width=True)
    except Exception:
        st.markdown("## 💎 OPAL Order Tool")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**📁 Upload Files**")
    kova_file = st.file_uploader("Cova Reorder Report (.xlsx)", type="xlsx", key="kova")
    ocs_file  = st.file_uploader("OCS Catalogue (.xlsx)",       type="xlsx", key="ocs")

    st.markdown("---")
    st.header("⚙️ Settings")

    # ── tax settings ──────────────────────────────────────────
    PROVINCE_RATES = {
        "Ontario (HST 13%)":              0.13,
        "Nova Scotia (HST 15%)":          0.15,
        "New Brunswick (HST 15%)":        0.15,
        "Newfoundland (HST 15%)":         0.15,
        "PEI (HST 15%)":                  0.15,
        "British Columbia (GST+PST 12%)": 0.12,
        "Manitoba (GST+PST 12%)":         0.12,
        "Saskatchewan (GST+PST 11%)":     0.11,
        "Quebec (GST+QST 14.975%)":       0.14975,
        "Alberta (GST 5%)":               0.05,
        "Custom":                         None,
    }
    province = st.selectbox("Province / Tax Region", list(PROVINCE_RATES.keys()), index=0)
    if PROVINCE_RATES[province] is None:
        tax_rate = st.number_input("Custom Tax Rate (%)", value=13.0, min_value=0.0, max_value=30.0, step=0.1) / 100
    else:
        tax_rate = PROVINCE_RATES[province]
        st.caption(f"Tax rate: {tax_rate*100:.3g}%")

    budget = st.number_input("Weekly Budget (tax-in)", value=30000, step=500, format="%d")

    st.markdown("**Shipping**")
    shipping_cost = st.number_input("Shipping Cost ($)", value=0, min_value=0, step=5, format="%d")
    ship_in_budget = st.checkbox("Deduct shipping from order budget", value=False,
                                  help="When checked, shipping is subtracted before calculating how much product you can order.")

    effective_budget = budget - shipping_cost if ship_in_budget else budget
    budget_pretax    = round(effective_budget / (1 + tax_rate), 2)
    tax_amount       = effective_budget - budget_pretax
    st.caption(f"Pre-tax: ${budget_pretax:,.2f}  |  Tax ({tax_rate*100:.3g}%): ${tax_amount:,.2f}"
               + (f"  |  Shipping deducted: ${shipping_cost:,}" if ship_in_budget and shipping_cost > 0 else ""))

    st.markdown("**Target Days of Supply**")

    with st.expander("🌸 Flower (by size)", expanded=True):
        fc1, fc2 = st.columns(2)
        with fc1:
            t_flower_1g   = st.number_input("1g",   value=7,  min_value=1, max_value=90, key="f1g")
            t_flower_35g  = st.number_input("3.5g", value=14, min_value=1, max_value=90, key="f35g")
            t_flower_5g   = st.number_input("5g",   value=14, min_value=1, max_value=90, key="f5g")
            t_flower_7g   = st.number_input("7g",   value=14, min_value=1, max_value=90, key="f7g")
        with fc2:
            t_flower_14g  = st.number_input("14g",  value=21, min_value=1, max_value=90, key="f14g")
            t_flower_28g  = st.number_input("28g",  value=30, min_value=1, max_value=90, key="f28g")
            t_flower_30g  = st.number_input("30g",  value=30, min_value=1, max_value=90, key="f30g")
            t_flower_other= st.number_input("Other", value=14, min_value=1, max_value=90, key="fother")

    FLOWER_SIZE_TARGET = {
        '1g': t_flower_1g, '3.5g': t_flower_35g, '5g': t_flower_5g,
        '7g': t_flower_7g, '14g': t_flower_14g,  '28g': t_flower_28g,
        '30g': t_flower_30g,
    }

    col1, col2 = st.columns(2)
    with col1:
        t_preroll   = st.number_input("Pre-Roll",     value=14, min_value=1, max_value=90)
        t_edibles   = st.number_input("Edibles",      value=14, min_value=1, max_value=90)
        t_vapes     = st.number_input("Vapes",        value=21, min_value=1, max_value=90)
        t_beverages = st.number_input("Beverages",    value=14, min_value=1, max_value=90)
        t_capsules  = st.number_input("Capsules",     value=21, min_value=1, max_value=90)
    with col2:
        t_conc      = st.number_input("Concentrates", value=14, min_value=1, max_value=90)
        t_topicals  = st.number_input("Topicals",     value=30, min_value=1, max_value=90)
        t_oil       = st.number_input("Oil",          value=21, min_value=1, max_value=90)
        t_seeds     = st.number_input("Seeds",        value=60, min_value=1, max_value=90)

    TARGET = {
        'Pre-Roll': t_preroll, 'Edibles': t_edibles, 'Vapes': t_vapes,
        'Beverages': t_beverages, 'Capsules': t_capsules, 'Concentrates': t_conc,
        'Topicals': t_topicals, 'Oil': t_oil, 'Seeds': t_seeds,
    }

# ── main ──────────────────────────────────────────────────────
try:
    st.image("opal_logo.png", width=280)
except Exception:
    st.markdown("# 💎 OPAL Order Tool")
st.markdown(
    f"<p style='color:#666;font-size:0.85rem;letter-spacing:0.08em;margin-top:-8px'>"
    f"ORDER TOOL &nbsp;·&nbsp; {date.today().strftime('%B %d, %Y').upper()}</p>",
    unsafe_allow_html=True
)

if not kova_file or not ocs_file:
    st.info("Upload your **Cova Reorder Report** and **OCS Catalogue** in the sidebar to get started with OPAL.")
    st.markdown("""
    **How to export from Cova:**
    Reports → Reorder → Export as .xlsx

    **How to export from OCS Wholesale:**
    Catalogue → Download Catalogue → Excel
    """)
    st.stop()

tab1, tab2, tab3 = st.tabs(["📦 Replenishment Order", "📋 Suggested Order", "🗂️ Menu Builder"])

# ── shared data loading ────────────────────────────────────────
@st.cache_data(show_spinner="Loading catalogue...")
def load_raw(kova_bytes, ocs_bytes):
    import re
    kova = pd.read_excel(io.BytesIO(kova_bytes), sheet_name='Reorder')
    ocs  = pd.read_excel(io.BytesIO(ocs_bytes),  sheet_name='MasterCatalogue')
    kova_ocs = kova[kova['Supplier'] == 'OCS'].copy()
    kova_ocs['Supplier Sku'] = kova_ocs['Supplier Sku'].str.strip()
    ocs['OCS Variant Number'] = ocs['OCS Variant Number'].str.strip()
    ocs_cols = ['OCS Variant Number','OCS Item Number','Unit Price','Pack Size','Stock Status','Plant Type']
    for _opt in ['Sub-Category']:
        if _opt in ocs.columns:
            ocs_cols.append(_opt)
    merged = kova_ocs.merge(ocs[ocs_cols], left_on='Supplier Sku', right_on='OCS Variant Number', how='left')
    if 'Sub-Category' not in merged.columns:
        merged['Sub-Category'] = ''
    else:
        merged['Sub-Category'] = merged['Sub-Category'].fillna('')

    def extract_size(sku):
        m = re.search(r'_(\d+\.?\d*[gG])_', str(sku))
        return m.group(1).lower() if m else None

    def extract_preroll_size(sku):
        m = re.search(r'_(\d+x\d+\.?\d*[gG])_', str(sku))
        return m.group(1).lower() if m else None

    def extract_product_size(sku):
        s = str(sku)
        if '___' in s:
            before = s.split('___')[0]
            idx = before.find('_')
            if idx >= 0:
                val = before[idx+1:]
                return val if val else None
        return None

    def map_strain(row):
        pt = str(row.get('Plant Type', '') or '').lower()
        if 'indica' in pt:  return 'Indica'
        if 'sativa' in pt:  return 'Sativa'
        if 'hybrid' in pt:  return 'Hybrid'
        if 'blend'  in pt:  return 'Blend'
        name = str(row.get('Product', '') or '').lower()
        for token in ['- indica', 'indica -', '(indica)']:
            if token in name: return 'Indica'
        for token in ['- sativa', 'sativa -', '(sativa)']:
            if token in name: return 'Sativa'
        for token in ['- hybrid', 'hybrid -', '(hybrid)']:
            if token in name: return 'Hybrid'
        for token in ['- blend', 'blend -', '(blend)']:
            if token in name: return 'Blend'
        return 'Unknown'

    merged['Flower Size']   = merged.apply(
        lambda r: extract_size(r['Supplier Sku']) if r['Classification'] == 'Flower' else None, axis=1)
    merged['Pre-Roll Size'] = merged.apply(
        lambda r: extract_preroll_size(r['Supplier Sku']) if r['Classification'] == 'Pre-Roll' else None, axis=1)
    merged['Product Size']  = merged['Supplier Sku'].apply(extract_product_size)
    merged['Strain'] = merged.apply(map_strain, axis=1)
    merged['Pack Size']  = pd.to_numeric(merged['Pack Size'],  errors='coerce').fillna(1).astype(int)
    merged['Unit Price'] = pd.to_numeric(merged['Unit Price'], errors='coerce')
    return merged, ocs

kova_bytes_raw = kova_file.read()
ocs_bytes_raw  = ocs_file.read()
merged_raw, ocs_df = load_raw(kova_bytes_raw, ocs_bytes_raw)

# ── process data ──────────────────────────────────────────────
@st.cache_data(show_spinner="Processing your data...")
def process(kova_bytes, ocs_bytes, target, flower_size_target, budget_pretax):
    merged = load_raw(kova_bytes, ocs_bytes)[0]
    active = merged[merged['Sales (30 Days)'] > 0].copy()
    active['Weekly Vel'] = (active['Sales (30 Days)'] / 4).round(2)
    active['Daily Vel']  = active['Sales (30 Days)'] / 30
    active['Days Left']  = active['Days of Stock Left (30 Days)'].round(1)
    active['Available']  = active['In Stock Qty'] + active['On Order']

    def assign_tier(w):
        if w >= 5:   return 'A'
        elif w >= 2: return 'B'
        elif w >= 1: return 'C'
        else:        return 'D'
    active['Tier'] = active['Weekly Vel'].apply(assign_tier)

    def calc_order(row):
        tier = row['Tier']; pack = int(row['Pack Size'])
        avail = row['Available']; dv = row['Daily Vel']
        dl = row['Days Left'] if pd.notna(row['Days Left']) else 0
        ins = row['In Stock Qty']
        if row['Classification'] == 'Flower':
            tgt = flower_size_target.get(row['Flower Size'], flower_size_target.get('other', 14))
        else:
            tgt = target.get(row['Classification'], 14)
        if tier == 'A':
            needed = max(0, math.ceil(dv * tgt) - avail)
            cases  = math.ceil(needed / pack) if needed > 0 else 0
        elif tier == 'B':
            if dl >= 10: return 0, 0, tier
            needed = max(0, math.ceil(dv * 10) - avail)
            cases  = math.ceil(needed / pack) if needed > 0 else 0
        elif tier == 'C':
            if dl >= 5 and ins > 0: return 0, 0, tier
            cases = 1
        else:
            if ins > 0 or row['On Order'] > 0: return 0, 0, tier
            cases = 1
        return cases, cases * pack, tier

    active[['Cases','Suggest','Tier']] = active.apply(lambda r: pd.Series(calc_order(r)), axis=1)
    active['Suggest']  = active['Suggest'].astype(int)
    active['Cases']    = active['Cases'].astype(int)
    active['Est Cost'] = (active['Suggest'] * active['Unit Price'].fillna(0)).round(2)

    needs = active[active['Suggest'] > 0].copy()
    needs['TP'] = needs['Tier'].map({'A':0,'B':1,'C':2,'D':3})
    needs['DL'] = needs['Days Left'].fillna(0)
    needs = needs.sort_values(['TP','DL']).reset_index(drop=True)

    cum = 0.0; in_b = []; deferred = []
    for _, row in needs.iterrows():
        if cum + row['Est Cost'] <= budget_pretax:
            cum += row['Est Cost']; in_b.append(row)
        else:
            deferred.append(row)

    order_df    = pd.DataFrame(in_b).sort_values(['Classification','DL']).reset_index(drop=True) if in_b else pd.DataFrame()
    deferred_df = pd.DataFrame(deferred).reset_index(drop=True) if deferred else pd.DataFrame()
    all_active  = active.sort_values(['Classification','DL' if 'DL' in active.columns else 'Days Left']).reset_index(drop=True)

    return order_df, deferred_df, all_active

order_df, deferred_df, all_active = process(kova_bytes_raw, ocs_bytes_raw, TARGET, FLOWER_SIZE_TARGET, budget_pretax)

if order_df.empty:
    st.warning("No items need ordering based on current stock levels and settings.")
    st.stop()

# ── shared style helpers ──────────────────────────────────────
def color_tier(val):
    colors = {'A':'background-color:#C6EFCE','B':'background-color:#E2EFDA',
              'C':'background-color:#FFEB9C','D':'background-color:#FFC7CE'}
    return colors.get(val, '')

def color_days(val):
    try:
        v = float(val)
        if v <= 3:  return 'background-color:#FFC7CE;color:#9C0006;font-weight:bold'
        if v <= 7:  return 'background-color:#FFEB9C;color:#9C6500;font-weight:bold'
        if v >= 14: return 'background-color:#C6EFCE;color:#276221'
    except: pass
    return ''

# ── workbook builders (defined before tabs) ───────────────────
def build_workbook(order_df, deferred_df, all_active, budget_pretax, target, today, tax_rate=0.13):
    wb = Workbook()

    tier_colours = {'A':'375623','B':'375623','C':'9C6500','D':'9C0006'}
    tier_fills   = {'A':'C6EFCE','B':'E2EFDA','C':'FFEB9C','D':'FFC7CE'}

    def write_order_sheet(ws, df, title):
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'A3'
        cols = [
            ('Category',14),('Product Name',40),('OCS Item #',12),('Tier',7),
            ('Days Left',10),('In Stock',10),('On Order',10),('Wkly Sales',10),
            ('30d Sales',10),('Daily Vel.',10),('Pack Size',10),('Cases',8),
            ('Suggest Order',13),('Unit Price',11),('Est. Cost',11),('OCS Variant #',18),
        ]
        for i, (_, w) in enumerate(cols, 1):
            ws.column_dimensions[ws.cell(row=1,column=i).column_letter].width = w
        hcell(ws, 1, 1, title, colspan=len(cols))
        for i, (h, _) in enumerate(cols, 1): hcell(ws, 2, i, h, wrap=True)
        ws.row_dimensions[1].height = 22; ws.row_dimensions[2].height = 30

        ROW = 3; prev_cat = None
        for _, row in df.iterrows():
            cat = row['Classification']
            if cat != prev_cat:
                for ci in range(1, len(cols)+1):
                    c = ws.cell(row=ROW, column=ci); c.fill = CAT_FILL; c.border = BORDER
                ws.cell(row=ROW,column=1).value = f'▶  {cat.upper()}'
                ws.cell(row=ROW,column=1).font = CAT_FONT
                ws.cell(row=ROW,column=1).alignment = _a('left')
                ROW += 1; prev_cat = cat

            alt = _p('F5FBFF') if ROW%2==0 else _p('FFFFFF')
            t = row['Tier']
            dl = row['Days Left'] if pd.notna(row.get('Days Left')) else 0

            def v(c, val, fmt=None, align='center', fill=alt):
                vcell(ws, ROW, c, val, fill=fill, fmt=fmt, align=align)

            v(1,  row['Classification'], align='left', fill=alt)
            v(2,  row['Product'],        align='left', fill=alt)
            v(3,  row['OCS Item Number'] if pd.notna(row.get('OCS Item Number')) else row['Supplier Sku'].split('_')[0])
            ct = ws.cell(row=ROW,column=4,value=t)
            ct.fill=_p(tier_fills[t]); ct.font=_f(tier_colours[t],bold=True); ct.alignment=_a('center'); ct.border=BORDER
            dl_c = ws.cell(row=ROW,column=5,value=round(dl,1))
            dl_c.number_format='0.0'; dl_c.alignment=_a('center'); dl_c.border=BORDER
            if dl<=3:   dl_c.fill=RED_FILL; dl_c.font=_f('9C0006',bold=True)
            elif dl<=7: dl_c.fill=YEL_FILL; dl_c.font=_f('9C6500',bold=True)
            else:       dl_c.fill=GRN_FILL; dl_c.font=_f('276221'); dl_c.fill=alt if dl<14 else GRN_FILL
            v(6,  int(row['In Stock Qty']))
            v(7,  int(row['On Order']))
            v(8,  int(row['Sales (7 Days)']))
            v(9,  int(row['Sales (30 Days)']))
            v(10, round(row['Daily Vel'],3), fmt='0.000')
            v(11, int(row['Pack Size']))
            v(12, int(row['Cases']))
            cs = ws.cell(row=ROW,column=13,value=int(row['Suggest']))
            cs.font=_f('000000',bold=True); cs.fill=_p('FFF2CC'); cs.alignment=_a('center'); cs.border=BORDER
            price = row['Unit Price'] if pd.notna(row.get('Unit Price')) else None
            v(14, price, fmt='"$"#,##0.00')
            v(15, row['Est Cost'] if pd.notna(row.get('Est Cost')) and row['Est Cost']>0 else None, fmt='"$"#,##0.00')
            v(16, row['Supplier Sku'], align='left')
            ROW += 1

        last = ROW-1
        for label, offset, formula in [
            ('SUBTOTAL (pre-tax)', 0, f'=SUM(O3:O{last})'),
            (f'Tax ({tax_rate*100:.3g}%)', 1, f'=O{ROW}*{tax_rate}'),
            ('TOTAL (tax-in)',    2, f'=O{ROW}+O{ROW+1}'),
        ]:
            r = ROW+offset
            for ci in range(1,len(cols)+1):
                ws.cell(row=r,column=ci).fill=_p('1F4E79'); ws.cell(row=r,column=ci).border=BORDER
            lbl=ws.cell(row=r,column=1,value=label); lbl.font=HDR_FONT; lbl.fill=_p('1F4E79'); lbl.alignment=_a('right'); lbl.border=BORDER
            cf=ws.cell(row=r,column=15,value=formula); cf.font=HDR_FONT; cf.fill=_p('1F4E79')
            cf.number_format='"$"#,##0.00'; cf.alignment=_a('center'); cf.border=BORDER

    # Sheet 1: Order Builder
    ws_ob = wb.active; ws_ob.title = 'Order Builder'
    write_order_sheet(ws_ob, order_df, f'OCS Order Builder — {today.strftime("%B %d, %Y")} — Budget: ${budget_pretax:,.0f} pre-tax (${round(budget_pretax*(1+tax_rate)):,} tax-in)')
    ws_ob.sheet_properties.tabColor = 'FFC000'

    # Sheet 2: Portal Order
    ws_portal = wb.create_sheet('OCS Portal Order')
    ws_portal.sheet_view.showGridLines = False
    ws_portal.freeze_panes = 'A3'
    pcols = [('OCS Item #',16),('OCS Variant #',22),('Product Name',48),('Category',14),
             ('Cases to Order',14),('Units per Case',13),('Total Units',12),('Unit Price',12),('Line Total',13)]
    for i,(_,w) in enumerate(pcols,1):
        ws_portal.column_dimensions[ws_portal.cell(row=1,column=i).column_letter].width=w
    hcell(ws_portal,1,1,f'OCS Wholesale Order — {today.strftime("%B %d, %Y")}  |  Enter CASES in the portal',colspan=len(pcols))
    for i,(h,_) in enumerate(pcols,1): hcell(ws_portal,2,i,h,wrap=True)
    PROW=3; prev_pcat=None
    portal_sorted = order_df.sort_values(['Classification','Product']).reset_index(drop=True)
    for _,row in portal_sorted.iterrows():
        cat=row['Classification']
        if cat!=prev_pcat:
            for ci in range(1,len(pcols)+1):
                c=ws_portal.cell(row=PROW,column=ci); c.fill=CAT_FILL; c.border=BORDER
            ws_portal.cell(row=PROW,column=1).value=f'▶  {cat.upper()}'
            ws_portal.cell(row=PROW,column=1).font=CAT_FONT; ws_portal.cell(row=PROW,column=1).alignment=_a('left')
            PROW+=1; prev_pcat=cat
        alt=_p('F0F7FF') if PROW%2==0 else _p('FFFFFF')
        item=str(row['OCS Item Number']).split('.')[0] if pd.notna(row.get('OCS Item Number')) else row['Supplier Sku'].split('_')[0]
        ci_=ws_portal.cell(row=PROW,column=1,value=item); ci_.font=_f('000000',bold=True,sz=11); ci_.fill=alt; ci_.border=BORDER; ci_.alignment=_a('center')
        vcell(ws_portal,PROW,2,row['Supplier Sku'],fill=alt,align='left')
        vcell(ws_portal,PROW,3,row['Product'],fill=alt,align='left')
        vcell(ws_portal,PROW,4,cat,fill=alt,align='left')
        cc=ws_portal.cell(row=PROW,column=5,value=int(row['Cases'])); cc.font=_f('000000',bold=True,sz=12); cc.fill=_p('FFF2CC'); cc.alignment=_a('center'); cc.border=BORDER
        vcell(ws_portal,PROW,6,int(row['Pack Size']),fill=alt,align='center')
        vcell(ws_portal,PROW,7,int(row['Suggest']),fill=alt,align='center')
        vcell(ws_portal,PROW,8,row['Unit Price'] if pd.notna(row.get('Unit Price')) else None,fill=alt,fmt='"$"#,##0.00',align='center')
        vcell(ws_portal,PROW,9,row['Est Cost'] if pd.notna(row.get('Est Cost')) and row['Est Cost']>0 else None,fill=alt,fmt='"$"#,##0.00',align='center')
        PROW+=1
    last_p=PROW-1
    for label,offset,formula in [('SUBTOTAL (pre-tax)',0,f'=SUM(I3:I{last_p})'),(f'Tax ({tax_rate*100:.3g}%)',1,f'=I{PROW}*{tax_rate}'),('TOTAL (tax-in)',2,f'=I{PROW}+I{PROW+1}')]:
        r=PROW+offset
        for ci in range(1,len(pcols)+1): ws_portal.cell(row=r,column=ci).fill=_p('1F4E79'); ws_portal.cell(row=r,column=ci).border=BORDER
        lbl=ws_portal.cell(row=r,column=1,value=label); lbl.font=HDR_FONT; lbl.fill=_p('1F4E79'); lbl.alignment=_a('right'); lbl.border=BORDER
        ws_portal.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        cf=ws_portal.cell(row=r,column=9,value=formula); cf.font=HDR_FONT; cf.fill=_p('1F4E79'); cf.number_format='"$"#,##0.00'; cf.alignment=_a('center'); cf.border=BORDER
    ws_portal.row_dimensions[1].height=22; ws_portal.row_dimensions[2].height=30
    ws_portal.sheet_properties.tabColor='375623'

    # Sheet 3: Deferred
    if not deferred_df.empty:
        ws_def = wb.create_sheet('Deferred (Next Order)')
        write_order_sheet(ws_def, deferred_df, f'Deferred Items — Carry to Next Order (${deferred_df["Est Cost"].sum():,.2f} pre-tax)')
        ws_def.sheet_properties.tabColor='843C0C'

    # Sheet 4: Full Inventory
    ws_inv = wb.create_sheet('Full Inventory')
    ws_inv.sheet_view.showGridLines = False; ws_inv.freeze_panes='A3'
    inv_cols=[('Category',14),('Product Name',40),('OCS Item #',12),('Tier',7),('Days Left',10),
              ('In Stock',10),('On Order',10),('7d Sales',9),('30d Sales',9),('60d Sales',9),
              ('Daily Vel.',10),('Wkly Vel.',10),('Suggest Order',12),('Unit Price',11),('Margin %',10),('OCS Variant #',18)]
    for i,(_,w) in enumerate(inv_cols,1):
        ws_inv.column_dimensions[ws_inv.cell(row=1,column=i).column_letter].width=w
    hcell(ws_inv,1,1,f'Full Active Inventory — {today.strftime("%B %d, %Y")}',colspan=len(inv_cols))
    for i,(h,_) in enumerate(inv_cols,1): hcell(ws_inv,2,i,h,wrap=True)
    ws_inv.row_dimensions[1].height=22; ws_inv.row_dimensions[2].height=30
    IROW=3; prev_icat=None
    for _,row in all_active.iterrows():
        cat=row['Classification']
        if cat!=prev_icat:
            for ci in range(1,len(inv_cols)+1):
                c=ws_inv.cell(row=IROW,column=ci); c.fill=CAT_FILL; c.border=BORDER
            ws_inv.cell(row=IROW,column=1).value=f'▶  {cat.upper()}'
            ws_inv.cell(row=IROW,column=1).font=CAT_FONT; ws_inv.cell(row=IROW,column=1).alignment=_a('left')
            IROW+=1; prev_icat=cat
        t=row['Tier']; alt=_p('F5FBFF') if IROW%2==0 else _p('FFFFFF')
        dl=row['Days Left'] if pd.notna(row.get('Days Left')) else 0
        suggest=int(row['Suggest'])
        def wi(c,val,fmt=None,align='center',fill=alt): vcell(ws_inv,IROW,c,val,fill=fill,fmt=fmt,align=align)
        wi(1,row['Classification'],align='left',fill=alt); wi(2,row['Product'],align='left',fill=alt)
        wi(3,row['OCS Item Number'] if pd.notna(row.get('OCS Item Number')) else row['Supplier Sku'].split('_')[0])
        ct2=ws_inv.cell(row=IROW,column=4,value=t); ct2.fill=_p(tier_fills[t]); ct2.font=_f(tier_colours[t],bold=True); ct2.alignment=_a('center'); ct2.border=BORDER
        dl2=ws_inv.cell(row=IROW,column=5,value=round(dl,1)); dl2.number_format='0.0'; dl2.alignment=_a('center'); dl2.border=BORDER
        if dl<=3: dl2.fill=RED_FILL; dl2.font=_f('9C0006',bold=True)
        elif dl<=7: dl2.fill=YEL_FILL; dl2.font=_f('9C6500',bold=True)
        else: dl2.fill=alt
        wi(6,int(row['In Stock Qty'])); wi(7,int(row['On Order']))
        wi(8,int(row['Sales (7 Days)'])); wi(9,int(row['Sales (30 Days)'])); wi(10,int(row['Sales (60 Days)']))
        wi(11,round(row['Daily Vel'],3),fmt='0.000'); wi(12,round(row['Weekly Vel'],2),fmt='0.00')
        cs2=ws_inv.cell(row=IROW,column=13,value=suggest); cs2.fill=_p('FFF2CC') if suggest>0 else alt; cs2.font=_f('000000',bold=(suggest>0)); cs2.alignment=_a('center'); cs2.border=BORDER
        wi(14,row['Unit Price'] if pd.notna(row.get('Unit Price')) else None,fmt='"$"#,##0.00')
        wi(15,row['Margin'] if pd.notna(row.get('Margin')) else None,fmt='0.0%')
        wi(16,row['Supplier Sku'],align='left')
        IROW+=1
    ws_inv.sheet_properties.tabColor='70AD47'

    wb.active = wb['Order Builder']
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def build_upload(order_df):
    upload = order_df[['Supplier Sku','Cases']].copy()
    upload.columns = ['SKU','Quantity']
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        upload.to_excel(writer, sheet_name='MasterCatalogue', index=False)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════
# TAB 1: REPLENISHMENT ORDER
# ══════════════════════════════════════════════════════════════
with tab1:
    pretax_total      = order_df['Est Cost'].sum()
    tax_amount_actual = pretax_total * tax_rate
    subtotal_taxin    = pretax_total + tax_amount_actual
    grand_total       = subtotal_taxin + (shipping_cost if not ship_in_budget else 0)
    tax_label         = f"Tax ({tax_rate*100:.3g}%)"

    # ── summary metrics ───────────────────────────────────────
    st.markdown("### This Week's Order")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("SKUs to Order",   len(order_df))
    m2.metric("Pre-Tax Total",   f"${pretax_total:,.2f}")
    m3.metric(tax_label,         f"${tax_amount_actual:,.2f}")
    m4.metric("Subtotal (Tax-In)", f"${subtotal_taxin:,.2f}")
    m5.metric("Shipping",        f"${shipping_cost:,}" if shipping_cost > 0 else "—")
    m6.metric("Grand Total",     f"${grand_total:,.2f}")

    st.markdown("---")

    # ── category summary ──────────────────────────────────────
    col_left, col_right = st.columns([1, 2])

    with col_left:
        st.markdown("**By Category**")
        cat_sum = order_df.groupby('Classification').agg(
            SKUs=('SKU','count'),
            Cost=('Est Cost','sum')
        ).sort_values('Cost', ascending=False).reset_index()
        cat_sum['Cost'] = cat_sum['Cost'].map('${:,.2f}'.format)
        st.dataframe(cat_sum, hide_index=True, use_container_width=True)

    with col_right:
        st.markdown("**By Velocity Tier**")
        tier_labels = {'A':'A — Fast (5+/wk)','B':'B — Mid (2–4/wk)','C':'C — Slow (~1/wk)','D':'D — Very Slow (<1/wk)'}
        tier_sum = order_df.groupby('Tier').agg(
            SKUs=('SKU','count'),
            Cost=('Est Cost','sum')
        ).reset_index()
        tier_sum['Tier'] = tier_sum['Tier'].map(tier_labels)
        tier_sum['Cost'] = tier_sum['Cost'].map('${:,.2f}'.format)
        st.dataframe(tier_sum, hide_index=True, use_container_width=True)

    st.markdown("---")

    # ── tier filter ───────────────────────────────────────────
    st.markdown("### Order Detail")
    tier_options = ['A — Fast (5+/wk)', 'B — Mid (2–4/wk)', 'C — Slow (~1/wk)', 'D — Very Slow (<1/wk)']
    tier_map     = {'A — Fast (5+/wk)':'A','B — Mid (2–4/wk)':'B','C — Slow (~1/wk)':'C','D — Very Slow (<1/wk)':'D'}
    selected_labels = st.multiselect(
        "Filter by Tier (select one or more to combine)",
        options=tier_options, default=tier_options, key="t1_tier_filter"
    )
    selected_tiers = [tier_map[l] for l in selected_labels] if selected_labels else list(tier_map.values())
    filtered_df = order_df[order_df['Tier'].isin(selected_tiers)]

    if filtered_df.empty:
        st.info("No items match the selected tiers.")
    else:
        f_pretax    = filtered_df['Est Cost'].sum()
        f_taxin     = f_pretax * (1 + tax_rate)
        f_grand     = f_taxin + (shipping_cost if not ship_in_budget else 0)
        fa1, fa2, fa3, fa4, fa5 = st.columns(5)
        fa1.metric("SKUs (filtered)", len(filtered_df))
        fa2.metric("Pre-Tax",   f"${f_pretax:,.2f}")
        fa3.metric(tax_label,   f"${f_pretax * tax_rate:,.2f}")
        fa4.metric("Shipping",  f"${shipping_cost:,}" if shipping_cost > 0 else "—")
        fa5.metric("Grand Total", f"${f_grand:,.2f}")

        display_cols = {
            'Classification':'Category','Product':'Product','Tier':'Tier',
            'Days Left':'Days Left','In Stock Qty':'In Stock','On Order':'On Order',
            'Sales (7 Days)':'7d Sales','Sales (30 Days)':'30d Sales',
            'Weekly Vel':'Wkly Vel','Cases':'Cases','Suggest':'Units',
            'Unit Price':'Unit Price','Est Cost':'Est Cost','Supplier Sku':'OCS Variant #'
        }
        show = filtered_df[[c for c in display_cols.keys() if c in filtered_df.columns]].rename(columns=display_cols)
        show['Unit Price'] = show['Unit Price'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
        show['Est Cost']   = show['Est Cost'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
        show['Days Left']  = show['Days Left'].round(1)
        st.dataframe(show.style.map(color_tier, subset=['Tier']).map(color_days, subset=['Days Left']),
                     hide_index=True, use_container_width=True, height=500)

    if not deferred_df.empty:
        with st.expander(f"📋 Deferred to Next Order ({len(deferred_df)} SKUs — ${deferred_df['Est Cost'].sum():,.2f} pre-tax)"):
            def_show = deferred_df[['Classification','Product','Tier','Days Left','In Stock Qty','Weekly Vel','Cases','Est Cost']].copy()
            def_show['Est Cost'] = def_show['Est Cost'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
            st.dataframe(def_show.rename(columns={'Classification':'Category','In Stock Qty':'In Stock','Weekly Vel':'Wkly Vel'}),
                         hide_index=True, use_container_width=True)

    st.markdown("---")

    # ── download files ────────────────────────────────────────
    st.markdown("### Download Files")
    if not filtered_df.empty:
        dl1, dl2 = st.columns(2)
        today = date.today()
        with dl1:
            workbook_buf = build_workbook(filtered_df, deferred_df, all_active, budget_pretax, TARGET, today, tax_rate)
            st.download_button(
                label="📥 Download Order Workbook (.xlsx)",
                data=workbook_buf,
                file_name=f'OCS_Order_Tool_{today.strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )
        with dl2:
            upload_buf = build_upload(filtered_df)
            st.download_button(
                label="📤 Download OCS Upload File (.xlsx)",
                data=upload_buf,
                file_name=f'OCS_Upload_{today.strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )

# ══════════════════════════════════════════════════════════════
# TAB 2: SUGGESTED ORDER (NO BUDGET CAP)
# ══════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 📋 Suggested Order")
    st.caption("All items that need restocking based on velocity — no hard budget cap. Review the total before submitting.")

    uncapped_df, _, _ = process(kova_bytes_raw, ocs_bytes_raw, TARGET, FLOWER_SIZE_TARGET, 999_999_999)

    if uncapped_df.empty:
        st.success("✅ Nothing needs restocking right now.")
    else:
        # ── tier filter ───────────────────────────────────────
        uc_tier_options = ['A — Fast (5+/wk)', 'B — Mid (2–4/wk)', 'C — Slow (~1/wk)', 'D — Very Slow (<1/wk)']
        uc_tier_map     = {'A — Fast (5+/wk)':'A','B — Mid (2–4/wk)':'B','C — Slow (~1/wk)':'C','D — Very Slow (<1/wk)':'D'}
        uc_selected_labels = st.multiselect(
            "Filter by Tier (select one or more to combine)",
            options=uc_tier_options, default=uc_tier_options, key="t2_tier_filter"
        )
        uc_selected_tiers = [uc_tier_map[l] for l in uc_selected_labels] if uc_selected_labels else list(uc_tier_map.values())
        uc_filtered = uncapped_df[uncapped_df['Tier'].isin(uc_selected_tiers)]

        uc_pretax = uc_filtered['Est Cost'].sum()
        uc_tax    = uc_pretax * tax_rate
        uc_taxin  = uc_pretax + uc_tax
        uc_grand  = uc_taxin + (shipping_cost if not ship_in_budget else 0)

        uc1, uc2, uc3, uc4, uc5 = st.columns(5)
        uc1.metric("SKUs to Order",   len(uc_filtered))
        uc2.metric("Pre-Tax Total",   f"${uc_pretax:,.2f}")
        uc3.metric(f"Tax ({tax_rate*100:.3g}%)", f"${uc_tax:,.2f}")
        uc4.metric("Shipping",        f"${shipping_cost:,}" if shipping_cost > 0 else "—")
        uc5.metric("Grand Total",     f"${uc_grand:,.2f}")

        st.markdown("---")

        if uc_filtered.empty:
            st.info("No items match the selected tiers.")
        else:
            uc_display_cols = {
                'Classification':'Category','Product':'Product','Tier':'Tier',
                'Days Left':'Days Left','In Stock Qty':'In Stock','On Order':'On Order',
                'Sales (7 Days)':'7d Sales','Sales (30 Days)':'30d Sales',
                'Weekly Vel':'Wkly Vel','Cases':'Cases','Suggest':'Units',
                'Unit Price':'Unit Price','Est Cost':'Est Cost','Supplier Sku':'OCS Variant #'
            }
            uc_show = uc_filtered[[c for c in uc_display_cols if c in uc_filtered.columns]].rename(columns=uc_display_cols).copy()
            uc_show['Unit Price'] = uc_show['Unit Price'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
            uc_show['Est Cost']   = uc_show['Est Cost'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
            uc_show['Days Left']  = uc_show['Days Left'].round(1)
            st.dataframe(
                uc_show.style.map(color_tier, subset=['Tier']).map(color_days, subset=['Days Left']),
                hide_index=True, use_container_width=True, height=500
            )

            st.markdown("---")
            uc_today = date.today()
            uc_upload = uc_filtered[['Supplier Sku','Cases']].copy()
            uc_upload.columns = ['SKU','Quantity']
            uc_buf = io.BytesIO()
            with pd.ExcelWriter(uc_buf, engine='openpyxl') as writer:
                uc_upload.to_excel(writer, sheet_name='MasterCatalogue', index=False)
            uc_buf.seek(0)
            st.download_button(
                label="📤 Download OCS Upload File (.xlsx)",
                data=uc_buf,
                file_name=f'OCS_SuggestedOrder_{uc_today.strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )

# ══════════════════════════════════════════════════════════════
# TAB 3: MENU BUILDER
# ══════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### 🗂️ Menu Builder")
    st.caption("Set how many SKUs you want on your shelf per category, size, and strain. The tool shows your gaps and suggests what to order.")

    today2 = date.today()
    STRAINS   = ['Indica','Sativa','Hybrid','Blend']
    ALL_CATS  = ['Flower','Pre-Roll','Vapes','Edibles','Concentrates','Beverages','Capsules','Oil','Topicals','Seeds']
    CAT_EMOJIS = {
        'Flower':'🌸','Pre-Roll':'🚬','Vapes':'💨','Edibles':'🍬',
        'Concentrates':'🔬','Beverages':'🥤','Capsules':'💊',
        'Oil':'🫙','Topicals':'🧴','Seeds':'🌱',
    }

    def sort_size_key(s):
        import re as _re
        m = _re.match(r'^(\d+)x(\d+\.?\d*)g?$', str(s))
        if m: return (0, int(m.group(1)), float(m.group(2)))
        m = _re.match(r'^(\d+\.?\d*)g$', str(s))
        if m: return (1, 0, float(m.group(1)))
        m = _re.match(r'^(\d+\.?\d*)ml$', str(s))
        if m: return (2, 0, float(m.group(1)))
        m = _re.match(r'^(\d+\.?\d*)mg$', str(s))
        if m: return (3, 0, float(m.group(1)))
        return (99, 0, 0)

    # Detect sizes per category
    cat_sizes = {}
    for cat in ALL_CATS:
        raw = merged_raw[(merged_raw['Classification']==cat) & merged_raw['Product Size'].notna()]['Product Size'].unique().tolist()
        try:    cat_sizes[cat] = sorted(raw, key=sort_size_key)
        except: cat_sizes[cat] = sorted(raw)

    # Sub-categories that don't belong under their OCS-assigned classification
    SUBCAT_EXCLUDE = {
        'Flower': {'Hash and Kief', 'Hash And Kief'},
    }

    # Sub-categories per category (auto-detected from OCS Sub-Category column)
    has_subcat  = 'Sub-Category' in merged_raw.columns
    subcat_map  = {}          # cat -> [subcats]
    subcat_sizes = {}         # (cat, subcat) -> [sizes]
    if has_subcat:
        for _cat in ALL_CATS:
            _excluded = SUBCAT_EXCLUDE.get(_cat, set())
            _sc_vals = [sc for sc in
                merged_raw[
                    (merged_raw['Classification']==_cat) & (merged_raw['Sub-Category']!='')
                ]['Sub-Category'].dropna().unique().tolist()
                if sc not in _excluded]
            if _sc_vals:
                subcat_map[_cat] = sorted(_sc_vals)
                for _sc in _sc_vals:
                    _raw = merged_raw[
                        (merged_raw['Classification']==_cat) &
                        (merged_raw['Sub-Category']==_sc) &
                        merged_raw['Product Size'].notna()
                    ]['Product Size'].unique().tolist()
                    try:    subcat_sizes[(_cat, _sc)] = sorted(_raw, key=sort_size_key)
                    except: subcat_sizes[(_cat, _sc)] = sorted(_raw)

    # ── target settings ───────────────────────────────────────
    st.markdown("#### Target SKU Counts")
    mb_targets = {}        # (cat, size, strain) for cats without sub-categories
    mb_subcat_targets = {} # (cat, subcat, size, strain) for cats with sub-categories

    for cat in ALL_CATS:
        sizes  = cat_sizes[cat]
        emoji  = CAT_EMOJIS[cat]
        subcats = subcat_map.get(cat, [])

        if subcats:
            with st.expander(f"{emoji} {cat} — by sub-category, size & strain", expanded=False):
                st.caption(f"{len(subcats)} sub-categories detected: {', '.join(subcats)}")
                for sc in subcats:
                    sc_sizes = subcat_sizes.get((cat, sc), [])
                    st.markdown(f"**🔹 {sc}**")
                    if not sc_sizes:
                        st.caption(f"No sizes detected for {sc}.")
                        continue
                    hdr = st.columns([2]+[1]*4)
                    hdr[0].markdown("**Size**")
                    for i, s in enumerate(STRAINS): hdr[i+1].markdown(f"**{s}**")
                    for size in sc_sizes:
                        rc = st.columns([2]+[1]*4)
                        rc[0].markdown(f"`{size}`")
                        for i, strain in enumerate(STRAINS):
                            default = 0
                            if cat == 'Flower' and sc == 'Whole Flower':
                                default = 4 if size in ['3.5g','7g'] and strain in ['Indica','Sativa'] else \
                                          2 if size in ['3.5g','7g'] and strain in ['Hybrid','Blend'] else 1
                            mb_subcat_targets[(cat, sc, size, strain)] = rc[i+1].number_input(
                                "", value=default, min_value=0, max_value=50,
                                key=f"mb_{cat}_{sc}_{size}_{strain}",
                                label_visibility="collapsed")
        else:
            with st.expander(f"{emoji} {cat} — by size & strain", expanded=(cat=='Flower')):
                if not sizes:
                    st.caption(f"No {cat} sizes detected in your catalogue.")
                    continue
                st.caption(f"{len(sizes)} sizes detected from your catalogue.")
                hdr = st.columns([2]+[1]*4)
                hdr[0].markdown("**Size**")
                for i, s in enumerate(STRAINS): hdr[i+1].markdown(f"**{s}**")
                for size in sizes:
                    rc = st.columns([2]+[1]*4)
                    rc[0].markdown(f"`{size}`")
                    for i, strain in enumerate(STRAINS):
                        default = 0
                        if cat == 'Flower':
                            default = 4 if size in ['3.5g','7g'] and strain in ['Indica','Sativa'] else \
                                      2 if size in ['3.5g','7g'] and strain in ['Hybrid','Blend'] else 1
                        mb_targets[(cat, size, strain)] = rc[i+1].number_input(
                            "", value=default, min_value=0, max_value=50,
                            key=f"mb_{cat}_{size}_{strain}", label_visibility="collapsed")

    st.markdown("---")

    # ── shelf counts ──────────────────────────────────────────
    on_shelf     = merged_raw[merged_raw['In Stock Qty'] > 0].copy()
    shelf_counts = on_shelf.groupby(['Classification','Product Size','Strain'])['SKU'].nunique().reset_index()
    shelf_counts.columns = ['Category','Size','Strain','On Shelf']

    # Sub-category shelf counts (for all cats that have sub-categories)
    subcat_shelf = pd.DataFrame()
    if has_subcat and subcat_map:
        sc_cats = list(subcat_map.keys())
        sc_raw = on_shelf[on_shelf['Classification'].isin(sc_cats)]
        if not sc_raw.empty:
            subcat_shelf = sc_raw.groupby(['Classification','Sub-Category','Product Size','Strain'])['SKU'].nunique().reset_index()
            subcat_shelf.columns = ['Category','Sub-Category','Size','Strain','On Shelf']

    # ── gap helpers ───────────────────────────────────────────
    def highlight_gap(val):
        if val > 0: return 'background-color:#FFC7CE;font-weight:bold'
        return 'background-color:#C6EFCE'

    def show_size_strain_grid(cat, sizes):
        rows = []
        for size in sizes:
            for strain in STRAINS:
                tgt = mb_targets.get((cat, size, strain), 0)
                cur = int(shelf_counts[
                    (shelf_counts['Category']==cat) &
                    (shelf_counts['Size']==size) &
                    (shelf_counts['Strain']==strain)
                ]['On Shelf'].sum())
                rows.append({'Size': size, 'Strain': strain, 'Target': tgt, 'On Shelf': cur, 'Gap': max(0, tgt-cur)})
        if not rows:
            return pd.DataFrame(0, index=sizes, columns=STRAINS)
        df = pd.DataFrame(rows)
        ps = df.pivot_table(index='Size', columns='Strain', values='On Shelf', fill_value=0).reindex(sizes).fillna(0).astype(int)
        pt = df.pivot_table(index='Size', columns='Strain', values='Target',   fill_value=0).reindex(sizes).fillna(0).astype(int)
        pg = df.pivot_table(index='Size', columns='Strain', values='Gap',      fill_value=0).reindex(sizes).fillna(0).astype(int)
        for piv in [ps, pt, pg]:
            for col in STRAINS:
                if col not in piv.columns: piv[col] = 0
        gc1, gc2, gc3 = st.columns(3)
        gc1.caption("On Shelf now");        gc1.dataframe(ps[STRAINS], use_container_width=True)
        gc2.caption("Your target");         gc2.dataframe(pt[STRAINS], use_container_width=True)
        gc3.caption("Gap (need to order)"); gc3.dataframe(pg[STRAINS].style.map(highlight_gap), use_container_width=True)
        return pg[STRAINS]

    def show_subcat_grid(cat, sc, sc_sizes):
        rows = []
        for size in sc_sizes:
            for strain in STRAINS:
                tgt = mb_subcat_targets.get((cat, sc, size, strain), 0)
                if subcat_shelf.empty:
                    cur = 0
                else:
                    cur = int(subcat_shelf[
                        (subcat_shelf['Category']==cat) &
                        (subcat_shelf['Sub-Category']==sc) &
                        (subcat_shelf['Size']==size) &
                        (subcat_shelf['Strain']==strain)
                    ]['On Shelf'].sum())
                rows.append({'Size': size, 'Strain': strain, 'Target': tgt, 'On Shelf': cur, 'Gap': max(0, tgt-cur)})
        if not rows:
            return pd.DataFrame(0, index=sc_sizes, columns=STRAINS)
        df = pd.DataFrame(rows)
        ps = df.pivot_table(index='Size', columns='Strain', values='On Shelf', fill_value=0).reindex(sc_sizes).fillna(0).astype(int)
        pt = df.pivot_table(index='Size', columns='Strain', values='Target',   fill_value=0).reindex(sc_sizes).fillna(0).astype(int)
        pg = df.pivot_table(index='Size', columns='Strain', values='Gap',      fill_value=0).reindex(sc_sizes).fillna(0).astype(int)
        for piv in [ps, pt, pg]:
            for col in STRAINS:
                if col not in piv.columns: piv[col] = 0
        gc1, gc2, gc3 = st.columns(3)
        gc1.caption("On Shelf now");        gc1.dataframe(ps[STRAINS], use_container_width=True)
        gc2.caption("Your target");         gc2.dataframe(pt[STRAINS], use_container_width=True)
        gc3.caption("Gap (need to order)"); gc3.dataframe(pg[STRAINS].style.map(highlight_gap), use_container_width=True)
        return pg[STRAINS]

    # ── gap display ───────────────────────────────────────────
    st.markdown("#### Current Shelf vs Target")
    cat_pivot_gaps = {}        # cat -> gap pivot  (no sub-categories)
    subcat_gaps    = {}        # (cat, subcat) -> gap pivot

    for cat in ALL_CATS:
        subcats = subcat_map.get(cat, [])
        if subcats:
            any_tgt = any(mb_subcat_targets.get((cat, sc, s, st_), 0) > 0
                          for sc in subcats
                          for s in subcat_sizes.get((cat, sc), [])
                          for st_ in STRAINS)
            if not any_tgt:
                continue
            st.markdown(f"**{CAT_EMOJIS[cat]} {cat}**")
            for sc in subcats:
                sc_sizes = subcat_sizes.get((cat, sc), [])
                if not sc_sizes: continue
                if not any(mb_subcat_targets.get((cat, sc, s, st_), 0) > 0 for s in sc_sizes for st_ in STRAINS):
                    continue
                st.markdown(f"*🔹 {sc}*")
                subcat_gaps[(cat, sc)] = show_subcat_grid(cat, sc, sc_sizes)
        else:
            sizes = cat_sizes[cat]
            if not sizes: continue
            if not any(mb_targets.get((cat, s, st_), 0) > 0 for s in sizes for st_ in STRAINS): continue
            st.markdown(f"**{CAT_EMOJIS[cat]} {cat}**")
            cat_pivot_gaps[cat] = show_size_strain_grid(cat, sizes)

    # ── suggestions ───────────────────────────────────────────
    total_gaps = (sum(int(pg.values.sum()) for pg in cat_pivot_gaps.values()) +
                  sum(int(pg.values.sum()) for pg in subcat_gaps.values()))
    if not cat_pivot_gaps and not subcat_gaps or total_gaps == 0:
        st.success("✅ Your shelf is fully stocked to target across all categories!")
    else:
        st.markdown("---")
        st.markdown("#### 💡 Suggested SKUs to Fill Gaps")
        st.caption("SKUs you've sold before (priority) or available on OCS — order 1 case each to fill gaps.")

        suggestions = []

        # Flat-category suggestions
        for cat, pg in cat_pivot_gaps.items():
            for size in cat_sizes[cat]:
                if size not in pg.index: continue
                for strain in STRAINS:
                    if strain not in pg.columns: continue
                    gap = int(pg.at[size, strain])
                    if gap <= 0: continue
                    candidates = merged_raw[
                        (merged_raw['Classification']==cat) &
                        (merged_raw['Product Size']==size) &
                        (merged_raw['Strain']==strain) &
                        (merged_raw['In Stock Qty']==0) &
                        (merged_raw['Stock Status']=='YES')
                    ].copy()
                    candidates['_sold'] = candidates['Sales (60 Days)'] > 0
                    candidates = candidates.sort_values(['_sold','Sales (60 Days)'], ascending=[False,False])
                    for _, r in candidates.head(gap).iterrows():
                        suggestions.append({
                            'Category': cat, 'Sub-Type': '', 'Size': size, 'Strain': strain,
                            'Product': r['Product'], 'OCS Variant #': r['Supplier Sku'],
                            'OCS Item #': str(r['OCS Item Number']).split('.')[0] if pd.notna(r.get('OCS Item Number')) else '—',
                            'Cases': 1, 'Pack Size': int(r['Pack Size']),
                            'Unit Price': r['Unit Price'] if pd.notna(r.get('Unit Price')) else None,
                            'Prev. Sold (60d)': int(r['Sales (60 Days)']) if pd.notna(r.get('Sales (60 Days)')) else 0,
                        })

        # Sub-category suggestions (Flower, Vapes, Topicals, etc.)
        for (cat, sc), pg in subcat_gaps.items():
            sc_sizes = subcat_sizes.get((cat, sc), [])
            for size in sc_sizes:
                if size not in pg.index: continue
                for strain in STRAINS:
                    if strain not in pg.columns: continue
                    gap = int(pg.at[size, strain])
                    if gap <= 0: continue
                    cand_mask = (
                        (merged_raw['Classification']==cat) &
                        (merged_raw['Product Size']==size) &
                        (merged_raw['Strain']==strain) &
                        (merged_raw['In Stock Qty']==0) &
                        (merged_raw['Stock Status']=='YES')
                    )
                    if has_subcat:
                        cand_mask &= (merged_raw['Sub-Category']==sc)
                    candidates = merged_raw[cand_mask].copy()
                    candidates['_sold'] = candidates['Sales (60 Days)'] > 0
                    candidates = candidates.sort_values(['_sold','Sales (60 Days)'], ascending=[False,False])
                    for _, r in candidates.head(gap).iterrows():
                        suggestions.append({
                            'Category': cat, 'Sub-Type': sc, 'Size': size, 'Strain': strain,
                            'Product': r['Product'], 'OCS Variant #': r['Supplier Sku'],
                            'OCS Item #': str(r['OCS Item Number']).split('.')[0] if pd.notna(r.get('OCS Item Number')) else '—',
                            'Cases': 1, 'Pack Size': int(r['Pack Size']),
                            'Unit Price': r['Unit Price'] if pd.notna(r.get('Unit Price')) else None,
                            'Prev. Sold (60d)': int(r['Sales (60 Days)']) if pd.notna(r.get('Sales (60 Days)')) else 0,
                        })

        if suggestions:
            sug_df = pd.DataFrame(suggestions)
            if sug_df['Sub-Type'].eq('').all():
                sug_df = sug_df.drop(columns=['Sub-Type'])
            sug_df['Est. Cost'] = (sug_df['Cases'] * sug_df['Pack Size'] * sug_df['Unit Price'].fillna(0)).round(2)
            sug_df['Unit Price'] = sug_df['Unit Price'].map(lambda x: f'${x:,.2f}' if pd.notna(x) and x > 0 else '—')
            sug_df['Est. Cost']  = sug_df['Est. Cost'].map(lambda x: f'${x:,.2f}' if x > 0 else '—')
            st.dataframe(sug_df, hide_index=True, use_container_width=True)

            menu_upload = pd.DataFrame([{'SKU': r['OCS Variant #'], 'Quantity': r['Cases']} for r in suggestions])
            menu_buf = io.BytesIO()
            with pd.ExcelWriter(menu_buf, engine='openpyxl') as writer:
                menu_upload.to_excel(writer, sheet_name='MasterCatalogue', index=False)
            menu_buf.seek(0)
            st.download_button(
                label="📤 Download OCS Upload File for Menu Order (.xlsx)",
                data=menu_buf,
                file_name=f'OCS_MenuOrder_{today2.strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )
        else:
            st.info("No matching SKUs found in the OCS catalogue to fill the gaps. Try adjusting your targets or check the catalogue for availability.")
