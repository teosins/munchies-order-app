import streamlit as st
import pandas as pd
import math
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Canna Order Tool", page_icon="🌿", layout="wide")

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
    st.markdown("## 🌿 Canna Order Tool")
    st.markdown("---")
    st.header("📁 Upload Files")
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
    budget_pretax = round(budget / (1 + tax_rate), 2)
    tax_amount = budget - budget_pretax
    st.caption(f"Pre-tax: ${budget_pretax:,.2f}  |  Tax ({tax_rate*100:.3g}%): ${tax_amount:,.2f}")

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
st.title("🌿 Canna Order Tool")
st.caption(f"Generating orders for: {date.today().strftime('%B %d, %Y')}")

if not kova_file or not ocs_file:
    st.info("Upload your **Cova Reorder Report** and **OCS Catalogue** in the sidebar to get started.")
    st.markdown("""
    **How to export from Cova:**
    Reports → Reorder → Export as .xlsx

    **How to export from OCS Wholesale:**
    Catalogue → Download Catalogue → Excel
    """)
    st.stop()

# ── process data ──────────────────────────────────────────────
@st.cache_data(show_spinner="Processing your data...")
def process(kova_bytes, ocs_bytes, target, flower_size_target, budget_pretax):
    kova = pd.read_excel(io.BytesIO(kova_bytes), sheet_name='Reorder')
    ocs  = pd.read_excel(io.BytesIO(ocs_bytes),  sheet_name='MasterCatalogue')

    kova_ocs = kova[kova['Supplier'] == 'OCS'].copy()
    kova_ocs['Supplier Sku'] = kova_ocs['Supplier Sku'].str.strip()
    ocs['OCS Variant Number'] = ocs['OCS Variant Number'].str.strip()

    ocs_cols = ['OCS Variant Number','OCS Item Number','Unit Price','Pack Size','Stock Status']
    merged = kova_ocs.merge(ocs[ocs_cols], left_on='Supplier Sku', right_on='OCS Variant Number', how='left')

    active = merged[merged['Sales (30 Days)'] > 0].copy()
    active['Pack Size']  = pd.to_numeric(active['Pack Size'],  errors='coerce').fillna(1).astype(int)
    active['Unit Price'] = pd.to_numeric(active['Unit Price'], errors='coerce')
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

    import re
    def extract_flower_size(sku):
        m = re.search(r'_(\d+\.?\d*[gG])_', str(sku))
        return m.group(1).lower() if m else None
    active['Flower Size'] = active.apply(
        lambda r: extract_flower_size(r['Supplier Sku']) if r['Classification'] == 'Flower' else None, axis=1)

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

kova_bytes = kova_file.read()
ocs_bytes  = ocs_file.read()
order_df, deferred_df, all_active = process(kova_bytes, ocs_bytes, TARGET, FLOWER_SIZE_TARGET, budget_pretax)

if order_df.empty:
    st.warning("No items need ordering based on current stock levels and settings.")
    st.stop()

pretax_total = order_df['Est Cost'].sum()
tax_amount_actual = pretax_total * tax_rate
total        = pretax_total + tax_amount_actual
tax_label    = f"Tax ({tax_rate*100:.3g}%)"

# ── summary metrics ───────────────────────────────────────────
st.markdown("### This Week's Order")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("SKUs to Order",  len(order_df))
m2.metric("Pre-Tax Total",  f"${pretax_total:,.2f}")
m3.metric(tax_label,        f"${tax_amount_actual:,.2f}")
m4.metric("Total (Tax-In)", f"${total:,.2f}")
m5.metric("Deferred SKUs",  len(deferred_df))

st.markdown("---")

# ── category summary ──────────────────────────────────────────
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

# ── order table ───────────────────────────────────────────────
st.markdown("### Order Detail")
display_cols = {
    'Classification':'Category','Product':'Product','Tier':'Tier',
    'Days Left':'Days Left','In Stock Qty':'In Stock','On Order':'On Order',
    'Sales (7 Days)':'7d Sales','Sales (30 Days)':'30d Sales',
    'Weekly Vel':'Wkly Vel','Cases':'Cases','Suggest':'Units',
    'Unit Price':'Unit Price','Est Cost':'Est Cost','Supplier Sku':'OCS Variant #'
}
show = order_df[[c for c in display_cols.keys() if c in order_df.columns]].rename(columns=display_cols)
show['Unit Price'] = show['Unit Price'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
show['Est Cost']   = show['Est Cost'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
show['Days Left']  = show['Days Left'].round(1)

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

styled = show.style.map(color_tier, subset=['Tier']).map(color_days, subset=['Days Left'])
st.dataframe(styled, hide_index=True, use_container_width=True, height=500)

if not deferred_df.empty:
    with st.expander(f"📋 Deferred to Next Order ({len(deferred_df)} SKUs — ${deferred_df['Est Cost'].sum():,.2f} pre-tax)"):
        def_show = deferred_df[['Classification','Product','Tier','Days Left','In Stock Qty','Weekly Vel','Cases','Est Cost']].copy()
        def_show['Est Cost'] = def_show['Est Cost'].map(lambda x: f'${x:,.2f}' if pd.notna(x) else '—')
        st.dataframe(def_show.rename(columns={'Classification':'Category','In Stock Qty':'In Stock','Weekly Vel':'Wkly Vel'}),
                     hide_index=True, use_container_width=True)

st.markdown("---")

# ── generate files ────────────────────────────────────────────
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

st.markdown("### Download Files")
dl1, dl2 = st.columns(2)

today = date.today()

with dl1:
    workbook_buf = build_workbook(order_df, deferred_df, all_active, budget_pretax, TARGET, today, tax_rate)
    st.download_button(
        label="📥 Download Order Workbook (.xlsx)",
        data=workbook_buf,
        file_name=f'OCS_Order_Tool_{today.strftime("%Y%m%d")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )

with dl2:
    upload_buf = build_upload(order_df)
    st.download_button(
        label="📤 Download OCS Upload File (.xlsx)",
        data=upload_buf,
        file_name=f'OCS_Upload_{today.strftime("%Y%m%d")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
