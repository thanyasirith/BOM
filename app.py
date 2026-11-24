import pandas as pd
import streamlit as st
import io

# --- 1. App Configuration ---
st.set_page_config(
    page_title="BOM Compare Program Pro",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. Custom CSS UI Styling ---
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(90deg, #1E3C72 0%, #2A5298 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #6c757d;
        margin-bottom: 20px;
    }
    .metric-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 15px 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-left: 5px solid #2A5298;
        transition: all 0.3s ease;
    }
    .metric-title {
        font-size: 0.8rem;
        color: #6c757d;
        font-weight: 600;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #1E3C72;
        margin-top: 2px;
    }
    .stDataFrame {
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🛠️ BOM Compare Program Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated Multi-BOM Analytics & Comparison System (Bluechips vs Zollner)</div>', unsafe_allow_html=True)

# --- 3. File Upload Section ---
st.info("💡 **Getting Started:** Upload an Excel file containing 2 sheets named **BCM** and **ZN**")
uploaded_file = st.file_uploader("📂 Upload BOM File (.xlsx)", type=["xlsx"])

# --- 4. Data Processing Logic ---
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        if 'BCM' not in xls.sheet_names or 'ZN' not in xls.sheet_names:
            st.error("❌ Error: Could not find required sheets 'BCM' or 'ZN' in the uploaded file.")
            st.stop()

        bcm_df = pd.read_excel(xls, sheet_name='BCM')
        zollner_df = pd.read_excel(xls, sheet_name='ZN')

        bcm_df.columns = [str(c).strip() for c in bcm_df.columns]
        zollner_df.columns = [str(c).strip() for c in zollner_df.columns]

        bcm_fg_col = None
        for col in ['BCM FG', 'FG', 'Material']:
            if col in bcm_df.columns:
                bcm_df.rename(columns={col: 'BCM FG'}, inplace=True)
                bcm_fg_col = 'BCM FG'
                break

        zn_fg_col = None
        for col in ['BCM FG', 'FG']:
            if col in zollner_df.columns:
                zollner_df.rename(columns={col: 'BCM FG'}, inplace=True)
                zn_fg_col = 'BCM FG'
                break

        if not bcm_fg_col or not zn_fg_col:
            st.error(f"❌ Error: Could not find BCM FG column. BCM columns found: {list(bcm_df.columns)}")
            st.stop()

        bcm_df['BCM FG'] = bcm_df['BCM FG'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        zollner_df['BCM FG'] = zollner_df['BCM FG'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        fg_list_bcm = [f for f in bcm_df['BCM FG'].unique() if str(f).strip() and str(f).lower() != 'nan']
        fg_list_zn = [f for f in zollner_df['BCM FG'].unique() if str(f).strip() and str(f).lower() != 'nan']
        all_fgs = sorted(list(set(fg_list_bcm + fg_list_zn)))

        col_select, col_btn = st.columns([3, 1])
        with col_select:
            selected_fg = st.selectbox("🎯 Select BCM FG BOM to compare:", ["All FGs"] + all_fgs)
        with col_btn:
            st.write(" ")
            st.write(" ")
            run_btn = st.button("🚀 Run Comparison", use_container_width=True, type="primary")

        if run_btn:
            with st.spinner("⚡ Comparing BOMs and Generating Executive Dashboard..."):

                if selected_fg != "All FGs":
                    bcm_df = bcm_df[bcm_df['BCM FG'] == selected_fg].copy()
                    zollner_df = zollner_df[zollner_df['BCM FG'] == selected_fg].copy()

                def find_bcm_article_col_idx(df):
                    for idx, col in enumerate(df.columns):
                        c_lower = str(col).strip().lower()
                        if 'bcm article' in c_lower or 'bcm_article' in c_lower:
                            return idx
                    for idx, col in enumerate(df.columns):
                        c_lower = str(col).strip().lower()
                        if 'article' in c_lower or 'component' in c_lower:
                            return idx
                    return 6

                def find_alt_item_group_idx(df):
                    for idx, col in enumerate(df.columns):
                        c_lower = str(col).strip().lower()
                        if 'altitemgroup' in c_lower or 'alt item group' in c_lower or 'alt_item_group' in c_lower:
                            return idx
                    return len(df.columns) - 1

                def find_bcm_zn_mapping_col_idx(df):
                    bcm_art_idx = find_bcm_article_col_idx(df)
                    alt_group_idx = find_alt_item_group_idx(df)

                    for idx in range(bcm_art_idx + 1, min(alt_group_idx, len(df.columns))):
                        col_name = str(df.columns[idx]).strip().lower()
                        if 'zn' in col_name or 'zollner' in col_name or 'art' in col_name:
                            return idx
                    if bcm_art_idx + 1 < len(df.columns):
                        return bcm_art_idx + 1
                    return None

                bcm_art_idx = find_bcm_article_col_idx(bcm_df)
                bcm_article_col = bcm_df.columns[bcm_art_idx]

                bcm_qty_col = None
                for col in ['Component quantity', 'Usage probability']:
                    if col in bcm_df.columns:
                        bcm_qty_col = col
                        break

                zn_article_col = 'BCM Article' if 'BCM Article' in zollner_df.columns else ('Component' if 'Component' in zollner_df.columns else None)
                zn_qty_col = 'Component Quantity' if 'Component Quantity' in zollner_df.columns else ('Component quantity' if 'Component quantity' in zollner_df.columns else None)

                if not bcm_qty_col or not bcm_article_col or not zn_article_col:
                    st.error("❌ Could not identify required Article or Quantity columns.")
                    st.stop()

                bcm_mapping_col_idx = find_bcm_zn_mapping_col_idx(bcm_df)

                bcm_df['BCM_Article_Clean'] = bcm_df[bcm_article_col].astype(str).str.strip()
                bcm_df['Display_Article'] = bcm_df['BCM_Article_Clean']

                missing_mapping_df = pd.DataFrame(columns=['BCM FG', 'AltItemGroup', 'Missing_ZN_Mapping'])

                if bcm_mapping_col_idx is not None:
                    mapping_values = bcm_df.iloc[:, bcm_mapping_col_idx].astype(str).str.strip().str.lower()
                    is_missing = bcm_df.iloc[:, bcm_mapping_col_idx].isna() | (mapping_values == '') | (mapping_values == 'nan') | (mapping_values == 'none')

                    bcm_df.loc[is_missing, 'Display_Article'] = bcm_df.loc[is_missing, 'BCM_Article_Clean'].apply(
                        lambda x: f"[{x}]" if x and x.lower() != 'nan' else x
                    )

                    missing_rows = bcm_df[is_missing].copy()
                    if not missing_rows.empty:
                        missing_mapping_df = missing_rows.groupby(['BCM FG', 'AltItemGroup'])['BCM_Article_Clean'].apply(
                            lambda x: ', '.join(sorted(set([str(i) for i in x if str(i) and str(i).lower() != 'nan'])))
                        ).reset_index()
                        missing_mapping_df.rename(columns={'BCM_Article_Clean': 'Missing_ZN_Mapping'}, inplace=True)

                # --- Match Key Logic ---
                bcm_map = bcm_df[['BCM FG', 'BCM_Article_Clean', 'AltItemGroup', bcm_qty_col]].drop_duplicates()
                bcm_map = bcm_map.sort_values(by=['BCM FG', 'BCM_Article_Clean', bcm_qty_col], ascending=[True, True, False])
                bcm_map['Rank'] = bcm_map.groupby(['BCM FG', 'BCM_Article_Clean']).cumcount() + 1
                bcm_map['Match_Key'] = bcm_map['BCM FG'] + "_" + bcm_map['BCM_Article_Clean'] + "_" + bcm_map['Rank'].astype(str)

                zollner_prep = zollner_df.copy()
                zollner_prep[zn_article_col] = zollner_prep[zn_article_col].astype(str).str.strip()
                zollner_prep = zollner_prep.sort_values(by=['BCM FG', zn_article_col, zn_qty_col], ascending=[True, True, False])
                zollner_prep['Rank'] = zollner_prep.groupby(['BCM FG', zn_article_col]).cumcount() + 1
                zollner_prep['Match_Key'] = zollner_prep['BCM FG'] + "_" + zollner_prep[zn_article_col] + "_" + zollner_prep['Rank'].astype(str)

                zollner_mapped = pd.merge(zollner_prep, bcm_map[['Match_Key', 'AltItemGroup']], on='Match_Key', how='left')

                def clean_join_display(series):
                    items = set()
                    for val in series.dropna():
                        s_val = str(val).strip()
                        if s_val and s_val.lower() not in ['nan', 'none']:
                            items.add(s_val)
                    sorted_items = sorted(items, key=lambda s: s.replace('[', '').replace(']', '').strip())
                    return ', '.join(sorted_items)

                def clean_join(series):
                    items = set()
                    for val in series.dropna():
                        s_val = str(val).strip()
                        if s_val and s_val.lower() not in ['nan', 'none']:
                            items.add(s_val)
                    return ', '.join(sorted(items))

                def clean_designators(series):
                    items = set()
                    for val in series.dropna():
                        s_val = str(val).replace(';', ',').replace(' ', '')
                        for part in s_val.split(','):
                            p = part.strip()
                            if p and p.lower() not in ['nan', 'none']:
                                items.add(p)
                    return ', '.join(sorted(items))

                bcm_grouped = bcm_df.groupby(['BCM FG', 'AltItemGroup'], as_index=False).agg({
                    bcm_qty_col: 'max', 
                    'Display_Article': clean_join_display, 
                    'BCM_Article_Clean': clean_join,   
                    'Designator': clean_designators if 'Designator' in bcm_df.columns else lambda x: ''
                })
                bcm_grouped.rename(columns={bcm_qty_col: 'Qty_BCM', 'Designator': 'BCM_Designator'}, inplace=True)

                zn_comp_col = 'ZN art.' if 'ZN art.' in zollner_df.columns else ('Component' if 'Component' in zollner_df.columns else zn_article_col)
                zollner_grouped = zollner_mapped.groupby(['BCM FG', 'AltItemGroup'], as_index=False).agg({
                    zn_qty_col: 'max',
                    zn_comp_col: clean_join,
                    'Designator': clean_designators if 'Designator' in zollner_df.columns else lambda x: ''
                })
                zollner_grouped.rename(columns={zn_qty_col: 'Qty_Zollner', zn_comp_col: 'Zollner_Articles', 'Designator': 'Zollner_Designator'}, inplace=True)

                result_df = pd.merge(bcm_grouped, zollner_grouped, on=['BCM FG', 'AltItemGroup'], how='outer')
                result_df = pd.merge(result_df, missing_mapping_df, on=['BCM FG', 'AltItemGroup'], how='left')
                result_df.fillna({'Qty_BCM': 0, 'Qty_Zollner': 0, 'Missing_ZN_Mapping': '', 'BCM_Designator': '', 'Zollner_Designator': ''}, inplace=True)

                result_df['Qty_BCM'] = pd.to_numeric(result_df['Qty_BCM'], errors='coerce').fillna(0).round(3)
                result_df['Qty_Zollner'] = pd.to_numeric(result_df['Qty_Zollner'], errors='coerce').fillna(0).round(3)

                def parse_desig_set(desig_str):
                    if pd.isna(desig_str): return set()
                    return set([x.strip() for x in str(desig_str).split(',') if x.strip() and x.strip().lower() not in ['nan', 'none']])

                def get_status(row):
                    zn_art = str(row['Zollner_Articles']).strip().lower()
                    if pd.isna(row['Zollner_Articles']) or zn_art == '' or zn_art in ['nan', 'none']:
                        return '⚠️ ZN Article Blank'
                    if str(row.get('Missing_ZN_Mapping', '')).strip() != '':
                        return '⚠️ Missing ZN Mapping'
                    if row['Qty_BCM'] == 0: return '❌ Missing in BCM'
                    if row['Qty_Zollner'] == 0: return '❌ Missing in ZN'
                    if row['Qty_BCM'] != row['Qty_Zollner']: return '⚠️ Qty Mismatch'

                    b_desig = parse_desig_set(row['BCM_Designator'])
                    z_desig = parse_desig_set(row['Zollner_Designator'])
                    if b_desig != z_desig and (b_desig or z_desig):
                        return '⚠️ Desig Mismatch'

                    return '✅ Match'

                def get_remark(row):
                    remarks = []
                    zn_art = str(row['Zollner_Articles']).strip().lower()
                    if pd.isna(row['Zollner_Articles']) or zn_art == '' or zn_art in ['nan', 'none']:
                        remarks.append('ZN did not provide an Article')
                    if str(row['Missing_ZN_Mapping']).strip() != '':
                        remarks.append(f'Missing ZN Mapping for BCM: [{row["Missing_ZN_Mapping"]}]')

                    b_desig = parse_desig_set(row['BCM_Designator'])
                    z_desig = parse_desig_set(row['Zollner_Designator'])
                    if b_desig != z_desig and (b_desig or z_desig):
                        diff_bcm = z_desig - b_desig
                        diff_zn = b_desig - z_desig
                        desig_msgs = []
                        if diff_bcm:
                            desig_msgs.append(f"Missing in BCM: {','.join(sorted(diff_bcm))}")
                        if diff_zn:
                            desig_msgs.append(f"Missing in ZN: {','.join(sorted(diff_zn))}")
                        remarks.append(f"Designator Mismatch ({' | '.join(desig_msgs)})")

                    return ' | '.join(remarks)

                result_df['Status'] = result_df.apply(get_status, axis=1)
                result_df['Remark'] = result_df.apply(get_remark, axis=1)

                def highlight_rows(row):
                    if '⚠️' in str(row['Status']) or 'Missing ZN Mapping' in str(row['Remark']) or 'Designator Mismatch' in str(row['Remark']):
                        return ['background-color: #ffcccc'] * len(row)
                    elif '❌' in str(row['Status']):
                        return ['background-color: #cce5ff'] * len(row)
                    else:
                        return [''] * len(row)

                df_web = result_df[['BCM FG', 'AltItemGroup', 'BCM_Designator', 'Zollner_Designator', 'Display_Article', 'Zollner_Articles', 'Qty_BCM', 'Qty_Zollner', 'Status', 'Remark']].copy()
                df_web.rename(columns={'Display_Article': 'BCM_Articles'}, inplace=True)

                df_excel = result_df[['BCM FG', 'AltItemGroup', 'BCM_Designator', 'Zollner_Designator', 'BCM_Article_Clean', 'Zollner_Articles', 'Qty_BCM', 'Qty_Zollner', 'Status', 'Remark']].copy()
                df_excel.rename(columns={'BCM_Article_Clean': 'BCM_Articles'}, inplace=True)

                # --- 5. Interactive Dashboard & Output UI ---
                st.write("---")

                total_groups = len(result_df)
                perfect_match = len(result_df[(result_df['Status'] == '✅ Match') & (result_df['Remark'] == '')])
                mismatches = len(result_df[result_df['Status'].str.contains('⚠️', na=False)])
                critical_missing = len(result_df[result_df['Status'].str.contains('❌', na=False)])

                tab1, tab2, tab3 = st.tabs(["📊 Executive Dashboard", "📋 Detailed BOM Table", "🔍 Issue Inspector"])

                with tab1:
                    st.markdown("### 📊 Key Performance Indicators")
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        st.markdown(f'<div class="metric-card"><div class="metric-title">📌 Total Alt Groups</div><div class="metric-value">{total_groups}</div></div>', unsafe_allow_html=True)
                    with c2:
                        st.markdown(f'<div class="metric-card" style="border-left-color: #28a745;"><div class="metric-title">✅ Perfect Matches</div><div class="metric-value" style="color: #28a745;">{perfect_match}</div></div>', unsafe_allow_html=True)
                    with c3:
                        st.markdown(f'<div class="metric-card" style="border-left-color: #ffc107;"><div class="metric-title">⚠️ Mismatches & Warnings</div><div class="metric-value" style="color: #d39e00;">{mismatches}</div></div>', unsafe_allow_html=True)
                    with c4:
                        st.markdown(f'<div class="metric-card" style="border-left-color: #dc3545;"><div class="metric-title">❌ Missing Components</div><div class="metric-value" style="color: #dc3545;">{critical_missing}</div></div>', unsafe_allow_html=True)

                    st.write(" ")
                    st.markdown("### 📋 FG-Level Summary Table (ตารางสรุปภาพรวมแต่ละ BCM FG)")

                    # 🌟 สร้างตารางสรุปยอดแยกตาม BCM FG แบบแนวนอน
                    fg_summary = result_df.groupby('BCM FG').agg(
                        Total_Items=('AltItemGroup', 'count'),
                        Perfect_Match=('Status', lambda x: (x == '✅ Match').sum()),
                        Qty_Mismatch=('Status', lambda x: (x == '⚠️ Qty Mismatch').sum()),
                        Desig_Mismatch=('Status', lambda x: (x == '⚠️ Desig Mismatch').sum()),
                        Missing_Mapping=('Status', lambda x: (x == '⚠️ Missing ZN Mapping').sum()),
                        ZN_Blank=('Status', lambda x: (x == '⚠️ ZN Article Blank').sum()),
                        Missing_Items=('Status', lambda x: x.str.contains('❌').sum())
                    ).reset_index()

                    st.dataframe(
                        fg_summary,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "BCM FG": st.column_config.TextColumn("BCM FG", width="medium"),
                            "Total_Items": st.column_config.NumberColumn("Total Items", width="small"),
                            "Perfect_Match": st.column_config.NumberColumn("✅ Match", width="small"),
                            "Qty_Mismatch": st.column_config.NumberColumn("⚠️ Qty Mismatch", width="small"),
                            "Desig_Mismatch": st.column_config.NumberColumn("⚠️ Desig Mismatch", width="small"),
                            "Missing_Mapping": st.column_config.NumberColumn("⚠️ Missing Mapping", width="small"),
                            "ZN_Blank": st.column_config.NumberColumn("⚠️ ZN Blank", width="small"),
                            "Missing_Items": st.column_config.NumberColumn("❌ Missing", width="small"),
                        }
                    )

                    st.write(" ")
                    st.markdown("### 📈 Status Distribution Breakdown")
                    status_summary = result_df['Status'].value_counts().reset_index()
                    status_summary.columns = ['Status', 'Count']
                    st.bar_chart(status_summary.set_index('Status'))

                with tab2:
                    st.markdown("### 📋 BOM Comparison Master Table")
                    show_only_issues = st.checkbox("🔍 Show Mismatches & Issues Only", value=False)

                    display_df = df_web.copy()
                    if show_only_issues:
                        display_df = display_df[display_df['Status'] != '✅ Match']

                    st.dataframe(
                        display_df.style.apply(highlight_rows, axis=1).format(precision=3), 
                        use_container_width=True, 
                        height=500,
                        column_config={
                            "BCM FG": st.column_config.TextColumn("BCM FG", width="small"),
                            "AltItemGroup": st.column_config.TextColumn("AltItemGroup", width="small"),
                            "BCM_Designator": st.column_config.TextColumn("BCM_Designator", width="medium"),
                            "Zollner_Designator": st.column_config.TextColumn("Zollner_Designator", width="medium"),
                            "BCM_Articles": st.column_config.TextColumn("BCM_Articles", width="medium"),
                            "Zollner_Articles": st.column_config.TextColumn("Zollner_Articles", width="medium"),
                            "Qty_BCM": st.column_config.NumberColumn("Qty_BCM", width="small"),
                            "Qty_Zollner": st.column_config.NumberColumn("Qty_Zollner", width="small"),
                            "Status": st.column_config.TextColumn("Status", width="medium"),
                            "Remark": st.column_config.TextColumn("Remark", width="large"),
                        }
                    )

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_excel.style.apply(highlight_rows, axis=1).to_excel(writer, index=False, sheet_name='Comparison_Result')
                    output.seek(0)

                    st.download_button(
                        label="📥 Download Detailed Report (Excel)",
                        data=output,
                        file_name=f"BOM_Comparison_{selected_fg}_Pro_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )

                with tab3:
                    st.markdown("### 🔍 Interactive Issue Inspector")
                    issues_only_df = df_web[df_web['Status'] != '✅ Match'].copy()

                    if not issues_only_df.empty:
                        status_filter = st.multiselect(
                            "Filter by Specific Issue Status:",
                            options=issues_only_df['Status'].unique().tolist(),
                            default=issues_only_df['Status'].unique().tolist()
                        )

                        filtered_issues = issues_only_df[issues_only_df['Status'].isin(status_filter)]
                        st.dataframe(
                            filtered_issues.style.apply(highlight_rows, axis=1).format(precision=3), 
                            use_container_width=True,
                            height=500,
                            column_config={
                                "BCM FG": st.column_config.TextColumn("BCM FG", width="small"),
                                "AltItemGroup": st.column_config.TextColumn("AltItemGroup", width="small"),
                                "BCM_Designator": st.column_config.TextColumn("BCM_Designator", width="medium"),
                                "Zollner_Designator": st.column_config.TextColumn("Zollner_Designator", width="medium"),
                                "BCM_Articles": st.column_config.TextColumn("BCM_Articles", width="medium"),
                                "Zollner_Articles": st.column_config.TextColumn("Zollner_Articles", width="medium"),
                                "Qty_BCM": st.column_config.NumberColumn("Qty_BCM", width="small"),
                                "Qty_Zollner": st.column_config.NumberColumn("Qty_Zollner", width="small"),
                                "Status": st.column_config.TextColumn("Status", width="medium"),
                                "Remark": st.column_config.TextColumn("Remark", width="large"),
                            }
                        )
                    else:
                        st.balloons()
                        st.success("✨ No Issues Found! All BOM items match perfectly across both systems.")

    except Exception as e:
        st.error(f"Error during processing: {e}")
