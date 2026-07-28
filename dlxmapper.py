import streamlit as st # type: ignore
import pandas as pd # type: ignore
import io
import os

# Set a wider layout for better dataframe viewing
st.set_page_config(
    page_title="DLX Mapper", 
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state for the uploader key to allow resetting
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# Initialize mode in session state if not present
if "mode" not in st.session_state:
    st.session_state.mode = "Demand Letter with Transmittal"  # default

# Initialize DL type for Transmittal Only mode
if "dl_type" not in st.session_state:
    st.session_state.dl_type = "DL1"  # default

# App Header Styling with Reset Button
col_header, col_reset = st.columns([4, 1])
with col_header:
    st.title("🎯 DLX Mapper")
with col_reset:
    st.write("") 
    if st.button("🔄 Reset App", use_container_width=True, type="secondary"):
        st.session_state.uploader_key += 1
        st.rerun()

st.write("---")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# =========================================================
# STEP 0: MODE SELECTION (Demand Letter vs Transmittal Only)
# =========================================================
st.subheader("📌 Select Output Type")

# Two buttons for mode selection
col_mode1, col_mode2 = st.columns(2)
with col_mode1:
    if st.button(
        "📄 Demand Letter",
        use_container_width=True,
        type="primary" if st.session_state.mode == "Demand Letter with Transmittal" else "secondary"
    ):
        st.session_state.mode = "Demand Letter with Transmittal"
        st.rerun()
with col_mode2:
    if st.button(
        "📨 Transmittal Only",
        use_container_width=True,
        type="primary" if st.session_state.mode == "Transmittal Only" else "secondary"
    ):
        st.session_state.mode = "Transmittal Only"
        st.rerun()

# Define campaign lists per mode
demand_letter_campaigns = [
    "CBS HOUSING LOAN",
    "PIF FORECLOSURE",
    "PIF HOME LOAN"
]

transmittal_only_campaigns = [
    "SBC HOME LOAN",
    "BPI",
    "BPI BANKO",
    "BPI CARDS EARLY",
    "BPI CARDS 30DPD",
    "BPI CARDS XDAYS",
    "BPI PL XDAYS",
    "BPI PL 30DPD",
    "BPI PL 60DPD",
    "BPI RBANK CARDS 30DPD",
    "BPI RBANK CARDS 60DPD",
    "ROBINSONS BPI",
    "BPI AUTO",
    "BDO HOMELOAN SKIP",
    "BDO 60DPD",
    "BDO AUTO LOAN"
]

# Choose the appropriate campaign list based on current mode and sort alphabetically
if st.session_state.mode == "Demand Letter with Transmittal":
    client_name_options = sorted(demand_letter_campaigns)
else:
    client_name_options = sorted(transmittal_only_campaigns)

st.write("---")

# 1. STEP 1: Configuration Control Panel (Moved to top)
st.subheader("⚙️ Select Section")

with st.container(border=True):
    # Use three columns: Campaign, Template, DL Type / Filter
    col1, col2, col3 = st.columns(3, gap="large")
    
    with col1:
        selected_client_name = st.selectbox(
            "👤 Select Campaign:", 
            client_name_options,
            help="Determines the target folder directory layout structure to pull the template configuration from."
        )
        
    with col2:
        # Determine the folder to look for templates based on mode
        if st.session_state.mode == "Transmittal Only":
            client_dir = os.path.join(SCRIPT_DIR, "Transmittal")
        else:
            client_dir = os.path.join(SCRIPT_DIR, selected_client_name)
        
        template_options = []
        if os.path.exists(client_dir) and os.path.isdir(client_dir):
            for file in os.listdir(client_dir):
                if file.endswith('.xlsx') and not file.startswith('~$'):
                    template_options.append(os.path.splitext(file)[0])
        template_options.sort()
        
        if not template_options:
            template_options = ["No templates found in folder"]
            
        selected_template = st.selectbox(
            "📐 Select Template:", 
            template_options,
            help="Select the specific DL blueprint mapping template layout required for file generations."
        )

    with col3:
        # For Transmittal Only: show DL Type dropdown
        if st.session_state.mode == "Transmittal Only":
            dl_type_options = ["DL1", "DL4", "DL11", "DL12", "DL13"]
            if st.session_state.dl_type not in dl_type_options:
                st.session_state.dl_type = dl_type_options[0]
            selected_dl_type = st.selectbox(
                "📑 DL Type:",
                dl_type_options,
                index=dl_type_options.index(st.session_state.dl_type),
                help="Select the Demand Letter type for Transmittal Only output."
            )
            st.session_state.dl_type = selected_dl_type
        else:
            # Demand Letter mode: placeholder for filter (will be filled after file upload)
            filter_placeholder = st.empty()
            st.session_state.filter_placeholder = filter_placeholder

# =========================================================
# DYNAMIC FILE PATH ROUTING & DIAGNOSTICS
# =========================================================

template_filename = os.path.join(client_dir, f"{selected_template}.xlsx")
template_exists = os.path.exists(template_filename)

required_lookup_name = None
if st.session_state.mode == "Demand Letter with Transmittal":
    if selected_client_name == "PIF HOME LOAN":
        required_lookup_name = "pif FOR DLX.xlsx"
    elif selected_client_name == "PIF FORECLOSURE":
        required_lookup_name = "pif fcl FOR DLX.xlsx"

if required_lookup_name:
    lookup_filename = os.path.join(SCRIPT_DIR, required_lookup_name)
    lookup_exists = os.path.exists(lookup_filename)
else:
    lookup_filename = None
    lookup_exists = False

if required_lookup_name:
    diagnostics_expanded = (not template_exists) or (not lookup_exists)
else:
    diagnostics_expanded = not template_exists

with st.expander("🔍 System File Path Diagnostics", expanded=diagnostics_expanded):
    c_left, c_right = st.columns(2)
    with c_left:
        if template_exists:
            st.success(f"📁 Template Found: `{os.path.basename(client_dir)}/{selected_template}.xlsx`")
        else:
            st.error(f"❌ Missing Template File: `{selected_template}.xlsx`")
            st.code(f"Required path: {template_filename}", language="bash")
            
    with c_right:
        if required_lookup_name:
            if lookup_exists:
                st.success(f"📊 Reference Database Found: `{required_lookup_name}`")
            else:
                st.warning(f"⚠️ Reference File `{required_lookup_name}` not found (Required for {selected_client_name})")
                st.code(f"Required path: {lookup_filename}", language="bash")
        else:
            st.info(f"ℹ️ No Reference Database required for `{selected_client_name}` in this mode.")

st.write("---")

# 2. STEP 2: Main Source File Upload
st.subheader("📥 Upload Source File")
uploaded_file = st.file_uploader(
    label="Drag and drop your source file configuration (Excel or CSV formats supported)", 
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
    key=f"file_uploader_{st.session_state.uploader_key}" 
)

# We'll store the uploaded file data in session state to reuse in the filter placeholder
if "df_source" not in st.session_state:
    st.session_state.df_source = None
if "dl_filter_options" not in st.session_state:
    st.session_state.dl_filter_options = []
if "dl_filter_counts" not in st.session_state:
    st.session_state.dl_filter_counts = {}
if "selected_dl_filter" not in st.session_state:
    st.session_state.selected_dl_filter = "All"

# 3. Processing Core Engine Routine Execution
if uploaded_file:
    if not template_exists:
        st.error("🛑 Processing halted: The requested template layout parameters could not be validated because the file is missing from local path layout directories.")
    else:
        try:
            # --- Read the file ---
            if uploaded_file.name.endswith('.csv'):
                try:
                    df_source = pd.read_csv(uploaded_file)
                except UnicodeDecodeError:
                    uploaded_file.seek(0)
                    df_source = pd.read_csv(uploaded_file, encoding='latin1')
            else:
                sheet_names = pd.ExcelFile(uploaded_file).sheet_names
                uploaded_file.seek(0)
                if "SUMMARY" in sheet_names:
                    st.info("📊 Found 'SUMMARY' sheet in the uploaded file. Using SUMMARY sheet for processing.")
                    df_source = pd.read_excel(uploaded_file, sheet_name="SUMMARY")
                    header_row_index = 0
                else:
                    df_temp = pd.read_excel(uploaded_file, header=None, nrows=15)
                    header_row_index = 0
                    for i, row in df_temp.iterrows():
                        row_values = [str(val).strip().upper() for val in row.values]
                        if any(key in row_values for key in ["ACCOUNT NUMBER", "OB/PRINCIPAL", "PLACEMENT", "CH NAME", "CH CODE"]):
                            header_row_index = i
                            break
                    uploaded_file.seek(0)
                    df_source = pd.read_excel(uploaded_file, header=header_row_index)
                    st.info(f"**🛠️ Auto-Header Scanner:** Successfully locked onto headers at Row {header_row_index + 1}:\n\n`{', '.join(df_source.columns.tolist())}`")
            
            df_source.columns = df_source.columns.astype(str).str.strip().str.upper()
            
            # Detect DL_TYPE column
            source_dl_col = None
            for possible in ["DL_TYPE", "DL TYPE", "DL-TYPE"]:
                if possible in df_source.columns:
                    source_dl_col = possible
                    break
            if source_dl_col is not None and source_dl_col != "DL_TYPE":
                df_source.rename(columns={source_dl_col: "DL_TYPE"}, inplace=True)
                source_dl_col = "DL_TYPE"

            # Store df_source and DL_TYPE info in session state for the filter placeholder
            st.session_state.df_source = df_source
            if "DL_TYPE" in df_source.columns and st.session_state.mode == "Demand Letter with Transmittal":
                unique_types = sorted(df_source["DL_TYPE"].dropna().unique())
                counts = df_source["DL_TYPE"].value_counts().to_dict()
                st.session_state.dl_filter_options = unique_types
                st.session_state.dl_filter_counts = counts
                if st.session_state.selected_dl_filter == "All" and len(unique_types) > 0:
                    st.session_state.selected_dl_filter = unique_types[0]
            else:
                st.session_state.dl_filter_options = []
                st.session_state.dl_filter_counts = {}

            # Update the placeholder in col3 (if in Demand Letter mode)
            if st.session_state.mode == "Demand Letter with Transmittal":
                filter_ph = st.session_state.get("filter_placeholder")
                if filter_ph is not None:
                    with filter_ph:
                        if "DL_TYPE" in df_source.columns and len(st.session_state.dl_filter_options) > 0:
                            summary_text = ", ".join([f"{k}: {v}" for k, v in st.session_state.dl_filter_counts.items()])
                            st.caption(f"📊 Counts: {summary_text}")
                            selected = st.selectbox(
                                "📑 Filter by DL_TYPE:",
                                options=["All"] + st.session_state.dl_filter_options,
                                index=0 if st.session_state.selected_dl_filter == "All" else (["All"] + st.session_state.dl_filter_options).index(st.session_state.selected_dl_filter),
                                help="Select a specific DL_TYPE to process only those rows."
                            )
                            st.session_state.selected_dl_filter = selected
                        else:
                            st.caption("ℹ️ No DL_TYPE column found in uploaded file.")
            
            # Apply DL_TYPE filter
            if st.session_state.mode == "Demand Letter with Transmittal" and "DL_TYPE" in df_source.columns:
                selected_filter = st.session_state.selected_dl_filter
                if selected_filter != "All":
                    df_source = df_source[df_source["DL_TYPE"] == selected_filter]
                    st.info(f"✅ Filtered to DL_TYPE = **{selected_filter}**. **{len(df_source)}** rows remaining.")
                else:
                    st.info(f"📊 Processing all DL_TYPE values. Total rows: **{len(df_source)}**")
            
            # Now proceed with mapping
            df_template_structure = pd.read_excel(template_filename, nrows=0)
            target_columns = [str(col).strip() for col in df_template_structure.columns.tolist()]

            if target_columns:
                mapping = {
                    "DF_2926": ["OB/PRINCIPAL"],
                    "DF_3179": ["OB/PRINCIPAL"],
                    "DF_5633": ["SOA DATE"],
                    "LEADS_OB": ["OB/PRINCIPAL"],
                    "DL_ADDRESS": ["ADDRESS"],
                    "LEADS_ACCTNO": ["ACCOUNT NUMBER"],
                    "LEADS_CHNAME": ["CH NAME"],
                    "LEADS_ENDO_DATE": ["ENDO DATE"],
                    "DF_1767": ["ACCOUNT TYPE"],
                    "DL_ADDRESS_TYPE_FULL": ["ADD TYPE"],
                    "LEADS_PLACEMENT": ["PLACEMENT"],
                    "FINAL_AREA": ["FINAL AREA"],
                    "AGENT_CODE": ["AGENT CODE", "AGENT_CODE", "LEADS_AGENTCODE"],
                    "AGENT_NAME": ["AGENT NAME", "AGENT_NAME", "LEADS_AGENTNAME", "LEADS_AGENT"],
                    "DF_5632": ["AMOUNT IN WORD", "AMOUNT IN WORDS"],
                }

                df_target = pd.DataFrame(index=df_source.index, columns=target_columns)

                for target_col in df_target.columns:
                    source_candidates = mapping.get(target_col)
                    if source_candidates:
                        for source_col in source_candidates:
                            if source_col in df_source.columns:
                                df_target[target_col] = df_source[source_col]
                                break
                    elif target_col == "LEADS_CHCODE" or target_col == "ACCT_TRANS_CODE":
                        if "CH CODE" in df_source.columns:
                            df_target[target_col] = df_source["CH CODE"]
                    elif target_col == "CLIENT_NAME":
                        df_target["CLIENT_NAME"] = selected_client_name
                    elif target_col == "DL_TYPE":
                        if st.session_state.mode == "Transmittal Only":
                            df_target["DL_TYPE"] = st.session_state.dl_type
                        else:
                            if "DL_TYPE" in df_source.columns:
                                df_target["DL_TYPE"] = df_source["DL_TYPE"]
                            else:
                                df_target["DL_TYPE"] = selected_template

                # Additional column-specific formatting (unchanged)
                if "DF_2926" in df_target.columns and "OB/PRINCIPAL" in df_source.columns:
                    df_target["DF_2926"] = df_source["OB/PRINCIPAL"]
                if "DF_3179" in df_target.columns and "OB/PRINCIPAL" in df_source.columns:
                    df_target["DF_3179"] = df_source["OB/PRINCIPAL"]
                if "LEADS_OB" in df_target.columns and "OB/PRINCIPAL" in df_source.columns:
                    df_target["LEADS_OB"] = df_source["OB/PRINCIPAL"]
                if "LEADS_ACCTNO" in df_target.columns and "ACCOUNT NUMBER" in df_source.columns:
                    if selected_client_name == "SBC HOME LOAN":
                        df_target["LEADS_ACCTNO"] = df_source["ACCOUNT NUMBER"]
                    else:
                        clean_acct = df_source["ACCOUNT NUMBER"].fillna("").astype(str).str.strip()
                        clean_acct = clean_acct.replace(r'\.0$', '', regex=True)
                        df_target["LEADS_ACCTNO"] = clean_acct
                if "ADDRESS_TYPE" in df_target.columns and "ADD TYPE" in df_source.columns:
                    clean_addr = df_source["ADD TYPE"].fillna("").astype(str).str.strip()
                    df_target["ADDRESS_TYPE"] = clean_addr.str.replace(r'(?i)\s*ADDRESS\s*', '', regex=True).str.strip()

                # Format numeric and date columns (same as before)
                if "DF_2926" in df_target.columns:
                    temp_num = pd.to_numeric(df_target["DF_2926"].astype(str).str.replace(',', '', regex=False), errors='coerce')
                    df_target["DF_2926"] = temp_num.apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
                if "DF_3179" in df_target.columns:
                    temp_num_3179 = pd.to_numeric(df_target["DF_3179"].astype(str).str.replace(',', '', regex=False), errors='coerce')
                    df_target["DF_3179"] = temp_num_3179.apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
                if "DF_5633" in df_target.columns:
                    df_target["DF_5633"] = pd.to_datetime(df_target["DF_5633"], errors='coerce').dt.strftime('%B %d, %Y').fillna("")
                if "LEADS_OB" in df_target.columns:
                    temp_num_ob = pd.to_numeric(df_target["LEADS_OB"].astype(str).str.replace(',', '', regex=False), errors='coerce')
                    df_target["LEADS_OB"] = temp_num_ob.apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
                if "LEADS_ENDO_DATE" in df_target.columns:
                    df_target["LEADS_ENDO_DATE"] = pd.to_datetime(df_target["LEADS_ENDO_DATE"], errors='coerce').dt.strftime('%B %d, %Y').fillna("")

                # =============================================================
                # LOOKUP LOGIC (Demand Letter only, with reference file)
                # =============================================================
                if st.session_state.mode == "Demand Letter with Transmittal" and required_lookup_name and lookup_exists:
                    # --- AGENT_NAME lookup via AGENT_CODE ---
                    if "AGENT_CODE" in df_source.columns:
                        try:
                            df_agent_lookup = pd.read_excel(lookup_filename, sheet_name="AGENT")
                            df_agent_lookup.columns = df_agent_lookup.columns.astype(str).str.strip().str.upper()
                            for col in df_agent_lookup.columns:
                                if df_agent_lookup[col].dtype == 'object':
                                    df_agent_lookup[col] = df_agent_lookup[col].astype(str).str.strip()
                            
                            temp_agent = df_source[["AGENT_CODE"]].copy()
                            temp_agent["AGENT_CODE"] = temp_agent["AGENT_CODE"].fillna("").astype(str).str.strip()
                            
                            merged_agent = pd.merge(temp_agent, df_agent_lookup, on="AGENT_CODE", how="left")
                            
                            if "AGENT_NAME" in df_target.columns and "AGENT_NAME" in merged_agent.columns:
                                df_target["AGENT_NAME"] = merged_agent["AGENT_NAME"]
                            if "CLIENT_NAME" in df_target.columns and "AGENT_NAME" in merged_agent.columns:
                                df_target["CLIENT_NAME"] = merged_agent["AGENT_NAME"]
                            if "AGENT_CODE" in df_target.columns and "AGENT_CODE" in merged_agent.columns:
                                df_target["AGENT_CODE"] = merged_agent["AGENT_CODE"]
                            
                            st.info("✅ AGENT_NAME successfully mapped from AGENT sheet using AGENT_CODE lookup.")
                        except ValueError:
                            st.warning("⚠️ AGENT sheet not found in reference file. Skipping AGENT_NAME lookup.")
                    
                    # --- PLACEMENT lookup for office details ---
                    if "PLACEMENT" in df_source.columns:
                        # Read the "ALL" sheet (first sheet) which contains the placement mapping
                        df_lookup = pd.read_excel(lookup_filename, sheet_name="ALL")
                        df_lookup.columns = df_lookup.columns.astype(str).str.strip().str.upper()
                        for col in df_lookup.columns:
                            if df_lookup[col].dtype == 'object':
                                df_lookup[col] = df_lookup[col].astype(str).str.strip()

                        # Prepare source placement values
                        temp_source = df_source[["PLACEMENT"]].copy()
                        temp_source["PLACEMENT"] = temp_source["PLACEMENT"].fillna("").astype(str).str.strip()

                        # Merge
                        merged_lookup = pd.merge(temp_source, df_lookup, on="PLACEMENT", how="left")

                        # Count matches
                        matched_count = merged_lookup["MAIN_OFFICE_ADDRESS"].notna().sum()
                        st.info(f"🔍 PLACEMENT lookup: {matched_count} out of {len(df_source)} rows matched.")

                        if matched_count == 0:
                            # Show sample values to help debug
                            source_vals = temp_source["PLACEMENT"].unique()[:5]
                            lookup_vals = df_lookup["PLACEMENT"].unique()[:5]
                            st.warning(f"⚠️ No PLACEMENT matches found. Check your source PLACEMENT values (e.g. {', '.join(source_vals)}) against lookup values (e.g. {', '.join(lookup_vals)}).")

                        # Map to target columns
                        placement_mappings = {
                            "MAIN_OFFICE_ADDRESS": "MAIN_OFFICE_ADDRESS",
                            "M_PHONE": "M_PHONE",
                            "M_TEL": "M_TEL",
                            "CLIENT_EMAIL": "CLIENT_EMAIL"
                        }
                        for target_col, lookup_col in placement_mappings.items():
                            if target_col in df_target.columns and lookup_col in merged_lookup.columns:
                                df_target[target_col] = merged_lookup[lookup_col]

                # Fill empty cells
                record_count = len(df_target)
                df_target = df_target.fillna("")

                st.subheader("📋 Preview Table")
                
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Processed Records", value=record_count)
                m2.metric(label="Destination Columns Mapped", value=len(df_target.columns))
                m3.metric(label="Output Format Extension", value=".xlsx (Excel Workbook)")
                
                st.write("")
                st.dataframe(df_target, use_container_width=True)

                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_target.to_excel(writer, index=False, sheet_name='Mapped_Data')
                excel_data = excel_buffer.getvalue()
                
                st.write("")
                custom_filename = f"{selected_client_name}_{selected_template}_{record_count}.xlsx"
                
                _, btn_col, _ = st.columns([1, 2, 1])
                with btn_col:
                    st.download_button(
                        label=f"⚡ Download {selected_client_name} {selected_template} Excel",
                        data=excel_data,
                        file_name=custom_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                
        except Exception as e:
            st.error(f"Execution Error Exception encountered during dataset processing: {e}")
