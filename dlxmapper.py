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

# App Header Styling with Reset Button
col_header, col_reset = st.columns([4, 1])
with col_header:
    st.title("🎯 DLX Mapper")
with col_reset:
    st.write("") 
    # FIXED: Replaced width="stretch" with use_container_width=True
    if st.button("🔄 Reset App", use_container_width=True, type="secondary"):
        st.session_state.uploader_key += 1
        st.rerun()

st.write("---")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. STEP 1: Configuration Control Panel (Moved to top)
st.subheader("⚙️ Select Section")

with st.container(border=True):
    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        client_name_options = [
            "PIF HOME LOAN",
            "PIF FORECLOSURE",
            "SBC HOME LOAN",
            "SBF HOME LOAN",
            "CBS HOUSING LOAN",
            "BDO HOME LOAN COLLECT",
            "UBP HOMELOAN MORTGAGE"
        ]
        selected_client_name = st.selectbox(
            "👤 Select Campaign:", 
            client_name_options,
            help="Determines the target folder directory layout structure to pull the template configuration from."
        )
        
    with col2:
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

# =========================================================
# DYNAMIC FILE PATH ROUTING & DIAGNOSTICS
# =========================================================

template_filename = os.path.join(SCRIPT_DIR, selected_client_name, f"{selected_template}.xlsx")
template_exists = os.path.exists(template_filename)

required_lookup_name = None
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
            st.success(f"📁 Template Found: `{selected_client_name}/{selected_template}.xlsx`")
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
            st.info(f"ℹ️ No Reference Database required for `{selected_client_name}`.")

st.write("---")

# 2. STEP 2: Main Source File Upload (Moved below selection)
st.subheader("📥 Upload Source File")
uploaded_file = st.file_uploader(
    label="Drag and drop your source file configuration (Excel or CSV formats supported)", 
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
    key=f"file_uploader_{st.session_state.uploader_key}" 
)

# 3. Processing Core Engine Routine Execution
if uploaded_file:
    if not template_exists:
        st.error("🛑 Processing halted: The requested template layout parameters could not be validated because the file is missing from local path layout directories.")
    else:
        try:
            if uploaded_file.name.endswith('.csv'):
                try:
                    df_source = pd.read_csv(uploaded_file)
                except UnicodeDecodeError:
                    uploaded_file.seek(0)
                    df_source = pd.read_csv(uploaded_file, encoding='latin1')
            else:
                # Check if the Excel file has a "SUMMARY" sheet
                sheet_names = pd.ExcelFile(uploaded_file).sheet_names
                uploaded_file.seek(0)
                
                # If "SUMMARY" sheet exists, use it; otherwise use the default logic
                if "SUMMARY" in sheet_names:
                    st.info("📊 Found 'SUMMARY' sheet in the uploaded file. Using SUMMARY sheet for processing.")
                    df_source = pd.read_excel(uploaded_file, sheet_name="SUMMARY")
                    header_row_index = 0
                else:
                    # Fallback to original behavior: auto-detect header row
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
            
            df_template_structure = pd.read_excel(template_filename, nrows=0)
            target_columns = [str(col).strip() for col in df_template_structure.columns.tolist()]

            if target_columns:
                
                # =================================================================
                # UPGRADED OMNI-MAPPING DICTIONARY
                # Automatically maps all known variations of Agent names and codes
                # =================================================================
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

                    # Agent Code Variations
                    "AGENT_CODE": ["AGENT CODE", "AGENT_CODE", "LEADS_AGENTCODE"],

                    # Agent Name Variations
                    "AGENT_NAME": ["AGENT NAME", "AGENT_NAME", "LEADS_AGENTNAME", "LEADS_AGENT"],

                    # PIF FORECLOSURE DL9
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
                        df_target["DL_TYPE"] = selected_template

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

                if required_lookup_name and lookup_exists and "AGENT_CODE" in df_source.columns:
                    # Read AGENT sheet for AGENT_NAME lookup based on AGENT_CODE
                    try:
                        df_agent_lookup = pd.read_excel(lookup_filename, sheet_name="AGENT")
                        df_agent_lookup.columns = df_agent_lookup.columns.astype(str).str.strip().str.upper()
                        for col in df_agent_lookup.columns:
                            if df_agent_lookup[col].dtype == 'object':
                                df_agent_lookup[col] = df_agent_lookup[col].astype(str).str.strip()
                        
                        # Merge source AGENT_CODE with AGENT sheet to get AGENT_NAME
                        temp_agent = df_source[["AGENT_CODE"]].copy()
                        temp_agent["AGENT_CODE"] = temp_agent["AGENT_CODE"].fillna("").astype(str).str.strip()
                        
                        merged_agent = pd.merge(temp_agent, df_agent_lookup, on="AGENT_CODE", how="left")
                        
                        # Map AGENT_NAME and AGENT_CODE from the lookup
                        if "AGENT_NAME" in df_target.columns and "AGENT_NAME" in merged_agent.columns:
                            df_target["AGENT_NAME"] = merged_agent["AGENT_NAME"]
                        
                        if "CLIENT_NAME" in df_target.columns and "AGENT_NAME" in merged_agent.columns:
                            df_target["CLIENT_NAME"] = merged_agent["AGENT_NAME"]
                        
                        if "AGENT_CODE" in df_target.columns and "AGENT_CODE" in merged_agent.columns:
                            df_target["AGENT_CODE"] = merged_agent["AGENT_CODE"]
                        
                        st.info("✅ AGENT_NAME successfully mapped from AGENT sheet using AGENT_CODE lookup.")
                    except ValueError:
                        st.warning("⚠️ AGENT sheet not found in reference file. Skipping AGENT_NAME lookup.")
                
                # Continue with PLACEMENT lookup for other fields (MAIN_OFFICE_ADDRESS, M_PHONE, M_TEL, CLIENT_EMAIL)
                if required_lookup_name and lookup_exists and "PLACEMENT" in df_source.columns:
                    df_lookup = pd.read_excel(lookup_filename)
                    df_lookup.columns = df_lookup.columns.astype(str).str.strip().str.upper() 
                    for col in df_lookup.columns:
                        if df_lookup[col].dtype == 'object':
                            df_lookup[col] = df_lookup[col].astype(str).str.strip()

                    temp_source = df_source[["PLACEMENT"]].copy()
                    temp_source["PLACEMENT"] = temp_source["PLACEMENT"].fillna("").astype(str).str.strip()

                    merged_lookup = pd.merge(temp_source, df_lookup, on="PLACEMENT", how="left")

                    placement_mappings = {
                        "MAIN_OFFICE_ADDRESS": "MAIN_OFFICE_ADDRESS",
                        "M_PHONE": "M_PHONE",
                        "M_TEL": "M_TEL",
                        "CLIENT_EMAIL": "CLIENT_EMAIL"
                    }

                    for target_col, lookup_col in placement_mappings.items():
                        if target_col in df_target.columns and lookup_col in merged_lookup.columns:
                            valid_series = merged_lookup[lookup_col]
                            df_target[target_col] = valid_series

                record_count = len(df_target)
                df_target = df_target.fillna("")

                st.subheader("📋 Preview Table")
                
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Processed Records", value=record_count)
                m2.metric(label="Destination Columns Mapped", value=len(df_target.columns))
                m3.metric(label="Output Format Extension", value=".xlsx (Excel Workbook)")
                
                st.write("")
                
                # FIXED: Replaced width="stretch" with use_container_width=True to prevent a secondary crash
                st.dataframe(df_target, use_container_width=True)

                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_target.to_excel(writer, index=False, sheet_name='Mapped_Data')
                excel_data = excel_buffer.getvalue()
                
                st.write("")
                
                custom_filename = f"{selected_client_name}_{selected_template}_{record_count}.xlsx"
                
                _, btn_col, _ = st.columns([1, 2, 1])
                with btn_col:
                    # FIXED: Replaced width="stretch" with use_container_width=True
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