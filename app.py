# app.py
# -----------------------------------------
# 30/30 Daily Report Combiner (Streamlit)
# -----------------------------------------
# Run with: streamlit run app.py
# Suggested requirements (requirements.txt):
# streamlit
# pandas
# numpy
# XlsxWriter
# openpyxl

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="30/30 Report Combiner", layout="wide")

st.title("📊 30/30 Daily Report Combiner")
st.markdown("### Combine Carwars and Tecobi reports with 30/30 validation")

# Define excluded agents at the top level
EXCLUDED_AGENTS = ['AJ Dhir', 'Aj Dhir', 'Thomas Williams', 'Mark Moore', 'Nicole Farr']

# Initialize session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
    st.session_state.result_buffer = None

def parse_time_to_excel(time_str):
    """Convert MM:SS or H:MM:SS format to Excel time (fraction of day)"""
    if pd.isna(time_str) or time_str == '' or time_str == '0:00' or time_str == 0:
        return 0
    if isinstance(time_str, (int, float)):
        return float(time_str)
    time_str = str(time_str).strip()
    if ':' not in time_str:
        return 0
    try:
        parts = time_str.split(':')
        if len(parts) == 2:  # MM:SS
            minutes = int(parts[0]); seconds = int(parts[1])
            total_hours = (minutes * 60 + seconds) / 3600
        elif len(parts) == 3:  # H:MM:SS
            hours = int(parts[0]); minutes = int(parts[1]); seconds = int(parts[2])
            total_hours = hours + minutes/60 + seconds/3600
        else:
            return 0
        return total_hours / 24  # Excel fraction of day
    except:
        return 0

def get_first_name(full_name):
    """Extract first name for sorting"""
    if pd.isna(full_name):
        return ""
    s = str(full_name).strip()
    return s.split()[0] if s else ""

# NOTE: renamed to avoid collisions with older cached/imported versions
def process_carwars_file_v2(df, location, exclude_list=None):
    """Process Carwars file and extract needed columns"""
    df.columns = df.columns.str.strip()

    # Filter out excluded agents (case-insensitive, partial match)
    if exclude_list and 'Agent Name' in df.columns:
        for name in exclude_list:
            df = df[~df['Agent Name'].astype(str).str.lower().str.contains(name.lower(), na=False)]

    processed = pd.DataFrame()
    processed['Agent Name'] = df.get('Agent Name', '').astype(str).str.strip()

    processed['Carwars_Unique_Outbound'] = pd.to_numeric(df.get('Unique Outbound', 0), errors='coerce').fillna(0)

    avg_talk_col = df.get('Avg Talk Time', None)
    if avg_talk_col is not None:
        processed['Carwars_Avg_Talk_Time'] = avg_talk_col.apply(parse_time_to_excel)
    else:
        processed['Carwars_Avg_Talk_Time'] = 0

    # Use Unique OB Text if available, otherwise Total OB Text
    if 'Unique OB Text' in df.columns:
        processed['Carwars_OB_Text'] = pd.to_numeric(df['Unique OB Text'], errors='coerce').fillna(0)
    else:
        processed['Carwars_OB_Text'] = pd.to_numeric(df.get('Total OB Text', 0), errors='coerce').fillna(0)

    processed['Location'] = location
    return processed

def process_tecobi_file(df, location, exclude_list=None):
    """Process Tecobi file and extract needed columns"""
    df.columns = df.columns.str.strip()

    # Build 'Agent Name' if first/last present; else try to use an existing name field
    if 'first_name' in df.columns and 'last_name' in df.columns:
        df['Agent Name'] = (df['first_name'].astype(str).str.strip() + ' ' + df['last_name'].astype(str).str.strip()).str.strip()
    elif 'Agent Name' not in df.columns:
        # fallback – try 'name'
        if 'name' in df.columns:
            df['Agent Name'] = df['name'].astype(str).str.strip()
        else:
            df['Agent Name'] = ""

    # Filter out excluded agents
    if exclude_list:
        for name in exclude_list:
            df = df[~df['Agent Name'].astype(str).str.lower().str.contains(name.lower(), na=False)]

    # Tecobi talk time
    if 'avg_outbound_call_duration' in df.columns:
        talk_time = pd.to_numeric(df['avg_outbound_call_duration'], errors='coerce').fillna(0)
    else:
        seconds = pd.to_numeric(df.get('seconds_clocked_in', 0), errors='coerce').fillna(0)
        calls = pd.to_numeric(df.get('outbound_calls', 0), errors='coerce').fillna(1).replace(0, 1)
        talk_time = seconds / calls

    processed = pd.DataFrame()
    processed['Agent Name'] = df['Agent Name'].astype(str).str.strip()
    processed['Tecobi_Outbound_Calls'] = pd.to_numeric(df.get('outbound_calls', 0), errors='coerce').fillna(0)
    processed['Tecobi_Talk_Time'] = talk_time
    processed['Tecobi_External_SMS'] = pd.to_numeric(df.get('external_sms', 0), errors='coerce').fillna(0)
    processed['Location'] = location
    return processed

def combine_location_data(carwars_df, tecobi_df, location):
    """Combine Carwars and Tecobi data for a single location"""
    carwars_loc = carwars_df[carwars_df['Location'] == location].copy()
    tecobi_loc = tecobi_df[tecobi_df['Location'] == location].copy()

    # Remove Location (not needed after filtering)
    carwars_loc = carwars_loc.drop(columns=['Location'], errors='ignore')
    tecobi_loc = tecobi_loc.drop(columns=['Location'], errors='ignore')

    combined = pd.merge(
        carwars_loc,
        tecobi_loc,
        on='Agent Name',
        how='outer'
    )

    numeric_columns = [
        'Carwars_Unique_Outbound', 'Carwars_Avg_Talk_Time', 'Carwars_OB_Text',
        'Tecobi_Outbound_Calls', 'Tecobi_Talk_Time', 'Tecobi_External_SMS'
    ]
    for col in numeric_columns:
        if col in combined.columns:
            combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0)

    combined['Calls'] = combined.get('Carwars_Unique_Outbound', 0) + combined.get('Tecobi_Outbound_Calls', 0)
    combined['Text'] = combined.get('Carwars_OB_Text', 0) + combined.get('Tecobi_External_SMS', 0)

    combined['Needs_Highlight'] = (combined['Calls'] < 30) | (combined['Text'] < 30)

    final = pd.DataFrame({
        'Agent Name': combined['Agent Name'],
        'Calls': combined['Calls'].astype(int),
        'Carwars Avg Talk Time': combined.get('Carwars_Avg_Talk_Time', 0),
        'Tecobi Talk Time': combined.get('Tecobi_Talk_Time', 0),
        'Text': combined['Text'].astype(int),
        'Needs_Highlight': combined['Needs_Highlight']
    })

    # Sort by first name
    final['First_Name'] = final['Agent Name'].apply(get_first_name)
    final = final.sort_values('First_Name', na_position='last').drop(columns=['First_Name'])
    return final

def create_formatted_excel(chattanooga_data, cleveland_data, dalton_data):
    """Create the final formatted Excel file matching 30/30 format"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter',
            'align': 'center', 'border': 1, 'bg_color': '#D7D7D7', 'font_size': 10
        })
        location_header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'font_size': 11, 'bg_color': '#B4C6E7', 'border': 1
        })
        time_format = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1})
        time_format_highlight = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1, 'bg_color': '#FFC7CE'})
        number_format = workbook.add_format({'align': 'center', 'border': 1})
        number_format_highlight = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFC7CE'})
        text_format = workbook.add_format({'align': 'left', 'border': 1})
        text_format_highlight = workbook.add_format({'align': 'left', 'border': 1, 'bg_color': '#FFC7CE'})
        empty_format = workbook.add_format({'border': 1})

        worksheet = writer.book.add_worksheet('Sheet1')

        # Column widths
        worksheet.set_column('A:A', 18)
        worksheet.set_column('B:B', 8)
        worksheet.set_column('C:C', 12)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:E', 8)

        worksheet.set_column('F:F', 18)
        worksheet.set_column('G:G', 8)
        worksheet.set_column('H:H', 12)
        worksheet.set_column('I:I', 12)
        worksheet.set_column('J:J', 8)

        worksheet.set_column('K:K', 18)
        worksheet.set_column('L:L', 8)
        worksheet.set_column('M:M', 12)
        worksheet.set_column('N:N', 12)
        worksheet.set_column('O:O', 8)

        # Visual separators (optional)
        worksheet.set_column('P:P', 2)
        worksheet.set_column('Q:Q', 2)

        # Headers
        worksheet.merge_range('A1:E1', 'Chattanooga', location_header_format)
        worksheet.merge_range('F1:J1', 'Cleveland', location_header_format)
        worksheet.merge_range('K1:O1', 'Dalton', location_header_format)

        headers = ['Agent Name', 'Calls', 'Carwars Avg\nTalk Time', 'Tecobi\nTalk Time\n(seconds)', 'Text']
        for i, h in enumerate(headers):
            worksheet.write(1, i, h, header_format)
            worksheet.write(1, 5 + i, h, header_format)
            worksheet.write(1, 10 + i, h, header_format)

        # Rows
        max_rows = max(len(chattanooga_data), len(cleveland_data), len(dalton_data))
        for row_idx in range(max_rows):
            excel_row = row_idx + 2

            # Helper to write a block (5 columns)
            def write_block(base_col, rowdata):
                needs = rowdata['Needs_Highlight']
                name_fmt = text_format_highlight if needs else text_format
                num_fmt = number_format_highlight if needs else number_format
                t_fmt = time_format_highlight if needs else time_format
                worksheet.write(excel_row, base_col + 0, rowdata['Agent Name'], name_fmt)
                worksheet.write(excel_row, base_col + 1, rowdata['Calls'], num_fmt)
                worksheet.write(excel_row, base_col + 2, rowdata['Carwars Avg Talk Time'], t_fmt)
                worksheet.write(excel_row, base_col + 3, rowdata['Tecobi Talk Time'], num_fmt)
                worksheet.write(excel_row, base_col + 4, rowdata['Text'], num_fmt)

            # Chattanooga
            if row_idx < len(chattanooga_data):
                write_block(0, chattanooga_data.iloc[row_idx])
            else:
                for c in range(0, 5):
                    worksheet.write(excel_row, c, '', empty_format)

            # Cleveland
            if row_idx < len(cleveland_data):
                write_block(5, cleveland_data.iloc[row_idx])
            else:
                for c in range(5, 10):
                    worksheet.write(excel_row, c, '', empty_format)

            # Dalton
            if row_idx < len(dalton_data):
                write_block(10, dalton_data.iloc[row_idx])
            else:
                for c in range(10, 15):
                    worksheet.write(excel_row, c, '', empty_format)

        worksheet.freeze_panes(2, 0)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)

    output.seek(0)
    return output

# -----------------------------------------
# UI
# -----------------------------------------
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Upload Files")
    st.markdown("Upload all 6 files (2 per location)")

    # File uploaders
    st.markdown("**Chattanooga Files:**")
    chatt_carwars = st.file_uploader("Chattanooga Carwars", type=['xlsx', 'xls', 'csv'], key="chatt_carwars")
    chatt_tecobi = st.file_uploader("Chattanooga Tecobi", type=['xlsx', 'xls', 'csv'], key="chatt_tecobi")

    st.markdown("**Cleveland Files:**")
    cleve_carwars = st.file_uploader("Cleveland Carwars", type=['xlsx', 'xls', 'csv'], key="cleve_carwars")
    cleve_tecobi = st.file_uploader("Cleveland Tecobi", type=['xlsx', 'xls', 'csv'], key="cleve_tecobi")

    st.markdown("**Dalton Files:**")
    dalton_carwars = st.file_uploader("Dalton Carwars", type=['xlsx', 'xls', 'csv'], key="dalton_carwars")
    dalton_tecobi = st.file_uploader("Dalton Tecobi", type=['xlsx', 'xls', 'csv'], key="dalton_tecobi")

with col2:
    st.subheader("⚙️ Process Files")
    st.info("ℹ️ The following agents will be automatically excluded: AJ Dhir, Thomas Williams, Mark Moore, Nicole Farr")

    all_files_uploaded = all([
        chatt_carwars, chatt_tecobi,
        cleve_carwars, cleve_tecobi,
        dalton_carwars, dalton_tecobi
    ])

    if all_files_uploaded:
        st.success("✅ All files uploaded!")

        if st.button("🔄 Process and Generate 30/30 Report", type="primary", use_container_width=True):
            try:
                with st.spinner("Processing files..."):
                    def read_file(uploaded_file):
                        name = uploaded_file.name.lower()
                        if name.endswith('.csv'):
                            return pd.read_csv(uploaded_file)
                        # openpyxl engine will be used automatically if installed
                        return pd.read_excel(uploaded_file)

                    # Carwars
                    carwars_files = {
                        'Chattanooga': process_carwars_file_v2(read_file(chatt_carwars), 'Chattanooga', exclude_list=EXCLUDED_AGENTS),
                        'Cleveland':   process_carwars_file_v2(read_file(cleve_carwars), 'Cleveland',   exclude_list=EXCLUDED_AGENTS),
                        'Dalton':      process_carwars_file_v2(read_file(dalton_carwars), 'Dalton',    exclude_list=EXCLUDED_AGENTS),
                    }

                    # Tecobi
                    tecobi_files = {
                        'Chattanooga': process_tecobi_file(read_file(chatt_tecobi), 'Chattanooga', exclude_list=EXCLUDED_AGENTS),
                        'Cleveland':   process_tecobi_file(read_file(cleve_tecobi), 'Cleveland',   exclude_list=EXCLUDED_AGENTS),
                        'Dalton':      process_tecobi_file(read_file(dalton_tecobi), 'Dalton',     exclude_list=EXCLUDED_AGENTS),
                    }

                    all_carwars = pd.concat(carwars_files.values(), ignore_index=True)
                    all_tecobi = pd.concat(tecobi_files.values(), ignore_index=True)

                    chattanooga_final = combine_location_data(all_carwars, all_tecobi, 'Chattanooga')
                    cleveland_final   = combine_location_data(all_carwars, all_tecobi, 'Cleveland')
                    dalton_final      = combine_location_data(all_carwars, all_tecobi, 'Dalton')

                    # Summary
                    st.markdown("### 📊 Summary")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric("Chattanooga", f"{len(chattanooga_final)} agents",
                                  f"{int(chattanooga_final['Needs_Highlight'].sum())} below 30/30")
                    with c2:
                        st.metric("Cleveland", f"{len(cleveland_final)} agents",
                                  f"{int(cleveland_final['Needs_Highlight'].sum())} below 30/30")
                    with c3:
                        st.metric("Dalton", f"{len(dalton_final)} agents",
                                  f"{int(dalton_final['Needs_Highlight'].sum())} below 30/30")

                    st.session_state.result_buffer = create_formatted_excel(
                        chattanooga_final, cleveland_final, dalton_final
                    )
                    st.session_state.processed = True

                st.success("✅ 30/30 Report generated successfully!")

            except Exception as e:
                st.error(f"❌ Error processing files: {e}")
                st.markdown("**Debug Information:**")
                st.code(str(e))
    else:
        st.warning("⚠️ Please upload all 6 files to continue")
        missing = []
        if not chatt_carwars:  missing.append("Chattanooga Carwars")
        if not chatt_tecobi:   missing.append("Chattanooga Tecobi")
        if not cleve_carwars:  missing.append("Cleveland Carwars")
        if not cleve_tecobi:   missing.append("Cleveland Tecobi")
        if not dalton_carwars: missing.append("Dalton Carwars")
        if not dalton_tecobi:  missing.append("Dalton Tecobi")
        if missing:
            st.markdown("**Missing files:**")
            for m in missing:
                st.markdown(f"- {m}")

# Download section
if st.session_state.processed and st.session_state.result_buffer:
    st.markdown("---")
    st.subheader("📥 Download Result")
    current_date = datetime.now().strftime("%m_%d_%Y")
    filename = f"30_30_Report_{current_date}_Formatted.xlsx"

    st.download_button(
        label="⬇️ Download 30/30 Report",
        data=st.session_state.result_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("🔄 Process New Files", use_container_width=True):
        st.session_state.processed = False
        st.session_state.result_buffer = None
        st.rerun()

# Instructions
st.markdown("---")
with st.expander("📖 Instructions & Info"):
    st.markdown("""
    ### How to Use:
    1. **Upload all 6 files** - Carwars and Tecobi files for each location
    2. **Click Process** - The app will combine and validate the data
    3. **Download** - Get your formatted Excel report

    ### 30/30 Validation:
    - Agents with **less than 30 calls OR less than 30 texts** are highlighted in light red
    - This helps quickly identify who hasn't met the 30/30 standard

    ### Data Processing:
    - **Calls** = Carwars "Unique Outbound" + Tecobi "outbound_calls"
    - **Text** = Carwars "Unique OB Text" + Tecobi "external_sms"
    - **Talk Times** are kept separate for each system
    - Agents are sorted alphabetically by first name
    - Handles agents appearing in only one system
    - Automatically excludes: AJ Dhir, Thomas Williams, Mark Moore, Nicole Farr
    """)
