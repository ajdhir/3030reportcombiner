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
    s = str(time_str).strip()
    if ':' not in s:
        return 0
    try:
        parts = s.split(':')
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

    # Trim and normalize Agent Name if present
    if 'Agent Name' in df.columns:
        df['Agent Name'] = df['Agent Name'].astype(str).str.strip()
    else:
        df['Agent Name'] = ""

    # Filter out "total" and "unassigned" rows (case-insensitive)
    df = df[~df['Agent Name'].str.contains(r'\btotal\b', case=False, na=False)]
    df = df[~df['Agent Name'].str.contains(r'\bunassigned\b', case=False, na=False)]

    # Filter out excluded agents (case-insensitive, partial match)
    if exclude_list:
        for name in exclude_list:
            df = df[~df['Agent Name'].str.lower().str.contains(name.lower(), na=False)]

    processed = pd.DataFrame()
    processed['Agent Name'] = df['Agent Name']

    processed['Carwars_Unique_Outbound'] = pd.to_numeric(df.get('Unique Outbound', 0), errors='coerce').fillna(0)

    if 'Avg Talk Time' in df.columns:
        processed['Carwars_Avg_Talk_Time'] = df['Avg Talk Time'].apply(parse_time_to_excel)
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
    elif 'Agent Name' in df.columns:
        df['Agent Name'] = df['Agent Name'].astype(str).str.strip()
    elif 'name' in df.columns:
        df['Agent Name'] = df['name'].astype(str).str.strip()
    else:
        df['Agent Name'] = ""

    # Filter out "total" and "unassigned" rows (case-insensitive)
    df = df[~df['Agent Name'].str.contains(r'\btotal\b', case=False, na=False)]
    df = df[~df['Agent Name'].str.contains(r'\bunassigned\b', case=False, na=False)]

    # Filter out excluded agents
    if exclude_list:
        for name in exclude_list:
            df = df[~df['Agent Name'].str.lower().str.contains(name.lower(), na=False)]

    # Tecobi talk time
    if 'avg_outbound_call_duration' in df.columns:
        talk_time = pd.to_numeric(df['avg_outbound_call_duration'], errors='coerce').fillna(0)
    else:
        seconds = pd.to_numeric(df.get('seconds_clocked_in', 0), errors='coerce').fillna(0)
        calls = pd.to_numeric(df.get('outbound_calls', 0), errors='coerce').fillna(1).replace(0, 1)
        talk_time = seconds / calls

    processed = pd.DataFrame()
    processed['Agent Name'] = df['Agent Name']
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

    # Final table
    final = pd.DataFrame({
        'Agent Name': combined['Agent Name'],
        'Calls': combined['Calls'].astype(int),
        'Carwars Avg Talk Time': combined.get('Carwars_Avg_Talk_Time', 0),
        'Tecobi Talk Time': combined.get('Tecobi_Talk_Time', 0),
        'Text': combined['Text'].astype(int)
    })

    # Sort by first name
    final['First_Name'] = final['Agent Name'].apply(get_first_name)
    final = final.sort_values('First_Name', na_position='last').drop(columns=['First_Name'])

    # Boolean flags for highlighting logic (used only during write)
    final['Name_Highlight'] = (final['Calls'] < 30) | (final['Text'] < 30)
    final['Calls_Highlight'] = (final['Calls'] < 30)
    final['Text_Highlight'] = (final['Text'] < 30)
    return final

def create_formatted_excel(chattanooga_data, cleveland_data, dalton_data):
    """Create the final formatted Excel file with thick RIGHT borders, totals, and a summary block in Q/R/S"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Base formats
        header_base = {'bold': True, 'text_wrap': True, 'valign': 'vcenter',
                       'align': 'center', 'border': 1, 'bg_color': '#D7D7D7', 'font_size': 10}
        header_format = workbook.add_format(header_base)
        header_format_thickright = workbook.add_format({**header_base, 'right': 2})

        location_header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'font_size': 11, 'bg_color': '#B4C6E7', 'border': 1
        })

        # Time formats
        time_format = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1})
        time_format_highlight = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1, 'bg_color': '#FFC7CE'})
        time_format_thickright = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1, 'right': 2})
        time_format_highlight_thickright = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'center', 'border': 1, 'bg_color': '#FFC7CE', 'right': 2})

        # Number formats
        number_format = workbook.add_format({'align': 'center', 'border': 1})
        number_format_highlight = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFC7CE'})
        number_format_thickright = workbook.add_format({'align': 'center', 'border': 1, 'right': 2})
        number_format_highlight_thickright = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFC7CE', 'right': 2})

        # Text formats
        text_format = workbook.add_format({'align': 'left', 'border': 1})
        text_format_highlight = workbook.add_format({'align': 'left', 'border': 1, 'bg_color': '#FFC7CE'})
        text_format_thickright = workbook.add_format({'align': 'left', 'border': 1, 'right': 2})
        text_format_highlight_thickright = workbook.add_format({'align': 'left', 'border': 1, 'bg_color': '#FFC7CE', 'right': 2})

        # Empty format
        empty_format = workbook.add_format({'border': 1})
        empty_format_thickright = workbook.add_format({'border': 1, 'right': 2})

        # Totals format
        total_number_format = workbook.add_format({'align': 'center', 'border': 1, 'bold': True})
        total_number_format_thickright = workbook.add_format({'align': 'center', 'border': 1, 'bold': True, 'right': 2})
        total_label_format = workbook.add_format({'align': 'left', 'border': 1, 'bold': True})

        # Summary block formats
        summary_header_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#E2EAFB'})
        summary_label_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'left'})
        summary_number_fmt = workbook.add_format({'border': 1, 'align': 'center'})

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

        # Spacer
        worksheet.set_column('P:P', 2)
        # Summary area columns (Q,R,S)
        worksheet.set_column('Q:Q', 14)
        worksheet.set_column('R:R', 10)
        worksheet.set_column('S:S', 10)

        # Headers
        worksheet.merge_range('A1:E1', 'Chattanooga', location_header_format)
        worksheet.merge_range('F1:J1', 'Cleveland', location_header_format)
        worksheet.merge_range('K1:O1', 'Dalton', location_header_format)

        headers = ['Agent Name', 'Calls', 'Carwars Avg\nTalk Time', 'Tecobi\nTalk Time\n(seconds)', 'Text']

        # Chattanooga header row (thick RIGHT on column E)
        for i, h in enumerate(headers):
            fmt = header_format_thickright if i == 4 else header_format
            worksheet.write(1, i, h, fmt)

        # Cleveland header row (thick RIGHT on column J)
        for i, h in enumerate(headers):
            base_col = 5 + i
            fmt = header_format_thickright if base_col == 9 else header_format
            worksheet.write(1, base_col, h, fmt)

        # Dalton header row (thick RIGHT on column O)
        for i, h in enumerate(headers):
            base_col = 10 + i
            fmt = header_format_thickright if base_col == 14 else header_format
            worksheet.write(1, base_col, h, fmt)

        # Rows
        max_rows = max(len(chattanooga_data), len(cleveland_data), len(dalton_data))

        for row_idx in range(max_rows):
            excel_row = row_idx + 2  # data starts on Excel row 3 (0-indexed -> 2)

            def write_block(base_col, rowdata, thickright_cols=None):
                """Write one 5-col block; thickright_cols are absolute excel columns needing right=2."""
                thickright_cols = thickright_cols or set()

                # Name format
                name_fmt = text_format_highlight if rowdata['Name_Highlight'] else text_format

                # Calls format (highlight only if Calls < 30)
                calls_fmt = number_format_highlight if rowdata['Calls_Highlight'] else number_format

                # Time formats
                carwars_time_fmt = time_format
                tecobi_time_fmt = number_format  # seconds as plain number

                # Text format (highlight only if Text < 30)
                text_num_fmt = number_format_highlight if rowdata['Text_Highlight'] else number_format

                # Apply thick RIGHT on the Text column for E, J, O
                abs_text_col = base_col + 4
                if abs_text_col in thickright_cols:
                    text_num_fmt = number_format_highlight_thickright if rowdata['Text_Highlight'] else number_format_thickright

                worksheet.write(excel_row, base_col + 0, rowdata['Agent Name'], name_fmt)
                worksheet.write(excel_row, base_col + 1, int(rowdata['Calls']), calls_fmt)
                worksheet.write(excel_row, base_col + 2, rowdata['Carwars Avg Talk Time'], carwars_time_fmt)
                worksheet.write(excel_row, base_col + 3, rowdata['Tecobi Talk Time'], tecobi_time_fmt)
                worksheet.write(excel_row, base_col + 4, int(rowdata['Text']), text_num_fmt)

            # Chattanooga (block base_col=0); thick RIGHT on column E -> absolute col 4
            if row_idx < len(chattanooga_data):
                write_block(0, chattanooga_data.iloc[row_idx], thickright_cols={4})
            else:
                for c in range(0, 5):
                    fmt = empty_format_thickright if c == 4 else empty_format  # thick-right on E
                    worksheet.write(excel_row, c, '', fmt)

            # Cleveland (block base_col=5); thick RIGHT on column J -> absolute col 9
            if row_idx < len(cleveland_data):
                write_block(5, cleveland_data.iloc[row_idx], thickright_cols={9})
            else:
                for c in range(5, 10):
                    fmt = empty_format_thickright if c == 9 else empty_format  # thick-right on J
                    worksheet.write(excel_row, c, '', fmt)

            # Dalton (block base_col=10); thick RIGHT on column O -> absolute col 14
            if row_idx < len(dalton_data):
                write_block(10, dalton_data.iloc[row_idx], thickright_cols={14})
            else:
                for c in range(10, 15):
                    fmt = empty_format_thickright if c == 14 else empty_format  # thick-right on O
                    worksheet.write(excel_row, c, '', fmt)

        # Totals row (after last data row)
        last_data_row = (max_rows - 1) + 2 if max_rows > 0 else 1  # last data row index
        totals_row = last_data_row + 1

        # Optional labels under Agent Name columns
        worksheet.write(totals_row, 0, "Totals", total_label_format)
        worksheet.write(totals_row, 5, "Totals", total_label_format)
        worksheet.write(totals_row, 10, "Totals", total_label_format)

        # Helper to write a SUM in a column (col_letter, start_row=3 to end_row=last_data_row+1 in Excel terms)
        def write_sum(col_idx, thick_right=False):
            col_letter = xlsx_col_letter(col_idx)
            start_row_excel = 3
            end_row_excel = last_data_row + 1  # convert 0-indexed to 1-indexed
            formula = f"=SUM({col_letter}{start_row_excel}:{col_letter}{end_row_excel})"
            fmt = total_number_format_thickright if thick_right else total_number_format
            worksheet.write_formula(totals_row, col_idx, formula, fmt)

        # Write sums for requested columns: B, E, G, J, L, O
        write_sum(1, thick_right=False)   # B
        write_sum(4, thick_right=True)    # E (thick RIGHT)
        write_sum(6, thick_right=False)   # G
        write_sum(9, thick_right=True)    # J (thick RIGHT)
        write_sum(11, thick_right=False)  # L
        write_sum(14, thick_right=True)   # O (thick RIGHT)

        # ---------- Summary block in Q/R/S ----------
        # Headers: R3 = Calls, S3 = Texts
        worksheet.write(2, 17, "Calls", summary_header_fmt)  # R3
        worksheet.write(2, 18, "Texts", summary_header_fmt)  # S3
        # Labels: Q4/Q5/Q6
        worksheet.write(3, 16, "Chattanooga", summary_label_fmt)  # Q4
        worksheet.write(4, 16, "Cleveland", summary_label_fmt)    # Q5
        worksheet.write(5, 16, "Dalton", summary_label_fmt)       # Q6

        # Totals row Excel index (1-based)
        totals_row_excel = totals_row + 1

        # Formulas pointing to the totals we just wrote (B/E, G/J, L/O)
        worksheet.write_formula(3, 17, f"=B{totals_row_excel}", summary_number_fmt)  # R4 calls (Chatt)
        worksheet.write_formula(3, 18, f"=E{totals_row_excel}", summary_number_fmt)  # S4 texts (Chatt)

        worksheet.write_formula(4, 17, f"=G{totals_row_excel}", summary_number_fmt)  # R5 calls (Cleve)
        worksheet.write_formula(4, 18, f"=J{totals_row_excel}", summary_number_fmt)  # S5 texts (Cleve)

        worksheet.write_formula(5, 17, f"=L{totals_row_excel}", summary_number_fmt)  # R6 calls (Dalton)
        worksheet.write_formula(5, 18, f"=O{totals_row_excel}", summary_number_fmt)  # S6 texts (Dalton)
        # --------------------------------------------

        worksheet.freeze_panes(2, 0)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)

    output.seek(0)
    return output

def xlsx_col_letter(col_idx):
    """Convert 0-based column index to Excel column letters."""
    letters = ''
    x = col_idx + 1
    while x > 0:
        x, rem = divmod(x - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

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
                                  f"{int(chattanooga_final['Name_Highlight'].sum())} below 30/30")
                    with c2:
                        st.metric("Cleveland", f"{len(cleveland_final)} agents",
                                  f"{int(cleveland_final['Name_Highlight'].sum())} below 30/30")
                    with c3:
                        st.metric("Dalton", f"{len(dalton_final)} agents",
                                  f"{int(dalton_final['Name_Highlight'].sum())} below 30/30")

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
    - Agent name is highlighted if **Calls < 30 OR Text < 30**
    - **Calls** cell highlighted red if **Calls < 30**
    - **Text** cell highlighted red if **Text < 30**
    - (Talk time cells are not highlighted)

    ### Data Processing:
    - **Calls** = Carwars "Unique Outbound" + Tecobi "outbound_calls"
    - **Text** = Carwars "Unique OB Text" + Tecobi "external_sms"
    - Filters out any rows whose Agent Name contains "total" or "unassigned"
    - Agents are sorted alphabetically by first name
    - Automatically excludes: AJ Dhir, Thomas Williams, Mark Moore, Nicole Farr
    """)
