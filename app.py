import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
from datetime import datetime
import re

st.set_page_config(page_title="30/30 Report Combiner", layout="wide")

st.title("📊 30/30 Daily Report Combiner")
st.markdown("### Combine Carwars and Tecobi reports with 30/30 validation")

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
        if len(parts) == 2:  # MM:SS format
            minutes = int(parts[0])
            seconds = int(parts[1])
            total_hours = (minutes * 60 + seconds) / 3600
        elif len(parts) == 3:  # H:MM:SS format
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = int(parts[2])
            total_hours = hours + minutes/60 + seconds/3600
        else:
            return 0
        
        # Convert to Excel time (fraction of day)
        return total_hours / 24
    except:
        return 0

def get_first_name(full_name):
    """Extract first name for sorting"""
    if pd.isna(full_name):
        return ""
    return str(full_name).strip().split()[0] if str(full_name).strip() else ""

def process_carwars_file(df, location):
    """Process Carwars file and extract needed columns"""
    # Standardize column names
    df.columns = df.columns.str.strip()
    
    # Create processed dataframe
    processed = pd.DataFrame()
    processed['Agent Name'] = df['Agent Name'].str.strip()
    processed['Carwars_Unique_Outbound'] = pd.to_numeric(df['Unique Outbound'], errors='coerce').fillna(0)
    processed['Carwars_Avg_Talk_Time'] = df['Avg Talk Time'].apply(parse_time_to_excel)
    
    # Use Unique OB Text if available, otherwise Total OB Text
    if 'Unique OB Text' in df.columns:
        processed['Carwars_OB_Text'] = pd.to_numeric(df['Unique OB Text'], errors='coerce').fillna(0)
    else:
        processed['Carwars_OB_Text'] = pd.to_numeric(df.get('Total OB Text', 0), errors='coerce').fillna(0)
    
    processed['Location'] = location
    
    return processed

def process_tecobi_file(df, location):
    """Process Tecobi file and extract needed columns"""
    # Standardize column names
    df.columns = df.columns.str.strip()
    
    # Create full name from first and last name
    df['Agent Name'] = (df['first_name'].str.strip() + ' ' + df['last_name'].str.strip()).str.strip()
    
    # Calculate talk time if needed
    if 'avg_outbound_call_duration' in df.columns:
        talk_time = pd.to_numeric(df['avg_outbound_call_duration'], errors='coerce').fillna(0)
    else:
        # Calculate from seconds_clocked_in / outbound_calls if available
        seconds = pd.to_numeric(df.get('seconds_clocked_in', 0), errors='coerce').fillna(0)
        calls = pd.to_numeric(df.get('outbound_calls', 0), errors='coerce').fillna(1)
        calls = calls.replace(0, 1)  # Avoid division by zero
        talk_time = seconds / calls
    
    # Create processed dataframe
    processed = pd.DataFrame()
    processed['Agent Name'] = df['Agent Name']
    processed['Tecobi_Outbound_Calls'] = pd.to_numeric(df.get('outbound_calls', 0), errors='coerce').fillna(0)
    processed['Tecobi_Talk_Time'] = talk_time
    processed['Tecobi_External_SMS'] = pd.to_numeric(df.get('external_sms', 0), errors='coerce').fillna(0)
    processed['Location'] = location
    
    return processed

def combine_location_data(carwars_df, tecobi_df, location):
    """Combine Carwars and Tecobi data for a single location"""
    
    # Filter data for this location
    carwars_loc = carwars_df[carwars_df['Location'] == location].copy()
    tecobi_loc = tecobi_df[tecobi_df['Location'] == location].copy()
    
    # Remove location column for merging
    carwars_loc = carwars_loc.drop('Location', axis=1)
    tecobi_loc = tecobi_loc.drop('Location', axis=1)
    
    # Merge on Agent Name with outer join to keep all agents
    combined = pd.merge(
        carwars_loc,
        tecobi_loc,
        on='Agent Name',
        how='outer'
    )
    
    # Fill NaN values with 0 for numeric columns
    numeric_columns = [
        'Carwars_Unique_Outbound', 'Carwars_Avg_Talk_Time', 'Carwars_OB_Text',
        'Tecobi_Outbound_Calls', 'Tecobi_Talk_Time', 'Tecobi_External_SMS'
    ]
    
    for col in numeric_columns:
        if col in combined.columns:
            combined[col] = combined[col].fillna(0)
    
    # Calculate combined metrics
    combined['Calls'] = combined['Carwars_Unique_Outbound'] + combined['Tecobi_Outbound_Calls']
    combined['Text'] = combined['Carwars_OB_Text'] + combined['Tecobi_External_SMS']
    
    # Check if meets 30/30 criteria
    combined['Needs_Highlight'] = (combined['Calls'] < 30) | (combined['Text'] < 30)
    
    # Create final dataframe for this location
    final = pd.DataFrame()
    final['Agent Name'] = combined['Agent Name']
    final['Calls'] = combined['Calls'].astype(int)
    final['Carwars Avg Talk Time'] = combined['Carwars_Avg_Talk_Time']
    final['Tecobi Talk Time'] = combined['Tecobi_Talk_Time']
    final['Text'] = combined['Text'].astype(int)
    final['Needs_Highlight'] = combined['Needs_Highlight']
    
    # Add first name for sorting
    final['First_Name'] = final['Agent Name'].apply(get_first_name)
    
    # Sort by first name alphabetically
    final = final.sort_values('First_Name', na_position='last')
    
    # Drop the sorting column but keep highlight flag
    final = final.drop('First_Name', axis=1)
    
    return final

def create_formatted_excel(chattanooga_data, cleveland_data, dalton_data):
    """Create the final formatted Excel file matching 30/30 format"""
    
    # Create a BytesIO buffer
    output = BytesIO()
    
    # Create Excel writer
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats matching the 30/30 formatted file
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1,
            'bg_color': '#D7D7D7',
            'font_size': 10
        })
        
        location_header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'font_size': 11,
            'bg_color': '#B4C6E7',
            'border': 1
        })
        
        time_format = workbook.add_format({
            'num_format': '[h]:mm:ss',
            'align': 'center',
            'border': 1
        })
        
        time_format_highlight = workbook.add_format({
            'num_format': '[h]:mm:ss',
            'align': 'center',
            'border': 1,
            'bg_color': '#FFC7CE'  # Light red
        })
        
        number_format = workbook.add_format({
            'align': 'center',
            'border': 1
        })
        
        number_format_highlight = workbook.add_format({
            'align': 'center',
            'border': 1,
            'bg_color': '#FFC7CE'  # Light red
        })
        
        text_format = workbook.add_format({
            'align': 'left',
            'border': 1
        })
        
        text_format_highlight = workbook.add_format({
            'align': 'left',
            'border': 1,
            'bg_color': '#FFC7CE'  # Light red
        })
        
        empty_format = workbook.add_format({
            'border': 1
        })
        
        # Create worksheet
        worksheet = writer.book.add_worksheet('Sheet1')
        
        # Set column widths matching the original
        worksheet.set_column('A:A', 18)  # Chattanooga - Agent Name
        worksheet.set_column('B:B', 8)   # Calls
        worksheet.set_column('C:C', 12)  # Carwars Avg Talk Time
        worksheet.set_column('D:D', 12)  # Tecobi Talk Time
        worksheet.set_column('E:E', 8)   # Text
        
        worksheet.set_column('F:F', 18)  # Cleveland - Agent Name
        worksheet.set_column('G:G', 8)   # Calls
        worksheet.set_column('H:H', 12)  # Carwars Avg Talk Time
        worksheet.set_column('I:I', 12)  # Tecobi Talk Time
        worksheet.set_column('J:J', 8)   # Text
        
        worksheet.set_column('K:K', 18)  # Dalton - Agent Name
        worksheet.set_column('L:L', 8)   # Calls
        worksheet.set_column('M:M', 12)  # Carwars Avg Talk Time
        worksheet.set_column('N:N', 12)  # Tecobi Talk Time
        worksheet.set_column('O:O', 8)   # Text
        
        # Add empty columns between sections for visual separation
        worksheet.set_column('P:P', 2)
        worksheet.set_column('Q:Q', 2)
        
        # Write location headers (row 0)
        worksheet.merge_range('A1:E1', 'Chattanooga', location_header_format)
        worksheet.merge_range('F1:J1', 'Cleveland', location_header_format)
        worksheet.merge_range('K1:O1', 'Dalton', location_header_format)
        
        # Write column headers (row 1)
        headers = ['Agent Name', 'Calls', 'Carwars Avg\nTalk Time', 'Tecobi\nTalk Time\n(seconds)', 'Text']
        
        # Chattanooga headers
        for i, header in enumerate(headers):
            worksheet.write(1, i, header, header_format)
        
        # Cleveland headers
        for i, header in enumerate(headers):
            worksheet.write(1, 5 + i, header, header_format)
        
        # Dalton headers
        for i, header in enumerate(headers):
            worksheet.write(1, 10 + i, header, header_format)
        
        # Write data for each location
        max_rows = max(len(chattanooga_data), len(cleveland_data), len(dalton_data))
        
        for row_idx in range(max_rows):
            excel_row = row_idx + 2  # Start from row 2 (0-indexed)
            
            # Chattanooga
            if row_idx < len(chattanooga_data):
                row_data = chattanooga_data.iloc[row_idx]
                needs_highlight = row_data['Needs_Highlight']
                
                # Choose format based on highlight need
                name_fmt = text_format_highlight if needs_highlight else text_format
                num_fmt = number_format_highlight if needs_highlight else number_format
                time_fmt = time_format_highlight if needs_highlight else time_format
                
                worksheet.write(excel_row, 0, row_data['Agent Name'], name_fmt)
                worksheet.write(excel_row, 1, row_data['Calls'], num_fmt)
                worksheet.write(excel_row, 2, row_data['Carwars Avg Talk Time'], time_fmt)
                worksheet.write(excel_row, 3, row_data['Tecobi Talk Time'], num_fmt)
                worksheet.write(excel_row, 4, row_data['Text'], num_fmt)
            else:
                # Write empty cells with borders
                for col in range(5):
                    worksheet.write(excel_row, col, '', empty_format)
            
            # Cleveland
            if row_idx < len(cleveland_data):
                row_data = cleveland_data.iloc[row_idx]
                needs_highlight = row_data['Needs_Highlight']
                
                # Choose format based on highlight need
                name_fmt = text_format_highlight if needs_highlight else text_format
                num_fmt = number_format_highlight if needs_highlight else number_format
                time_fmt = time_format_highlight if needs_highlight else time_format
                
                worksheet.write(excel_row, 5, row_data['Agent Name'], name_fmt)
                worksheet.write(excel_row, 6, row_data['Calls'], num_fmt)
                worksheet.write(excel_row, 7, row_data['Carwars Avg Talk Time'], time_fmt)
                worksheet.write(excel_row, 8, row_data['Tecobi Talk Time'], num_fmt)
                worksheet.write(excel_row, 9, row_data['Text'], num_fmt)
            else:
                # Write empty cells with borders
                for col in range(5, 10):
                    worksheet.write(excel_row, col, '', empty_format)
            
            # Dalton
            if row_idx < len(dalton_data):
                row_data = dalton_data.iloc[row_idx]
                needs_highlight = row_data['Needs_Highlight']
                
                # Choose format based on highlight need
                name_fmt = text_format_highlight if needs_highlight else text_format
                num_fmt = number_format_highlight if needs_highlight else number_format
                time_fmt = time_format_highlight if needs_highlight else time_format
                
                worksheet.write(excel_row, 10, row_data['Agent Name'], name_fmt)
                worksheet.write(excel_row, 11, row_data['Calls'], num_fmt)
                worksheet.write(excel_row, 12, row_data['Carwars Avg Talk Time'], time_fmt)
                worksheet.write(excel_row, 13, row_data['Tecobi Talk Time'], num_fmt)
                worksheet.write(excel_row, 14, row_data['Text'], num_fmt)
            else:
                # Write empty cells with borders
                for col in range(10, 15):
                    worksheet.write(excel_row, col, '', empty_format)
        
        # Freeze panes (freeze first 2 rows)
        worksheet.freeze_panes(2, 0)
        
        # Set print settings for better printing
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide
    
    output.seek(0)
    return output

# Main app interface
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
    
    # Check if all files are uploaded
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
                    # Define excluded agents list here
                    EXCLUDED_AGENTS = ['AJ Dhir', 'Aj Dhir', 'Thomas Williams', 'Mark Moore', 'Nicole Farr']
                    
                    # Read all files
                    carwars_files = {}
                    tecobi_files = {}
                    
                    # Helper function to read file
                    def read_file(uploaded_file):
                        if uploaded_file.name.endswith('.csv'):
                            return pd.read_csv(uploaded_file)
                        else:
                            return pd.read_excel(uploaded_file)
                    
                    # Read Carwars files with exclusion list
                    carwars_files['Chattanooga'] = process_carwars_file(
                        read_file(chatt_carwars), 'Chattanooga', EXCLUDED_AGENTS
                    )
                    carwars_files['Cleveland'] = process_carwars_file(
                        read_file(cleve_carwars), 'Cleveland', EXCLUDED_AGENTS
                    )
                    carwars_files['Dalton'] = process_carwars_file(
                        read_file(dalton_carwars), 'Dalton', EXCLUDED_AGENTS
                    )
                    
                    # Read Tecobi files with exclusion list
                    tecobi_files['Chattanooga'] = process_tecobi_file(
                        read_file(chatt_tecobi), 'Chattanooga', EXCLUDED_AGENTS
                    )
                    tecobi_files['Cleveland'] = process_tecobi_file(
                        read_file(cleve_tecobi), 'Cleveland', EXCLUDED_AGENTS
                    )
                    tecobi_files['Dalton'] = process_tecobi_file(
                        read_file(dalton_tecobi), 'Dalton', EXCLUDED_AGENTS
                    )
                    
                    # Combine all Carwars data
                    all_carwars = pd.concat(carwars_files.values(), ignore_index=True)
                    
                    # Combine all Tecobi data
                    all_tecobi = pd.concat(tecobi_files.values(), ignore_index=True)
                    
                    # Process each location
                    chattanooga_final = combine_location_data(all_carwars, all_tecobi, 'Chattanooga')
                    cleveland_final = combine_location_data(all_carwars, all_tecobi, 'Cleveland')
                    dalton_final = combine_location_data(all_carwars, all_tecobi, 'Dalton')
                    
                    # Show summary statistics
                    st.markdown("### 📊 Summary")
                    col_sum1, col_sum2, col_sum3 = st.columns(3)
                    
                    with col_sum1:
                        st.metric("Chattanooga", 
                                 f"{len(chattanooga_final)} agents",
                                 f"{sum(chattanooga_final['Needs_Highlight'])} below 30/30")
                    
                    with col_sum2:
                        st.metric("Cleveland", 
                                 f"{len(cleveland_final)} agents",
                                 f"{sum(cleveland_final['Needs_Highlight'])} below 30/30")
                    
                    with col_sum3:
                        st.metric("Dalton", 
                                 f"{len(dalton_final)} agents",
                                 f"{sum(dalton_final['Needs_Highlight'])} below 30/30")
                    
                    # Create formatted Excel file
                    st.session_state.result_buffer = create_formatted_excel(
                        chattanooga_final, cleveland_final, dalton_final
                    )
                    st.session_state.processed = True
                    
                st.success("✅ 30/30 Report generated successfully!")
                
            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
                st.markdown("**Debug Information:**")
                st.code(str(e))
    else:
        st.warning("⚠️ Please upload all 6 files to continue")
        
        # Show which files are missing
        missing_files = []
        if not chatt_carwars: missing_files.append("Chattanooga Carwars")
        if not chatt_tecobi: missing_files.append("Chattanooga Tecobi")
        if not cleve_carwars: missing_files.append("Cleveland Carwars")
        if not cleve_tecobi: missing_files.append("Cleveland Tecobi")
        if not dalton_carwars: missing_files.append("Dalton Carwars")
        if not dalton_tecobi: missing_files.append("Dalton Tecobi")
        
        if missing_files:
            st.markdown("**Missing files:**")
            for file in missing_files:
                st.markdown(f"- {file}")

# Download section
if st.session_state.processed and st.session_state.result_buffer:
    st.markdown("---")
    st.subheader("📥 Download Result")
    
    # Generate filename with current date
    current_date = datetime.now().strftime("%m_%d_%Y")
    filename = f"30_30_Report_{current_date}_Formatted.xlsx"
    
    st.download_button(
        label="⬇️ Download 30/30 Report",
        data=st.session_state.result_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    # Add reset button
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
    """)
