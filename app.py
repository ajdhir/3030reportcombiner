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

def process_carwars_file_v2(df, location, exclude_list=None):
    """Process Carwars file and extract needed columns"""
    # Standardize column names
    df.columns = df.columns.str.strip()
    
    # Filter out specific excluded agents
    if exclude_list:
        for name in exclude_list:
            df = df[~df['Agent Name'].str.lower().str.contains(name.lower(), na=False)]
    
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

def process_tecobi_file(df, location, exclude_list=None):
    """Process Tecobi file and extract needed columns"""
    # Standardize column names
    df.columns = df.columns.str.strip()
    
    # Create full name from first and last name
    df['Agent Name'] = (df['first_name'].str.strip() + ' ' + df['last_name'].str.strip()).str.strip()
    
    # Filter out specific excluded agents
    if exclude_list:
        for name in exclude_list:
            df = df[~df['Agent Name'].str.lower().str.contains(name.lower(), na=False)]
    
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
                
                # Choose format based on highligh
