"""
Excel conversion functions for different process types
"""

import pandas as pd
import io
from typing import List, Tuple, Dict, Optional
import re
import datetime
import os


def extract_file_number(filename: str) -> int:
    """Extract sequence number from filename (1), (2), etc."""
    filename = filename.split('.')[0]
    match = re.search(r'\((\d+)\)', filename)
    if match:
        try:
            return int(match.group(1))
        except:
            pass
    return 0


def sort_files_by_sequence(files: List[Tuple[str, bytes]]) -> List[Tuple[str, bytes]]:
    """Sort files by extracted sequence number"""
    files_info = []
    
    for filename, data in files:
        file_number = extract_file_number(filename)
        files_info.append({
            'number': file_number,
            'filename': filename,
            'data': data,
        })
    
    files_info.sort(key=lambda x: x['number'])
    
    return [(info['filename'], info['data']) for info in files_info]


def load_port_code_mapping() -> Dict[str, str]:
    """Load port code mapping from Port Code List.xlsx"""
    port_mapping = {}
    reverse_port_mapping = {}  # For searching by code
    
    try:
        # First check if the port code file exists in the current directory
        port_file_path = "Port Code List.xlsx"
        
        if os.path.exists(port_file_path):
            # Load the port code mapping file
            port_df = pd.read_excel(port_file_path)
            
            # Clean column names (remove newlines)
            port_df.columns = [col.replace('\n', ' ').strip() for col in port_df.columns]
            
            print(f"Port file columns: {port_df.columns.tolist()}")
            
            # Look for the correct columns - based on the file structure
            # Column A: Sl no
            # Column B: Location Code
            # Column D: Location Name
            # Column E: State
            
            location_code_col = None
            location_name_col = None
            state_col = None
            
            # Try to find columns by common patterns
            for col in port_df.columns:
                col_lower = str(col).lower()
                if 'code' in col_lower and 'location' in col_lower:
                    location_code_col = col
                elif 'name' in col_lower and 'location' in col_lower:
                    location_name_col = col
                elif 'state' in col_lower:
                    state_col = col
            
            # If not found by patterns, use positional mapping
            if not location_code_col and len(port_df.columns) >= 2:
                location_code_col = port_df.columns[1]  # Column B
            if not location_name_col and len(port_df.columns) >= 4:
                location_name_col = port_df.columns[3]  # Column D
            if not state_col and len(port_df.columns) >= 5:
                state_col = port_df.columns[4]  # Column E
            
            print(f"Using Location Code column: '{location_code_col}'")
            print(f"Using Location Name column: '{location_name_col}'")
            print(f"Using State column: '{state_col if state_col else 'Not found'}'")
            
            # Create comprehensive mapping
            for _, row in port_df.iterrows():
                try:
                    location_code = row[location_code_col] if location_code_col in row else None
                    location_name = row[location_name_col] if location_name_col in row else None
                    
                    if pd.notna(location_code) and pd.notna(location_name):
                        location_code_str = str(location_code).strip()
                        location_name_str = str(location_name).strip()
                        
                        # Add primary mapping
                        port_mapping[location_name_str] = location_code_str
                        
                        # Store reverse mapping
                        reverse_port_mapping[location_code_str] = location_name_str
                        
                        # Add mapping with state if available
                        if state_col and state_col in row:
                            state = str(row[state_col]).strip() if pd.notna(row[state_col]) else ''
                            if state:
                                port_mapping[f"{location_name_str} - {state}"] = location_code_str
                        
                        # Extract 6-digit PIN codes if present
                        pin_codes = re.findall(r'\b(\d{6})\b', location_name_str)
                        for pin in pin_codes:
                            port_mapping[pin] = location_code_str
                        
                        # Create variations of the location name for better matching
                        # 1. Remove common prefixes/suffixes
                        clean_name = location_name_str
                        prefixes_to_remove = ['CUSTOM HOUSE', 'CUSTOMS HOUSE', 'SEAPORT', 'PORT', 
                                            'AIR CARGO COMPLEX', 'ACC', 'ICD', 'LCS', 'LAND CUSTOMS STATION']
                        
                        for prefix in prefixes_to_remove:
                            if prefix in clean_name.upper():
                                clean_name = re.sub(f'{prefix}[,\\s]*', '', clean_name, flags=re.IGNORECASE).strip()
                        
                        if clean_name and clean_name != location_name_str:
                            port_mapping[clean_name] = location_code_str
                        
                        # 2. Create acronym or short form
                        words = re.split(r'[,\s-]+', clean_name)
                        if len(words) > 1:
                            # Take first letters of major words
                            short_form = ''.join([w[0].upper() for w in words if w and len(w) > 2])
                            if short_form:
                                port_mapping[short_form] = location_code_str
                        
                        # 3. Create mapping without special characters
                        clean_name_no_special = re.sub(r'[^\w\s-]', ' ', location_name_str)
                        clean_name_no_special = re.sub(r'\s+', ' ', clean_name_no_special).strip()
                        if clean_name_no_special and clean_name_no_special != location_name_str:
                            port_mapping[clean_name_no_special] = location_code_str
                
                except Exception as e:
                    print(f"Error processing row {_}: {e}")
                    continue
            
            print(f"Loaded {len(port_mapping)} port code mappings")
            
            # Debug: Print some sample mappings
            print("\nSample port mappings:")
            sample_count = 0
            for key, value in list(port_mapping.items()):
                if 'NHAVA' in key.upper() or 'SAHAR' in key.upper() or '400707' in key or '400099' in key:
                    display_key = key if len(key) <= 60 else key[:57] + "..."
                    print(f"  '{display_key}' -> '{value}'")
                    sample_count += 1
                if sample_count >= 10:
                    break
        
        else:
            print(f"Port Code List file not found at: {port_file_path}")
            
    except Exception as e:
        print(f"Error loading port code mapping: {e}")
        import traceback
        traceback.print_exc()
    
    return port_mapping


def get_port_code(port_name: str, port_mapping: Dict[str, str]) -> str:
    """Get port code from port name using mapping"""
    if not port_name or pd.isna(port_name):
        return ''
    
    port_str = str(port_name).strip()
    if not port_str:
        return ''
    
    # If it's already a 6-character code (like INNSA1)
    if len(port_str) == 6 and port_str.isalnum() and port_str[:2].isalpha() and port_str[2:].isalnum():
        return port_str
    
    # Check if it's already in the mapping (exact match)
    if port_str in port_mapping:
        return port_mapping[port_str]
    
    # Clean the port string
    port_clean = port_str.upper()
    
    # Try to find 6-digit PIN code in the string
    pin_match = re.search(r'\b(\d{6})\b', port_clean)
    if pin_match:
        pin_code = pin_match.group(1)
        if pin_code in port_mapping:
            return port_mapping[pin_code]
    
    # Try partial matches with priority
    # 1. Try matching by common keywords first
    common_keywords = {
        'NHAVA SHEVA': 'INNSA1',
        'JNCH': 'INNSA1',
        'SAHAR ANDHERI': 'INBOM4',
        'SAHAR, ANDHERI': 'INBOM4',
        'MUMBAI AIRPORT': 'INBOM4',
        'AIR CARGO SAHAR': 'INBOM4',
        '400707': 'INNSA1',  # PIN for NHAVA SHEVA
        '400099': 'INBOM4',  # PIN for SAHAR ANDHERI
    }
    
    for keyword, code in common_keywords.items():
        if keyword in port_clean:
            return code
    
    # 2. Try case-insensitive partial matching
    port_clean_lower = port_str.lower()
    best_match = None
    best_match_score = 0
    
    for map_key, map_code in port_mapping.items():
        map_key_lower = str(map_key).lower()
        
        # Calculate match score
        score = 0
        
        # Exact substring match
        if map_key_lower in port_clean_lower or port_clean_lower in map_key_lower:
            score = 100
        
        # Word-based matching
        port_words = set(re.findall(r'\b\w+\b', port_clean_lower))
        map_words = set(re.findall(r'\b\w+\b', map_key_lower))
        common_words = port_words.intersection(map_words)
        
        if common_words:
            score = max(score, len(common_words) * 20)
        
        # If this is a better match, update
        if score > best_match_score:
            best_match_score = score
            best_match = map_code
    
    if best_match and best_match_score > 30:  # Threshold for a good match
        return best_match
    
    # 3. Try acronym matching
    port_acronym = ''.join([w[0].upper() for w in re.findall(r'\b\w+\b', port_clean) if w])
    if port_acronym and port_acronym in port_mapping:
        return port_mapping[port_acronym]
    
    # 4. Try fuzzy matching for common ports
    fuzzy_matches = {
        'JNCH, NHAVA SHEVA': 'INNSA1',
        'NHAVA SHEVA PORT': 'INNSA1',
        'SAHAR AIR CARGO': 'INBOM4',
        'ANDHERI AIR CARGO': 'INBOM4',
        'MUMBAI SAHAR': 'INBOM4',
    }
    
    for fuzzy_key, fuzzy_code in fuzzy_matches.items():
        if fuzzy_key in port_clean:
            return fuzzy_code
    
    # If no match found, return original (or try to extract a code if possible)
    # Try to extract any 6-char alphanumeric code
    code_match = re.search(r'\b([A-Z]{2}\d{4}|[A-Z]{2}\d{3}[A-Z]|[A-Z]{6})\b', port_clean)
    if code_match:
        return code_match.group(1)
    
    return port_str  # Return original if no match


def get_currency_code(currency_name: str) -> str:
    """Convert currency name to standard 3-letter code"""
    if not currency_name or pd.isna(currency_name):
        return ''
    
    currency_str = str(currency_name).strip().upper()
    if not currency_str:
        return ''
    
    # Currency mapping dictionary
    currency_map = {
        # Major currencies
        'US DOLLARS': 'USD',
        'USD': 'USD',
        'US DOLLAR': 'USD',
        'U.S. DOLLAR': 'USD',
        'U.S. DOLLARS': 'USD',
        'DOLLAR': 'USD',
        'DOLLARS': 'USD',
        
        'EURO': 'EUR',
        'EUROS': 'EUR',
        'EUR': 'EUR',
        
        'POUND': 'GBP',
        'POUNDS': 'GBP',
        'BRITISH POUND': 'GBP',
        'GBP': 'GBP',
        
        'YEN': 'JPY',
        'JAPANESE YEN': 'JPY',
        'JPY': 'JPY',
        
        'AUSTRALIAN DOLLAR': 'AUD',
        'AUD': 'AUD',
        
        'CANADIAN DOLLAR': 'CAD',
        'CAD': 'CAD',
        
        'SWISS FRANC': 'CHF',
        'CHF': 'CHF',
        
        'YUAN': 'CNY',
        'CHINESE YUAN': 'CNY',
        'RENMINBI': 'CNY',
        'CNY': 'CNY',
        
        # Indian currency
        'INDIAN RUPEE': 'INR',
        'RUPEE': 'INR',
        'RUPES': 'INR',
        'RS': 'INR',
        '₹': 'INR',
        'INR': 'INR',
        
        # Other currencies
        'SINGAPORE DOLLAR': 'SGD',
        'SGD': 'SGD',
        
        'HONG KONG DOLLAR': 'HKD',
        'HKD': 'HKD',
        
        'NEW ZEALAND DOLLAR': 'NZD',
        'NZD': 'NZD',
        
        'SWEDISH KRONA': 'SEK',
        'SEK': 'SEK',
        
        'NORWEGIAN KRONE': 'NOK',
        'NOK': 'NOK',
        
        'DANISH KRONE': 'DKK',
        'DKK': 'DKK',
    }
    
    # Check for exact match
    if currency_str in currency_map:
        return currency_map[currency_str]
    
    # Check for partial match
    for key, code in currency_map.items():
        if key in currency_str or currency_str in key:
            return code
    
    # If currency is already 3 uppercase letters, assume it's already a code
    if len(currency_str) == 3 and currency_str.isalpha() and currency_str.isupper():
        return currency_str
    
    # Return original if no match
    return currency_str


def merge_excel_files(files: List[Tuple[str, bytes]]) -> pd.DataFrame:
    """Merge multiple Excel files into one DataFrame"""
    if not files:
        raise ValueError("No files provided")
    
    # Sort files by sequence
    sorted_files = sort_files_by_sequence(files)
    
    merged_data = []
    
    for i, (filename, file_data) in enumerate(sorted_files):
        try:
            print(f"\nProcessing file {i+1}: {filename}")
            
            # Determine engine based on file extension
            if filename.lower().endswith('.xls'):
                # For .xls files, try xlrd engine
                try:
                    df = pd.read_excel(io.BytesIO(file_data), engine='xlrd', header=0)
                except ImportError:
                    raise ImportError(
                        "Missing dependency: xlrd library is required to read .xls files. "
                        "Please install it using: pip install xlrd"
                    )
            elif filename.lower().endswith('.xlsx'):
                # For .xlsx files, use openpyxl engine
                df = pd.read_excel(io.BytesIO(file_data), engine='openpyxl', header=0)
            else:
                # Try default engine
                df = pd.read_excel(io.BytesIO(file_data), header=0)
            
            print(f"  File shape: {df.shape}")
            print(f"  File columns: {df.columns.tolist()}")
            
            # For BRC files specifically - SIMPLE LOGIC
            if i == 0:
                # First file - keep all rows
                print(f"  First file - keeping all {len(df)} rows")
                merged_data.append(df)
            else:
                # For subsequent files, we need to check if first row is header
                # BRC files have header row with specific column names
                if len(df) > 0:
                    first_row = df.iloc[0] if not df.empty else None
                    is_header = False
                    
                    if first_row is not None:
                        # Check if first row contains header-like values
                        header_count = 0
                        for cell in first_row:
                            if isinstance(cell, str):
                                cell_lower = str(cell).lower()
                                # BRC header keywords
                                if any(keyword in cell_lower for keyword in ['brc', 'sb', 'date', 'number', 'port', 'invoice', 'currency', 'realization']):
                                    header_count += 1
                        
                        # If more than 3 columns look like headers, it's probably a header row
                        if header_count > 3:
                            is_header = True
                            print(f"  Detected header row - skipping it")
                    
                    if is_header:
                        # Skip only the first row (header)
                        data_to_append = df.iloc[1:].reset_index(drop=True)
                    else:
                        # No header detected, keep all rows
                        data_to_append = df
                    
                    print(f"  Appending {len(data_to_append)} data rows")
                    merged_data.append(data_to_append)
                
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            continue
    
    if not merged_data:
        raise ValueError("No valid data found")
    
    # Concatenate all dataframes
    print(f"\nConcatenating {len(merged_data)} dataframes")
    combined_df = pd.concat(merged_data, ignore_index=True)
    
    print(f"Final merged DataFrame shape: {combined_df.shape}")
    
    return combined_df


def convert_dbk_disbursement(df: pd.DataFrame) -> pd.DataFrame:
    """Convert DBK Disbursement Excel format to required output format"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1
    serial_number = 1
    
    # Helper function to convert to datetime and format as string
    def convert_and_format_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            # Format as '09-Jul-2025' (with 4-digit year)
            return date_value.strftime('%d-%b-%Y')
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # Special handling for two-digit years - convert to four-digit (20XX)
            # Look for patterns like dd-mmm-yy or dd-mmm-yyyy
            date_formats = [
                '%d-%b-%y',      # 09-Jul-25 (will convert 25 to 2025)
                '%d-%b-%Y',      # 09-Jul-2025
                '%Y-%m-%d',      # 2025-07-09
                '%d/%m/%Y',      # 09/07/2025
                '%m/%d/%Y',      # 07/09/2025
                '%Y/%m/%d',      # 2025/07/09
                '%d-%m-%Y',      # 09-07-2025
                '%m-%d-%Y',      # 07-09-2025
            ]
            
            # Try each format
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    # Force 4-digit year output
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    continue
            
            # Try with case-insensitive month
            if any(month in date_str.upper() for month in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                                                          'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']):
                # Try with two-digit year first (most common case)
                for month in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                             'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']:
                    if month in date_str.upper():
                        # Convert to proper case
                        date_str_proper = date_str.upper().replace(month, month.capitalize())
                        try:
                            # Try with two-digit year format
                            date_obj = pd.to_datetime(date_str_proper, format='%d-%b-%y')
                            return date_obj.strftime('%d-%b-%Y')
                        except:
                            try:
                                # Try with four-digit year format
                                date_obj = pd.to_datetime(date_str_proper, format='%d-%b-%Y')
                                return date_obj.strftime('%d-%b-%Y')
                            except:
                                pass
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            # Excel serial date starts from 1899-12-30
            try:
                date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        return ''
    
    # Helper function to convert to number
    def convert_to_number(value):
        if pd.isna(value):
            return None
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it
            if isinstance(value, str):
                cleaned = ''.join(c for c in value if c.isdigit() or c == '.' or c == '-')
                if not cleaned:
                    return None
                
                num_value = float(cleaned)
                return int(num_value) if num_value == int(num_value) else num_value
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return None
    
    # Iterate through each row in the merged data
    for index, row in df.iterrows():
        try:
            # Extract values
            shipping_bill_no = convert_to_number(row['SB No']) if 'SB No' in df.columns else None
            shipping_bill_date = convert_and_format_date(row['SB Date']) if 'SB Date' in df.columns else ''
            scroll_no = convert_to_number(row['Custom Scroll No']) if 'Custom Scroll No' in df.columns else None
            scroll_date = convert_and_format_date(row['Custom Scroll Date']) if 'Custom Scroll Date' in df.columns else ''
            port = str(row['Location']) if 'Location' in df.columns and not pd.isna(row['Location']) else ''
            amount = convert_to_number(row['IgstAmount']) if 'IgstAmount' in df.columns else None
            
            # Debug for first few rows
            if index < 3:
                print(f"\nRow {index} conversion:")
                print(f"  SB Date raw: {row['SB Date'] if 'SB Date' in df.columns else 'N/A'} -> {shipping_bill_date}")
                print(f"  Scroll Date raw: {row['Custom Scroll Date'] if 'Custom Scroll Date' in df.columns else 'N/A'} -> {scroll_date}")
            
            # Create a new row in the required format
            new_row = {
                'S No.': int(serial_number),
                'Port': port,
                'Shipping Bill No.': shipping_bill_no if shipping_bill_no is not None else 0,
                'Shipping Bill Date': shipping_bill_date,  # Formatted as string '09-Jul-25'
                'Scroll No.': scroll_no if scroll_no is not None else 0,
                'Scroll Date': scroll_date,  # Formatted as string '31-Oct-25'
                'Drawback': '',
                'STR': '',
                'Amount': amount if amount is not None else 0
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            continue
    
    # Create the result DataFrame
    result_df = pd.DataFrame(result_data)
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    if not result_df.empty:
        print(f"  Sample Shipping Bill Dates: {result_df['Shipping Bill Date'].head(3).tolist()}")
        print(f"  Sample Scroll Dates: {result_df['Scroll Date'].head(3).tolist()}")
    
    # Create the header rows
    header_rows = pd.DataFrame([
        {'S No.': 'IEC Name - Alfa', 'Port': '', 'Shipping Bill No.': '', 'Shipping Bill Date': '', 
         'Scroll No.': '', 'Scroll Date': '', 'Drawback': '', 'STR': '', 'Amount': ''},
        {'S No.': 'Report Generated From - 2021/01/01To2021/06/30', 'Port': '', 'Shipping Bill No.': '', 
         'Shipping Bill Date': '', 'Scroll No.': '', 'Scroll Date': '', 'Drawback': '', 'STR': '', 'Amount': ''},
        {'S No.': 'Report Generated On - 16/02/2022 12:58:50', 'Port': '', 'Shipping Bill No.': '', 
         'Shipping Bill Date': '', 'Scroll No.': '', 'Scroll Date': '', 'Drawback': '', 'STR': '', 'Amount': ''},
        # Column headers
        {'S No.': 'S No.', 'Port': 'Port', 'Shipping Bill No.': 'Shipping Bill No.',
         'Shipping Bill Date': 'Shipping Bill Date', 'Scroll No.': 'Scroll No.',
         'Scroll Date': 'Scroll Date', 'Drawback': 'Drawback', 'STR': 'STR', 'Amount': 'Amount'}
    ])
    
    # Concatenate header rows with data
    final_df = pd.concat([header_rows, result_df], ignore_index=True)
    
    return final_df

def convert_dbk_pendency(df: pd.DataFrame) -> pd.DataFrame:
    """Convert DBK Pendency Excel format"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns for DBK Pendency: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1 (for data rows, starting from row 5)
    serial_number = 1
    
    # Helper function to convert to datetime and format as string
    def convert_and_format_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            # Format as '31-OCT-2025'
            return date_value.strftime('%d-%b-%Y').upper()
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # Try common date formats
            date_formats = [
                '%d-%b-%Y',      # 31-OCT-2025
                '%d-%b-%y',      # 31-OCT-25
                '%Y-%m-%d',      # 2025-10-31
                '%d/%m/%Y',      # 31/10/2025
                '%m/%d/%Y',      # 10/31/2025
                '%Y/%m/%d',      # 2025/10/31
                '%d-%m-%Y',      # 31-10-2025
                '%m-%d-%Y',      # 10-31-2025
            ]
            
            # Try each format
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    # Format as '31-OCT-2025'
                    return date_obj.strftime('%d-%b-%Y').upper()
                except:
                    continue
            
            # Try with case-insensitive month
            if any(month in date_str.upper() for month in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                                                          'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']):
                # Try with two-digit year first (most common case)
                try:
                    date_obj = pd.to_datetime(date_str, format='%d-%b-%y')
                    return date_obj.strftime('%d-%b-%Y').upper()
                except:
                    try:
                        date_obj = pd.to_datetime(date_str, format='%d-%b-%Y')
                        return date_obj.strftime('%d-%b-%Y').upper()
                    except:
                        pass
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y').upper()
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            # Excel serial date starts from 1899-12-30
            try:
                date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                return date_obj.strftime('%d-%b-%Y').upper()
            except:
                return ''
        
        return ''
    
    # Helper function to convert to number
    def convert_to_number(value):
        if pd.isna(value):
            return None
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it
            if isinstance(value, str):
                cleaned = ''.join(c for c in value if c.isdigit() or c == '.' or c == '-')
                if not cleaned:
                    return None
                
                num_value = float(cleaned)
                return int(num_value) if num_value == int(num_value) else num_value
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return None
    
    # Iterate through each row in the merged data (starting from row 1 as data starts from row 1 in merged file)
    for index, row in df.iterrows():
        try:
            # Extract values based on your mapping:
            # Output column 2 (Shipping Bill No.) -> merged file column A (SB No)
            # Output column 3 (Shipping Bill Date) -> merged file column B (SB Date)
            # Output column 4 (LEO Date) -> merged file column G (Leo Date)
            # Output column 5 (Amount) -> merged file column E (DBK Amount RS)
            # Output column 6 (Current Queue) -> merged file column F (Curr Queue)
            
            # Map the columns
            shipping_bill_no = convert_to_number(row['SB No']) if 'SB No' in df.columns else None
            shipping_bill_date = convert_and_format_date(row['SB Date']) if 'SB Date' in df.columns else ''
            leo_date = convert_and_format_date(row['Leo Date']) if 'Leo Date' in df.columns else ''
            amount = convert_to_number(row['DBK Amount RS']) if 'DBK Amount RS' in df.columns else None
            current_queue = str(row['Curr Queue']) if 'Curr Queue' in df.columns and not pd.isna(row['Curr Queue']) else ''
            
            # Create a new row in the required format
            new_row = {
                'S No.': int(serial_number),
                'Shipping Bill No.': shipping_bill_no if shipping_bill_no is not None else 0,
                'Shipping Bill Date': shipping_bill_date,
                'LEO Date': leo_date,
                'Amount': amount if amount is not None else 0,
                'Current Queue': current_queue
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            continue
    
    # Create the result DataFrame with data rows
    result_df = pd.DataFrame(result_data)
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    if not result_df.empty:
        print(f"  First few rows:")
        print(result_df.head())
    
    # Create the header rows exactly as in the output file format
    header_rows = pd.DataFrame([
        {'S No.': 'IEC Name-Alfa', 'Shipping Bill No.': '', 'Shipping Bill Date': '', 
         'LEO Date': '', 'Amount': '', 'Current Queue': ''},
        {'S No.': 'Location Name-All Locations', 'Shipping Bill No.': '', 'Shipping Bill Date': '', 
         'LEO Date': '', 'Amount': '', 'Current Queue': ''},
        {'S No.': 'Report Generated From-2021/07/01To2021/11/30', 'Shipping Bill No.': '', 'Shipping Bill Date': '', 
         'LEO Date': '', 'Amount': '', 'Current Queue': ''},
        {'S No.': 'Report Generated On-16/02/2022 12:48:50', 'Shipping Bill No.': '', 'Shipping Bill Date': '', 
         'LEO Date': '', 'Amount': '', 'Current Queue': ''},
        # Column headers (row 5)
        {'S No.': 'S No.', 'Shipping Bill No.': 'Shipping Bill No.', 'Shipping Bill Date': 'Shipping Bill Date',
         'LEO Date': 'LEO Date', 'Amount': 'Amount', 'Current Queue': 'Current Queue'}
    ])
    
    # Concatenate header rows with data rows
    final_df = pd.concat([header_rows, result_df], ignore_index=True)
    
    return final_df


def convert_brc(df: pd.DataFrame, brc_type: Optional[str] = None) -> pd.DataFrame:
    """Convert BRC Excel format after merging"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns for BRC: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    print(f"BRC Type selected: {brc_type}")
    
    # Load port code mapping
    port_mapping = load_port_code_mapping()
    print(f"Loaded {len(port_mapping)} port code mappings")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1 (for data rows, starting from row 4)
    serial_number = 1
    
    # Helper function to clean and convert to number
    def convert_to_number(value):
        if pd.isna(value):
            return ''
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it (remove commas, etc.)
            if isinstance(value, str):
                # Remove commas and any non-numeric except decimal point and minus
                value_str = str(value)
                value_str = value_str.replace(',', '').strip()
                if not value_str:
                    return ''
                
                try:
                    num_value = float(value_str)
                    # Convert to int if it's a whole number
                    return int(num_value) if num_value == int(num_value) else num_value
                except:
                    return ''
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return ''
    
    # Helper function to convert date to dd-mmm-yyyy format
    def convert_and_format_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            # Format as '07-Nov-2025'
            return date_value.strftime('%d-%b-%Y')
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # Try common date formats including dd-mm-yyyy
            date_formats = [
                '%d-%m-%Y',      # 07-11-2025
                '%d/%m/%Y',      # 07/11/2025
                '%d-%b-%Y',      # 07-Nov-2025
                '%d-%b-%y',      # 07-Nov-25
                '%Y-%m-%d',      # 2025-11-07
                '%d.%m.%Y',      # 07.11.2025
            ]
            
            # Try each format
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    # Format as '07-Nov-2025'
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    continue
            
            # Try with case-insensitive month names
            month_patterns = {
                'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
                'may': 'May', 'jun': 'Jun', 'jul': 'Jul', 'aug': 'Aug',
                'sep': 'Sep', 'oct': 'Oct', 'nov': 'Nov', 'dec': 'Dec'
            }
            
            for month_abbr, month_proper in month_patterns.items():
                if month_abbr in date_str.lower():
                    # Try to parse with the proper month abbreviation
                    try:
                        # Replace the month abbreviation in the string
                        date_str_modified = date_str.lower().replace(month_abbr, month_proper)
                        date_obj = pd.to_datetime(date_str_modified, format='%d-%b-%Y')
                        return date_obj.strftime('%d-%b-%Y')
                    except:
                        pass
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            # Excel serial date starts from 1899-12-30
            try:
                date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        return ''
    
    # Create column mapping based on the input file structure
    column_indices = {}
    
    # Debug: Print all columns with indices
    print("\n=== DEBUG: All columns in input DataFrame ===")
    for i, col in enumerate(df.columns):
        print(f"  Column {i} ('{col}'): Sample data = {df[col].iloc[0] if len(df) > 0 else 'N/A'}")
    
    # Try to find columns by exact name first
    expected_columns = {
        'BRC Number': 'brc_number',
        'BRC Date': 'brc_date',
        'BRC Status': 'brc_status',
        'Invoice Number': 'invoice_number',
        'SB NUMBER': 'sb_number',
        'PORT CODE': 'port_code',
        'SB Date': 'sb_date',
        'REALISED VALUE': 'realised_value',
        'CURRENCY': 'currency',
        'REALIZATION_DATE': 'realization_date',
        'BRC Utlisation Status': 'brc_utilisation'
    }
    
    # First pass: try exact column names
    for col_name in df.columns:
        if col_name in expected_columns:
            column_indices[expected_columns[col_name]] = list(df.columns).index(col_name)
            print(f"Found '{col_name}' at column {column_indices[expected_columns[col_name]]}")
    
    # Second pass: try case-insensitive matching
    if len(column_indices) < len(expected_columns):
        for col_name in df.columns:
            col_lower = str(col_name).lower()
            for expected, map_name in expected_columns.items():
                if map_name not in column_indices:
                    if expected.lower() in col_lower:
                        column_indices[map_name] = list(df.columns).index(col_name)
                        print(f"Found '{col_name}' (matches '{expected}') at column {column_indices[map_name]}")
    
    # Third pass: look specifically for port-related columns
    for i, col_name in enumerate(df.columns):
        col_lower = str(col_name).lower()
        if 'port' in col_lower:
            if 'port_code' not in column_indices:
                column_indices['port_code'] = i
                print(f"Found port column at index {i}: '{col_name}'")
    
    # If still not found, use default position 6 (column G/7)
    if 'port_code' not in column_indices and len(df.columns) > 6:
        column_indices['port_code'] = 6
        print(f"Using default port code column at index 6: '{df.columns[6] if 6 < len(df.columns) else 'N/A'}'")
    
    # Map other columns if not found
    other_mappings = {
        'brc_number': 0,
        'brc_date': 1,
        'brc_status': 2,
        'invoice_number': 7,
        'sb_number': 4,
        'sb_date': 5,
        'realised_value': 25,
        'currency': 26,
        'realization_date': 24,
        'brc_utilisation': 22
    }
    
    for map_name, default_pos in other_mappings.items():
        if map_name not in column_indices and default_pos < len(df.columns):
            column_indices[map_name] = default_pos
            print(f"Using positional mapping for {map_name}: column {default_pos} ('{df.columns[default_pos] if default_pos < len(df.columns) else 'N/A'}')")
    
    print(f"\nFinal column mapping: {column_indices}")
    
    # Debug: Show sample data from each mapped column
    if len(df) > 0:
        print("\n=== Sample data from mapped columns ===")
        for col_name, col_index in column_indices.items():
            if col_index < len(df.columns):
                sample_value = df.iloc[0, col_index] if not df.empty else 'N/A'
                print(f"  {col_name} (col {col_index}): {sample_value}")
    
    # Iterate through each row in the merged data
    for index, row in df.iterrows():
        try:
            # Skip rows that are completely empty
            if row.isna().all():
                continue
            
            # Skip header rows (rows that contain column headers)
            is_header = False
            header_keywords = ['brc number', 'brc date', 'brc status', 'invoice number', 
                              'sb number', 'port code', 'sb date', 'realised value', 
                              'currency', 'realization_date', 'brc utilisation status']
            
            for cell in row:
                if isinstance(cell, str):
                    cell_lower = cell.lower()
                    # Check if the cell contains any of the header keywords
                    if any(keyword in cell_lower for keyword in header_keywords):
                        is_header = True
                        break
            
            if is_header:
                print(f"Skipping header row {index}")
                continue
            
            # Extract values using the column indices
            def get_value(col_name, default=''):
                if col_name in column_indices:
                    col_idx = column_indices[col_name]
                    if col_idx < len(row):
                        value = row.iloc[col_idx]
                        return value if not pd.isna(value) else default
                return default
            
            brc_number = get_value('brc_number')
            brc_date = get_value('brc_date')
            brc_status = get_value('brc_status')
            invoice_number = get_value('invoice_number')
            sb_number = get_value('sb_number')
            port_name = get_value('port_code')  # This contains the port name/code
            sb_date = get_value('sb_date')
            realised_value = get_value('realised_value')
            currency = get_value('currency')
            realization_date = get_value('realization_date')
            brc_utilisation = get_value('brc_utilisation')
            
            # Debug for first few rows
            if index < 3:
                print(f"\nRow {index} extracted values:")
                print(f"  BRC Number: {brc_number}")
                print(f"  BRC Date: {brc_date}")
                print(f"  BRC Status: {brc_status}")
                print(f"  Invoice Number: {invoice_number}")
                print(f"  SB Number: {sb_number}")
                print(f"  PORT NAME/CODE: '{port_name}'")
                print(f"  SB Date: {sb_date}")
                print(f"  Realised Value: {realised_value}")
                print(f"  Currency: '{currency}'")
                print(f"  Realization Date: {realization_date}")
                print(f"  BRC Utlisation Status: '{brc_utilisation}'")
            
            # Validate we have at least some data
            if pd.isna(brc_number) and pd.isna(sb_number):
                print(f"Skipping row {index} - no data")
                continue
            
            # Convert port name to short form code using mapping
            port_code_clean = get_port_code(port_name, port_mapping)
            
            # Convert currency to 3-letter code
            currency_clean = get_currency_code(currency)
            
            # Debug port and currency conversion
            if index < 3:
                print(f"  Port code conversion: '{port_name}' -> '{port_code_clean}'")
                print(f"  Currency conversion: '{currency}' -> '{currency_clean}'")
            
            # Create a new row in the output format
            new_row = {
                'BRC Number': str(brc_number).strip() if brc_number is not None and not pd.isna(brc_number) else '',
                'BRC Date': convert_and_format_date(brc_date) if brc_date is not None and not pd.isna(brc_date) else '',
                'BRC Status': str(brc_status).strip() if brc_status is not None and not pd.isna(brc_status) else '',
                'Bill ID': convert_to_number(invoice_number) if invoice_number is not None and not pd.isna(invoice_number) else '',
                'SHB No': convert_to_number(sb_number) if sb_number is not None and not pd.isna(sb_number) else '',
                'SHB Port': port_code_clean,  # Short form port code
                'SHB Date': convert_and_format_date(sb_date) if sb_date is not None and not pd.isna(sb_date) else '',
                'Realised Value': convert_to_number(realised_value) if realised_value is not None and not pd.isna(realised_value) else '',
                'Currency': currency_clean,  # 3-letter currency code
                'Date of Realization': convert_and_format_date(realization_date) if realization_date is not None and not pd.isna(realization_date) else '',
                'BRC Utlisation Status': str(brc_utilisation).strip() if brc_utilisation is not None and not pd.isna(brc_utilisation) else '',
                'BRC Lot': ''
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    # Create the result DataFrame with data rows
    result_df = pd.DataFrame(result_data, columns=[
        'BRC Number', 'BRC Date', 'BRC Status', 'Bill ID', 'SHB No', 
        'SHB Port', 'SHB Date', 'Realised Value', 'Currency', 
        'Date of Realization', 'BRC Utlisation Status', 'BRC Lot'
    ])
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    print(f"  Total rows processed: {len(result_data)}")
    if not result_df.empty:
        print(f"  First 3 rows:")
        print(result_df.head(3))
        print(f"\n  SHB Port sample: {result_df['SHB Port'].head(5).tolist()}")
        print(f"  Currency sample: {result_df['Currency'].head(5).tolist()}")
    
    # Create the header rows to match the output file format
    # Based on "BRC_upload.xlsx" format - 3 rows: 2 empty rows + header row
    header_rows = pd.DataFrame([
        # Row 1: Empty
        {col: '' for col in result_df.columns},
        # Row 2: Empty
        {col: '' for col in result_df.columns},
        # Row 3: Column headers
        {
            'BRC Number': 'BRC Number',
            'BRC Date': 'BRC Date',
            'BRC Status': 'BRC Status',
            'Bill ID': 'Bill ID',
            'SHB No': 'SHB No',
            'SHB Port': 'SHB Port',
            'SHB Date': 'SHB Date',
            'Realised Value': 'Realised Value',
            'Currency': 'Currency',
            'Date of Realization': 'Date of Realization',
            'BRC Utlisation Status': 'BRC Utlisation Status',
            'BRC Lot': 'BRC Lot'
        }
    ])
    
    # Concatenate header rows with data rows
    final_df = pd.concat([header_rows, result_df], ignore_index=True)
    
    return final_df


def convert_irm(df: pd.DataFrame) -> pd.DataFrame:
    """Convert IRM Excel format"""
    if df.empty:
        return df
    
    # Add IRM specific conversion
    return df


def convert_igst_scroll(df: pd.DataFrame) -> pd.DataFrame:
    """Convert IGST Scroll Excel format"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns for IGST Scroll: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1 (for data rows, starting from row 7 in output)
    serial_number = 1
    
    # Helper function to clean and convert to number
    def convert_to_number(value):
        if pd.isna(value):
            return ''
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it (remove currency symbols, commas, etc.)
            if isinstance(value, str):
                # Remove currency symbols, commas, and any non-numeric except decimal point and minus
                value_str = str(value)
                # Remove ₹ symbol and commas
                value_str = value_str.replace('₹', '').replace(',', '').strip()
                if not value_str:
                    return ''
                
                try:
                    num_value = float(value_str)
                    # Convert to int if it's a whole number
                    return int(num_value) if num_value == int(num_value) else num_value
                except:
                    return ''
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return ''
    
    # Helper function to convert date to dd-mmm-yyyy format
    def convert_and_format_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            # Format as '10-Jul-2025' (with 4-digit year)
            return date_value.strftime('%d-%b-%Y')
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # Try common date formats including the ones in your example
            date_formats = [
                '%Y-%m-%d',      # 2025-07-10
                '%d-%b-%y',      # 10-Jul-25
                '%d-%b-%Y',      # 10-Jul-2025
                '%d/%m/%Y',      # 10/07/2025
                '%m/%d/%Y',      # 07/10/2025
                '%Y/%m/%d',      # 2025/07/10
                '%d-%m-%Y',      # 10-07-2025
                '%m-%d-%Y',      # 07-10-2025
                '%d.%m.%Y',      # 10.07.2025
            ]
            
            # Try each format
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    # Force 4-digit year output
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    continue
            
            # Try with case-insensitive month
            if any(month in date_str.upper() for month in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                                                          'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']):
                # Try with two-digit year first (most common case)
                try:
                    date_obj = pd.to_datetime(date_str, format='%d-%b-%y')
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    try:
                        date_obj = pd.to_datetime(date_str, format='%d-%b-%Y')
                        return date_obj.strftime('%d-%b-%Y')
                    except:
                        pass
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            # Excel serial date starts from 1899-12-30
            try:
                date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        return ''
    
    # Debug: Show first few rows of input
    print("\nFirst few rows of input DataFrame:")
    print(df.head(10))
    
    # Iterate through each row in the input data
    # Start from row 0 since pandas already read the header
    for index, row in df.iterrows():
        try:
            # Skip rows that are completely empty
            if row.isna().all():
                print(f"Skipping empty row {index}")
                continue
            
            # Debug: Show row data
            if index < 5:  # Show first 5 data rows for debugging
                print(f"\nProcessing row {index}:")
                for i, col in enumerate(df.columns):
                    print(f"  {col}: {row[col]}")
            
            # Extract values using column names (more reliable)
            shipping_bill_no = row['Shipping Bill No.'] if 'Shipping Bill No.' in df.columns else None
            shipping_bill_date = row['Shipping Bill Date'] if 'Shipping Bill Date' in df.columns else None
            igst_scroll_no = row['IGST Scroll No'] if 'IGST Scroll No' in df.columns else None
            igst_scroll_date = row['IGST Scroll Date'] if 'IGST Scroll Date' in df.columns else None
            scroll_amount = row['Scroll Amount(INR)'] if 'Scroll Amount(INR)' in df.columns else None
            status_pfms = row['Scroll Status At PFMS'] if 'Scroll Status At PFMS' in df.columns else None
            status_pao = row['Scroll Status At PAO'] if 'Scroll Status At PAO' in df.columns else None
            bank_response = row['Bank Response Code'] if 'Bank Response Code' in df.columns else None
            bank_transaction_id = row['Bank Transaction ID'] if 'Bank Transaction ID' in df.columns else None
            
            # Debug specific row for troubleshooting
            if index < 3:
                print(f"\nRow {index} extracted values:")
                print(f"  Shipping Bill No.: {shipping_bill_no}")
                print(f"  Shipping Bill Date: {shipping_bill_date}")
                print(f"  IGST Scroll No: {igst_scroll_no}")
                print(f"  IGST Scroll Date: {igst_scroll_date}")
                print(f"  Scroll Amount: {scroll_amount}")
                print(f"  Status PFMS: {status_pfms}")
                print(f"  Status PAO: {status_pao}")
                print(f"  Bank Response: {bank_response}")
                print(f"  Bank Transaction ID: {bank_transaction_id}")
            
            # Check if this is a header row by looking at column names
            # If the first column contains the header text, skip it
            if isinstance(shipping_bill_no, str) and 'Shipping Bill No.' in shipping_bill_no:
                print(f"Skipping header row {index}")
                continue
            
            # Validate we have at least some data
            if (pd.isna(shipping_bill_no) or shipping_bill_no == '') and (pd.isna(igst_scroll_no) or igst_scroll_no == ''):
                print(f"Skipping row {index} - no data")
                continue
            
            # Create a new row in the output format
            new_row = {
                'S No.': int(serial_number),
                'Shipping Bill No.': convert_to_number(shipping_bill_no) if shipping_bill_no is not None and not pd.isna(shipping_bill_no) else '',
                'Shipping Bill Date': convert_and_format_date(shipping_bill_date) if shipping_bill_date is not None and not pd.isna(shipping_bill_date) else '',
                'IGST Scroll Number': str(igst_scroll_no).strip() if igst_scroll_no is not None and not pd.isna(igst_scroll_no) else '',
                'IGST Scroll Date': convert_and_format_date(igst_scroll_date) if igst_scroll_date is not None and not pd.isna(igst_scroll_date) else '',
                'Scroll Amount(INR)': convert_to_number(scroll_amount) if scroll_amount is not None and not pd.isna(scroll_amount) else '',
                'Scroll Status At PFMS': str(status_pfms).strip() if status_pfms is not None and not pd.isna(status_pfms) else '',
                'Scroll Status At PAO': str(status_pao).strip() if status_pao is not None and not pd.isna(status_pao) else '',
                'Bank Response Code': str(bank_response).strip() if bank_response is not None and not pd.isna(bank_response) else '',
                'Bank Transaction Details': str(bank_transaction_id).strip() if bank_transaction_id is not None and not pd.isna(bank_transaction_id) else ''
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    # Create the result DataFrame with data rows
    result_df = pd.DataFrame(result_data, columns=[
        'S No.', 'Shipping Bill No.', 'Shipping Bill Date', 'IGST Scroll Number',
        'IGST Scroll Date', 'Scroll Amount(INR)', 'Scroll Status At PFMS',
        'Scroll Status At PAO', 'Bank Response Code', 'Bank Transaction Details'
    ])
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    print(f"  Total rows processed: {len(result_data)}")
    if not result_df.empty:
        print(f"  First few rows:")
        print(result_df.head())
    
    # Create the header rows to match the output file format
    # Based on "IGST Scroll uploding.xlsx" format - 6 header rows
    header_data = [
        # Row 1: IEC Name (only in first column)
        ['IEC Name  -  ALFA'] + [''] * 9,
        # Row 2: Location Name (only in first column)
        ['Location Name  -  NHAVA SHEVA SEA (INNSA1)'] + [''] * 9,
        # Row 3: Report Generated From (only in first column)
        ['Report Generated From -  2024/01/01  To  2024/06/30'] + [''] * 9,
        # Row 4: Report Generated On (only in first column)
        ['Report Generated On  -  30/07/2024 15:05:26'] + [''] * 9,
        # Row 5: Column Headers
        ['S No.', 'Shipping Bill No.', 'Shipping Bill Date', 'IGST Scroll Number',
         'IGST Scroll Date', 'Scroll Amount(INR)', 'Scroll Status', 'Scroll Status',
         'Bank Response Code', 'Bank Transaction Details'],
        # Row 6: Sub-headers (only for Scroll Status columns)
        [''] * 5 + ['', 'At PFMS', 'At PAO', '', '']
    ]
    
    # Create header DataFrame
    header_df = pd.DataFrame(header_data, columns=[
        'S No.', 'Shipping Bill No.', 'Shipping Bill Date', 'IGST Scroll Number',
        'IGST Scroll Date', 'Scroll Amount(INR)', 'Scroll Status At PFMS',
        'Scroll Status At PAO', 'Bank Response Code', 'Bank Transaction Details'
    ])
    
    # Concatenate header rows with data rows
    final_df = pd.concat([header_df, result_df], ignore_index=True)
    
    return final_df

def convert_rodtep_scroll(df: pd.DataFrame) -> pd.DataFrame:
    """Convert RODTEP Scroll Excel format"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns for RODTEP Scroll: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1 (for data rows, starting from row 4)
    serial_number = 1
    
    # Helper function to clean and convert to number
    def convert_to_number(value):
        if pd.isna(value):
            return ''
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it (remove currency symbols, commas, etc.)
            if isinstance(value, str):
                # Remove currency symbols, commas, and any non-numeric except decimal point and minus
                value_str = str(value)
                # Remove ₹ symbol and commas
                value_str = value_str.replace('₹', '').replace(',', '').strip()
                if not value_str:
                    return ''
                
                try:
                    num_value = float(value_str)
                    # Convert to int if it's a whole number
                    return int(num_value) if num_value == int(num_value) else num_value
                except:
                    return ''
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return ''
    
    # Helper function to convert date from dd.mm.yyyy format
    def convert_dot_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            return date_value.strftime('%d-%b-%Y')
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # First try the dd.mm.yyyy format explicitly
            if '.' in date_str:
                try:
                    parts = date_str.split('.')
                    if len(parts) == 3:
                        day, month, year = parts
                        # Ensure 4-digit year
                        if len(year) == 2:
                            year = '20' + year
                        date_obj = datetime.datetime(int(year), int(month), int(day))
                        return date_obj.strftime('%d-%b-%Y')
                except:
                    pass
            
            # Try other common formats
            date_formats = [
                '%d.%m.%Y',      # 27.10.2025
                '%d-%m-%Y',      # 27-10-2025
                '%d/%m/%Y',      # 27/10/2025
                '%d-%b-%Y',      # 27-Oct-2025
                '%d-%b-%y',      # 27-Oct-25
                '%Y-%m-%d',      # 2025-10-27
            ]
            
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    continue
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            try:
                # Excel serial date starts from 1899-12-30
                if date_value > 0:
                    date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                    return date_obj.strftime('%d-%b-%Y')
                else:
                    return ''
            except:
                return ''
        
        return ''
    
    # Check the column names and map them
    print("\nDebug: Column names in input DataFrame:")
    for i, col in enumerate(df.columns):
        print(f"  Column {i}: '{col}'")
    
    # Try to find the correct columns by name or position
    column_map = {}
    
    # Look for column names in the DataFrame
    for col_name in df.columns:
        col_lower = str(col_name).lower()
        
        if 'sb number' in col_lower or 'shipping bill' in col_lower:
            column_map['sb_number'] = col_name
        elif 'sb date' in col_lower or 'shipping bill date' in col_lower:
            column_map['sb_date'] = col_name
        elif 'scroll number' in col_lower:
            column_map['scroll_number'] = col_name
        elif 'scroll date' in col_lower:
            column_map['scroll_date'] = col_name
        elif 'location' in col_lower:
            column_map['location'] = col_name
        elif 'sanctioned' in col_lower or 'amount' in col_lower:
            column_map['amount'] = col_name
    
    print(f"\nColumn mapping found: {column_map}")
    
    # If column mapping failed, use positional mapping
    if not column_map:
        print("Using positional mapping...")
        if len(df.columns) >= 7:
            column_map = {
                'sb_number': df.columns[0],
                'sb_date': df.columns[1],
                'scroll_number': df.columns[2],
                'scroll_date': df.columns[3],
                'location': df.columns[5],
                'amount': df.columns[6]
            }
    
    # Iterate through each row in the input data
    for index, row in df.iterrows():
        try:
            # Skip rows that are completely empty or header rows
            if row.isna().all():
                continue
            
            # Check if this is a header row
            if index == 0:
                # Check if any cell contains header-like text
                header_keywords = ['sb', 'scroll', 'date', 'number', 'location', 'amount']
                is_header = False
                for cell in row:
                    if isinstance(cell, str):
                        cell_lower = cell.lower()
                        if any(keyword in cell_lower for keyword in header_keywords):
                            is_header = True
                            break
                if is_header:
                    print(f"Skipping header row {index}")
                    continue
            
            # Extract values using the column map
            sb_number = row[column_map.get('sb_number')] if 'sb_number' in column_map else None
            sb_date = row[column_map.get('sb_date')] if 'sb_date' in column_map else None
            scroll_number = row[column_map.get('scroll_number')] if 'scroll_number' in column_map else None
            scroll_date = row[column_map.get('scroll_date')] if 'scroll_date' in column_map else None
            location = row[column_map.get('location')] if 'location' in column_map else None
            amount = row[column_map.get('amount')] if 'amount' in column_map else None
            
            # Debug for first few rows
            if index < 3:
                print(f"\nRow {index} values:")
                print(f"  SB Number: {sb_number}")
                print(f"  SB Date: {sb_date}")
                print(f"  Scroll Number: {scroll_number}")
                print(f"  Scroll Date: {scroll_date}")
                print(f"  Location: {location}")
                print(f"  Amount: {amount}")
            
            # Validate we have at least some data
            if pd.isna(sb_number) and pd.isna(scroll_number):
                print(f"Skipping row {index} - no data")
                continue
            
            # Create a new row in the output format
            new_row = {
                'Sr. No.': int(serial_number),
                'SHB No': convert_to_number(sb_number),
                'Date': convert_dot_date(sb_date),
                'Scroll No': convert_to_number(scroll_number),
                'Scroll Date': convert_dot_date(scroll_date),
                'Scroll Amt': convert_to_number(amount),
                'Port': str(location) if location is not None and not pd.isna(location) else ''
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    # Create the result DataFrame with data rows
    result_df = pd.DataFrame(result_data)
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    if not result_df.empty:
        print(f"  First few rows:")
        print(result_df.head())
    
    # If result is empty, return empty DataFrame with headers
    if result_df.empty:
        print("Warning: No data was processed. Returning empty DataFrame with headers.")
    
    # Create the header rows to match the output file format
    # Based on "RoDTEP Scroll Uploading.xlsx" format - 3 header rows
    header_rows = pd.DataFrame([
        # Row 1: Empty
        {'Sr. No.': '', 'SHB No': '', 'Date': '', 'Scroll No': '', 'Scroll Date': '', 'Scroll Amt': '', 'Port': ''},
        # Row 2: Empty
        {'Sr. No.': '', 'SHB No': '', 'Date': '', 'Scroll No': '', 'Scroll Date': '', 'Scroll Amt': '', 'Port': ''},
        # Row 3: Column headers
        {'Sr. No.': 'Sr. No.', 'SHB No': 'SHB No', 'Date': 'Date', 'Scroll No': 'Scroll No', 
         'Scroll Date': 'Scroll Date', 'Scroll Amt': 'Scroll Amt', 'Port': 'Port'}
    ])
    
    # Concatenate header rows with data rows
    final_df = pd.concat([header_rows, result_df], ignore_index=True)
    
    return final_df

def convert_rodtep_scrip(df: pd.DataFrame) -> pd.DataFrame:
    """Convert RODTEP Scrip Excel format"""
    if df.empty:
        return df
    
    print(f"Input DataFrame columns for RODTEP Scrip: {df.columns.tolist()}")
    print(f"Input DataFrame shape: {df.shape}")
    
    # Create a new DataFrame with the required format
    result_data = []
    
    # Start serial number from 1 (for data rows, starting from row 4 in output)
    serial_number = 1
    
    # Helper function to clean and convert to number (for amounts) - remove currency symbols and commas
    def convert_to_number(value):
        if pd.isna(value):
            return ''
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value) if value == int(value) else value
            
            # If string, clean it (remove currency symbols, commas, etc.)
            if isinstance(value, str):
                value_str = str(value).strip()
                if not value_str or value_str.upper() == 'N.A' or value_str.upper() == 'NA':
                    return ''
                
                # Remove currency symbols like ₹, $, €, £, etc.
                value_str = re.sub(r'[₹$€£¥]', '', value_str)
                
                # Remove commas
                value_str = value_str.replace(',', '').strip()
                
                # Remove any other non-numeric characters except decimal point and minus
                value_str = re.sub(r'[^\d.-]', '', value_str)
                
                if not value_str:
                    return ''
                
                try:
                    num_value = float(value_str)
                    # Convert to int if it's a whole number
                    return int(num_value) if num_value == int(num_value) else num_value
                except:
                    return ''
            
            return pd.to_numeric(value, errors='coerce')
        except Exception as e:
            print(f"Error converting number '{value}': {e}")
            return ''
    
    # Helper function to clean and convert to number for non-amount fields (like SB NUMBER, SCRIP NUMBER)
    def convert_to_integer(value):
        if pd.isna(value):
            return ''
        try:
            # If already a number
            if isinstance(value, (int, float)):
                return int(value)
            
            # If string, clean it (remove non-digit characters)
            if isinstance(value, str):
                value_str = str(value).strip()
                if not value_str:
                    return ''
                
                # Remove any non-digit characters
                value_str = re.sub(r'[^\d]', '', value_str)
                
                if not value_str:
                    return ''
                
                try:
                    return int(value_str)
                except:
                    return ''
            
            return int(value) if not pd.isna(value) else ''
        except Exception as e:
            print(f"Error converting to integer '{value}': {e}")
            return ''
    
    # Helper function to convert date to dd-mmm-yyyy format, handling "N.A" values
    def convert_and_format_date(date_value):
        if pd.isna(date_value):
            return ''
        
        # Check if it's "N.A" or "NA" (case-insensitive)
        if isinstance(date_value, str):
            date_str = str(date_value).strip().upper()
            if date_str == 'N.A' or date_str == 'NA':
                return date_str  # Return as-is for SCRIP TRANSFER DATE
        
        # If it's already a datetime object
        if isinstance(date_value, (pd.Timestamp, datetime.datetime)):
            # Format as '09-Jul-2025'
            return date_value.strftime('%d-%b-%Y')
        
        # If it's a string
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str:
                return ''
            
            # Skip if it's "N.A" or "NA" (already handled above)
            if date_str.upper() in ['N.A', 'NA']:
                return date_str.upper()
            
            # Try date formats including dd.mm.yyyy (from your input example)
            date_formats = [
                '%d.%m.%Y',      # 22.07.2025 (from your input)
                '%d-%b-%Y',      # 09-Jul-2025
                '%d-%b-%y',      # 09-Jul-25
                '%Y-%m-%d',      # 2025-07-09
                '%d/%m/%Y',      # 09/07/2025
                '%m/%d/%Y',      # 07/09/2025
                '%Y/%m/%d',      # 2025/07/09
                '%d-%m-%Y',      # 09-07-2025
                '%m-%d-%Y',      # 07-09-2025
            ]
            
            # Try each format
            for fmt in date_formats:
                try:
                    date_obj = pd.to_datetime(date_str, format=fmt)
                    # Format as '09-Jul-2025'
                    return date_obj.strftime('%d-%b-%Y')
                except:
                    continue
            
            # Try with case-insensitive month names
            month_patterns = {
                'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
                'may': 'May', 'jun': 'Jun', 'jul': 'Jul', 'aug': 'Aug',
                'sep': 'Sep', 'oct': 'Oct', 'nov': 'Nov', 'dec': 'Dec'
            }
            
            for month_abbr, month_proper in month_patterns.items():
                if month_abbr in date_str.lower():
                    # Try to parse with the proper month abbreviation
                    try:
                        # Replace the month abbreviation in the string
                        date_str_modified = date_str.lower().replace(month_abbr, month_proper)
                        date_obj = pd.to_datetime(date_str_modified, format='%d-%b-%Y')
                        return date_obj.strftime('%d-%b-%Y')
                    except:
                        pass
            
            # Last resort: pandas flexible parser
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.isna(date_obj):
                    return ''
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        # If it's a number (Excel serial date), convert it
        if isinstance(date_value, (int, float)):
            # Excel serial date starts from 1899-12-30
            try:
                date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_value)
                return date_obj.strftime('%d-%b-%Y')
            except:
                return ''
        
        return ''
    
    # Debug: Show first few rows of input
    print("\nFirst few rows of input DataFrame:")
    print(df.head(10))
    
    # Check the column names and map them
    print("\nDebug: Column names in input DataFrame:")
    for i, col in enumerate(df.columns):
        print(f"  Column {i}: '{col}'")
    
    # Map column names from input to variables for clarity
    # Based on your input file structure:
    # A: Scrip No, B: Scrip Issue Date, C: Scrip Exp Date, D: Scrip Issued Amount
    # E: Scrip Balance, F: Scrip Transfer Date, G: Scrip Status, H: Scroll Number, I: SB Number
    
    column_map = {}
    for col_name in df.columns:
        col_lower = str(col_name).lower()
        
        if 'scrip no' in col_lower:
            column_map['scrip_no'] = col_name
        elif 'scrip issue date' in col_lower:
            column_map['scrip_issue_date'] = col_name
        elif 'scrip exp date' in col_lower or 'scrip expiry date' in col_lower:
            column_map['scrip_exp_date'] = col_name
        elif 'scrip issued amount' in col_lower:
            column_map['scrip_issued_amount'] = col_name
        elif 'scrip balance' in col_lower:
            column_map['scrip_balance'] = col_name
        elif 'scrip transfer date' in col_lower:
            column_map['scrip_transfer_date'] = col_name
        elif 'scrip status' in col_lower:
            column_map['scrip_status'] = col_name
        elif 'scroll number' in col_lower:
            column_map['scroll_number'] = col_name
        elif 'sb number' in col_lower or 'shipping bill' in col_lower:
            column_map['sb_number'] = col_name
    
    print(f"\nColumn mapping found: {column_map}")
    
    # If column mapping failed, use positional mapping (based on your input file)
    if not column_map:
        print("Using positional mapping...")
        if len(df.columns) >= 9:
            column_map = {
                'scrip_no': df.columns[0],
                'scrip_issue_date': df.columns[1],
                'scrip_exp_date': df.columns[2],
                'scrip_issued_amount': df.columns[3],
                'scrip_balance': df.columns[4],
                'scrip_transfer_date': df.columns[5],
                'scrip_status': df.columns[6],
                'scroll_number': df.columns[7],
                'sb_number': df.columns[8]
            }
    
    # Iterate through each row in the input data
    for index, row in df.iterrows():
        try:
            # Skip rows that are completely empty
            if row.isna().all():
                continue
            
            # Check if this is a header row
            if index == 0:
                # Check if any cell contains header-like text
                header_keywords = ['scrip', 'scroll', 'sb', 'date', 'number', 'amount', 'balance', 'status']
                is_header = False
                for cell in row:
                    if isinstance(cell, str):
                        cell_lower = cell.lower()
                        if any(keyword in cell_lower for keyword in header_keywords):
                            is_header = True
                            break
                if is_header:
                    print(f"Skipping header row {index}")
                    continue
            
            # Extract values using the column map
            # Note: We'll still extract scroll_number for debugging, but output column 2 will be blank
            scroll_number = row[column_map.get('scroll_number')] if 'scroll_number' in column_map else None
            sb_number = row[column_map.get('sb_number')] if 'sb_number' in column_map else None
            scrip_no = row[column_map.get('scrip_no')] if 'scrip_no' in column_map else None
            scrip_issue_date = row[column_map.get('scrip_issue_date')] if 'scrip_issue_date' in column_map else None
            scrip_exp_date = row[column_map.get('scrip_exp_date')] if 'scrip_exp_date' in column_map else None
            scrip_issued_amount = row[column_map.get('scrip_issued_amount')] if 'scrip_issued_amount' in column_map else None
            scrip_balance = row[column_map.get('scrip_balance')] if 'scrip_balance' in column_map else None
            scrip_transfer_date = row[column_map.get('scrip_transfer_date')] if 'scrip_transfer_date' in column_map else None
            scrip_status = row[column_map.get('scrip_status')] if 'scrip_status' in column_map else None
            
            # Debug for first few rows
            if index < 3:
                print(f"\nRow {index} values:")
                print(f"  Scroll Number (input col 8): {scroll_number} (will be blank in output)")
                print(f"  SB Number (input col 9): {sb_number}")
                print(f"  Scrip No (input col 1): {scrip_no}")
                print(f"  Scrip Issue Date (input col 2): {scrip_issue_date}")
                print(f"  Scrip Exp Date (input col 3): {scrip_exp_date}")
                print(f"  Scrip Issued Amount (input col 4): '{scrip_issued_amount}'")
                print(f"  Scrip Balance (input col 5): '{scrip_balance}'")
                print(f"  Scrip Transfer Date (input col 6): '{scrip_transfer_date}'")
                print(f"  Scrip Status (input col 7): {scrip_status}")
            
            # Validate we have at least some data
            if pd.isna(scrip_no) and pd.isna(scroll_number) and pd.isna(sb_number):
                print(f"Skipping row {index} - no data")
                continue
            
            # Special handling for SCRIP TRANSFER DATE - preserve "N.A" if present
            scrip_transfer_date_formatted = ''
            if scrip_transfer_date is not None and not pd.isna(scrip_transfer_date):
                scrip_transfer_date_str = str(scrip_transfer_date).strip().upper()
                if scrip_transfer_date_str in ['N.A', 'NA']:
                    scrip_transfer_date_formatted = scrip_transfer_date_str
                else:
                    scrip_transfer_date_formatted = convert_and_format_date(scrip_transfer_date)
            
            # Create a new row in the output format according to your requirements
            # Output column 2 (SCROLL NUMBER) is now blank as requested
            new_row = {
                'Sr. No': int(serial_number),
                'SCROLL NUMBER': '',  # BLANK as per new requirement (output column 2)
                'SB NUMBER': convert_to_integer(sb_number) if sb_number is not None and not pd.isna(sb_number) else '',
                'SB DATE': '',  # Blank as per requirements
                'SB AMOUNT': '',  # Blank as per requirements
                'SCRIP NUMBER': convert_to_integer(scrip_no) if scrip_no is not None and not pd.isna(scrip_no) else '',
                'SCRIP ISSUE DATE': convert_and_format_date(scrip_issue_date) if scrip_issue_date is not None and not pd.isna(scrip_issue_date) else '',
                'SCRIP EXPIRY DATE': convert_and_format_date(scrip_exp_date) if scrip_exp_date is not None and not pd.isna(scrip_exp_date) else '',
                'SCRIP ISSUE AMOUNT': convert_to_number(scrip_issued_amount) if scrip_issued_amount is not None and not pd.isna(scrip_issued_amount) else '',
                'SCRIP BALANCE AMOUNT': convert_to_number(scrip_balance) if scrip_balance is not None and not pd.isna(scrip_balance) else '',
                'SCRIP TRANSFER DATE': scrip_transfer_date_formatted,
                'SCRIP STATUS': str(scrip_status).strip() if scrip_status is not None and not pd.isna(scrip_status) else '',
                'Application Ref. No': ''  # Blank as per requirements
            }
            result_data.append(new_row)
            serial_number += 1
            
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    # Create the result DataFrame with data rows
    result_df = pd.DataFrame(result_data, columns=[
        'Sr. No', 'SCROLL NUMBER', 'SB NUMBER', 'SB DATE', 'SB AMOUNT',
        'SCRIP NUMBER', 'SCRIP ISSUE DATE', 'SCRIP EXPIRY DATE',
        'SCRIP ISSUE AMOUNT', 'SCRIP BALANCE AMOUNT',
        'SCRIP TRANSFER DATE', 'SCRIP STATUS', 'Application Ref. No'
    ])
    
    print(f"\nOutput DataFrame info:")
    print(f"  Shape: {result_df.shape}")
    print(f"  Total rows processed: {len(result_data)}")
    if not result_df.empty:
        print(f"  First few rows:")
        print(result_df.head())
        print(f"\n  Sample amounts (after cleaning):")
        print(f"    SCRIP ISSUE AMOUNT: {result_df['SCRIP ISSUE AMOUNT'].head(3).tolist()}")
        print(f"    SCRIP BALANCE AMOUNT: {result_df['SCRIP BALANCE AMOUNT'].head(3).tolist()}")
        print(f"    SCRIP TRANSFER DATE: {result_df['SCRIP TRANSFER DATE'].head(3).tolist()}")
        print(f"    SCROLL NUMBER (should be blank): {result_df['SCROLL NUMBER'].head(3).tolist()}")
    
    # Create the header rows to match the output file format
    # Row 1: Empty
    # Row 2: Empty  
    # Row 3: Column headers (headers at row 3 as requested)
    # Row 4 onwards: Data (data from row 4 as requested)
    
    header_rows = pd.DataFrame([
        # Row 1: Empty
        {col: '' for col in result_df.columns},
        # Row 2: Empty
        {col: '' for col in result_df.columns},
        # Row 3: Column headers (headers at row 3)
        {
            'Sr. No': 'Sr. No',
            'SCROLL NUMBER': 'SCROLL NUMBER',
            'SB NUMBER': 'SB NUMBER',
            'SB DATE': 'SB DATE',
            'SB AMOUNT': 'SB AMOUNT',
            'SCRIP NUMBER': 'SCRIP NUMBER',
            'SCRIP ISSUE DATE': 'SCRIP ISSUE DATE',
            'SCRIP EXPIRY DATE': 'SCRIP EXPIRY DATE',
            'SCRIP ISSUE AMOUNT': 'SCRIP ISSUE AMOUNT',
            'SCRIP BALANCE AMOUNT': 'SCRIP BALANCE AMOUNT',
            'SCRIP TRANSFER DATE': 'SCRIP TRANSFER DATE',
            'SCRIP STATUS': 'SCRIP STATUS',
            'Application Ref. No': 'Application Ref. No'
        }
    ])
    
    # Concatenate header rows with data rows
    final_df = pd.concat([header_rows, result_df], ignore_index=True)
    
    return final_df

def process_excel(process_type: str, files: List[Tuple[str, bytes]], brc_type: Optional[str] = None) -> pd.DataFrame:
    """
    Main processing function
    Returns: Processed DataFrame
    """
    
    # Processes that require merging first
    if process_type in ['dbk_disbursement', 'dbk_pendency', 'brc']:
        # Merge files first
        merged_df = merge_excel_files(files)
        
        # Then apply conversion
        if process_type == 'dbk_disbursement':
            return convert_dbk_disbursement(merged_df)
        elif process_type == 'dbk_pendency':
            return convert_dbk_pendency(merged_df)
        else:  # brc
            # Pass brc_type to convert_brc function
            return convert_brc(merged_df, brc_type)
    
    else:
        # For other processes, only one file expected
        if not files:
            raise ValueError("No file provided")
        
        # Get the single file
        filename, file_data = files[0]
        
        # Determine engine based on file extension
        if filename.lower().endswith('.xls'):
            # For .xls files, try xlrd engine
            try:
                df = pd.read_excel(io.BytesIO(file_data), engine='xlrd')
            except ImportError:
                raise ImportError(
                    "Missing dependency: xlrd library is required to read .xls files. "
                    "Please install it using: pip install xlrd"
                )
        elif filename.lower().endswith('.xlsx'):
            # For .xlsx files, use openpyxl engine
            df = pd.read_excel(io.BytesIO(file_data), engine='openpyxl')
        else:
            # Try default engine
            df = pd.read_excel(io.BytesIO(file_data))
        
        # Apply conversion based on process type
        converters = {
            'irm': convert_irm,
            'igst_scroll': convert_igst_scroll,
            'rodtep_scroll': convert_rodtep_scroll,
            'rodtep_scrip': convert_rodtep_scrip,
        }
        
        converter = converters.get(process_type)
        if converter:
            return converter(df)
        else:
            return df


def get_process_display_name(process_type: str) -> str:
    """Get display name for process type"""
    names = {
        'dbk_disbursement': 'DBK Disbursement',
        'dbk_pendency': 'DBK Pendency',
        'brc': 'BRC',
        'irm': 'IRM',
        'igst_scroll': 'IGST Scroll',
        'rodtep_scroll': 'RODTEP Scroll',
        'rodtep_scrip': 'RODTEP Scrip',
    }
    return names.get(process_type, process_type)


def get_process_filename(process_type: str) -> str:
    """Get filename for process type"""
    names = {
        'dbk_disbursement': 'DBK_Disbursement',
        'dbk_pendency': 'DBK_Pendency',
        'brc': 'BRC',
        'irm': 'IRM',
        'igst_scroll': 'IGST_Scroll',
        'rodtep_scroll': 'RODTEP_Scroll',
        'rodtep_scrip': 'RODTEP_Scrip',
    }
    return names.get(process_type, 'Processed')