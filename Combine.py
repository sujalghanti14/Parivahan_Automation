import pandas as pd
import numpy as np
import re
import os
from pathlib import Path

def process_multiple_car_manufacturer_data():
    """
    Process multiple Excel files from a folder and compile them into one sheet.
    Transform car manufacturer fuel data from wide format to long (tidy) format.
    Includes location information extracted from A1 cell of each file.
    """
    
    # Step 1: Define folder path containing Excel files
    folder_path = r"C:\Users\sujal\Documents\Raam Group\Parivahan\Raw Data\2W\2W INDIA 2025\Sept"
    
    print("📁 Processing multiple Excel files from folder...")
    print(f"Folder path: {folder_path}")
    
    # Get all Excel files in the folder
    excel_files = []
    for file_path in Path(folder_path).glob("*.xlsx"):
        if not file_path.name.startswith('~'):  # Skip temporary Excel files
            excel_files.append(file_path)
    
    for file_path in Path(folder_path).glob("*.xls"):
        if not file_path.name.startswith('~'):  # Skip temporary Excel files
            excel_files.append(file_path)
    
    print(f"📊 Found {len(excel_files)} Excel files to process:")
    for i, file_path in enumerate(excel_files, 1):
        print(f"  {i}. {file_path.name}")
    
    if not excel_files:
        print("❌ No Excel files found in the specified folder!")
        return None
    
    # Step 2: Process each file and collect all data
    all_data = []
    successful_files = 0
    failed_files = []
    
    print(f"\n{'='*80}")
    print("🔄 PROCESSING FILES...")
    print(f"{'='*80}")
    
    for i, file_path in enumerate(excel_files, 1):
        print(f"\n📄 Processing file {i}/{len(excel_files)}: {file_path.name}")

        try:
            # Process individual file
            df_processed = process_single_file(str(file_path))
            
            if df_processed is not None and len(df_processed) > 0:
                # Add source file information
                df_processed['Source_File'] = file_path.name
                all_data.append(df_processed)
                successful_files += 1
                print(f"✅ Successfully processed: {len(df_processed)} records")
            else:
                print(f"⚠️ No data found in file: {file_path.name}")
                failed_files.append(file_path.name)
                
        except Exception as e:
            print(f"❌ Error processing {file_path.name}: {str(e)}")
            failed_files.append(file_path.name)
            continue
    
    # Step 3: Combine all processed data
    if not all_data:
        print("❌ No data was successfully processed from any file!")
        return None
    
    print(f"\n🔗 Combining data from {len(all_data)} files...")
    
    # Concatenate all dataframes
    df_combined = pd.concat(all_data, ignore_index=True)
    
    # Reorder columns for better readability
    column_order = ['Maker', 'Fuel Type', 'Count', 'Location', 'Month_Year', 'Source_File']
    df_combined = df_combined[column_order]
    
    print(f"✅ Data combination complete!")
    print(f"Combined data shape: {df_combined.shape[0]} rows × {df_combined.shape[1]} columns")
    
    # Step 4: Save combined data to Excel file
    output_path = r'C:\Users\sujal\Documents\Raam Group\Parivahan\Scrap\combined_car_data_all_locations.xlsx'
    
    print(f"\n💾 Saving combined data to Excel file...")
    
    # Save with proper formatting
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='Combined_Car_Data', index=False)
        
        # Get workbook and worksheet for formatting
        workbook = writer.book
        worksheet = writer.sheets['Combined_Car_Data']
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Step 5: Display final results and summary
    print(f"\n{'='*90}")
    print(f"🎉 BATCH PROCESSING COMPLETE!")
    print(f"{'='*90}")
    print(f"📊 FINAL SUMMARY:")
    print(f"   • Total files processed: {successful_files}/{len(excel_files)}")
    print(f"   • Output file: combined_car_data_all_locations.xlsx")
    print(f"   • Total manufacturers: {df_combined['Maker'].nunique():,}")
    print(f"   • Total fuel types: {df_combined['Fuel Type'].nunique():,}")
    print(f"   • Total records (non-zero): {len(df_combined):,}")
    print(f"   • Total vehicle count: {df_combined['Count'].sum():,}")
    print(f"   • Unique locations: {df_combined['Location'].nunique():,}")
    
    if failed_files:
        print(f"\n⚠️ FAILED FILES ({len(failed_files)}):")
        for file in failed_files:
            print(f"   • {file}")
    
    print(f"{'='*90}")
    
    # Show sample of final combined output
    print(f"\n📋 SAMPLE OF COMBINED OUTPUT:")
    print(f"{'Maker':<25} | {'Fuel Type':<15} | {'Count':>6} | {'Month_Year':<10} | {'Source File':<20}")
    print(f"{'-'*25} | {'-'*15} | {'-'*6} | {'-'*10} | {'-'*20}")
    
    for _, row in df_combined.head(15).iterrows():
        maker_name = str(row['Maker'])[:24] if pd.notna(row['Maker']) else "Unknown"
        fuel_type = str(row['Fuel Type'])[:14] if pd.notna(row['Fuel Type']) else "Unknown"
        count = int(row['Count']) if pd.notna(row['Count']) else 0
        month_year = str(row['Month_Year'])[:9] if pd.notna(row['Month_Year']) else "Unknown"
        source_file = str(row['Source_File'])[:19] if pd.notna(row['Source_File']) else "Unknown"
        print(f"{maker_name:<25} | {fuel_type:<15} | {count:>6,} | {month_year:<10} | {source_file:<20}")
    
    if len(df_combined) > 15:
        print(f"{'...':<25} | {'...':<15} | {'...':>6} | {'...':<10} | {'...':<20}")
    
    # Show top fuel types across all locations
    print(f"\n🏆 TOP FUEL TYPES ACROSS ALL LOCATIONS:")
    fuel_totals = df_combined.groupby('Fuel Type')['Count'].sum().sort_values(ascending=False)
    for fuel_type, total_count in fuel_totals.head(10).items():
        print(f"   {fuel_type:<25}: {total_count:>10,}")
    
    # Show data by location
    print(f"\n🗺️ DATA BY LOCATION:")
    location_summary = df_combined.groupby('Location').agg({
        'Count': 'sum',
        'Maker': 'nunique',
        'Fuel Type': 'nunique'
    }).sort_values('Count', ascending=False)
    
    print(f"{'Location':<50} | {'Total Count':>12} | {'Makers':>8} | {'Fuel Types':>11}")
    print(f"{'-'*50} | {'-'*12} | {'-'*8} | {'-'*11}")
    
    for location, row in location_summary.head(10).iterrows():
        location_short = str(location)[:49] if pd.notna(location) else "Unknown"
        print(f"{location_short:<50} | {int(row['Count']):>12,} | {int(row['Maker']):>8} | {int(row['Fuel Type']):>11}")
    
    print(f"\n✅ Combined Excel file saved successfully at:")
    print(f"   {output_path}")
    
    return df_combined

def process_single_file(file_path):
    """
    Process a single Excel file and return processed data.
    """
    try:
        print(f"   📖 Reading file: {os.path.basename(file_path)}")
        
        # Read A1 cell to extract location information
        df_header = pd.read_excel(file_path, nrows=1, header=None)
        a1_content = str(df_header.iloc[0, 0]) if not df_header.empty else ""
        
        # Extract location information
        location_info = extract_location_info(a1_content)
        print(f"   📍 Location: {location_info['full_location'][:60]}...")
        
        # Read Excel file, skipping first 3 rows to get to the actual data
        df = pd.read_excel(file_path, skiprows=3)
        
        if df.empty:
            print(f"   ⚠️ Empty dataframe after skipping rows")
            return None
        
        # Fix column naming - manufacturer names are in the second column
        column_names = ['S_No', 'Maker'] + list(df.columns[2:])
        df.columns = column_names
        
        # Identify fuel type columns (all columns except 'S_No' and 'Maker')
        fuel_type_columns = df.columns[2:].tolist()
        
        # Remove the S_No column as we don't need it
        df = df.drop('S_No', axis=1)
        
        # Clean all fuel type columns
        def clean_count_value(value):
            if pd.isna(value) or value == '' or value == '-':
                return 0
            try:
                cleaned_value = str(value).replace(',', '').strip()
                return int(float(cleaned_value))
            except (ValueError, TypeError):
                return 0
        
        # Apply cleaning function to all fuel type columns
        for col in fuel_type_columns:
            df[col] = df[col].apply(clean_count_value)
        
        # Transform from wide format to long format
        df_long = pd.melt(
            df,
            id_vars=['Maker'],
            value_vars=fuel_type_columns,
            var_name='Fuel Type',
            value_name='Count'
        )
        
        # Add location information
        df_long['Location'] = location_info['full_location']
        df_long['Month_Year'] = location_info['month_year']
        
        # Remove rows where Count is zero
        df_final = df_long[df_long['Count'] > 0].copy()
        
        print(f"   ✅ Processed {len(df_final)} non-zero records")
        
        return df_final
        
    except Exception as e:
        print(f"   ❌ Error in process_single_file: {str(e)}")
        return None

def extract_location_info(a1_content):
    """
    Extract location information from A1 cell content.
    Expected format: "Maker Wise Fuel Data of DHULE - MH18, Maharashtra (JUL,2025)"
    
    Returns dictionary with extracted information.
    """
    location_info = {
        'full_location': 'Unknown',
        'city': 'Unknown',
        'rto_code': 'Unknown',
        'state': 'Unknown',
        'month_year': 'Unknown'
    }
    
    try:
        # Store the full content
        location_info['full_location'] = a1_content
        
        # Extract city and RTO code using regex
        city_rto_pattern = r'of\s+([A-Z\s]+)\s*-\s*([A-Z0-9]+)\s*,\s*([A-Z\s]+)'
        city_rto_match = re.search(city_rto_pattern, a1_content.upper())
        
        if city_rto_match:
            location_info['city'] = city_rto_match.group(1).strip()
            location_info['rto_code'] = city_rto_match.group(2).strip()
            location_info['state'] = city_rto_match.group(3).strip()
        
        # Extract month and year using regex
        month_year_pattern = r'\(([A-Z]{3})[,\s]+(\d{4})\)'
        month_year_match = re.search(month_year_pattern, a1_content.upper())
        
        if month_year_match:
            month = month_year_match.group(1)
            year = month_year_match.group(2)
            location_info['month_year'] = f"{month},{year}"
        
        # If no specific patterns found, try to extract basic info
        if location_info['city'] == 'Unknown':
            alt_pattern = r'of\s+([A-Z\s\-0-9]+)'
            alt_match = re.search(alt_pattern, a1_content.upper())
            if alt_match:
                extracted = alt_match.group(1).strip()
                if '-' in extracted:
                    parts = extracted.split('-')
                    location_info['city'] = parts[0].strip()
                    if len(parts) > 1:
                        location_info['rto_code'] = parts[1].strip().split(',')[0].strip()
    
    except Exception as e:
        print(f"⚠️ Warning: Could not extract location info from '{a1_content}'. Error: {e}")
    
    return location_info

# Run the batch processing function
if __name__ == "__main__":
    combined_data = process_multiple_car_manufacturer_data()