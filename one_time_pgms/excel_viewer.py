#!/usr/bin/env python3
"""
Excel File Viewer for comparing 绕嵌排 sheets from old and new folders
"""

import os
import glob
from flask import Flask, render_template, request, jsonify
import pandas as pd

app = Flask(__name__)

# Paths to the folders
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OLD_FOLDER = os.path.join(BASE_DIR, 'old_payroll')
NEW_FOLDER = os.path.join(BASE_DIR, 'new_payroll')
# Available sheet names
SHEET_NAMES = ['绕嵌排', '精加工', '喷漆装配']

def get_available_files():
    """Get all available files from both folders"""
    # Get files from old folder
    old_files = []
    for ext in ['*.xls', '*.xlsx']:
        old_files.extend(glob.glob(os.path.join(OLD_FOLDER, ext)))
    
    # Get files from new folder
    new_files = []
    for ext in ['*.xls', '*.xlsx']:
        new_files.extend(glob.glob(os.path.join(NEW_FOLDER, ext)))
    
    # Extract base names without extension
    old_basenames = {os.path.splitext(os.path.basename(f))[0]: f for f in old_files}
    new_basenames = {os.path.splitext(os.path.basename(f))[0]: f for f in new_files}
    
    # Find common files that exist in both folders
    common_files = sorted(set(old_basenames.keys()) & set(new_basenames.keys()))
    
    # Create file list with full paths
    file_list = []
    for basename in common_files:
        file_list.append({
            'name': basename,
            'old_path': old_basenames[basename],
            'new_path': new_basenames[basename]
        })
    
    return file_list

def read_excel_sheet(file_path, sheet_name):
    """Read specific sheet from Excel file with flexible matching"""
    try:
        # Try to read the sheet directly first
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        # If sheet not found, try to find similar sheet names
        try:
            xl = pd.ExcelFile(file_path)
            available_sheets = xl.sheet_names
            
            # Define mapping for common variations
            sheet_variations = {
                '精加工': ['金加工', '精加工'],
                '喷漆装配': ['装配喷漆', '喷漆装配']
            }
            
            # Look for exact matches first
            if sheet_name in available_sheets:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                return df
            
            # Look for variations
            for target_name, variations in sheet_variations.items():
                if sheet_name == target_name:
                    for variation in variations:
                        if variation in available_sheets:
                            df = pd.read_excel(file_path, sheet_name=variation)
                            return df
            
            # Look for sheets containing the target name
            matching_sheets = [s for s in available_sheets if sheet_name in s]
            if matching_sheets:
                # Use the first matching sheet
                df = pd.read_excel(file_path, sheet_name=matching_sheets[0])
                return df
            else:
                # Return empty DataFrame if no matching sheet found
                return pd.DataFrame({'Error': [f'Sheet "{sheet_name}" not found. Available sheets: {available_sheets}']})
        except Exception as e2:
            return pd.DataFrame({'Error': [f'Error reading file: {str(e2)}']})

@app.route('/')
def index():
    """Main page showing the first file"""
    files = get_available_files()
    if not files:
        return "No matching files found in both folders"
    
    # Get the first file data with default sheet
    current_file = files[0]
    default_sheet = SHEET_NAMES[0]
    
    # Read data from both files
    old_data = read_excel_sheet(current_file['old_path'], default_sheet)
    new_data = read_excel_sheet(current_file['new_path'], default_sheet)
    
    # Convert to HTML tables
    old_html = old_data.to_html(classes='table table-striped table-bordered', index=False, escape=False) if not old_data.empty else "<p>No data available</p>"
    new_html = new_data.to_html(classes='table table-striped table-bordered', index=False, escape=False) if not new_data.empty else "<p>No data available</p>"
    
    return render_template('index.html', 
                         old_data=old_html,
                         new_data=new_html,
                         current_file=current_file['name'],
                         files=files,
                         current_index=0,
                         sheet_names=SHEET_NAMES,
                         current_sheet=default_sheet)

@app.route('/get_file/<int:file_index>/<sheet_name>')
def get_file(file_index, sheet_name):
    """Get specific file data by index and sheet name"""
    files = get_available_files()
    
    if file_index < 0 or file_index >= len(files):
        return jsonify({'error': 'Invalid file index'})
    
    current_file = files[file_index]
    
    # Read data from both files
    old_data = read_excel_sheet(current_file['old_path'], sheet_name)
    new_data = read_excel_sheet(current_file['new_path'], sheet_name)
    
    # Convert to HTML tables
    old_html = old_data.to_html(classes='table table-striped table-bordered', index=False, escape=False) if not old_data.empty else "<p>No data available</p>"
    new_html = new_data.to_html(classes='table table-striped table-bordered', index=False, escape=False) if not new_data.empty else "<p>No data available</p>"
    
    return jsonify({
        'old_data': old_html,
        'new_data': new_html,
        'current_file': current_file['name'],
        'current_index': file_index,
        'total_files': len(files),
        'current_sheet': sheet_name
    })

if __name__ == '__main__':
    # Create templates directory if it doesn't exist
    templates_dir = os.path.join(os.path.dirname(__file__), 'templates')
    os.makedirs(templates_dir, exist_ok=True)
    
    app.run(debug=True, port=5000)
