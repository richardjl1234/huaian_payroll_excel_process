#!/usr/bin/env python3
"""
Sheet generator module for processing Excel files.
Generates sheet contents from Excel files for further processing.
"""

import os
import openpyxl
import xlrd
import pandas as pd
import logging
from collections import namedtuple
from typing import List, Generator

# Import existing functions
try:
    from .sheet_processor import  get_all_data_from_sheet
    from .config import setup_global_logging
except ImportError:
    # Fallback for when running directly
    import sys
    import os
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from excel_processor.sheet_processor import  get_all_data_from_sheet
    from excel_processor.config import setup_global_logging

# Set up logging using global configuration
setup_global_logging()
logger = logging.getLogger(__name__)

# Define named tuples for structured data
SheetContents = namedtuple('SheetContents', ['raw_sheet_contents', 'file_name', 'sheet_name'])


def get_excel_files() -> List[str]:
    """
    Get all Excel files from new_payroll and old_payroll folders.
    
    Returns:
        List[str]: List of Excel file names
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)  # Go up one level to project root
    new_payroll_path = os.path.join(project_root, 'new_payroll')
    old_payroll_path = os.path.join(project_root, 'old_payroll')
    
    excel_files = []
    
    # Get files from new_payroll folder
    if os.path.exists(new_payroll_path):
        for file in os.listdir(new_payroll_path):
            file_path = os.path.join(new_payroll_path, file)
            if os.path.isfile(file_path) and file.lower().endswith(('.xls', '.xlsx')):
                excel_files.append(file)
    
    # Get files from old_payroll folder
    if os.path.exists(old_payroll_path):
        for file in os.listdir(old_payroll_path):
            file_path = os.path.join(old_payroll_path, file)
            if os.path.isfile(file_path) and file.lower().endswith(('.xls', '.xlsx')):
                excel_files.append(file)
    
    # Sort files by name
    excel_files.sort()
    
    return excel_files


def process_one_file(file_path: str, file_name: str):
    """
    Process a single Excel file and return workbook and sheet names.
    
    Parameters:
        file_path (str): Full path to the Excel file
        file_name (str): Name of the Excel file
        
    Returns:
        tuple: (success: bool, workbook: object, sheet_names: List[str], error_msg: str)
    """
    try:
        workbook = None
        sheet_names = []
        
        # Read Excel file based on file type
        if file_name.lower().endswith('.xlsx'):
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            sheet_names = workbook.sheetnames
        elif file_name.lower().endswith('.xls'):
            workbook = xlrd.open_workbook(file_path)
            sheet_names = workbook.sheet_names()
        
        return True, workbook, sheet_names, ""
    except Exception as e:
        return False, None, [], str(e)


def sheet_gen(excel_files: List[str]) -> Generator[SheetContents, None, None]:
    """
    Generator function that yields sheet contents from Excel files.
    
    Parameters:
        excel_files (List[str]): List of Excel file names
        
    Yields:
        SheetContents: Named tuple containing raw_sheet_contents, file_name, sheet_name
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)  # Go up one level to project root
    new_payroll_path = os.path.join(project_root, 'new_payroll')
    old_payroll_path = os.path.join(project_root, 'old_payroll')
    
    for file_name in excel_files:
        logger.info(f"Processing file: {file_name}")
        
        # Determine file path
        file_path = None
        if os.path.exists(os.path.join(new_payroll_path, file_name)):
            file_path = os.path.join(new_payroll_path, file_name)
        elif os.path.exists(os.path.join(old_payroll_path, file_name)):
            file_path = os.path.join(old_payroll_path, file_name)
        else:
            logger.warning(f"File not found: {file_name}")
            continue
        
        # Process the file
        success, workbook, sheet_names, error_msg = process_one_file(file_path, file_name)
        
        if not success:
            logger.error(f"Failed to process file {file_name}: {error_msg}")
            continue
        
        # Process each sheet
        for sheet_name in sheet_names:
            # Skip sheets containing '汇总' in the name
            if '汇总' in sheet_name or '统计' in sheet_name or 'deleted' in sheet_name:
                logger.info(f"Skipping sheet '{sheet_name}' in file '{file_name}' (contains '汇总' or '统计' or 'deleted')")
                continue
            
            try:
                # Get raw sheet contents using the new get_all_data_from_sheet function
                df_summary = get_all_data_from_sheet(file_name, sheet_name)
                
                # Yield the sheet contents
                yield SheetContents(
                    raw_sheet_contents=df_summary,
                    file_name=file_name,
                    sheet_name=sheet_name
                )
                
            except Exception as e:
                logger.error(f"Error processing sheet '{sheet_name}' in file '{file_name}': {str(e)}")
                continue


def test_sheet_gen():
    """
    Test function for sheet_gen to verify it works correctly.
    """
    print("=" * 50)
    print("Testing sheet_gen function...")
    print("=" * 50)
    
    # Get Excel files
    excel_files = get_excel_files()
    print(f"Found {len(excel_files)} Excel files")
    
    if not excel_files:
        print("No Excel files found. Please ensure there are files in new_payroll or old_payroll folders.")
        return
    
    # Test sheet_gen with first 2 files
    test_files = excel_files[:]
    print(f"Testing with files: {test_files}")
    
    sheet_count = 0
    for sheet_contents in sheet_gen(test_files):
        print(f"Sheet {sheet_count + 1}:")
        print(f"  File: {sheet_contents.file_name}")
        print(f"  Sheet: {sheet_contents.sheet_name}")
        print(f"  Shape: {sheet_contents.raw_sheet_contents.shape}")
        print(f"  Columns: {list(sheet_contents.raw_sheet_contents.columns)}")
        # print(f" row_sheet_contents: {sheet_contents.raw_sheet_contents.to_string()}")
        print()
        sheet_count += 1
        
    
    print(f"Test completed. Processed {sheet_count} sheets.")
    return sheet_count


if __name__ == "__main__":
    # Run the test function when executed directly
    test_sheet_gen()
