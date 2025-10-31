#!/usr/bin/env python3
"""
Batch processing program for payroll Excel files.
Processes all files in new_payroll and old_payroll folders without frontend.
Refactored to reuse functions from sheet_gen and df_gen modules.
"""

import os
import sqlite3
import logging
import sys

# Import existing functions from modules
from excel_processor.sheet_processor import load_df_to_db
from excel_processor.sheet_gen import sheet_gen, get_excel_files, SheetContents
from excel_processor.df_gen import df_gen, SplitDataFrame
from excel_processor.config import setup_global_logging, DATABASE_PATH

# Set up logging using global configuration
setup_global_logging()
logger = logging.getLogger(__name__)




def clean_database_tables():
    """Clean payroll_details and load_log tables in the database."""
    try:
        conn = sqlite3.connect(DATABASE_PATH)
        conn.execute("DELETE FROM payroll_details")
        conn.execute("DELETE FROM load_log")
        conn.commit()
        conn.close()
        logger.info("Database tables cleaned successfully")
    except Exception as e:
        logger.error(f"Error cleaning database tables: {e}")


def _process_excel_files(excel_files, clean_db=True):
    """
    Private function to process Excel files and load to database.
    
    Parameters:
        excel_files (list): List of Excel file names to process
        clean_db (bool): Whether to clean database tables before processing
    
    Returns:
        tuple: (total_sheets, total_dataframes, successful_loads, failed_loads)
    """
    if clean_db:
        clean_database_tables()
    
    logger.info(f"Found {len(excel_files)} Excel files to process")
    
    # Process files using sheet_gen and df_gen
    total_sheets = 0
    total_dataframes = 0
    successful_loads = 0
    failed_loads = 0
    
    for sheet_contents in sheet_gen(excel_files):
        logger.info(f"Processing sheet: {sheet_contents.file_name} - {sheet_contents.sheet_name}")
        total_sheets += 1
        
        for split_df in df_gen(sheet_contents):
            logger.info(f"  Generated dataframe: Table {split_df.table_index}, Shape={split_df.split_df.shape}")
            total_dataframes += 1
            
            # Call load_df_to_db to load the dataframe to database
            result = load_df_to_db(
                split_df.split_df, 
                split_df.file_name, 
                split_df.sheet_name, 
                split_df.table_index
            )
            
            if "Successfully" in result:
                successful_loads += 1
                logger.info(f"    ✓ Successfully loaded to database: {result}")
            else:
                failed_loads += 1
                logger.error(f"    ✗ Failed to load to database: {result}")
    
    return total_sheets, total_dataframes, successful_loads, failed_loads


def batch_process_main():
    """
    Main batch processing logic with database loading.
    Complete end-to-end pipeline from files to database.
    """
    logger.info("Starting batch process main logic (with database loading)...")

    # Get all Excel files
    excel_files = get_excel_files()
    # excel_files = excel_files[:4]
    
    # Process files using the common logic
    total_sheets, total_dataframes, successful_loads, failed_loads = _process_excel_files(excel_files, clean_db=True)
    
    logger.info(f"Batch process completed. Processed {total_sheets} sheets and {total_dataframes} dataframes.")
    logger.info(f"Database loading results: {successful_loads} successful, {failed_loads} failed")


def process_single_file(file_name: str):
    """
    Process a single Excel file specified by command line parameter.
    
    Parameters:
        file_name (str): Name of the Excel file to process
    """
    # Check if file exists in either new_payroll or old_payroll
    current_dir = os.path.dirname(os.path.abspath(__file__))
    new_payroll_path = os.path.join(current_dir, 'new_payroll')
    old_payroll_path = os.path.join(current_dir, 'old_payroll')
    
    file_exists = False
    if os.path.exists(os.path.join(new_payroll_path, file_name)):
        file_exists = True
    elif os.path.exists(os.path.join(old_payroll_path, file_name)):
        file_exists = True
    
    if not file_exists:
        raise FileNotFoundError(f"File '{file_name}' not found in new_payroll or old_payroll folders")
    
    # Process only this file
    logger.info(f"Processing single file: {file_name}")
    excel_files = [file_name]
    
    # Process files using the common logic (don't clean DB for single file processing)
    total_sheets, total_dataframes, successful_loads, failed_loads = _process_excel_files(excel_files, clean_db=False)
    
    logger.info(f"Single file processing completed. Processed {total_sheets} sheets and {total_dataframes} dataframes.")
    logger.info(f"Database loading results: {successful_loads} successful, {failed_loads} failed")


if __name__ == "__main__":
    import sys
    
    # Check for command line parameter
    if len(sys.argv) > 1:
        # Process single file mode
        file_name = sys.argv[1]
        try:
            process_single_file(file_name)
        except FileNotFoundError as e:
            logger.error(f"Error: {e}")
            logger.info("Available files in new_payroll and old_payroll folders:")
            all_files = get_excel_files()
            for f in all_files:
                logger.info(f"  - {f}")
        except Exception as e:
            logger.error(f"Error processing file '{file_name}': {e}")
    else:
        # Normal mode - use batch_process_main function for complete processing
        batch_process_main()
