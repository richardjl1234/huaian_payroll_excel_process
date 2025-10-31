#!/usr/bin/env python3
"""
Debug script to identify the decimal conversion issue in the 202507.xls file.
"""

import os
import pandas as pd
import xlrd
from decimal import Decimal, ROUND_HALF_UP
import logging
import sqlite3
from excel_processor.sheet_processor import get_all_data_from_sheet, split_raw_sheet_contents
from excel_processor.special_logic import special_logic_preprocess_df

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def inspect_problematic_data():
    """Inspect the problematic data in 202507.xls file, sheet '喷漆装配'"""
    
    file_name = "202507.xls"
    sheet_name = "喷漆装配"
    
    logger.info(f"Debugging file: {file_name}, sheet: {sheet_name}")
    
    try:
        # Step 1: Get raw sheet data
        df_summary = get_all_data_from_sheet(file_name, sheet_name)
        logger.info(f"Raw sheet shape: {df_summary.shape}")
        logger.info(f"Raw sheet columns: {list(df_summary.columns)}")
        
        # Step 2: Split into multiple dataframes
        dfs = split_raw_sheet_contents(df_summary)
        logger.info(f"Number of dataframes after splitting: {len(dfs)}")
        
        # Focus on the problematic table (Table 1)
        if len(dfs) >= 1:
            df_table1 = dfs[0]
            logger.info(f"Table 1 shape: {df_table1.shape}")
            logger.info(f"Table 1 columns: {list(df_table1.columns)}")
            
            # Step 3: Apply special logic preprocessing
            df_processed = special_logic_preprocess_df(df_table1, sheet_name, file_name, 1)
            logger.info(f"After special logic - shape: {df_processed.shape}")
            logger.info(f"After special logic - columns: {list(df_processed.columns)}")
            
            # Step 4: Check for problematic numeric columns
            numeric_columns = ['计件数量', '系数', '定额', '金额']
            
            for col in numeric_columns:
                if col in df_processed.columns:
                    logger.info(f"\n--- Checking column: {col} ---")
                    logger.info(f"Data type: {df_processed[col].dtype}")
                    logger.info(f"Sample values: {df_processed[col].head(10).tolist()}")
                    
                    # Check for problematic values
                    for idx, value in df_processed[col].items():
                        try:
                            # Try to convert to numeric
                            pd.to_numeric([value], errors='raise')
                        except Exception as e:
                            logger.error(f"Problematic value at index {idx}: '{value}' - Error: {e}")
            
            # Step 5: Test the decimal conversion that's failing
            logger.info("\n--- Testing decimal conversion ---")
            for col in numeric_columns:
                if col in df_processed.columns:
                    logger.info(f"Testing column: {col}")
                    for idx, value in df_processed[col].items():
                        try:
                            # This is the exact conversion that happens in sheet_processor.py
                            numeric_val = pd.to_numeric([value], errors='coerce').fillna(0.0)[0]
                            decimal_val = Decimal(str(numeric_val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                            float_val = float(decimal_val)
                            logger.info(f"  Index {idx}: '{value}' -> {numeric_val} -> {decimal_val} -> {float_val}")
                        except Exception as e:
                            logger.error(f"  ERROR at index {idx}: '{value}' - {type(e).__name__}: {e}")
            
            # Step 6: Print the entire problematic dataframe for manual inspection
            logger.info("\n--- Complete Table 1 Data ---")
            logger.info(df_processed.to_string())
            
        else:
            logger.error("No dataframes found after splitting!")
            
    except Exception as e:
        logger.error(f"Error during debugging: {e}")
        import traceback
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    inspect_problematic_data()
