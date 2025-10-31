#!/usr/bin/env python3
"""
Debug script to identify the specific problematic value in the 202507.xls file.
"""

import os
import pandas as pd
import xlrd
from decimal import Decimal, ROUND_HALF_UP
import logging
from excel_processor.sheet_processor import get_all_data_from_sheet, split_raw_sheet_contents

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def find_problematic_value():
    """Find the exact problematic value causing the decimal conversion error"""
    
    file_name = "202507.xls"
    sheet_name = "喷漆装配"
    
    logger.info(f"Debugging file: {file_name}, sheet: {sheet_name}")
    
    try:
        # Step 1: Get raw sheet data
        df_summary = get_all_data_from_sheet(file_name, sheet_name)
        
        # Step 2: Split into multiple dataframes
        dfs = split_raw_sheet_contents(df_summary)
        
        # Focus on the problematic table (Table 1)
        if len(dfs) >= 1:
            df_table1 = dfs[0]
            
            # Step 3: Find rows where '职员全名' is '前装'
            qianzhuang_rows = df_table1[df_table1['职员全名'] == '前装']
            logger.info(f"Found {len(qianzhuang_rows)} rows with '前装'")
            
            # Step 4: Check the problematic values in '计件数量' column
            logger.info("\n--- Checking '计件数量' values for '前装' rows ---")
            for idx, row in qianzhuang_rows.iterrows():
                jijian_count = row['计件数量']
                logger.info(f"Row {idx}: '计件数量' = '{jijian_count}' (type: {type(jijian_count)})")
                
                # Test the problematic conversion
                try:
                    if pd.notna(jijian_count) and jijian_count != '':
                        # This is the exact conversion that's failing
                        decimal_val = Decimal(str(jijian_count))
                        logger.info(f"  ✓ Successfully converted to Decimal: {decimal_val}")
                    else:
                        logger.info(f"  ⚠ Empty or NaN value")
                except Exception as e:
                    logger.error(f"  ✗ ERROR converting '{jijian_count}': {type(e).__name__}: {e}")
            
            # Step 5: Print the complete '前装' rows for inspection
            logger.info("\n--- Complete '前装' rows data ---")
            logger.info(qianzhuang_rows.to_string())
            
        else:
            logger.error("No dataframes found after splitting!")
            
    except Exception as e:
        logger.error(f"Error during debugging: {e}")
        import traceback
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    find_problematic_value()
