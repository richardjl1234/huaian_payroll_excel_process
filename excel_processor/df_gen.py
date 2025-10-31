#!/usr/bin/env python3
"""
DataFrame generator module for processing sheet contents.
Generates split dataframes from sheet contents for database loading.
"""

import logging
from collections import namedtuple
from typing import Generator

# Import existing functions
try:
    from .sheet_processor import split_raw_sheet_contents
    from .config import setup_global_logging
except ImportError:
    # Fallback for when running directly
    import sys
    import os
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from excel_processor.sheet_processor import split_raw_sheet_contents
    from excel_processor.config import setup_global_logging

# Set up logging using global configuration
setup_global_logging()
logger = logging.getLogger(__name__)

# Define named tuples for structured data
SplitDataFrame = namedtuple('SplitDataFrame', ['split_df', 'file_name', 'sheet_name', 'table_index'])


def df_gen(sheet_contents) -> Generator[SplitDataFrame, None, None]:
    """
    Generator function that yields split dataframes from sheet contents.
    
    Parameters:
        sheet_contents: Named tuple containing raw_sheet_contents, file_name, sheet_name
        
    Yields:
        SplitDataFrame: Named tuple containing split_df, file_name, sheet_name, table_index
    """
    try:
        # Use the new split_raw_sheet_contents function to split the sheet into multiple dataframes
        dfs = split_raw_sheet_contents(sheet_contents.raw_sheet_contents)
        
        # Yield each split dataframe with table index
        for table_index, df in enumerate(dfs):
            # If the dataframe is empty, still yield it but log a warning
            if df.empty:
                logger.warning(f"Empty dataframe found: File={sheet_contents.file_name}, Sheet={sheet_contents.sheet_name}, Table={table_index + 1}")
            
            yield SplitDataFrame(
                split_df=df,
                file_name=sheet_contents.file_name,
                sheet_name=sheet_contents.sheet_name,
                table_index=table_index + 1  # Start from 1 instead of 0
            )
            
    except Exception as e:
        logger.error(f"Error splitting dataframes for file '{sheet_contents.file_name}', sheet '{sheet_contents.sheet_name}': {str(e)}")


def test_df_gen():
    """
    Test function for df_gen to verify it works correctly.
    """
    print("=" * 50)
    print("Testing df_gen function...")
    print("=" * 50)
    
    # Import sheet_gen to get test data
    try:
        from .sheet_gen import get_excel_files, sheet_gen
    except ImportError:
        # Fallback for when running directly
        import sys
        import os
        sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from excel_processor.sheet_gen import get_excel_files, sheet_gen
    
    # Get Excel files
    excel_files = get_excel_files()
    print(f"Found {len(excel_files)} Excel files")
    
    if not excel_files:
        print("No Excel files found. Please ensure there are files in new_payroll or old_payroll folders.")
        return
    
    # Test with first file
    test_files = excel_files[:1]
    print(f"Testing with file: {test_files[0]}")
    
    df_count = 0
    sheet_count = 0
    
    for sheet_contents in sheet_gen(test_files):
        sheet_count += 1
        print(f"\nSheet {sheet_count}: {sheet_contents.file_name} - {sheet_contents.sheet_name}")
        print(f"  Original shape: {sheet_contents.raw_sheet_contents.shape}")
        
        # Test df_gen for this sheet
        for split_df in df_gen(sheet_contents):
            df_count += 1
            print(f"  Table {split_df.table_index}:")
            print(f"    Shape: {split_df.split_df.shape}")
            print(f"    Columns: {list(split_df.split_df.columns)}")
            print(f"    First few rows:")
            print(split_df.split_df.head(3).to_string())
            print()
        
        # Limit test to first sheet for quick testing
        break
    
    print(f"Test completed. Processed {sheet_count} sheets and {df_count} dataframes.")
    return df_count


def test_df_gen_with_mock_data():
    """
    Test function for df_gen using mock data to verify the logic.
    This is useful when there are no actual Excel files available.
    """
    print("=" * 50)
    print("Testing df_gen with mock data...")
    print("=" * 50)
    
    # Create a mock SheetContents object
    from collections import namedtuple
    SheetContents = namedtuple('SheetContents', ['raw_sheet_contents', 'file_name', 'sheet_name'])
    
    # Create a simple mock dataframe
    import pandas as pd
    mock_df = pd.DataFrame({
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'Salary': [50000, 60000, 70000]
    })
    
    mock_sheet_contents = SheetContents(
        raw_sheet_contents=mock_df,
        file_name='test_file.xlsx',
        sheet_name='test_sheet'
    )
    
    print(f"Mock sheet contents:")
    print(f"  File: {mock_sheet_contents.file_name}")
    print(f"  Sheet: {mock_sheet_contents.sheet_name}")
    print(f"  Shape: {mock_sheet_contents.raw_sheet_contents.shape}")
    print()
    
    # Test df_gen with mock data
    df_count = 0
    for split_df in df_gen(mock_sheet_contents):
        df_count += 1
        print(f"Generated dataframe {df_count}:")
        print(f"  Table index: {split_df.table_index}")
        print(f"  Shape: {split_df.split_df.shape}")
        print(f"  Data:")
        print(split_df.split_df.to_string())
        print()
    
    print(f"Mock test completed. Generated {df_count} dataframes.")
    return df_count


if __name__ == "__main__":
    # Run the test functions when executed directly
    print("Running df_gen tests...")
    print()
    
    # First test with actual data if available
    actual_df_count = test_df_gen()
    
    print("\n" + "="*50)
    print("If no actual files were found, running mock test...")
    print("="*50)
    
    # If no actual data was processed, run mock test
    if actual_df_count == 0:
        test_df_gen_with_mock_data()
