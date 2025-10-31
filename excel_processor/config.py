"""
Configuration module for shared constants and settings.
"""

import logging
import os
from pathlib import Path


# Database configuration
# Database file location - moved to parent folder
CURRENT_DIR = Path(__file__).parent
DATABASE_PATH =  str(CURRENT_DIR.parent.parent / "payroll_database.db")

# Expected columns for payroll data processing
expected_columns = ['职员全名', '日期','客户名称', '型号', '工序全名', '工序', '计件数量', '系数', '定额', '金额', '备注']

# Minimum number of expected columns that should be found in a valid dataframe
COMMON_COL_COUNT = 4

# Global logging configuration
def setup_global_logging():
    """
    Set up global logging configuration for all programs.
    This ensures consistent logging behavior across batch_process.py, sheet_gen.py, 
    sheet_processor.py, and df_gen.py.
    """
    # Get the project root directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    log_file_path = os.path.join(project_root, 'log_batch.txt')
    
    # Configure root logger
    root_logger = logging.getLogger()
    
    # Clear any existing handlers to avoid conflicts
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # Set up handlers
    stream_handler = logging.StreamHandler()
    file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    stream_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    
    # Add handlers to root logger
    root_logger.addHandler(stream_handler)
    root_logger.addHandler(file_handler)
    root_logger.setLevel(logging.INFO)
    
    return root_logger

# Initialize global logging when this module is imported
setup_global_logging()
